from __future__ import annotations

import base64
import io
import math
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal
from functools import lru_cache
from pathlib import Path
from typing import Iterable, Mapping, Sequence
from zoneinfo import ZoneInfo

from PIL import Image, ImageDraw, ImageFont
from sqlalchemy import create_engine, text
from sqlalchemy.pool import NullPool

from reporting_bot.config import Settings
from reporting_bot.simple_xlsx import STYLE, Workbook
from reporting_bot.ks_scope import ROSTER, DEPARTMENTS


REPORT_TITLES = {
    "all": "КС — Общий отчёт",
    "plan": "КС — План-факт",
    "managers": "КС — Менеджеры",
    "funnel": "КС — Воронка",
    "projects": "КС — Проекты и оплаты",
}

SERVICE_LABELS = {
    "sout": "СОУТ",
    "obuchenie": "Обучение",
    "opk": "ОПК",
    "opk_eth": "ОПК ЭТХ",
    "autsorsing": "Аутсорсинг",
    "pk": "ПК",
    "ppk": "ППК",
    "suot": "СУОТ",
    "sbkts": "СБКТС",
    "other": "Другие услуги",
    "hassp": "ХАССП",
    "audit": "Аудит",
    "zamery": "Замеры",
}

WAITING_STATUSES = {
    "Ждём ШР и реквизиты",
    "Ждем штатное расписание",
    "Ждем реквизиты",
}
@dataclass(frozen=True)
class FilterOption:
    token: str
    label: str


@dataclass(frozen=True)
class KsFilterOptions:
    departments: tuple[FilterOption, ...]
    managers: tuple[FilterOption, ...]
    products: tuple[FilterOption, ...]


@dataclass(frozen=True)
class KsFilters:
    department_token: str = "-"
    manager_token: str = "-"
    service: str = "-"


@dataclass(frozen=True)
class KsEvent:
    kind: str
    event_id: str
    event_at: date | datetime
    manager: str
    department: str
    service: str
    amount: Decimal = Decimal(0)
    reference: str = "—"
    status: str = "—"
    stage: str = "—"
    current_step: str = "—"
    paid_amount: Decimal | None = None
    cohort_at: date | datetime | None = None
    project_id: str = ""
    is_agreed: bool = False


@dataclass(frozen=True)
class KsReportData:
    report_kind: str
    date_from: date
    date_to: date
    comparison_from: date | None
    comparison_to: date | None
    comparison_mode: str
    filters: KsFilters
    filter_labels: Mapping[str, str]
    events: tuple[KsEvent, ...]
    comparison_events: tuple[KsEvent, ...]


@dataclass(frozen=True)
class KsReportArtifact:
    workbook: bytes
    workbook_filename: str
    chart: bytes
    chart_filename: str
    caption: str


from reporting_bot.ks_scope import (
    DEPARTMENT_OPTIONS_SQL, MANAGER_OPTIONS_SQL, PRODUCT_OPTIONS_SQL,
    REQUESTS_SQL, OFFERS_SQL, PROJECTS_SQL, PAYMENTS_SQL, CALLS_SQL,
)


class KsReportService:
    def __init__(self, settings: Settings) -> None:
        self._settings = settings
        self._database_url = settings.database_url.replace(
            "postgresql://", "postgresql+psycopg://", 1
        )

    def filter_options(self, department_token: str = "-") -> KsFilterOptions:
        if department_token != "-" and department_token not in dict(DEPARTMENTS):
            raise ValueError("Список отделов КС обновлён. Запустите команду отчёта заново.")
        engine = self._engine()
        try:
            with engine.connect() as connection:
                with connection.begin():
                    self._read_only(connection)
                    departments = tuple(
                        FilterOption(str(row["token"]), str(row["label"]))
                        for row in connection.execute(text(DEPARTMENT_OPTIONS_SQL)).mappings()
                    )
                    managers = tuple(
                        FilterOption(str(row["token"]), str(row["label"] or "Без имени"))
                        for row in connection.execute(
                            text(MANAGER_OPTIONS_SQL),
                            {"department_token": department_token},
                        ).mappings()
                    )
                    products = tuple(
                        FilterOption(str(row["token"]), _service_label(str(row["label"])))
                        for row in connection.execute(text(PRODUCT_OPTIONS_SQL)).mappings()
                    )
            return KsFilterOptions(departments, managers, products)
        finally:
            engine.dispose()

    def create(
        self,
        report_kind: str,
        date_from: date,
        date_to: date,
        filters: KsFilters | None = None,
        comparison_mode: str = "previous",
    ) -> KsReportArtifact:
        if report_kind not in REPORT_TITLES:
            raise ValueError("Неизвестный отчёт КС.")
        if date_to < date_from:
            raise ValueError("Дата окончания не может быть раньше даты начала.")
        if (date_to - date_from).days > 366:
            raise ValueError("Максимальный период отчёта — 366 дней.")
        if comparison_mode not in {"previous", "year", "none"}:
            raise ValueError("Неизвестный режим сравнения.")

        selected_filters = filters or KsFilters()
        comparison_from, comparison_to = _comparison_period(
            date_from, date_to, comparison_mode
        )
        events, comparison_events, labels = self._load(
            date_from,
            date_to,
            comparison_from,
            comparison_to,
            selected_filters,
        )
        data = KsReportData(
            report_kind=report_kind,
            date_from=date_from,
            date_to=date_to,
            comparison_from=comparison_from,
            comparison_to=comparison_to,
            comparison_mode=comparison_mode,
            filters=selected_filters,
            filter_labels=labels,
            events=events,
            comparison_events=comparison_events,
        )
        stem = f"Отчет_КС_{report_kind}_{date_from.isoformat()}_{date_to.isoformat()}"
        return KsReportArtifact(
            workbook=build_ks_workbook(data),
            workbook_filename=stem + ".xlsx",
            chart=build_ks_chart(data),
            chart_filename=stem + ".png",
            caption=_artifact_caption(data),
        )

    def _engine(self):
        return create_engine(
            self._database_url,
            poolclass=NullPool,
            connect_args={"connect_timeout": self._settings.connect_timeout_seconds},
        )

    def _read_only(self, connection) -> None:
        connection.execute(text("SET TRANSACTION ISOLATION LEVEL REPEATABLE READ, READ ONLY"))
        connection.execute(
            text("SELECT set_config('statement_timeout', :timeout, true)"),
            {"timeout": f"{max(self._settings.statement_timeout_ms, 30_000)}ms"},
        )

    def _load(
        self,
        date_from: date,
        date_to: date,
        comparison_from: date | None,
        comparison_to: date | None,
        filters: KsFilters,
    ) -> tuple[tuple[KsEvent, ...], tuple[KsEvent, ...], Mapping[str, str]]:
        engine = self._engine()
        try:
            with engine.connect() as connection:
                with connection.begin():
                    self._read_only(connection)
                    options = _load_filter_options(connection, filters.department_token)
                    _validate_filters(options, filters)
                    events = _load_events(connection, date_from, date_to, filters)
                    comparison_events = (
                        _load_events(connection, comparison_from, comparison_to, filters)
                        if comparison_from is not None and comparison_to is not None
                        else ()
                    )
            labels = {
                "department": _option_label(options.departments, filters.department_token, "Все отделы"),
                "manager": _option_label(options.managers, filters.manager_token, "Все менеджеры"),
                "product": _option_label(options.products, filters.service, "Все продукты"),
            }
            return events, comparison_events, labels
        finally:
            engine.dispose()


def _load_filter_options(connection, department_token: str) -> KsFilterOptions:
    departments = tuple(
        FilterOption(str(row["token"]), str(row["label"]))
        for row in connection.execute(text(DEPARTMENT_OPTIONS_SQL)).mappings()
    )
    managers = tuple(
        FilterOption(str(row["token"]), str(row["label"] or "Без имени"))
        for row in connection.execute(
            text(MANAGER_OPTIONS_SQL), {"department_token": department_token}
        ).mappings()
    )
    products = tuple(
        FilterOption(str(row["token"]), _service_label(str(row["label"])))
        for row in connection.execute(text(PRODUCT_OPTIONS_SQL)).mappings()
    )
    return KsFilterOptions(departments, managers, products)


def _validate_filters(options: KsFilterOptions, filters: KsFilters) -> None:
    choices = (
        (filters.department_token, {o.token for o in options.departments}),
        (filters.manager_token, {o.token for o in options.managers}),
        (filters.service, {o.token for o in options.products}),
    )
    for selected, allowed in choices:
        if selected != "-" and not set(selected.split(",")) <= allowed:
            raise ValueError("Состав КС или фильтры изменились. Запустите команду отчёта заново.")


def _option_label(
    options: Sequence[FilterOption], token: str, default: str
) -> str:
    if token == "-":
        return default
    labels_by_token = {option.token: option.label for option in options}
    return ", ".join(
        labels_by_token.get(value, value)
        for value in token.split(",")
        if value
    )


def _load_events(
    connection,
    date_from: date,
    date_to: date,
    filters: KsFilters,
) -> tuple[KsEvent, ...]:
    params = {
        "date_from": date_from.isoformat(),
        "date_to_exclusive": (date_to + timedelta(days=1)).isoformat(),
        "department_token": filters.department_token,
        "manager_token": filters.manager_token,
        "service": filters.service,
    }
    events: list[KsEvent] = []
    for kind, sql in (
        ("Лид", REQUESTS_SQL),
        ("КП", OFFERS_SQL),
        ("Проект", PROJECTS_SQL),
        ("Платёж", PAYMENTS_SQL),
        ("Звонок", CALLS_SQL),
    ):
        result_rows = list(connection.execute(text(sql), params).mappings())
        if len(result_rows) > 10000:
            raise ValueError(f"Более 10 000 записей «{kind}». Сократите период или выберите отдел: отчёт не обрезается автоматически.")
        for row in result_rows:
            events.append(
                KsEvent(
                    kind=kind,
                    event_id=str(row["event_id"]),
                    event_at=row["event_at"],
                    manager=str(row["manager"] or "Не назначен"),
                    department=str(row["department"] or "Без отдела"),
                    service=str(row["service"] or ""),
                    amount=Decimal(str(row.get("amount") or 0)),
                    reference=str(row.get("reference") or "—"),
                    status=str(row.get("status") or "—"),
                    stage=str(row.get("stage") or "—"),
                    current_step=str(row.get("current_step") or "—"),
                    paid_amount=Decimal(str(row.get("paid_amount") or 0)),
                    cohort_at=row.get("cohort_at"),
                    project_id=str(row.get("project_id") or ""),
                    is_agreed=bool(row.get("is_agreed") or False),
                )
            )
    return tuple(
        sorted(
            events,
            key=lambda row: (row.event_at.isoformat(), row.kind, row.event_id),
        )
    )


def _comparison_period(
    date_from: date,
    date_to: date,
    mode: str,
) -> tuple[date | None, date | None]:
    if mode == "none":
        return None, None
    if mode == "year":
        def previous_year(value: date) -> date:
            try:
                return value.replace(year=value.year - 1)
            except ValueError:
                return value.replace(year=value.year - 1, day=28)

        return previous_year(date_from), previous_year(date_to)
    duration = date_to - date_from
    previous_to = date_from - timedelta(days=1)
    return previous_to - duration, previous_to


def _service_label(value: str | None) -> str:
    if not value:
        return "Не указан"
    normalized = value.strip().lower()
    if normalized == "соут":
        normalized = "sout"
    return SERVICE_LABELS.get(normalized, value)


def _events(rows: Iterable[KsEvent], kind: str) -> list[KsEvent]:
    return [row for row in rows if row.kind == kind]


def _sum(rows: Iterable[KsEvent]) -> Decimal:
    return sum((row.amount for row in rows), Decimal(0))


def _ratio(actual: Decimal | int, base: Decimal | int) -> float | None:
    if not base:
        return None
    return float(actual) / float(base)


def _change(actual: Decimal | int, previous: Decimal | int) -> float | None:
    if not previous:
        return None
    return (float(actual) - float(previous)) / float(previous)


def _metrics(rows: Sequence[KsEvent]) -> dict[str, Decimal | int | float | None]:
    leads = len(_events(rows, "Лид"))
    offers = len(_events(rows, "КП"))
    projects = _events(rows, "Проект")
    payments = _events(rows, "Платёж")
    return {
        "leads": leads,
        "offers": offers,
        "projects": len(projects),
        "cash": _sum(payments),
        "occurrence": _sum(projects),
        "conversion": _ratio(_funnel_counts(rows)["Успех"], leads),
        "calls": len(_events(rows, "Звонок")),
    }


def _working_days(date_from: date, date_to: date) -> int:
    return sum(
        1
        for offset in range((date_to - date_from).days + 1)
        if (date_from + timedelta(days=offset)).weekday() < 5
    )


def _forecast(value: Decimal, date_from: date, date_to: date) -> Decimal:
    today = datetime.now(ZoneInfo("Europe/Moscow")).date()
    if today >= date_to or today < date_from:
        return value
    elapsed = _working_days(date_from, today)
    total = _working_days(date_from, date_to)
    if not elapsed:
        return value
    return value * Decimal(total) / Decimal(elapsed)


def build_ks_workbook(data: KsReportData) -> bytes:
    workbook = Workbook(REPORT_TITLES[data.report_kind])
    if data.report_kind == "all":
        _build_all_workbook(workbook, data)
    elif data.report_kind == "plan":
        _build_plan_workbook(workbook, data)
    elif data.report_kind == "managers":
        _build_managers_workbook(workbook, data)
    elif data.report_kind == "funnel":
        _build_funnel_workbook(workbook, data)
    else:
        _build_projects_workbook(workbook, data)
    _build_methodology(workbook, data)
    return workbook.to_bytes()


def _selected_roster(data: KsReportData):
    tokens = set(data.filters.manager_token.split(","))
    return [row for row in ROSTER
            if (data.filters.department_token == "-" or row[2] == data.filters.department_token)
            and (data.filters.manager_token == "-" or row[0].replace("-", "")[:6] in tokens)]


def _build_methodology(workbook: Workbook, data: KsReportData) -> None:
    sheet = workbook.add_sheet("Правила расчёта")
    sheet.append(["Показатель", "Источник и правило"], STYLE["table_header"])
    rules = (
        ("Касса", "payment.amount > 0; deletedAt пусто; chargeAt в выбранном периоде. Частичные и несколько платежей суммируются. Возвраты не учитываются. Платежи без chargeAt или связи с проектом не включаются."),
        ("Возникновение", "projects.sale_price один раз на ID проекта. Дата — самое раннее systemlogs.created_at, когда expert_id или measurer_id изменилось с пустого на непустое. Повторное назначение не создаёт новый запуск. Без даты в журнале проект не включается; дата создания не подставляется."),
        ("Менеджер", "Проекты и касса: projects.manager_id. Лиды и КП: requests_clone.manager_id. Звонки: calls.user_id. Текущий ответственный на момент выгрузки, не эксперт и не СКС."),
        ("Лиды и конверсия", "Все неудалённые и неархивные заявки, созданные в периоде. Успех — их текущий успешный статус. Конверсия = Успех / число заявок × 100%."),
        ("КП", "Число offers по дате offers.created_at, ответственный — менеджер связанной заявки. Не ограничивается датой создания заявки."),
        ("Оплаты и остатки проектов", "Для проектов с возникновением в периоде берутся все положительные неудалённые платежи на дату выгрузки. Остаток каждого = max(sale_price − оплаты, 0). Это не касса за выбранный период и не исторический остаток на конец периода."),
        ("Согласование", "Текущий этап «Распечатка» либо наличие этого этапа в projects_steps_history."),
        ("План-ориентир", "Факт предыдущего аналогичного периода, не утверждённый план."),
        ("Распределение", "Сентябрь 2026.xlsx, лист 01.09, A3:A29. Только 22 согласованные карточки, включая пять ГТО. Это распределение применяется и к сравнению; исторические переводы между отделами не восстанавливаются."),
        ("Источник распределения", "https://disk.360.yandex.ru/edit/d/7XvglQIX1UyEi3GeuyXIWCPegnqahzm72s0qoIz-cKg6RGI0MnBrcWN6Zw?from_public=1"),
    )
    for name, rule in rules:
        r = sheet.append([name, rule], STYLE["table_text"])
        sheet.row_heights[r] = 65
    sheet.widths = {0: 30, 1: 115}
    roster_sheet = workbook.add_sheet("Состав КС")
    roster_sheet.append(["Менеджер", "Отдел КС", "ID CRM", "Группа CRM"], STYLE["table_header"])
    for uid, name, dept, group in _selected_roster(data):
        roster_sheet.append([name, dict(DEPARTMENTS)[dept], uid, group], STYLE["table_text"])
    roster_sheet.widths = {0: 30, 1: 20, 2: 40, 3: 17}
    roster_sheet.freeze_rows = 1


def _period_text(data: KsReportData) -> str:
    return f"Период: {data.date_from:%d.%m.%Y}–{data.date_to:%d.%m.%Y}"


def _filters_text(data: KsReportData) -> str:
    return (
        f"Отдел: {data.filter_labels['department']} · "
        f"Менеджер: {data.filter_labels['manager']} · "
        f"Продукт: {data.filter_labels['product']}"
    )


def _comparison_text(data: KsReportData) -> str:
    if data.comparison_from is None or data.comparison_to is None:
        return "Сравнение отключено"
    return f"Сравнение: {data.comparison_from:%d.%m.%Y}–{data.comparison_to:%d.%m.%Y}"


def _excel_column(count: int) -> str:
    result = ""
    while count:
        count, remainder = divmod(count - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _title(sheet, title: str, data: KsReportData, columns: int) -> None:
    last = _excel_column(columns)
    sheet.append([title] + [None] * (columns - 1), STYLE["title"])
    sheet.append([_period_text(data)] + [None] * (columns - 1), STYLE["subtitle"])
    sheet.append([_filters_text(data)] + [None] * (columns - 1), STYLE["subtitle"])
    sheet.append([_comparison_text(data)] + [None] * (columns - 1), STYLE["subtitle"])
    sheet.merges.extend([f"A1:{last}1", f"A2:{last}2", f"A3:{last}3", f"A4:{last}4"])
    sheet.row_heights.update({1: 28, 2: 22, 3: 22, 4: 22})


def _build_plan_workbook(
    workbook: Workbook,
    data: KsReportData,
    *,
    sheet_name: str = "Сводка",
    include_support: bool = True,
) -> None:
    current = _metrics(data.events)
    previous = _metrics(data.comparison_events)
    sheet = workbook.add_sheet(sheet_name)
    _title(sheet, REPORT_TITLES[data.report_kind], data, 8)
    sheet.append([None] * 8)
    sheet.append(
        ["Показатель", "План-ориентир", "Факт", "% выполнения", "Прогноз", "Пред. период", "Изменение", "Единица"],
        STYLE["table_header"],
    )
    rows = (
        ("Выручка по кассе", previous["cash"], current["cash"], _forecast(Decimal(current["cash"]), data.date_from, data.date_to), "руб."),
        ("Возникновение", previous["occurrence"], current["occurrence"], _forecast(Decimal(current["occurrence"]), data.date_from, data.date_to), "руб."),
        ("Количество лидов", previous["leads"], current["leads"], current["leads"], "шт."),
        ("Количество КП", previous["offers"], current["offers"], current["offers"], "шт."),
        ("Запущенные проекты", previous["projects"], current["projects"], current["projects"], "шт."),
        ("Конверсия", previous["conversion"], current["conversion"], current["conversion"], "%"),
    )
    for label, plan, fact, forecast, unit in rows:
        is_money = unit == "руб."
        is_percent = unit == "%"
        value_style = STYLE["table_money"] if is_money else STYLE["table_percent"] if is_percent else STYLE["table_number"]
        sheet.append(
            [label, plan if plan is not None else "—", fact if fact is not None else "—", _ratio(fact or 0, plan or 0), forecast if forecast is not None else "—", plan if plan is not None else "—", _change(fact or 0, plan or 0), unit],
            [STYLE["table_text"], value_style, value_style, STYLE["table_percent"], value_style, value_style, STYLE["table_percent"], STYLE["table_center"]],
        )
    sheet.append([None] * 8)
    note = sheet.append(
        ["Касса — положительные платежи по дате оплаты, без возвратов. Возникновение — стоимость проекта один раз по первому назначению замерщика или эксперта из журнала CRM. КП — по дате создания. Лиды и конверсия — когорта заявок. Менеджер — ответственный за проект/заявку, не эксперт. План-ориентир равен факту предыдущего периода, это не загруженный план."] + [None] * 7,
        STYLE["note"],
    )
    sheet.merges.append(f"A{note}:H{note}")
    sheet.row_heights[note] = 72
    sheet.widths = {0: 27, 1: 20, 2: 18, 3: 18, 4: 18, 5: 18, 6: 16, 7: 12}
    sheet.freeze_rows = 6

    if include_support:
        _build_dimension_sheet(workbook, data, "По отделам", "department")
        _build_dimension_sheet(workbook, data, "По продуктам", "service")
        _build_detail_sheet(workbook, data, data.events)


def _dimension_rows(events: Sequence[KsEvent], field: str) -> dict[str, list[KsEvent]]:
    result: dict[str, list[KsEvent]] = defaultdict(list)
    for event in events:
        if field == "service" and event.kind == "Звонок":
            continue
        key = getattr(event, field)
        if field == "service":
            key = _service_label(key)
        result[str(key or "Не указан")].append(event)
    return result


def _build_dimension_sheet(
    workbook: Workbook,
    data: KsReportData,
    name: str,
    field: str,
) -> None:
    sheet = workbook.add_sheet(name)
    _title(sheet, name, data, 9)
    sheet.append([None] * 9)
    sheet.append(
        ["Разрез", "Лиды", "КП", "Проекты", "Конверсия", "Касса, руб.", "Возникновение, руб.", "План-ориентир, руб.", "Изменение"],
        STYLE["table_header"],
    )
    current_groups = _dimension_rows(data.events, field)
    previous_groups = _dimension_rows(data.comparison_events, field)
    keys = set(current_groups) | set(previous_groups)
    if field == "department":
        keys.update(dict(DEPARTMENTS)[r[2]] for r in _selected_roster(data))
    keys = sorted(keys)
    for key in keys:
        current = _metrics(current_groups.get(key, []))
        previous = _metrics(previous_groups.get(key, []))
        sheet.append(
            [key, current["leads"], current["offers"], current["projects"], current["conversion"] or "—", current["cash"], current["occurrence"], previous["occurrence"], _change(Decimal(current["occurrence"]), Decimal(previous["occurrence"]))],
            [STYLE["table_text"], STYLE["table_number"], STYLE["table_number"], STYLE["table_number"], STYLE["table_percent"], STYLE["table_money"], STYLE["table_money"], STYLE["table_money"], STYLE["table_percent"]],
        )
    sheet.widths = {0: 27, 1: 12, 2: 12, 3: 13, 4: 15, 5: 18, 6: 22, 7: 23, 8: 15}
    sheet.freeze_rows = 6
    if keys:
        sheet.auto_filter = f"A6:I{len(sheet.rows)}"


def _manager_groups(events: Sequence[KsEvent]) -> dict[str, list[KsEvent]]:
    return _dimension_rows(events, "manager")


def _build_managers_workbook(
    workbook: Workbook,
    data: KsReportData,
    *,
    include_detail: bool = True,
) -> None:
    sheet = workbook.add_sheet("Менеджеры")
    _title(sheet, REPORT_TITLES["managers"], data, 10)
    sheet.append([None] * 10)
    sheet.append(
        ["Менеджер", "Отдел", "Звонки", "Лиды", "КП", "Проекты", "Конверсия", "Касса, руб.", "Возникновение, руб.", "Изменение возникновения"],
        STYLE["table_header"],
    )
    current_groups = _manager_groups(data.events)
    previous_groups = _manager_groups(data.comparison_events)
    roster_departments = {r[1]: dict(DEPARTMENTS)[r[2]] for r in _selected_roster(data)}
    for manager in sorted(set(current_groups) | set(previous_groups) | set(roster_departments)):
        rows = current_groups.get(manager, [])
        current = _metrics(rows)
        previous = _metrics(previous_groups.get(manager, []))
        department = next((row.department for row in rows if row.department), roster_departments.get(manager, "Без отдела"))
        sheet.append(
            [manager, department, current["calls"], current["leads"], current["offers"], current["projects"], current["conversion"] or "—", current["cash"], current["occurrence"], _change(Decimal(current["occurrence"]), Decimal(previous["occurrence"]))],
            [STYLE["table_text"], STYLE["table_text"], STYLE["table_number"], STYLE["table_number"], STYLE["table_number"], STYLE["table_number"], STYLE["table_percent"], STYLE["table_money"], STYLE["table_money"], STYLE["table_percent"]],
        )
    sheet.widths = {0: 28, 1: 24, 2: 12, 3: 12, 4: 12, 5: 12, 6: 15, 7: 18, 8: 22, 9: 24}
    sheet.freeze_rows = 6
    if len(sheet.rows) > 6:
        sheet.auto_filter = f"A6:J{len(sheet.rows)}"
    if include_detail:
        _build_detail_sheet(workbook, data, data.events)


def _funnel_counts(events: Sequence[KsEvent]) -> dict[str, int]:
    leads = _events(events, "Лид")
    counts = {
        "Новая": len(leads),
        "Ждём ШР": 0,
        "Думает": 0,
        "Успех": 0,
        "Отказ": 0,
    }
    waiting = {value.casefold() for value in WAITING_STATUSES}
    for row in leads:
        values = {row.status.strip().casefold(), row.stage.strip().casefold()}
        if values & waiting:
            counts["Ждём ШР"] += 1
        elif "думает" in values:
            counts["Думает"] += 1
        elif values & {"успех", "успешно"}:
            counts["Успех"] += 1
        elif "отказ" in values:
            counts["Отказ"] += 1
    return counts


def _planned_funnel(data: KsReportData) -> dict[str, int]:
    actual = _funnel_counts(data.events)
    previous_metrics = _metrics(data.comparison_events)
    projects = _events(data.events, "Проект")
    average_check = float(_sum(projects)) / len(projects) if projects else 0
    target_amount = float(previous_metrics["occurrence"])
    target_success = math.ceil(target_amount / average_check) if average_check and target_amount else actual["Успех"]
    scale = target_success / actual["Успех"] if actual["Успех"] else 1
    return {
        stage: target_success if stage == "Успех" else math.ceil(count * scale)
        for stage, count in actual.items()
    }


def _build_funnel_workbook(
    workbook: Workbook,
    data: KsReportData,
    *,
    include_detail: bool = True,
) -> None:
    actual = _funnel_counts(data.events)
    planned = _planned_funnel(data)
    sheet = workbook.add_sheet("Воронка")
    _title(sheet, REPORT_TITLES["funnel"], data, 6)
    sheet.append([None] * 6)
    sheet.append(
        ["Этап", "Факт, шт.", "Доля лидов", "Плановая воронка, шт.", "Дефицит", "Пред. период, шт."],
        STYLE["table_header"],
    )
    previous = _funnel_counts(data.comparison_events)
    for stage in ("Новая", "Ждём ШР", "Думает", "Успех", "Отказ"):
        sheet.append(
            [stage, actual[stage], _ratio(actual[stage], actual["Новая"]), planned[stage], max(planned[stage] - actual[stage], 0), previous[stage]],
            [STYLE["table_text"], STYLE["table_number"], STYLE["table_percent"], STYLE["table_number"], STYLE["table_number"], STYLE["table_number"]],
        )
    sheet.append([None] * 6)
    note = sheet.append(
        ["Воронка когортная: «Новая» — все заявки, поступившие в выбранный период, а остальные этапы — их текущие статусы. Конверсия = Успех / Новая. Плановая воронка масштабирует текущую структуру до количества успехов, необходимого для достижения ориентира предыдущего периода при текущем среднем чеке."] + [None] * 5,
        STYLE["note"],
    )
    sheet.merges.append(f"A{note}:F{note}")
    sheet.row_heights[note] = 64
    sheet.widths = {0: 23, 1: 16, 2: 17, 3: 26, 4: 16, 5: 22}
    sheet.freeze_rows = 6

    manager_sheet = workbook.add_sheet("По менеджерам")
    _title(manager_sheet, "Конверсия по менеджерам", data, 7)
    manager_sheet.append([None] * 7)
    manager_sheet.append(["Менеджер", "Лиды", "Ждём ШР", "Думает", "Успех", "Отказ", "Конверсия в успех"], STYLE["table_header"])
    for manager, rows in sorted(_manager_groups(data.events).items()):
        counts = _funnel_counts(rows)
        manager_sheet.append(
            [manager, counts["Новая"], counts["Ждём ШР"], counts["Думает"], counts["Успех"], counts["Отказ"], _ratio(counts["Успех"], counts["Новая"])],
            [STYLE["table_text"], STYLE["table_number"], STYLE["table_number"], STYLE["table_number"], STYLE["table_number"], STYLE["table_number"], STYLE["table_percent"]],
        )
    manager_sheet.widths = {0: 29, 1: 13, 2: 16, 3: 13, 4: 13, 5: 13, 6: 21}
    manager_sheet.freeze_rows = 6
    if include_detail:
        _build_detail_sheet(workbook, data, _events(data.events, "Лид"))


def _project_summary(events: Sequence[KsEvent]) -> dict[str, Decimal | int]:
    projects = _events(events, "Проект")
    payments = [row for row in _events(events, "Платёж") if row.amount > 0]
    occurrence = _sum(projects)
    paid_by_project: dict[str, Decimal] = defaultdict(Decimal)
    for payment in payments:
        paid_by_project[payment.project_id] += payment.amount

    # Cash is period-based; project balances use all recorded positive payments.
    for project in projects:
        if project.paid_amount is not None:
            paid_by_project[project.event_id] = project.paid_amount
    paid = sum((paid_by_project[row.event_id] for row in projects), Decimal(0))

    def subset_totals(subset: Sequence[KsEvent]) -> tuple[Decimal, Decimal, Decimal]:
        project_ids = {row.event_id for row in subset}
        subset_paid = sum(
            (amount for project_id, amount in paid_by_project.items() if project_id in project_ids),
            Decimal(0),
        )
        subset_awaiting = sum(
            (
                max(row.amount - paid_by_project.get(row.event_id, Decimal(0)), Decimal(0))
                for row in subset
            ),
            Decimal(0),
        )
        return _sum(subset), subset_paid, subset_awaiting

    project_amount, _, awaiting = subset_totals(projects)
    agreed = [row for row in projects if row.is_agreed]
    approval = [
        row
        for row in projects
        if not row.is_agreed and "соглас" in row.current_step.casefold()
    ]
    agreed_amount, agreed_paid, agreed_awaiting = subset_totals(agreed)
    approval_amount, approval_paid, approval_awaiting = subset_totals(approval)
    return {
        "projects": len(projects),
        "occurrence": project_amount or occurrence,
        "paid": paid,
        "awaiting": awaiting,
        "agreed": len(agreed),
        "agreed_amount": agreed_amount,
        "agreed_paid": agreed_paid,
        "agreed_awaiting": agreed_awaiting,
        "approval": len(approval),
        "approval_amount": approval_amount,
        "approval_paid": approval_paid,
        "approval_awaiting": approval_awaiting,
    }


def _build_projects_workbook(
    workbook: Workbook,
    data: KsReportData,
    *,
    include_detail: bool = True,
) -> None:
    current = _project_summary(data.events)
    previous = _project_summary(data.comparison_events)
    sheet = workbook.add_sheet("Проекты и оплаты")
    _title(sheet, REPORT_TITLES["projects"], data, 7)
    sheet.append([None] * 7)
    sheet.append(["Показатель", "Факт", "Пред. период", "Изменение", "Оплачено, руб.", "Ожидает оплаты, руб.", "Сумма проектов, руб."], STYLE["table_header"])
    sheet.append(
        ["Проекты, запущенные в периоде", current["projects"], previous["projects"], _change(int(current["projects"]), int(previous["projects"])), current["paid"], current["awaiting"], current["occurrence"]],
        [STYLE["total_text"], STYLE["total_number"], STYLE["total_number"], STYLE["total_percent"], STYLE["total_money"], STYLE["total_money"], STYLE["total_money"]],
    )
    sheet.append(
        ["Согласованные проекты", current["agreed"], previous["agreed"], _change(int(current["agreed"]), int(previous["agreed"])), current["agreed_paid"], current["agreed_awaiting"], current["agreed_amount"]],
        [STYLE["table_text"], STYLE["table_number"], STYLE["table_number"], STYLE["table_percent"], STYLE["table_money"], STYLE["table_money"], STYLE["table_money"]],
    )
    sheet.append(
        ["Проекты на согласовании", current["approval"], previous["approval"], _change(int(current["approval"]), int(previous["approval"])), current["approval_paid"], current["approval_awaiting"], current["approval_amount"]],
        [STYLE["table_text"], STYLE["table_number"], STYLE["table_number"], STYLE["table_percent"], STYLE["table_money"], STYLE["table_money"], STYLE["table_money"]],
    )
    sheet.append([None] * 7)
    note = sheet.append(
        ["Проект включён по первому назначению замерщика или эксперта в периоде. Стоимость — projects.sale_price на момент выгрузки. Оплачено здесь — все положительные неудалённые платежи этих проектов на момент выгрузки, не касса за период. Остаток — сумма max(стоимость проекта − его оплаты, 0). Возвраты исключены. Согласован — достиг «Распечатки». Проекты без подтверждённой даты назначения не включаются."] + [None] * 6,
        STYLE["note"],
    )
    sheet.merges.append(f"A{note}:G{note}")
    sheet.row_heights[note] = 78
    sheet.widths = {0: 28, 1: 15, 2: 18, 3: 16, 4: 20, 5: 24, 6: 24}
    sheet.freeze_rows = 6
    if include_detail:
        _build_detail_sheet(workbook, data, [row for row in data.events if row.kind in {"Проект", "Платёж"}])


def _build_all_workbook(workbook: Workbook, data: KsReportData) -> None:
    _build_plan_workbook(
        workbook,
        data,
        sheet_name="Общая сводка",
        include_support=False,
    )
    _build_dimension_sheet(workbook, data, "По отделам", "department")
    _build_dimension_sheet(workbook, data, "По продуктам", "service")
    _build_managers_workbook(workbook, data, include_detail=False)
    _build_funnel_workbook(workbook, data, include_detail=False)
    _build_projects_workbook(workbook, data, include_detail=False)
    _build_detail_sheet(workbook, data, data.events)


def _build_detail_sheet(
    workbook: Workbook,
    data: KsReportData,
    events: Sequence[KsEvent],
) -> None:
    sheet = workbook.add_sheet("Детализация")
    _title(sheet, "Детализация расчёта", data, 15)
    sheet.append([None] * 15)
    sheet.append(["Тип", "Дата события", "Дата заявки (когорта)", "Менеджер", "Отдел", "Продукт", "Сумма, руб.", "Номер / ID", "Статус", "Этап заявки", "Этап проекта", "ID записи", "ID проекта", "Все оплаты проекта, руб.", "Остаток проекта, руб."], STYLE["table_header"])
    for row in events:
        sheet.append(
            [row.kind, row.event_at, row.cohort_at or "—", row.manager, row.department, _service_label(row.service), row.amount, row.reference if row.reference != "—" else row.event_id, row.status, row.stage, row.current_step, row.event_id, row.project_id or "—", row.paid_amount if row.kind == "Проект" else None, max(row.amount - (row.paid_amount or Decimal(0)), Decimal(0)) if row.kind == "Проект" else None],
            [STYLE["table_text"], STYLE["table_center"], STYLE["table_center"], STYLE["table_text"], STYLE["table_text"], STYLE["table_text"], STYLE["table_money"], STYLE["table_text"], STYLE["table_text"], STYLE["table_text"], STYLE["table_text"], STYLE["table_text"], STYLE["table_text"], STYLE["table_money"], STYLE["table_money"]],
        )
    sheet.widths = {0: 14, 1: 18, 2: 23, 3: 28, 4: 24, 5: 22, 6: 18, 7: 25, 8: 22, 9: 25, 10: 31, 11: 39, 12: 39, 13: 25, 14: 25}
    sheet.freeze_rows = 6
    if events:
        sheet.auto_filter = f"A6:O{len(sheet.rows)}"


@lru_cache(maxsize=32)
def _font(size: int, bold: bool = False):
    asset_name = "ks-report-bold.ttf.b64" if bold else "ks-report-regular.ttf.b64"
    asset_path = Path(__file__).with_name("assets") / asset_name
    try:
        font_bytes = base64.b64decode(asset_path.read_text(encoding="ascii"))
        return ImageFont.truetype(io.BytesIO(font_bytes), size)
    except (OSError, ValueError):
        pass
    candidates = (
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/dejavu/DejaVuSans.ttf",
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf" if bold else "/System/Library/Fonts/Supplemental/Arial.ttf",
    )
    for path in candidates:
        try:
            return ImageFont.truetype(path, size)
        except OSError:
            continue
    return ImageFont.load_default(size=size)


def _chart_value(value: float, money: bool) -> str:
    if money:
        if abs(value) >= 1_000_000:
            return f"{value / 1_000_000:.1f} млн"
        if abs(value) >= 1_000:
            return f"{value / 1_000:.0f} тыс."
        return f"{value:.0f} руб."
    return f"{value:.0f}"


def _change_label(actual: Decimal | int | float, previous: Decimal | int | float) -> str:
    change = _change(actual, previous)
    if change is None:
        return "нет базы сравнения"
    sign = "+" if change >= 0 else ""
    return f"к сравнению {sign}{change * 100:.1f}%"


def _bar_chart(
    title: str,
    subtitle: str,
    labels: Sequence[str],
    current: Sequence[float],
    comparison: Sequence[float],
    *,
    money: bool = False,
    comparison_label: str = "Сравнение",
    cards: Sequence[tuple[str, str, str]] = (),
) -> bytes:
    image = Image.new("RGB", (1200, 1200), "#F4F7FB")
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((45, 45, 1155, 1155), radius=28, fill="#FFFFFF", outline="#DCE5EF", width=2)
    title_size = 42
    title_font = _font(title_size, True)
    while title_size > 28 and draw.textbbox((0, 0), title, font=title_font)[2] > 1020:
        title_size -= 2
        title_font = _font(title_size, True)
    draw.text((85, 78), title, font=title_font, fill="#173A5E")
    draw.text((85, 137), subtitle, font=_font(24), fill="#5B6B7C")
    if cards:
        card_width = 245
        for index, (label, value, detail) in enumerate(cards[:4]):
            left = 85 + index * 255
            draw.rounded_rectangle(
                (left, 188, left + card_width, 310),
                radius=16,
                fill="#F7F9FC",
                outline="#DCE5EF",
                width=2,
            )
            draw.text((left + 16, 204), label, font=_font(17), fill="#5B6B7C")
            draw.text((left + 16, 239), value, font=_font(26, True), fill="#173A5E")
            draw.text((left + 16, 279), detail, font=_font(14), fill="#7A8998")
        legend_top = 337
        start_y = 425
        available_height = 625
    else:
        legend_top = 188
        start_y = 300
        available_height = 780
    draw.rounded_rectangle((85, legend_top, 1105, legend_top + 64), radius=18, fill="#EAF2FA")
    draw.rectangle((115, legend_top + 22, 143, legend_top + 42), fill="#255985")
    draw.text((156, legend_top + 14), "Выбранный период", font=_font(20), fill="#213547")
    draw.rectangle((430, legend_top + 22, 458, legend_top + 42), fill="#E89A3C")
    draw.text((471, legend_top + 14), comparison_label, font=_font(20), fill="#213547")

    if not labels:
        draw.text((420, 570), "Нет данных за выбранный период", font=_font(30, True), fill="#66788A")
    else:
        maximum = max([*current, *comparison, 1.0])
        row_height = min(128, available_height // max(len(labels), 1))
        bar_left = 330
        bar_width = 680
        for index, label in enumerate(labels):
            y = start_y + index * row_height
            display = label if len(label) <= 23 else label[:22] + "…"
            draw.text((85, y + 16), display, font=_font(19, True), fill="#213547")
            current_width = int(bar_width * max(current[index], 0) / maximum)
            compare_width = int(bar_width * max(comparison[index], 0) / maximum)
            draw.rounded_rectangle((bar_left, y + 7, bar_left + current_width, y + 32), radius=8, fill="#255985")
            draw.rounded_rectangle((bar_left, y + 40, bar_left + compare_width, y + 65), radius=8, fill="#E89A3C")
            draw.text((bar_left + current_width + 10, y + 4), _chart_value(current[index], money), font=_font(17, True), fill="#173A5E")
            draw.text((bar_left + compare_width + 10, y + 37), _chart_value(comparison[index], money), font=_font(17), fill="#7C5630")
    draw.text((85, 1092), "EcoStar Reports · данные CRM", font=_font(20), fill="#7A8998")
    output = io.BytesIO()
    image.save(output, format="PNG", optimize=True)
    return output.getvalue()


def _summary_chart(data: KsReportData) -> bytes:
    current = _metrics(data.events)
    previous = _metrics(data.comparison_events)
    projects = _project_summary(data.events)
    image = Image.new("RGB", (1200, 1200), "#F4F7FB")
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle(
        (45, 45, 1155, 1155),
        radius=28,
        fill="#FFFFFF",
        outline="#DCE5EF",
        width=2,
    )
    draw.text((85, 78), "КС · Общий отчёт", font=_font(42, True), fill="#173A5E")
    draw.text(
        (85, 137),
        f"{data.date_from:%d.%m.%Y}–{data.date_to:%d.%m.%Y} · {data.filter_labels['department']}",
        font=_font(24),
        fill="#5B6B7C",
    )

    conversion = current["conversion"]
    previous_projects = _project_summary(data.comparison_events)
    cards = (
        ("Касса", _chart_value(float(current["cash"]), True), _change_label(current["cash"], previous["cash"])),
        ("Возникновение", _chart_value(float(current["occurrence"]), True), _change_label(current["occurrence"], previous["occurrence"])),
        ("Ожидает оплаты", _chart_value(float(projects["awaiting"]), True), _change_label(projects["awaiting"], previous_projects["awaiting"])),
        ("Лиды", str(current["leads"]), _change_label(current["leads"], previous["leads"])),
        ("Проекты · согласовано", f"{current['projects']} · {projects['agreed']}", "первый запуск · «Распечатка»"),
        ("Конверсия", f"{float(conversion) * 100:.1f}%" if conversion is not None else "—", "успех / новая"),
    )
    for index, (label, value, detail) in enumerate(cards):
        column = index % 3
        row = index // 3
        left = 85 + column * 340
        top = 188 + row * 112
        draw.rounded_rectangle(
            (left, top, left + 310, top + 98),
            radius=16,
            fill="#EAF2FA" if row == 0 else "#F7F9FC",
            outline="#DCE5EF",
            width=2,
        )
        draw.text((left + 18, top + 12), label, font=_font(16), fill="#5B6B7C")
        draw.text((left + 18, top + 38), value, font=_font(25, True), fill="#173A5E")
        draw.text((left + 18, top + 72), detail, font=_font(13), fill="#7A8998")

    panel_top = 425
    draw.rounded_rectangle((85, panel_top, 570, 730), radius=18, fill="#F7F9FC", outline="#DCE5EF", width=2)
    draw.text((108, panel_top + 18), "Деньги: факт и сравнение", font=_font(21, True), fill="#213547")
    draw.rectangle((108, panel_top + 56, 128, panel_top + 70), fill="#255985")
    draw.text((138, panel_top + 49), "Факт", font=_font(15), fill="#5B6B7C")
    draw.rectangle((222, panel_top + 56, 242, panel_top + 70), fill="#E89A3C")
    draw.text((252, panel_top + 49), "Сравнение", font=_font(15), fill="#5B6B7C")
    current_values = (float(current["cash"]), float(current["occurrence"]))
    previous_values = (float(previous["cash"]), float(previous["occurrence"]))
    maximum = max(*current_values, *previous_values, 1.0)
    for index, label in enumerate(("Касса", "Возникновение")):
        top = panel_top + 95 + index * 92
        draw.text((108, top), label, font=_font(16, True), fill="#213547")
        current_width = int(275 * max(current_values[index], 0) / maximum)
        previous_width = int(275 * max(previous_values[index], 0) / maximum)
        draw.rounded_rectangle((245, top, 245 + current_width, top + 19), radius=6, fill="#255985")
        draw.rounded_rectangle((245, top + 27, 245 + previous_width, top + 46), radius=6, fill="#E89A3C")
        draw.text((245, top + 52), f"{_chart_value(current_values[index], True)} / {_chart_value(previous_values[index], True)}", font=_font(14), fill="#5B6B7C")

    actual_funnel = _funnel_counts(data.events)
    funnel_values = (
        ("Лиды", int(current["leads"])),
        ("КП", int(current["offers"])),
        ("Проекты", int(current["projects"])),
        ("Успех", actual_funnel["Успех"]),
    )
    draw.rounded_rectangle((590, panel_top, 1105, 730), radius=18, fill="#F7F9FC", outline="#DCE5EF", width=2)
    draw.text((614, panel_top + 18), "Воронка за период", font=_font(21, True), fill="#213547")
    funnel_max = max((value for _, value in funnel_values), default=1) or 1
    for index, (label, value) in enumerate(funnel_values):
        top = panel_top + 68 + index * 55
        width = int(330 * value / funnel_max)
        draw.text((614, top), label, font=_font(16), fill="#213547")
        draw.rounded_rectangle((738, top + 2, 738 + width, top + 22), radius=7, fill="#255985" if index < 3 else "#2E9C73")
        draw.text((1080, top - 2), str(value), anchor="ra", font=_font(17, True), fill="#173A5E")

    manager_groups = _manager_groups(data.events)
    ranking = sorted(
        manager_groups,
        key=lambda name: float(_metrics(manager_groups[name])["occurrence"]),
        reverse=True,
    )[:3]
    draw.rounded_rectangle((85, 750, 1105, 1060), radius=18, fill="#F7F9FC", outline="#DCE5EF", width=2)
    draw.text((108, 770), "Топ-3 менеджера по возникновению", font=_font(21, True), fill="#213547")
    ranking_max = max((float(_metrics(manager_groups[name])["occurrence"]) for name in ranking), default=1.0) or 1.0
    if ranking:
        for index, name in enumerate(ranking):
            value = float(_metrics(manager_groups[name])["occurrence"])
            top = 822 + index * 70
            display = name if len(name) <= 24 else name[:23] + "…"
            draw.text((108, top), display, font=_font(17, True), fill="#213547")
            width = int(560 * value / ranking_max)
            draw.rounded_rectangle((430, top + 2, 430 + width, top + 25), radius=7, fill="#255985")
            draw.text((1010, top - 2), _chart_value(value, True), font=_font(17, True), fill="#173A5E")
    else:
        draw.text((410, 885), "Нет данных по менеджерам", font=_font(18), fill="#7A8998")

    draw.text((85, 1092), "EcoStar Reports · данные CRM", font=_font(20), fill="#7A8998")
    output = io.BytesIO()
    image.save(output, format="PNG", optimize=True)
    return output.getvalue()


def build_ks_chart(data: KsReportData) -> bytes:
    subtitle = f"{data.date_from:%d.%m.%Y}–{data.date_to:%d.%m.%Y} · {data.filter_labels['department']}"
    if data.report_kind == "all":
        return _summary_chart(data)
    if data.report_kind == "plan":
        current = _metrics(data.events)
        previous = _metrics(data.comparison_events)
        return _bar_chart(
            "КС · План-факт",
            subtitle,
            ("Касса", "Возникновение"),
            (float(current["cash"]), float(current["occurrence"])),
            (float(previous["cash"]), float(previous["occurrence"])),
            money=True,
            comparison_label="План-ориентир",
            cards=(
                ("Касса", _chart_value(float(current["cash"]), True), _change_label(current["cash"], previous["cash"])),
                ("Возникновение", _chart_value(float(current["occurrence"]), True), _change_label(current["occurrence"], previous["occurrence"])),
                ("Лиды", str(current["leads"]), _change_label(current["leads"], previous["leads"])),
                ("Конверсия", f"{float(current['conversion']) * 100:.1f}%" if current["conversion"] is not None else "—", "успех / новая"),
            ),
        )
    if data.report_kind == "managers":
        current = _metrics(data.events)
        previous = _metrics(data.comparison_events)
        current_groups = _manager_groups(data.events)
        previous_groups = _manager_groups(data.comparison_events)
        ranking = sorted(
            current_groups,
            key=lambda name: float(_metrics(current_groups[name])["occurrence"]),
            reverse=True,
        )[:8]
        return _bar_chart(
            "КС · Менеджеры по возникновению",
            subtitle,
            ranking,
            [float(_metrics(current_groups[name])["occurrence"]) for name in ranking],
            [float(_metrics(previous_groups.get(name, []))["occurrence"]) for name in ranking],
            money=True,
            cards=(
                ("Менеджеры", str(len(current_groups)), "с активностью за период"),
                ("Звонки", str(current["calls"]), _change_label(current["calls"], previous["calls"])),
                ("Лиды", str(current["leads"]), _change_label(current["leads"], previous["leads"])),
                ("Проекты", str(current["projects"]), _change_label(current["projects"], previous["projects"])),
            ),
        )
    if data.report_kind == "funnel":
        current = _metrics(data.events)
        previous = _metrics(data.comparison_events)
        actual = _funnel_counts(data.events)
        planned = _planned_funnel(data)
        labels = ("Новая", "Ждём ШР", "Думает", "Успех")
        return _bar_chart(
            "КС · Фактическая и плановая воронка",
            subtitle,
            labels,
            [float(actual[label]) for label in labels],
            [float(planned[label]) for label in labels],
            comparison_label="Необходимо",
            cards=(
                ("Лиды", str(current["leads"]), _change_label(current["leads"], previous["leads"])),
                ("КП", str(current["offers"]), _change_label(current["offers"], previous["offers"])),
                ("Успех", str(actual["Успех"]), "успешные заявки"),
                ("Конверсия", f"{float(current['conversion']) * 100:.1f}%" if current["conversion"] is not None else "—", "успех / новая"),
            ),
        )
    current = _project_summary(data.events)
    previous = _project_summary(data.comparison_events)
    return _bar_chart(
        "КС · Проекты и оплаты",
        subtitle,
        ("Сумма проектов", "Оплачено", "Ожидает оплаты"),
        [float(current["occurrence"]), float(current["paid"]), float(current["awaiting"])],
        [float(previous["occurrence"]), float(previous["paid"]), float(previous["awaiting"])],
        money=True,
        cards=(
            ("Проекты · согласовано", f"{current['projects']} · {current['agreed']}", "первый запуск · «Распечатка»"),
            ("Сумма проектов", _chart_value(float(current["occurrence"]), True), _change_label(current["occurrence"], previous["occurrence"])),
            ("Оплачено", _chart_value(float(current["paid"]), True), _change_label(current["paid"], previous["paid"])),
            ("Ожидает оплаты", _chart_value(float(current["awaiting"]), True), _change_label(current["awaiting"], previous["awaiting"])),
        ),
    )


def _artifact_caption(data: KsReportData) -> str:
    return (
        f"<b>{REPORT_TITLES[data.report_kind]}</b> · "
        f"{data.date_from:%d.%m.%Y}–{data.date_to:%d.%m.%Y}\n"
        f"{data.filter_labels['department']} · {data.filter_labels['manager']} · "
        f"{data.filter_labels['product']}"
    )
