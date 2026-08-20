from __future__ import annotations

import io
import math
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal
from typing import Iterable, Mapping, Sequence
from zoneinfo import ZoneInfo

from PIL import Image, ImageDraw, ImageFont
from sqlalchemy import create_engine, text
from sqlalchemy.pool import NullPool

from reporting_bot.config import Settings
from reporting_bot.simple_xlsx import STYLE, Workbook


REPORT_TITLES = {
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
    paid_amount: Decimal = Decimal(0)


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


DEPARTMENT_OPTIONS_SQL = """
SELECT left(replace(d.id::text, '-', ''), 6) AS token, d.name AS label
FROM departments d
WHERE d.group_id = 1
ORDER BY d.name
"""


MANAGER_OPTIONS_SQL = """
SELECT DISTINCT
  left(replace(u.id::text, '-', ''), 6) AS token,
  concat_ws(' ', u.last_name, u.first_name) AS label
FROM users u
LEFT JOIN department_members dm ON dm.member_id = u.id
LEFT JOIN departments d ON d.id = dm.department_id AND d.group_id = 1
WHERE u.is_active
  AND u.group_id = 1
  AND (
    :department_token = '-'
    OR left(replace(coalesce(d.id::text, ''), '-', ''), 6) = :department_token
  )
ORDER BY label
"""


PRODUCT_OPTIONS_SQL = """
SELECT left(md5(service), 6) AS token, service AS label
FROM (
  SELECT lower(service) AS service FROM requests_clone
  WHERE group_id = 1 AND service IS NOT NULL AND service <> ''
  UNION
  SELECT lower(service) AS service FROM projects
  WHERE group_id = 1 AND service IS NOT NULL AND service <> ''
) services
ORDER BY service
"""


REQUESTS_SQL = """
SELECT
  r.id::text AS event_id,
  r.created_at AS event_at,
  concat_ws(' ', u.last_name, u.first_name) AS manager,
  coalesce(d.name, 'Без отдела') AS department,
  coalesce(r.service, '') AS service,
  coalesce(r.price, 0) AS amount,
  coalesce(r.status, '—') AS status,
  coalesce(r.stage, '—') AS stage
FROM requests_clone r
LEFT JOIN users u ON u.id = r.manager_id
LEFT JOIN LATERAL (
  SELECT dep.id, dep.name
  FROM department_members dm
  JOIN departments dep ON dep.id = dm.department_id AND dep.group_id = 1
  WHERE dm.member_id = r.manager_id
  ORDER BY dep.name
  LIMIT 1
) d ON true
WHERE r.group_id = 1
  AND r.created_at >= CAST(:date_from AS DATE)
  AND r.created_at < CAST(:date_to_exclusive AS DATE)
  AND r.deleted_at IS NULL
  AND r.archived_at IS NULL
  AND coalesce(r.status, '') NOT IN ('Не лид', 'Дубль')
  AND (:department_token = '-' OR left(replace(coalesce(d.id::text, ''), '-', ''), 6) = :department_token)
  AND (:manager_token = '-' OR left(replace(coalesce(r.manager_id::text, ''), '-', ''), 6) = :manager_token)
  AND (:service = '-' OR left(md5(lower(coalesce(r.service, ''))), 6) = :service)
ORDER BY r.created_at, r.id
LIMIT 10000
"""


OFFERS_SQL = """
SELECT DISTINCT ON (o.request_id)
  o.request_id::text AS event_id,
  o.created_at AS event_at,
  concat_ws(' ', u.last_name, u.first_name) AS manager,
  coalesce(d.name, 'Без отдела') AS department,
  coalesce(nullif(o.service, ''), r.service, '') AS service,
  coalesce(r.price, 0) AS amount,
  coalesce(o.number, '—') AS reference
FROM offers o
JOIN requests_clone r ON r.id = o.request_id AND r.group_id = 1
LEFT JOIN users u ON u.id = coalesce(o.user_id, r.manager_id)
LEFT JOIN LATERAL (
  SELECT dep.id, dep.name
  FROM department_members dm
  JOIN departments dep ON dep.id = dm.department_id AND dep.group_id = 1
  WHERE dm.member_id = coalesce(o.user_id, r.manager_id)
  ORDER BY dep.name
  LIMIT 1
) d ON true
WHERE o.is_sent IS TRUE
  AND o.created_at >= CAST(:date_from AS DATE)
  AND o.created_at < CAST(:date_to_exclusive AS DATE)
  AND (:department_token = '-' OR left(replace(coalesce(d.id::text, ''), '-', ''), 6) = :department_token)
  AND (:manager_token = '-' OR left(replace(coalesce(coalesce(o.user_id, r.manager_id)::text, ''), '-', ''), 6) = :manager_token)
  AND (:service = '-' OR left(md5(lower(coalesce(nullif(o.service, ''), r.service, ''))), 6) = :service)
ORDER BY o.request_id, o.created_at, o.id
LIMIT 10000
"""


PROJECTS_SQL = """
SELECT
  p.id::text AS event_id,
  p.created_at AS event_at,
  concat_ws(' ', u.last_name, u.first_name) AS manager,
  coalesce(d.name, 'Без отдела') AS department,
  coalesce(p.service, '') AS service,
  coalesce(p.sale_price, 0) AS amount,
  coalesce(p.contract_number, '—') AS reference,
  coalesce(p.current_step, '—') AS current_step,
  coalesce(p.paid_amount, 0) AS paid_amount
FROM projects p
LEFT JOIN users u ON u.id = p.manager_id
LEFT JOIN LATERAL (
  SELECT dep.id, dep.name
  FROM department_members dm
  JOIN departments dep ON dep.id = dm.department_id AND dep.group_id = 1
  WHERE dm.member_id = p.manager_id
  ORDER BY dep.name
  LIMIT 1
) d ON true
WHERE p.group_id = 1
  AND p.created_at >= CAST(:date_from AS DATE)
  AND p.created_at < CAST(:date_to_exclusive AS DATE)
  AND (:department_token = '-' OR left(replace(coalesce(d.id::text, ''), '-', ''), 6) = :department_token)
  AND (:manager_token = '-' OR left(replace(coalesce(p.manager_id::text, ''), '-', ''), 6) = :manager_token)
  AND (:service = '-' OR left(md5(lower(coalesce(p.service, ''))), 6) = :service)
ORDER BY p.created_at, p.id
LIMIT 10000
"""


PAYMENTS_SQL = """
SELECT
  pay.id::text AS event_id,
  pay."chargeAt" AS event_at,
  concat_ws(' ', u.last_name, u.first_name) AS manager,
  coalesce(d.name, 'Без отдела') AS department,
  coalesce(p.service, '') AS service,
  coalesce(pay.amount, 0) AS amount,
  coalesce(pay."contractNumber", p.contract_number, '—') AS reference,
  coalesce(p.current_step, '—') AS current_step
FROM payment pay
LEFT JOIN projects p ON p.id = pay."projectId"
LEFT JOIN users u ON u.id = coalesce(p.manager_id, pay."userId")
LEFT JOIN LATERAL (
  SELECT dep.id, dep.name
  FROM department_members dm
  JOIN departments dep ON dep.id = dm.department_id AND dep.group_id = 1
  WHERE dm.member_id = coalesce(p.manager_id, pay."userId")
  ORDER BY dep.name
  LIMIT 1
) d ON true
WHERE pay.group_id = 1
  AND pay."deletedAt" IS NULL
  AND pay."chargeAt" >= CAST(:date_from AS DATE)
  AND pay."chargeAt" < CAST(:date_to_exclusive AS DATE)
  AND (:department_token = '-' OR left(replace(coalesce(d.id::text, ''), '-', ''), 6) = :department_token)
  AND (:manager_token = '-' OR left(replace(coalesce(coalesce(p.manager_id, pay."userId")::text, ''), '-', ''), 6) = :manager_token)
  AND (:service = '-' OR left(md5(lower(coalesce(p.service, ''))), 6) = :service)
ORDER BY pay."chargeAt", pay.id
LIMIT 10000
"""


CALLS_SQL = """
SELECT
  c.id::text AS event_id,
  c.call_date AS event_at,
  concat_ws(' ', u.last_name, u.first_name) AS manager,
  coalesce(d.name, 'Без отдела') AS department,
  '' AS service,
  0 AS amount,
  coalesce(c.direction::text, '—') AS reference
FROM calls c
JOIN users u ON u.id = c.user_id AND u.group_id = 1
LEFT JOIN LATERAL (
  SELECT dep.id, dep.name
  FROM department_members dm
  JOIN departments dep ON dep.id = dm.department_id AND dep.group_id = 1
  WHERE dm.member_id = c.user_id
  ORDER BY dep.name
  LIMIT 1
) d ON true
WHERE c.call_date >= CAST(:date_from AS DATE)
  AND c.call_date < CAST(:date_to_exclusive AS DATE)
  AND (:department_token = '-' OR left(replace(coalesce(d.id::text, ''), '-', ''), 6) = :department_token)
  AND (:manager_token = '-' OR left(replace(c.user_id::text, '-', ''), 6) = :manager_token)
  AND :service = '-'
ORDER BY c.call_date, c.id
LIMIT 10000
"""


class KsReportService:
    def __init__(self, settings: Settings) -> None:
        self._settings = settings
        self._database_url = settings.database_url.replace(
            "postgresql://", "postgresql+psycopg://", 1
        )

    def filter_options(self, department_token: str = "-") -> KsFilterOptions:
        engine = self._engine()
        try:
            with engine.connect() as connection:
                with connection.begin():
                    self._read_only(connection)
                    departments = tuple(
                        FilterOption(str(row["token"]), str(row["label"]))
                        for row in connection.execute(text(DEPARTMENT_OPTIONS_SQL)).mappings()
                    )
                    managers = (
                        tuple(
                            FilterOption(str(row["token"]), str(row["label"] or "Без имени"))
                            for row in connection.execute(
                                text(MANAGER_OPTIONS_SQL),
                                {"department_token": department_token},
                            ).mappings()
                        )
                        if department_token != "-"
                        else ()
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
        connection.execute(text("SET TRANSACTION READ ONLY"))
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
                    events = _load_events(connection, date_from, date_to, filters)
                    comparison_events = (
                        _load_events(connection, comparison_from, comparison_to, filters)
                        if comparison_from is not None and comparison_to is not None
                        else ()
                    )
                    options = _load_filter_options(connection, filters.department_token)
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
    managers = (
        tuple(
            FilterOption(str(row["token"]), str(row["label"] or "Без имени"))
            for row in connection.execute(
                text(MANAGER_OPTIONS_SQL), {"department_token": department_token}
            ).mappings()
        )
        if department_token != "-"
        else ()
    )
    products = tuple(
        FilterOption(str(row["token"]), _service_label(str(row["label"])))
        for row in connection.execute(text(PRODUCT_OPTIONS_SQL)).mappings()
    )
    return KsFilterOptions(departments, managers, products)


def _option_label(
    options: Sequence[FilterOption], token: str, default: str
) -> str:
    if token == "-":
        return default
    return next((option.label for option in options if option.token == token), token)


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
        for row in connection.execute(text(sql), params).mappings():
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
    normalized = value.lower()
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
        "conversion": _ratio(len(projects), leads),
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
    if data.report_kind == "plan":
        _build_plan_workbook(workbook, data)
    elif data.report_kind == "managers":
        _build_managers_workbook(workbook, data)
    elif data.report_kind == "funnel":
        _build_funnel_workbook(workbook, data)
    else:
        _build_projects_workbook(workbook, data)
    return workbook.to_bytes()


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


def _build_plan_workbook(workbook: Workbook, data: KsReportData) -> None:
    current = _metrics(data.events)
    previous = _metrics(data.comparison_events)
    sheet = workbook.add_sheet("Сводка")
    _title(sheet, REPORT_TITLES["plan"], data, 8)
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
        ["План-ориентир временно равен факту предыдущего аналогичного периода. Прогноз рассчитывается по рабочим дням; для завершённого периода прогноз равен факту."] + [None] * 7,
        STYLE["note"],
    )
    sheet.merges.append(f"A{note}:H{note}")
    sheet.row_heights[note] = 38
    sheet.widths = {0: 27, 1: 20, 2: 18, 3: 18, 4: 18, 5: 18, 6: 16, 7: 12}
    sheet.freeze_rows = 6

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
    keys = sorted(set(current_groups) | set(previous_groups))
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


def _build_managers_workbook(workbook: Workbook, data: KsReportData) -> None:
    sheet = workbook.add_sheet("Менеджеры")
    _title(sheet, REPORT_TITLES["managers"], data, 10)
    sheet.append([None] * 10)
    sheet.append(
        ["Менеджер", "Отдел", "Звонки", "Лиды", "КП", "Проекты", "Конверсия", "Касса, руб.", "Возникновение, руб.", "Изменение возникновения"],
        STYLE["table_header"],
    )
    current_groups = _manager_groups(data.events)
    previous_groups = _manager_groups(data.comparison_events)
    for manager in sorted(set(current_groups) | set(previous_groups)):
        rows = current_groups.get(manager, [])
        current = _metrics(rows)
        previous = _metrics(previous_groups.get(manager, []))
        department = next((row.department for row in rows if row.department), "Без отдела")
        sheet.append(
            [manager, department, current["calls"], current["leads"], current["offers"], current["projects"], current["conversion"] or "—", current["cash"], current["occurrence"], _change(Decimal(current["occurrence"]), Decimal(previous["occurrence"]))],
            [STYLE["table_text"], STYLE["table_text"], STYLE["table_number"], STYLE["table_number"], STYLE["table_number"], STYLE["table_number"], STYLE["table_percent"], STYLE["table_money"], STYLE["table_money"], STYLE["table_percent"]],
        )
    sheet.widths = {0: 28, 1: 24, 2: 12, 3: 12, 4: 12, 5: 12, 6: 15, 7: 18, 8: 22, 9: 24}
    sheet.freeze_rows = 6
    if len(sheet.rows) > 6:
        sheet.auto_filter = f"A6:J{len(sheet.rows)}"
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
    for row in leads:
        if row.status in WAITING_STATUSES or row.stage in WAITING_STATUSES:
            counts["Ждём ШР"] += 1
        elif row.status == "Думает" or row.stage == "Думает":
            counts["Думает"] += 1
        elif row.status == "Успешно":
            counts["Успех"] += 1
        elif row.status == "Отказ":
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


def _build_funnel_workbook(workbook: Workbook, data: KsReportData) -> None:
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
        ["Плановая воронка масштабирует текущую структуру до количества успешных сделок, необходимого для достижения ориентира предыдущего периода при текущем среднем чеке."] + [None] * 5,
        STYLE["note"],
    )
    sheet.merges.append(f"A{note}:F{note}")
    sheet.row_heights[note] = 42
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
    _build_detail_sheet(workbook, data, _events(data.events, "Лид"))


def _project_summary(events: Sequence[KsEvent]) -> dict[str, Decimal | int]:
    projects = _events(events, "Проект")
    payments = _events(events, "Платёж")
    occurrence = _sum(projects)
    paid = _sum(payments)
    awaiting = sum((max(row.amount - row.paid_amount, Decimal(0)) for row in projects), Decimal(0))
    approval = [row for row in projects if "соглас" in row.current_step.lower()]
    return {
        "projects": len(projects),
        "occurrence": occurrence,
        "paid": paid,
        "awaiting": awaiting,
        "approval": len(approval),
        "approval_amount": _sum(approval),
    }


def _build_projects_workbook(workbook: Workbook, data: KsReportData) -> None:
    current = _project_summary(data.events)
    previous = _project_summary(data.comparison_events)
    sheet = workbook.add_sheet("Проекты и оплаты")
    _title(sheet, REPORT_TITLES["projects"], data, 7)
    sheet.append([None] * 7)
    sheet.append(["Показатель", "Факт", "Пред. период", "Изменение", "Оплачено, руб.", "Ожидает оплаты, руб.", "Сумма проектов, руб."], STYLE["table_header"])
    sheet.append(
        ["Созданные проекты", current["projects"], previous["projects"], _change(int(current["projects"]), int(previous["projects"])), current["paid"], current["awaiting"], current["occurrence"]],
        [STYLE["total_text"], STYLE["total_number"], STYLE["total_number"], STYLE["total_percent"], STYLE["total_money"], STYLE["total_money"], STYLE["total_money"]],
    )
    sheet.append(
        ["Проекты на согласовании", current["approval"], previous["approval"], _change(int(current["approval"]), int(previous["approval"])), "—", "—", current["approval_amount"]],
        [STYLE["table_text"], STYLE["table_number"], STYLE["table_number"], STYLE["table_percent"], STYLE["table_money"], STYLE["table_money"], STYLE["table_money"]],
    )
    sheet.append([None] * 7)
    note = sheet.append(
        ["Созданные проекты отбираются по projects.created_at. Оплата — платежи payment по chargeAt без удалённых записей. Ожидаемая сумма = стоимость проекта минус projects.paid_amount, не ниже нуля."] + [None] * 6,
        STYLE["note"],
    )
    sheet.merges.append(f"A{note}:G{note}")
    sheet.row_heights[note] = 44
    sheet.widths = {0: 28, 1: 15, 2: 18, 3: 16, 4: 20, 5: 24, 6: 24}
    sheet.freeze_rows = 6
    _build_detail_sheet(workbook, data, [row for row in data.events if row.kind in {"Проект", "Платёж"}])


def _build_detail_sheet(
    workbook: Workbook,
    data: KsReportData,
    events: Sequence[KsEvent],
) -> None:
    sheet = workbook.add_sheet("Детализация")
    _title(sheet, "Детализация расчёта", data, 10)
    sheet.append([None] * 10)
    sheet.append(["Тип", "Дата", "Менеджер", "Отдел", "Продукт", "Сумма, руб.", "Номер / ID", "Статус", "Этап заявки", "Этап проекта"], STYLE["table_header"])
    for row in events:
        sheet.append(
            [row.kind, row.event_at, row.manager, row.department, _service_label(row.service), row.amount, row.reference if row.reference != "—" else row.event_id, row.status, row.stage, row.current_step],
            [STYLE["table_text"], STYLE["table_center"], STYLE["table_text"], STYLE["table_text"], STYLE["table_text"], STYLE["table_money"], STYLE["table_text"], STYLE["table_text"], STYLE["table_text"], STYLE["table_text"]],
        )
    sheet.widths = {0: 14, 1: 18, 2: 28, 3: 24, 4: 22, 5: 18, 6: 25, 7: 22, 8: 25, 9: 31}
    sheet.freeze_rows = 6
    if events:
        sheet.auto_filter = f"A6:J{len(sheet.rows)}"


def _font(size: int, bold: bool = False):
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
    return ImageFont.load_default()


def _chart_value(value: float, money: bool) -> str:
    if money:
        if abs(value) >= 1_000_000:
            return f"{value / 1_000_000:.1f} млн"
        if abs(value) >= 1_000:
            return f"{value / 1_000:.0f} тыс."
        return f"{value:.0f} руб."
    return f"{value:.0f}"


def _bar_chart(
    title: str,
    subtitle: str,
    labels: Sequence[str],
    current: Sequence[float],
    comparison: Sequence[float],
    *,
    money: bool = False,
    comparison_label: str = "Сравнение",
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
    draw.rounded_rectangle((85, 188, 1105, 252), radius=18, fill="#EAF2FA")
    draw.rectangle((115, 210, 143, 230), fill="#255985")
    draw.text((156, 202), "Выбранный период", font=_font(22), fill="#213547")
    draw.rectangle((430, 210, 458, 230), fill="#E89A3C")
    draw.text((471, 202), comparison_label, font=_font(22), fill="#213547")

    if not labels:
        draw.text((420, 570), "Нет данных за выбранный период", font=_font(30, True), fill="#66788A")
    else:
        maximum = max([*current, *comparison, 1.0])
        start_y = 300
        row_height = min(150, 780 // max(len(labels), 1))
        label_width = 280
        bar_left = 365
        bar_width = 650
        for index, label in enumerate(labels):
            y = start_y + index * row_height
            display = label if len(label) <= 23 else label[:22] + "…"
            draw.text((85, y + 18), display, font=_font(23, True), fill="#213547")
            current_width = int(bar_width * max(current[index], 0) / maximum)
            compare_width = int(bar_width * max(comparison[index], 0) / maximum)
            draw.rounded_rectangle((bar_left, y + 8, bar_left + current_width, y + 38), radius=9, fill="#255985")
            draw.rounded_rectangle((bar_left, y + 48, bar_left + compare_width, y + 78), radius=9, fill="#E89A3C")
            draw.text((bar_left + current_width + 12, y + 7), _chart_value(current[index], money), font=_font(20, True), fill="#173A5E")
            draw.text((bar_left + compare_width + 12, y + 47), _chart_value(comparison[index], money), font=_font(20), fill="#7C5630")
    draw.text((85, 1092), "EcoStar Reports · данные CRM", font=_font(20), fill="#7A8998")
    output = io.BytesIO()
    image.save(output, format="PNG", optimize=True)
    return output.getvalue()


def build_ks_chart(data: KsReportData) -> bytes:
    subtitle = f"{data.date_from:%d.%m.%Y}–{data.date_to:%d.%m.%Y} · {data.filter_labels['department']}"
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
        )
    if data.report_kind == "managers":
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
        )
    if data.report_kind == "funnel":
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
    )


def _artifact_caption(data: KsReportData) -> str:
    return (
        f"<b>{REPORT_TITLES[data.report_kind]}</b> · "
        f"{data.date_from:%d.%m.%Y}–{data.date_to:%d.%m.%Y}\n"
        f"{data.filter_labels['department']} · {data.filter_labels['manager']} · "
        f"{data.filter_labels['product']}"
    )
