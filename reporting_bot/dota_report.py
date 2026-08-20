from __future__ import annotations

import calendar
import json
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from decimal import Decimal
from typing import Iterable, Mapping, Sequence

from sqlalchemy import create_engine, text
from sqlalchemy.pool import NullPool

from reporting_bot.config import Settings
from reporting_bot.simple_xlsx import STYLE, Workbook


CATEGORIES = (
    "ОПР",
    "Аутсорсинг",
    "СУОТ",
    "Обучение",
    "Аудит",
    "Другие услуги",
)

LAUNCH_PLANS = {
    "ОПР": Decimal("3000000"),
    "Аутсорсинг": Decimal("3400000"),
    "СУОТ": Decimal("300000"),
    "Обучение": Decimal("3600000"),
    "Аудит": Decimal("0"),
    "Другие услуги": Decimal("1100000"),
}

RELEASE_PLANS = {
    "ОПР": Decimal("2500000"),
    "Аутсорсинг": Decimal("3400000"),
    "СУОТ": Decimal("200000"),
    "Обучение": Decimal("3200000"),
    "Аудит": Decimal("0"),
    "Другие услуги": Decimal("350000"),
}

DEFAULT_MONTHLY_PLANS = {
    "2026-08": {
        "Запуск": LAUNCH_PLANS,
        "Выпуск": RELEASE_PLANS,
    }
}


DOTA_EVENTS_SQL = """
WITH classified AS (
  SELECT
    h.id,
    h.project_id,
    h.created_at AS event_at,
    h.previous_step,
    h.new_step,
    p.contract_number,
    p.sale_price,
    p.workplace_count,
    p.real_workplace_count,
    coalesce(
      to_jsonb(company_row)->>'name',
      to_jsonb(company_row)->>'short_name',
      to_jsonb(company_row)->>'full_name',
      '—'
    ) AS company_name,
    concat_ws(' ', manager.last_name, manager.first_name) AS manager,
    CASE
      WHEN manager.group_id = 2 THEN 'ГТО'
      WHEN manager.role IN ('partnerManager', 'agent') THEN 'ОАП'
      WHEN manager.role = 'corpManager' THEN 'КАМ'
      WHEN trim(manager.last_name) IN ('Бурдейная', 'Шергина') THEN 'КАМ'
      WHEN trim(manager.last_name) = 'Васильева' AND trim(manager.first_name) = 'Татьяна' THEN 'КАМ'
      WHEN manager.id IS NOT NULL THEN 'ОП'
      ELSE '—'
    END AS manager_department,
    concat_ws(' ', owner_user.last_name, owner_user.first_name) AS event_owner,
    CASE
      WHEN p.service IN ('opk', 'opk_eth') THEN 'ОПР'
      WHEN p.service = 'autsorsing' THEN 'Аутсорсинг'
      WHEN p.service = 'suot' THEN 'СУОТ'
      WHEN p.service = 'obuchenie' THEN 'Обучение'
      WHEN p.service = 'audit' THEN 'Аудит'
      WHEN p.service = 'other' THEN 'Другие услуги'
    END AS category,
    CASE
      WHEN p.service IN ('opk', 'opk_eth', 'autsorsing', 'suot', 'audit')
        AND h.new_step = 'Принять проект'
        THEN 'Запуск'
      WHEN p.service = 'obuchenie'
        AND h.new_step = 'Передать заявку в УЦ'
        THEN 'Запуск'
      WHEN p.service = 'other'
        AND h.new_step = 'Передать специалисту исходные данные'
        THEN 'Запуск'
      WHEN p.service IN ('opk', 'opk_eth', 'suot')
        AND h.new_step = 'Согласовать отчет у клиента'
        THEN 'Выпуск'
      WHEN p.service = 'obuchenie'
        AND h.new_step = 'Выдать удостоверение Заказчику'
        THEN 'Выпуск'
      WHEN p.service IN ('autsorsing', 'audit')
        AND h.new_step = 'Подготовить документы'
        THEN 'Выпуск'
      WHEN p.service = 'other'
        AND h.new_step = 'Передать документы Заказчику'
        THEN 'Выпуск'
    END AS metric
  FROM projects_steps_history h
  JOIN projects p ON p.id = h.project_id
  LEFT JOIN companies_clone company_row ON company_row.id = p.company_id
  LEFT JOIN users manager ON manager.id = p.manager_id
  LEFT JOIN users owner_user ON owner_user.id = h.owner_id
  WHERE h.created_at >= CAST(:date_from AS DATE)
    AND h.created_at < CAST(:date_to_exclusive AS DATE)
    AND p.group_id IN (1, 2)
    AND p.service IN ('opk', 'opk_eth', 'autsorsing', 'suot', 'obuchenie', 'audit', 'other')
), ranked AS (
  SELECT *, row_number() OVER (
    PARTITION BY project_id, metric
    ORDER BY event_at, id
  ) AS event_rank
  FROM classified
  WHERE metric IS NOT NULL
)
SELECT
  metric,
  category,
  project_id::text AS project_id,
  contract_number,
  company_name,
  manager,
  manager_department,
  event_owner,
  event_at,
  previous_step,
  new_step,
  coalesce(sale_price, 0) AS sale_price,
  CASE
    WHEN metric = 'Выпуск' THEN coalesce(real_workplace_count, 0)
    ELSE coalesce(workplace_count, 0)
  END AS workplace_count
FROM ranked
WHERE event_rank = 1
ORDER BY event_at, category, project_id
"""


@dataclass(frozen=True)
class DotaEvent:
    metric: str
    category: str
    project_id: str
    contract_number: str | None
    company_name: str
    manager: str
    manager_department: str
    event_owner: str
    event_at: datetime
    previous_step: str
    new_step: str
    amount: Decimal
    workplace_count: int


@dataclass(frozen=True)
class DotaReportData:
    date_from: date
    date_to: date
    previous_date_from: date
    previous_date_to: date
    events: tuple[DotaEvent, ...]
    previous_events: tuple[DotaEvent, ...]
    plans: Mapping[str, Mapping[str, Decimal]] = field(
        default_factory=lambda: DEFAULT_MONTHLY_PLANS["2026-08"]
    )
    plan_month: str | None = "2026-08"


class DotaReportService:
    def __init__(self, settings: Settings) -> None:
        self._settings = settings
        self._database_url = settings.database_url.replace(
            "postgresql://", "postgresql+psycopg://", 1
        )
        self._monthly_plans = _parse_monthly_plans(settings.dota_plans_json)

    def create(self, date_from: date, date_to: date) -> tuple[bytes, str]:
        if date_to < date_from:
            raise ValueError("Дата окончания не может быть раньше даты начала.")
        if (date_to - date_from).days > 366:
            raise ValueError("Максимальный период отчёта — 366 дней.")
        previous_date_from = _previous_month_date(date_from)
        previous_date_to = _previous_month_date(date_to)
        data = self._load(date_from, date_to, previous_date_from, previous_date_to)
        filename = f"Отчет_ДОТ_{date_from.isoformat()}_{date_to.isoformat()}.xlsx"
        return build_dota_workbook(data), filename

    def _load(
        self,
        date_from: date,
        date_to: date,
        previous_date_from: date,
        previous_date_to: date,
    ) -> DotaReportData:
        engine = create_engine(
            self._database_url,
            poolclass=NullPool,
            connect_args={"connect_timeout": self._settings.connect_timeout_seconds},
        )
        try:
            with engine.connect() as connection:
                with connection.begin():
                    connection.execute(text("SET TRANSACTION READ ONLY"))
                    connection.execute(
                        text("SELECT set_config('statement_timeout', :timeout, true)"),
                        {"timeout": f"{max(self._settings.statement_timeout_ms, 30_000)}ms"},
                    )
                    events = _load_events(connection, date_from, date_to)
                    previous_events = _load_events(
                        connection,
                        previous_date_from,
                        previous_date_to,
                    )
            return DotaReportData(
                date_from=date_from,
                date_to=date_to,
                previous_date_from=previous_date_from,
                previous_date_to=previous_date_to,
                events=events,
                previous_events=previous_events,
                plans=_plans_for_period(self._monthly_plans, date_from, date_to),
                plan_month=(
                    date_from.strftime("%Y-%m")
                    if (date_from.year, date_from.month) == (date_to.year, date_to.month)
                    else None
                ),
            )
        finally:
            engine.dispose()


def _load_events(connection, date_from: date, date_to: date) -> tuple[DotaEvent, ...]:
    params = {
        "date_from": date_from.isoformat(),
        "date_to_exclusive": (date_to + timedelta(days=1)).isoformat(),
    }
    return tuple(
        DotaEvent(
            metric=str(row["metric"]),
            category=str(row["category"]),
            project_id=str(row["project_id"]),
            contract_number=row["contract_number"],
            company_name=str(row["company_name"] or "—"),
            manager=str(row["manager"] or "—"),
            manager_department=str(row["manager_department"] or "—"),
            event_owner=str(row["event_owner"] or "—"),
            event_at=row["event_at"],
            previous_step=str(row["previous_step"]),
            new_step=str(row["new_step"]),
            amount=Decimal(str(row["sale_price"] or 0)),
            workplace_count=int(row["workplace_count"] or 0),
        )
        for row in connection.execute(text(DOTA_EVENTS_SQL), params).mappings()
    )


def _previous_month_date(value: date) -> date:
    year = value.year
    month = value.month - 1
    if month == 0:
        year -= 1
        month = 12
    day = min(value.day, calendar.monthrange(year, month)[1])
    return date(year, month, day)


def _parse_monthly_plans(
    raw: str,
) -> dict[str, dict[str, dict[str, Decimal]]]:
    result = {
        month: {
            metric: {category: Decimal(value) for category, value in values.items()}
            for metric, values in metrics.items()
        }
        for month, metrics in DEFAULT_MONTHLY_PLANS.items()
    }
    if not raw:
        return result
    try:
        parsed = json.loads(raw)
        result.update(
            {
                str(month): {
                    str(metric): {
                        str(category): Decimal(str(value))
                        for category, value in values.items()
                    }
                    for metric, values in metrics.items()
                }
                for month, metrics in parsed.items()
            }
        )
        return result
    except (json.JSONDecodeError, AttributeError, TypeError, ValueError) as exc:
        raise ValueError("DOTA_PLANS_JSON должен содержать корректный JSON с месячными планами.") from exc


def _plans_for_period(
    monthly_plans: Mapping[str, Mapping[str, Mapping[str, Decimal]]],
    date_from: date,
    date_to: date,
) -> Mapping[str, Mapping[str, Decimal]]:
    if (date_from.year, date_from.month) != (date_to.year, date_to.month):
        return {}
    return monthly_plans.get(date_from.strftime("%Y-%m"), {})


def build_dota_workbook(data: DotaReportData) -> bytes:
    period = f"Период: {data.date_from:%d.%m.%Y}–{data.date_to:%d.%m.%Y}"
    workbook = Workbook(f"Отчёт ДОТ {period}")
    _build_summary_sheet(workbook, data, period)
    launches = tuple(event for event in data.events if event.metric == "Запуск")
    releases = tuple(event for event in data.events if event.metric == "Выпуск")
    _build_daily_sheet(workbook, "Запуск", launches, data.date_from, data.date_to, period)
    _build_daily_sheet(workbook, "Выпуск", releases, data.date_from, data.date_to, period)
    _build_detail_sheet(workbook, "Детали запуска", launches, period)
    _build_detail_sheet(workbook, "Детали выпуска", releases, period)
    return workbook.to_bytes()


def _title(sheet, title: str, period: str, columns: int) -> None:
    sheet.append([title] + [None] * (columns - 1), STYLE["title"])
    sheet.append([period] + [None] * (columns - 1), STYLE["subtitle"])
    last_column = _excel_column(columns)
    sheet.merges.extend([f"A1:{last_column}1", f"A2:{last_column}2"])
    sheet.row_heights[1] = 27
    sheet.row_heights[2] = 22


def _excel_column(count: int) -> str:
    value = count
    result = ""
    while value:
        value, remainder = divmod(value - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _events(
    rows: Sequence[DotaEvent],
    metric: str,
    category: str | None = None,
) -> list[DotaEvent]:
    return [
        row
        for row in rows
        if row.metric == metric and (category is None or row.category == category)
    ]


def _sum_amount(rows: Iterable[DotaEvent]) -> Decimal:
    return sum((row.amount for row in rows), Decimal(0))


def _ratio(actual: Decimal | int, baseline: Decimal | int | None) -> float | None:
    if not baseline:
        return None
    return float(actual) / float(baseline)


def _change(actual: Decimal | int, previous: Decimal | int) -> float | None:
    if not previous:
        return None
    return (float(actual) - float(previous)) / float(previous)


def _build_summary_sheet(workbook: Workbook, data: DotaReportData, period: str) -> None:
    columns = 10
    sheet = workbook.add_sheet("Сводная")
    _title(sheet, "Отчёт ДОТ", period, columns)
    sheet.append([None] * columns)

    for metric in ("Запуск", "Выпуск"):
        plans = data.plans.get(metric, {})
        section_row = sheet.append([metric.upper()] + [None] * (columns - 1), STYLE["section"])
        sheet.merges.append(f"A{section_row}:J{section_row}")
        sheet.append(
            [
                "Направление",
                "План, руб.",
                "Факт, руб.",
                "Выполнение",
                "Пред. период, руб.",
                "Изменение",
                "Проекты, шт.",
                "Пред. период, шт.",
                "Средний чек, руб.",
                "Осталось, руб.",
            ],
            STYLE["table_header"],
        )
        for category in CATEGORIES:
            current = _events(data.events, metric, category)
            previous = _events(data.previous_events, metric, category)
            amount = _sum_amount(current)
            previous_amount = _sum_amount(previous)
            plan = plans.get(category)
            average = amount / len(current) if current else Decimal(0)
            sheet.append(
                [
                    category,
                    plan if plan is not None else "—",
                    amount,
                    _ratio(amount, plan),
                    previous_amount,
                    _change(amount, previous_amount),
                    len(current),
                    len(previous),
                    average,
                    max(plan - amount, Decimal(0)) if plan is not None else "—",
                ],
                [
                    STYLE["table_text"],
                    STYLE["table_money"],
                    STYLE["table_money"],
                    STYLE["table_percent"],
                    STYLE["table_money"],
                    STYLE["table_percent"],
                    STYLE["table_number"],
                    STYLE["table_number"],
                    STYLE["table_money"],
                    STYLE["table_money"],
                ],
            )

        current_all = _events(data.events, metric)
        previous_all = _events(data.previous_events, metric)
        has_complete_plan = all(category in plans for category in CATEGORIES)
        total_plan = sum((plans[category] for category in CATEGORIES), Decimal(0)) if has_complete_plan else None
        total_amount = _sum_amount(current_all)
        previous_total = _sum_amount(previous_all)
        sheet.append(
            [
                "Итого",
                total_plan if total_plan is not None else "—",
                total_amount,
                _ratio(total_amount, total_plan),
                previous_total,
                _change(total_amount, previous_total),
                len(current_all),
                len(previous_all),
                total_amount / len(current_all) if current_all else Decimal(0),
                max(total_plan - total_amount, Decimal(0)) if total_plan is not None else "—",
            ],
            [
                STYLE["total_text"],
                STYLE["total_money"],
                STYLE["total_money"],
                STYLE["total_percent"],
                STYLE["total_money"],
                STYLE["total_percent"],
                STYLE["total_number"],
                STYLE["total_number"],
                STYLE["total_money"],
                STYLE["total_money"],
            ],
        )
        sheet.append([None] * columns)

    plan_note = (
        f"План: {data.plan_month}, полный месяц. "
        if data.plan_month and data.plans
        else "План для выбранного периода не настроен. "
    )
    note_row = sheet.append(
        [
            plan_note
            + f"Сравнение: {data.previous_date_from:%d.%m.%Y}–{data.previous_date_to:%d.%m.%Y}. "
            + "Повтор проекта по одному показателю не учитывается.",
        ]
        + [None] * (columns - 1),
        STYLE["note"],
    )
    sheet.merges.append(f"A{note_row}:J{note_row}")
    sheet.row_heights[note_row] = 58
    sheet.widths = {
        0: 22,
        1: 17,
        2: 17,
        3: 15,
        4: 20,
        5: 15,
        6: 16,
        7: 20,
        8: 21,
        9: 18,
    }
    sheet.freeze_rows = 5


def _build_daily_sheet(
    workbook: Workbook,
    title: str,
    rows: Sequence[DotaEvent],
    date_from: date,
    date_to: date,
    period: str,
) -> None:
    columns = 1 + len(CATEGORIES) * 2 + 1 + 2
    sheet = workbook.add_sheet(title)
    _title(sheet, title, period, columns)
    sheet.append([None] * columns)
    headers = ["Дата"]
    for category in CATEGORIES:
        headers.append(f"{category}, руб.")
        if category == "ОПР":
            headers.append("ОПР, РМ факт" if title == "Выпуск" else "ОПР, РМ")
        headers.append(f"{category}, шт.")
    headers.extend(["Всего, руб.", "Всего, шт."])
    sheet.append(headers, STYLE["table_header"])

    by_day: dict[date, dict[str, list[DotaEvent]]] = defaultdict(lambda: defaultdict(list))
    for row in rows:
        by_day[row.event_at.date()][row.category].append(row)

    current = date_from
    while current <= date_to:
        values: list[object] = [current]
        styles = [STYLE["table_center"]]
        day_amount = Decimal(0)
        day_count = 0
        for category in CATEGORIES:
            category_rows = by_day[current].get(category, [])
            amount = _sum_amount(category_rows)
            count = len(category_rows)
            values.append(amount)
            styles.append(STYLE["table_money"])
            if category == "ОПР":
                values.append(sum(row.workplace_count for row in category_rows))
                styles.append(STYLE["table_number"])
            values.append(count)
            styles.append(STYLE["table_number"])
            day_amount += amount
            day_count += count
        values.extend([day_amount, day_count])
        styles.extend([STYLE["total_money"], STYLE["total_number"]])
        sheet.append(values, styles)
        current += timedelta(days=1)

    totals: list[object] = ["Итого"]
    total_styles = [STYLE["total_text"]]
    for category in CATEGORIES:
        category_rows = [row for row in rows if row.category == category]
        totals.append(_sum_amount(category_rows))
        total_styles.append(STYLE["total_money"])
        if category == "ОПР":
            totals.append(sum(row.workplace_count for row in category_rows))
            total_styles.append(STYLE["total_number"])
        totals.append(len(category_rows))
        total_styles.append(STYLE["total_number"])
    totals.extend([_sum_amount(rows), len(rows)])
    total_styles.extend([STYLE["total_money"], STYLE["total_number"]])
    sheet.append(totals, total_styles)
    sheet.widths = {0: 14}
    for index, header in enumerate(headers[1:], start=1):
        sheet.widths[index] = 17 if "руб." in header else 14
    sheet.freeze_rows = 4
    sheet.auto_filter = f"A4:{_excel_column(columns)}{max(4, len(sheet.rows) - 1)}"


def _build_detail_sheet(
    workbook: Workbook,
    title: str,
    rows: Sequence[DotaEvent],
    period: str,
) -> None:
    columns = 11
    sheet = workbook.add_sheet(title)
    _title(sheet, title, period, columns)
    sheet.append([None] * columns)
    sheet.append(
        [
            "Направление",
            "Проект ID",
            "Договор",
            "Компания",
            "Менеджер",
            "Отдел менеджера",
            "Дата события",
            "Стоимость, руб.",
            "РМ факт" if title == "Детали выпуска" else "РМ",
            "Этап выпуска" if title == "Детали выпуска" else "Этап запуска",
            "Изменил этап",
        ],
        STYLE["table_header"],
    )
    for row in rows:
        sheet.append(
            [
                row.category,
                row.project_id,
                row.contract_number or "—",
                row.company_name,
                row.manager,
                row.manager_department,
                row.event_at,
                row.amount,
                row.workplace_count,
                row.new_step,
                row.event_owner,
            ],
            [
                STYLE["table_text"],
                STYLE["table_text"],
                STYLE["table_text"],
                STYLE["table_text"],
                STYLE["table_text"],
                STYLE["table_text"],
                STYLE["table_center"],
                STYLE["table_money"],
                STYLE["table_number"],
                STYLE["table_text"],
                STYLE["table_text"],
            ],
        )
    sheet.widths = {
        0: 19,
        1: 38,
        2: 22,
        3: 39,
        4: 28,
        5: 19,
        6: 21,
        7: 20,
        8: 10,
        9: 34,
        10: 27,
    }
    sheet.freeze_rows = 4
    sheet.auto_filter = f"A4:K{max(4, len(sheet.rows))}"
