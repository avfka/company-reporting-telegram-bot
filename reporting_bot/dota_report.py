from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal
from typing import Iterable, Sequence

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
    coalesce(
      to_jsonb(company_row)->>'name',
      to_jsonb(company_row)->>'short_name',
      to_jsonb(company_row)->>'full_name',
      '—'
    ) AS company_name,
    concat_ws(' ', manager.last_name, manager.first_name) AS manager,
    concat_ws(' ', manager_sks.last_name, manager_sks.first_name) AS manager_sks,
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
        AND h.previous_step = 'Получить документы'
        AND h.new_step = 'Принять проект'
        THEN 'Запуск'
      WHEN p.service = 'obuchenie'
        AND h.previous_step = 'Получить оплату'
        AND h.new_step = 'Передать заявку в УЦ'
        THEN 'Запуск'
      WHEN p.service = 'other'
        AND h.previous_step = 'Получить оплату'
        AND h.new_step = 'Передать специалисту исходные данные'
        THEN 'Запуск'
      WHEN p.service IN ('opk', 'opk_eth', 'suot', 'audit')
        AND h.previous_step = 'В работе у эксперта'
        AND h.new_step = 'Согласовать отчет у клиента'
        THEN 'Выпуск'
      WHEN p.service = 'obuchenie'
        AND h.previous_step = 'Передать заявку в УЦ'
        AND h.new_step = 'Выдать удостоверение Заказчику'
        THEN 'Выпуск'
      WHEN p.service = 'autsorsing'
        AND h.previous_step = 'Отправить документы клиенту'
        AND h.new_step = 'Закрыть проект'
        THEN 'Выпуск'
      WHEN p.service = 'other'
        AND h.previous_step = 'Передать документы Заказчику'
        AND h.new_step = 'Закрыть проект'
        THEN 'Выпуск'
    END AS metric
  FROM projects_steps_history h
  JOIN projects p ON p.id = h.project_id
  LEFT JOIN companies_clone company_row ON company_row.id = p.company_id
  LEFT JOIN users manager ON manager.id = p.manager_id
  LEFT JOIN users manager_sks ON manager_sks.id = p.manager_sks_id
  LEFT JOIN users owner_user ON owner_user.id = h.owner_id
  WHERE h.created_at >= CAST(:date_from AS DATE)
    AND h.created_at < CAST(:date_to_exclusive AS DATE)
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
  manager_sks,
  event_owner,
  event_at,
  previous_step,
  new_step,
  coalesce(sale_price, 0) AS sale_price,
  coalesce(workplace_count, 0) AS workplace_count
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
    manager_sks: str
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


class DotaReportService:
    def __init__(self, settings: Settings) -> None:
        self._settings = settings
        self._database_url = settings.database_url.replace(
            "postgresql://", "postgresql+psycopg://", 1
        )

    def create(self, date_from: date, date_to: date) -> tuple[bytes, str]:
        if date_to < date_from:
            raise ValueError("Дата окончания не может быть раньше даты начала.")
        if (date_to - date_from).days > 366:
            raise ValueError("Максимальный период отчёта — 366 дней.")
        duration = (date_to - date_from).days + 1
        previous_date_to = date_from - timedelta(days=1)
        previous_date_from = previous_date_to - timedelta(days=duration - 1)
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
            manager_sks=str(row["manager_sks"] or "—"),
            event_owner=str(row["event_owner"] or "—"),
            event_at=row["event_at"],
            previous_step=str(row["previous_step"]),
            new_step=str(row["new_step"]),
            amount=Decimal(str(row["sale_price"] or 0)),
            workplace_count=int(row["workplace_count"] or 0),
        )
        for row in connection.execute(text(DOTA_EVENTS_SQL), params).mappings()
    )


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


def _ratio(actual: Decimal | int, baseline: Decimal | int) -> float | None:
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

    for metric, plans in (("Запуск", LAUNCH_PLANS), ("Выпуск", RELEASE_PLANS)):
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
            plan = plans[category]
            average = amount / len(current) if current else Decimal(0)
            sheet.append(
                [
                    category,
                    plan,
                    amount,
                    _ratio(amount, plan),
                    previous_amount,
                    _change(amount, previous_amount),
                    len(current),
                    len(previous),
                    average,
                    max(plan - amount, Decimal(0)),
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
        total_plan = sum(plans.values(), Decimal(0))
        total_amount = _sum_amount(current_all)
        previous_total = _sum_amount(previous_all)
        sheet.append(
            [
                "Итого",
                total_plan,
                total_amount,
                _ratio(total_amount, total_plan),
                previous_total,
                _change(total_amount, previous_total),
                len(current_all),
                len(previous_all),
                total_amount / len(current_all) if current_all else Decimal(0),
                max(total_plan - total_amount, Decimal(0)),
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

    note_row = sheet.append(
        [
            "Планы перенесены из образца за август 2026 и применяются к выбранному периоду без пересчёта. "
            f"Предыдущий сопоставимый период: {data.previous_date_from:%d.%m.%Y}–{data.previous_date_to:%d.%m.%Y}. "
            "Повторный одинаковый переход одного проекта внутри периода учитывается один раз.",
        ]
        + [None] * (columns - 1),
        STYLE["note"],
    )
    sheet.merges.append(f"A{note_row}:J{note_row}")
    sheet.row_heights[note_row] = 42
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
    columns = 2 + len(CATEGORIES) * 2 + 1
    sheet = workbook.add_sheet(title)
    _title(sheet, title, period, columns)
    sheet.append([None] * columns)
    headers = ["Дата"]
    for category in CATEGORIES:
        headers.extend([f"{category}, руб.", f"{category}, шт."])
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
            values.extend([amount, count])
            styles.extend([STYLE["table_money"], STYLE["table_number"]])
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
        totals.extend([_sum_amount(category_rows), len(category_rows)])
        total_styles.extend([STYLE["total_money"], STYLE["total_number"]])
    totals.extend([_sum_amount(rows), len(rows)])
    total_styles.extend([STYLE["total_money"], STYLE["total_number"]])
    sheet.append(totals, total_styles)
    sheet.widths = {0: 14}
    for index in range(1, columns):
        sheet.widths[index] = 16 if index % 2 else 13
    sheet.freeze_rows = 4
    sheet.auto_filter = f"A4:{_excel_column(columns)}{max(4, len(sheet.rows) - 1)}"


def _build_detail_sheet(
    workbook: Workbook,
    title: str,
    rows: Sequence[DotaEvent],
    period: str,
) -> None:
    columns = 12
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
            "Специалист СКС",
            "Дата события",
            "Стоимость, руб.",
            "РМ",
            "Предыдущий этап",
            "Новый этап",
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
                row.manager_sks,
                row.event_at,
                row.amount,
                row.workplace_count,
                row.previous_step,
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
                STYLE["table_text"],
            ],
        )
    sheet.widths = {
        0: 19,
        1: 38,
        2: 22,
        3: 39,
        4: 28,
        5: 28,
        6: 21,
        7: 20,
        8: 10,
        9: 31,
        10: 34,
        11: 27,
    }
    sheet.freeze_rows = 4
    sheet.auto_filter = f"A4:L{max(4, len(sheet.rows))}"
