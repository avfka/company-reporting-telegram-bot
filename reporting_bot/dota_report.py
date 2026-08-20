from __future__ import annotations

import calendar
import base64
import io
import json
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from decimal import Decimal
from functools import lru_cache
from pathlib import Path
from typing import Iterable, Mapping, Sequence

from PIL import Image, ImageDraw, ImageFont
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


@dataclass(frozen=True)
class DotaReportArtifact:
    workbook: bytes
    workbook_filename: str
    chart: bytes
    chart_filename: str
    caption: str


class DotaReportService:
    def __init__(self, settings: Settings) -> None:
        self._settings = settings
        self._database_url = settings.database_url.replace(
            "postgresql://", "postgresql+psycopg://", 1
        )
        self._monthly_plans = _parse_monthly_plans(settings.dota_plans_json)

    def create(self, date_from: date, date_to: date) -> DotaReportArtifact:
        if date_to < date_from:
            raise ValueError("Дата окончания не может быть раньше даты начала.")
        if (date_to - date_from).days > 366:
            raise ValueError("Максимальный период отчёта — 366 дней.")
        previous_date_from = _previous_month_date(date_from)
        previous_date_to = _previous_month_date(date_to)
        data = self._load(date_from, date_to, previous_date_from, previous_date_to)
        stem = f"Отчет_ДОТ_{date_from.isoformat()}_{date_to.isoformat()}"
        return DotaReportArtifact(
            workbook=build_dota_workbook(data),
            workbook_filename=stem + ".xlsx",
            chart=build_dota_chart(data),
            chart_filename=stem + ".png",
            caption=f"<b>Отчёт ДОТ</b> · {date_from:%d.%m.%Y}–{date_to:%d.%m.%Y}",
        )

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


@lru_cache(maxsize=32)
def _dota_font(size: int, bold: bool = False):
    asset_name = "ks-report-bold.ttf.b64" if bold else "ks-report-regular.ttf.b64"
    asset_path = Path(__file__).with_name("assets") / asset_name
    try:
        font_bytes = base64.b64decode(asset_path.read_text(encoding="ascii"))
        return ImageFont.truetype(io.BytesIO(font_bytes), size)
    except (OSError, ValueError):
        return ImageFont.load_default(size=size)


def _dota_money(value: Decimal | float) -> str:
    number = float(value)
    if abs(number) >= 1_000_000:
        return f"{number / 1_000_000:.1f} млн ₽".replace(".", ",")
    if abs(number) >= 1_000:
        return f"{number / 1_000:.0f} тыс. ₽"
    return f"{number:.0f} ₽"


def _dota_change(current: Decimal | int, previous: Decimal | int) -> str:
    change = _change(current, previous)
    if change is None:
        return "нет базы сравнения"
    sign = "+" if change >= 0 else ""
    return f"к прошлому месяцу {sign}{change * 100:.1f}%".replace(".", ",")


def _dota_plan_label(actual: Decimal, plans: Mapping[str, Decimal]) -> str:
    if not all(category in plans for category in CATEGORIES):
        return "план не задан"
    plan = sum((plans[category] for category in CATEGORIES), Decimal(0))
    if not plan:
        return "план — 0 ₽"
    return f"{float(actual / plan) * 100:.0f}% месячного плана"


def _draw_dota_card(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    label: str,
    value: str,
    detail: str,
    accent: str,
) -> None:
    left, top, right, bottom = box
    draw.rounded_rectangle(box, radius=18, fill="#F8FAFD", outline="#DCE5EF", width=2)
    draw.rounded_rectangle((left, top, left + 8, bottom), radius=4, fill=accent)
    draw.text((left + 22, top + 15), label, font=_dota_font(16), fill="#5B6B7C")
    draw.text((left + 22, top + 45), value, font=_dota_font(27, True), fill="#173A5E")
    draw.text((left + 22, top + 82), detail, font=_dota_font(13), fill="#7A8998")


def _dota_buckets(
    data: DotaReportData,
) -> tuple[str, list[str], list[float], list[float]]:
    by_month = (data.date_to - data.date_from).days > 45
    launch: defaultdict[date, Decimal] = defaultdict(Decimal)
    release: defaultdict[date, Decimal] = defaultdict(Decimal)
    for event in data.events:
        event_date = event.event_at.date()
        key = event_date.replace(day=1) if by_month else event_date - timedelta(days=event_date.weekday())
        target = launch if event.metric == "Запуск" else release
        target[key] += event.amount
    keys = sorted(set(launch) | set(release))
    if by_month:
        labels = [key.strftime("%m.%Y") for key in keys]
        title = "Динамика по месяцам"
    else:
        labels = [
            f"{max(key, data.date_from):%d.%m}–{min(key + timedelta(days=6), data.date_to):%d.%m}"
            for key in keys
        ]
        title = "Динамика по неделям"
    return (
        title,
        labels,
        [float(launch[key]) for key in keys],
        [float(release[key]) for key in keys],
    )


def build_dota_chart(data: DotaReportData) -> bytes:
    image = Image.new("RGB", (1200, 1500), "#F4F7FB")
    draw = ImageDraw.Draw(image)
    blue = "#255985"
    orange = "#E89A3C"
    green = "#2E9C73"
    draw.rounded_rectangle((45, 35, 1155, 1465), radius=30, fill="#FFFFFF", outline="#DCE5EF", width=2)
    draw.text((85, 70), "ДОТ · Запуски и выпуски", font=_dota_font(41, True), fill="#173A5E")
    draw.text(
        (85, 126),
        f"{data.date_from:%d.%m.%Y}–{data.date_to:%d.%m.%Y} · сравнение {data.previous_date_from:%d.%m.%Y}–{data.previous_date_to:%d.%m.%Y}",
        font=_dota_font(21),
        fill="#5B6B7C",
    )

    launches = _events(data.events, "Запуск")
    releases = _events(data.events, "Выпуск")
    previous_launches = _events(data.previous_events, "Запуск")
    previous_releases = _events(data.previous_events, "Выпуск")
    launch_amount = _sum_amount(launches)
    release_amount = _sum_amount(releases)
    previous_launch_amount = _sum_amount(previous_launches)
    previous_release_amount = _sum_amount(previous_releases)
    launch_rm = sum(row.workplace_count for row in launches if row.category == "ОПР")
    release_rm = sum(row.workplace_count for row in releases if row.category == "ОПР")
    previous_launch_rm = sum(row.workplace_count for row in previous_launches if row.category == "ОПР")
    previous_release_rm = sum(row.workplace_count for row in previous_releases if row.category == "ОПР")

    card_top = 178
    card_width = 245
    cards = (
        ("Запуск, руб.", _dota_money(launch_amount), _dota_plan_label(launch_amount, data.plans.get("Запуск", {})), blue),
        ("Выпуск, руб.", _dota_money(release_amount), _dota_plan_label(release_amount, data.plans.get("Выпуск", {})), orange),
        ("Запущено проектов", str(len(launches)), _dota_change(len(launches), len(previous_launches)), blue),
        ("Выпущено проектов", str(len(releases)), _dota_change(len(releases), len(previous_releases)), orange),
    )
    for index, card in enumerate(cards):
        left = 85 + index * 255
        _draw_dota_card(draw, (left, card_top, left + card_width, card_top + 112), *card)

    _draw_dota_card(
        draw,
        (85, 305, 585, 417),
        "ОПР · РМ при запуске",
        f"{launch_rm:,}".replace(",", " "),
        _dota_change(launch_rm, previous_launch_rm),
        blue,
    )
    _draw_dota_card(
        draw,
        (605, 305, 1105, 417),
        "ОПР · фактические РМ при выпуске",
        f"{release_rm:,}".replace(",", " "),
        _dota_change(release_rm, previous_release_rm),
        orange,
    )

    panel_top = 440
    draw.rounded_rectangle((85, panel_top, 1105, 830), radius=20, fill="#F8FAFD", outline="#DCE5EF", width=2)
    draw.text((108, panel_top + 18), "Направления: запуск и выпуск", font=_dota_font(22, True), fill="#213547")
    draw.rectangle((745, panel_top + 23, 767, panel_top + 38), fill=blue)
    draw.text((778, panel_top + 17), "Запуск", font=_dota_font(15), fill="#5B6B7C")
    draw.rectangle((875, panel_top + 23, 897, panel_top + 38), fill=orange)
    draw.text((908, panel_top + 17), "Выпуск", font=_dota_font(15), fill="#5B6B7C")
    category_values = [
        (
            category,
            float(_sum_amount(_events(data.events, "Запуск", category))),
            float(_sum_amount(_events(data.events, "Выпуск", category))),
        )
        for category in CATEGORIES
    ]
    maximum = max((value for _, launch, release in category_values for value in (launch, release)), default=1.0) or 1.0
    for index, (category, launch_value, release_value) in enumerate(category_values):
        top = panel_top + 67 + index * 50
        draw.text((108, top + 7), category, font=_dota_font(16, True), fill="#213547")
        bar_left = 300
        launch_width = int(540 * launch_value / maximum)
        release_width = int(540 * release_value / maximum)
        if launch_width:
            draw.rounded_rectangle((bar_left, top, bar_left + launch_width, top + 17), radius=6, fill=blue)
        if release_width:
            draw.rounded_rectangle((bar_left, top + 23, bar_left + release_width, top + 40), radius=6, fill=orange)
        draw.text((870, top - 3), _dota_money(launch_value), font=_dota_font(14, True), fill="#173A5E")
        draw.text((990, top + 20), _dota_money(release_value), font=_dota_font(14), fill="#7C5630")

    department_top = 852
    draw.rounded_rectangle((85, department_top, 1105, 1122), radius=20, fill="#F8FAFD", outline="#DCE5EF", width=2)
    draw.text((108, department_top + 18), "Запуски по отделам", font=_dota_font(22, True), fill="#213547")
    headers = (("Отдел", 108), ("Деньги", 300), ("Проекты", 520), ("ОПР, РМ", 700), ("Средний чек", 865))
    for label, x in headers:
        draw.text((x, department_top + 58), label, font=_dota_font(15, True), fill="#5B6B7C")
    draw.line((108, department_top + 84, 1080, department_top + 84), fill="#DCE5EF", width=2)
    department_rows: dict[str, list[DotaEvent]] = defaultdict(list)
    for event in launches:
        department_rows[event.manager_department].append(event)
    department_order = [name for name in ("ГТО", "КАМ", "ОАП", "ОП", "—") if name in department_rows]
    for index, department in enumerate(department_order[:5]):
        rows = department_rows[department]
        amount = _sum_amount(rows)
        rm = sum(row.workplace_count for row in rows if row.category == "ОПР")
        average = amount / len(rows) if rows else Decimal(0)
        top = department_top + 98 + index * 32
        draw.text((108, top), department, font=_dota_font(16, True), fill="#213547")
        draw.text((300, top), _dota_money(amount), font=_dota_font(16), fill="#173A5E")
        draw.text((520, top), str(len(rows)), font=_dota_font(16), fill="#173A5E")
        draw.text((700, top), f"{rm:,}".replace(",", " "), font=_dota_font(16), fill="#173A5E")
        draw.text((865, top), _dota_money(average), font=_dota_font(16), fill="#173A5E")

    trend_top = 1145
    draw.rounded_rectangle((85, trend_top, 1105, 1410), radius=20, fill="#F8FAFD", outline="#DCE5EF", width=2)
    trend_title, labels, launch_values, release_values = _dota_buckets(data)
    draw.text((108, trend_top + 18), trend_title, font=_dota_font(22, True), fill="#213547")
    draw.rectangle((775, trend_top + 23, 797, trend_top + 38), fill=blue)
    draw.text((808, trend_top + 17), "Запуск", font=_dota_font(15), fill="#5B6B7C")
    draw.rectangle((905, trend_top + 23, 927, trend_top + 38), fill=orange)
    draw.text((938, trend_top + 17), "Выпуск", font=_dota_font(15), fill="#5B6B7C")
    if labels:
        maximum = max([*launch_values, *release_values, 1.0])
        chart_left, chart_right = 125, 1065
        chart_bottom = trend_top + 215
        group_width = (chart_right - chart_left) / len(labels)
        bar_width = min(28, max(8, int(group_width / 3)))
        for index, label in enumerate(labels):
            center = chart_left + group_width * (index + 0.5)
            launch_height = int(115 * launch_values[index] / maximum)
            release_height = int(115 * release_values[index] / maximum)
            draw.rounded_rectangle((int(center - bar_width - 3), chart_bottom - launch_height, int(center - 3), chart_bottom), radius=4, fill=blue)
            draw.rounded_rectangle((int(center + 3), chart_bottom - release_height, int(center + bar_width + 3), chart_bottom), radius=4, fill=orange)
            if launch_height and len(labels) <= 6:
                value_label = _dota_money(launch_values[index])
                bbox = draw.textbbox((0, 0), value_label, font=_dota_font(10, True))
                draw.text(
                    (center - bar_width / 2 - 3 - (bbox[2] - bbox[0]) / 2, chart_bottom - launch_height - 16),
                    value_label,
                    font=_dota_font(10, True),
                    fill="#173A5E",
                )
            if release_height and len(labels) <= 6:
                value_label = _dota_money(release_values[index])
                bbox = draw.textbbox((0, 0), value_label, font=_dota_font(10))
                draw.text(
                    (center + bar_width / 2 + 3 - (bbox[2] - bbox[0]) / 2, chart_bottom - release_height - 16),
                    value_label,
                    font=_dota_font(10),
                    fill="#7C5630",
                )
            display = label if len(label) <= 13 else label[:12] + "…"
            bbox = draw.textbbox((0, 0), display, font=_dota_font(12))
            draw.text((center - (bbox[2] - bbox[0]) / 2, chart_bottom + 8), display, font=_dota_font(12), fill="#5B6B7C")
        draw.line((chart_left, chart_bottom, chart_right, chart_bottom), fill="#C8D4E1", width=2)
    else:
        draw.text((420, trend_top + 125), "Нет данных за выбранный период", font=_dota_font(20), fill="#7A8998")

    draw.text((85, 1430), "EcoStar Reports · данные CRM", font=_dota_font(18), fill="#7A8998")
    output = io.BytesIO()
    image.save(output, format="PNG", optimize=True)
    return output.getvalue()
