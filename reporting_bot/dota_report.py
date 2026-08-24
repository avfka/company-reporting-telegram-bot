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
class DotaChartArtifact:
    content: bytes
    filename: str
    caption: str


@dataclass(frozen=True)
class DotaReportArtifact:
    workbook: bytes
    workbook_filename: str
    chart: bytes
    chart_filename: str
    caption: str
    detail_charts: tuple[DotaChartArtifact, ...] = ()


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
        detail_charts = build_dota_detail_charts(data, stem)
        return DotaReportArtifact(
            workbook=build_dota_workbook(data),
            workbook_filename=stem + ".xlsx",
            chart=build_dota_chart(data),
            chart_filename=stem + ".png",
            caption=f"<b>Отчёт ДОТ</b> · {date_from:%d.%m.%Y}–{date_to:%d.%m.%Y}",
            detail_charts=detail_charts,
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


def _dota_share_rows(
    rows: Sequence[DotaEvent],
    label_getter,
    *,
    max_items: int = 4,
) -> list[tuple[str, Decimal]]:
    totals: defaultdict[str, Decimal] = defaultdict(Decimal)
    for row in rows:
        totals[str(label_getter(row) or "—")] += row.amount
    ranked = sorted(
        ((label, amount) for label, amount in totals.items() if amount > 0),
        key=lambda item: item[1],
        reverse=True,
    )
    if len(ranked) <= max_items:
        return ranked
    return [*ranked[:max_items], ("Остальные", sum((amount for _, amount in ranked[max_items:]), Decimal(0)))]


def _draw_dota_donut(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    title: str,
    rows: Sequence[tuple[str, Decimal]],
) -> None:
    left, top, right, bottom = box
    colors = ("#255985", "#E89A3C", "#2E9C73", "#7C5AA6", "#98A7B6")
    draw.rounded_rectangle(box, radius=20, fill="#F8FAFD", outline="#DCE5EF", width=2)
    draw.text((left + 23, top + 18), title, font=_dota_font(21, True), fill="#213547")
    total = sum((amount for _, amount in rows), Decimal(0))
    pie_box = (left + 120, top + 64, left + 380, top + 324)
    if total:
        angle = -90.0
        for index, (_, amount) in enumerate(rows):
            next_angle = angle + 360.0 * float(amount / total)
            draw.pieslice(pie_box, start=angle, end=next_angle, fill=colors[index % len(colors)], outline="#FFFFFF", width=3)
            angle = next_angle
        draw.ellipse((left + 190, top + 134, left + 310, top + 254), fill="#F8FAFD")
        total_label = _dota_money(total)
        bbox = draw.textbbox((0, 0), total_label, font=_dota_font(18, True))
        draw.text((left + 250 - (bbox[2] - bbox[0]) / 2, top + 174), total_label, font=_dota_font(18, True), fill="#173A5E")
        legend_top = top + 342
        for index, (label, amount) in enumerate(rows):
            y = legend_top + index * 27
            draw.rounded_rectangle((left + 26, y + 3, left + 42, y + 19), radius=4, fill=colors[index % len(colors)])
            display = label if len(label) <= 22 else label[:21] + "…"
            share = float(amount / total) * 100
            draw.text((left + 54, y), display, font=_dota_font(14, True), fill="#213547")
            detail = f"{share:.0f}% · {_dota_money(amount)}"
            detail_box = draw.textbbox((0, 0), detail, font=_dota_font(13))
            draw.text((right - 24 - (detail_box[2] - detail_box[0]), y + 1), detail, font=_dota_font(13), fill="#5B6B7C")
    else:
        draw.ellipse(pie_box, fill="#E8EEF5")
        draw.ellipse((left + 190, top + 134, left + 310, top + 254), fill="#F8FAFD")
        draw.text((left + 176, top + 350), "Нет данных за период", font=_dota_font(16), fill="#7A8998")


def _dota_daily_series(
    data: DotaReportData,
    value_getter,
    *,
    category: str | None = None,
) -> tuple[list[date], list[float], list[float]]:
    dates: list[date] = []
    current = data.date_from
    while current <= data.date_to:
        dates.append(current)
        current += timedelta(days=1)
    launch: defaultdict[date, float] = defaultdict(float)
    release: defaultdict[date, float] = defaultdict(float)
    for event in data.events:
        if category is not None and event.category != category:
            continue
        target = launch if event.metric == "Запуск" else release
        target[event.event_at.date()] += float(value_getter(event))
    return dates, [launch[value] for value in dates], [release[value] for value in dates]


def _draw_dota_daily_panel(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    title: str,
    dates: Sequence[date],
    launch: Sequence[float],
    release: Sequence[float],
    formatter,
) -> None:
    left, top, right, bottom = box
    blue = "#255985"
    orange = "#E89A3C"
    draw.rounded_rectangle(box, radius=20, fill="#F8FAFD", outline="#DCE5EF", width=2)
    draw.text((left + 23, top + 18), title, font=_dota_font(21, True), fill="#213547")
    draw.rectangle((right - 330, top + 25, right - 308, top + 40), fill=blue)
    draw.text((right - 298, top + 18), "Запуск", font=_dota_font(14), fill="#5B6B7C")
    draw.rectangle((right - 190, top + 25, right - 168, top + 40), fill=orange)
    draw.text((right - 158, top + 18), "Выпуск", font=_dota_font(14), fill="#5B6B7C")

    chart_left, chart_right = left + 74, right - 30
    chart_top, chart_bottom = top + 72, bottom - 43
    maximum = max([*launch, *release, 1.0])
    draw.line((chart_left, chart_bottom, chart_right, chart_bottom), fill="#C8D4E1", width=2)
    draw.line((chart_left, chart_top, chart_right, chart_top), fill="#E1E8F0", width=1)
    maximum_label = formatter(maximum)
    draw.text((left + 16, chart_top - 8), maximum_label, font=_dota_font(11), fill="#7A8998")
    draw.text((left + 50, chart_bottom - 7), "0", font=_dota_font(11), fill="#7A8998")
    if not dates:
        draw.text((left + 400, top + 130), "Нет данных за период", font=_dota_font(17), fill="#7A8998")
        return

    group_width = (chart_right - chart_left) / len(dates)
    bar_width = max(1, min(14, int(group_width * 0.34)))
    label_every = max(1, (len(dates) + 9) // 10)
    for index, current_date in enumerate(dates):
        center = chart_left + group_width * (index + 0.5)
        launch_height = int((chart_bottom - chart_top - 8) * launch[index] / maximum)
        release_height = int((chart_bottom - chart_top - 8) * release[index] / maximum)
        if launch_height:
            draw.rectangle((int(center - bar_width - 1), chart_bottom - launch_height, int(center - 1), chart_bottom), fill=blue)
        if release_height:
            draw.rectangle((int(center + 1), chart_bottom - release_height, int(center + bar_width + 1), chart_bottom), fill=orange)
        if index % label_every == 0 or index == len(dates) - 1:
            label = current_date.strftime("%d.%m")
            bbox = draw.textbbox((0, 0), label, font=_dota_font(10))
            draw.text((center - (bbox[2] - bbox[0]) / 2, chart_bottom + 9), label, font=_dota_font(10), fill="#5B6B7C")


def _dota_number(value: float) -> str:
    return f"{value:,.0f}".replace(",", " ")


def _dota_compact(value: float, *, money: bool = False) -> str:
    absolute = abs(value)
    if absolute >= 1_000_000:
        result = f"{value / 1_000_000:.1f}м".replace(".0м", "м").replace(".", ",")
    elif absolute >= 1_000:
        result = f"{value / 1_000:.0f}к"
    else:
        result = f"{value:.0f}"
    return result + (" ₽" if money else "")


def _dota_breakdown_rows(
    rows: Sequence[DotaEvent],
    label_getter,
    *,
    max_items: int = 4,
) -> list[tuple[str, Decimal, int, int]]:
    totals: dict[str, tuple[Decimal, int, int]] = {}
    for row in rows:
        label = str(label_getter(row) or "—")
        amount, projects, workplaces = totals.get(label, (Decimal(0), 0, 0))
        totals[label] = (amount + row.amount, projects + 1, workplaces + row.workplace_count)
    ranked = sorted(
        ((label, *values) for label, values in totals.items() if values[0] > 0),
        key=lambda item: item[1],
        reverse=True,
    )
    if len(ranked) <= max_items:
        return ranked
    remaining = ranked[max_items:]
    return [
        *ranked[:max_items],
        (
            "Остальные",
            sum((row[1] for row in remaining), Decimal(0)),
            sum(row[2] for row in remaining),
            sum(row[3] for row in remaining),
        ),
    ]


def _draw_right_text(
    draw: ImageDraw.ImageDraw,
    right: int,
    y: int,
    text_value: str,
    font,
    fill: str,
) -> None:
    bounds = draw.textbbox((0, 0), text_value, font=font)
    draw.text((right - (bounds[2] - bounds[0]), y), text_value, font=font, fill=fill)


def _build_dota_donut_detail(
    data: DotaReportData,
    *,
    category: str,
    title: str,
    subtitle: str,
    label_getter,
    include_workplaces: bool,
) -> bytes:
    image = Image.new("RGB", (1200, 900), "#F1F4F8")
    draw = ImageDraw.Draw(image)
    navy = "#173A5E"
    colors = ("#F47C2B", "#FFBC28", "#70AD47", "#A84D0B", "#4E88B5")
    draw.rounded_rectangle((38, 32, 1162, 866), radius=28, fill="#FFFFFF", outline="#D5DEE8", width=2)
    draw.text((78, 68), title, font=_dota_font(38, True), fill=navy)
    draw.text((78, 118), subtitle, font=_dota_font(20), fill="#5B6B7C")
    draw.text(
        (78, 151),
        f"Период: {data.date_from:%d.%m.%Y}–{data.date_to:%d.%m.%Y}",
        font=_dota_font(17),
        fill="#7A8998",
    )

    launches = _events(data.events, "Запуск", category)
    rows = _dota_breakdown_rows(launches, label_getter)
    total = sum((row[1] for row in rows), Decimal(0))
    pie_box = (92, 225, 612, 745)
    if total:
        angle = -90.0
        for index, (_, amount, _, _) in enumerate(rows):
            next_angle = angle + 360.0 * float(amount / total)
            draw.pieslice(
                pie_box,
                start=angle,
                end=next_angle,
                fill=colors[index % len(colors)],
                outline="#FFFFFF",
                width=5,
            )
            angle = next_angle
        draw.ellipse((237, 370, 467, 600), fill="#FFFFFF")
        total_label = _dota_money(total)
        bounds = draw.textbbox((0, 0), total_label, font=_dota_font(27, True))
        draw.text((352 - (bounds[2] - bounds[0]) / 2, 446), total_label, font=_dota_font(27, True), fill=navy)
        count_label = f"{len(launches)} проектов"
        bounds = draw.textbbox((0, 0), count_label, font=_dota_font(17))
        draw.text((352 - (bounds[2] - bounds[0]) / 2, 486), count_label, font=_dota_font(17), fill="#6B7A89")
    else:
        draw.ellipse(pie_box, fill="#E5EBF2")
        draw.ellipse((237, 370, 467, 600), fill="#FFFFFF")
        draw.text((273, 470), "Нет данных", font=_dota_font(22, True), fill="#7A8998")

    table_left, table_top, table_right = 650, 220, 1122
    draw.rounded_rectangle((table_left, table_top, table_right, 758), radius=18, fill="#F8FAFD", outline="#DCE5EF", width=2)
    draw.rounded_rectangle((table_left, table_top, table_right, table_top + 62), radius=18, fill=navy)
    draw.rectangle((table_left, table_top + 44, table_right, table_top + 62), fill=navy)
    draw.text((table_left + 20, table_top + 19), "Структура запусков", font=_dota_font(19, True), fill="#FFFFFF")
    columns_y = table_top + 83
    draw.text((table_left + 20, columns_y), "Отдел" if include_workplaces else "Компания", font=_dota_font(14, True), fill="#657587")
    _draw_right_text(draw, table_left + 326, columns_y, "Деньги", _dota_font(14, True), "#657587")
    _draw_right_text(draw, table_left + 396, columns_y, "Проекты", _dota_font(14, True), "#657587")
    _draw_right_text(draw, table_right - 20, columns_y, "РМ" if include_workplaces else "Доля", _dota_font(14, True), "#657587")
    draw.line((table_left + 18, columns_y + 28, table_right - 18, columns_y + 28), fill="#DCE5EF", width=1)
    for index, (label, amount, projects, workplaces) in enumerate(rows):
        y = columns_y + 48 + index * 70
        color = colors[index % len(colors)]
        draw.rounded_rectangle((table_left + 20, y + 2, table_left + 36, y + 18), radius=4, fill=color)
        display = label if len(label) <= 20 else label[:19] + "…"
        draw.text((table_left + 47, y), display, font=_dota_font(16, True), fill="#213547")
        _draw_right_text(draw, table_left + 326, y, _dota_money(amount), _dota_font(15), "#34495E")
        _draw_right_text(draw, table_left + 396, y, str(projects), _dota_font(15), "#34495E")
        last_value = str(workplaces) if include_workplaces else (f"{float(amount / total) * 100:.0f}%" if total else "0%")
        _draw_right_text(draw, table_right - 20, y, last_value, _dota_font(15, True), navy)
        if index < len(rows) - 1:
            draw.line((table_left + 20, y + 42, table_right - 20, y + 42), fill="#E8EDF3", width=1)

    draw.text((78, 818), "EcoStar Reports · данные CRM", font=_dota_font(16), fill="#8291A1")
    output = io.BytesIO()
    image.save(output, format="PNG", optimize=True)
    return output.getvalue()


def _draw_dota_dark_daily_panel(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    title: str,
    dates: Sequence[date],
    launch: Sequence[float],
    release: Sequence[float],
    *,
    money: bool,
) -> None:
    left, top, right, bottom = box
    launch_fill, launch_outline = "#203D58", "#4DB7F0"
    release_fill, release_outline = "#5B3A20", "#F29B38"
    draw.rounded_rectangle(box, radius=18, fill="#333536", outline="#5A5E62", width=2)
    draw.text((left + 24, top + 19), title, font=_dota_font(22, True), fill="#F5F7FA")
    legend_y = top + 25
    draw.rounded_rectangle((right - 310, legend_y, right - 288, legend_y + 16), radius=3, fill=launch_fill, outline=launch_outline, width=2)
    draw.text((right - 278, top + 18), "Запуск", font=_dota_font(14), fill="#D9E0E7")
    draw.rounded_rectangle((right - 165, legend_y, right - 143, legend_y + 16), radius=3, fill=release_fill, outline=release_outline, width=2)
    draw.text((right - 133, top + 18), "Выпуск", font=_dota_font(14), fill="#D9E0E7")

    chart_left, chart_right = left + 82, right - 32
    chart_top, chart_bottom = top + 76, bottom - 50
    maximum = max([*launch, *release, 1.0])
    for step in range(5):
        y = int(chart_bottom - (chart_bottom - chart_top) * step / 4)
        draw.line((chart_left, y, chart_right, y), fill="#55595D", width=1)
        value = maximum * step / 4
        label = _dota_compact(value, money=money)
        _draw_right_text(draw, chart_left - 10, y - 8, label, _dota_font(11), "#AEB7C0")
    if not dates:
        draw.text((left + 520, top + 160), "Нет данных за период", font=_dota_font(18), fill="#AEB7C0")
        return

    group_width = (chart_right - chart_left) / len(dates)
    bar_width = max(2, min(16, int(group_width * 0.32)))
    label_every = max(1, (len(dates) + 13) // 14)
    chart_height = chart_bottom - chart_top - 12
    for index, current_date in enumerate(dates):
        center = chart_left + group_width * (index + 0.5)
        values = (
            (launch[index], int(center - bar_width - 1), launch_fill, launch_outline),
            (release[index], int(center + 1), release_fill, release_outline),
        )
        for value, x, fill, outline in values:
            height = int(chart_height * value / maximum)
            if not height:
                continue
            draw.rounded_rectangle((x, chart_bottom - height, x + bar_width, chart_bottom), radius=2, fill=fill, outline=outline, width=2)
            if group_width >= 28:
                label = _dota_compact(value, money=money)
                bounds = draw.textbbox((0, 0), label, font=_dota_font(10, True))
                label_x = x + bar_width / 2 - (bounds[2] - bounds[0]) / 2
                draw.text((label_x, chart_bottom - height - 18), label, font=_dota_font(10, True), fill="#F2F4F6")
        if index % label_every == 0 or index == len(dates) - 1:
            label = current_date.strftime("%d.%m")
            bounds = draw.textbbox((0, 0), label, font=_dota_font(11))
            draw.text((center - (bounds[2] - bounds[0]) / 2, chart_bottom + 12), label, font=_dota_font(11), fill="#C3CAD1")


def _build_dota_movement_detail(
    data: DotaReportData,
    *,
    category: str,
    title: str,
    include_workplaces: bool,
) -> bytes:
    panel_count = 3 if include_workplaces else 2
    height = 1310 if include_workplaces else 980
    image = Image.new("RGB", (1400, height), "#242526")
    draw = ImageDraw.Draw(image)
    draw.text((62, 48), title, font=_dota_font(38, True), fill="#FFFFFF")
    draw.text(
        (62, 98),
        f"Динамика по календарным дням · {data.date_from:%d.%m.%Y}–{data.date_to:%d.%m.%Y}",
        font=_dota_font(19),
        fill="#B9C0C7",
    )

    category_rows = [row for row in data.events if row.category == category]
    launches = _events(category_rows, "Запуск")
    releases = _events(category_rows, "Выпуск")
    cards = (
        ("Запуск, руб.", _dota_money(_sum_amount(launches)), "#4DB7F0"),
        ("Выпуск, руб.", _dota_money(_sum_amount(releases)), "#F29B38"),
        ("Запущено проектов", str(len(launches)), "#4DB7F0"),
        ("Выпущено проектов", str(len(releases)), "#F29B38"),
    )
    for index, (label, value, accent) in enumerate(cards):
        left = 62 + index * 324
        draw.rounded_rectangle((left, 148, left + 300, 254), radius=12, fill="#333536", outline="#565A5E", width=2)
        draw.rectangle((left, 148, left + 7, 254), fill=accent)
        draw.text((left + 22, 165), label, font=_dota_font(15), fill="#B9C0C7")
        draw.text((left + 22, 198), value, font=_dota_font(26, True), fill="#FFFFFF")

    dates, money_launch, money_release = _dota_daily_series(data, lambda row: row.amount, category=category)
    _, project_launch, project_release = _dota_daily_series(data, lambda row: 1, category=category)
    panels = [
        ("Деньги по дням, руб.", money_launch, money_release, True),
        ("Проекты по дням, шт.", project_launch, project_release, False),
    ]
    if include_workplaces:
        _, rm_launch, rm_release = _dota_daily_series(data, lambda row: row.workplace_count, category=category)
        panels.append(("Рабочие места по дням, РМ", rm_launch, rm_release, False))
    for index, (panel_title, launch, release, money) in enumerate(panels):
        top = 286 + index * 322
        _draw_dota_dark_daily_panel(draw, (62, top, 1338, top + 292), panel_title, dates, launch, release, money=money)

    draw.text((62, height - 37), "EcoStar Reports · данные CRM", font=_dota_font(15), fill="#8C949C")
    output = io.BytesIO()
    image.save(output, format="PNG", optimize=True)
    return output.getvalue()


def build_dota_detail_charts(data: DotaReportData, stem: str) -> tuple[DotaChartArtifact, ...]:
    period = f"{data.date_from:%d.%m.%Y}–{data.date_to:%d.%m.%Y}"
    return (
        DotaChartArtifact(
            content=_build_dota_donut_detail(
                data,
                category="ОПР",
                title="ОПР · Запуски по отделам",
                subtitle="Доля выручки, проекты и рабочие места",
                label_getter=lambda row: row.manager_department,
                include_workplaces=True,
            ),
            filename=stem + "_ОПР_доли.png",
            caption=f"<b>ДОТ · ОПР — доли запусков</b> · {period}",
        ),
        DotaChartArtifact(
            content=_build_dota_donut_detail(
                data,
                category="Обучение",
                title="Обучение · Запуски по компаниям",
                subtitle="Доля выручки и количество проектов",
                label_getter=lambda row: row.company_name,
                include_workplaces=False,
            ),
            filename=stem + "_Обучение_доли.png",
            caption=f"<b>ДОТ · Обучение — доли запусков</b> · {period}",
        ),
        DotaChartArtifact(
            content=_build_dota_movement_detail(
                data,
                category="ОПР",
                title="ОПР · Движение по дням",
                include_workplaces=True,
            ),
            filename=stem + "_ОПР_по_дням.png",
            caption=f"<b>ДОТ · ОПР — движение по дням</b> · {period}",
        ),
        DotaChartArtifact(
            content=_build_dota_movement_detail(
                data,
                category="Обучение",
                title="Обучение · Движение по дням",
                include_workplaces=False,
            ),
            filename=stem + "_Обучение_по_дням.png",
            caption=f"<b>ДОТ · Обучение — движение по дням</b> · {period}",
        ),
    )


def build_dota_chart(data: DotaReportData) -> bytes:
    image = Image.new("RGB", (1200, 1800), "#F4F7FB")
    draw = ImageDraw.Draw(image)
    blue = "#255985"
    orange = "#E89A3C"
    draw.rounded_rectangle((45, 35, 1155, 1765), radius=30, fill="#FFFFFF", outline="#DCE5EF", width=2)
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

    opr_launches = _events(data.events, "Запуск", "ОПР")
    training_launches = _events(data.events, "Запуск", "Обучение")
    _draw_dota_donut(
        draw,
        (85, 315, 585, 815),
        "ОПР · доли запусков по отделам",
        _dota_share_rows(opr_launches, lambda row: row.manager_department),
    )
    _draw_dota_donut(
        draw,
        (605, 315, 1105, 815),
        "Обучение · доли запусков по компаниям",
        _dota_share_rows(training_launches, lambda row: row.company_name),
    )

    dates, money_launch, money_release = _dota_daily_series(data, lambda row: row.amount)
    _, project_launch, project_release = _dota_daily_series(data, lambda row: 1)
    _, rm_launch, rm_release = _dota_daily_series(data, lambda row: row.workplace_count, category="ОПР")
    _draw_dota_daily_panel(draw, (85, 840, 1105, 1125), "Деньги по дням", dates, money_launch, money_release, _dota_money)
    _draw_dota_daily_panel(draw, (85, 1145, 1105, 1430), "Проекты по дням", dates, project_launch, project_release, _dota_number)
    _draw_dota_daily_panel(draw, (85, 1450, 1105, 1735), "ОПР · РМ по дням", dates, rm_launch, rm_release, _dota_number)

    draw.text((85, 1742), "EcoStar Reports · данные CRM", font=_dota_font(17), fill="#7A8998")
    output = io.BytesIO()
    image.save(output, format="PNG", optimize=True)
    return output.getvalue()
