from __future__ import annotations

import base64
import io
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, time, timedelta
from decimal import Decimal
from functools import lru_cache
from pathlib import Path
from statistics import mean
from typing import Iterable, Sequence

from PIL import Image, ImageDraw, ImageFont
from sqlalchemy import create_engine, text
from sqlalchemy.pool import NullPool

from reporting_bot.config import Settings
from reporting_bot.chart_layout import fitted_text
from reporting_bot.simple_xlsx import STYLE, Workbook


SCS_SPECIALISTS = (
    "Иванова Елена",
    "Шергина Надежда",
    "Григорьева Алёна",
    "Хасанова Азалия",
    "Сластина Ангелина",
    "Орлова Полина",
    "Максимович Анастасия",
    "Кирпиченко Полина",
)
PRIMARY_SPECIALISTS = (
    "Иванова Елена",
    "Максимович Анастасия",
    "Хасанова Азалия",
    "Орлова Полина",
    "Сластина Ангелина",
    "Григорьева Алёна",
)
SENDING_TASK_SPECIALISTS = ("Балакирева Диана", "Кулешева Владислава")

PLAN_SEND_DAYS = 4.0
PLAN_TASK_SECONDS = 30 * 60
PLAN_APPROVED_COUNT = 43
PLAN_APPROVED_AMOUNT = Decimal("1500000")
ANOMALY_DURATION_DAYS = 200.0

SERVICE_LABELS = {
    "sout": "СОУТ",
    "obuchenie": "Обучение",
    "opk": "ОПК",
    "autsorsing": "Аутсорсинг",
    "pk": "ПК",
    "suot": "СУОТ",
    "sbkts": "СБКТС",
    "opk_eth": "ОПК ЭТХ",
    "ppk": "ППК",
    "other": "Другая услуга",
    "hassp": "ХАССП",
    "audit": "Аудит",
    "zamery": "Замеры",
}


AGREEMENT_SQL = """
WITH target_candidates AS (
  SELECT
    h.project_id,
    h.created_at AS target_at,
    p.service,
    p.manager_sks_id,
    p.contract_number,
    p.sale_price,
    row_number() OVER (PARTITION BY h.project_id ORDER BY h.created_at, h.id) AS target_rank
  FROM projects_steps_history h
  JOIN projects p ON p.id = h.project_id
  WHERE h.created_at >= CAST(:date_from AS DATE)
    AND h.created_at < CAST(:date_to_exclusive AS DATE)
    AND (
      (p.service = 'sout' AND h.new_step LIKE 'Актуализировать РМ%')
      OR
      (COALESCE(p.service, '') <> 'sout' AND h.new_step = 'Распечатка')
    )
    AND h.previous_step = 'Согласовать отчет у клиента'
), targets AS (
  SELECT * FROM target_candidates WHERE target_rank = 1
), start_events AS (
  SELECT project_id, created_at
  FROM projects_steps_history
  WHERE new_step = 'Согласовать отчет у клиента'
), paired AS (
  SELECT t.project_id, t.target_at, t.service, t.manager_sks_id,
         t.contract_number, t.sale_price, max(s.created_at) AS start_at
  FROM targets t
  JOIN start_events s ON s.project_id = t.project_id AND s.created_at <= t.target_at
  GROUP BY t.project_id, t.target_at, t.service, t.manager_sks_id,
           t.contract_number, t.sale_price
)
SELECT
  paired.project_id::text AS project_id,
  paired.contract_number,
  paired.service,
  concat_ws(' ', u.last_name, u.first_name) AS specialist,
  paired.start_at,
  paired.target_at,
  extract(epoch FROM (paired.target_at - paired.start_at)) / 86400.0 AS duration_days,
  coalesce(paired.sale_price, 0) AS sale_price
FROM paired
LEFT JOIN users u ON u.id = paired.manager_sks_id
WHERE paired.start_at IS NOT NULL AND paired.target_at >= paired.start_at
ORDER BY paired.target_at, paired.project_id
"""


SENDING_SQL = """
WITH target_candidates AS (
  SELECT
    h.project_id,
    h.created_at AS target_at,
    p.service,
    p.manager_sks_id AS specialist_id,
    p.contract_number,
    p.sale_price,
    row_number() OVER (PARTITION BY h.project_id ORDER BY h.created_at, h.id) AS target_rank
  FROM projects_steps_history h
  JOIN projects p ON p.id = h.project_id
  WHERE h.created_at >= CAST(:date_from AS DATE)
    AND h.created_at < CAST(:date_to_exclusive AS DATE)
    AND (
      (p.service = 'sout' AND h.new_step LIKE 'Выгрузить протоколы%ФСА')
      OR
      (COALESCE(p.service, '') <> 'sout' AND h.new_step = 'Закрыть проект')
    )
), targets AS (
  SELECT * FROM target_candidates WHERE target_rank = 1
), start_events AS (
  SELECT project_id, created_at
  FROM projects_steps_history
  WHERE new_step = 'Отправить документы клиенту'
), paired AS (
  SELECT t.project_id, t.target_at, t.service, t.specialist_id,
         t.contract_number, t.sale_price, max(s.created_at) AS start_at
  FROM targets t
  JOIN start_events s ON s.project_id = t.project_id AND s.created_at <= t.target_at
  GROUP BY t.project_id, t.target_at, t.service, t.specialist_id,
           t.contract_number, t.sale_price
)
SELECT
  paired.project_id::text AS project_id,
  paired.contract_number,
  paired.service,
  concat_ws(' ', u.last_name, u.first_name) AS specialist,
  paired.start_at,
  paired.target_at,
  extract(epoch FROM (paired.target_at - paired.start_at)) / 86400.0 AS duration_days,
  coalesce(paired.sale_price, 0) AS sale_price
FROM paired
LEFT JOIN users u ON u.id = paired.specialist_id
WHERE paired.start_at IS NOT NULL
  AND paired.target_at >= paired.start_at
ORDER BY paired.target_at, paired.project_id
"""


TASK_DURATION_SQL = """
WITH completed_tasks AS (
  SELECT
    replace(
      lower(regexp_replace(trim(title), '\\s+', ' ', 'g')),
      'ё',
      'е'
    ) AS normalized_title,
    created_at,
    closed_at
  FROM tasks_clone
  WHERE done
    AND created_at IS NOT NULL
    AND closed_at IS NOT NULL
    AND closed_at >= CAST(:date_from AS DATE)
    AND closed_at < CAST(:date_to_exclusive AS DATE)
    AND closed_at >= created_at
    AND closed_at - created_at < INTERVAL '14 hours'
)
SELECT
  'Договор и счет' AS category,
  created_at,
  closed_at
FROM completed_tasks
WHERE (
    normalized_title LIKE '%договор%'
    OR normalized_title LIKE '%счет%'
  )
  AND normalized_title NOT LIKE '%эдо%'
ORDER BY closed_at
"""


SENDING_TASK_SQL = """
SELECT concat_ws(' ', u.last_name, u.first_name) AS specialist, count(DISTINCT t.id) AS task_count
FROM tasks_clone t
JOIN task_user tu ON tu.task_id = t.id
JOIN users u ON u.id = tu.user_id
WHERE t.done
  AND t.closed_at >= CAST(:date_from AS DATE)
  AND t.closed_at < CAST(:date_to_exclusive AS DATE)
  AND lower(trim(t.title)) = 'отправка'
  AND concat_ws(' ', u.last_name, u.first_name) IN ('Балакирева Диана', 'Кулешева Владислава')
GROUP BY u.last_name, u.first_name
ORDER BY u.last_name, u.first_name
"""


SENDING_ASSIGNED_TASK_SQL = """
WITH selected_tasks AS (
  SELECT DISTINCT
    t.id,
    concat_ws(' ', u.last_name, u.first_name) AS specialist
  FROM tasks_clone t
  JOIN task_user tu ON tu.task_id = t.id
  JOIN users u ON u.id = tu.user_id
  WHERE t.done
    AND t.closed_at >= CAST(:date_from AS DATE)
    AND t.closed_at < CAST(:date_to_exclusive AS DATE)
    AND concat_ws(' ', u.last_name, u.first_name) IN ('Балакирева Диана', 'Кулешева Владислава')
)
SELECT specialist, count(DISTINCT id) AS task_count
FROM selected_tasks
GROUP BY specialist
UNION ALL
SELECT '__total__' AS specialist, count(DISTINCT id) AS task_count
FROM selected_tasks
ORDER BY specialist
"""


@dataclass(frozen=True)
class ProjectMetric:
    project_id: str
    contract_number: str | None
    service: str | None
    specialist: str
    start_at: datetime
    target_at: datetime
    duration_days: float
    sale_price: Decimal


@dataclass(frozen=True)
class TaskMetric:
    category: str
    created_at: datetime
    closed_at: datetime
    working_seconds: float


@dataclass(frozen=True)
class SksReportData:
    date_from: date
    date_to: date
    agreements: tuple[ProjectMetric, ...]
    sends: tuple[ProjectMetric, ...]
    tasks: tuple[TaskMetric, ...]
    sending_tasks: dict[str, int]
    assigned_sending_tasks: dict[str, int] = field(default_factory=dict)
    assigned_sending_task_total: int = 0


@dataclass(frozen=True)
class SksReportArtifact:
    workbook: bytes
    workbook_filename: str
    chart: bytes
    chart_filename: str
    caption: str


def working_seconds(start: datetime, end: datetime) -> float:
    """Return overlap with Moscow workdays 08:00–17:30, excluding weekends."""
    if end <= start:
        return 0.0
    total = 0.0
    current_day = start.date()
    while current_day <= end.date():
        if current_day.weekday() < 5:
            work_start = datetime.combine(current_day, time(8, 0))
            work_end = datetime.combine(current_day, time(17, 30))
            overlap_start = max(start, work_start)
            overlap_end = min(end, work_end)
            if overlap_end > overlap_start:
                total += (overlap_end - overlap_start).total_seconds()
        current_day += timedelta(days=1)
    return total


class SksReportService:
    def __init__(self, settings: Settings) -> None:
        self._settings = settings
        self._database_url = settings.database_url.replace(
            "postgresql://", "postgresql+psycopg://", 1
        )

    def create(self, date_from: date, date_to: date) -> SksReportArtifact:
        if date_to < date_from:
            raise ValueError("Дата окончания не может быть раньше даты начала.")
        if (date_to - date_from).days > 366:
            raise ValueError("Максимальный период отчёта — 366 дней.")
        data = self._load(date_from, date_to)
        stem = f"Отчет_СКС_{date_from.isoformat()}_{date_to.isoformat()}"
        return SksReportArtifact(
            workbook=build_sks_workbook(data),
            workbook_filename=stem + ".xlsx",
            chart=build_sks_chart(data),
            chart_filename=stem + ".png",
            caption=f"<b>Отчёт СКС</b> · {date_from:%d.%m.%Y}–{date_to:%d.%m.%Y}",
        )

    def _load(self, date_from: date, date_to: date) -> SksReportData:
        engine = create_engine(
            self._database_url,
            poolclass=NullPool,
            connect_args={"connect_timeout": self._settings.connect_timeout_seconds},
        )
        params = {
            "date_from": date_from.isoformat(),
            "date_to_exclusive": (date_to + timedelta(days=1)).isoformat(),
        }
        try:
            with engine.connect() as connection:
                with connection.begin():
                    connection.execute(text("SET TRANSACTION READ ONLY"))
                    connection.execute(
                        text("SELECT set_config('statement_timeout', :timeout, true)"),
                        {"timeout": f"{max(self._settings.statement_timeout_ms, 30_000)}ms"},
                    )
                    agreements = tuple(
                        _project_metric(row)
                        for row in connection.execute(text(AGREEMENT_SQL), params).mappings()
                    )
                    sends = tuple(
                        _project_metric(row)
                        for row in connection.execute(text(SENDING_SQL), params).mappings()
                    )
                    tasks = tuple(
                        TaskMetric(
                            category=str(row["category"]),
                            created_at=row["created_at"],
                            closed_at=row["closed_at"],
                            working_seconds=working_seconds(row["created_at"], row["closed_at"]),
                        )
                        for row in connection.execute(text(TASK_DURATION_SQL), params).mappings()
                    )
                    sending_tasks = {name: 0 for name in SENDING_TASK_SPECIALISTS}
                    for row in connection.execute(text(SENDING_TASK_SQL), params).mappings():
                        sending_tasks[str(row["specialist"])] = int(row["task_count"])
                    assigned_sending_tasks = {name: 0 for name in SENDING_TASK_SPECIALISTS}
                    assigned_sending_task_total = 0
                    for row in connection.execute(text(SENDING_ASSIGNED_TASK_SQL), params).mappings():
                        specialist = str(row["specialist"])
                        if specialist == "__total__":
                            assigned_sending_task_total = int(row["task_count"])
                        else:
                            assigned_sending_tasks[specialist] = int(row["task_count"])
            return SksReportData(
                date_from=date_from,
                date_to=date_to,
                agreements=agreements,
                sends=sends,
                tasks=tasks,
                sending_tasks=sending_tasks,
                assigned_sending_tasks=assigned_sending_tasks,
                assigned_sending_task_total=assigned_sending_task_total,
            )
        finally:
            engine.dispose()


def _project_metric(row) -> ProjectMetric:
    return ProjectMetric(
        project_id=str(row["project_id"]),
        contract_number=row["contract_number"],
        service=row["service"],
        specialist=str(row["specialist"] or "Агентский канал / не назначен"),
        start_at=row["start_at"],
        target_at=row["target_at"],
        duration_days=float(row["duration_days"]),
        sale_price=Decimal(str(row["sale_price"] or 0)),
    )


def _service_label(value: str | None) -> str:
    if not value:
        return "Не указана"
    return SERVICE_LABELS.get(value, value)


def _average(values: Iterable[float]) -> float | None:
    materialized = list(values)
    return mean(materialized) if materialized else None


def _service_category_average(rows: Iterable[ProjectMetric]) -> float | None:
    """Match the stakeholder sheet: mean of the СОУТ and other-service means."""
    valid_rows = _duration_rows(rows)
    category_averages = [
        value
        for value in (
            _average(row.duration_days for row in valid_rows if row.service == "sout"),
            _average(row.duration_days for row in valid_rows if row.service != "sout"),
        )
        if value is not None
    ]
    return _average(category_averages)


def _duration_rows(rows: Iterable[ProjectMetric]) -> list[ProjectMetric]:
    return [row for row in rows if row.duration_days <= ANOMALY_DURATION_DAYS]


def _anomaly_rows(rows: Iterable[ProjectMetric]) -> list[ProjectMetric]:
    return [row for row in rows if row.duration_days > ANOMALY_DURATION_DAYS]


def _variance(actual: float | Decimal | None, plan: float | Decimal, *, lower_is_better: bool) -> float | None:
    if actual is None or float(plan) == 0:
        return None
    actual_value = float(actual)
    plan_value = float(plan)
    return (plan_value - actual_value) / plan_value if lower_is_better else (actual_value - plan_value) / plan_value


def _subset(rows: Sequence[ProjectMetric], names: Iterable[str]) -> list[ProjectMetric]:
    allowed = set(names)
    return [row for row in rows if row.specialist in allowed]


def _summary_values(data: SksReportData, names: Iterable[str]) -> dict[str, object]:
    agreements = _subset(data.agreements, names)
    sends = _subset(data.sends, names)
    return {
        "agreement_days": _average(row.duration_days for row in _duration_rows(agreements)),
        "send_days": _average(row.duration_days for row in _duration_rows(sends)),
        "approved_count": len(agreements),
        "approved_amount": sum((row.sale_price for row in agreements), Decimal(0)),
    }


def _title(sheet, text_value: str, period: str, columns: int) -> None:
    sheet.append([text_value] + [None] * (columns - 1), STYLE["title"])
    sheet.append([period] + [None] * (columns - 1), STYLE["subtitle"])
    sheet.merges.extend([f"A1:{_excel_column(columns)}1", f"A2:{_excel_column(columns)}2"])
    sheet.row_heights[1] = 26
    sheet.row_heights[2] = 22


def _excel_column(count: int) -> str:
    value = count
    result = ""
    while value:
        value, remainder = divmod(value - 1, 26)
        result = chr(65 + remainder) + result
    return result


@lru_cache(maxsize=32)
def _sks_font(size: int, bold: bool = False):
    asset_name = "ks-report-bold.ttf.b64" if bold else "ks-report-regular.ttf.b64"
    asset_path = Path(__file__).with_name("assets") / asset_name
    try:
        font_bytes = base64.b64decode(asset_path.read_text(encoding="ascii"))
        return ImageFont.truetype(io.BytesIO(font_bytes), size)
    except (OSError, ValueError):
        return ImageFont.load_default(size=size)


def _sks_days(value: object) -> str:
    if value is None:
        return "—"
    return f"{float(value):.1f} дня".replace(".", ",")


def _sks_money(value: Decimal | float | object) -> str:
    number = float(value)
    if abs(number) >= 1_000_000:
        return f"{number / 1_000_000:.2f}".replace(".", ",") + " млн руб."
    if abs(number) >= 1_000:
        return f"{number / 1_000:.0f} тыс. руб."
    return f"{number:.0f} руб."


def _sks_time(seconds: float | None) -> str:
    if seconds is None:
        return "—"
    rounded = max(0, round(seconds))
    hours, remainder = divmod(rounded, 3600)
    minutes, seconds_value = divmod(remainder, 60)
    return f"{hours:d}:{minutes:02d}:{seconds_value:02d}"


def _sks_plan_detail(actual: object, plan: object, *, lower_is_better: bool) -> tuple[str, str]:
    difference = _variance(actual, plan, lower_is_better=lower_is_better)
    if difference is None:
        return "нет данных для сравнения", "#7A8998"
    sign = "+" if difference >= 0 else ""
    label = f"к плану {sign}{difference * 100:.0f}%"
    return label, "#2E9C73" if difference >= 0 else "#C45A54"


def _draw_sks_card(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    label: str,
    value: str,
    detail: str,
    accent: str,
    detail_color: str = "#7A8998",
) -> None:
    left, top, right, bottom = box
    draw.rounded_rectangle(box, radius=18, fill="#F8FAFD", outline="#DCE5EF", width=2)
    draw.rounded_rectangle((left, top, left + 8, bottom), radius=4, fill=accent)
    fitted_text(draw, (left + 22, top + 15), label, _sks_font, width=right-left-40, size=18, fill="#5B6B7C")
    fitted_text(draw, (left + 22, top + 44), value, _sks_font, width=right-left-40, size=30, bold=True)
    fitted_text(draw, (left + 22, top + 82), detail, _sks_font, width=right-left-40, size=15, fill=detail_color)


def _draw_sks_duration_panel(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    data: SksReportData,
) -> None:
    left, top, right, bottom = box
    blue = "#255985"
    orange = "#E89A3C"
    draw.rounded_rectangle(box, radius=20, fill="#F8FAFD", outline="#DCE5EF", width=2)
    draw.text((left + 24, top + 18), "Скорость работы по специалистам", font=_sks_font(23, True), fill="#213547")
    draw.text((left + 24, top + 51), "Среднее число календарных дней · без аномалий свыше 200 дней", font=_sks_font(14), fill="#7A8998")
    draw.rectangle((left + 24, top + 83, left + 42, top + 98), fill=blue)
    draw.text((left + 52, top + 77), "Согласование", font=_sks_font(17), fill="#5B6B7C")
    draw.rectangle((left + 225, top + 83, left + 243, top + 98), fill=orange)
    draw.text((left + 253, top + 77), "Отправка", font=_sks_font(17), fill="#5B6B7C")

    rows: list[tuple[str, float | None, float | None]] = []
    for name in SCS_SPECIALISTS:
        values = _summary_values(data, (name,))
        rows.append((name, values["agreement_days"], values["send_days"]))
    maximum = max(
        [PLAN_SEND_DAYS, 1.0]
        + [float(value) for _, agreement, sending in rows for value in (agreement, sending) if value is not None]
    )
    chart_left = left + 290
    chart_right = right - 140
    chart_width = chart_right - chart_left
    row_top = top + 125
    row_height = 48
    plan_x = chart_left + chart_width * min(1.0, PLAN_SEND_DAYS / maximum)
    draw.line((plan_x, row_top - 8, plan_x, bottom - 38), fill="#AAB8C6", width=2)
    draw.text((min(plan_x, right - 230), bottom - 32), "план отправки: 4 дня", font=_sks_font(16), fill="#5B6B7C")
    for index, (name, agreement, sending) in enumerate(rows):
        y = row_top + index * row_height
        fitted_text(draw, (left + 24, y + 9), name, _sks_font, width=250, size=18, bold=True)
        draw.line((chart_left, y + 14, chart_right, y + 14), fill="#E5EBF2", width=13)
        draw.line((chart_left, y + 35, chart_right, y + 35), fill="#E5EBF2", width=13)
        for value, bar_y, color in ((agreement, y + 14, blue), (sending, y + 35, orange)):
            if value is None:
                draw.text((chart_right + 15, bar_y - 9), "—", font=_sks_font(17), fill="#5B6B7C")
                continue
            end_x = chart_left + chart_width * min(1.0, float(value) / maximum)
            draw.line((chart_left, bar_y, end_x, bar_y), fill=color, width=13)
            label = f"{float(value):.1f}".replace(".", ",")
            draw.text((chart_right + 15, bar_y - 11), label + " дн.", font=_sks_font(17, True), fill=color)


def _draw_sks_approval_panel(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    data: SksReportData,
) -> None:
    left, top, right, bottom = box
    draw.rounded_rectangle(box, radius=20, fill="#F8FAFD", outline="#DCE5EF", width=2)
    draw.text((left + 24, top + 18), "Согласованные отчёты", font=_sks_font(23, True), fill="#213547")
    draw.text((left + 24, top + 51), "Стоимость проектов и количество переходов на целевой этап", font=_sks_font(14), fill="#7A8998")
    rows: list[tuple[str, int, Decimal]] = []
    for name in SCS_SPECIALISTS:
        values = _summary_values(data, (name,))
        rows.append((name, int(values["approved_count"]), values["approved_amount"]))
    maximum = max([Decimal(1)] + [amount for _, _, amount in rows])
    chart_left = left + 290
    chart_right = right - 285
    chart_width = chart_right - chart_left
    row_top = top + 95
    row_height = 48
    colors = ("#255985", "#2E7896", "#35949A", "#2E9C73", "#75A85A", "#D39A3E", "#8A6CAB")
    for index, (name, count, amount) in enumerate(rows):
        y = row_top + index * row_height
        color = colors[index % len(colors)]
        fitted_text(draw, (left + 24, y + 5), name, _sks_font, width=250, size=18, bold=True)
        draw.line((chart_left, y + 15, chart_right, y + 15), fill="#E5EBF2", width=18)
        if amount > 0:
            end_x = chart_left + chart_width * float(amount / maximum)
            draw.line((chart_left, y + 15, end_x, y + 15), fill=color, width=18)
        fitted_text(draw, (chart_right + 16, y + 4), _sks_money(amount), _sks_font, width=170, size=19, bold=True)
        badge = f"{count} шт."
        badge_box = draw.textbbox((0, 0), badge, font=_sks_font(12, True))
        badge_width = badge_box[2] - badge_box[0] + 18
        draw.rounded_rectangle((right - badge_width - 20, y, right - 20, y + 30), radius=10, fill="#EAF0F6")
        draw.text((right - badge_width - 11, y + 5), badge, font=_sks_font(12, True), fill="#5B6B7C")


def build_sks_chart(data: SksReportData) -> bytes:
    image = Image.new("RGB", (1200, 2000), "#F4F7FB")
    draw = ImageDraw.Draw(image)
    blue = "#255985"
    orange = "#E89A3C"
    green = "#2E9C73"
    purple = "#7C5AA6"
    draw.rounded_rectangle((45, 35, 1155, 1965), radius=30, fill="#FFFFFF", outline="#DCE5EF", width=2)
    draw.text((85, 70), "СКС · Итоги работы", font=_sks_font(41, True), fill="#173A5E")
    draw.text((85, 126), f"{data.date_from:%d.%m.%Y}–{data.date_to:%d.%m.%Y} · данные CRM", font=_sks_font(21), fill="#5B6B7C")

    primary = _summary_values(data, PRIMARY_SPECIALISTS)
    task_average = _average(task.working_seconds for task in data.tasks)
    all_send_average = _service_category_average(data.sends)
    send_detail, send_color = _sks_plan_detail(all_send_average, PLAN_SEND_DAYS, lower_is_better=True)
    task_detail, task_color = _sks_plan_detail(task_average, PLAN_TASK_SECONDS, lower_is_better=True)
    count_detail, count_color = _sks_plan_detail(len(data.agreements), PLAN_APPROVED_COUNT, lower_is_better=False)
    amount_detail, amount_color = _sks_plan_detail(primary["approved_amount"], PLAN_APPROVED_AMOUNT, lower_is_better=False)
    cards = (
        ("Согласование · 6 спец.", _sks_days(primary["agreement_days"]), "среднее за период", blue, "#7A8998"),
        ("Отправка · вся выборка", _sks_days(all_send_average), send_detail, orange, send_color),
        ("Все задачи · Б/К", f"{data.assigned_sending_task_total} шт.", "Балакирева / Кулешева", purple, "#7A8998"),
        ("Задачи · договор/счет", _sks_time(task_average), task_detail, green, task_color),
        ("Согласовано · вся выборка", f"{len(data.agreements)} шт.", count_detail, blue, count_color),
        ("Сумма · 6 спец.", _sks_money(primary["approved_amount"]), amount_detail, orange, amount_color),
    )
    for index, card in enumerate(cards):
        column = index % 3
        row = index // 3
        left = 85 + column * 350
        top = 178 + row * 130
        _draw_sks_card(draw, (left, top, left + 330, top + 112), *card)

    _draw_sks_duration_panel(draw, (85, 455, 1105, 1025), data)
    _draw_sks_approval_panel(draw, (85, 1050, 1105, 1555), data)

    sout_rows = [row for row in data.agreements if row.service == "sout"]
    other_rows = [row for row in data.agreements if row.service != "sout"]
    anomaly_agreement = len(_anomaly_rows(data.agreements))
    anomaly_send = len(_anomaly_rows(data.sends))
    sending_diana = data.sending_tasks.get("Балакирева Диана", 0)
    sending_vladislava = data.sending_tasks.get("Кулешева Владислава", 0)
    operational_cards = (
        ("СОУТ · вся выборка", f"{len(sout_rows)} шт.", "согласовано за период", blue),
        ("Другие · вся выборка", f"{len(other_rows)} шт.", "согласовано за период", green),
        ("Задачи «отправка»", f"{sending_diana + sending_vladislava} шт.", f"Балакирева {sending_diana} · Кулешева {sending_vladislava}", orange),
        ("Аномалии > 200 дней", f"{anomaly_agreement + anomaly_send} записей", f"согласование {anomaly_agreement} · отправка {anomaly_send}", purple),
    )
    draw.text((85, 1590), "Операционный контроль", font=_sks_font(23, True), fill="#213547")
    for index, (label, value, detail, accent) in enumerate(operational_cards):
        left = 85 + index * 255
        _draw_sks_card(draw, (left, 1635, left + 245, 1755), label, value, detail, accent)

    draw.rounded_rectangle((85, 1790, 1105, 1920), radius=18, fill="#F1F5F9")
    draw.text((108, 1810), "Как читать показатели", font=_sks_font(17, True), fill="#213547")
    draw.text((108, 1841), "• Положительный процент означает результат лучше плана. План применяется к выбранному периоду без пересчёта.", font=_sks_font(14), fill="#5B6B7C")
    draw.text((108, 1868), "• Среднее отправки = среднее между показателями СОУТ и остальных услуг, как в Google-таблице.", font=_sks_font(14), fill="#5B6B7C")
    draw.text((108, 1895), "• Задачи: Пн–Пт 08:00–17:30; календарный срок от 14 ч исключается до расчёта.", font=_sks_font(14), fill="#5B6B7C")
    draw.text((85, 1932), "EcoStar Reports · подробности и список проектов — в Excel", font=_sks_font(16), fill="#7A8998")

    output = io.BytesIO()
    image.save(output, format="PNG", optimize=True)
    return output.getvalue()


def build_sks_workbook(data: SksReportData) -> bytes:
    period = f"Период: {data.date_from:%d.%m.%Y}–{data.date_to:%d.%m.%Y}"
    workbook = Workbook(f"Отчет СКС {period}")
    _build_summary_sheet(workbook, data, period)
    _build_metric_sheet(workbook, "Согласование", data.agreements, period, specialists=SCS_SPECIALISTS, total_label="Итого СКС (8 специалистов)")
    sending_specialists = tuple(sorted({row.specialist for row in data.sends}))
    _build_metric_sheet(workbook, "Отправка", data.sends, period, specialists=sending_specialists, total_label="Итого: все СКС")
    _build_task_sheet(workbook, data, period)
    _build_detail_sheet(workbook, "Проекты — согласование", _subset(data.agreements, SCS_SPECIALISTS), period)
    _build_detail_sheet(workbook, "Проекты — отправка", data.sends, period)
    _build_anomaly_sheet(workbook, data, period)
    return workbook.to_bytes()


def _build_summary_sheet(workbook: Workbook, data: SksReportData, period: str) -> None:
    sheet = workbook.add_sheet("Сводка")
    _title(sheet, "Отчёт СКС", period, 8)
    sheet.append([None] * 8)
    sheet.append(["Ключевые показатели · область расчёта указана в названии"] + [None] * 7, STYLE["section"])
    sheet.merges.append("A4:H4")
    sheet.append(
        ["Показатель", "План", "Факт", "Отклонение к плану", "Единица", "Метод", None, None],
        STYLE["table_header"],
    )

    primary = _summary_values(data, PRIMARY_SPECIALISTS)
    task_average = _average(task.working_seconds for task in data.tasks)
    all_send_average = _service_category_average(data.sends)
    kpis = [
        ("Среднее время отправки · все СКС", PLAN_SEND_DAYS, all_send_average, True, "дни", "меньше — лучше", "days"),
        ("Среднее время задачи · < 14 часов", PLAN_TASK_SECONDS / 86400, None if task_average is None else task_average / 86400, True, "ч:м:с", "меньше — лучше", "time"),
        ("Согласовано отчетов · все СКС", PLAN_APPROVED_COUNT, len(data.agreements), False, "шт.", "больше — лучше", "number"),
        ("Согласовано отчетов · основные 6", PLAN_APPROVED_AMOUNT, primary["approved_amount"], False, "руб.", "больше — лучше", "money"),
    ]
    for label, plan, actual, lower, unit, method, kind in kpis:
        value_style = {
            "days": STYLE["table_days"],
            "time": STYLE["table_time"],
            "number": STYLE["table_number"],
            "money": STYLE["table_money"],
        }[kind]
        sheet.append(
            [label, plan, actual if actual is not None else "—", _variance(actual, plan, lower_is_better=lower), unit, method, None, None],
            [STYLE["table_text"], value_style, value_style, STYLE["table_percent"], STYLE["table_center"], STYLE["table_text"], STYLE["base"], STYLE["base"]],
        )

    sheet.append([None] * 8)
    sheet.append(["Согласование по специалистам СКС"] + [None] * 7, STYLE["section"])
    sheet.merges.append("A11:H11")
    sheet.append(
        ["Специалист", "Согласование, дни", "Согласовано, шт.", "% к плану 43", "Согласовано, руб.", "% к плану 1 500 000", None, None],
        STYLE["table_header"],
    )
    for name in (*SCS_SPECIALISTS, "Итого СКС (8 специалистов)"):
        names = SCS_SPECIALISTS if name == "Итого СКС (8 специалистов)" else (name,)
        values = _summary_values(data, names)
        total = name == "Итого СКС (8 специалистов)"
        styles = [
            STYLE["total_text"] if total else STYLE["table_text"],
            STYLE["total_days"] if total else STYLE["table_days"],
            STYLE["total_number"] if total else STYLE["table_number"],
            STYLE["total_percent"] if total else STYLE["table_percent"],
            STYLE["total_money"] if total else STYLE["table_money"],
            STYLE["total_percent"] if total else STYLE["table_percent"],
            STYLE["base"],
            STYLE["base"],
        ]
        sheet.append(
            [
                name,
                values["agreement_days"] if values["agreement_days"] is not None else "—",
                values["approved_count"],
                _variance(values["approved_count"], PLAN_APPROVED_COUNT, lower_is_better=False),
                values["approved_amount"],
                _variance(values["approved_amount"], PLAN_APPROVED_AMOUNT, lower_is_better=False),
                None,
                None,
            ],
            styles,
        )

    sheet.append([None] * 8)
    sheet.append(["Контроль согласованных отчётов · количество по всем СКС, деньги и сроки по основным 6"] + [None] * 7, STYLE["section"])
    sheet.merges.append(f"A{len(sheet.rows)}:H{len(sheet.rows)}")
    sheet.append(
        ["Категория", "Все СКС, шт.", "Основные 6, руб.", "Основные 6, дни", "Аномалий > 200 дней · основные 6", None, None, None],
        STYLE["table_header"],
    )
    primary_agreements = _subset(data.agreements, PRIMARY_SPECIALISTS)
    for label, all_rows, primary_rows in (
        ("СОУТ", [row for row in data.agreements if row.service == "sout"], [row for row in primary_agreements if row.service == "sout"]),
        ("Другие услуги", [row for row in data.agreements if row.service != "sout"], [row for row in primary_agreements if row.service != "sout"]),
        ("Итого", list(data.agreements), primary_agreements),
    ):
        valid_rows = _duration_rows(primary_rows)
        total = label == "Итого"
        sheet.append(
            [
                label,
                len(all_rows),
                sum((row.sale_price for row in primary_rows), Decimal(0)),
                _average(row.duration_days for row in valid_rows) if valid_rows else "—",
                len(_anomaly_rows(primary_rows)),
                None,
                None,
                None,
            ],
            [
                STYLE["total_text"] if total else STYLE["table_text"],
                STYLE["total_number"] if total else STYLE["table_number"],
                STYLE["total_money"] if total else STYLE["table_money"],
                STYLE["total_days"] if total else STYLE["table_days"],
                STYLE["total_number"] if total else STYLE["table_number"],
                STYLE["base"],
                STYLE["base"],
                STYLE["base"],
            ],
        )

    sheet.append([None] * 8)
    sheet.append(["Отправка · задачи исполнителей Балакирева Диана и Кулешева Владислава"] + [None] * 7, STYLE["section"])
    sheet.merges.append(f"A{len(sheet.rows)}:H{len(sheet.rows)}")
    sheet.append(
        ["Исполнитель", "Закрыто задач, шт.", None, None, "Единица", "Контроль", None, None],
        STYLE["table_header"],
    )
    for name in (*SENDING_TASK_SPECIALISTS, "Итого"):
        total = name == "Итого"
        task_count = data.assigned_sending_task_total if total else data.assigned_sending_tasks.get(name, 0)
        sheet.append(
            [name, task_count, None, None, "задач", "Итого без двойного счёта", None, None],
            [STYLE["total_text"] if total else STYLE["table_text"], STYLE["total_number"] if total else STYLE["table_number"], STYLE["base"], STYLE["base"], STYLE["total_text"] if total else STYLE["table_text"], STYLE["total_text"] if total else STYLE["table_text"], STYLE["base"], STYLE["base"]],
        )

    sheet.append([None] * 8)
    sheet.append(["Закрытые задачи с точным названием «отправка»"] + [None] * 7, STYLE["section"])
    sheet.merges.append(f"A{len(sheet.rows)}:H{len(sheet.rows)}")
    for name in SENDING_TASK_SPECIALISTS:
        sheet.append(
            [name, data.sending_tasks.get(name, 0), "шт.", None, None, None, None, None],
            [STYLE["table_text"], STYLE["table_number"], STYLE["table_center"], STYLE["base"], STYLE["base"], STYLE["base"], STYLE["base"], STYLE["base"]],
        )

    sheet.append([None] * 8)
    note_row = sheet.append(
        [
            "Отклонение: положительное значение означает выполнение лучше плана. Как в Google-таблице: среднее время отправки — среднее из показателей СОУТ и остальных услуг по всем СКС; количество согласованных — по всем СКС; сумма — по основным 6 специалистам. В блоке отправки считаются все закрытые задачи Дианы и Владиславы, итог без двойного счёта задач, назначенных обеим. Проекты свыше 200 дней исключаются из времени. Задачи «Договор/Счет» длительностью 14 часов и более исключаются полностью.",
            None, None, None, None, None, None, None,
        ],
        STYLE["note"],
    )
    sheet.merges.append(f"A{note_row}:H{note_row}")
    sheet.row_heights[note_row] = 34
    sheet.widths = {0: 31, 1: 19, 2: 19, 3: 19, 4: 17, 5: 23, 6: 22, 7: 22}
    sheet.freeze_rows = 12


def _build_metric_sheet(
    workbook: Workbook,
    title: str,
    rows: Sequence[ProjectMetric],
    period: str,
    *,
    specialists: Sequence[str],
    total_label: str,
) -> None:
    sheet = workbook.add_sheet(title)
    _title(sheet, title, period, 8)
    sheet.append([None] * 8)
    sheet.append(
        ["Специалист", "Услуга", "Проектов, шт.", "Аномалий > 200 дней", "Среднее, дни", "Минимум, дни", "Максимум, дни", "Сумма проектов, руб."],
        STYLE["table_header"],
    )
    groups: dict[tuple[str, str], list[ProjectMetric]] = defaultdict(list)
    visible_names = set(specialists)
    for row in rows:
        if row.specialist in visible_names:
            groups[(row.specialist, _service_label(row.service))].append(row)

    for specialist in specialists:
        specialist_groups = sorted(
            ((service, values) for (name, service), values in groups.items() if name == specialist),
            key=lambda item: item[0],
        )
        if not specialist_groups:
            sheet.append([specialist, "—", 0, 0, "—", "—", "—", 0], [STYLE["table_text"], STYLE["table_text"], STYLE["table_number"], STYLE["table_number"], STYLE["table_days"], STYLE["table_days"], STYLE["table_days"], STYLE["table_money"]])
            continue
        for service, values in specialist_groups:
            durations = [value.duration_days for value in _duration_rows(values)]
            sheet.append(
                [specialist, service, len(values), len(_anomaly_rows(values)), mean(durations) if durations else "—", min(durations) if durations else "—", max(durations) if durations else "—", sum((value.sale_price for value in values), Decimal(0))],
                [STYLE["table_text"], STYLE["table_text"], STYLE["table_number"], STYLE["table_number"], STYLE["table_days"], STYLE["table_days"], STYLE["table_days"], STYLE["table_money"]],
            )

    selected_rows = _subset(rows, specialists)
    durations = [row.duration_days for row in _duration_rows(selected_rows)]
    sheet.append(
        [total_label, "Все услуги", len(selected_rows), len(_anomaly_rows(selected_rows)), mean(durations) if durations else "—", min(durations) if durations else "—", max(durations) if durations else "—", sum((row.sale_price for row in selected_rows), Decimal(0))],
        [STYLE["total_text"], STYLE["total_text"], STYLE["total_number"], STYLE["total_number"], STYLE["total_days"], STYLE["total_days"], STYLE["total_days"], STYLE["total_money"]],
    )
    sheet.widths = {0: 28, 1: 22, 2: 16, 3: 22, 4: 18, 5: 18, 6: 18, 7: 23}
    sheet.freeze_rows = 4
    sheet.auto_filter = f"A4:H{len(sheet.rows)}"


def _build_task_sheet(workbook: Workbook, data: SksReportData, period: str) -> None:
    sheet = workbook.add_sheet("Задачи")
    _title(sheet, "Задачи", period, 7)
    sheet.append([None] * 7)
    sheet.append(
        ["Название", "Закрыто, шт.", "Среднее рабочее время", "План", "Отклонение к плану", "Минимум", "Максимум"],
        STYLE["table_header"],
    )
    seconds = [task.working_seconds for task in data.tasks]
    average = mean(seconds) if seconds else None
    sheet.append(
        [
            "Договор и счет",
            len(data.tasks),
            "—" if average is None else average / 86400,
            PLAN_TASK_SECONDS / 86400,
            _variance(average, PLAN_TASK_SECONDS, lower_is_better=True),
            "—" if not seconds else min(seconds) / 86400,
            "—" if not seconds else max(seconds) / 86400,
        ],
        [
            STYLE["total_text"],
            STYLE["total_number"],
            STYLE["total_time"],
            STYLE["total_time"],
            STYLE["total_percent"],
            STYLE["total_time"],
            STYLE["total_time"],
        ],
    )

    sheet.append([None] * 7)
    section = sheet.append(["Закрытые задачи с точным названием «отправка»"] + [None] * 6, STYLE["section"])
    sheet.merges.append(f"A{section}:G{section}")
    sheet.append(["Специалист", "Закрыто, шт.", None, None, None, None, None], STYLE["table_header"])
    for name in SENDING_TASK_SPECIALISTS:
        sheet.append([name, data.sending_tasks.get(name, 0), None, None, None, None, None], [STYLE["table_text"], STYLE["table_number"], STYLE["base"], STYLE["base"], STYLE["base"], STYLE["base"], STYLE["base"]])

    sheet.append([None] * 7)
    note = sheet.append(
        ["Рабочее время рассчитано по будням 08:00–17:30 (Москва); ночи и выходные исключены. Задачи с полной календарной длительностью 14 часов и более полностью исключены. Учитываются выполненные задачи, название которых содержит «Договор» или «Счет», но не содержит «ЭДО»; регистр, ё/е и лишние пробелы не влияют." ] + [None] * 6,
        STYLE["note"],
    )
    sheet.merges.append(f"A{note}:G{note}")
    sheet.row_heights[note] = 58
    sheet.widths = {0: 29, 1: 17, 2: 23, 3: 18, 4: 22, 5: 18, 6: 18}
    sheet.freeze_rows = 4


def _build_detail_sheet(workbook: Workbook, title: str, rows: Sequence[ProjectMetric], period: str) -> None:
    sheet = workbook.add_sheet(title)
    _title(sheet, title, period, 9)
    sheet.append([None] * 9)
    sheet.append(
        ["Проект ID", "Договор", "Услуга", "Специалист СКС", "Начало этапа", "Целевой этап", "Дней", "Статус длительности", "Стоимость, руб."],
        STYLE["table_header"],
    )
    for row in rows:
        sheet.append(
            [row.project_id, row.contract_number or "—", _service_label(row.service), row.specialist, row.start_at, row.target_at, row.duration_days, "Аномалия > 200 дней" if row.duration_days > ANOMALY_DURATION_DAYS else "Учитывается", row.sale_price],
            [STYLE["table_text"], STYLE["table_text"], STYLE["table_text"], STYLE["table_text"], STYLE["table_center"], STYLE["table_center"], STYLE["table_days"], STYLE["table_text"], STYLE["table_money"]],
        )
    sheet.widths = {0: 38, 1: 22, 2: 18, 3: 28, 4: 21, 5: 21, 6: 14, 7: 24, 8: 21}
    sheet.freeze_rows = 4
    sheet.auto_filter = f"A4:I{max(4, len(sheet.rows))}"


def _build_anomaly_sheet(workbook: Workbook, data: SksReportData, period: str) -> None:
    sheet = workbook.add_sheet("Аномалии")
    _title(sheet, "Аномалии длительности", period, 10)
    sheet.append([None] * 10)
    sheet.append(
        ["Показатель", "Проект ID", "Договор", "Услуга", "Специалист СКС", "Начало этапа", "Целевой этап", "Дней", "Стоимость, руб.", "Учёт"],
        STYLE["table_header"],
    )
    rows = [
        *(("Согласование", row) for row in _anomaly_rows(_subset(data.agreements, SCS_SPECIALISTS))),
        *(("Отправка", row) for row in _anomaly_rows(data.sends)),
    ]
    for metric, row in sorted(rows, key=lambda item: (item[0], -item[1].duration_days, item[1].project_id)):
        sheet.append(
            [metric, row.project_id, row.contract_number or "—", _service_label(row.service), row.specialist, row.start_at, row.target_at, row.duration_days, row.sale_price, "Исключён только из статистики времени"],
            [STYLE["table_text"], STYLE["table_text"], STYLE["table_text"], STYLE["table_text"], STYLE["table_text"], STYLE["table_center"], STYLE["table_center"], STYLE["table_days"], STYLE["table_money"], STYLE["table_text"]],
        )
    if not rows:
        sheet.append(["Аномалий за выбранный период нет"] + [None] * 9, STYLE["note"])
        sheet.merges.append("A5:J5")
    sheet.widths = {0: 18, 1: 38, 2: 22, 3: 18, 4: 28, 5: 21, 6: 21, 7: 14, 8: 21, 9: 38}
    sheet.freeze_rows = 4
    sheet.auto_filter = f"A4:J{max(4, len(sheet.rows))}"
