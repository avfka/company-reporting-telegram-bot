from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from decimal import Decimal
from statistics import mean
from typing import Iterable, Sequence

from sqlalchemy import create_engine, text
from sqlalchemy.pool import NullPool

from reporting_bot.config import Settings
from reporting_bot.simple_xlsx import STYLE, Workbook


PRIMARY_SPECIALISTS = (
    "Иванова Елена",
    "Максимович Анастасия",
    "Хасанова Азалия",
    "Орлова Полина",
    "Сластина Ангелина",
    "Григорьева Алёна",
)
SPECIAL_SPECIALIST = "Шергина Надежда"
SENDING_TASK_SPECIALISTS = ("Балакирева Диана", "Кулешева Владислава")

PLAN_SEND_DAYS = 4.0
PLAN_TASK_SECONDS = 30 * 60
PLAN_APPROVED_COUNT = 43
PLAN_APPROVED_AMOUNT = Decimal("1500000")

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
      (p.service = 'sout' AND h.new_step = 'Актуализировать РМ / Прикрепить декларацию')
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
    p.manager_sks_id,
    p.contract_number,
    p.sale_price,
    row_number() OVER (PARTITION BY h.project_id ORDER BY h.created_at, h.id) AS target_rank
  FROM projects_steps_history h
  JOIN projects p ON p.id = h.project_id
  WHERE h.created_at >= CAST(:date_from AS DATE)
    AND h.created_at < CAST(:date_to_exclusive AS DATE)
    AND (
      (p.service = 'sout' AND h.new_step = 'Выгрузить протоколы в ФСА')
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
)
SELECT
  CASE
    WHEN normalized_title IN ('договор', 'договор и счет', 'счет и договор') THEN 'Договор'
    ELSE 'Счет'
  END AS category,
  created_at,
  closed_at
FROM completed_tasks
WHERE normalized_title IN ('договор', 'договор и счет', 'счет и договор', 'счет')
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


def working_seconds(start: datetime, end: datetime) -> float:
    """Return overlap with Moscow workdays 09:00–17:30, excluding weekends."""
    if end <= start:
        return 0.0
    total = 0.0
    current_day = start.date()
    while current_day <= end.date():
        if current_day.weekday() < 5:
            work_start = datetime.combine(current_day, time(9, 0))
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

    def create(self, date_from: date, date_to: date) -> tuple[bytes, str]:
        if date_to < date_from:
            raise ValueError("Дата окончания не может быть раньше даты начала.")
        if (date_to - date_from).days > 366:
            raise ValueError("Максимальный период отчёта — 366 дней.")
        data = self._load(date_from, date_to)
        filename = f"Отчет_СКС_{date_from.isoformat()}_{date_to.isoformat()}.xlsx"
        return build_sks_workbook(data), filename

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
            return SksReportData(
                date_from=date_from,
                date_to=date_to,
                agreements=agreements,
                sends=sends,
                tasks=tasks,
                sending_tasks=sending_tasks,
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
        "agreement_days": _average(row.duration_days for row in agreements),
        "send_days": _average(row.duration_days for row in sends),
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


def build_sks_workbook(data: SksReportData) -> bytes:
    period = f"Период: {data.date_from:%d.%m.%Y}–{data.date_to:%d.%m.%Y}"
    workbook = Workbook(f"Отчет СКС {period}")
    _build_summary_sheet(workbook, data, period)
    _build_metric_sheet(workbook, "Согласование", data.agreements, period, include_primary_total=True)
    _build_metric_sheet(workbook, "Отправка", data.sends, period, include_primary_total=True, include_all_total=True)
    _build_task_sheet(workbook, data, period)
    _build_detail_sheet(workbook, "Проекты — согласование", data.agreements, period)
    _build_detail_sheet(workbook, "Проекты — отправка", data.sends, period)
    return workbook.to_bytes()


def _build_summary_sheet(workbook: Workbook, data: SksReportData, period: str) -> None:
    sheet = workbook.add_sheet("Сводка")
    _title(sheet, "Отчёт СКС", period, 8)
    sheet.append([None] * 8)
    sheet.append(["Ключевые показатели · основные 6 специалистов"] + [None] * 7, STYLE["section"])
    sheet.merges.append("A4:H4")
    sheet.append(
        ["Показатель", "План", "Факт", "Отклонение к плану", "Единица", "Метод", None, None],
        STYLE["table_header"],
    )

    primary = _summary_values(data, PRIMARY_SPECIALISTS)
    task_average = _average(task.working_seconds for task in data.tasks)
    kpis = [
        ("Среднее время отправки", PLAN_SEND_DAYS, primary["send_days"], True, "дни", "меньше — лучше", "days"),
        ("Среднее время выполнения задачи", PLAN_TASK_SECONDS / 86400, None if task_average is None else task_average / 86400, True, "ч:м:с", "меньше — лучше", "time"),
        ("Согласовано отчетов", PLAN_APPROVED_COUNT, primary["approved_count"], False, "шт.", "больше — лучше", "number"),
        ("Согласовано отчетов", PLAN_APPROVED_AMOUNT, primary["approved_amount"], False, "руб.", "больше — лучше", "money"),
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
    sheet.append(["Показатели по специалистам"] + [None] * 7, STYLE["section"])
    sheet.merges.append("A11:H11")
    sheet.append(
        ["Специалист", "Согласование, дни", "Отправка, дни", "% к плану 4 дня", "Согласовано, шт.", "% к плану 43", "Согласовано, руб.", "% к плану 1 500 000"],
        STYLE["table_header"],
    )
    for name in (*PRIMARY_SPECIALISTS, "Итого 6 специалистов", SPECIAL_SPECIALIST):
        names = PRIMARY_SPECIALISTS if name == "Итого 6 специалистов" else (name,)
        values = _summary_values(data, names)
        total = name == "Итого 6 специалистов"
        styles = [
            STYLE["total_text"] if total else STYLE["table_text"],
            STYLE["total_days"] if total else STYLE["table_days"],
            STYLE["total_days"] if total else STYLE["table_days"],
            STYLE["total_percent"] if total else STYLE["table_percent"],
            STYLE["total_number"] if total else STYLE["table_number"],
            STYLE["total_percent"] if total else STYLE["table_percent"],
            STYLE["total_money"] if total else STYLE["table_money"],
            STYLE["total_percent"] if total else STYLE["table_percent"],
        ]
        sheet.append(
            [
                name,
                values["agreement_days"] if values["agreement_days"] is not None else "—",
                values["send_days"] if values["send_days"] is not None else "—",
                _variance(values["send_days"], PLAN_SEND_DAYS, lower_is_better=True),
                values["approved_count"],
                _variance(values["approved_count"], PLAN_APPROVED_COUNT, lower_is_better=False),
                values["approved_amount"],
                _variance(values["approved_amount"], PLAN_APPROVED_AMOUNT, lower_is_better=False),
            ],
            styles,
        )

    sheet.append([None] * 8)
    sheet.append(["Отправка · все специалисты СКС, включая агентский канал"] + [None] * 7, STYLE["section"])
    sheet.merges.append(f"A{len(sheet.rows)}:H{len(sheet.rows)}")
    all_send_avg = _average(row.duration_days for row in data.sends)
    sheet.append(
        ["Все услуги", len(data.sends), all_send_avg if all_send_avg is not None else "—", _variance(all_send_avg, PLAN_SEND_DAYS, lower_is_better=True), "проектов", None, None, None],
        [STYLE["total_text"], STYLE["total_number"], STYLE["total_days"], STYLE["total_percent"], STYLE["total_text"], STYLE["base"], STYLE["base"], STYLE["base"]],
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
            "Отклонение: положительное значение означает выполнение лучше плана. Для времени меньше — лучше; для количества и суммы больше — лучше. План применяется к выбранному периоду без пересчёта.",
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
    include_primary_total: bool,
    include_all_total: bool = False,
) -> None:
    sheet = workbook.add_sheet(title)
    _title(sheet, title, period, 7)
    sheet.append([None] * 7)
    sheet.append(
        ["Специалист", "Услуга", "Проектов, шт.", "Среднее, дни", "Минимум, дни", "Максимум, дни", "Сумма проектов, руб."],
        STYLE["table_header"],
    )
    groups: dict[tuple[str, str], list[ProjectMetric]] = defaultdict(list)
    visible_names = set(PRIMARY_SPECIALISTS) | {SPECIAL_SPECIALIST}
    for row in rows:
        if row.specialist in visible_names:
            groups[(row.specialist, _service_label(row.service))].append(row)

    for specialist in (*PRIMARY_SPECIALISTS, SPECIAL_SPECIALIST):
        specialist_groups = sorted(
            ((service, values) for (name, service), values in groups.items() if name == specialist),
            key=lambda item: item[0],
        )
        if not specialist_groups:
            sheet.append([specialist, "—", 0, "—", "—", "—", 0], [STYLE["table_text"], STYLE["table_text"], STYLE["table_number"], STYLE["table_days"], STYLE["table_days"], STYLE["table_days"], STYLE["table_money"]])
            continue
        for service, values in specialist_groups:
            durations = [value.duration_days for value in values]
            sheet.append(
                [specialist, service, len(values), mean(durations), min(durations), max(durations), sum((value.sale_price for value in values), Decimal(0))],
                [STYLE["table_text"], STYLE["table_text"], STYLE["table_number"], STYLE["table_days"], STYLE["table_days"], STYLE["table_days"], STYLE["table_money"]],
            )

    if include_primary_total:
        primary = _subset(rows, PRIMARY_SPECIALISTS)
        durations = [row.duration_days for row in primary]
        sheet.append(
            ["Итого 6 специалистов", "Все услуги", len(primary), mean(durations) if durations else "—", min(durations) if durations else "—", max(durations) if durations else "—", sum((row.sale_price for row in primary), Decimal(0))],
            [STYLE["total_text"], STYLE["total_text"], STYLE["total_number"], STYLE["total_days"], STYLE["total_days"], STYLE["total_days"], STYLE["total_money"]],
        )
    if include_all_total:
        durations = [row.duration_days for row in rows]
        sheet.append(
            ["Все СКС + агентский канал", "Все услуги", len(rows), mean(durations) if durations else "—", min(durations) if durations else "—", max(durations) if durations else "—", sum((row.sale_price for row in rows), Decimal(0))],
            [STYLE["total_text"], STYLE["total_text"], STYLE["total_number"], STYLE["total_days"], STYLE["total_days"], STYLE["total_days"], STYLE["total_money"]],
        )
    sheet.widths = {0: 28, 1: 22, 2: 16, 3: 18, 4: 18, 5: 18, 6: 23}
    sheet.freeze_rows = 4
    sheet.auto_filter = f"A4:G{len(sheet.rows)}"


def _build_task_sheet(workbook: Workbook, data: SksReportData, period: str) -> None:
    sheet = workbook.add_sheet("Задачи")
    _title(sheet, "Задачи", period, 7)
    sheet.append([None] * 7)
    sheet.append(
        ["Название", "Закрыто, шт.", "Среднее рабочее время", "План", "Отклонение к плану", "Минимум", "Максимум"],
        STYLE["table_header"],
    )
    for category in ("Договор", "Счет", "Итого"):
        values = list(data.tasks) if category == "Итого" else [task for task in data.tasks if task.category == category]
        seconds = [task.working_seconds for task in values]
        average = mean(seconds) if seconds else None
        total = category == "Итого"
        styles = [
            STYLE["total_text"] if total else STYLE["table_text"],
            STYLE["total_number"] if total else STYLE["table_number"],
            STYLE["total_time"] if total else STYLE["table_time"],
            STYLE["total_time"] if total else STYLE["table_time"],
            STYLE["total_percent"] if total else STYLE["table_percent"],
            STYLE["total_time"] if total else STYLE["table_time"],
            STYLE["total_time"] if total else STYLE["table_time"],
        ]
        sheet.append(
            [
                category,
                len(values),
                "—" if average is None else average / 86400,
                PLAN_TASK_SECONDS / 86400,
                _variance(average, PLAN_TASK_SECONDS, lower_is_better=True),
                "—" if not seconds else min(seconds) / 86400,
                "—" if not seconds else max(seconds) / 86400,
            ],
            styles,
        )

    sheet.append([None] * 7)
    section = sheet.append(["Закрытые задачи с точным названием «отправка»"] + [None] * 6, STYLE["section"])
    sheet.merges.append(f"A{section}:G{section}")
    sheet.append(["Специалист", "Закрыто, шт.", None, None, None, None, None], STYLE["table_header"])
    for name in SENDING_TASK_SPECIALISTS:
        sheet.append([name, data.sending_tasks.get(name, 0), None, None, None, None, None], [STYLE["table_text"], STYLE["table_number"], STYLE["base"], STYLE["base"], STYLE["base"], STYLE["base"], STYLE["base"]])

    sheet.append([None] * 7)
    note = sheet.append(
        ["Рабочее время рассчитано по будням 09:00–17:30 (Москва); ночи и выходные исключены. Включены выполненные задачи «Договор», «Счет», «Договор и счет» и «Счет и договор»; регистр, ё/е и лишние пробелы не влияют. Комбинированные названия относятся к категории «Договор»." ] + [None] * 6,
        STYLE["note"],
    )
    sheet.merges.append(f"A{note}:G{note}")
    sheet.row_heights[note] = 46
    sheet.widths = {0: 29, 1: 17, 2: 23, 3: 18, 4: 22, 5: 18, 6: 18}
    sheet.freeze_rows = 4


def _build_detail_sheet(workbook: Workbook, title: str, rows: Sequence[ProjectMetric], period: str) -> None:
    sheet = workbook.add_sheet(title)
    _title(sheet, title, period, 8)
    sheet.append([None] * 8)
    sheet.append(
        ["Проект ID", "Договор", "Услуга", "Специалист СКС", "Начало этапа", "Целевой этап", "Дней", "Стоимость, руб."],
        STYLE["table_header"],
    )
    for row in rows:
        sheet.append(
            [row.project_id, row.contract_number or "—", _service_label(row.service), row.specialist, row.start_at, row.target_at, row.duration_days, row.sale_price],
            [STYLE["table_text"], STYLE["table_text"], STYLE["table_text"], STYLE["table_text"], STYLE["table_center"], STYLE["table_center"], STYLE["table_days"], STYLE["table_money"]],
        )
    sheet.widths = {0: 38, 1: 22, 2: 18, 3: 28, 4: 21, 5: 21, 6: 14, 7: 21}
    sheet.freeze_rows = 4
    sheet.auto_filter = f"A4:H{max(4, len(sheet.rows))}"
