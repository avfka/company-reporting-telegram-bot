import io
import zipfile
from datetime import date, datetime
from decimal import Decimal

from PIL import Image

from reporting_bot.sks_report import (
    AGREEMENT_SQL,
    ANOMALY_DURATION_DAYS,
    ProjectMetric,
    SENDING_SQL,
    SksReportData,
    TASK_DURATION_SQL,
    TaskMetric,
    _summary_values,
    build_sks_chart,
    build_sks_workbook,
    working_seconds,
)


def test_working_seconds_excludes_nights_and_weekends() -> None:
    start = datetime(2026, 7, 3, 17, 0)  # Friday
    end = datetime(2026, 7, 6, 9, 30)  # Monday
    assert working_seconds(start, end) == 60 * 60


def test_working_seconds_caps_to_workday() -> None:
    assert working_seconds(
        datetime(2026, 7, 6, 8, 0),
        datetime(2026, 7, 6, 18, 0),
    ) == 8.5 * 60 * 60


def test_task_query_matches_titles_containing_contract_or_invoice_except_edo() -> None:
    assert "'Договор и счет' AS category" in TASK_DURATION_SQL
    assert "regexp_replace(trim(title), '\\s+', ' ', 'g')" in TASK_DURATION_SQL
    assert "'ё'," in TASK_DURATION_SQL
    assert "normalized_title LIKE '%договор%'" in TASK_DURATION_SQL
    assert "normalized_title LIKE '%счет%'" in TASK_DURATION_SQL
    assert "normalized_title NOT LIKE '%эдо%'" in TASK_DURATION_SQL


def test_sks_stage_queries_keep_close_project_and_accept_stage_label_variants() -> None:
    assert "h.new_step LIKE 'Актуализировать РМ%'" in AGREEMENT_SQL
    assert "h.new_step LIKE 'Выгрузить протоколы%ФСА'" in SENDING_SQL
    assert "h.new_step = 'Закрыть проект'" in SENDING_SQL


def test_build_sks_workbook_creates_valid_xlsx_package() -> None:
    metric = ProjectMetric(
        project_id="project-1",
        contract_number="СКС-1",
        service="sout",
        specialist="Иванова Елена",
        start_at=datetime(2026, 7, 1, 9, 0),
        target_at=datetime(2026, 7, 3, 12, 0),
        duration_days=2.125,
        sale_price=Decimal("125000.50"),
    )
    task = TaskMetric(
        category="Договор и счет",
        created_at=datetime(2026, 7, 1, 17, 0),
        closed_at=datetime(2026, 7, 2, 9, 30),
        working_seconds=3600,
    )
    content = build_sks_workbook(
        SksReportData(
            date_from=date(2026, 7, 1),
            date_to=date(2026, 7, 31),
            agreements=(metric,),
            sends=(metric,),
            tasks=(task,),
            sending_tasks={"Балакирева Диана": 3, "Кулешева Владислава": 4},
        )
    )

    with zipfile.ZipFile(io.BytesIO(content)) as archive:
        assert archive.testzip() is None
        assert len([name for name in archive.namelist() if name.startswith("xl/worksheets/sheet")]) == 7
        workbook_xml = archive.read("xl/workbook.xml").decode("utf-8")
        summary_xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
        workbook_content = "".join(
            archive.read(name).decode("utf-8")
            for name in archive.namelist()
            if name.startswith("xl/") and name.endswith(".xml")
        )
    assert "Сводка" in workbook_xml
    assert "Отчёт СКС" in summary_xml
    assert "Иванова Елена" in summary_xml
    assert "Договор и счет" in workbook_content
    assert "Аномалии" in workbook_xml


def test_build_sks_chart_creates_readable_png() -> None:
    metric = ProjectMetric(
        project_id="project-1",
        contract_number="СКС-1",
        service="sout",
        specialist="Иванова Елена",
        start_at=datetime(2026, 7, 1, 9, 0),
        target_at=datetime(2026, 7, 3, 12, 0),
        duration_days=2.125,
        sale_price=Decimal("125000.50"),
    )
    chart = build_sks_chart(
        SksReportData(
            date_from=date(2026, 7, 1),
            date_to=date(2026, 7, 31),
            agreements=(metric,),
            sends=(metric,),
            tasks=(
                TaskMetric(
                    category="Договор и счет",
                    created_at=datetime(2026, 7, 1, 17, 0),
                    closed_at=datetime(2026, 7, 2, 9, 30),
                    working_seconds=3600,
                ),
            ),
            sending_tasks={"Балакирева Диана": 3, "Кулешева Владислава": 4},
        )
    )

    image = Image.open(io.BytesIO(chart))
    assert image.format == "PNG"
    assert image.size == (1200, 2000)


def test_projects_over_200_days_are_excluded_only_from_time_statistics() -> None:
    normal = ProjectMetric(
        project_id="normal",
        contract_number="СКС-1",
        service="sout",
        specialist="Иванова Елена",
        start_at=datetime(2026, 7, 1, 9, 0),
        target_at=datetime(2026, 7, 3, 9, 0),
        duration_days=2.0,
        sale_price=Decimal("100000"),
    )
    anomaly = ProjectMetric(
        project_id="anomaly",
        contract_number="СКС-2",
        service="pk",
        specialist="Иванова Елена",
        start_at=datetime(2025, 7, 1, 9, 0),
        target_at=datetime(2026, 7, 20, 9, 0),
        duration_days=ANOMALY_DURATION_DAYS + 0.01,
        sale_price=Decimal("250000"),
    )
    data = SksReportData(
        date_from=date(2026, 7, 1),
        date_to=date(2026, 7, 31),
        agreements=(normal, anomaly),
        sends=(normal, anomaly),
        tasks=(),
        sending_tasks={},
    )

    values = _summary_values(data, ("Иванова Елена",))
    content = build_sks_workbook(data)

    assert values["agreement_days"] == 2.0
    assert values["send_days"] == 2.0
    assert values["approved_count"] == 2
    assert values["approved_amount"] == Decimal("350000")
    with zipfile.ZipFile(io.BytesIO(content)) as archive:
        workbook_content = "".join(
            archive.read(name).decode("utf-8")
            for name in archive.namelist()
            if name.startswith("xl/") and name.endswith(".xml")
        )
    assert "Аномалия &gt; 200 дней" in workbook_content
    assert "Исключён только из статистики времени" in workbook_content
