import io
import zipfile
from datetime import date, datetime
from decimal import Decimal

from reporting_bot.sks_report import (
    ProjectMetric,
    SksReportData,
    TASK_DURATION_SQL,
    TaskMetric,
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


def test_task_query_includes_combined_contract_and_invoice_titles() -> None:
    contract_titles = "('договор', 'договор и счет', 'счет и договор')"
    all_titles = "('договор', 'договор и счет', 'счет и договор', 'счет')"

    assert f"normalized_title IN {contract_titles}" in TASK_DURATION_SQL
    assert f"normalized_title IN {all_titles}" in TASK_DURATION_SQL
    assert "regexp_replace(trim(title), '\\s+', ' ', 'g')" in TASK_DURATION_SQL
    assert "'ё'," in TASK_DURATION_SQL
    assert "LIKE '%договор%'" not in TASK_DURATION_SQL


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
        category="Договор",
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
        assert len([name for name in archive.namelist() if name.startswith("xl/worksheets/sheet")]) == 6
        workbook_xml = archive.read("xl/workbook.xml").decode("utf-8")
        summary_xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
    assert "Сводка" in workbook_xml
    assert "Отчёт СКС" in summary_xml
    assert "Иванова Елена" in summary_xml
