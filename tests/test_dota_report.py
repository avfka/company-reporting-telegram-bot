import io
import zipfile
from datetime import date, datetime
from decimal import Decimal

from reporting_bot.dota_report import DotaEvent, DotaReportData, build_dota_workbook


def event(metric: str, category: str, project_id: str, event_at: datetime) -> DotaEvent:
    return DotaEvent(
        metric=metric,
        category=category,
        project_id=project_id,
        contract_number="ДОТ-1",
        company_name="ООО Пример",
        manager="Иванов Иван",
        manager_sks="Иванова Елена",
        event_owner="Петров Пётр",
        event_at=event_at,
        previous_step="Получить документы",
        new_step="Принять проект",
        amount=Decimal("125000.50"),
        workplace_count=18,
    )


def test_build_dota_workbook_creates_summary_daily_and_detail_sheets() -> None:
    content = build_dota_workbook(
        DotaReportData(
            date_from=date(2026, 8, 1),
            date_to=date(2026, 8, 3),
            previous_date_from=date(2026, 7, 29),
            previous_date_to=date(2026, 7, 31),
            events=(
                event("Запуск", "ОПР", "launch-1", datetime(2026, 8, 3, 10, 0)),
                event("Выпуск", "Обучение", "release-1", datetime(2026, 8, 3, 12, 0)),
            ),
            previous_events=(
                event("Запуск", "ОПР", "previous-1", datetime(2026, 7, 30, 10, 0)),
            ),
        )
    )

    with zipfile.ZipFile(io.BytesIO(content)) as archive:
        assert archive.testzip() is None
        assert len([name for name in archive.namelist() if name.startswith("xl/worksheets/sheet")]) == 5
        workbook_xml = archive.read("xl/workbook.xml").decode("utf-8")
        summary_xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
        launch_xml = archive.read("xl/worksheets/sheet2.xml").decode("utf-8")
    assert "Сводная" in workbook_xml
    assert "Детали запуска" in workbook_xml
    assert "Отчёт ДОТ" in summary_xml
    assert "ОПР" in summary_xml
    assert "03.08.2026" in launch_xml
