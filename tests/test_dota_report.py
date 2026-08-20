import io
import zipfile
from datetime import date, datetime
from decimal import Decimal

from reporting_bot.dota_report import (
    DOTA_EVENTS_SQL,
    DotaEvent,
    DotaReportData,
    _parse_monthly_plans,
    _previous_month_date,
    build_dota_workbook,
)


def event(metric: str, category: str, project_id: str, event_at: datetime) -> DotaEvent:
    return DotaEvent(
        metric=metric,
        category=category,
        project_id=project_id,
        contract_number="ДОТ-1",
        company_name="ООО Пример",
        manager="Иванов Иван",
        manager_department="ОП",
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
    assert "ОПР, РМ" in launch_xml


def test_previous_period_uses_same_dates_of_previous_month() -> None:
    assert _previous_month_date(date(2026, 8, 1)) == date(2026, 7, 1)
    assert _previous_month_date(date(2026, 8, 18)) == date(2026, 7, 18)
    assert _previous_month_date(date(2026, 3, 31)) == date(2026, 2, 28)


def test_dota_sql_uses_target_stage_without_previous_stage_restriction() -> None:
    assert "h.new_step = 'Принять проект'" in DOTA_EVENTS_SQL
    assert "h.new_step = 'Подготовить документы'" in DOTA_EVENTS_SQL
    assert "h.new_step = 'Передать документы Заказчику'" in DOTA_EVENTS_SQL
    assert "h.previous_step =" not in DOTA_EVENTS_SQL
    assert "real_workplace_count" in DOTA_EVENTS_SQL
    assert "manager_department" in DOTA_EVENTS_SQL
    assert "manager_sks" not in DOTA_EVENTS_SQL


def test_monthly_plans_can_be_supplied_as_json() -> None:
    plans = _parse_monthly_plans(
        '{"2026-09":{"Запуск":{"ОПР":123},"Выпуск":{"ОПР":456}}}'
    )

    assert plans["2026-09"]["Запуск"]["ОПР"] == Decimal("123")
    assert plans["2026-09"]["Выпуск"]["ОПР"] == Decimal("456")


def test_detail_sheets_have_department_and_only_relevant_stage_column() -> None:
    content = build_dota_workbook(
        DotaReportData(
            date_from=date(2026, 8, 1),
            date_to=date(2026, 8, 3),
            previous_date_from=date(2026, 7, 1),
            previous_date_to=date(2026, 7, 3),
            events=(event("Запуск", "ОПР", "launch-1", datetime(2026, 8, 3, 10)),),
            previous_events=(),
        )
    )

    with zipfile.ZipFile(io.BytesIO(content)) as archive:
        details_xml = archive.read("xl/worksheets/sheet4.xml").decode("utf-8")

    assert "Отдел менеджера" in details_xml
    assert "Этап запуска" in details_xml
    assert "Специалист СКС" not in details_xml
    assert "Предыдущий этап" not in details_xml
