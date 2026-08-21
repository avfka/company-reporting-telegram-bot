import io
import zipfile
from dataclasses import replace
from datetime import date, datetime
from decimal import Decimal

from PIL import Image

from reporting_bot.dota_report import (
    DOTA_EVENTS_SQL,
    DotaEvent,
    DotaReportData,
    _dota_daily_series,
    _dota_share_rows,
    _parse_monthly_plans,
    _previous_month_date,
    build_dota_chart,
    build_dota_detail_charts,
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


def test_build_dota_chart_creates_readable_png() -> None:
    content = build_dota_chart(
        DotaReportData(
            date_from=date(2026, 8, 1),
            date_to=date(2026, 8, 18),
            previous_date_from=date(2026, 7, 1),
            previous_date_to=date(2026, 7, 18),
            events=(
                event("Запуск", "ОПР", "launch-1", datetime(2026, 8, 3, 10)),
                event("Выпуск", "ОПР", "release-1", datetime(2026, 8, 10, 10)),
            ),
            previous_events=(
                event("Запуск", "ОПР", "previous-1", datetime(2026, 7, 3, 10)),
            ),
        )
    )

    image = Image.open(io.BytesIO(content))
    assert image.format == "PNG"
    assert image.size == (1200, 1800)


def test_dota_infographic_uses_daily_values_and_groups_small_pie_slices() -> None:
    rows = tuple(
        replace(
            event("Запуск", "Обучение", str(index), datetime(2026, 8, index, 10)),
            company_name=f"Компания {index}",
            amount=Decimal(index * 100),
        )
        for index in range(1, 5)
    )
    data = DotaReportData(
        date_from=date(2026, 8, 1),
        date_to=date(2026, 8, 4),
        previous_date_from=date(2026, 7, 1),
        previous_date_to=date(2026, 7, 4),
        events=rows,
        previous_events=(),
    )

    shares = _dota_share_rows(rows, lambda row: row.company_name, max_items=2)
    dates, launch, release = _dota_daily_series(data, lambda row: row.amount)

    assert [label for label, _ in shares] == ["Компания 4", "Компания 3", "Остальные"]
    assert shares[-1][1] == Decimal("300")
    assert dates == [date(2026, 8, day) for day in range(1, 5)]
    assert launch == [100.0, 200.0, 300.0, 400.0]
    assert release == [0.0, 0.0, 0.0, 0.0]


def test_dota_builds_four_separate_detail_charts() -> None:
    data = DotaReportData(
        date_from=date(2026, 8, 1),
        date_to=date(2026, 8, 4),
        previous_date_from=date(2026, 7, 1),
        previous_date_to=date(2026, 7, 4),
        events=(
            event("Запуск", "ОПР", "opr-launch", datetime(2026, 8, 1, 10)),
            event("Выпуск", "ОПР", "opr-release", datetime(2026, 8, 2, 10)),
            event("Запуск", "Обучение", "training-launch", datetime(2026, 8, 3, 10)),
            event("Выпуск", "Обучение", "training-release", datetime(2026, 8, 4, 10)),
        ),
        previous_events=(),
    )

    charts = build_dota_detail_charts(data, "Отчет_ДОТ")

    assert len(charts) == 4
    assert [chart.filename for chart in charts] == [
        "Отчет_ДОТ_ОПР_доли.png",
        "Отчет_ДОТ_Обучение_доли.png",
        "Отчет_ДОТ_ОПР_по_дням.png",
        "Отчет_ДОТ_Обучение_по_дням.png",
    ]
    sizes = [Image.open(io.BytesIO(chart.content)).size for chart in charts]
    assert sizes == [(1200, 900), (1200, 900), (1400, 1310), (1400, 980)]
