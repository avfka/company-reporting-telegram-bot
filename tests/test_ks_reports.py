import io
import zipfile
from datetime import date, datetime
from decimal import Decimal

from PIL import Image

from reporting_bot.ks_reports import (
    KsEvent,
    KsFilters,
    KsReportData,
    build_ks_chart,
    build_ks_workbook,
)


def sample_data(report_kind: str) -> KsReportData:
    current = (
        KsEvent("Лид", "lead-1", datetime(2026, 7, 1, 10), "Иванова Елена", "КС", "sout", status="Успешно"),
        KsEvent("Лид", "lead-2", datetime(2026, 7, 2, 10), "Иванова Елена", "КС", "sout", status="Думает"),
        KsEvent("КП", "lead-1", datetime(2026, 7, 2, 11), "Иванова Елена", "КС", "sout", reference="КП-1"),
        KsEvent("Проект", "project-1", date(2026, 7, 3), "Иванова Елена", "КС", "sout", Decimal("150000"), "Д-1", current_step="Распечатка", paid_amount=Decimal("50000"), project_id="project-1", is_agreed=True),
        KsEvent("Платёж", "payment-1", date(2026, 7, 4), "Иванова Елена", "КС", "sout", Decimal("50000"), "Д-1", project_id="project-1"),
        KsEvent("Звонок", "call-1", datetime(2026, 7, 1, 9), "Иванова Елена", "КС", "", reference="outbound"),
    )
    previous = (
        KsEvent("Лид", "lead-0", datetime(2026, 6, 1, 10), "Иванова Елена", "КС", "sout", status="Успешно"),
        KsEvent("Проект", "project-0", date(2026, 6, 3), "Иванова Елена", "КС", "sout", Decimal("100000"), "Д-0", paid_amount=Decimal("100000"), project_id="project-0"),
        KsEvent("Платёж", "payment-0", date(2026, 6, 4), "Иванова Елена", "КС", "sout", Decimal("100000"), "Д-0", project_id="project-0"),
    )
    return KsReportData(
        report_kind=report_kind,
        date_from=date(2026, 7, 1),
        date_to=date(2026, 7, 31),
        comparison_from=date(2026, 6, 1),
        comparison_to=date(2026, 6, 30),
        comparison_mode="previous",
        filters=KsFilters(),
        filter_labels={"department": "Все отделы", "manager": "Все менеджеры", "product": "Все продукты"},
        events=current,
        comparison_events=previous,
    )


def test_all_ks_reports_create_valid_workbooks_and_png_charts() -> None:
    for report_kind in ("all", "plan", "managers", "funnel", "projects"):
        data = sample_data(report_kind)
        workbook = build_ks_workbook(data)
        with zipfile.ZipFile(io.BytesIO(workbook)) as archive:
            assert archive.testzip() is None
            content = "".join(
                archive.read(name).decode("utf-8")
                for name in archive.namelist()
                if name.startswith("xl/") and name.endswith(".xml")
            )
        assert "Иванова Елена" in content

        chart = build_ks_chart(data)
        image = Image.open(io.BytesIO(chart))
        assert image.format == "PNG"
        assert image.size == (1200, 1200)


def test_combined_report_contains_every_analytics_section() -> None:
    workbook = build_ks_workbook(sample_data("all"))
    with zipfile.ZipFile(io.BytesIO(workbook)) as archive:
        workbook_xml = archive.read("xl/workbook.xml").decode("utf-8")
    for sheet_name in (
        "Общая сводка",
        "По отделам",
        "По продуктам",
        "Менеджеры",
        "Воронка",
        "По менеджерам",
        "Проекты и оплаты",
        "Детализация",
    ):
        assert sheet_name in workbook_xml


def test_product_filter_uses_compact_hash_token() -> None:
    from reporting_bot.ks_reports import PROJECTS_SQL

    assert "left(md5(lower(coalesce(p.service, ''))), 6)" in PROJECTS_SQL


def test_manager_filter_accepts_multiple_compact_tokens() -> None:
    from reporting_bot.ks_reports import PROJECTS_SQL

    assert "ANY(string_to_array(:manager_token, ','))" in PROJECTS_SQL


def test_ks_sql_uses_request_cohort_experts_and_positive_payments() -> None:
    from reporting_bot.ks_reports import OFFERS_SQL, PAYMENTS_SQL, PROJECTS_SQL, REQUESTS_SQL

    assert "NOT IN ('Не лид', 'Дубль')" not in REQUESTS_SQL
    assert "o.id::text AS event_id" in OFFERS_SQL
    assert "o.is_sent" not in OFFERS_SQL
    assert "r.created_at >= CAST(:date_from AS DATE)" in OFFERS_SQL
    assert "p.expert_id IS NOT NULL" in PROJECTS_SQL
    assert "u.id = p.expert_id" in PROJECTS_SQL
    assert "psh.new_step = 'Распечатка'" in PROJECTS_SQL
    assert "pay.amount > 0" in PAYMENTS_SQL
    assert "u.id = p.expert_id" in PAYMENTS_SQL
    assert "r.created_at >= CAST(:date_from AS DATE)" in PAYMENTS_SQL


def test_conversion_is_success_over_all_requests_in_cohort() -> None:
    from reporting_bot.ks_reports import _metrics

    rows = (
        KsEvent("Лид", "1", date(2026, 7, 1), "А", "КС", "sout", status="Успешно"),
        KsEvent("Лид", "2", date(2026, 7, 1), "А", "КС", "sout", status="Не лид"),
        KsEvent("Лид", "3", date(2026, 7, 1), "А", "КС", "sout", status="Думает"),
        KsEvent("Проект", "p1", date(2026, 7, 2), "А", "КС", "sout"),
        KsEvent("Проект", "p2", date(2026, 7, 2), "А", "КС", "sout"),
    )

    metrics = _metrics(rows)

    assert metrics["leads"] == 3
    assert metrics["projects"] == 2
    assert metrics["conversion"] == 1 / 3


def test_project_summary_sums_partial_payments_excludes_returns_and_tracks_agreement() -> None:
    from reporting_bot.ks_reports import _project_summary

    rows = (
        KsEvent("Проект", "p1", date(2026, 7, 1), "Эксперт", "КС", "sout", Decimal("100"), project_id="p1", is_agreed=True),
        KsEvent("Проект", "p2", date(2026, 7, 1), "Эксперт", "КС", "sout", Decimal("80"), project_id="p2", current_step="Согласовать отчет у клиента"),
        KsEvent("Платёж", "pay1", date(2026, 7, 2), "Эксперт", "КС", "sout", Decimal("30"), project_id="p1"),
        KsEvent("Платёж", "pay2", date(2026, 7, 3), "Эксперт", "КС", "sout", Decimal("20"), project_id="p1"),
        KsEvent("Платёж", "pay3", date(2026, 7, 3), "Эксперт", "КС", "sout", Decimal("10"), project_id="p2"),
        KsEvent("Платёж", "refund", date(2026, 7, 4), "Эксперт", "КС", "sout", Decimal("-15"), project_id="p1"),
    )

    summary = _project_summary(rows)

    assert summary["projects"] == 2
    assert summary["occurrence"] == Decimal("180")
    assert summary["paid"] == Decimal("60")
    assert summary["awaiting"] == Decimal("120")
    assert summary["agreed"] == 1
    assert summary["agreed_paid"] == Decimal("50")
    assert summary["agreed_awaiting"] == Decimal("50")
    assert summary["approval"] == 1
