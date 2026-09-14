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
        assert image.width == 1200
        assert image.height == {"all": 1200, "plan": 720, "managers": 720, "funnel": 914, "projects": 810}[report_kind]


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

    assert "left(md5(CASE" in PROJECTS_SQL
    assert "= 'соут' THEN 'sout'" in PROJECTS_SQL


def test_manager_filter_accepts_multiple_compact_tokens() -> None:
    from reporting_bot.ks_reports import PROJECTS_SQL

    assert "ANY(string_to_array(:manager_token, ','))" in PROJECTS_SQL


def test_ks_sql_uses_event_dates_managers_and_positive_payments() -> None:
    from reporting_bot.ks_reports import OFFERS_SQL, PAYMENTS_SQL, PROJECTS_SQL, REQUESTS_SQL

    assert "NOT IN ('Не лид', 'Дубль')" not in REQUESTS_SQL
    assert "o.id::text AS event_id" in OFFERS_SQL
    assert "o.is_sent" not in OFFERS_SQL
    assert "o.created_at >= CAST(:date_from AS DATE)" in OFFERS_SQL
    assert "r.created_at >= CAST(:date_from AS DATE)" not in OFFERS_SQL
    assert "k.user_id=p.manager_id" in PROJECTS_SQL
    assert "min(l.created_at)" in PROJECTS_SQL
    assert "('measurer_id','expert_id')" in PROJECTS_SQL
    assert "psh.new_step = 'Распечатка'" in PROJECTS_SQL
    assert "pay.amount > 0" in PAYMENTS_SQL
    assert "k.user_id=p.manager_id" in PAYMENTS_SQL
    assert 'pay."chargeAt" >= CAST(:date_from AS DATE)' in PAYMENTS_SQL
    assert "r.created_at >= CAST(:date_from AS DATE)" not in PAYMENTS_SQL


def test_roster_is_explicit_and_has_only_five_approved_gto_users() -> None:
    from collections import Counter
    from reporting_bot.ks_scope import ROSTER

    assert len(ROSTER) == len({r[0] for r in ROSTER}) == 22
    assert Counter(r[2] for r in ROSTER) == {"op": 9, "siber": 4, "kam": 3, "skam": 1, "aop": 5}
    assert sum(r[3] == 2 for r in ROSTER) == 5
    assert len({r[0].replace('-', '')[:6] for r in ROSTER}) == 22
    assert next(r[2] for r in ROSTER if r[1] == "Шергина Надежда") == "kam"


def test_manual_manager_reply_supports_new_department_tokens() -> None:
    from reporting_bot.telegram import _manual_manager_context
    for token in ('op', 'siber', 'kam', 'skam', 'aop'):
        result = _manual_manager_context(f'KSM|ksa|20260901|20260914|{token}|-|-'.replace('|-|-', '|-|p'))
        assert result is not None
        assert result[3] == token


def test_stale_department_filter_fails_explicitly() -> None:
    import pytest
    from reporting_bot.ks_reports import _validate_filters, KsFilterOptions, FilterOption
    options = KsFilterOptions((FilterOption('op', 'ОП'),), (), ())
    with pytest.raises(ValueError, match="заново"):
        _validate_filters(options, KsFilters(department_token='cde48f'))


def test_project_balance_does_not_use_only_current_period_cash() -> None:
    from reporting_bot.ks_reports import _metrics, _project_summary

    rows = (
        KsEvent("Проект", "p1", date(2026, 9, 1), "А", "ОП", "sout", Decimal(100), paid_amount=Decimal(80), project_id="p1"),
        KsEvent("Платёж", "pay1", date(2026, 9, 2), "А", "ОП", "sout", Decimal(20), project_id="p1"),
        KsEvent("Платёж", "old-project-pay", date(2026, 9, 2), "А", "ОП", "sout", Decimal(400), project_id="old"),
    )
    assert _metrics(rows)["cash"] == 420
    assert _project_summary(rows)["paid"] == 80
    assert _project_summary(rows)["awaiting"] == 20


def test_normalized_service_and_auditable_workbook() -> None:
    from reporting_bot.ks_reports import _service_label
    assert _service_label(" соут ") == _service_label("SOUT") == "СОУТ"
    with zipfile.ZipFile(io.BytesIO(build_ks_workbook(sample_data("all")))) as archive:
        content = ''.join(archive.read(n).decode() for n in archive.namelist() if n.endswith('.xml'))
    for value in ("Состав КС", "Правила расчёта", "ID проекта", "Все оплаты проекта, руб.", "Айсина Юлианна"):
        assert value in content


def test_query_row_limit_raises_instead_of_returning_partial_data() -> None:
    import pytest
    from reporting_bot.ks_reports import _load_events

    class Result:
        def mappings(self):
            return [{}] * 10001

    class Connection:
        def execute(self, *args):
            return Result()

    with pytest.raises(ValueError, match="не обрезается"):
        _load_events(Connection(), date(2026, 9, 1), date(2026, 9, 14), KsFilters())


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
