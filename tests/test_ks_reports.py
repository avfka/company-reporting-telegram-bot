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
        KsEvent("Проект", "project-1", date(2026, 7, 3), "Иванова Елена", "КС", "sout", Decimal("150000"), "Д-1", current_step="Согласовать отчет у клиента", paid_amount=Decimal("50000")),
        KsEvent("Платёж", "payment-1", date(2026, 7, 4), "Иванова Елена", "КС", "sout", Decimal("50000"), "Д-1"),
        KsEvent("Звонок", "call-1", datetime(2026, 7, 1, 9), "Иванова Елена", "КС", "", reference="outbound"),
    )
    previous = (
        KsEvent("Лид", "lead-0", datetime(2026, 6, 1, 10), "Иванова Елена", "КС", "sout", status="Успешно"),
        KsEvent("Проект", "project-0", date(2026, 6, 3), "Иванова Елена", "КС", "sout", Decimal("100000"), "Д-0", paid_amount=Decimal("100000")),
        KsEvent("Платёж", "payment-0", date(2026, 6, 4), "Иванова Елена", "КС", "sout", Decimal("100000"), "Д-0"),
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
    for report_kind in ("plan", "managers", "funnel", "projects"):
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


def test_product_filter_uses_compact_hash_token() -> None:
    from reporting_bot.ks_reports import PROJECTS_SQL

    assert "left(md5(lower(coalesce(p.service, ''))), 6)" in PROJECTS_SQL
