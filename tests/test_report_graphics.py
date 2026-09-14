"""Graphic regression cases; fixtures contain no production data."""
import io
from dataclasses import replace
from datetime import date, datetime, timedelta
from decimal import Decimal

import pytest
from PIL import Image, ImageDraw

from reporting_bot.chart_layout import fitted_text, daily_panel
from reporting_bot.ks_reports import _font, build_ks_chart
from reporting_bot.sks_report import SCS_SPECIALISTS, ProjectMetric, SksReportData, build_sks_chart
from reporting_bot.dota_report import DotaReportData, build_dota_chart, build_dota_detail_charts
from test_ks_reports import sample_data
from test_dota_report import event


def demo_dota(days=31):
    start = date(2026, 8, 1)
    rows = []
    for i in range(days):
        for category in ("ОПР", "Обучение"):
            for metric in ("Запуск", "Выпуск"):
                if i % 7 < 5:
                    rows.append(replace(event(metric, category, f"{i}-{category}-{metric}",
                                              datetime.combine(start + timedelta(days=i), datetime.min.time())),
                                        amount=Decimal((i * 139 + (250 if metric == "Запуск" else 90)) % 850 * 900),
                                        company_name=("Компания с очень длинным названием для проверки подписей" if i % 5 == 0 else f"Компания {i % 5}"),
                                        manager_department=("ОП", "КАМ", "ОАП", "ГТО", "Другой отдел")[i % 5]))
    return DotaReportData(start, start + timedelta(days=days-1), date(2026, 7, 1), date(2026, 7, 31), tuple(rows), ())


def demo_sks():
    rows = tuple(ProjectMetric(str(i), f"СКС-{i}", "sout" if i % 2 else "opk", name,
                               datetime(2026, 8, 1), datetime(2026, 8, 3),
                               float(i + 1), Decimal(123456 * (i+1)))
                 for i, name in enumerate(SCS_SPECIALISTS))
    return SksReportData(date(2026, 8, 1), date(2026, 8, 31), rows, rows, (), {})


def test_fit_uses_glyph_width_for_long_cyrillic_labels():
    draw = ImageDraw.Draw(Image.new("RGB", (300, 100)))
    label = fitted_text(draw, (0, 0), "Очень длинное название организации " * 5,
                        _font, width=190, size=22, minimum=16)
    assert label.endswith("...")
    assert draw.textlength(label, font=_font(16)) <= 190


@pytest.mark.parametrize("days", [1, 7, 31, 46, 366])
def test_daily_panel_preserves_input_and_fits_long_periods(days):
    dates = [date(2026, 1, 1) + timedelta(days=i) for i in range(days)]
    launches = [float(i % 9) for i in range(days)]
    releases = [float(i % 5) for i in range(days)]
    before = (list(launches), list(releases))
    image = Image.new("RGB", (1200, 320), "white")
    daily_panel(ImageDraw.Draw(image), (5, 5, 1195, 315), "Дни", dates, launches, releases,
                font=_font, formatter=lambda n: str(n))
    assert (launches, releases) == before


@pytest.mark.parametrize("kind", ["all", "plan", "managers", "funnel", "projects"])
def test_empty_charts_and_missing_comparison(kind):
    data = replace(sample_data(kind), events=(), comparison_events=(), comparison_from=None,
                   comparison_to=None, comparison_mode="none")
    assert Image.open(io.BytesIO(build_ks_chart(data))).width == 1200


def test_eight_manager_rows_get_space_and_no_data_mutation():
    data = sample_data("managers")
    rows = tuple(replace(data.events[3], manager=f"Менеджер с длинным именем {i}", event_id=f"p{i}") for i in range(8))
    data = replace(data, events=rows)
    assert Image.open(io.BytesIO(build_ks_chart(data))).size == (1200, 1330)
    assert data.events == rows


def test_dota_delivery_stays_one_overview_and_four_details():
    data = demo_dota()
    assert Image.open(io.BytesIO(build_dota_chart(data))).size == (1200, 1800)
    details = build_dota_detail_charts(data, "demo")
    assert len(details) == 4
    assert len({chart.filename for chart in details}) == 4
    for chart in details:
        Image.open(io.BytesIO(chart.content)).verify()


def test_empty_sks_renders():
    data = replace(demo_sks(), agreements=(), sends=())
    assert Image.open(io.BytesIO(build_sks_chart(data))).size == (1200, 2000)


if __name__ == "__main__":
    import sys
    from pathlib import Path
    output = Path(sys.argv[1])
    output.mkdir(parents=True, exist_ok=True)
    # Generated test-only images for visual review, never CRM figures.
    for kind in ("all", "plan", "managers", "funnel", "projects"):
        (output / f"demo-ks-{kind}.png").write_bytes(build_ks_chart(sample_data(kind)))
    (output / "demo-sks.png").write_bytes(build_sks_chart(demo_sks()))
    data = demo_dota()
    (output / "demo-dota.png").write_bytes(build_dota_chart(data))
    for i, chart in enumerate(build_dota_detail_charts(data, "demo")):
        (output / f"demo-dota-{i}.png").write_bytes(chart.content)
    for i, chart in enumerate(build_dota_detail_charts(demo_dota(366), "demo")):
        if i > 1:
            (output / f"demo-dota-year-{i}.png").write_bytes(chart.content)
    print(output)
