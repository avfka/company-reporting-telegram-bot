import json

import pytest

from reporting_bot.reports import ReportCatalog


def write_reports(tmp_path, reports):
    path = tmp_path / "reports.json"
    path.write_text(json.dumps(reports), encoding="utf-8")
    return path


def test_catalog_loads_parameterized_select(tmp_path) -> None:
    path = write_reports(
        tmp_path,
        [
            {
                "id": "sales_by_day",
                "title": "Sales",
                "description": "Daily sales",
                "parameters": ["day"],
                "sql": "SELECT total FROM sales WHERE day = :day",
            }
        ],
    )
    report = ReportCatalog.from_file(str(path), 50).get("sales_by_day")
    assert report is not None
    assert report.parameters == ("day",)


@pytest.mark.parametrize(
    "sql",
    [
        "DELETE FROM users",
        "WITH changed AS (UPDATE users SET admin = true RETURNING *) SELECT * FROM changed",
        "SELECT * FROM users; DROP TABLE users",
    ],
)
def test_catalog_rejects_mutating_sql(tmp_path, sql) -> None:
    path = write_reports(
        tmp_path,
        [{"id": "unsafe_query", "title": "Unsafe", "description": "Unsafe", "sql": sql}],
    )
    with pytest.raises(ValueError):
        ReportCatalog.from_file(str(path), 50)
