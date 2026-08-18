from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping


REPORT_ID_RE = re.compile(r"^[a-z][a-z0-9_]{1,63}$")
FORBIDDEN_SQL_RE = re.compile(
    r"\b(insert|update|delete|merge|alter|drop|truncate|create|grant|revoke|copy|call|do|vacuum|refresh)\b",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class Report:
    report_id: str
    title: str
    description: str
    sql: str
    parameters: tuple[str, ...]
    max_rows: int


class ReportCatalog:
    def __init__(self, reports: Mapping[str, Report]) -> None:
        self._reports = dict(reports)

    @classmethod
    def from_file(cls, path: str, default_max_rows: int) -> "ReportCatalog":
        source = Path(path)
        if not source.is_absolute():
            source = Path(__file__).resolve().parent.parent / source
        payload = json.loads(source.read_text(encoding="utf-8"))
        if not isinstance(payload, list):
            raise ValueError("Reports file must contain a JSON array")

        reports: dict[str, Report] = {}
        for item in payload:
            report = cls._parse(item, default_max_rows)
            if report.report_id in reports:
                raise ValueError(f"Duplicate report id: {report.report_id}")
            reports[report.report_id] = report
        if not reports:
            raise ValueError("At least one report is required")
        return cls(reports)

    @staticmethod
    def _parse(item: Mapping[str, Any], default_max_rows: int) -> Report:
        report_id = str(item.get("id", "")).strip()
        title = str(item.get("title", "")).strip()
        description = str(item.get("description", "")).strip()
        sql = str(item.get("sql", "")).strip()
        parameters = tuple(str(value).strip() for value in item.get("parameters", []))
        max_rows = int(item.get("max_rows", default_max_rows))

        if not REPORT_ID_RE.fullmatch(report_id):
            raise ValueError(f"Invalid report id: {report_id!r}")
        if not title or not description:
            raise ValueError(f"Report {report_id} needs a title and description")
        if not sql.lower().startswith(("select", "with")):
            raise ValueError(f"Report {report_id} must be a SELECT query")
        if ";" in sql.rstrip(";") or FORBIDDEN_SQL_RE.search(sql):
            raise ValueError(f"Report {report_id} contains forbidden SQL")
        if len(parameters) != len(set(parameters)) or any(not value.isidentifier() for value in parameters):
            raise ValueError(f"Report {report_id} has invalid parameters")
        if max_rows <= 0 or max_rows > 500:
            raise ValueError(f"Report {report_id} max_rows must be between 1 and 500")

        return Report(report_id, title, description, sql.rstrip(";"), parameters, max_rows)

    def get(self, report_id: str) -> Report | None:
        return self._reports.get(report_id)

    def all(self) -> tuple[Report, ...]:
        return tuple(self._reports.values())
