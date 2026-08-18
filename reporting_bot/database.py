from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Mapping

from sqlalchemy import create_engine, text
from sqlalchemy.pool import NullPool

from reporting_bot.config import Settings
from reporting_bot.reports import Report


@dataclass(frozen=True)
class QueryResult:
    columns: tuple[str, ...]
    rows: tuple[tuple[Any, ...], ...]
    truncated: bool


class ReportExecutor:
    def __init__(self, settings: Settings) -> None:
        if not settings.database_url.startswith(("postgresql://", "postgresql+psycopg://")):
            raise ValueError("DATABASE_URL must be a PostgreSQL URL")
        self._settings = settings
        self._database_url = settings.database_url.replace(
            "postgresql://", "postgresql+psycopg://", 1
        )

    def run(self, report: Report, parameters: Mapping[str, str]) -> QueryResult:
        connect_args = {"connect_timeout": self._settings.connect_timeout_seconds}
        engine = create_engine(
            self._database_url,
            poolclass=NullPool,
            connect_args=connect_args,
        )
        try:
            with engine.connect() as connection:
                with connection.begin():
                    connection.execute(text("SET TRANSACTION READ ONLY"))
                    connection.execute(
                        text("SELECT set_config('statement_timeout', :timeout, true)"),
                        {"timeout": f"{self._settings.statement_timeout_ms}ms"},
                    )
                    cursor = connection.execute(text(report.sql), dict(parameters))
                    rows = cursor.fetchmany(report.max_rows + 1)
                    columns = tuple(cursor.keys())
                    return QueryResult(
                        columns=columns,
                        rows=tuple(tuple(row) for row in rows[: report.max_rows]),
                        truncated=len(rows) > report.max_rows,
                    )
        finally:
            engine.dispose()
