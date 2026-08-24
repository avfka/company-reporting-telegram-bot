from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from sqlalchemy import create_engine, text
from sqlalchemy.pool import NullPool

from reporting_bot.config import Settings


CRM_TABLES = (
    "clients",
    "client_companies",
    "companies",
    "companies_clone",
    "projects",
    "requests",
    "requests_clone",
    "tasks",
    "users",
)


@dataclass(frozen=True)
class CrmBridgeRepository:
    settings: Settings

    def schema(self) -> dict[str, Any]:
        database_url = self.settings.database_url.replace(
            "postgresql://", "postgresql+psycopg://", 1
        )
        engine = create_engine(
            database_url,
            poolclass=NullPool,
            connect_args={"connect_timeout": self.settings.connect_timeout_seconds},
        )
        try:
            with engine.connect() as connection:
                with connection.begin():
                    connection.execute(text("SET TRANSACTION READ ONLY"))
                    connection.execute(
                        text("SELECT set_config('statement_timeout', :timeout, true)"),
                        {"timeout": f"{self.settings.statement_timeout_ms}ms"},
                    )
                    columns = connection.execute(
                        text(
                            """
                            SELECT table_name, ordinal_position, column_name,
                                   data_type, udt_name, is_nullable
                            FROM information_schema.columns
                            WHERE table_schema = 'public'
                              AND table_name = ANY(:tables)
                            ORDER BY table_name, ordinal_position
                            """
                        ),
                        {"tables": list(CRM_TABLES)},
                    ).mappings()
                    foreign_keys = connection.execute(
                        text(
                            """
                            SELECT tc.table_name, kcu.column_name,
                                   ccu.table_name AS referenced_table,
                                   ccu.column_name AS referenced_column
                            FROM information_schema.table_constraints tc
                            JOIN information_schema.key_column_usage kcu
                              ON tc.constraint_name = kcu.constraint_name
                             AND tc.constraint_schema = kcu.constraint_schema
                            JOIN information_schema.constraint_column_usage ccu
                              ON ccu.constraint_name = tc.constraint_name
                             AND ccu.constraint_schema = tc.constraint_schema
                            WHERE tc.constraint_type = 'FOREIGN KEY'
                              AND tc.table_schema = 'public'
                              AND tc.table_name = ANY(:tables)
                            ORDER BY tc.table_name, kcu.column_name
                            """
                        ),
                        {"tables": list(CRM_TABLES)},
                    ).mappings()
                    return {
                        "columns": [dict(row) for row in columns],
                        "foreign_keys": [dict(row) for row in foreign_keys],
                    }
        finally:
            engine.dispose()
