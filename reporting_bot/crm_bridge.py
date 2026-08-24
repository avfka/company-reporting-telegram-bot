from __future__ import annotations

from dataclasses import dataclass
import re
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

    def find_contacts(self, phone: str) -> dict[str, Any]:
        phone_digits = re.sub(r"\D", "", phone)
        if len(phone_digits) < 7:
            raise ValueError("Phone must contain at least 7 digits")

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
                    rows = connection.execute(
                        text(
                            """
                            SELECT
                              c.id::text AS client_id,
                              c.name AS client_name,
                              c.phone,
                              c.second_phone,
                              p.id::text AS project_id,
                              p.contract_number,
                              p.service,
                              p.current_step,
                              p.current_step_updated_at,
                              coalesce(p.closed, false) AS closed,
                              coalesce(company.name, company.ur_name) AS company_name,
                              concat_ws(' ', manager.first_name, manager.last_name) AS manager_name,
                              manager.telegram_id AS manager_telegram_id
                            FROM clients c
                            LEFT JOIN projects p
                              ON p.client_id = c.id
                             AND (p.group_id = 1 OR p.group_id IS NULL)
                            LEFT JOIN companies_clone company ON company.id = p.company_id
                            LEFT JOIN users manager ON manager.id = coalesce(p.manager_id, p.expert_id)
                            WHERE c.deleted_at IS NULL
                              AND c.archived_at IS NULL
                              AND (c.group_id = 1 OR c.group_id IS NULL)
                              AND (
                                regexp_replace(coalesce(c.phone, ''), '\\D', '', 'g') = :phone
                                OR regexp_replace(coalesce(c.second_phone, ''), '\\D', '', 'g') = :phone
                              )
                            ORDER BY c.id, coalesce(p.closed, false),
                                     p.current_step_updated_at DESC NULLS LAST,
                                     p.created_at DESC NULLS LAST
                            LIMIT 100
                            """
                        ),
                        {"phone": phone_digits},
                    ).mappings()
                    return self._group_contacts(rows)
        finally:
            engine.dispose()

    @staticmethod
    def _group_contacts(rows) -> dict[str, Any]:
        contacts: dict[str, dict[str, Any]] = {}
        for row in rows:
            client_id = str(row["client_id"])
            contact = contacts.setdefault(
                client_id,
                {
                    "id": client_id,
                    "name": row["client_name"],
                    "phone": row["phone"],
                    "second_phone": row["second_phone"],
                    "projects": [],
                },
            )
            if row["project_id"]:
                project_id = str(row["project_id"])
                contact["projects"].append(
                    {
                        "id": project_id,
                        "contract_number": row["contract_number"],
                        "service": row["service"],
                        "current_step": row["current_step"],
                        "current_step_updated_at": (
                            row["current_step_updated_at"].isoformat()
                            if row["current_step_updated_at"]
                            else None
                        ),
                        "closed": bool(row["closed"]),
                        "company_name": row["company_name"],
                        "manager_name": row["manager_name"] or None,
                        "manager_telegram_id": row["manager_telegram_id"],
                        "url": f"https://crm.lab-ecostar.ru/projects/{project_id}",
                    }
                )
        return {"contacts": list(contacts.values())}
