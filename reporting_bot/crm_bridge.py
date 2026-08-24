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

    def find_staff(self, telegram_id: int) -> dict[str, Any]:
        if telegram_id <= 0:
            raise ValueError("Telegram ID must be a positive integer")

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
                    row = connection.execute(
                        text(
                            """
                            SELECT id::text AS id, telegram_id, email,
                                   first_name, last_name, is_active, role,
                                   group_id
                            FROM users
                            WHERE telegram_id::text = :telegram_id
                              AND coalesce(is_active, true) = true
                            ORDER BY id
                            LIMIT 1
                            """
                        ),
                        {"telegram_id": str(telegram_id)},
                    ).mappings().first()
                    if row is None:
                        return {"staff": None}
                    return {
                        "staff": {
                            "id": str(row["id"]),
                            "telegram_id": row["telegram_id"],
                            "email": row["email"],
                            "first_name": row["first_name"],
                            "last_name": row["last_name"],
                            "name": " ".join(
                                part
                                for part in (row["first_name"], row["last_name"])
                                if part
                            )
                            or None,
                            "role": row["role"],
                            "group_id": row["group_id"],
                        }
                    }
        finally:
            engine.dispose()

    def find_contacts(
        self,
        phone: str,
        staff_telegram_id: int | None = None,
    ) -> dict[str, Any]:
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
                            WITH selected_staff AS (
                              SELECT id
                              FROM users
                              WHERE telegram_id::text = :staff_telegram_id
                                AND coalesce(is_active, true) = true
                              LIMIT 1
                            )
                            SELECT
                              c.id::text AS client_id,
                              c.name AS client_name,
                              c.phone,
                              c.second_phone,
                              p.id::text AS project_id,
                              p.company_id::text AS company_id,
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
                              AND (
                                :staff_telegram_id = ''
                                OR EXISTS (
                                  SELECT 1
                                  FROM selected_staff staff
                                  WHERE staff.id = p.manager_id
                                     OR staff.id = p.expert_id
                                     OR staff.id = p.measurer_id
                                     OR staff.id = p.manager_sks_id
                                )
                              )
                            ORDER BY c.id, coalesce(p.closed, false),
                                     p.current_step_updated_at DESC NULLS LAST,
                                     p.created_at DESC NULLS LAST
                            LIMIT 100
                            """
                        ),
                        {
                            "phone": phone_digits,
                            "staff_telegram_id": (
                                str(staff_telegram_id)
                                if staff_telegram_id is not None
                                else ""
                            ),
                        },
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
                        "company_id": row["company_id"],
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
