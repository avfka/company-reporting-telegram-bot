from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
from typing import Mapping, Sequence
from zoneinfo import ZoneInfo

from sqlalchemy import create_engine, text
from sqlalchemy.pool import NullPool

from reporting_bot.config import Settings
from reporting_bot.simple_xlsx import STYLE, Workbook


PARTNER_CREATED_FROM = date(2025, 10, 1)
GTO_GROUP_ID = 2
PARTNER_MANAGER_ROLE = "partnerManager"


PARTNERS_SQL = """
WITH selected_partners AS (
  SELECT p.*
  FROM partners p
  WHERE p.created_at >= CAST(:partner_created_from AS DATE)
    AND coalesce(p.group_id, 0) <> :gto_group_id
), paid_projects AS (
  SELECT p.*
  FROM projects p
  JOIN selected_partners partner ON partner.id = p.partner_id
  JOIN users responsible ON responsible.id = p.manager_id
    AND responsible.role = :partner_manager_role
  WHERE p.is_paid IS TRUE
), project_totals AS (
  SELECT
    partner_id,
    count(*) AS paid_project_count,
    sum(coalesce(paid_amount, 0)) AS paid_amount
  FROM paid_projects
  GROUP BY partner_id
), fee_totals AS (
  SELECT
    fee.partner_id,
    count(DISTINCT fee.project_id) AS fee_project_count,
    count(*) AS fee_count,
    sum(coalesce(fee.amount, 0)) AS agent_fee_amount
  FROM partner_fee fee
  JOIN paid_projects project
    ON project.id = fee.project_id
   AND project.partner_id = fee.partner_id
  GROUP BY fee.partner_id
)
SELECT
  partner.id::text AS partner_id,
  partner.created_at,
  partner.name,
  partner.category,
  partner.phone,
  partner.email,
  partner.contacts::text AS contacts,
  concat_ws(' ', manager.last_name, manager.first_name) AS manager,
  partner.archived_at,
  coalesce(project_totals.paid_project_count, 0) AS paid_project_count,
  coalesce(project_totals.paid_amount, 0) AS paid_amount,
  coalesce(fee_totals.fee_project_count, 0) AS fee_project_count,
  coalesce(fee_totals.fee_count, 0) AS fee_count,
  coalesce(fee_totals.agent_fee_amount, 0) AS agent_fee_amount
FROM selected_partners partner
LEFT JOIN users manager ON manager.id = partner.manager_id
LEFT JOIN project_totals ON project_totals.partner_id = partner.id
LEFT JOIN fee_totals ON fee_totals.partner_id = partner.id
ORDER BY partner.created_at, partner.name, partner.id
"""


PAID_PROJECTS_SQL = """
WITH selected_partners AS (
  SELECT p.id, p.name
  FROM partners p
  WHERE p.created_at >= CAST(:partner_created_from AS DATE)
    AND coalesce(p.group_id, 0) <> :gto_group_id
), fee_totals AS (
  SELECT
    fee.project_id,
    fee.partner_id,
    count(*) AS fee_count,
    sum(coalesce(fee.amount, 0)) AS agent_fee_amount
  FROM partner_fee fee
  GROUP BY fee.project_id, fee.partner_id
)
SELECT
  partner.id::text AS partner_id,
  partner.name AS partner_name,
  project.id::text AS project_id,
  project.created_at,
  project.contract_number,
  project.service,
  project.payment_date,
  coalesce(project.paid_amount, 0) AS paid_amount,
  coalesce(project.sale_price, 0) AS sale_price,
  coalesce(fee_totals.fee_count, 0) AS fee_count,
  coalesce(fee_totals.agent_fee_amount, 0) AS agent_fee_amount
FROM projects project
JOIN selected_partners partner ON partner.id = project.partner_id
JOIN users responsible ON responsible.id = project.manager_id
  AND responsible.role = :partner_manager_role
LEFT JOIN fee_totals
  ON fee_totals.project_id = project.id
 AND fee_totals.partner_id = partner.id
WHERE project.is_paid IS TRUE
ORDER BY partner.name, project.payment_date, project.created_at, project.id
"""


AGENT_FEES_SQL = """
SELECT
  partner.id::text AS partner_id,
  partner.name AS partner_name,
  fee.id::text AS fee_id,
  fee.fee_date,
  fee.created_at AS fee_created_at,
  project.id::text AS project_id,
  project.contract_number,
  project.service,
  coalesce(project.paid_amount, 0) AS project_paid_amount,
  coalesce(fee.amount, 0) AS agent_fee_amount
FROM partner_fee fee
JOIN partners partner ON partner.id = fee.partner_id
JOIN projects project
  ON project.id = fee.project_id
 AND project.partner_id = fee.partner_id
JOIN users responsible ON responsible.id = project.manager_id
  AND responsible.role = :partner_manager_role
WHERE partner.created_at >= CAST(:partner_created_from AS DATE)
  AND coalesce(partner.group_id, 0) <> :gto_group_id
  AND project.is_paid IS TRUE
ORDER BY partner.name, fee.fee_date, fee.created_at, fee.id
"""


CONTROL_SQL = """
WITH selected_partners AS (
  SELECT p.id
  FROM partners p
  WHERE p.created_at >= CAST(:partner_created_from AS DATE)
    AND coalesce(p.group_id, 0) <> :gto_group_id
), paid_projects AS (
  SELECT p.id, p.partner_id, coalesce(p.paid_amount, 0) AS paid_amount
  FROM projects p
  JOIN selected_partners partner ON partner.id = p.partner_id
  JOIN users responsible ON responsible.id = p.manager_id
    AND responsible.role = :partner_manager_role
  WHERE p.is_paid IS TRUE
)
SELECT
  (
    SELECT count(*)
    FROM partners p
    WHERE p.created_at >= CAST(:partner_created_from AS DATE)
      AND p.group_id = :gto_group_id
  ) AS excluded_gto_partners,
  (
    SELECT count(*)
    FROM projects project
    JOIN selected_partners partner ON partner.id = project.partner_id
    WHERE project.is_paid IS TRUE
      AND NOT EXISTS (
        SELECT 1 FROM users responsible
        WHERE responsible.id = project.manager_id
          AND responsible.role = :partner_manager_role
      )
  ) AS excluded_non_partner_manager_projects,
  (
    SELECT count(*)
    FROM paid_projects
    WHERE paid_amount <= 0
  ) AS paid_projects_without_amount,
  (
    SELECT count(*)
    FROM partner_fee fee
    JOIN paid_projects project ON project.id = fee.project_id
    WHERE fee.partner_id <> project.partner_id
  ) AS mismatched_fee_rows,
  (
    SELECT count(*)
    FROM partner_fee fee
    JOIN projects project
      ON project.id = fee.project_id
     AND project.partner_id = fee.partner_id
    JOIN selected_partners partner ON partner.id = fee.partner_id
    JOIN users responsible ON responsible.id = project.manager_id
      AND responsible.role = :partner_manager_role
    WHERE project.is_paid IS NOT TRUE
  ) AS fees_on_unpaid_projects
"""


SERVICE_LABELS = {
    "sout": "СОУТ",
    "opk": "ОПР",
    "opk_eth": "ОПР (ЭТХ)",
    "autsorsing": "Аутсорсинг",
    "suot": "СУОТ",
    "obuchenie": "Обучение",
    "audit": "Аудит",
    "other": "Другие услуги",
}


@dataclass(frozen=True)
class PartnerSummary:
    partner_id: str
    created_at: datetime
    name: str
    category: str
    phone: str
    email: str
    contacts: str
    manager: str
    archived_at: datetime | None
    paid_project_count: int
    paid_amount: Decimal
    fee_project_count: int
    fee_count: int
    agent_fee_amount: Decimal


@dataclass(frozen=True)
class PaidPartnerProject:
    partner_id: str
    partner_name: str
    project_id: str
    created_at: date
    contract_number: str
    service: str
    payment_date: date | None
    paid_amount: Decimal
    sale_price: Decimal
    fee_count: int
    agent_fee_amount: Decimal


@dataclass(frozen=True)
class AgentFee:
    partner_id: str
    partner_name: str
    fee_id: str
    fee_date: date | None
    fee_created_at: datetime
    project_id: str
    contract_number: str
    service: str
    project_paid_amount: Decimal
    agent_fee_amount: Decimal


@dataclass(frozen=True)
class AgentsReportData:
    generated_on: date
    partners: tuple[PartnerSummary, ...]
    projects: tuple[PaidPartnerProject, ...]
    fees: tuple[AgentFee, ...]
    control: Mapping[str, int]


@dataclass(frozen=True)
class AgentsReportArtifact:
    workbook: bytes
    workbook_filename: str
    caption: str


class AgentsReportService:
    def __init__(self, settings: Settings) -> None:
        self._settings = settings
        self._database_url = settings.database_url.replace(
            "postgresql://", "postgresql+psycopg://", 1
        )

    def create(self) -> AgentsReportArtifact:
        data = self._load()
        stem = (
            f"Отчет_Партнеры_с_{PARTNER_CREATED_FROM.isoformat()}_"
            f"на_{data.generated_on.isoformat()}"
        )
        return AgentsReportArtifact(
            workbook=build_agents_workbook(data),
            workbook_filename=stem + ".xlsx",
            caption=(
                f"<b>Отчёт по партнёрам</b> · созданы с "
                f"{PARTNER_CREATED_FROM:%d.%m.%Y} · без ГТО\n"
                "Только проекты с ответственным в роли «Менеджер партнёров»."
            ),
        )

    def _load(self) -> AgentsReportData:
        engine = create_engine(
            self._database_url,
            poolclass=NullPool,
            connect_args={"connect_timeout": self._settings.connect_timeout_seconds},
        )
        params = {
            "partner_created_from": PARTNER_CREATED_FROM,
            "gto_group_id": GTO_GROUP_ID,
            "partner_manager_role": PARTNER_MANAGER_ROLE,
        }
        try:
            with engine.connect() as connection:
                with connection.begin():
                    connection.execute(text("SET TRANSACTION ISOLATION LEVEL REPEATABLE READ, READ ONLY"))
                    connection.execute(
                        text("SELECT set_config('statement_timeout', :timeout, true)"),
                        {"timeout": f"{max(self._settings.statement_timeout_ms, 30_000)}ms"},
                    )
                    partner_rows = connection.execute(text(PARTNERS_SQL), params).mappings()
                    partners = tuple(_partner_from_row(row) for row in partner_rows)
                    project_rows = connection.execute(text(PAID_PROJECTS_SQL), params).mappings()
                    projects = tuple(_project_from_row(row) for row in project_rows)
                    fee_rows = connection.execute(text(AGENT_FEES_SQL), params).mappings()
                    fees = tuple(_fee_from_row(row) for row in fee_rows)
                    control_row = connection.execute(text(CONTROL_SQL), params).mappings().one()
        finally:
            engine.dispose()
        return AgentsReportData(
            generated_on=datetime.now(ZoneInfo("Europe/Moscow")).date(),
            partners=partners,
            projects=projects,
            fees=fees,
            control={key: int(value or 0) for key, value in control_row.items()},
        )


def _decimal(value: object) -> Decimal:
    return Decimal(str(value or 0))


def _text(value: object) -> str:
    return str(value or "").strip()


def _partner_from_row(row: Mapping[str, object]) -> PartnerSummary:
    return PartnerSummary(
        partner_id=_text(row["partner_id"]),
        created_at=row["created_at"],
        name=_text(row["name"]) or "—",
        category=_text(row["category"]) or "—",
        phone=_text(row["phone"]) or "—",
        email=_text(row["email"]) or "—",
        contacts=_text(row["contacts"]) or "—",
        manager=_text(row["manager"]) or "—",
        archived_at=row["archived_at"],
        paid_project_count=int(row["paid_project_count"] or 0),
        paid_amount=_decimal(row["paid_amount"]),
        fee_project_count=int(row["fee_project_count"] or 0),
        fee_count=int(row["fee_count"] or 0),
        agent_fee_amount=_decimal(row["agent_fee_amount"]),
    )


def _project_from_row(row: Mapping[str, object]) -> PaidPartnerProject:
    return PaidPartnerProject(
        partner_id=_text(row["partner_id"]),
        partner_name=_text(row["partner_name"]) or "—",
        project_id=_text(row["project_id"]),
        created_at=row["created_at"],
        contract_number=_text(row["contract_number"]) or "—",
        service=_text(row["service"]),
        payment_date=row["payment_date"],
        paid_amount=_decimal(row["paid_amount"]),
        sale_price=_decimal(row["sale_price"]),
        fee_count=int(row["fee_count"] or 0),
        agent_fee_amount=_decimal(row["agent_fee_amount"]),
    )


def _fee_from_row(row: Mapping[str, object]) -> AgentFee:
    return AgentFee(
        partner_id=_text(row["partner_id"]),
        partner_name=_text(row["partner_name"]) or "—",
        fee_id=_text(row["fee_id"]),
        fee_date=row["fee_date"],
        fee_created_at=row["fee_created_at"],
        project_id=_text(row["project_id"]),
        contract_number=_text(row["contract_number"]) or "—",
        service=_text(row["service"]),
        project_paid_amount=_decimal(row["project_paid_amount"]),
        agent_fee_amount=_decimal(row["agent_fee_amount"]),
    )


def _excel_column(count: int) -> str:
    result = ""
    while count:
        count, remainder = divmod(count - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _title(sheet, title: str, subtitle: str, columns: int) -> None:
    sheet.append([title] + [None] * (columns - 1), STYLE["title"])
    sheet.append([subtitle] + [None] * (columns - 1), STYLE["subtitle"])
    last_column = _excel_column(columns)
    sheet.merges.extend([f"A1:{last_column}1", f"A2:{last_column}2"])
    sheet.row_heights[1] = 27
    sheet.row_heights[2] = 22


def _service_label(value: str) -> str:
    return SERVICE_LABELS.get(value, value or "—")


def _ratio(value: Decimal, baseline: Decimal) -> float | str:
    return float(value / baseline) if baseline else "—"


def build_agents_workbook(data: AgentsReportData) -> bytes:
    subtitle = (
        f"Партнёры созданы с {PARTNER_CREATED_FROM:%d.%m.%Y} · "
        f"состояние базы на {data.generated_on:%d.%m.%Y} · ГТО исключено"
    )
    workbook = Workbook("Отчёт по партнёрам")
    _build_summary_sheet(workbook, data, subtitle)
    _build_partners_sheet(workbook, data.partners, subtitle)
    _build_projects_sheet(workbook, data.projects, subtitle)
    _build_fees_sheet(workbook, data.fees, subtitle)
    _build_control_sheet(workbook, data, subtitle)
    return workbook.to_bytes()


def _build_summary_sheet(workbook: Workbook, data: AgentsReportData, subtitle: str) -> None:
    sheet = workbook.add_sheet("Сводка")
    _title(sheet, "Отчёт по партнёрам", subtitle, 7)
    sheet.append([None] * 7)
    sheet.append(
        ["Показатель", "Значение", "Единица", "Источник", None, None, None],
        STYLE["table_header"],
    )
    paid_amount = sum((row.paid_amount for row in data.partners), Decimal(0))
    fee_amount = sum((row.agent_fee_amount for row in data.partners), Decimal(0))
    paid_projects = sum(row.paid_project_count for row in data.partners)
    values = (
        ("Партнёры в отчёте", len(data.partners), "шт.", "partners.created_at; без group_id = 2", STYLE["table_number"]),
        ("Партнёры с оплаченными проектами", sum(row.paid_project_count > 0 for row in data.partners), "шт.", "projects.partner_id + is_paid", STYLE["table_number"]),
        ("Оплаченные проекты", paid_projects, "шт.", "projects.is_paid = true", STYLE["table_number"]),
        ("Оплаченная сумма", paid_amount, "руб.", "projects.paid_amount", STYLE["table_money"]),
        ("Проекты с агентскими выплатами", sum(row.fee_project_count for row in data.partners), "шт.", "partner_fee.project_id", STYLE["table_number"]),
        ("Агентские выплаты", fee_amount, "руб.", "partner_fee.amount", STYLE["table_money"]),
        ("После агентских выплат", paid_amount - fee_amount, "руб.", "Оплаченная сумма − выплаты", STYLE["table_money"]),
        ("Доля агентских выплат", _ratio(fee_amount, paid_amount), "%", "Выплаты / оплаченная сумма", STYLE["table_percent"]),
    )
    for label, value, unit, source, value_style in values:
        sheet.append(
            [label, value, unit, source, None, None, None],
            [STYLE["table_text"], value_style, STYLE["table_center"], STYLE["table_text"], STYLE["base"], STYLE["base"], STYLE["base"]],
        )
    sheet.append([None] * 7)
    note_row = sheet.append(
        [
            "В отчёт включены все партнёры, созданные с 01.10.2025, кроме группы ГТО. "
            "Проект относится к партнёру по projects.partner_id. Учитываются только проекты, у которых "
            "ответственный (projects.manager_id) имеет роль «Менеджер партнёров» (partnerManager), и is_paid = true. "
            "Сумма оплаты берётся из paid_amount. Агентская выплата учитывается только когда partner_fee связан "
            "с тем же партнёром и оплаченным проектом."
        ] + [None] * 6,
        STYLE["note"],
    )
    sheet.merges.append(f"A{note_row}:G{note_row}")
    sheet.row_heights[note_row] = 96
    sheet.widths = {0: 35, 1: 21, 2: 14, 3: 38, 4: 12, 5: 12, 6: 12}
    sheet.freeze_rows = 4


def _build_partners_sheet(workbook: Workbook, rows: Sequence[PartnerSummary], subtitle: str) -> None:
    columns = 15
    sheet = workbook.add_sheet("Партнёры")
    _title(sheet, "Партнёры", subtitle, columns)
    sheet.append([None] * columns)
    sheet.append(
        [
            "Дата создания", "Партнёр", "Категория", "Менеджер", "Телефон", "Email", "Контакты", "Статус",
            "Оплаченные проекты, шт.", "Оплаченная сумма, руб.", "Проекты с выплатами, шт.", "Выплаты, записей",
            "Агентские выплаты, руб.", "После выплат, руб.", "Доля выплат",
        ],
        STYLE["table_header"],
    )
    for row in rows:
        balance = row.paid_amount - row.agent_fee_amount
        sheet.append(
            [
                row.created_at, row.name, row.category, row.manager, row.phone, row.email, row.contacts,
                "Архивный" if row.archived_at else "Активный", row.paid_project_count, row.paid_amount,
                row.fee_project_count, row.fee_count, row.agent_fee_amount, balance,
                _ratio(row.agent_fee_amount, row.paid_amount),
            ],
            [
                STYLE["table_center"], STYLE["table_text"], STYLE["table_text"], STYLE["table_text"],
                STYLE["table_text"], STYLE["table_text"], STYLE["table_text"], STYLE["table_center"],
                STYLE["table_number"], STYLE["table_money"], STYLE["table_number"], STYLE["table_number"],
                STYLE["table_money"], STYLE["table_money"], STYLE["table_percent"],
            ],
        )
    sheet.append(
        [
            "Итого", None, None, None, None, None, None, None,
            sum(row.paid_project_count for row in rows),
            sum((row.paid_amount for row in rows), Decimal(0)),
            sum(row.fee_project_count for row in rows),
            sum(row.fee_count for row in rows),
            sum((row.agent_fee_amount for row in rows), Decimal(0)),
            sum((row.paid_amount - row.agent_fee_amount for row in rows), Decimal(0)),
            _ratio(
                sum((row.agent_fee_amount for row in rows), Decimal(0)),
                sum((row.paid_amount for row in rows), Decimal(0)),
            ),
        ],
        [
            STYLE["total_text"], STYLE["base"], STYLE["base"], STYLE["base"], STYLE["base"], STYLE["base"],
            STYLE["base"], STYLE["base"], STYLE["total_number"], STYLE["total_money"], STYLE["total_number"],
            STYLE["total_number"], STYLE["total_money"], STYLE["total_money"], STYLE["total_percent"],
        ],
    )
    sheet.widths = {0: 18, 1: 37, 2: 18, 3: 28, 4: 19, 5: 27, 6: 34, 7: 14, 8: 22, 9: 22, 10: 24, 11: 20, 12: 23, 13: 23, 14: 17}
    sheet.freeze_rows = 4
    sheet.auto_filter = f"A4:O{max(4, len(sheet.rows) - 1)}"


def _build_projects_sheet(workbook: Workbook, rows: Sequence[PaidPartnerProject], subtitle: str) -> None:
    columns = 12
    sheet = workbook.add_sheet("Оплаченные проекты")
    _title(sheet, "Оплаченные проекты", subtitle, columns)
    sheet.append([None] * columns)
    sheet.append(
        [
            "Партнёр", "Проект ID", "Дата проекта", "Договор", "Услуга", "Дата оплаты",
            "Оплачено, руб.", "Стоимость, руб.", "Выплат, записей", "Агентские выплаты, руб.",
            "После выплат, руб.", "Доля выплат",
        ],
        STYLE["table_header"],
    )
    for row in rows:
        sheet.append(
            [
                row.partner_name, row.project_id, row.created_at, row.contract_number, _service_label(row.service),
                row.payment_date or "—", row.paid_amount, row.sale_price, row.fee_count, row.agent_fee_amount,
                row.paid_amount - row.agent_fee_amount, _ratio(row.agent_fee_amount, row.paid_amount),
            ],
            [
                STYLE["table_text"], STYLE["table_text"], STYLE["table_center"], STYLE["table_text"],
                STYLE["table_text"], STYLE["table_center"], STYLE["table_money"], STYLE["table_money"],
                STYLE["table_number"], STYLE["table_money"], STYLE["table_money"], STYLE["table_percent"],
            ],
        )
    sheet.append(
        [
            "Итого", None, None, None, None, None,
            sum((row.paid_amount for row in rows), Decimal(0)),
            sum((row.sale_price for row in rows), Decimal(0)),
            sum(row.fee_count for row in rows),
            sum((row.agent_fee_amount for row in rows), Decimal(0)),
            sum((row.paid_amount - row.agent_fee_amount for row in rows), Decimal(0)),
            _ratio(
                sum((row.agent_fee_amount for row in rows), Decimal(0)),
                sum((row.paid_amount for row in rows), Decimal(0)),
            ),
        ],
        [
            STYLE["total_text"], STYLE["base"], STYLE["base"], STYLE["base"], STYLE["base"], STYLE["base"],
            STYLE["total_money"], STYLE["total_money"], STYLE["total_number"], STYLE["total_money"],
            STYLE["total_money"], STYLE["total_percent"],
        ],
    )
    sheet.widths = {0: 38, 1: 38, 2: 17, 3: 23, 4: 19, 5: 17, 6: 21, 7: 21, 8: 19, 9: 24, 10: 23, 11: 17}
    sheet.freeze_rows = 4
    sheet.auto_filter = f"A4:L{max(4, len(sheet.rows) - 1)}"


def _build_fees_sheet(workbook: Workbook, rows: Sequence[AgentFee], subtitle: str) -> None:
    columns = 10
    sheet = workbook.add_sheet("Агентские выплаты")
    _title(sheet, "Агентские выплаты", subtitle, columns)
    sheet.append([None] * columns)
    sheet.append(
        [
            "Партнёр", "Дата выплаты", "Дата записи", "Выплата ID", "Проект ID", "Договор",
            "Услуга", "Оплачено по проекту, руб.", "Агентская выплата, руб.", "Доля от оплаты",
        ],
        STYLE["table_header"],
    )
    for row in rows:
        sheet.append(
            [
                row.partner_name, row.fee_date or "—", row.fee_created_at, row.fee_id, row.project_id,
                row.contract_number, _service_label(row.service), row.project_paid_amount,
                row.agent_fee_amount, _ratio(row.agent_fee_amount, row.project_paid_amount),
            ],
            [
                STYLE["table_text"], STYLE["table_center"], STYLE["table_center"], STYLE["table_text"],
                STYLE["table_text"], STYLE["table_text"], STYLE["table_text"], STYLE["table_money"],
                STYLE["table_money"], STYLE["table_percent"],
            ],
        )
    sheet.append(
        [
            "Итого", None, None, None, None, None, None, None,
            sum((row.agent_fee_amount for row in rows), Decimal(0)), None,
        ],
        [
            STYLE["total_text"], STYLE["base"], STYLE["base"], STYLE["base"], STYLE["base"],
            STYLE["base"], STYLE["base"], STYLE["base"], STYLE["total_money"], STYLE["base"],
        ],
    )
    sheet.widths = {0: 38, 1: 17, 2: 20, 3: 38, 4: 38, 5: 23, 6: 19, 7: 25, 8: 25, 9: 19}
    sheet.freeze_rows = 4
    sheet.auto_filter = f"A4:J{max(4, len(sheet.rows) - 1)}"


def _build_control_sheet(workbook: Workbook, data: AgentsReportData, subtitle: str) -> None:
    sheet = workbook.add_sheet("Контроль")
    _title(sheet, "Контроль расчёта", subtitle, 6)
    sheet.append([None] * 6)
    sheet.append(["Проверка", "Значение", "Ожидание", "Результат", "Что сделано", None], STYLE["table_header"])
    checks = (
        ("Партнёры ГТО", data.control.get("excluded_gto_partners", 0), "Не включать", "Исключено", "Исключены из всех листов"),
        ("Оплаченные проекты других ответственных", data.control.get("excluded_non_partner_manager_projects", 0), "Не включать", "Исключено", "Ответственный не имеет роли «Менеджер партнёров» или не указан"),
        ("Оплаченные проекты без суммы", data.control.get("paid_projects_without_amount", 0), "0", None, "Оставлены для прозрачности с суммой 0"),
        ("Выплата закреплена не за партнёром проекта", data.control.get("mismatched_fee_rows", 0), "0", None, "Не включена в агентские выплаты"),
        ("Выплаты по проектам без зелёного флага", data.control.get("fees_on_unpaid_projects", 0), "Не включать", "Исключено", "Исключены из итогов"),
    )
    for label, value, expected, fixed_result, action in checks:
        result = fixed_result or ("OK" if value == 0 else "Требует внимания")
        sheet.append(
            [label, value, expected, result, action, None],
            [STYLE["table_text"], STYLE["table_number"], STYLE["table_center"], STYLE["table_center"], STYLE["table_text"], STYLE["base"]],
        )
    sheet.append([None] * 6)
    rule_row = sheet.append(["Правила включения"] + [None] * 5, STYLE["section"])
    sheet.merges.append(f"A{rule_row}:F{rule_row}")
    rules = (
        "Партнёр: partners.created_at ≥ 01.10.2025 и group_id ≠ 2 (ГТО).",
        "Партнёр проекта: projects.partner_id.",
        "Ответственный проекта: projects.manager_id → users.role = partnerManager («Менеджер партнёров»).",
        "Роль ответственного партнёра (partners.manager_id) и менеджера СКС не используется для этого фильтра.",
        "Партнёры без подходящих проектов остаются в списке с нулевыми показателями.",
        "Оплаченный проект: projects.is_paid = true; сумма: projects.paid_amount.",
        "Агентская выплата: partner_fee.amount только при совпадении partner_id у выплаты и проекта.",
        "Архивные партнёры не исключаются: их статус показан на листе «Партнёры».",
    )
    for rule in rules:
        row_number = sheet.append([rule] + [None] * 5, STYLE["note"])
        sheet.merges.append(f"A{row_number}:F{row_number}")
        sheet.row_heights[row_number] = 30
    sheet.widths = {0: 43, 1: 17, 2: 19, 3: 23, 4: 45, 5: 12}
    sheet.freeze_rows = 4
