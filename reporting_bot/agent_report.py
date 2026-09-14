"""Individual agent report. Period refers to project creation, not cash receipts."""
from __future__ import annotations

import html
import re
from dataclasses import dataclass
from datetime import date, timedelta
from decimal import Decimal
from typing import Mapping
from uuid import UUID

from sqlalchemy import create_engine, text
from sqlalchemy.pool import NullPool

from reporting_bot.agents_report import AgentsReportArtifact, _service_label, _title
from reporting_bot.config import Settings
from reporting_bot.simple_xlsx import Formula, STYLE, Workbook


SEARCH_SQL = """
SELECT id::text AS partner_id, name
FROM partners
WHERE coalesce(group_id, 0) <> 2
  AND replace(lower(name), 'ё', 'е') LIKE :search ESCAPE '\\'
ORDER BY name, id
LIMIT 21
"""

PARTNER_SQL = """
SELECT id::text AS partner_id, name FROM partners
WHERE id = CAST(:partner_id AS uuid) AND coalesce(group_id, 0) <> 2
"""

PROJECTS_SQL = """
SELECT p.id::text AS project_id, p.created_at, p.service, p.sale_price,
       c.name AS company_name, c.inn,
       concat_ws(' ', u.first_name, u.last_name) AS manager_sks,
       a.name AS partner_name, p.paid_amount::numeric AS paid_amount,
       p.contract_number, p.is_paid
FROM projects p
JOIN partners a ON a.id = p.partner_id
LEFT JOIN companies_clone c ON c.id = p.company_id
LEFT JOIN users u ON u.id = p.manager_sks_id
WHERE p.partner_id = CAST(:partner_id AS uuid)
  AND coalesce(a.group_id, 0) <> 2
  AND p.created_at >= :date_from AND p.created_at < :date_to_exclusive
  AND p.paid_amount > 0
ORDER BY p.created_at DESC, p.id
"""

TOTAL_SQL = """
SELECT coalesce(sum(p.paid_amount::numeric), 0) AS paid_amount
FROM projects p JOIN partners a ON a.id = p.partner_id
WHERE coalesce(a.group_id, 0) <> 2
  AND p.created_at >= :date_from AND p.created_at < :date_to_exclusive
  AND p.paid_amount > 0
"""


def validate_period(date_from: date, date_to: date) -> None:
    if date_to < date_from:
        raise ValueError("Дата окончания раньше даты начала.")
    if (date_to - date_from).days > 366:
        raise ValueError("Максимальный период отчёта — 366 дней.")


@dataclass(frozen=True)
class AgentOption:
    partner_id: str
    name: str


@dataclass(frozen=True)
class AgentReportData:
    partner: AgentOption
    date_from: date
    date_to: date
    projects: tuple[Mapping[str, object], ...]
    all_agents_paid_amount: Decimal


class AgentReportService:
    def __init__(self, settings: Settings) -> None:
        self.settings = settings

    def _query(self, operation):
        engine = create_engine(
            self.settings.database_url.replace("postgresql://", "postgresql+psycopg://", 1),
            poolclass=NullPool,
            connect_args={"connect_timeout": self.settings.connect_timeout_seconds},
        )
        try:
            with engine.connect() as connection, connection.begin():
                # One consistent snapshot for detail and share denominator.
                connection.execute(text("SET TRANSACTION ISOLATION LEVEL REPEATABLE READ, READ ONLY"))
                connection.execute(text("SELECT set_config('statement_timeout', :timeout, true)"),
                                   {"timeout": f"{max(self.settings.statement_timeout_ms, 30000)}ms"})
                return operation(connection)
        finally:
            engine.dispose()

    def search(self, query: str) -> tuple[AgentOption, ...]:
        query = " ".join(query.lower().replace("ё", "е").split())
        if not 2 <= len(query) <= 150:
            raise ValueError("Введите от 2 до 150 символов имени или названия агента.")
        escaped = query.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")
        return self._query(lambda c: tuple(AgentOption(**row) for row in c.execute(
            text(SEARCH_SQL), {"search": f"%{escaped}%"}).mappings()))

    def _load(self, partner_id: str, date_from: date, date_to: date) -> AgentReportData:
        validate_period(date_from, date_to)
        params = {"partner_id": str(UUID(partner_id)), "date_from": date_from,
                  "date_to_exclusive": date_to + timedelta(days=1)}

        def load(connection):
            partner = connection.execute(text(PARTNER_SQL), params).mappings().one_or_none()
            if partner is None:
                raise ValueError("Агент не найден или относится к группе ГТО. Выберите другого агента.")
            projects = tuple(dict(row) for row in connection.execute(text(PROJECTS_SQL), params).mappings())
            total = connection.execute(text(TOTAL_SQL), params).scalar_one()
            return AgentReportData(AgentOption(**partner), date_from, date_to, projects, Decimal(str(total)))

        return self._query(load)

    def create(self, partner_id: str, date_from: date, date_to: date) -> AgentsReportArtifact:
        data = self._load(partner_id, date_from, date_to)
        safe_name = re.sub(r'[^\w .-]', '_', data.partner.name)[:60].strip() or "Агент"
        return AgentsReportArtifact(
            build_agent_workbook(data),
            f"Отчет_Агент_{safe_name}_{date_from}_{date_to}.xlsx",
            f"<b>Отчёт по агенту</b> · {html.escape(data.partner.name)}\n"
            f"Проекты созданы {date_from:%d.%m.%Y}–{date_to:%d.%m.%Y}. "
            f"С положительной оплатой: {len(data.projects)}. Без ГТО.",
        )


def build_agent_workbook(data: AgentReportData) -> bytes:
    workbook = Workbook("Отчёт по агенту (оплаты)")
    summary = workbook.add_sheet("Сводка")
    detail = workbook.add_sheet("Проекты")
    period = f"Дата создания проекта: {data.date_from:%d.%m.%Y}–{data.date_to:%d.%m.%Y}"
    _title(summary, "Отчёт по агенту (оплаты)", period, 5)
    summary.widths = {0: 62, 1: 22, 2: 25, 3: 24, 4: 24}
    summary.append([])
    summary.append(["Агент", "Проекты с оплатой, шт.", "Оплаченная сумма, руб.", "Средний чек, руб.", "Доля в оплатах агентов"], STYLE["table_header"])
    summary.row_heights[4] = 36
    count = len(data.projects)
    amount = sum((Decimal(str(p["paid_amount"])) for p in data.projects), Decimal(0))
    last = max(6, 5 + count)
    summary.append([data.partner.name, Formula(f"COUNTA('Проекты'!I6:I{last})", count),
                    Formula(f"SUM('Проекты'!H6:H{last})", amount),
                    Formula('IF(B5=0,0,C5/B5)', amount / count if count else 0),
                    Formula('IF(C8=0,0,C5/C8)', amount / data.all_agents_paid_amount if data.all_agents_paid_amount else 0)],
                   [STYLE[k] for k in ("table_text", "table_number", "table_money", "table_money", "table_percent")])
    summary.row_heights[5] = max(48, 15 * ((len(data.partner.name) + 49) // 50))
    summary.append([])
    summary.append(["База для расчёта доли"], STYLE["section"])
    summary.append(["Оплаченная сумма всех агентов без ГТО", None, data.all_agents_paid_amount],
                   [STYLE["table_text"], STYLE["base"], STYLE["table_money"]])
    summary.append([])
    for note in (
        "Период отбирает проекты по дате создания включительно, а не платежи по дате поступления.",
        "Учитываются проекты с paid_amount > 0, включая частичную оплату. Зелёный флаг не обязателен.",
        "Оплаченная сумма — текущее накопленное значение в карточке проекта на момент выгрузки.",
        "Средний чек = оплаченная сумма / число проектов с оплатой. Компания может иметь несколько проектов.",
        "Доля = оплата выбранного агента / оплата всех агентов без ГТО за тот же период создания проектов.",
        "Дата создания самого агента не ограничена. Привязка агента берётся из текущей карточки проекта.",
    ):
        row = summary.append([note], STYLE["note"])
        summary.merges.append(f"A{row}:E{row}")
        summary.row_heights[row] = 24

    _title(detail, data.partner.name, period, 11)
    detail.append(["Источник: CRM. projects; partners; companies_clone; users. Одна строка — один проект."], STYLE["note"])
    detail.merges.append("A3:K3")
    detail.append([])
    detail.append(["Дата создания", "Услуга", "Стоимость, руб.", "Ком. название компании", "ИНН компании",
                   "Менеджер СКС", "Имя партнёра", "Оплаченная сумма, руб.", "ID проекта", "Номер договора", "Зелёный флаг оплаты"], STYLE["table_header"])
    detail.widths = dict(enumerate([16, 18, 20, 48, 18, 32, 48, 23, 40, 24, 19]))
    detail.freeze_rows = 5
    detail.row_heights[5] = 36
    styles = [STYLE[k] for k in ("table_date", "table_text", "table_money", "table_text", "table_id",
                                "table_text", "table_text", "table_money", "table_id", "table_text", "table_center")]
    for p in data.projects:
        row = detail.append([p["created_at"], _service_label(str(p["service"] or "")), p["sale_price"],
                             p["company_name"], p["inn"], p["manager_sks"], p["partner_name"], p["paid_amount"],
                             p["project_id"], p["contract_number"], "Да" if p["is_paid"] else "Нет"], styles)
        detail.row_heights[row] = max(32, 15 * max((len(str(p[k] or "")) + 39) // 40 for k in ("company_name", "partner_name")))
    if not count:
        detail.append([None] * 11)
        summary.append(["За выбранный период проектов с положительной оплатой не найдено."], STYLE["note"])
        summary.merges.append(f"A{len(summary.rows)}:E{len(summary.rows)}")
    detail.auto_filter = f"A5:K{last}"
    detail.append(["Итого", None, Formula(f"SUM(C6:C{last})", sum((Decimal(str(p["sale_price"] or 0)) for p in data.projects), Decimal(0))),
                   None, None, None, None, Formula(f"SUM(H6:H{last})", amount)],
                  [STYLE["total_text"]] * 2 + [STYLE["total_money"]] + [STYLE["total_text"]] * 4 + [STYLE["total_money"]])
    return workbook.to_bytes()
