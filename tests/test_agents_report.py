from datetime import date, datetime
from decimal import Decimal
from zipfile import ZipFile
from io import BytesIO
import sqlite3

from reporting_bot.agents_report import (
    AGENT_FEES_SQL,
    PAID_PROJECTS_SQL,
    PARTNERS_SQL,
    CONTROL_SQL,
    PARTNER_MANAGER_ROLE,
    AgentFee,
    AgentsReportData,
    PaidPartnerProject,
    PartnerSummary,
    build_agents_workbook,
)


def test_all_report_queries_exclude_gto_partners() -> None:
    for query in (PARTNERS_SQL, PAID_PROJECTS_SQL, AGENT_FEES_SQL):
        assert "coalesce(" in query
        assert "<> :gto_group_id" in query


def test_all_agent_project_and_fee_queries_filter_project_responsible_role():
    assert PARTNER_MANAGER_ROLE == "partnerManager"
    for query, alias in ((PARTNERS_SQL, "p"), (PAID_PROJECTS_SQL, "project"),
                         (AGENT_FEES_SQL, "project"), (CONTROL_SQL, "p")):
        assert f"JOIN users responsible ON responsible.id = {alias}.manager_id" in query
        assert "responsible.role = :partner_manager_role" in query
        assert "responsible.id = partner.manager_id" not in query
        assert "responsible.id = project.manager_sks_id" not in query
    # Control excludes absent owners too, not just owners with a different role.
    assert "AND NOT EXISTS (" in CONTROL_SQL
    assert "excluded_non_partner_manager_projects" in CONTROL_SQL
    # The unpaid-fee diagnostic must use the same responsible filter.
    unpaid_query = CONTROL_SQL.split("AS mismatched_fee_rows,")[1]
    assert "responsible.id = project.manager_id" in unpaid_query
    assert "responsible.role = :partner_manager_role" in unpaid_query


def test_partner_owner_filter_applies_to_every_sheet_not_just_project_totals():
    for query, alias in ((PARTNERS_SQL, "p"), (PAID_PROJECTS_SQL, "p"),
                         (AGENT_FEES_SQL, "partner"), (CONTROL_SQL, "p")):
        assert f"JOIN users partner_owner ON partner_owner.id = {alias}.manager_id" in query
        assert "partner_owner.role = :partner_manager_role" in query


def test_mixed_partner_and_project_owners_do_not_leak_into_report():
    # Exercise the real joins/aggregations in an isolated database, adapting only
    # PostgreSQL cast syntax. No company database writes are needed for this test.
    with sqlite3.connect(":memory:") as db:
        db.row_factory = sqlite3.Row
        db.create_function("concat_ws", -1, lambda sep, *values: sep.join(v for v in values if v is not None))
        db.executescript("""
          CREATE TABLE users(id TEXT, first_name TEXT, last_name TEXT, role TEXT);
          INSERT INTO users VALUES ('pm','Марина','Менеджер','partnerManager'), ('other','Артур','Другой','manager');
          CREATE TABLE partners(id TEXT, name TEXT, manager_id TEXT, created_at TEXT, group_id INT,
            category TEXT, phone TEXT, email TEXT, contacts TEXT, archived_at TEXT);
          INSERT INTO partners(id,name,manager_id,created_at,group_id) VALUES
            ('good','Подходящий','pm','2026-01-01',1),
            ('wrong','Другой менеджер','other','2026-01-01',1),
            ('missing','Без менеджера',NULL,'2026-01-01',1),
            ('empty','Без проектов','pm','2026-01-01',1),
            ('gto','ГТО','pm','2026-01-01',2),
            ('old','Старый','pm','2025-09-30',1);
          CREATE TABLE projects(id TEXT,partner_id TEXT,manager_id TEXT,is_paid BOOLEAN,paid_amount REAL,
            sale_price REAL,created_at TEXT,contract_number TEXT,service TEXT,payment_date TEXT);
          INSERT INTO projects(id,partner_id,manager_id,is_paid,paid_amount) VALUES
            ('p1','good','pm',1,100), ('p2','good','other',1,200),
            ('p3','wrong','pm',1,300), ('p4','missing','pm',1,400),
            ('p5','gto','pm',1,500), ('p6','old','pm',1,600),
            ('p7','good','pm',0,700), ('p8','good',NULL,1,800);
          CREATE TABLE partner_fee(id TEXT,partner_id TEXT,project_id TEXT,amount REAL,fee_date TEXT,created_at TEXT);
          INSERT INTO partner_fee(id,partner_id,project_id,amount) VALUES
            ('f1','good','p1',10), ('f2','good','p2',20), ('f3','wrong','p3',30),
            ('f4','missing','p4',40), ('f5','gto','p5',50), ('f6','old','p6',60),
            ('f7','good','p7',70), ('f8','good','p1',5), ('f9','wrong','p1',90);
        """)
        params = dict(partner_created_from="2025-10-01", gto_group_id=2, partner_manager_role=PARTNER_MANAGER_ROLE)
        def run(query):
            query = query.replace("::text", "").replace("CAST(:partner_created_from AS DATE)", ":partner_created_from")
            return [dict(row) for row in db.execute(query, params)]
        partners = run(PARTNERS_SQL)
        assert {p["partner_id"] for p in partners} == {"good", "empty"}
        assert {p["manager"] for p in partners} == {"Менеджер Марина"}
        assert sum(p["paid_project_count"] for p in partners) == 1
        assert sum(p["paid_amount"] for p in partners) == 100
        assert sum(p["agent_fee_amount"] for p in partners) == 15
        projects = run(PAID_PROJECTS_SQL)
        assert [p["project_id"] for p in projects] == ["p1"]
        assert projects[0]["agent_fee_amount"] == 15
        assert {f["fee_id"] for f in run(AGENT_FEES_SQL)} == {"f1", "f8"}
        control = run(CONTROL_SQL)[0]
        assert control["excluded_non_partner_manager_partners"] == 2
        assert control["excluded_non_partner_manager_projects"] == 2
        assert control["excluded_gto_partners"] == 1
        assert control["fees_on_unpaid_projects"] == 1
        assert control["mismatched_fee_rows"] == 1


def test_agents_workbook_contains_expected_sheets_and_values() -> None:
    partner = PartnerSummary(
        partner_id="partner-1",
        created_at=datetime(2025, 10, 2, 9, 0),
        name="Партнёр Тест",
        category="Агент",
        phone="+70000000000",
        email="test@example.com",
        contacts="Контакт",
        manager="Иванов Иван",
        archived_at=None,
        paid_project_count=1,
        paid_amount=Decimal("100000"),
        fee_project_count=1,
        fee_count=1,
        agent_fee_amount=Decimal("15000"),
    )
    project = PaidPartnerProject(
        partner_id="partner-1",
        partner_name="Партнёр Тест",
        project_id="project-1",
        created_at=date(2025, 11, 1),
        contract_number="Д-1",
        service="opk",
        payment_date=date(2025, 11, 10),
        paid_amount=Decimal("100000"),
        sale_price=Decimal("120000"),
        fee_count=1,
        agent_fee_amount=Decimal("15000"),
    )
    fee = AgentFee(
        partner_id="partner-1",
        partner_name="Партнёр Тест",
        fee_id="fee-1",
        fee_date=date(2025, 11, 12),
        fee_created_at=datetime(2025, 11, 12, 10, 0),
        project_id="project-1",
        contract_number="Д-1",
        service="opk",
        project_paid_amount=Decimal("100000"),
        agent_fee_amount=Decimal("15000"),
    )
    content = build_agents_workbook(
        AgentsReportData(
            generated_on=date(2026, 8, 28),
            partners=(partner,),
            projects=(project,),
            fees=(fee,),
            control={
                "excluded_gto_partners": 2,
                "paid_projects_without_amount": 0,
                "mismatched_fee_rows": 0,
                "fees_on_unpaid_projects": 0,
            },
        )
    )

    with ZipFile(BytesIO(content)) as archive:
        workbook_xml = archive.read("xl/workbook.xml").decode("utf-8")
        shared = "".join(
            archive.read(name).decode("utf-8")
            for name in archive.namelist()
            if name.startswith("xl/worksheets/sheet")
        )

    for sheet_name in ("Сводка", "Партнёры", "Оплаченные проекты", "Агентские выплаты", "Контроль"):
        assert sheet_name in workbook_xml
    assert "Партнёр Тест" in shared
    assert "100000" in shared
    assert "15000" in shared
    assert "ГТО" in shared
    assert "Менеджер партнёров" in shared
    assert "Оплаченные проекты других ответственных" in shared
