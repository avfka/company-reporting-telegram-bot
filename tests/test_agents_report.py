from datetime import date, datetime
from decimal import Decimal
from zipfile import ZipFile
from io import BytesIO

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
