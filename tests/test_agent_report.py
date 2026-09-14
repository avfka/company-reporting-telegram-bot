import asyncio
from datetime import date
from decimal import Decimal
from io import BytesIO
from unittest.mock import AsyncMock
from zipfile import ZipFile
import xml.etree.ElementTree as ET

import pytest

from reporting_bot.agent_report import (
    AgentOption, AgentReportData, AgentReportService, PROJECTS_SQL, TOTAL_SQL,
    PARTNER_SQL, SEARCH_SQL, build_agent_workbook, validate_period,
)
from reporting_bot.agent_dialog import search_prompt
from reporting_bot.config import Settings
from reporting_bot.reports import ReportCatalog
from reporting_bot.telegram import IncomingMessage, IncomingCallback, handle_message, handle_callback

PARTNER = AgentOption("12345678-1234-1234-1234-123456789012", "Тестовый агент")
START, END = date(2026, 8, 1), date(2026, 8, 31)
SETTINGS = Settings("token", "secret", frozenset({42}), "postgresql+psycopg://localhost/test")
NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def sample_data(empty=False):
    def project(id, amount, paid):
        return dict(project_id=id, created_at=START, service="sout", sale_price=Decimal("10000"),
                    company_name='ООО "Компания"', inn="0012345678", manager_sks="Имя Фамилия",
                    partner_name=PARTNER.name, paid_amount=Decimal(amount), contract_number="Д-1", is_paid=paid)
    return AgentReportData(PARTNER, START, END, () if empty else (
        project("id-1", "10000", True), project("id-2", "2500.50", False)), Decimal("25001"))


@pytest.mark.parametrize("empty", [False, True])
def test_workbook_reconciles_and_preserves_reference_columns(empty):
    with ZipFile(BytesIO(build_agent_workbook(sample_data(empty)))) as z:
        summary = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
        detail = ET.fromstring(z.read("xl/worksheets/sheet2.xml"))
        assert 'Сводка' in z.read("xl/workbook.xml").decode()
        assert float(summary.find(".//s:c[@r='C5']/s:v", NS).text) == (0 if empty else 12500.5)
        assert float(summary.find(".//s:c[@r='B5']/s:v", NS).text) == (0 if empty else 2)
        assert float(summary.find(".//s:c[@r='E5']/s:v", NS).text) == (0 if empty else .5)
        assert summary.find(".//s:c[@r='C5']/s:f", NS) is not None
        if not empty:
            assert detail.find(".//s:c[@r='A6']/s:v", NS) is not None  # numeric Excel date
            assert detail.find(".//s:c[@r='E6']/s:is/s:t", NS).text == "0012345678"
            assert detail.find(".//s:c[@r='K7']/s:is/s:t", NS).text == "Нет"  # partial stays
        assert detail.find("s:autoFilter", NS) is not None


def test_query_population_matches_template_and_does_not_duplicate_projects():
    assert "p.paid_amount > 0" in PROJECTS_SQL and "p.paid_amount > 0" in TOTAL_SQL
    assert "p.created_at >= :date_from AND p.created_at < :date_to_exclusive" in PROJECTS_SQL
    assert "WHERE p.partner_id = CAST(:partner_id AS uuid)" in PROJECTS_SQL
    assert "partner_fee" not in PROJECTS_SQL
    assert "p.is_paid IS TRUE" not in PROJECTS_SQL
    for query in (PROJECTS_SQL, TOTAL_SQL, PARTNER_SQL, SEARCH_SQL):
        assert "<> 2" in query
        assert "2025" not in query


@pytest.mark.parametrize("start,end", [(END, START), (START, date(2028, 1, 1))])
def test_bad_period_rejected_before_database(start, end):
    service = AgentReportService(SETTINGS)
    with pytest.raises(ValueError):
        service._load(PARTNER.partner_id, start, end)


def test_one_day_is_allowed_and_service_excludes_day_after_end(monkeypatch):
    validate_period(START, START)
    class Result:
        def mappings(self): return self
        def one_or_none(self): return dict(partner_id=PARTNER.partner_id, name=PARTNER.name)
        def scalar_one(self): return Decimal(0)
        def __iter__(self): return iter(())
    class Connection:
        def execute(self, sql, params):
            assert params["date_from"] == START
            assert params["date_to_exclusive"] == date(2026, 9, 1)
            return Result()
    service = AgentReportService(SETTINGS)
    monkeypatch.setattr(service, "_query", lambda operation: operation(Connection()))
    assert service._load(PARTNER.partner_id, START, END).projects == ()


def test_agent_missing_or_gto_rejected(monkeypatch):
    class Connection:
        def execute(self, *args): return self
        def mappings(self): return self
        def one_or_none(self): return None
    service = AgentReportService(SETTINGS)
    monkeypatch.setattr(service, "_query", lambda operation: operation(Connection()))
    with pytest.raises(ValueError, match="ГТО"):
        service._load(PARTNER.partner_id, START, END)


def test_search_escapes_wildcards_and_normalizes_yo(monkeypatch):
    class Connection:
        def execute(self, sql, params):
            assert params == {"search": "%алена\\_\\%%"}
            return self
        def mappings(self): return []
    service = AgentReportService(SETTINGS)
    monkeypatch.setattr(service, "_query", lambda operation: operation(Connection()))
    assert service.search(" Алёна_% ") == ()


def message(text, **kwargs):
    send, search, create = AsyncMock(), AsyncMock(return_value=(PARTNER,)), AsyncMock()
    asyncio.run(handle_message(IncomingMessage(42, 42, text, **kwargs), SETTINGS, ReportCatalog({}),
                               send, AsyncMock(), search_agents=search, send_agent_report=create))
    return send, search, create


def test_command_calendar_and_direct_dates():
    send, _, create = message("/report_agent")
    assert send.call_args.args[2]["inline_keyboard"][0][0]["callback_data"].startswith("agent:")
    create.assert_not_awaited()
    send, _, _ = message("/report_agent 2026-08-01 2026-08-31")
    assert send.call_args.args[2]["force_reply"] is True


def test_search_requires_confirmation_and_callback_fits():
    prompt, _ = search_prompt(START, END)
    send, search, create = message("Тестовый", reply_to_text=prompt)
    search.assert_awaited_once_with("Тестовый")
    button = send.call_args.args[2]["inline_keyboard"][0][0]
    assert len(button["callback_data"].encode()) <= 64
    create.assert_not_awaited()
    dispatch = AsyncMock()
    asyncio.run(handle_callback(IncomingCallback("cb", 42, 42, button["callback_data"]), SETTINGS,
                               AsyncMock(), AsyncMock(), AsyncMock(), send_agent_report=dispatch))
    dispatch.assert_awaited_once_with(42, PARTNER.partner_id, START, END)


@pytest.mark.parametrize("payload", ["agent:preset:2026-08-01:2026-08-31", "agent:to:2026-08-01:2026-08-31"])
def test_date_selection_prompts_for_agent_without_generating(payload):
    send, create = AsyncMock(), AsyncMock()
    asyncio.run(handle_callback(IncomingCallback("cb", 42, 42, payload), SETTINGS, send,
                               AsyncMock(), AsyncMock(), send_agent_report=create))
    assert send.call_args.args[2]["force_reply"]
    create.assert_not_awaited()


@pytest.mark.parametrize("user_id,chat_type", [(99, "private"), (42, "group")])
def test_agent_selection_enforces_authorization(user_id, chat_type):
    create = AsyncMock()
    asyncio.run(handle_callback(IncomingCallback("cb", 42, user_id,
        f"ag:go:20260801:20260831:{PARTNER.partner_id.replace('-', '')}", chat_type), SETTINGS,
        AsyncMock(), AsyncMock(), AsyncMock(), send_agent_report=create))
    create.assert_not_awaited()


@pytest.mark.parametrize("results,expected", [((), "не найден"), ((PARTNER,) * 21, "больше 20")])
def test_ambiguous_or_missing_search_does_not_silently_pick(results, expected):
    send = AsyncMock()
    asyncio.run(handle_message(IncomingMessage(42, 42, "Агент", reply_to_text=search_prompt(START, END)[0]),
        SETTINGS, ReportCatalog({}), send, AsyncMock(), search_agents=AsyncMock(return_value=results)))
    assert expected in send.call_args.args[1]
    assert send.call_args.args[2]["force_reply"]


@pytest.mark.parametrize("payload", ["ag:go:20260831:20260801:bad", "ag:go:20260801:20260831:bad", "ag:go:x"])
def test_malformed_selection_does_not_query(payload):
    send, create = AsyncMock(), AsyncMock()
    asyncio.run(handle_callback(IncomingCallback("cb", 42, 42, payload), SETTINGS, send,
                               AsyncMock(), AsyncMock(), send_agent_report=create))
    assert "/report_agent" in send.call_args.args[1]
    create.assert_not_awaited()
