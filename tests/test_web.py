import asyncio
from datetime import date
from unittest.mock import AsyncMock

import httpx

from reporting_bot.config import Settings
from reporting_bot.web import create_app
from reporting_bot.agents_report import AgentsReportArtifact


def configured_settings() -> Settings:
    return Settings(
        telegram_bot_token="token",
        telegram_webhook_secret="secret",
        allowed_user_ids=frozenset({42}),
        database_url="postgresql+psycopg://user:pass@localhost/db",
        crm_bridge_token="bridge-secret",
    )


async def request(app, method, path, **kwargs):
    transport = httpx.ASGITransport(app=app)
    async with httpx.AsyncClient(
        transport=transport,
        base_url="http://test",
    ) as client:
        return await client.request(method, path, **kwargs)


def test_individual_agent_callback_sends_workbook_through_webhook(monkeypatch):
    calls = []
    pid = "12345678-1234-1234-1234-123456789012"
    def create(self, partner_id, start, end):
        calls.append((partner_id, start, end))
        return AgentsReportArtifact(b"test-xlsx", "agent.xlsx", "Agent report")
    document = AsyncMock()
    monkeypatch.setattr("reporting_bot.web.AgentReportService.create", create)
    monkeypatch.setattr("reporting_bot.web.TelegramClient.send_document", document)
    monkeypatch.setattr("reporting_bot.web.TelegramClient.send_message", AsyncMock())
    monkeypatch.setattr("reporting_bot.web.TelegramClient.answer_callback_query", AsyncMock())
    response = asyncio.run(request(create_app(configured_settings()), "POST", "/telegram/webhook",
        headers={"X-Telegram-Bot-Api-Secret-Token": "secret"},
        json={"callback_query": {"id": "cb", "from": {"id": 42},
              "message": {"chat": {"id": 42, "type": "private"}},
              "data": f"ag:go:20260801:20260831:{pid.replace('-', '')}"}}))
    assert response.status_code == 200
    assert calls == [(pid, date(2026, 8, 1), date(2026, 8, 31))]
    document.assert_awaited_once_with(42, b"test-xlsx", "agent.xlsx", "Agent report")


def test_health_does_not_expose_secrets() -> None:
    response = asyncio.run(
        request(create_app(configured_settings()), "GET", "/health")
    )
    assert response.status_code == 200
    assert response.json()["status"] == "ready"
    assert "token" not in response.text
    assert "postgresql" not in response.text


def test_webhook_rejects_invalid_secret() -> None:
    response = asyncio.run(
        request(
            create_app(configured_settings()),
            "POST",
            "/telegram/webhook",
            headers={"X-Telegram-Bot-Api-Secret-Token": "wrong"},
            json={"update_id": 1},
        )
    )
    assert response.status_code == 401


def test_webhook_accepts_setup_update_before_database_is_configured() -> None:
    setup_settings = Settings(
        telegram_bot_token="token",
        telegram_webhook_secret="secret",
        allowed_user_ids=frozenset(),
        database_url="",
    )
    response = asyncio.run(
        request(
            create_app(setup_settings),
            "POST",
            "/telegram/webhook",
            headers={"X-Telegram-Bot-Api-Secret-Token": "secret"},
            json={"update_id": 1},
        )
    )
    assert response.status_code == 200


def test_crm_schema_rejects_invalid_bridge_token() -> None:
    response = asyncio.run(
        request(
            create_app(configured_settings()),
            "GET",
            "/internal/crm/schema",
            headers={"X-CRM-Bridge-Token": "wrong"},
        )
    )
    assert response.status_code == 401


def test_crm_schema_returns_repository_payload(monkeypatch) -> None:
    monkeypatch.setattr(
        "reporting_bot.web.CrmBridgeRepository.schema",
        lambda self: {"columns": [{"table_name": "clients"}], "foreign_keys": []},
    )
    response = asyncio.run(
        request(
            create_app(configured_settings()),
            "GET",
            "/internal/crm/schema",
            headers={"X-CRM-Bridge-Token": "bridge-secret"},
        )
    )
    assert response.status_code == 200
    assert response.json()["columns"][0]["table_name"] == "clients"


def test_crm_contacts_returns_repository_payload(monkeypatch) -> None:
    monkeypatch.setattr(
        "reporting_bot.web.CrmBridgeRepository.find_contacts",
        lambda self, phone: {"contacts": [{"id": "client-1", "phone": phone}]},
    )
    response = asyncio.run(
        request(
            create_app(configured_settings()),
            "GET",
            "/internal/crm/contacts?phone=%2B79991234567",
            headers={"X-CRM-Bridge-Token": "bridge-secret"},
        )
    )
    assert response.status_code == 200
    assert response.json()["contacts"][0]["phone"] == "+79991234567"


def test_crm_contacts_rejects_missing_phone(monkeypatch) -> None:
    response = asyncio.run(
        request(
            create_app(configured_settings()),
            "GET",
            "/internal/crm/contacts",
            headers={"X-CRM-Bridge-Token": "bridge-secret"},
        )
    )
    assert response.status_code == 400
