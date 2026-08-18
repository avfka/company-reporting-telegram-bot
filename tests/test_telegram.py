import asyncio

from reporting_bot.config import Settings
from reporting_bot.database import QueryResult
from reporting_bot.reports import Report, ReportCatalog
from reporting_bot.telegram import IncomingMessage, handle_message


def settings() -> Settings:
    return Settings(
        telegram_bot_token="token",
        telegram_webhook_secret="secret",
        allowed_user_ids=frozenset({42}),
        database_url="postgresql+psycopg://user:pass@localhost/db",
    )


def catalog() -> ReportCatalog:
    report = Report("db_status", "DB status", "Status", "SELECT 1", (), 10)
    return ReportCatalog({report.report_id: report})


def test_whoami_is_available_before_authorization() -> None:
    sent = []

    async def send(chat_id, text):
        sent.append((chat_id, text))

    async def run_report(report, parameters):
        raise AssertionError("should not run")

    asyncio.run(
        handle_message(
            IncomingMessage(chat_id=9, user_id=999, text="/whoami"),
            settings(),
            catalog(),
            send,
            run_report,
        )
    )
    assert "999" in sent[0][1]


def test_unauthorized_user_cannot_run_report() -> None:
    sent = []

    async def send(chat_id, text):
        sent.append(text)

    async def run_report(report, parameters):
        raise AssertionError("should not run")

    asyncio.run(
        handle_message(
            IncomingMessage(chat_id=9, user_id=999, text="/run db_status"),
            settings(),
            catalog(),
            send,
            run_report,
        )
    )
    assert "Доступ запрещён" in sent[0]


def test_authorized_user_receives_report() -> None:
    sent = []

    async def send(chat_id, text):
        sent.append(text)

    async def run_report(report, parameters):
        return QueryResult(("value",), ((1,),), False)

    asyncio.run(
        handle_message(
            IncomingMessage(chat_id=9, user_id=42, text="/run db_status"),
            settings(),
            catalog(),
            send,
            run_report,
        )
    )
    assert "Формирую" in sent[0]
    assert "<pre>" in sent[1]
    assert "1" in sent[1]


def test_group_chat_cannot_receive_report() -> None:
    sent = []

    async def send(chat_id, text):
        sent.append(text)

    async def run_report(report, parameters):
        raise AssertionError("should not run")

    asyncio.run(
        handle_message(
            IncomingMessage(chat_id=-100, user_id=42, text="/run db_status", chat_type="group"),
            settings(),
            catalog(),
            send,
            run_report,
        )
    )
    assert "только в личном чате" in sent[0]
