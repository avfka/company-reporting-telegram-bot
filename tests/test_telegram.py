import asyncio

from reporting_bot.config import Settings
from reporting_bot.database import QueryResult
from reporting_bot.reports import Report, ReportCatalog
from reporting_bot.telegram import (
    IncomingCallback,
    IncomingMessage,
    handle_callback,
    handle_message,
    parse_callback,
)


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


def test_reports_sks_accepts_direct_date_range() -> None:
    sent = []
    generated = []

    async def send(chat_id, text, reply_markup=None):
        sent.append((chat_id, text, reply_markup))

    async def run_report(report, parameters):
        raise AssertionError("should not run")

    async def send_sks(chat_id, date_from, date_to):
        generated.append((chat_id, date_from.isoformat(), date_to.isoformat()))

    asyncio.run(
        handle_message(
            IncomingMessage(chat_id=9, user_id=42, text="/reports_sks 2026-07-01 2026-07-31"),
            settings(),
            catalog(),
            send,
            run_report,
            send_sks,
        )
    )
    assert "Excel" in sent[0][1]
    assert generated == [(9, "2026-07-01", "2026-07-31")]


def test_reports_sks_shows_date_buttons() -> None:
    sent = []

    async def send(chat_id, text, reply_markup=None):
        sent.append((text, reply_markup))

    async def run_report(report, parameters):
        raise AssertionError("should not run")

    async def send_sks(chat_id, date_from, date_to):
        raise AssertionError("should wait for a date selection")

    asyncio.run(
        handle_message(
            IncomingMessage(chat_id=9, user_id=42, text="/reports_sks"),
            settings(),
            catalog(),
            send,
            run_report,
            send_sks,
        )
    )
    assert sent[0][1]["inline_keyboard"]


def test_reports_dota_accepts_direct_date_range() -> None:
    sent = []
    generated = []

    async def send(chat_id, text, reply_markup=None):
        sent.append((chat_id, text, reply_markup))

    async def run_report(report, parameters):
        raise AssertionError("should not run")

    async def send_dota(chat_id, date_from, date_to):
        generated.append((chat_id, date_from.isoformat(), date_to.isoformat()))

    asyncio.run(
        handle_message(
            IncomingMessage(chat_id=9, user_id=42, text="/reports_dota 2026-08-01 2026-08-31"),
            settings(),
            catalog(),
            send,
            run_report,
            None,
            send_dota,
        )
    )
    assert "ДОТ" in sent[0][1]
    assert generated == [(9, "2026-08-01", "2026-08-31")]


def test_reports_dota_shows_date_buttons() -> None:
    sent = []

    async def send(chat_id, text, reply_markup=None):
        sent.append((text, reply_markup))

    async def run_report(report, parameters):
        raise AssertionError("should not run")

    async def send_dota(chat_id, date_from, date_to):
        raise AssertionError("should wait for a date selection")

    asyncio.run(
        handle_message(
            IncomingMessage(chat_id=9, user_id=42, text="/reports_dota"),
            settings(),
            catalog(),
            send,
            run_report,
            None,
            send_dota,
        )
    )
    assert sent[0][1]["inline_keyboard"]
    assert sent[0][1]["inline_keyboard"][0][0]["callback_data"].startswith("dota:")


def test_sks_callback_runs_selected_period() -> None:
    generated = []
    answered = []

    async def send(chat_id, text, reply_markup=None):
        pass

    async def answer(callback_query_id):
        answered.append(callback_query_id)

    async def send_sks(chat_id, date_from, date_to):
        generated.append((chat_id, date_from.isoformat(), date_to.isoformat()))

    asyncio.run(
        handle_callback(
            IncomingCallback("callback-1", 9, 42, "sks:to:2026-07-01:2026-07-31"),
            settings(),
            send,
            answer,
            send_sks,
        )
    )
    assert answered == ["callback-1"]
    assert generated == [(9, "2026-07-01", "2026-07-31")]


def test_end_date_calendar_uses_wide_single_line_prompt() -> None:
    sent = []

    async def send(chat_id, text, reply_markup=None):
        sent.append((chat_id, text, reply_markup))

    async def answer(callback_query_id):
        pass

    async def send_sks(chat_id, date_from, date_to):
        raise AssertionError("should wait for an end date")

    asyncio.run(
        handle_callback(
            IncomingCallback("callback-wide", 9, 42, "sks:from:2026-08-07"),
            settings(),
            send,
            answer,
            send_sks,
        )
    )

    assert len(sent) == 1
    assert "\n" not in sent[0][1]
    assert sent[0][1] == (
        "Дата начала отчёта: <b>07.08.2026</b> · "
        "Выберите дату окончания отчётного периода:"
    )
    assert len(sent[0][2]["inline_keyboard"][1]) == 7


def test_end_date_calendar_keeps_wide_prompt_when_switching_month() -> None:
    sent = []

    async def send(chat_id, text, reply_markup=None):
        sent.append((text, reply_markup))

    async def answer(callback_query_id):
        pass

    async def send_sks(chat_id, date_from, date_to):
        raise AssertionError("should wait for an end date")

    asyncio.run(
        handle_callback(
            IncomingCallback("callback-month", 9, 42, "sks:month_to:2026-08-07:2026-09"),
            settings(),
            send,
            answer,
            send_sks,
        )
    )

    assert sent[0][0].startswith("Дата начала отчёта: <b>07.08.2026</b>")
    assert sent[0][1]["inline_keyboard"][0][1]["text"] == "Сентябрь 2026"


def test_dota_callback_runs_selected_period() -> None:
    generated = []
    answered = []

    async def send(chat_id, text, reply_markup=None):
        pass

    async def answer(callback_query_id):
        answered.append(callback_query_id)

    async def send_sks(chat_id, date_from, date_to):
        raise AssertionError("should not run SKS")

    async def send_dota(chat_id, date_from, date_to):
        generated.append((chat_id, date_from.isoformat(), date_to.isoformat()))

    asyncio.run(
        handle_callback(
            IncomingCallback("callback-2", 9, 42, "dota:to:2026-08-01:2026-08-31"),
            settings(),
            send,
            answer,
            send_sks,
            send_dota,
        )
    )
    assert answered == ["callback-2"]
    assert generated == [(9, "2026-08-01", "2026-08-31")]


def test_parse_callback_reads_chat_and_sender() -> None:
    parsed = parse_callback(
        {
            "callback_query": {
                "id": "cb",
                "from": {"id": 42},
                "message": {"chat": {"id": 9, "type": "private"}},
                "data": "sks:noop",
            }
        }
    )
    assert parsed is not None
    assert parsed.chat_id == 9
    assert parsed.user_id == 42
