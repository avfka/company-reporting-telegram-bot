from __future__ import annotations

import html
import logging
import shlex
from dataclasses import dataclass
from typing import Any, Awaitable, Callable, Mapping

import httpx

from reporting_bot.config import Settings
from reporting_bot.database import QueryResult
from reporting_bot.reports import Report, ReportCatalog


logger = logging.getLogger(__name__)


class TelegramClient:
    def __init__(self, token: str) -> None:
        self._base_url = f"https://api.telegram.org/bot{token}"

    async def send_message(self, chat_id: int, text: str) -> None:
        async with httpx.AsyncClient(timeout=10) as client:
            response = await client.post(
                f"{self._base_url}/sendMessage",
                json={"chat_id": chat_id, "text": text, "parse_mode": "HTML"},
            )
            response.raise_for_status()


@dataclass(frozen=True)
class IncomingMessage:
    chat_id: int
    user_id: int
    text: str
    chat_type: str = "private"


def parse_message(update: Mapping[str, Any]) -> IncomingMessage | None:
    message = update.get("message")
    if not isinstance(message, Mapping):
        return None
    sender = message.get("from")
    chat = message.get("chat")
    text_value = message.get("text")
    if not isinstance(sender, Mapping) or not isinstance(chat, Mapping) or not isinstance(text_value, str):
        return None
    try:
        return IncomingMessage(
            chat_id=int(chat["id"]),
            user_id=int(sender["id"]),
            text=text_value.strip(),
            chat_type=str(chat.get("type", "")),
        )
    except (KeyError, TypeError, ValueError):
        return None


def _help_text() -> str:
    return (
        "<b>Бот отчётности</b>\n\n"
        "/reports — список доступных отчётов\n"
        "/run &lt;отчёт&gt; [параметр=значение] — сформировать отчёт\n"
        "/whoami — показать ваш Telegram ID\n"
        "/help — помощь"
    )


def _report_list(catalog: ReportCatalog) -> str:
    lines = ["<b>Доступные отчёты</b>"]
    for report in catalog.all():
        suffix = ""
        if report.parameters:
            suffix = " · параметры: " + ", ".join(report.parameters)
        lines.append(f"\n<code>{html.escape(report.report_id)}</code> — {html.escape(report.title)}{html.escape(suffix)}")
        lines.append(html.escape(report.description))
    return "\n".join(lines)


def _parse_run(text_value: str, catalog: ReportCatalog) -> tuple[Report, dict[str, str]]:
    try:
        parts = shlex.split(text_value)
    except ValueError as exc:
        raise ValueError("Не удалось разобрать параметры. Проверьте кавычки.") from exc
    if len(parts) < 2:
        raise ValueError("Формат: /run &lt;отчёт&gt; [параметр=значение]")
    report = catalog.get(parts[1])
    if report is None:
        raise ValueError("Неизвестный отчёт. Используйте /reports.")

    parameters: dict[str, str] = {}
    for item in parts[2:]:
        if "=" not in item:
            raise ValueError(f"Параметр {html.escape(item)} должен иметь формат имя=значение.")
        key, value = item.split("=", 1)
        if key not in report.parameters:
            raise ValueError(f"Параметр {html.escape(key)} не разрешён для этого отчёта.")
        parameters[key] = value
    missing = [name for name in report.parameters if name not in parameters]
    if missing:
        raise ValueError("Не заданы параметры: " + ", ".join(html.escape(value) for value in missing))
    return report, parameters


def _format_result(report: Report, result: QueryResult) -> list[str]:
    heading = f"<b>{html.escape(report.title)}</b>"
    if not result.rows:
        return [heading + "\n\nНет данных."]

    widths = []
    for index, column in enumerate(result.columns):
        values = [str(row[index]) if row[index] is not None else "—" for row in result.rows]
        widths.append(min(32, max(len(str(column)), *(len(value) for value in values))))

    def render(values: tuple[Any, ...] | tuple[str, ...]) -> str:
        cells = []
        for index, value in enumerate(values):
            display = "—" if value is None else str(value)
            if len(display) > widths[index]:
                display = display[: max(1, widths[index] - 1)] + "…"
            cells.append(display.ljust(widths[index]))
        return " | ".join(cells)

    table_lines = [render(result.columns), "-+-".join("-" * width for width in widths)]
    table_lines.extend(render(row) for row in result.rows)
    if result.truncated:
        table_lines.append("… результат ограничен настройкой max_rows")

    messages: list[str] = []
    current = heading + "\n<pre>"
    for line in table_lines:
        escaped = html.escape(line)
        if len(current) + len(escaped) + len("\n</pre>") > 3900:
            messages.append(current + "</pre>")
            current = "<pre>"
        current += escaped + "\n"
    messages.append(current + "</pre>")
    return messages


async def handle_message(
    message: IncomingMessage,
    settings: Settings,
    catalog: ReportCatalog,
    send_message: Callable[[int, str], Awaitable[None]],
    run_report: Callable[[Report, Mapping[str, str]], Awaitable[QueryResult]],
) -> None:
    command = message.text.split(maxsplit=1)[0].split("@", 1)[0].lower()

    if message.chat_type != "private":
        await send_message(message.chat_id, "Из соображений безопасности отчёты доступны только в личном чате с ботом.")
        return
    if command == "/whoami":
        await send_message(message.chat_id, f"Ваш Telegram ID: <code>{message.user_id}</code>")
        return
    if message.user_id not in settings.allowed_user_ids:
        await send_message(message.chat_id, "Доступ запрещён. Передайте администратору ID из команды /whoami.")
        return
    if command in ("/start", "/help"):
        await send_message(message.chat_id, _help_text())
        return
    if command == "/reports":
        await send_message(message.chat_id, _report_list(catalog))
        return
    if command == "/run":
        try:
            report, parameters = _parse_run(message.text, catalog)
        except ValueError as exc:
            await send_message(message.chat_id, str(exc))
            return
        logger.info("Running report", extra={"report_id": report.report_id, "telegram_user_id": message.user_id})
        await send_message(message.chat_id, "Формирую отчёт…")
        result = await run_report(report, parameters)
        for part in _format_result(report, result):
            await send_message(message.chat_id, part)
        return
    await send_message(message.chat_id, _help_text())
