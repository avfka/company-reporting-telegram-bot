"""Stateless agent selection; UUID and dates fit in Telegram's 64-byte callbacks."""
from __future__ import annotations

import re
from datetime import date
from uuid import UUID

from reporting_bot.agent_report import validate_period


def search_prompt(date_from: date, date_to: date, notice: str = ""):
    validate_period(date_from, date_to)
    return (
        f"{notice}\n" if notice else ""
    ) + (
        "Введите имя или часть названия агента в ответ на это сообщение (минимум 2 символа).\n"
        "ГТО исключено. Период — дата создания проекта.\n"
        f"AGENT_PERIOD:{date_from.isoformat()}:{date_to.isoformat()}"
    ), {"force_reply": True, "selective": True, "input_field_placeholder": "Имя или название агента"}


async def handle_agent_reply(message, send_message, search_agents) -> bool:
    if message.text.startswith("/"):
        return False
    match = re.search(r"(?:^|\n)AGENT_PERIOD:(\d{4}-\d{2}-\d{2}):(\d{4}-\d{2}-\d{2})$", message.reply_to_text or "")
    if not match:
        return False
    if search_agents is None:
        await send_message(message.chat_id, "Отчёт по агенту временно недоступен.")
        return True
    try:
        start, end = (date.fromisoformat(v) for v in match.groups())
        validate_period(start, end)
        options = await search_agents(message.text)
        if not options:
            prompt, markup = search_prompt(start, end, "Агент не найден. Попробуйте другую часть названия.")
        elif len(options) > 20:
            prompt, markup = search_prompt(start, end, "Найдено больше 20 агентов. Уточните название.")
        else:
            prompt = f"Выберите агента. Проекты созданы {start:%d.%m.%Y}–{end:%d.%m.%Y}."
            markup = {"inline_keyboard": [[{
                "text": f"{option.name[:90]} · {UUID(option.partner_id).hex[-6:]}",
                "callback_data": f"ag:go:{start:%Y%m%d}:{end:%Y%m%d}:{UUID(option.partner_id).hex}",
            }] for option in options]}
        await send_message(message.chat_id, prompt, markup)
    except ValueError as exc:
        await send_message(message.chat_id, str(exc) + " Повторите /report_agent.")
    return True


async def handle_agent_selection(callback, send_message, send_agent_report) -> bool:
    if not callback.data.startswith("ag:"):
        return False
    try:
        parts = callback.data.split(":")
        if len(parts) != 5 or parts[:2] != ["ag", "go"]:
            raise ValueError("Некорректная кнопка выбора агента.")
        start, end = (date.fromisoformat(v) for v in parts[2:4])
        validate_period(start, end)
        partner_id = str(UUID(parts[4]))
        if send_agent_report is None:
            await send_message(callback.chat_id, "Отчёт по агенту временно недоступен.")
            return True
        await send_message(callback.chat_id, "Формирую Excel-отчёт по выбранному агенту…")
        await send_agent_report(callback.chat_id, partner_id, start, end)
    except ValueError as exc:
        await send_message(callback.chat_id, str(exc) + " Повторите /report_agent.")
    return True
