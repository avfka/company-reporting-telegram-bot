from __future__ import annotations

import calendar
import html
import logging
import re
import shlex
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Any, Awaitable, Callable, Mapping
from zoneinfo import ZoneInfo

import httpx

from reporting_bot.config import Settings
from reporting_bot.database import QueryResult
from reporting_bot.ks_reports import KsFilterOptions, KsFilters
from reporting_bot.reports import Report, ReportCatalog


logger = logging.getLogger(__name__)


KS_COMMANDS = {
    "/reports_ks_all": ("ksa", "all", "КС — Общий отчёт"),
    "/reports_ks_plan": ("ksp", "plan", "КС — План-факт"),
    "/reports_ks_managers": ("ksm", "managers", "КС — Менеджеры"),
    "/reports_ks_funnel": ("ksf", "funnel", "КС — Воронка"),
    "/reports_ks_projects": ("ksj", "projects", "КС — Проекты и оплаты"),
}
KS_PREFIXES = {prefix: (kind, title) for prefix, kind, title in KS_COMMANDS.values()}


class TelegramClient:
    def __init__(self, token: str) -> None:
        self._base_url = f"https://api.telegram.org/bot{token}"

    async def send_message(
        self,
        chat_id: int,
        text: str,
        reply_markup: Mapping[str, Any] | None = None,
    ) -> None:
        payload: dict[str, Any] = {
            "chat_id": chat_id,
            "text": text,
            "parse_mode": "HTML",
        }
        if reply_markup is not None:
            payload["reply_markup"] = reply_markup
        async with httpx.AsyncClient(timeout=10) as client:
            response = await client.post(
                f"{self._base_url}/sendMessage",
                json=payload,
            )
            response.raise_for_status()

    async def send_document(
        self,
        chat_id: int,
        content: bytes,
        filename: str,
        caption: str,
    ) -> None:
        async with httpx.AsyncClient(timeout=30) as client:
            response = await client.post(
                f"{self._base_url}/sendDocument",
                data={"chat_id": str(chat_id), "caption": caption, "parse_mode": "HTML"},
                files={
                    "document": (
                        filename,
                        content,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                },
            )
            response.raise_for_status()

    async def send_photo(
        self,
        chat_id: int,
        content: bytes,
        filename: str,
        caption: str,
    ) -> None:
        async with httpx.AsyncClient(timeout=30) as client:
            response = await client.post(
                f"{self._base_url}/sendPhoto",
                data={"chat_id": str(chat_id), "caption": caption, "parse_mode": "HTML"},
                files={"photo": (filename, content, "image/png")},
            )
            response.raise_for_status()

    async def answer_callback_query(self, callback_query_id: str) -> None:
        async with httpx.AsyncClient(timeout=10) as client:
            response = await client.post(
                f"{self._base_url}/answerCallbackQuery",
                json={"callback_query_id": callback_query_id},
            )
            response.raise_for_status()


@dataclass(frozen=True)
class IncomingMessage:
    chat_id: int
    user_id: int
    text: str
    chat_type: str = "private"
    reply_to_text: str | None = None


@dataclass(frozen=True)
class IncomingCallback:
    callback_query_id: str
    chat_id: int
    user_id: int
    data: str
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
            reply_to_text=(
                str(message["reply_to_message"].get("text", ""))
                if isinstance(message.get("reply_to_message"), Mapping)
                else None
            ),
        )
    except (KeyError, TypeError, ValueError):
        return None


def parse_callback(update: Mapping[str, Any]) -> IncomingCallback | None:
    callback = update.get("callback_query")
    if not isinstance(callback, Mapping):
        return None
    sender = callback.get("from")
    message = callback.get("message")
    data = callback.get("data")
    if not isinstance(sender, Mapping) or not isinstance(message, Mapping) or not isinstance(data, str):
        return None
    chat = message.get("chat")
    if not isinstance(chat, Mapping):
        return None
    try:
        return IncomingCallback(
            callback_query_id=str(callback["id"]),
            chat_id=int(chat["id"]),
            user_id=int(sender["id"]),
            data=data,
            chat_type=str(chat.get("type", "")),
        )
    except (KeyError, TypeError, ValueError):
        return None


def _help_text() -> str:
    return (
        "<b>Бот отчётности</b>\n\n"
        "/reports — список доступных отчётов\n"
        "/reports_sks — подробный Excel-отчёт СКС с выбором дат\n"
        "/reports_dota — инфографика и Excel-отчёт ДОТ по запускам и выпускам\n"
        "/reports_ks — меню аналитики КС: общий отчёт и отдельные разделы\n"
        "/run &lt;отчёт&gt; [параметр=значение] — сформировать отчёт\n"
        "/whoami — показать ваш Telegram ID\n"
        "/help — помощь"
    )


def _report_list(catalog: ReportCatalog) -> str:
    lines = [
        "<b>Доступные отчёты</b>",
        "\n<code>reports_sks</code> — Подробный отчёт СКС",
        "Excel со сводкой, показателями по специалистам и проектной детализацией. Команда: /reports_sks",
        "\n<code>reports_dota</code> — Отчёт ДОТ по запускам и выпускам",
        "Инфографика и Excel со сводкой, динамикой и детализацией проектов. Команда: /reports_dota",
        "\n<code>reports_ks</code> — Аналитика КС",
        "Общий отчёт и четыре отдельных раздела с фильтрами, Excel и графиком. Команда: /reports_ks",
    ]
    for report in catalog.all():
        suffix = ""
        if report.parameters:
            suffix = " · параметры: " + ", ".join(report.parameters)
        lines.append(f"\n<code>{html.escape(report.report_id)}</code> — {html.escape(report.title)}{html.escape(suffix)}")
        lines.append(html.escape(report.description))
    return "\n".join(lines)


def _month_shift(value: date, months: int) -> date:
    month_index = value.year * 12 + value.month - 1 + months
    return date(month_index // 12, month_index % 12 + 1, 1)


def _month_title(value: date) -> str:
    names = (
        "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
        "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь",
    )
    return f"{names[value.month - 1]} {value.year}"


def _end_date_prompt(date_from: date) -> str:
    # Telegram sizes an inline keyboard to the message bubble. Keep this prompt
    # on one sufficiently long line so the second calendar does not collapse.
    return (
        f"Дата начала отчёта: <b>{date_from:%d.%m.%Y}</b> · "
        "Выберите дату окончания отчётного периода:"
    )


def _calendar_markup(
    report_prefix: str,
    mode: str,
    month: date,
    start: date | None = None,
) -> dict[str, Any]:
    if report_prefix not in {"sks", "dota", *KS_PREFIXES}:
        raise ValueError("Unknown report prefix")
    if mode not in {"from", "to"}:
        raise ValueError("Unknown calendar mode")
    previous = _month_shift(month, -1)
    following = _month_shift(month, 1)
    if mode == "from":
        previous_data = f"{report_prefix}:month_from:{previous:%Y-%m}"
        following_data = f"{report_prefix}:month_from:{following:%Y-%m}"
    else:
        if start is None:
            raise ValueError("End-date calendar needs a start date")
        previous_data = f"{report_prefix}:month_to:{start.isoformat()}:{previous:%Y-%m}"
        following_data = f"{report_prefix}:month_to:{start.isoformat()}:{following:%Y-%m}"

    keyboard: list[list[dict[str, str]]] = [
        [
            {"text": "‹", "callback_data": previous_data},
            {"text": _month_title(month), "callback_data": f"{report_prefix}:noop"},
            {"text": "›", "callback_data": following_data},
        ],
        [{"text": value, "callback_data": f"{report_prefix}:noop"} for value in ("Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс")],
    ]
    for week in calendar.Calendar(firstweekday=0).monthdatescalendar(month.year, month.month):
        row = []
        for day in week:
            if day.month != month.month:
                row.append({"text": "·", "callback_data": f"{report_prefix}:noop"})
            elif mode == "from":
                row.append({"text": str(day.day), "callback_data": f"{report_prefix}:from:{day.isoformat()}"})
            elif start is not None and day < start:
                row.append({"text": "·", "callback_data": f"{report_prefix}:noop"})
            else:
                row.append({"text": str(day.day), "callback_data": f"{report_prefix}:to:{start.isoformat()}:{day.isoformat()}"})
        keyboard.append(row)
    return {"inline_keyboard": keyboard}


def _report_entry_markup(report_prefix: str, today: date) -> dict[str, Any]:
    current_month = today.replace(day=1)
    previous_month = _month_shift(current_month, -1)
    previous_month_end = current_month - timedelta(days=1)
    return {
        "inline_keyboard": [
            [{"text": "Последние 7 дней", "callback_data": f"{report_prefix}:preset:{(today - timedelta(days=6)).isoformat()}:{today.isoformat()}"}],
            [{"text": "Текущий месяц", "callback_data": f"{report_prefix}:preset:{current_month.isoformat()}:{today.isoformat()}"}],
            [{"text": "Прошлый месяц", "callback_data": f"{report_prefix}:preset:{previous_month.isoformat()}:{previous_month_end.isoformat()}"}],
            [{"text": "Выбрать даты", "callback_data": f"{report_prefix}:month_from:{current_month:%Y-%m}"}],
        ]
    }


def _ks_menu_markup() -> dict[str, Any]:
    return {
        "inline_keyboard": [
            [{"text": title, "callback_data": f"ks:open:{prefix}"}]
            for prefix, (_, title) in KS_PREFIXES.items()
        ]
    }


def _compact_date(value: date) -> str:
    return value.strftime("%Y%m%d")


def _callback_date(value: str) -> date:
    if "-" in value:
        return date.fromisoformat(value)
    return datetime.strptime(value, "%Y%m%d").date()


def _button_rows(buttons: list[dict[str, str]], columns: int = 2) -> list[list[dict[str, str]]]:
    return [buttons[index:index + columns] for index in range(0, len(buttons), columns)]


def _ks_department_markup(
    prefix: str,
    date_from: date,
    date_to: date,
    options: KsFilterOptions,
) -> dict[str, Any]:
    start = _compact_date(date_from)
    end = _compact_date(date_to)
    buttons = [{"text": "Все отделы", "callback_data": f"{prefix}:dept:{start}:{end}:-"}]
    buttons.extend(
        {"text": option.label, "callback_data": f"{prefix}:dept:{start}:{end}:{option.token}"}
        for option in options.departments
    )
    return {"inline_keyboard": _button_rows(buttons)}


def _ks_manager_markup(
    prefix: str,
    date_from: date,
    date_to: date,
    department_token: str,
    service: str,
    comparison: str,
    options: KsFilterOptions,
) -> dict[str, Any]:
    start = _compact_date(date_from)
    end = _compact_date(date_to)
    common = f"{prefix}:go:{start}:{end}:{department_token}:{service}:{comparison}"
    buttons = [
        {"text": "Все менеджеры", "callback_data": common + ":-"},
        {
            "text": "✍️ Ввести несколько менеджеров",
            "callback_data": f"{prefix}:mm:{start}:{end}:{department_token}:{service}:{comparison}",
        },
    ]
    if department_token != "-":
        buttons.extend(
            {"text": option.label, "callback_data": common + f":{option.token}"}
            for option in options.managers
        )
    return {"inline_keyboard": _button_rows(buttons)}


def _ks_product_markup(
    prefix: str,
    date_from: date,
    date_to: date,
    department_token: str,
    options: KsFilterOptions,
) -> dict[str, Any]:
    start = _compact_date(date_from)
    end = _compact_date(date_to)
    buttons = [{"text": "Все продукты", "callback_data": f"{prefix}:pr:{start}:{end}:{department_token}:-"}]
    buttons.extend(
        {"text": option.label, "callback_data": f"{prefix}:pr:{start}:{end}:{department_token}:{option.token}"}
        for option in options.products
    )
    return {"inline_keyboard": _button_rows(buttons)}


def _ks_comparison_markup(
    prefix: str,
    date_from: date,
    date_to: date,
    department_token: str,
    service: str,
) -> dict[str, Any]:
    common = (
        f"{prefix}:co:{_compact_date(date_from)}:{_compact_date(date_to)}:"
        f"{department_token}:{service}"
    )
    return {
        "inline_keyboard": [
            [{"text": "Предыдущий аналогичный период", "callback_data": common + ":p"}],
            [{"text": "Тот же период прошлого года", "callback_data": common + ":y"}],
            [{"text": "Без сравнения", "callback_data": common + ":n"}],
        ]
    }


def _ks_manual_manager_prompt(
    prefix: str,
    date_from: date,
    date_to: date,
    department_token: str,
    service: str,
    comparison: str,
    error: str | None = None,
) -> tuple[str, dict[str, Any]]:
    marker = (
        f"KSM|{prefix}|{_compact_date(date_from)}|{_compact_date(date_to)}|"
        f"{department_token}|{service}|{comparison}"
    )
    heading = f"{html.escape(error)}\n\n" if error else ""
    text_value = (
        heading
        + "Введите ФИО менеджеров через запятую или каждое с новой строки.\n"
        + "Например: <code>Иванова Елена, Максимович Анастасия</code>\n\n"
        + f"Код выбора: <code>{marker}</code>"
    )
    return text_value, {
        "force_reply": True,
        "selective": True,
        "input_field_placeholder": "Фамилия Имя, Фамилия Имя",
    }


def _manual_manager_context(reply_to_text: str | None) -> tuple[str, date, date, str, str, str] | None:
    if not reply_to_text:
        return None
    match = re.search(
        r"KSM\|(ksa|ksp|ksm|ksf|ksj)\|(\d{8})\|(\d{8})\|([a-f0-9-]+)\|([a-f0-9-]+)\|([pyn])",
        reply_to_text,
    )
    if match is None:
        return None
    prefix, start, end, department_token, service, comparison = match.groups()
    return prefix, _callback_date(start), _callback_date(end), department_token, service, comparison


def _resolve_manager_tokens(raw_names: str, options: KsFilterOptions) -> str:
    names = [
        " ".join(value.split())
        for value in re.split(r"[,;\n]+", raw_names)
        if value.strip()
    ]
    if not names:
        raise ValueError("Не указано ни одного менеджера.")
    if len(names) > 10:
        raise ValueError("За один раз можно выбрать не более 10 менеджеров.")

    def normalized(value: str) -> str:
        return " ".join(value.casefold().replace("ё", "е").split())

    by_name = {normalized(option.label): option.token for option in options.managers}
    tokens: list[str] = []
    missing: list[str] = []
    for name in names:
        token = by_name.get(normalized(name))
        if token is None:
            missing.append(name)
        elif token not in tokens:
            tokens.append(token)
    if missing:
        raise ValueError("Не найдены менеджеры: " + ", ".join(missing) + ". Проверьте ФИО.")
    return ",".join(tokens)


def _parse_report_dates(text_value: str, command_name: str) -> tuple[date, date] | None:
    parts = shlex.split(text_value)
    if len(parts) == 1:
        return None
    if len(parts) != 3:
        raise ValueError(f"Формат: /{command_name} ГГГГ-ММ-ДД ГГГГ-ММ-ДД")
    try:
        date_from = date.fromisoformat(parts[1])
        date_to = date.fromisoformat(parts[2])
    except ValueError as exc:
        raise ValueError("Даты нужны в формате ГГГГ-ММ-ДД, например 2026-07-01.") from exc
    if date_to < date_from:
        raise ValueError("Дата окончания не может быть раньше даты начала.")
    if (date_to - date_from).days > 366:
        raise ValueError("Максимальный период отчёта — 366 дней.")
    return date_from, date_to


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
    send_message: Callable[..., Awaitable[None]],
    run_report: Callable[[Report, Mapping[str, str]], Awaitable[QueryResult]],
    send_sks_report: Callable[[int, date, date], Awaitable[None]] | None = None,
    send_dota_report: Callable[[int, date, date], Awaitable[None]] | None = None,
    send_ks_report: Callable[[int, str, date, date, KsFilters, str], Awaitable[None]] | None = None,
    load_ks_filters: Callable[[str], Awaitable[KsFilterOptions]] | None = None,
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
    manual_context = _manual_manager_context(message.reply_to_text)
    if manual_context is not None:
        if send_ks_report is None or load_ks_filters is None:
            await send_message(message.chat_id, "Отчёты КС временно недоступны.")
            return
        prefix, date_from, date_to, department_token, service, comparison_code = manual_context
        report_kind, report_title = KS_PREFIXES[prefix]
        comparison = {"p": "previous", "y": "year", "n": "none"}[comparison_code]
        options = await load_ks_filters(department_token)
        try:
            manager_token = _resolve_manager_tokens(message.text, options)
        except ValueError as exc:
            prompt, reply_markup = _ks_manual_manager_prompt(
                prefix,
                date_from,
                date_to,
                department_token,
                service,
                comparison_code,
                str(exc),
            )
            await send_message(message.chat_id, prompt, reply_markup)
            return
        await send_message(message.chat_id, f"Формирую график и Excel-отчёт {report_title}…")
        await send_ks_report(
            message.chat_id,
            report_kind,
            date_from,
            date_to,
            KsFilters(department_token, manager_token, service),
            comparison,
        )
        return
    if command in ("/start", "/help"):
        await send_message(message.chat_id, _help_text())
        return
    if command == "/reports":
        await send_message(message.chat_id, _report_list(catalog))
        return
    if command in ("/reports_ks", "reports_ks"):
        await send_message(
            message.chat_id,
            "<b>Аналитика КС</b>\n\nВыберите отчёт:",
            _ks_menu_markup(),
        )
        return
    normalized_ks_command = command if command.startswith("/") else "/" + command
    if normalized_ks_command in KS_COMMANDS:
        prefix, report_kind, report_title = KS_COMMANDS[normalized_ks_command]
        if send_ks_report is None or load_ks_filters is None:
            await send_message(message.chat_id, "Отчёты КС временно недоступны.")
            return
        try:
            selected = _parse_report_dates(message.text, normalized_ks_command.removeprefix("/"))
        except ValueError as exc:
            await send_message(message.chat_id, str(exc))
            return
        if selected is None:
            today = datetime.now(ZoneInfo("Europe/Moscow")).date()
            await send_message(
                message.chat_id,
                f"<b>{report_title}</b>\n\nВыберите период или отправьте даты после команды.",
                _report_entry_markup(prefix, today),
            )
            return
        options = await load_ks_filters("-")
        await send_message(
            message.chat_id,
            f"<b>{report_title}</b>\n\nВыберите отдел:",
            _ks_department_markup(prefix, *selected, options),
        )
        return
    if command in ("/reports_sks", "reports_sks"):
        if send_sks_report is None:
            await send_message(message.chat_id, "Отчёт СКС временно недоступен.")
            return
        try:
            selected = _parse_report_dates(message.text, "reports_sks")
        except ValueError as exc:
            await send_message(message.chat_id, str(exc))
            return
        if selected is None:
            today = datetime.now(ZoneInfo("Europe/Moscow")).date()
            await send_message(
                message.chat_id,
                "<b>Отчёт СКС</b>\n\nВыберите период или отправьте команду:\n<code>/reports_sks 2026-07-01 2026-07-31</code>",
                _report_entry_markup("sks", today),
            )
            return
        await send_message(message.chat_id, "Формирую Excel-отчёт СКС…")
        await send_sks_report(message.chat_id, *selected)
        return
    if command in ("/reports_dota", "reports_dota"):
        if send_dota_report is None:
            await send_message(message.chat_id, "Отчёт ДОТ временно недоступен.")
            return
        try:
            selected = _parse_report_dates(message.text, "reports_dota")
        except ValueError as exc:
            await send_message(message.chat_id, str(exc))
            return
        if selected is None:
            today = datetime.now(ZoneInfo("Europe/Moscow")).date()
            await send_message(
                message.chat_id,
                "<b>Отчёт ДОТ</b>\n\nВыберите период или отправьте команду:\n<code>/reports_dota 2026-08-01 2026-08-31</code>",
                _report_entry_markup("dota", today),
            )
            return
        await send_message(message.chat_id, "Формирую инфографику и Excel-отчёт ДОТ…")
        await send_dota_report(message.chat_id, *selected)
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


async def handle_callback(
    callback: IncomingCallback,
    settings: Settings,
    send_message: Callable[..., Awaitable[None]],
    answer_callback: Callable[[str], Awaitable[None]],
    send_sks_report: Callable[[int, date, date], Awaitable[None]],
    send_dota_report: Callable[[int, date, date], Awaitable[None]] | None = None,
    send_ks_report: Callable[[int, str, date, date, KsFilters, str], Awaitable[None]] | None = None,
    load_ks_filters: Callable[[str], Awaitable[KsFilterOptions]] | None = None,
) -> None:
    await answer_callback(callback.callback_query_id)
    if callback.chat_type != "private":
        await send_message(callback.chat_id, "Из соображений безопасности отчёты доступны только в личном чате с ботом.")
        return
    if callback.user_id not in settings.allowed_user_ids:
        await send_message(callback.chat_id, "Доступ запрещён. Используйте /whoami и передайте ID администратору.")
        return
    if callback.data.startswith("ks:open:"):
        prefix = callback.data.split(":", 2)[2]
        if prefix not in KS_PREFIXES:
            return
        _, report_title = KS_PREFIXES[prefix]
        today = datetime.now(ZoneInfo("Europe/Moscow")).date()
        await send_message(
            callback.chat_id,
            f"<b>{report_title}</b>\n\nВыберите период:",
            _report_entry_markup(prefix, today),
        )
        return
    if not callback.data.startswith(("sks:", "dota:", *(prefix + ":" for prefix in KS_PREFIXES))):
        return
    parts = callback.data.split(":")
    report_prefix = parts[0]
    if report_prefix == "sks":
        report_title = "СКС"
        report_command = "/reports_sks"
        send_report = send_sks_report
    elif report_prefix == "dota":
        report_title = "ДОТ"
        report_command = "/reports_dota"
        send_report = send_dota_report
    else:
        report_kind, report_title = KS_PREFIXES[report_prefix]
        report_command = next(command for command, values in KS_COMMANDS.items() if values[0] == report_prefix)
        send_report = None
        if send_ks_report is None or load_ks_filters is None:
            await send_message(callback.chat_id, f"Отчёт {report_title} временно недоступен.")
            return
    if report_prefix not in KS_PREFIXES and send_report is None:
        await send_message(callback.chat_id, f"Отчёт {report_title} временно недоступен.")
        return
    try:
        action = parts[1]
        if action == "noop":
            return
        if action == "preset" and len(parts) == 4:
            date_from = date.fromisoformat(parts[2])
            date_to = date.fromisoformat(parts[3])
            if report_prefix in KS_PREFIXES:
                options = await load_ks_filters("-")
                await send_message(
                    callback.chat_id,
                    f"<b>{report_title}</b>\n\nВыберите отдел:",
                    _ks_department_markup(report_prefix, date_from, date_to, options),
                )
                return
            message_text = (
                "Формирую инфографику и Excel-отчёт ДОТ…"
                if report_prefix == "dota"
                else f"Формирую Excel-отчёт {report_title}…"
            )
            await send_message(callback.chat_id, message_text)
            await send_report(callback.chat_id, date_from, date_to)
            return
        if action == "month_from" and len(parts) == 3:
            month = date.fromisoformat(parts[2] + "-01")
            await send_message(callback.chat_id, "Выберите <b>дату начала</b>:", _calendar_markup(report_prefix, "from", month))
            return
        if action == "from" and len(parts) == 3:
            date_from = date.fromisoformat(parts[2])
            await send_message(
                callback.chat_id,
                _end_date_prompt(date_from),
                _calendar_markup(report_prefix, "to", date_from.replace(day=1), date_from),
            )
            return
        if action == "month_to" and len(parts) == 4:
            date_from = date.fromisoformat(parts[2])
            month = date.fromisoformat(parts[3] + "-01")
            await send_message(callback.chat_id, _end_date_prompt(date_from), _calendar_markup(report_prefix, "to", month, date_from))
            return
        if action == "to" and len(parts) == 4:
            date_from = date.fromisoformat(parts[2])
            date_to = date.fromisoformat(parts[3])
            if date_to < date_from:
                raise ValueError("Дата окончания раньше даты начала.")
            if (date_to - date_from).days > 366:
                raise ValueError("Максимальный период отчёта — 366 дней.")
            if report_prefix in KS_PREFIXES:
                options = await load_ks_filters("-")
                await send_message(
                    callback.chat_id,
                    f"<b>{report_title}</b>\n\nВыберите отдел:",
                    _ks_department_markup(report_prefix, date_from, date_to, options),
                )
                return
            message_text = (
                "Формирую инфографику и Excel-отчёт ДОТ…"
                if report_prefix == "dota"
                else f"Формирую Excel-отчёт {report_title}…"
            )
            await send_message(callback.chat_id, message_text)
            await send_report(callback.chat_id, date_from, date_to)
            return
        if report_prefix in KS_PREFIXES and action == "dept" and len(parts) == 5:
            date_from = _callback_date(parts[2])
            date_to = _callback_date(parts[3])
            department_token = parts[4]
            options = await load_ks_filters(department_token)
            await send_message(
                callback.chat_id,
                "Выберите продукт:",
                _ks_product_markup(report_prefix, date_from, date_to, department_token, options),
            )
            return
        if report_prefix in KS_PREFIXES and action == "pr" and len(parts) == 6:
            date_from = _callback_date(parts[2])
            date_to = _callback_date(parts[3])
            department_token, service = parts[4:6]
            await send_message(
                callback.chat_id,
                "Выберите период для сравнения:",
                _ks_comparison_markup(report_prefix, date_from, date_to, department_token, service),
            )
            return
        if report_prefix in KS_PREFIXES and action == "co" and len(parts) == 7:
            date_from = _callback_date(parts[2])
            date_to = _callback_date(parts[3])
            department_token, service, comparison = parts[4:7]
            if comparison not in {"p", "y", "n"}:
                raise ValueError("Неизвестный период сравнения.")
            options = await load_ks_filters(department_token)
            await send_message(
                callback.chat_id,
                "Выберите одного менеджера или введите несколько вручную:",
                _ks_manager_markup(
                    report_prefix,
                    date_from,
                    date_to,
                    department_token,
                    service,
                    comparison,
                    options,
                ),
            )
            return
        if report_prefix in KS_PREFIXES and action == "mm" and len(parts) == 7:
            date_from = _callback_date(parts[2])
            date_to = _callback_date(parts[3])
            department_token, service, comparison = parts[4:7]
            prompt, reply_markup = _ks_manual_manager_prompt(
                report_prefix,
                date_from,
                date_to,
                department_token,
                service,
                comparison,
            )
            await send_message(callback.chat_id, prompt, reply_markup)
            return
        if report_prefix in KS_PREFIXES and action == "go" and len(parts) == 8:
            date_from = _callback_date(parts[2])
            date_to = _callback_date(parts[3])
            department_token, service, comparison_code, manager_token = parts[4:8]
            comparison = {"p": "previous", "y": "year", "n": "none"}.get(comparison_code)
            if comparison is None:
                raise ValueError("Неизвестный период сравнения.")
            await send_message(callback.chat_id, f"Формирую график и Excel-отчёт {report_title}…")
            await send_ks_report(
                callback.chat_id,
                report_kind,
                date_from,
                date_to,
                KsFilters(department_token, manager_token, service),
                comparison,
            )
            return
    except (IndexError, ValueError) as exc:
        await send_message(callback.chat_id, f"Не удалось применить выбор: {html.escape(str(exc))}")
        return
    await send_message(callback.chat_id, f"Не удалось распознать выбор. Отправьте {report_command} ещё раз.")
