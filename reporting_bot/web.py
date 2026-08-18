from __future__ import annotations

import asyncio
import json
import logging
import secrets
from typing import Any, Mapping

from reporting_bot.config import Settings
from reporting_bot.database import ReportExecutor
from reporting_bot.reports import ReportCatalog
from reporting_bot.telegram import TelegramClient, handle_message, parse_message


logger = logging.getLogger(__name__)
MAX_REQUEST_BYTES = 1_000_000


class ReportingBotApp:
    def __init__(self, settings: Settings) -> None:
        self.settings = settings

    async def __call__(self, scope, receive, send) -> None:
        scope_type = scope.get("type")
        if scope_type == "lifespan":
            await self._lifespan(receive, send)
            return
        if scope_type != "http":
            return

        method = str(scope.get("method", "GET")).upper()
        path = str(scope.get("path", "/"))
        if method == "GET" and path == "/":
            await self._json(
                send,
                200,
                {"service": "company-reporting-telegram-bot", "status": "ok"},
            )
            return
        if method == "GET" and path == "/health":
            await self._json(send, 200, self._health())
            return
        if method == "POST" and path == "/telegram/webhook":
            await self._webhook(scope, receive, send)
            return
        await self._json(send, 404, {"detail": "Not found"})

    def _health(self) -> dict[str, Any]:
        return {
            "status": "ready" if self.settings.ready else "configuration_required",
            "telegram_configured": bool(self.settings.telegram_bot_token),
            "database_configured": bool(self.settings.database_url),
            "access_list_configured": bool(self.settings.allowed_user_ids),
            "webhook_secret_configured": bool(self.settings.telegram_webhook_secret),
        }

    async def _webhook(self, scope, receive, send) -> None:
        if not self.settings.ready:
            await self._json(
                send,
                503,
                {"detail": "Service configuration is incomplete"},
            )
            return

        headers = {key.lower(): value for key, value in scope.get("headers", [])}
        supplied_secret = headers.get(
            b"x-telegram-bot-api-secret-token", b""
        ).decode("utf-8")
        if not supplied_secret or not secrets.compare_digest(
            supplied_secret,
            self.settings.telegram_webhook_secret,
        ):
            await self._json(send, 401, {"detail": "Invalid webhook secret"})
            return

        try:
            raw_body = await self._body(receive)
            update = json.loads(raw_body)
            if not isinstance(update, Mapping):
                raise ValueError("Webhook body must be an object")
        except (UnicodeDecodeError, json.JSONDecodeError, ValueError):
            await self._json(send, 400, {"detail": "Invalid JSON body"})
            return

        message = parse_message(update)
        if message is None:
            await self._json(send, 200, {"ok": True})
            return

        try:
            catalog = ReportCatalog.from_file(
                self.settings.reports_path,
                self.settings.default_max_rows,
            )
            telegram = TelegramClient(self.settings.telegram_bot_token)
            executor = ReportExecutor(self.settings)

            async def run_report(report, parameters):
                return await asyncio.to_thread(executor.run, report, parameters)

            await handle_message(
                message,
                self.settings,
                catalog,
                telegram.send_message,
                run_report,
            )
        except Exception:
            logger.exception("Failed to process Telegram update")
            try:
                telegram = TelegramClient(self.settings.telegram_bot_token)
                await telegram.send_message(
                    message.chat_id,
                    "Не удалось сформировать отчёт. Обратитесь к администратору.",
                )
            except Exception:
                logger.exception("Failed to send Telegram error message")
        await self._json(send, 200, {"ok": True})

    @staticmethod
    async def _body(receive) -> str:
        body = bytearray()
        more_body = True
        while more_body:
            event = await receive()
            if event.get("type") == "http.disconnect":
                raise ValueError("Client disconnected")
            if event.get("type") != "http.request":
                continue
            body.extend(event.get("body", b""))
            if len(body) > MAX_REQUEST_BYTES:
                raise ValueError("Request body too large")
            more_body = bool(event.get("more_body", False))
        return body.decode("utf-8")

    @staticmethod
    async def _json(send, status: int, payload: Mapping[str, Any]) -> None:
        body = json.dumps(
            payload,
            ensure_ascii=False,
            separators=(",", ":"),
        ).encode("utf-8")
        await send(
            {
                "type": "http.response.start",
                "status": status,
                "headers": [
                    (b"content-type", b"application/json; charset=utf-8"),
                    (b"content-length", str(len(body)).encode("ascii")),
                    (b"cache-control", b"no-store"),
                ],
            }
        )
        await send({"type": "http.response.body", "body": body})

    @staticmethod
    async def _lifespan(receive, send) -> None:
        while True:
            event = await receive()
            if event["type"] == "lifespan.startup":
                await send({"type": "lifespan.startup.complete"})
            elif event["type"] == "lifespan.shutdown":
                await send({"type": "lifespan.shutdown.complete"})
                return


def create_app(settings: Settings | None = None) -> ReportingBotApp:
    return ReportingBotApp(settings or Settings.from_env())
