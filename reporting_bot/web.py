from __future__ import annotations

import asyncio
import json
import logging
import secrets
from typing import Any, Mapping

from reporting_bot.config import Settings
from reporting_bot.crm_bridge import CrmBridgeRepository
from reporting_bot.database import ReportExecutor
from reporting_bot.dota_report import DotaReportService
from reporting_bot.ks_reports import KsReportService
from reporting_bot.reports import ReportCatalog
from reporting_bot.sks_report import SksReportService
from reporting_bot.telegram import (
    TelegramClient,
    handle_callback,
    handle_message,
    parse_callback,
    parse_message,
)


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
        if method == "GET" and path == "/internal/crm/schema":
            await self._crm_schema(scope, send)
            return
        await self._json(send, 404, {"detail": "Not found"})

    def _health(self) -> dict[str, Any]:
        return {
            "status": "ready" if self.settings.ready else "configuration_required",
            "telegram_configured": bool(self.settings.telegram_bot_token),
            "database_configured": bool(self.settings.database_url),
            "access_list_configured": bool(self.settings.allowed_user_ids),
            "webhook_secret_configured": bool(self.settings.telegram_webhook_secret),
            "crm_bridge_configured": bool(self.settings.crm_bridge_token),
        }

    async def _crm_schema(self, scope, send) -> None:
        if not self._valid_bridge_token(scope):
            await self._json(send, 401, {"detail": "Invalid bridge token"})
            return
        try:
            payload = await asyncio.to_thread(
                CrmBridgeRepository(self.settings).schema
            )
        except Exception:
            logger.exception("Failed to inspect CRM schema")
            await self._json(send, 503, {"detail": "CRM is unavailable"})
            return
        await self._json(send, 200, payload)

    def _valid_bridge_token(self, scope) -> bool:
        if not self.settings.crm_bridge_token:
            return False
        headers = {key.lower(): value for key, value in scope.get("headers", [])}
        supplied = headers.get(b"x-crm-bridge-token", b"").decode("utf-8")
        return bool(supplied) and secrets.compare_digest(
            supplied,
            self.settings.crm_bridge_token,
        )

    async def _webhook(self, scope, receive, send) -> None:
        if not self.settings.webhook_ready:
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
        callback = parse_callback(update)
        if message is None and callback is None:
            await self._json(send, 200, {"ok": True})
            return

        try:
            catalog = ReportCatalog.from_file(
                self.settings.reports_path,
                self.settings.default_max_rows,
            )
            telegram = TelegramClient(self.settings.telegram_bot_token)
            executor = ReportExecutor(self.settings)
            sks_service = SksReportService(self.settings)
            dota_service = DotaReportService(self.settings)
            ks_service = KsReportService(self.settings)

            async def run_report(report, parameters):
                return await asyncio.to_thread(executor.run, report, parameters)

            async def send_sks_report(chat_id, date_from, date_to):
                artifact = await asyncio.to_thread(
                    sks_service.create,
                    date_from,
                    date_to,
                )
                await telegram.send_photo(
                    chat_id,
                    artifact.chart,
                    artifact.chart_filename,
                    artifact.caption,
                )
                await telegram.send_document(
                    chat_id,
                    artifact.workbook,
                    artifact.workbook_filename,
                    artifact.caption,
                )

            async def send_dota_report(chat_id, date_from, date_to):
                artifact = await asyncio.to_thread(
                    dota_service.create,
                    date_from,
                    date_to,
                )
                await telegram.send_photo(
                    chat_id,
                    artifact.chart,
                    artifact.chart_filename,
                    artifact.caption,
                )
                for chart in artifact.detail_charts:
                    await telegram.send_photo(
                        chat_id,
                        chart.content,
                        chart.filename,
                        chart.caption,
                    )
                await telegram.send_document(
                    chat_id,
                    artifact.workbook,
                    artifact.workbook_filename,
                    artifact.caption,
                )

            async def load_ks_filters(department_token):
                return await asyncio.to_thread(
                    ks_service.filter_options,
                    department_token,
                )

            async def send_ks_report(
                chat_id,
                report_kind,
                date_from,
                date_to,
                filters,
                comparison_mode,
            ):
                artifact = await asyncio.to_thread(
                    ks_service.create,
                    report_kind,
                    date_from,
                    date_to,
                    filters,
                    comparison_mode,
                )
                await telegram.send_photo(
                    chat_id,
                    artifact.chart,
                    artifact.chart_filename,
                    artifact.caption,
                )
                await telegram.send_document(
                    chat_id,
                    artifact.workbook,
                    artifact.workbook_filename,
                    artifact.caption,
                )

            if callback is not None:
                await handle_callback(
                    callback,
                    self.settings,
                    telegram.send_message,
                    telegram.answer_callback_query,
                    send_sks_report,
                    send_dota_report,
                    send_ks_report,
                    load_ks_filters,
                )
            elif message is not None:
                await handle_message(
                    message,
                    self.settings,
                    catalog,
                    telegram.send_message,
                    run_report,
                    send_sks_report,
                    send_dota_report,
                    send_ks_report,
                    load_ks_filters,
                )
        except Exception:
            logger.exception("Failed to process Telegram update")
            try:
                telegram = TelegramClient(self.settings.telegram_bot_token)
                await telegram.send_message(
                    callback.chat_id if callback is not None else message.chat_id,
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
