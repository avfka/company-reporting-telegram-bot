from __future__ import annotations

import os
import sys

import httpx


def required(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise SystemExit(f"Missing environment variable: {name}")
    return value


def main() -> None:
    vercel_environment = os.getenv("VERCEL_ENV", "").strip()
    if vercel_environment and vercel_environment != "production":
        print(f"Skipping webhook setup for Vercel environment: {vercel_environment}")
        return

    token = required("TELEGRAM_BOT_TOKEN")
    secret = required("TELEGRAM_WEBHOOK_SECRET")
    app_url = required("APP_URL").rstrip("/")
    response = httpx.post(
        f"https://api.telegram.org/bot{token}/setWebhook",
        json={
            "url": f"{app_url}/telegram/webhook",
            "secret_token": secret,
            "allowed_updates": ["message"],
            "drop_pending_updates": False,
        },
        timeout=15,
    )
    response.raise_for_status()
    payload = response.json()
    if not payload.get("ok"):
        raise SystemExit(f"Telegram rejected webhook: {payload}")
    print("Webhook configured successfully")


if __name__ == "__main__":
    try:
        main()
    except httpx.HTTPError as exc:
        print(f"Webhook configuration failed: {exc}", file=sys.stderr)
        raise SystemExit(1) from exc
