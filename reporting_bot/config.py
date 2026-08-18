from __future__ import annotations

import os
from dataclasses import dataclass


def _positive_int(name: str, default: int) -> int:
    raw = os.getenv(name, str(default)).strip()
    try:
        value = int(raw)
    except ValueError as exc:
        raise ValueError(f"{name} must be an integer") from exc
    if value <= 0:
        raise ValueError(f"{name} must be positive")
    return value


def _user_ids(raw: str) -> frozenset[int]:
    if not raw.strip():
        return frozenset()
    try:
        return frozenset(int(value.strip()) for value in raw.split(",") if value.strip())
    except ValueError as exc:
        raise ValueError("ALLOWED_TELEGRAM_USER_IDS must contain comma-separated integers") from exc


@dataclass(frozen=True)
class Settings:
    telegram_bot_token: str
    telegram_webhook_secret: str
    allowed_user_ids: frozenset[int]
    database_url: str
    reports_path: str = "reports.json"
    statement_timeout_ms: int = 10_000
    connect_timeout_seconds: int = 5
    default_max_rows: int = 50

    @classmethod
    def from_env(cls) -> "Settings":
        return cls(
            telegram_bot_token=os.getenv("TELEGRAM_BOT_TOKEN", "").strip(),
            telegram_webhook_secret=os.getenv("TELEGRAM_WEBHOOK_SECRET", "").strip(),
            allowed_user_ids=_user_ids(os.getenv("ALLOWED_TELEGRAM_USER_IDS", "")),
            database_url=os.getenv("DATABASE_URL", "").strip(),
            reports_path=os.getenv("REPORTS_PATH", "reports.json").strip() or "reports.json",
            statement_timeout_ms=_positive_int("DB_STATEMENT_TIMEOUT_MS", 10_000),
            connect_timeout_seconds=_positive_int("DB_CONNECT_TIMEOUT_SECONDS", 5),
            default_max_rows=_positive_int("DEFAULT_MAX_ROWS", 50),
        )

    @property
    def webhook_ready(self) -> bool:
        return bool(self.telegram_bot_token and self.telegram_webhook_secret)

    @property
    def ready(self) -> bool:
        return bool(self.webhook_ready and self.database_url and self.allowed_user_ids)
