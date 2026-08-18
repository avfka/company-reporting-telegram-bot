from reporting_bot.config import Settings


def test_ready_requires_all_runtime_secrets() -> None:
    settings = Settings(
        telegram_bot_token="token",
        telegram_webhook_secret="secret",
        allowed_user_ids=frozenset({42}),
        database_url="postgresql+psycopg://user:pass@localhost/db",
    )
    assert settings.ready is True


def test_not_ready_without_allowlist() -> None:
    settings = Settings(
        telegram_bot_token="token",
        telegram_webhook_secret="secret",
        allowed_user_ids=frozenset(),
        database_url="postgresql+psycopg://user:pass@localhost/db",
    )
    assert settings.ready is False
    assert settings.webhook_ready is True
