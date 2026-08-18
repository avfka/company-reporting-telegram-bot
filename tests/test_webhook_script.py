from scripts import set_webhook


def test_preview_build_skips_webhook_setup(monkeypatch, capsys) -> None:
    monkeypatch.setenv("VERCEL_ENV", "preview")
    monkeypatch.delenv("TELEGRAM_BOT_TOKEN", raising=False)

    set_webhook.main()

    assert "Skipping webhook setup" in capsys.readouterr().out


def test_production_webhook_subscribes_to_calendar_callbacks(monkeypatch) -> None:
    captured = {}

    class Response:
        def raise_for_status(self):
            pass

        def json(self):
            return {"ok": True}

    def post(url, json, timeout):
        captured.update({"url": url, "json": json, "timeout": timeout})
        return Response()

    monkeypatch.setenv("VERCEL_ENV", "production")
    monkeypatch.setenv("TELEGRAM_BOT_TOKEN", "token")
    monkeypatch.setenv("TELEGRAM_WEBHOOK_SECRET", "secret")
    monkeypatch.setenv("APP_URL", "https://example.com")
    monkeypatch.setattr(set_webhook.httpx, "post", post)

    set_webhook.main()

    assert captured["json"]["allowed_updates"] == ["message", "callback_query"]
