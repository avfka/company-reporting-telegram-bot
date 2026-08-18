from scripts import set_webhook


def test_preview_build_skips_webhook_setup(monkeypatch, capsys) -> None:
    monkeypatch.setenv("VERCEL_ENV", "preview")
    monkeypatch.delenv("TELEGRAM_BOT_TOKEN", raising=False)

    set_webhook.main()

    assert "Skipping webhook setup" in capsys.readouterr().out
