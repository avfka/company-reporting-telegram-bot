# Telegram-бот для отчётности по PostgreSQL

Безопасный MVP корпоративного Telegram-бота: пользователь выбирает заранее утверждённый отчёт, бот выполняет параметризованный `SELECT` в read-only транзакции и возвращает компактную таблицу.

## Что уже есть

- минимальный ASGI webhook API без тяжёлого веб-фреймворка;
- доступ только для Telegram ID из белого списка;
- выдача отчётов только в личных чатах, без групп;
- `/whoami`, `/reports`, `/run <report> [name=value]`;
- SQL только из версионируемого `reports.json`;
- PostgreSQL read-only транзакция, таймаут подключения и запроса;
- лимит строк и разбиение длинных Telegram-сообщений;
- health endpoint `/health` без раскрытия секретов;
- конфигурация для Vercel и Docker;
- тесты ключевых ограничений безопасности;
- CI с тестами и аудитом production-зависимостей.

## Почему нет произвольного SQL из чата

Текстовые SQL-запросы от пользователей дают слишком большой риск утечки или изменения данных. В этом проекте запросы добавляет разработчик в `reports.json`, а значения пользователей передаются только как bind-параметры. Дополнительно запрос исполняется в read-only транзакции. Для подключения к продакшен-БД всё равно нужен отдельный пользователь только с `SELECT`.

## Быстрый запуск

1. Создайте бота через `@BotFather` и получите токен.
2. Скопируйте `.env.example` в `.env` и заполните значения.
3. Создайте read-only пользователя PostgreSQL по примеру `docs/postgres_readonly_role.sql`.
4. Установите зависимости и запустите API:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements-dev.txt
set -a; source .env; set +a
uvicorn app:app --reload
```

5. Для локального webhook используйте HTTPS-туннель. Задайте публичный URL и зарегистрируйте webhook:

```bash
export APP_URL=https://your-public-url.example
python scripts/set_webhook.py
```

Если Telegram ID ещё неизвестен, сначала задайте токен и webhook-secret, зарегистрируйте webhook и отправьте боту `/whoami`. После этого добавьте полученный ID в `ALLOWED_TELEGRAM_USER_IDS` и настройте `DATABASE_URL`.

## Команды бота

```text
/whoami
/reports
/run db_status
/run table_catalog
/run daily_sales report_date=2026-08-18
```

`daily_sales` — только пример из `examples/business_reports.json`; адаптируйте имена таблиц и полей и перенесите запись в `reports.json`.

## Добавление отчёта

Каждая запись в `reports.json` содержит:

- `id` — команда отчёта;
- `title` и `description` — подпись в Telegram;
- `parameters` — разрешённые имена параметров;
- `max_rows` — лимит от 1 до 500;
- `sql` — один `SELECT` или `WITH ... SELECT` без изменяющих операций.

Пример:

```json
{
  "id": "sales_by_period",
  "title": "Продажи за период",
  "description": "Выручка по отделам",
  "parameters": ["date_from", "date_to"],
  "max_rows": 100,
  "sql": "SELECT department, SUM(total_amount) AS revenue FROM orders WHERE created_at >= CAST(:date_from AS DATE) AND created_at < CAST(:date_to AS DATE) + INTERVAL '1 day' GROUP BY department ORDER BY revenue DESC"
}
```

## Переменные окружения

| Переменная | Назначение |
|---|---|
| `TELEGRAM_BOT_TOKEN` | токен от BotFather |
| `TELEGRAM_WEBHOOK_SECRET` | случайная строка из букв, цифр, `_`, `-` |
| `APP_URL` | production URL приложения без завершающего `/` |
| `ALLOWED_TELEGRAM_USER_IDS` | Telegram ID через запятую |
| `ALLOWED_TELEGRAM_USER_IDS_EXTRA` | дополнительные Telegram ID без замены основного списка |
| `DATABASE_URL` | URL PostgreSQL для read-only пользователя |
| `REPORTS_PATH` | путь к JSON с отчётами |
| `DB_STATEMENT_TIMEOUT_MS` | максимальное время SQL-запроса |
| `DB_CONNECT_TIMEOUT_SECONDS` | таймаут подключения к БД |
| `DEFAULT_MAX_ROWS` | лимит строк по умолчанию |

## Деплой на Vercel

Vercel распознаёт переменную `app` в `app.py` как ASGI-приложение. Добавьте переменные в Production и выполните production deploy. Build hook автоматически зарегистрирует Telegram webhook только для production-окружения.

Важно: Vercel должен иметь сетевой доступ к PostgreSQL. Если база доступна только внутри корпоративной сети/VPN, разместите сервис в той же сети или настройте безопасный прокси/туннель; не открывайте БД всему интернету.
