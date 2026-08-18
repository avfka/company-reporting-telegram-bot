FROM python:3.12-slim AS base

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

RUN addgroup --system app && adduser --system --ingroup app app

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py reports.json ./
COPY reporting_bot ./reporting_bot

FROM base AS test

COPY requirements-dev.txt pyproject.toml ./
RUN pip install --no-cache-dir -r requirements-dev.txt
COPY scripts ./scripts
COPY tests ./tests
RUN pytest -q

FROM base AS audit

RUN pip install --no-cache-dir pip-audit && pip-audit -r requirements.txt

FROM base AS production

USER app
EXPOSE 8000

CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000"]
