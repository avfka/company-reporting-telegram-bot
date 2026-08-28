from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal
from typing import Mapping, Sequence

from sqlalchemy import create_engine, text
from sqlalchemy.pool import NullPool

from reporting_bot.config import Settings
from reporting_bot.simple_xlsx import STYLE, Workbook


COMPANIES_SQL = """
SELECT
  company.id::text AS company_id,
  company.created_at,
  company.modified_at,
  company.name,
  company.ur_name,
  company.inn,
  company.ogrn,
  company.kpp,
  company.category,
  company.status,
  company.group_id,
  company.archived_at,
  concat_ws(' ', manager.last_name, manager.first_name) AS manager,
  concat_ws(' ', region.full_type, region.name) AS region,
  okved.code AS okved_code,
  okved.name AS okved_name,
  company.main_activity,
  company.employee_count,
  company.capital,
  company.income,
  company.general_director,
  company.okato,
  company.oktmo,
  company.okpo,
  company.okogu,
  coalesce(company.request_count, 0) AS request_count,
  coalesce(company.project_count, 0) AS project_count,
  coalesce(company.task_count, 0) AS task_count,
  company.ur_address,
  company.fiz_address,
  company.mailing_address,
  company.bank_bic,
  company.bank_name,
  company.bank_account,
  company.bank_coraccount,
  company.bank_address,
  company.bank_swift,
  array_to_string(company.synced_to_onec, ', ') AS synced_to_onec
FROM companies_clone company
LEFT JOIN users manager ON manager.id = company.manager_id
LEFT JOIN regions region ON region.id = company.region_id
LEFT JOIN okveds okved ON okved.id = company.okved_id
WHERE company.created_at >= CAST(:date_from AS DATE)
  AND company.created_at < CAST(:date_to_exclusive AS DATE)
ORDER BY company.created_at, company.name, company.id
"""


@dataclass(frozen=True)
class CompanyRow:
    company_id: str
    created_at: datetime
    modified_at: datetime
    name: str
    ur_name: str
    inn: str
    ogrn: str
    kpp: str
    category: str
    status: str
    group_id: int | None
    archived_at: datetime | None
    manager: str
    region: str
    okved_code: str
    okved_name: str
    main_activity: str
    employee_count: int | None
    capital: Decimal | None
    income: Decimal | None
    general_director: str
    okato: str
    oktmo: str
    okpo: str
    okogu: str
    request_count: int
    project_count: int
    task_count: int
    ur_address: str
    fiz_address: str
    mailing_address: str
    bank_bic: str
    bank_name: str
    bank_account: str
    bank_coraccount: str
    bank_address: str
    bank_swift: str
    synced_to_onec: str


@dataclass(frozen=True)
class CompaniesReportData:
    date_from: date
    date_to: date
    companies: tuple[CompanyRow, ...]


@dataclass(frozen=True)
class CompaniesReportArtifact:
    workbook: bytes
    workbook_filename: str
    caption: str


class CompaniesReportService:
    def __init__(self, settings: Settings) -> None:
        self._settings = settings
        self._database_url = settings.database_url.replace(
            "postgresql://", "postgresql+psycopg://", 1
        )

    def create(self, date_from: date, date_to: date) -> CompaniesReportArtifact:
        if date_to < date_from:
            raise ValueError("Дата окончания не может быть раньше даты начала.")
        if (date_to - date_from).days > 366:
            raise ValueError("Максимальный период выгрузки — 366 дней.")
        data = self._load(date_from, date_to)
        stem = f"Компании_{date_from.isoformat()}_{date_to.isoformat()}"
        return CompaniesReportArtifact(
            workbook=build_companies_workbook(data),
            workbook_filename=stem + ".xlsx",
            caption=f"<b>Выгрузка компаний</b> · {date_from:%d.%m.%Y}–{date_to:%d.%m.%Y}",
        )

    def _load(self, date_from: date, date_to: date) -> CompaniesReportData:
        engine = create_engine(
            self._database_url,
            poolclass=NullPool,
            connect_args={"connect_timeout": self._settings.connect_timeout_seconds},
        )
        try:
            with engine.connect() as connection:
                with connection.begin():
                    connection.execute(text("SET TRANSACTION READ ONLY"))
                    connection.execute(
                        text("SELECT set_config('statement_timeout', :timeout, true)"),
                        {"timeout": f"{max(self._settings.statement_timeout_ms, 30_000)}ms"},
                    )
                    rows = connection.execute(
                        text(COMPANIES_SQL),
                        {
                            "date_from": date_from,
                            "date_to_exclusive": date_to + timedelta(days=1),
                        },
                    ).mappings()
                    companies = tuple(_company_from_row(row) for row in rows)
        finally:
            engine.dispose()
        return CompaniesReportData(date_from, date_to, companies)


def _text(value: object) -> str:
    return str(value or "").strip()


def _decimal_or_none(value: object) -> Decimal | None:
    return None if value is None else Decimal(str(value))


def _company_from_row(row: Mapping[str, object]) -> CompanyRow:
    return CompanyRow(
        company_id=_text(row["company_id"]),
        created_at=row["created_at"],
        modified_at=row["modified_at"],
        name=_text(row["name"]) or "—",
        ur_name=_text(row["ur_name"]) or "—",
        inn=_text(row["inn"]) or "—",
        ogrn=_text(row["ogrn"]) or "—",
        kpp=_text(row["kpp"]) or "—",
        category=_text(row["category"]) or "—",
        status=_text(row["status"]) or "—",
        group_id=int(row["group_id"]) if row["group_id"] is not None else None,
        archived_at=row["archived_at"],
        manager=_text(row["manager"]) or "—",
        region=_text(row["region"]) or "—",
        okved_code=_text(row["okved_code"]) or "—",
        okved_name=_text(row["okved_name"]) or "—",
        main_activity=_text(row["main_activity"]) or "—",
        employee_count=int(row["employee_count"]) if row["employee_count"] is not None else None,
        capital=_decimal_or_none(row["capital"]),
        income=_decimal_or_none(row["income"]),
        general_director=_text(row["general_director"]) or "—",
        okato=_text(row["okato"]) or "—",
        oktmo=_text(row["oktmo"]) or "—",
        okpo=_text(row["okpo"]) or "—",
        okogu=_text(row["okogu"]) or "—",
        request_count=int(row["request_count"] or 0),
        project_count=int(row["project_count"] or 0),
        task_count=int(row["task_count"] or 0),
        ur_address=_text(row["ur_address"]) or "—",
        fiz_address=_text(row["fiz_address"]) or "—",
        mailing_address=_text(row["mailing_address"]) or "—",
        bank_bic=_text(row["bank_bic"]) or "—",
        bank_name=_text(row["bank_name"]) or "—",
        bank_account=_text(row["bank_account"]) or "—",
        bank_coraccount=_text(row["bank_coraccount"]) or "—",
        bank_address=_text(row["bank_address"]) or "—",
        bank_swift=_text(row["bank_swift"]) or "—",
        synced_to_onec=_text(row["synced_to_onec"]) or "—",
    )


def _excel_column(count: int) -> str:
    result = ""
    while count:
        count, remainder = divmod(count - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _title(sheet, title: str, period: str, columns: int) -> None:
    sheet.append([title] + [None] * (columns - 1), STYLE["title"])
    sheet.append([period] + [None] * (columns - 1), STYLE["subtitle"])
    last_column = _excel_column(columns)
    sheet.merges.extend([f"A1:{last_column}1", f"A2:{last_column}2"])
    sheet.row_heights[1] = 27
    sheet.row_heights[2] = 22


def _group_label(group_id: int | None) -> str:
    if group_id == 1:
        return "Экостар"
    if group_id == 2:
        return "ГТО"
    return "Не указана"


def build_companies_workbook(data: CompaniesReportData) -> bytes:
    period = f"Период создания: {data.date_from:%d.%m.%Y}–{data.date_to:%d.%m.%Y}"
    workbook = Workbook(f"Выгрузка компаний {period}")
    _build_summary_sheet(workbook, data, period)
    _build_companies_sheet(workbook, data.companies, period)
    _build_requisites_sheet(workbook, data.companies, period)
    return workbook.to_bytes()


def _build_summary_sheet(workbook: Workbook, data: CompaniesReportData, period: str) -> None:
    rows = data.companies
    sheet = workbook.add_sheet("Сводка")
    _title(sheet, "Выгрузка компаний", period, 6)
    sheet.append([None] * 6)
    sheet.append(["Показатель", "Значение", "Единица", "Правило", None, None], STYLE["table_header"])
    metrics = (
        ("Компании за период", len(rows), "шт.", "companies_clone.created_at в выбранном периоде"),
        ("Экостар", sum(row.group_id == 1 for row in rows), "шт.", "group_id = 1"),
        ("ГТО", sum(row.group_id == 2 for row in rows), "шт.", "group_id = 2"),
        ("Активные", sum(row.archived_at is None for row in rows), "шт.", "archived_at не заполнено"),
        ("Архивные", sum(row.archived_at is not None for row in rows), "шт.", "archived_at заполнено"),
        ("С заполненным ИНН", sum(row.inn != "—" for row in rows), "шт.", "ИНН не пустой"),
        ("С назначенным менеджером", sum(row.manager != "—" for row in rows), "шт.", "manager_id заполнен"),
        ("Проектов в карточках", sum(row.project_count for row in rows), "шт.", "Сумма companies_clone.project_count"),
    )
    for label, value, unit, rule in metrics:
        sheet.append(
            [label, value, unit, rule, None, None],
            [STYLE["table_text"], STYLE["table_number"], STYLE["table_center"], STYLE["table_text"], STYLE["base"], STYLE["base"]],
        )
    sheet.append([None] * 6)
    note = sheet.append(
        ["Выгрузка сформирована по актуальной таблице companies_clone. Обе границы выбранного периода включены. На листе «Компании» — основные сведения, на листе «Реквизиты» — адреса и банковские данные."] + [None] * 5,
        STYLE["note"],
    )
    sheet.merges.append(f"A{note}:F{note}")
    sheet.row_heights[note] = 55
    sheet.widths = {0: 34, 1: 18, 2: 14, 3: 43, 4: 12, 5: 12}
    sheet.freeze_rows = 4


def _build_companies_sheet(workbook: Workbook, rows: Sequence[CompanyRow], period: str) -> None:
    columns = 29
    sheet = workbook.add_sheet("Компании")
    _title(sheet, "Компании", period, columns)
    sheet.append([None] * columns)
    sheet.append(
        [
            "Дата создания", "Дата изменения", "Название", "Юридическое наименование", "ИНН", "ОГРН", "КПП",
            "Группа", "Статус", "Категория", "Менеджер", "Регион", "Код ОКВЭД", "ОКВЭД", "Основной вид деятельности",
            "Сотрудников", "Капитал, руб.", "Доход, руб.", "Генеральный директор", "ОКАТО", "ОКТМО", "ОКПО", "ОКОГУ",
            "Заявки, шт.", "Проекты, шт.", "Задачи, шт.", "Дата архивации", "Синхронизация 1С", "Компания ID",
        ],
        STYLE["table_header"],
    )
    for row in rows:
        sheet.append(
            [
                row.created_at, row.modified_at, row.name, row.ur_name, row.inn, row.ogrn, row.kpp,
                _group_label(row.group_id), "Архивная" if row.archived_at else (row.status or "Активная"),
                row.category, row.manager, row.region, row.okved_code, row.okved_name, row.main_activity,
                row.employee_count if row.employee_count is not None else "—",
                row.capital if row.capital is not None else "—", row.income if row.income is not None else "—",
                row.general_director, row.okato, row.oktmo, row.okpo, row.okogu, row.request_count,
                row.project_count, row.task_count, row.archived_at or "—", row.synced_to_onec, row.company_id,
            ],
            [
                STYLE["table_center"], STYLE["table_center"], STYLE["table_text"], STYLE["table_text"],
                STYLE["table_id"], STYLE["table_id"], STYLE["table_id"], STYLE["table_center"],
                STYLE["table_center"], STYLE["table_text"], STYLE["table_text"], STYLE["table_text"],
                STYLE["table_text"], STYLE["table_text"], STYLE["table_text"], STYLE["table_number"],
                STYLE["table_money"], STYLE["table_money"], STYLE["table_text"], STYLE["table_id"],
                STYLE["table_id"], STYLE["table_id"], STYLE["table_id"], STYLE["table_number"],
                STYLE["table_number"], STYLE["table_number"], STYLE["table_center"], STYLE["table_text"],
                STYLE["table_id"],
            ],
        )
    sheet.widths = {
        0: 19, 1: 19, 2: 38, 3: 43, 4: 17, 5: 18, 6: 15, 7: 14, 8: 16, 9: 18,
        10: 28, 11: 27, 12: 16, 13: 42, 14: 42, 15: 15, 16: 18, 17: 18, 18: 32,
        19: 17, 20: 17, 21: 17, 22: 17, 23: 15, 24: 15, 25: 15, 26: 19, 27: 23, 28: 38,
    }
    sheet.freeze_rows = 4
    sheet.auto_filter = f"A4:AC{max(4, len(sheet.rows))}"


def _build_requisites_sheet(workbook: Workbook, rows: Sequence[CompanyRow], period: str) -> None:
    columns = 14
    sheet = workbook.add_sheet("Реквизиты")
    _title(sheet, "Адреса и банковские реквизиты", period, columns)
    sheet.append([None] * columns)
    sheet.append(
        [
            "Компания", "Юридическое наименование", "ИНН", "КПП", "Юридический адрес", "Фактический адрес",
            "Почтовый адрес", "БИК", "Банк", "Расчётный счёт", "Корреспондентский счёт", "Адрес банка",
            "SWIFT", "Компания ID",
        ],
        STYLE["table_header"],
    )
    for row in rows:
        sheet.append(
            [
                row.name, row.ur_name, row.inn, row.kpp, row.ur_address, row.fiz_address, row.mailing_address,
                row.bank_bic, row.bank_name, row.bank_account, row.bank_coraccount, row.bank_address,
                row.bank_swift, row.company_id,
            ],
            [
                STYLE["table_text"], STYLE["table_text"], STYLE["table_id"], STYLE["table_id"],
                STYLE["table_text"], STYLE["table_text"], STYLE["table_text"], STYLE["table_id"],
                STYLE["table_text"], STYLE["table_id"], STYLE["table_id"], STYLE["table_text"],
                STYLE["table_id"], STYLE["table_id"],
            ],
        )
    sheet.widths = {0: 38, 1: 43, 2: 17, 3: 15, 4: 47, 5: 47, 6: 47, 7: 16, 8: 34, 9: 25, 10: 27, 11: 40, 12: 18, 13: 38}
    sheet.freeze_rows = 4
    sheet.auto_filter = f"A4:N{max(4, len(sheet.rows))}"
