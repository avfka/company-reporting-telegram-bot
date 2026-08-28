from datetime import date, datetime
from decimal import Decimal
from io import BytesIO
from zipfile import ZipFile

from reporting_bot.companies_report import CompanyRow, CompaniesReportData, build_companies_workbook


def test_companies_workbook_contains_summary_details_and_requisites() -> None:
    company = CompanyRow(
        company_id="company-1",
        created_at=datetime(2026, 8, 10, 9, 30),
        modified_at=datetime(2026, 8, 11, 10, 0),
        name="Компания Тест",
        ur_name='ООО "Компания Тест"',
        inn="7700000000",
        ogrn="1000000000000",
        kpp="770001001",
        category="client",
        status="active",
        group_id=1,
        archived_at=None,
        manager="Иванов Иван",
        region="г Москва",
        okved_code="62.01",
        okved_name="Разработка ПО",
        main_activity="Разработка ПО",
        employee_count=10,
        capital=Decimal("10000"),
        income=Decimal("5000000"),
        general_director="Петров Пётр",
        okato="1",
        oktmo="2",
        okpo="3",
        okogu="4",
        request_count=2,
        project_count=1,
        task_count=3,
        ur_address="Москва",
        fiz_address="Москва",
        mailing_address="Москва",
        bank_bic="044525225",
        bank_name="Банк",
        bank_account="40700000000000000000",
        bank_coraccount="30100000000000000000",
        bank_address="Москва",
        bank_swift="TEST",
        synced_to_onec="yes",
    )
    content = build_companies_workbook(
        CompaniesReportData(date(2026, 8, 1), date(2026, 8, 31), (company,))
    )
    with ZipFile(BytesIO(content)) as archive:
        workbook_xml = archive.read("xl/workbook.xml").decode("utf-8")
        sheets_xml = "".join(
            archive.read(name).decode("utf-8")
            for name in archive.namelist()
            if name.startswith("xl/worksheets/sheet")
        )
    for sheet_name in ("Сводка", "Компании", "Реквизиты"):
        assert sheet_name in workbook_xml
    assert "Компания Тест" in sheets_xml
    assert "7700000000" in sheets_xml
    assert "Экостар" in sheets_xml
