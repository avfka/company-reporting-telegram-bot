from datetime import datetime

from reporting_bot.crm_bridge import CrmBridgeRepository


def test_group_contacts_puts_open_project_first() -> None:
    rows = [
        {
            "client_id": "client-1",
            "client_name": "Анна",
            "phone": "+79991234567",
            "second_phone": None,
            "project_id": "project-1",
            "company_id": "company-1",
            "contract_number": "42/26",
            "service": "sout",
            "current_step": "Согласовать отчет у клиента",
            "current_step_updated_at": datetime(2026, 8, 24, 12, 0),
            "closed": False,
            "company_name": "ООО Пример",
            "manager_name": "Алиша Менеджер",
            "manager_telegram_id": "547997434",
        }
    ]

    payload = CrmBridgeRepository._group_contacts(rows)

    project = payload["contacts"][0]["projects"][0]
    assert project["company_id"] == "company-1"
    assert project["current_step"] == "Согласовать отчет у клиента"
    assert project["url"].endswith("/project-1")
