"""Approved September 2026 KS roster, keyed by CRM user identity.

Source: September 2026.xlsx, 01.09, A3:A29, supplied by the user.
This is reporting classification only; CRM departments are not modified.
"""

DEPARTMENTS = (("op", "ОП"), ("siber", "ОП Сибирь"), ("kam", "КАМ"),
               ("skam", "СпецКАМ"), ("aop", "АОП"))

# Inactive duplicate identities are deliberately not matched by display name.
ROSTER = (
    ("3993e377-8b48-4e5f-b2e5-cb371797fa5b", "Десюкевич Наталья", "op", 1),
    ("1e60e0bf-cbbb-4782-998f-295cd9a8340b", "Шереметьев Владислав", "op", 1),
    ("9f03b421-d638-413a-a1e8-e78ce436e706", "Игнатьев Кирилл", "op", 1),
    ("574a298b-1e62-4b97-bdf1-991e97ecead2", "Медякова Ксения", "op", 1),
    ("834950f8-bcf6-4df3-b86c-c7ad847dbedd", "Тикунов Михаил", "op", 1),
    ("04476a9f-c9de-4d1d-a3ac-e69b99d1b109", "Витовский Артур", "op", 1),
    ("fafdcb95-4b03-4371-b7c6-fa757a2b0445", "Головин Дмитрий", "op", 1),
    ("3d53949d-f199-442e-9e75-0ec2b1c329d6", "Шакиров Руслан", "op", 1),
    ("3c63f04a-3c85-43e5-bc26-ed3fc66845a8", "Васильева Анастасия", "op", 1),
    ("974e31a5-ad58-49d0-ba5a-8471c5caa882", "Айсина Юлианна", "siber", 2),
    ("03b66d40-d48a-4827-a3d3-b8e01d7aac19", "Мельникова Ольга", "siber", 2),
    ("c3d117f0-b5e6-46be-a931-5563c57f0ce9", "Рудакова Ирина", "siber", 2),
    ("485a0abf-67f5-4d47-ac0f-ecd8e89febc3", "Воробьева Наталья", "siber", 2),
    ("19641278-eea5-4dbd-a375-a0cf50f1409e", "Васильева Татьяна", "kam", 1),
    ("815a3c6e-1819-45b7-8e05-f27ea8b2ffb1", "Шергина Надежда", "kam", 1),
    ("c3cbfd78-a733-4f8b-b2b1-b3a349d31b80", "Бурдейная Дарья", "kam", 1),
    ("591a29a4-2bd8-4c22-969b-503dda7f6e9a", "Мохова Жанна", "skam", 1),
    ("52cdce66-c2a6-48c7-a0b8-647ba6f244f0", "Журавский Александр", "aop", 1),
    ("faf9c273-04cf-4d5c-a66e-54da06af5120", "Степанова Марина", "aop", 1),
    ("7a792798-cb2b-423b-ab5d-dbaf0ce1b7d5", "Холявченко Алёна", "aop", 1),
    ("2f4f2f7f-698c-4147-99ce-78871a8bcc2a", "Шаманская Вероника", "aop", 1),
    ("4316114b-1d6a-4e09-aa46-a6343aef4093", "Васильева Алёна", "aop", 2),
)


def _literal(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


ROSTER_CTE = "WITH ks_roster(user_id, manager, department_token, department, group_id) AS (VALUES\n" + ",\n".join(
    "(" + ", ".join((_literal(uid) + "::uuid", _literal(name), _literal(dept),
                      _literal(dict(DEPARTMENTS)[dept]), str(group))) + ")"
    for uid, name, dept, group in ROSTER
) + ")\n"


def service_sql(expr: str) -> str:
    # CRM contains both Latin service codes and the Cyrillic spelling of СОУТ.
    return f"CASE WHEN lower(trim(coalesce({expr}, ''))) = 'соут' THEN 'sout' ELSE lower(trim(coalesce({expr}, ''))) END"


def filters(service: str) -> str:
    normalized = service_sql(service)
    return f"""
  AND (:department_token = '-' OR k.department_token = :department_token)
  AND (:manager_token = '-' OR left(replace(k.user_id::text, '-', ''), 6) = ANY(string_to_array(:manager_token, ',')))
  AND (:service = '-' OR left(md5({normalized}), 6) = :service)
"""


DEPARTMENT_OPTIONS_SQL = ROSTER_CTE + "SELECT DISTINCT department_token AS token, department AS label FROM ks_roster ORDER BY label"
MANAGER_OPTIONS_SQL = ROSTER_CTE + """
SELECT left(replace(k.user_id::text, '-', ''), 6) AS token, k.manager AS label
FROM ks_roster k JOIN users u ON u.id = k.user_id
WHERE :department_token = '-' OR k.department_token = :department_token
ORDER BY label
"""
PRODUCT_OPTIONS_SQL = ROSTER_CTE + f"""
SELECT left(md5(service), 6) AS token, service AS label FROM (
 SELECT {service_sql('r.service')} AS service FROM requests_clone r JOIN ks_roster k ON k.user_id=r.manager_id
 UNION
 SELECT {service_sql('p.service')} AS service FROM projects p JOIN ks_roster k ON k.user_id=p.manager_id
) s WHERE service NOT IN ('', 'zamery', 'замеры') ORDER BY service
"""

REQUESTS_SQL = ROSTER_CTE + f"""
SELECT r.id::text AS event_id, r.created_at AS event_at, r.created_at AS cohort_at,
 k.manager, k.department, {service_sql('r.service')} AS service,
 coalesce(r.price, 0) AS amount, coalesce(r.status, '—') AS status, coalesce(r.stage, '—') AS stage
FROM requests_clone r JOIN ks_roster k ON k.user_id=r.manager_id
WHERE r.group_id IN (1,2) AND r.deleted_at IS NULL AND r.archived_at IS NULL
 AND r.created_at >= CAST(:date_from AS DATE) AND r.created_at < CAST(:date_to_exclusive AS DATE)
""" + filters('r.service') + "ORDER BY r.created_at, r.id LIMIT 10001"

OFFERS_SQL = ROSTER_CTE + f"""
SELECT o.id::text AS event_id, o.created_at AS event_at, r.created_at AS cohort_at,
 k.manager, k.department, {service_sql("coalesce(nullif(o.service, ''), r.service)")} AS service,
 coalesce(r.price,0) AS amount, coalesce(o.number,'—') AS reference
FROM offers o JOIN requests_clone r ON r.id=o.request_id
JOIN ks_roster k ON k.user_id=r.manager_id
WHERE r.group_id IN (1,2) AND r.deleted_at IS NULL AND r.archived_at IS NULL
 AND o.created_at >= CAST(:date_from AS DATE) AND o.created_at < CAST(:date_to_exclusive AS DATE)
""" + filters("coalesce(nullif(o.service, ''), r.service)") + "ORDER BY o.created_at,o.id LIMIT 10001"

# Take the first actual assignment, across BOTH fields and all historical dates.
# A replacement of a non-empty assignee is not the initial assignment.
FIRST_ASSIGNMENT_JOIN = """
JOIN LATERAL (
 SELECT min(l.created_at) AS assigned_at
 FROM systemlogs l
 CROSS JOIN LATERAL jsonb_array_elements(
   CASE WHEN jsonb_typeof(l.details::jsonb->'updates')='array'
     THEN l.details::jsonb->'updates' ELSE '[]'::jsonb END
 ) change
 WHERE l.target_type='project' AND l.target_id=p.id
   AND l.event IN ('project:created','project:updated')
   AND change->>'property' IN ('measurer_id','expert_id')
   AND nullif(trim(change->>'after'), '') IS NOT NULL
   AND nullif(trim(change->>'before'), '') IS NULL
) assigned ON true
"""

PROJECTS_SQL = ROSTER_CTE + """
, assignment_candidates AS MATERIALIZED (
 SELECT DISTINCT l.target_id FROM systemlogs l
 WHERE l.target_type='project' AND l.event IN ('project:created','project:updated')
   AND l.created_at >= CAST(:date_from AS DATE)
   AND l.created_at < CAST(:date_to_exclusive AS DATE)
   AND EXISTS (
     SELECT 1 FROM jsonb_array_elements(CASE
       WHEN jsonb_typeof(l.details::jsonb->'updates')='array'
       THEN l.details::jsonb->'updates' ELSE '[]'::jsonb END) change
     WHERE change->>'property' IN ('measurer_id','expert_id')
       AND nullif(trim(change->>'after'), '') IS NOT NULL
       AND nullif(trim(change->>'before'), '') IS NULL
   )
)
""" + f"""
SELECT p.id::text AS event_id, assigned.assigned_at AS event_at,
 r.created_at AS cohort_at, p.id::text AS project_id, k.manager, k.department,
 {service_sql('p.service')} AS service, coalesce(p.sale_price,0) AS amount,
 coalesce(p.contract_number,'—') AS reference, coalesce(p.current_step,'—') AS current_step,
 coalesce(paid.total,0) AS paid_amount,
 (p.current_step='Распечатка' OR EXISTS (
   SELECT 1 FROM projects_steps_history psh WHERE psh.project_id=p.id AND psh.new_step = 'Распечатка'
 )) AS is_agreed
FROM projects p JOIN ks_roster k ON k.user_id=p.manager_id
JOIN assignment_candidates candidate ON candidate.target_id=p.id
LEFT JOIN requests_clone r ON r.id=p.request_id
""" + FIRST_ASSIGNMENT_JOIN + """
LEFT JOIN LATERAL (
 SELECT sum(pay.amount) AS total FROM payment pay
 WHERE pay."projectId"=p.id AND pay."deletedAt" IS NULL AND pay.amount>0 AND pay.group_id IN (1,2)
) paid ON true
WHERE p.group_id IN (1,2)
 AND assigned.assigned_at >= CAST(:date_from AS DATE)
 AND assigned.assigned_at < CAST(:date_to_exclusive AS DATE)
""" + filters('p.service') + "ORDER BY assigned.assigned_at,p.id LIMIT 10001"

PAYMENTS_SQL = ROSTER_CTE + f"""
SELECT pay.id::text AS event_id, pay."chargeAt" AS event_at,
 r.created_at AS cohort_at, p.id::text AS project_id, k.manager, k.department,
 {service_sql('p.service')} AS service, pay.amount,
 coalesce(pay."contractNumber",p.contract_number,'—') AS reference,
 coalesce(p.current_step,'—') AS current_step
FROM payment pay JOIN projects p ON p.id=pay."projectId"
JOIN ks_roster k ON k.user_id=p.manager_id
LEFT JOIN requests_clone r ON r.id=p.request_id
WHERE p.group_id IN (1,2) AND pay.group_id IN (1,2)
 AND pay."deletedAt" IS NULL AND pay.amount > 0
 AND pay."chargeAt" >= CAST(:date_from AS DATE)
 AND pay."chargeAt" < CAST(:date_to_exclusive AS DATE)
""" + filters('p.service') + 'ORDER BY pay."chargeAt",pay.id LIMIT 10001'

CALLS_SQL = ROSTER_CTE + """
SELECT c.id::text AS event_id,c.call_date AS event_at,k.manager,k.department,
 '' AS service,0 AS amount,coalesce(c.direction::text,'—') AS reference
FROM calls c JOIN ks_roster k ON k.user_id=c.user_id
WHERE c.call_date >= CAST(:date_from AS DATE) AND c.call_date < CAST(:date_to_exclusive AS DATE)
 AND :service = '-'
""" + filters("''") + "ORDER BY c.call_date,c.id LIMIT 10001"
