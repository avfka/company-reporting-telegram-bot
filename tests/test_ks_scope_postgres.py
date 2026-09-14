"""Read-only SQL fixtures. Set KS_TEST_DATABASE_URL to execute in PostgreSQL.

All fixture relations are CTEs; no CRM rows or schemas are modified.
"""
import json
import os
from datetime import datetime

import pytest
from sqlalchemy import create_engine, text

from reporting_bot.ks_scope import PROJECTS_SQL, PAYMENTS_SQL, ROSTER


@pytest.fixture
def connection():
    url = os.environ.get("KS_TEST_DATABASE_URL")
    if not url:
        pytest.skip("KS_TEST_DATABASE_URL is not configured")
    engine = create_engine(url)
    with engine.connect() as conn:
        conn.execute(text("SET TRANSACTION READ ONLY"))
        yield conn
    engine.dispose()


def uid(n):
    return f"00000000-0000-0000-0000-{n:012d}"


def fixture_ctes():
    manager = ROSTER[0][0]
    gto = ROSTER[9][0]
    projects = ','.join(
        f"('{uid(i)}'::uuid, '{owner}'::uuid, null::uuid, 'соут', 100::numeric, 'contract', 'Распечатка', {group})"
        for i,owner,group in ((1,manager,1),(2,manager,1),(3,manager,1),(4,manager,1),
                              (5,manager,1),(6,manager,1),(7,uid(999),2),(8,gto,2))
    )
    assignments = (
        (1,'2026-08-31 23:59:59','measurer_id',None,uid(50)),
        (1,'2026-09-02','expert_id',None,uid(51)),
        (2,'2026-09-01','measurer_id',None,uid(50)),
        (2,'2026-09-04','expert_id',None,uid(51)),
        (2,'2026-09-05','expert_id',uid(51),uid(52)),
        (3,'2026-09-14 23:59:59','measurer_id',None,uid(50)),
        (4,'2026-09-03','expert_id',uid(51),uid(52)),
        (5,'2026-09-03','expert_id',uid(51),None),
        (6,'2026-09-15','expert_id',None,uid(51)),
        (7,'2026-09-03','expert_id',None,uid(51)),
        (8,'2026-09-03','expert_id',None,uid(51)),
    )
    logs = ','.join(
        f"('{uid(i)}'::uuid,'{at}'::timestamp,'project','project:updated','{json.dumps({'updates':[{'property':field,'before':before,'after':after}]})}'::json)"
        for i,at,field,before,after in assignments
    )
    payments = ','.join(
        f"('{uid(100+i)}'::uuid,'{uid(project)}'::uuid,{at}, {amount}::float8, null::text, {deleted}, {group})"
        for i,project,at,amount,deleted,group in (
            (1,1,"'2026-09-02'::date",25,'null::timestamp',1),
            (2,2,"'2026-08-31'::date",10,'null::timestamp',1),
            (3,2,"'2026-09-04'::date",20,'null::timestamp',1),
            (4,2,"'2026-09-05'::date",30,'null::timestamp',1),
            (5,2,"'2026-09-06'::date",-5,'null::timestamp',1),
            (6,2,"'2026-09-06'::date",7,"'2026-09-07'::timestamp",1),
            (7,2,"'2026-09-15'::date",15,'null::timestamp',1),
            (8,7,"'2026-09-03'::date",40,'null::timestamp',2),
            (9,8,"'2026-09-03'::date",40,'null::timestamp',2),
        )
    )
    return f"""WITH projects(id,manager_id,request_id,service,sale_price,contract_number,current_step,group_id) AS (VALUES {projects}),
      systemlogs(target_id,created_at,target_type,event,details) AS (VALUES {logs}),
      payment(id,"projectId","chargeAt",amount,"contractNumber","deletedAt",group_id) AS (VALUES {payments}),
      requests_clone(id,created_at) AS (SELECT null::uuid,null::timestamp WHERE false),
      projects_steps_history(project_id,new_step) AS (SELECT null::uuid,null::text WHERE false),
    """


PARAMS = dict(date_from='2026-09-01', date_to_exclusive='2026-09-15',
              department_token='-', manager_token='-', service='-')


def run(connection, sql, **filters):
    return list(connection.execute(text(fixture_ctes() + sql.removeprefix('WITH ')), PARAMS | filters).mappings())


def test_occurrence_first_of_two_assignments_once_with_boundaries(connection):
    rows = run(connection, PROJECTS_SQL)
    assert {r['event_id'] for r in rows} == {uid(2),uid(3),uid(8)}
    assert len(rows) == 3
    assert sum(r['amount'] for r in rows) == 300
    p2 = next(r for r in rows if r['event_id'] == uid(2))
    assert p2['event_at'] == datetime(2026,9,1)
    assert p2['paid_amount'] == 75  # Includes payments outside this period, not refunds/deleted rows.
    assert p2['manager'] == ROSTER[0][1]
    assert p2['service'] == 'sout'


def test_cash_payment_dates_and_explicit_roster_not_request_cohort(connection):
    rows = run(connection, PAYMENTS_SQL)
    assert {r['event_id'] for r in rows} == {uid(101),uid(103),uid(104),uid(109)}
    assert sum(r['amount'] for r in rows) == 115


def test_department_and_manager_filter_scope(connection):
    rows = run(connection, PROJECTS_SQL, department_token='siber')
    assert [r['event_id'] for r in rows] == [uid(8)]
    rows = run(connection, PROJECTS_SQL, manager_token=ROSTER[0][0].replace('-','')[:6])
    assert {r['event_id'] for r in rows} == {uid(2),uid(3)}
