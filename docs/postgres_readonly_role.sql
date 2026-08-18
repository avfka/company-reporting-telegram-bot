-- Run in psql as a database administrator.
-- Before running, set a strong password in a psql variable:
-- \set reporting_bot_password 'replace-with-a-strong-password'

CREATE ROLE reporting_bot LOGIN PASSWORD :'reporting_bot_password';
GRANT CONNECT ON DATABASE company TO reporting_bot;
GRANT USAGE ON SCHEMA public TO reporting_bot;
GRANT SELECT ON ALL TABLES IN SCHEMA public TO reporting_bot;
ALTER DEFAULT PRIVILEGES IN SCHEMA public GRANT SELECT ON TABLES TO reporting_bot;
ALTER ROLE reporting_bot SET default_transaction_read_only = on;
ALTER ROLE reporting_bot SET statement_timeout = '10s';
