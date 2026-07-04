"""Unit tests for the read-only SQL guard — no database required."""

from __future__ import annotations

import pytest

from setup.agent.sql_tool import SQLGuardError, validate_readonly_sql

ALLOWED = [
    "SELECT 1",
    "select count(*) from patients",
    "  WITH x AS (SELECT 1 AS n) SELECT n FROM x",
    "SELECT * FROM claims WHERE status = 'denied';",
    "EXPLAIN SELECT 1",
    "TABLE payers",
    # write-ish words inside string literals / identifiers must not trip the guard
    "SELECT * FROM tasks WHERE title = 'update the chart'",
    "SELECT created_at, updated_at FROM patients",
    "SELECT * FROM call_logs",
]

REJECTED = [
    "",
    "   ",
    "DROP TABLE patients",
    "DELETE FROM patients",
    "INSERT INTO patients (id) VALUES ('x')",
    "UPDATE claims SET status = 'paid'",
    "TRUNCATE claims",
    "CREATE TABLE evil (id int)",
    "ALTER TABLE claims ADD COLUMN x int",
    "GRANT ALL ON patients TO public",
    "SELECT 1; DROP TABLE patients",
    "SELECT 1; SELECT 2",
    "COPY patients TO '/tmp/exfil.csv'",
    "SELECT pg_sleep(60)",
    "DO $$ BEGIN NULL; END $$",
    "SET search_path TO evil",
    "-- sneaky\nDELETE FROM patients",
    "WITH x AS (SELECT 1) INSERT INTO patients SELECT * FROM x",
]


@pytest.mark.parametrize("sql", ALLOWED)
def test_allows_read_queries(sql):
    validate_readonly_sql(sql)


@pytest.mark.parametrize("sql", REJECTED)
def test_rejects_writes_and_multistatements(sql):
    with pytest.raises(SQLGuardError):
        validate_readonly_sql(sql)
