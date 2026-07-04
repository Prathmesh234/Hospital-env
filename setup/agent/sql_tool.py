"""Guarded, read-only SQL execution for agents.

The guard is defense-in-depth, not a parser: the query text is stripped of
string literals and comments, then checked to be a single statement whose
first keyword is ``SELECT`` (or ``WITH``) and which contains no write/DDL
keywords anywhere. On top of that, every query runs inside a transaction
that is explicitly ``READ ONLY`` with a server-side ``statement_timeout``,
and is rolled back afterwards regardless of outcome.
"""

from __future__ import annotations

import csv
import io
import json
import re
import time
import uuid
from dataclasses import dataclass, field
from datetime import date, datetime
from datetime import time as dt_time
from decimal import Decimal
from typing import Any

from sqlalchemy import text

from setup.db.postgres import get_engine

DEFAULT_ROW_LIMIT = 50
MAX_ROW_LIMIT = 1000
DEFAULT_TIMEOUT_MS = 15_000

_FORBIDDEN_KEYWORDS = {
    "insert", "update", "delete", "merge", "upsert",
    "drop", "alter", "create", "truncate", "rename",
    "grant", "revoke", "security",
    "copy", "vacuum", "cluster", "reindex", "refresh",
    "set", "reset", "call", "do", "execute", "prepare", "deallocate",
    "listen", "notify", "lock", "comment", "checkpoint", "discard",
    "pg_sleep", "pg_read_file", "pg_write_file", "lo_import", "lo_export",
}

_STRING_OR_COMMENT = re.compile(
    r"""
    '(?:[^']|'')*'          # single-quoted string (with '' escapes)
    | "(?:[^"]|"")*"        # double-quoted identifier
    | \$\$.*?\$\$           # dollar-quoted string
    | --[^\n]*              # line comment
    | /\*.*?\*/             # block comment
    """,
    re.VERBOSE | re.DOTALL,
)


class SQLGuardError(Exception):
    """Raised when a query is rejected by the read-only guard."""


@dataclass
class QueryResult:
    columns: list[str]
    rows: list[list[Any]] = field(default_factory=list)
    truncated: bool = False
    elapsed_ms: float = 0.0

    @property
    def rowcount(self) -> int:
        return len(self.rows)

    def to_json(self, indent: int | None = 2) -> str:
        records = [dict(zip(self.columns, row, strict=True)) for row in self.rows]
        return json.dumps(
            {"rowcount": self.rowcount, "truncated": self.truncated, "rows": records},
            indent=indent,
            default=str,
        )

    def to_csv(self) -> str:
        buf = io.StringIO()
        writer = csv.writer(buf)
        writer.writerow(self.columns)
        writer.writerows(self.rows)
        return buf.getvalue()


def _scrub(sql: str) -> str:
    """Remove string literals, quoted identifiers, and comments."""
    return _STRING_OR_COMMENT.sub(" ", sql)


def validate_readonly_sql(sql: str) -> None:
    """Raise :class:`SQLGuardError` unless ``sql`` is a single read-only statement."""
    scrubbed = _scrub(sql).strip()
    if not scrubbed:
        raise SQLGuardError("Empty query.")

    # single statement only: a trailing semicolon is fine, embedded ones are not
    if ";" in scrubbed.rstrip().rstrip(";"):
        raise SQLGuardError("Multiple statements are not allowed — submit one query at a time.")

    first_word = re.match(r"\s*(\w+)", scrubbed)
    if first_word is None or first_word.group(1).lower() not in {"select", "with", "table", "explain"}:
        raise SQLGuardError(
            "Only read queries are allowed (must start with SELECT, WITH, TABLE, or EXPLAIN)."
        )

    words = {w.lower() for w in re.findall(r"\w+", scrubbed)}
    hits = sorted(words & _FORBIDDEN_KEYWORDS)
    if hits:
        raise SQLGuardError(f"Query rejected — forbidden keyword(s): {', '.join(hits)}.")


def _jsonable(value: Any) -> Any:
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, (datetime, date, dt_time)):
        return value.isoformat()
    if isinstance(value, uuid.UUID):
        return str(value)
    if isinstance(value, (bytes, memoryview)):
        return bytes(value).hex()
    return value


def run_sql(
    sql: str,
    limit: int = DEFAULT_ROW_LIMIT,
    timeout_ms: int = DEFAULT_TIMEOUT_MS,
) -> QueryResult:
    """Execute a guarded read-only query and return a :class:`QueryResult`.

    Fetches ``limit + 1`` rows so truncation is reported honestly. The
    transaction is always rolled back.
    """
    validate_readonly_sql(sql)
    limit = max(1, min(int(limit), MAX_ROW_LIMIT))

    engine = get_engine()
    started = time.perf_counter()
    with engine.connect() as conn:
        tx = conn.begin()
        try:
            conn.exec_driver_sql("SET TRANSACTION READ ONLY")
            conn.exec_driver_sql(f"SET LOCAL statement_timeout = {int(timeout_ms)}")
            cursor = conn.execute(text(sql.rstrip().rstrip(";")))
            columns = list(cursor.keys())
            fetched = cursor.fetchmany(limit + 1)
        finally:
            tx.rollback()

    elapsed_ms = (time.perf_counter() - started) * 1000
    truncated = len(fetched) > limit
    rows = [[_jsonable(v) for v in row] for row in fetched[:limit]]
    return QueryResult(columns=columns, rows=rows, truncated=truncated, elapsed_ms=elapsed_ms)
