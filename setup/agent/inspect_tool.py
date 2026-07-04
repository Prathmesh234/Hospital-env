"""Schema introspection helpers for agents: table lists, column details, samples."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from sqlalchemy import text

from setup.agent.sql_tool import run_sql
from setup.db.postgres import get_engine
from setup.models import Base

_DOMAIN_LABELS = {
    "base": "catalogs",
    "providers": "organization",
    "patients": "patients",
    "scheduling": "scheduling",
    "encounters": "encounters",
    "clinical": "clinical",
    "medications": "medications",
    "labs": "labs_imaging",
    "billing": "billing",
    "communications": "communications",
    "operations": "operations",
}


def _table_domains() -> dict[str, str]:
    """Map ``table_name -> domain`` from the model module each class lives in."""
    domains: dict[str, str] = {}
    for mapper in Base.registry.mappers:
        module = mapper.class_.__module__.rsplit(".", 1)[-1]
        domains[mapper.local_table.name] = _DOMAIN_LABELS.get(module, module)
    return domains


def list_tables(domain: str | None = None, with_counts: bool = True) -> list[dict[str, Any]]:
    """All tables (optionally one domain), each with its live row count."""
    domains = _table_domains()
    names = sorted(Base.metadata.tables.keys())
    out: list[dict[str, Any]] = []

    counts: dict[str, int] = {}
    if with_counts:
        engine = get_engine()
        with engine.connect() as conn:
            union = " UNION ALL ".join(
                f"SELECT '{name}' AS t, count(*) AS n FROM {name}" for name in names
            )
            for t, n in conn.execute(text(union)):
                counts[t] = n

    for name in names:
        dom = domains.get(name, "?")
        if domain and dom != domain:
            continue
        out.append({"table": name, "domain": dom, "rows": counts.get(name)})
    return out


def list_domains() -> list[str]:
    return sorted(set(_table_domains().values()))


@dataclass
class ColumnInfo:
    name: str
    type: str
    nullable: bool
    primary_key: bool
    foreign_keys: list[str]


def describe_table(name: str) -> dict[str, Any]:
    """Columns, PKs, FKs, and indexes for one table (from the ORM metadata)."""
    table = Base.metadata.tables.get(name)
    if table is None:
        known = ", ".join(sorted(Base.metadata.tables.keys()))
        raise KeyError(f"Unknown table {name!r}. Known tables: {known}")

    columns = [
        ColumnInfo(
            name=col.name,
            type=str(col.type),
            nullable=bool(col.nullable),
            primary_key=col.primary_key,
            foreign_keys=[str(fk.target_fullname) for fk in col.foreign_keys],
        )
        for col in table.columns
    ]
    indexes = [
        {"name": idx.name, "columns": [c.name for c in idx.columns], "unique": bool(idx.unique)}
        for idx in table.indexes
    ]
    referenced_by = sorted(
        {
            fk.parent.table.name
            for other in Base.metadata.tables.values()
            for fk in other.foreign_keys
            if fk.column.table.name == name
        }
    )
    return {
        "table": name,
        "domain": _table_domains().get(name, "?"),
        "columns": columns,
        "indexes": indexes,
        "referenced_by": referenced_by,
    }


def sample_table(name: str, limit: int = 5):
    """A few rows from one table, via the guarded SQL runner."""
    if name not in Base.metadata.tables:
        raise KeyError(f"Unknown table {name!r}.")
    return run_sql(f"SELECT * FROM {name} LIMIT {int(limit)}", limit=limit)
