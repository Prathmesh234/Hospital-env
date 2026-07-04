"""Read-only MongoDB access for agents (the document side-store).

Only the four Hospital-env collections are reachable, and only via ``find``
/ ``count`` — no writes, no aggregation pipelines with ``$out``/``$merge``.
"""

from __future__ import annotations

import json
from datetime import date, datetime
from typing import Any

from bson import ObjectId

from setup.db.mongo import get_sync_db
from setup.nosql.collections import COLLECTIONS

ALLOWED_COLLECTIONS = {spec["name"] for spec in COLLECTIONS}
DEFAULT_LIMIT = 20
MAX_LIMIT = 200

_FORBIDDEN_OPERATORS = {"$where", "$function", "$accumulator", "$out", "$merge"}


class MongoGuardError(Exception):
    """Raised when a Mongo query is rejected."""


def _check_collection(name: str) -> None:
    if name not in ALLOWED_COLLECTIONS:
        raise MongoGuardError(
            f"Unknown collection {name!r}. Available: {', '.join(sorted(ALLOWED_COLLECTIONS))}"
        )


def _check_filter(node: Any) -> None:
    if isinstance(node, dict):
        for key, value in node.items():
            if key in _FORBIDDEN_OPERATORS:
                raise MongoGuardError(f"Operator {key!r} is not allowed.")
            _check_filter(value)
    elif isinstance(node, list):
        for item in node:
            _check_filter(item)


def _jsonable(value: Any) -> Any:
    if isinstance(value, ObjectId):
        return str(value)
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, dict):
        return {k: _jsonable(v) for k, v in value.items()}
    if isinstance(value, list):
        return [_jsonable(v) for v in value]
    return value


def parse_json_arg(raw: str | None, what: str) -> dict[str, Any] | None:
    if raw is None or raw.strip() == "":
        return None
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise MongoGuardError(f"Invalid JSON for {what}: {exc}") from exc
    if not isinstance(parsed, dict):
        raise MongoGuardError(f"{what} must be a JSON object.")
    return parsed


def collection_stats() -> list[dict[str, Any]]:
    db = get_sync_db()
    existing = set(db.list_collection_names())
    return [
        {
            "collection": name,
            "documents": db[name].count_documents({}) if name in existing else 0,
        }
        for name in sorted(ALLOWED_COLLECTIONS)
    ]


def find(
    collection: str,
    filter_: dict[str, Any] | None = None,
    projection: dict[str, Any] | None = None,
    sort: list[tuple[str, int]] | None = None,
    limit: int = DEFAULT_LIMIT,
) -> list[dict[str, Any]]:
    """Guarded ``find`` returning JSON-serializable documents."""
    _check_collection(collection)
    filter_ = filter_ or {}
    _check_filter(filter_)
    limit = max(1, min(int(limit), MAX_LIMIT))

    cursor = get_sync_db()[collection].find(filter_, projection)
    if sort:
        cursor = cursor.sort(sort)
    return [_jsonable(doc) for doc in cursor.limit(limit)]


def count(collection: str, filter_: dict[str, Any] | None = None) -> int:
    _check_collection(collection)
    filter_ = filter_ or {}
    _check_filter(filter_)
    return get_sync_db()[collection].count_documents(filter_)
