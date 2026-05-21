"""MongoDB collection bootstrapping (indexes + JSON schema validators).

These collections complement the relational schema; see ``docs/schema.md``
for the document shapes.
"""

from __future__ import annotations

from typing import Any

from pymongo import ASCENDING, DESCENDING, TEXT
from pymongo.database import Database
from pymongo.errors import CollectionInvalid


COLLECTIONS: list[dict[str, Any]] = [
    {
        "name": "clinical_notes",
        "indexes": [
            ([("patient_id", ASCENDING), ("written_at", DESCENDING)], {}),
            ([("encounter_id", ASCENDING)], {}),
            ([("author_provider_id", ASCENDING)], {}),
            ([("free_text", TEXT)], {"default_language": "english"}),
        ],
    },
    {
        "name": "audit_logs",
        "indexes": [
            ([("patient_id", ASCENDING), ("occurred_at", DESCENDING)], {}),
            ([("actor_provider_id", ASCENDING), ("occurred_at", DESCENDING)], {}),
            ([("resource_type", ASCENDING), ("occurred_at", DESCENDING)], {}),
        ],
    },
    {
        "name": "imaging_metadata",
        "indexes": [
            ([("patient_id", ASCENDING), ("study_date", DESCENDING)], {}),
            ([("study_uid", ASCENDING)], {"unique": True}),
            ([("imaging_study_id", ASCENDING)], {}),
        ],
    },
    {
        "name": "vitals_streams",
        "indexes": [
            (
                [("patient_id", ASCENDING), ("metric", ASCENDING), ("started_at", DESCENDING)],
                {},
            ),
            ([("encounter_id", ASCENDING)], {}),
            ([("device_id", ASCENDING)], {}),
        ],
    },
]


def ensure_collections(db: Database) -> list[str]:
    """Create collections (if missing) and ensure indexes. Returns created names."""
    created: list[str] = []
    existing = set(db.list_collection_names())
    for spec in COLLECTIONS:
        name: str = spec["name"]
        if name not in existing:
            try:
                db.create_collection(name)
                created.append(name)
            except CollectionInvalid:
                pass
        coll = db[name]
        for keys, opts in spec["indexes"]:
            coll.create_index(keys, **opts)
    return created


def drop_collections(db: Database) -> list[str]:
    """Drop every Hospital-env-managed collection. Returns dropped names."""
    dropped: list[str] = []
    for spec in COLLECTIONS:
        name = spec["name"]
        if name in db.list_collection_names():
            db.drop_collection(name)
            dropped.append(name)
    return dropped
