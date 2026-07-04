"""Deterministic MongoDB document seeding, derived from the PostgreSQL rows.

The relational workbook is the single source of truth; this module projects
document-shaped satellites off it (clinical notes, HIPAA audit events,
imaging metadata, IoT vitals streams). Everything is generated from a fixed
RNG seed over deterministically-ordered inputs, so repeated runs produce the
same documents and eval gold answers stay stable.
"""

from __future__ import annotations

import random
from datetime import datetime, timedelta
from typing import Any

from pymongo.database import Database
from sqlalchemy import text
from sqlalchemy.engine import Engine

RNG_SEED = 2026

_NOTE_TEMPLATES = {
    "inpatient": (
        "progress_note",
        "Hospital day {day}. Patient {name} remains admitted for {reason}. "
        "Overnight events reviewed; vitals stable. Continue current management, "
        "monitor response to therapy, reassess disposition in the morning.",
    ),
    "emergency": (
        "ed_triage_note",
        "ED triage: {name} presents with {reason}. Acuity assigned, initial "
        "workup ordered including labs and imaging as indicated. Reassess after results.",
    ),
    "ambulatory": (
        "clinic_note",
        "Clinic visit for {reason}. {name} seen in consultation; history reviewed, "
        "examination performed. Plan discussed with patient, follow-up arranged.",
    ),
    "virtual": (
        "telehealth_note",
        "Telehealth encounter regarding {reason}. {name} interviewed by video; "
        "medication adherence reviewed, no red-flag symptoms elicited.",
    ),
}

_AUDIT_ACTIONS = ["view", "view", "view", "view", "update", "print", "export"]
_AUDIT_RESOURCES = ["chart", "lab_result", "medication_list", "insurance", "demographics"]

_VITALS_METRICS = [
    ("heart_rate", 62, 104, "bpm"),
    ("spo2", 90, 100, "%"),
    ("respiratory_rate", 12, 24, "breaths/min"),
]


def _rows(engine: Engine, sql: str) -> list[dict[str, Any]]:
    with engine.connect() as conn:
        result = conn.execute(text(sql))
        return [dict(zip(result.keys(), row, strict=True)) for row in result]


def _seed_clinical_notes(engine: Engine, db: Database, rng: random.Random) -> int:
    encounters = _rows(
        engine,
        """
        SELECT e.id AS encounter_id, e.patient_id, e.encounter_class, e.chief_complaint,
               e.admitted_at, e.attending_provider_id,
               p.first_name || ' ' || p.last_name AS patient_name
        FROM encounters e JOIN patients p ON p.id = e.patient_id
        WHERE e.admitted_at IS NOT NULL
        ORDER BY e.admitted_at, e.id
        """,
    )
    docs: list[dict[str, Any]] = []
    for enc in encounters:
        note_type, template = _NOTE_TEMPLATES.get(
            enc["encounter_class"], _NOTE_TEMPLATES["ambulatory"]
        )
        n_notes = 2 if enc["encounter_class"] == "inpatient" else 1
        for day in range(1, n_notes + 1):
            written_at = enc["admitted_at"] + timedelta(hours=6 + 24 * (day - 1))
            docs.append(
                {
                    "patient_id": str(enc["patient_id"]),
                    "encounter_id": str(enc["encounter_id"]),
                    "author_provider_id": str(enc["attending_provider_id"])
                    if enc["attending_provider_id"]
                    else None,
                    "note_type": note_type,
                    "written_at": written_at,
                    "free_text": template.format(
                        day=day,
                        name=enc["patient_name"],
                        reason=(enc["chief_complaint"] or "ongoing care").lower(),
                    ),
                    "cosigned": rng.random() < 0.4,
                }
            )
    if docs:
        db["clinical_notes"].insert_many(docs)
    return len(docs)


def _seed_audit_logs(engine: Engine, db: Database, rng: random.Random) -> int:
    patients = _rows(engine, "SELECT id FROM patients ORDER BY mrn")
    providers = _rows(engine, "SELECT id FROM providers ORDER BY npi")
    base = datetime(2025, 1, 6, 7, 30)
    docs = []
    for i in range(60):
        actor = providers[rng.randrange(len(providers))]
        patient = patients[rng.randrange(len(patients))]
        docs.append(
            {
                "event_id": f"AUD-2025-{i + 1:04d}",
                "actor_provider_id": str(actor["id"]),
                "patient_id": str(patient["id"]),
                "action": rng.choice(_AUDIT_ACTIONS),
                "resource_type": rng.choice(_AUDIT_RESOURCES),
                "occurred_at": base + timedelta(minutes=37 * i + rng.randrange(20)),
                "source_ip": f"10.20.{rng.randrange(1, 5)}.{rng.randrange(2, 250)}",
                "success": rng.random() > 0.05,
            }
        )
    db["audit_logs"].insert_many(docs)
    return len(docs)


def _seed_imaging_metadata(engine: Engine, db: Database, rng: random.Random) -> int:
    studies = _rows(
        engine,
        """
        SELECT s.id AS imaging_study_id, s.study_uid, s.accession_number, s.performed_at,
               s.series_count, s.image_count,
               o.patient_id, o.modality, o.body_part
        FROM imaging_studies s JOIN imaging_orders o ON o.id = s.imaging_order_id
        ORDER BY s.accession_number
        """,
    )
    docs = []
    for study in studies:
        series_count = study["series_count"] or rng.randrange(1, 4)
        docs.append(
            {
                "imaging_study_id": str(study["imaging_study_id"]),
                "patient_id": str(study["patient_id"]),
                "study_uid": study["study_uid"],
                "accession_number": study["accession_number"],
                "modality": study["modality"],
                "body_part": study["body_part"],
                "study_date": study["performed_at"],
                "station": f"{study['modality']}-STATION-{rng.randrange(1, 4)}",
                "series": [
                    {
                        "series_number": s + 1,
                        "instance_count": rng.randrange(40, 320),
                        "description": f"{study['body_part']} series {s + 1}",
                    }
                    for s in range(series_count)
                ],
            }
        )
    if docs:
        db["imaging_metadata"].insert_many(docs)
    return len(docs)


def _seed_vitals_streams(engine: Engine, db: Database, rng: random.Random) -> int:
    monitored = _rows(
        engine,
        """
        SELECT e.id AS encounter_id, e.patient_id, e.admitted_at
        FROM encounters e
        WHERE e.encounter_class IN ('inpatient', 'emergency')
          AND e.admitted_at IS NOT NULL
        ORDER BY e.admitted_at, e.id
        """,
    )
    docs = []
    for enc_idx, enc in enumerate(monitored):
        for metric, low, high, unit in _VITALS_METRICS:
            started = enc["admitted_at"] + timedelta(hours=1)
            samples = [
                {"offset_s": 900 * n, "value": rng.randrange(low, high + 1)} for n in range(24)
            ]
            docs.append(
                {
                    "patient_id": str(enc["patient_id"]),
                    "encounter_id": str(enc["encounter_id"]),
                    "device_id": f"MON-{enc_idx + 1:02d}",
                    "metric": metric,
                    "unit": unit,
                    "started_at": started,
                    "interval_s": 900,
                    "samples": samples,
                }
            )
    if docs:
        db["vitals_streams"].insert_many(docs)
    return len(docs)


def seed_documents(engine: Engine, db: Database, drop_first: bool = True) -> dict[str, int]:
    """Populate all four Mongo collections. Returns ``{collection: doc_count}``."""
    rng = random.Random(RNG_SEED)
    if drop_first:
        for name in ("clinical_notes", "audit_logs", "imaging_metadata", "vitals_streams"):
            db[name].delete_many({})
    return {
        "clinical_notes": _seed_clinical_notes(engine, db, rng),
        "audit_logs": _seed_audit_logs(engine, db, rng),
        "imaging_metadata": _seed_imaging_metadata(engine, db, rng),
        "vitals_streams": _seed_vitals_streams(engine, db, rng),
    }
