"""Load a multi-sheet xlsx workbook into the hospital schema.

Each worksheet is mapped to the SQLAlchemy table whose ``__tablename__``
matches the sheet name. The header row defines the column subset to
populate; blank cells become ``NULL``. Sheets are ordered to respect
foreign-key dependencies.
"""

from __future__ import annotations

import json
from datetime import date, datetime, time
from pathlib import Path
from typing import Any

import pandas as pd
from sqlalchemy import inspect
from sqlalchemy.orm import Session

from setup.models import Base


# Load order respects FK dependencies — see docs/schema.md
LOAD_ORDER: list[str] = [
    # 1. catalogs
    "icd10_codes",
    "cpt_codes",
    "loinc_codes",
    "snomed_codes",
    "rxnorm_concepts",
    # 2. independent organizational entities
    "locations",
    "specialties",
    "shifts",
    "appointment_types",
    "payers",
    "pharmacies",
    # 3. org hierarchy
    "departments",
    "units",
    "rooms",
    "beds",
    # 4. people
    "providers",
    "provider_specialties",
    "provider_licenses",
    # 5. patients core
    "patients",
    "patient_identifiers",
    "patient_addresses",
    "patient_contacts",
    "emergency_contacts",
    "patient_consents",
    # 6. plans & coverage
    "insurance_plans",
    "patient_coverages",
    "authorizations",
    # 7. scheduling
    "appointment_slots",
    "referrals",
    "appointments",
    "appointment_status_history",
    "waitlist_entries",
    "appointment_reminders",
    # 8. encounters
    "encounters",
    "bed_assignments",
    "encounter_diagnoses",
    "encounter_procedures",
    # 9. clinical
    "problem_list_entries",
    "allergies",
    "allergy_reactions",
    "vital_signs",
    "clinical_observations",
    "care_plans",
    "care_plan_goals",
    "medications",
    "prescriptions",
    "medication_administrations",
    "medication_reconciliations",
    # 10. labs / imaging
    "lab_orders",
    "lab_specimens",
    "lab_results",
    "imaging_orders",
    "imaging_studies",
    "imaging_reports",
    # 11. billing
    "claims",
    "claim_lines",
    "charges",
    "payments",
    "adjustments",
    "claim_denials",
    "claim_appeals",
    "patient_statements",
    # 12. communications
    "patient_message_threads",
    "patient_messages",
    "call_logs",
    "insurance_correspondences",
    "inter_provider_messages",
    # 13. operations
    "staff_schedules",
    "on_call_assignments",
    "pharmacy_inventory",
    "equipment",
    "tasks",
    "audit_logs_summary",
]


def _coerce(value: Any, column_type: Any) -> Any:
    """Best-effort cell → Python value coercion driven by SQLAlchemy column type."""
    if value is None or (isinstance(value, float) and pd.isna(value)) or value == "":
        return None
    py_type = getattr(column_type, "python_type", None)
    try:
        if py_type is bool:
            if isinstance(value, str):
                return value.strip().lower() in {"true", "t", "1", "yes", "y"}
            return bool(value)
        if py_type is int:
            return int(value)
        if py_type is float:
            return float(value)
        if py_type is datetime:
            if isinstance(value, datetime):
                return value
            if isinstance(value, date):
                return datetime.combine(value, time.min)
            return datetime.fromisoformat(str(value).replace("Z", "+00:00"))
        if py_type is date:
            if isinstance(value, datetime):
                return value.date()
            if isinstance(value, date):
                return value
            return date.fromisoformat(str(value))
        if py_type is time:
            if isinstance(value, time):
                return value
            return time.fromisoformat(str(value))
        if py_type is dict or py_type is list:
            if isinstance(value, str):
                return json.loads(value)
            return value
    except (ValueError, TypeError):
        # fall through and return as-is; SQLAlchemy may still bind it
        return value
    return value


def _rows_for_sheet(df: pd.DataFrame, table) -> list[dict[str, Any]]:
    """Convert a DataFrame for one sheet into kwargs dicts for the SQLAlchemy table."""
    columns = {c.name: c for c in table.columns}
    rows: list[dict[str, Any]] = []
    for _, raw in df.iterrows():
        row: dict[str, Any] = {}
        for col_name in df.columns:
            col = columns.get(col_name)
            if col is None:
                continue
            row[col_name] = _coerce(raw[col_name], col.type)
        rows.append(row)
    return rows


def _collect_sheet_index(path: Path) -> dict[str, tuple[Path, pd.ExcelFile]]:
    """Build a {sheet_name: (file_path, ExcelFile)} index.

    If ``path`` is a file, only that workbook is opened. If it's a directory,
    every ``*.xlsx`` file in the directory (recursively) is opened. If multiple
    workbooks contain the same sheet name, the alphabetically-first file wins
    and the others are ignored for that sheet.
    """
    if path.is_dir():
        files = sorted(p for p in path.rglob("*.xlsx") if not p.name.startswith("~$"))
    else:
        files = [path]
    if not files:
        raise FileNotFoundError(f"No xlsx files found at {path}")

    index: dict[str, tuple[Path, pd.ExcelFile]] = {}
    for f in files:
        xls = pd.ExcelFile(f)
        for sheet in xls.sheet_names:
            index.setdefault(sheet, (f, xls))
    return index


def load_workbook(session: Session, path: str | Path) -> dict[str, int]:
    """Load an xlsx file (or directory of xlsx files) into the hospital database.

    Returns a ``{table_name: row_count}`` dict.
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(path)

    sheet_index = _collect_sheet_index(path)
    available = set(sheet_index.keys())
    counts: dict[str, int] = {}

    tables_by_name = {t.name: t for t in Base.metadata.sorted_tables}

    def _load(table_name: str) -> None:
        if table_name not in available or table_name not in tables_by_name:
            return
        _, xls = sheet_index[table_name]
        df = xls.parse(table_name, dtype=object)
        if df.empty:
            counts[table_name] = 0
            return
        df = df.where(df.notna(), None)
        rows = _rows_for_sheet(df, tables_by_name[table_name])
        if not rows:
            counts[table_name] = 0
            return
        session.execute(tables_by_name[table_name].insert(), rows)
        counts[table_name] = len(rows)

    for table_name in LOAD_ORDER:
        _load(table_name)

    leftover = sorted(available - set(LOAD_ORDER))
    for table_name in leftover:
        _load(table_name)

    session.commit()
    return counts


def list_expected_sheets() -> list[str]:
    """All sheet names the loader knows how to populate, in load order."""
    return list(LOAD_ORDER)


def diff_workbook(path: str | Path) -> dict[str, list[str]]:
    """Report missing / extra sheets compared to the canonical load order.

    Accepts either a single xlsx file or a directory of xlsx files.
    """
    path = Path(path)
    sheet_index = _collect_sheet_index(path)
    available = set(sheet_index.keys())
    expected = set(LOAD_ORDER)
    return {
        "missing_sheets": sorted(expected - available),
        "extra_sheets": sorted(available - expected),
    }


# Tiny convenience accessor for the inspector (kept for completeness)
def _inspect_db_metadata():
    return inspect(Base.metadata)
