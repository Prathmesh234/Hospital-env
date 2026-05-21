"""Encounter, claim, lab order schemas (minimal read models)."""

from __future__ import annotations

import uuid
from datetime import date, datetime

from setup.schemas.common import ORMModel


class EncounterRead(ORMModel):
    id: uuid.UUID
    patient_id: uuid.UUID
    encounter_class: str
    status: str
    admission_type: str | None = None
    chief_complaint: str | None = None
    attending_provider_id: uuid.UUID | None = None
    location_id: uuid.UUID
    admitted_at: datetime | None = None
    discharged_at: datetime | None = None
    discharge_disposition: str | None = None
    triage_acuity: int | None = None


class ClaimRead(ORMModel):
    id: uuid.UUID
    patient_id: uuid.UUID
    encounter_id: uuid.UUID | None = None
    coverage_id: uuid.UUID
    claim_number: str
    claim_type: str
    status: str
    service_start_date: date
    service_end_date: date | None = None
    total_charge: float
    total_allowed: float | None = None
    total_paid: float | None = None
    patient_responsibility: float | None = None
    submitted_at: datetime | None = None
    paid_at: datetime | None = None


class LabOrderRead(ORMModel):
    id: uuid.UUID
    patient_id: uuid.UUID
    encounter_id: uuid.UUID | None = None
    ordering_provider_id: uuid.UUID
    ordered_at: datetime
    priority: str | None = None
    status: str
    panel_loinc_code: str | None = None
    fasting_required: bool


class LabResultRead(ORMModel):
    id: uuid.UUID
    lab_order_id: uuid.UUID
    loinc_code: str
    value_numeric: float | None = None
    value_text: str | None = None
    units: str | None = None
    reference_range_low: float | None = None
    reference_range_high: float | None = None
    interpretation: str | None = None
    is_critical: bool
    resulted_at: datetime
    status: str | None = None


class PrescriptionRead(ORMModel):
    id: uuid.UUID
    patient_id: uuid.UUID
    prescriber_provider_id: uuid.UUID
    medication_id: uuid.UUID
    dose: str
    route: str
    frequency: str
    start_date: date
    end_date: date | None = None
    status: str


class PatientMessageRead(ORMModel):
    id: uuid.UUID
    thread_id: uuid.UUID
    channel: str
    direction: str
    body: str
    sent_at: datetime
    read_at: datetime | None = None
