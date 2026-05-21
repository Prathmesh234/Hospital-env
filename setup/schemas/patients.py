"""Patient-related Pydantic schemas."""

from __future__ import annotations

import uuid
from datetime import date, datetime

from pydantic import Field

from setup.schemas.common import ORMModel


class PatientCreate(ORMModel):
    mrn: str
    first_name: str
    last_name: str
    date_of_birth: date
    middle_name: str | None = None
    prefix: str | None = None
    suffix: str | None = None
    sex_at_birth: str | None = None
    gender_identity: str | None = None
    pronouns: str | None = None
    race: str | None = None
    ethnicity: str | None = None
    preferred_language: str | None = None
    marital_status: str | None = None
    religion: str | None = None
    ssn_last4: str | None = Field(default=None, max_length=4)
    is_deceased: bool = False
    deceased_at: datetime | None = None
    primary_provider_id: uuid.UUID | None = None
    vip_status: str | None = "none"
    notes: str | None = None


class PatientRead(PatientCreate):
    id: uuid.UUID
    created_at: datetime
    updated_at: datetime


class PatientAddressRead(ORMModel):
    id: uuid.UUID
    patient_id: uuid.UUID
    address_use: str
    line1: str
    line2: str | None = None
    city: str
    state: str
    postal_code: str
    country: str
    is_primary: bool
    valid_from: date | None = None
    valid_to: date | None = None


class PatientContactRead(ORMModel):
    id: uuid.UUID
    patient_id: uuid.UUID
    contact_system: str
    contact_value: str
    contact_use: str | None = None
    is_primary: bool
    consent_to_contact: bool


class EmergencyContactRead(ORMModel):
    id: uuid.UUID
    patient_id: uuid.UUID
    name: str
    relationship_type: str = Field(alias="relationship")
    phone: str
    email: str | None = None
    priority_rank: int
    has_medical_poa: bool
