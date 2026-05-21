"""Medication endpoints (formulary + prescriptions + MAR)."""

from __future__ import annotations

import uuid

from fastapi import APIRouter, Depends, Query
from sqlalchemy import select
from sqlalchemy.orm import Session

from setup.db.postgres import get_session
from setup.models import Medication, MedicationAdministration, Prescription
from setup.schemas.clinical import PrescriptionRead
from setup.schemas.common import ORMModel

router = APIRouter(prefix="/medications", tags=["medications"])


class MedicationRead(ORMModel):
    id: uuid.UUID
    rxcui: str | None = None
    ndc: str | None = None
    name: str
    strength: str | None = None
    dosage_form: str | None = None
    route: str | None = None
    is_on_formulary: bool


class MedicationAdministrationRead(ORMModel):
    id: uuid.UUID
    prescription_id: uuid.UUID
    patient_id: uuid.UUID
    administered_at: object
    dose_given: str
    status: str


@router.get("/formulary", response_model=list[MedicationRead])
def list_formulary(
    session: Session = Depends(get_session),
    limit: int = Query(100, ge=1, le=500),
    on_formulary: bool | None = True,
) -> list[Medication]:
    stmt = select(Medication).order_by(Medication.name)
    if on_formulary is not None:
        stmt = stmt.where(Medication.is_on_formulary == on_formulary)
    return list(session.scalars(stmt.limit(limit)))


@router.get("/patients/{patient_id}/prescriptions", response_model=list[PrescriptionRead])
def patient_prescriptions(
    patient_id: uuid.UUID, session: Session = Depends(get_session)
) -> list[Prescription]:
    stmt = (
        select(Prescription)
        .where(Prescription.patient_id == patient_id)
        .order_by(Prescription.start_date.desc())
    )
    return list(session.scalars(stmt))


@router.get(
    "/prescriptions/{prescription_id}/administrations",
    response_model=list[MedicationAdministrationRead],
)
def prescription_mar(
    prescription_id: uuid.UUID, session: Session = Depends(get_session)
) -> list[MedicationAdministration]:
    stmt = (
        select(MedicationAdministration)
        .where(MedicationAdministration.prescription_id == prescription_id)
        .order_by(MedicationAdministration.administered_at)
    )
    return list(session.scalars(stmt))
