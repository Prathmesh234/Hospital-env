"""Clinical endpoints: labs, vitals, problem list."""

from __future__ import annotations

import uuid

from fastapi import APIRouter, Depends, Query
from sqlalchemy import select
from sqlalchemy.orm import Session

from setup.db.postgres import get_session
from setup.models import LabOrder, LabResult, ProblemListEntry, VitalSign
from setup.schemas.clinical import LabOrderRead, LabResultRead
from setup.schemas.common import ORMModel

router = APIRouter(prefix="/clinical", tags=["clinical"])


class VitalSignRead(ORMModel):
    id: uuid.UUID
    patient_id: uuid.UUID
    measured_at: object
    systolic_bp: int | None = None
    diastolic_bp: int | None = None
    heart_rate: int | None = None
    respiratory_rate: int | None = None
    temperature_c: float | None = None
    spo2: int | None = None
    pain_score: int | None = None


class ProblemListEntryRead(ORMModel):
    id: uuid.UUID
    patient_id: uuid.UUID
    description: str
    clinical_status: str
    severity: str | None = None
    icd10_code: str | None = None


@router.get("/patients/{patient_id}/vitals", response_model=list[VitalSignRead])
def patient_vitals(
    patient_id: uuid.UUID,
    session: Session = Depends(get_session),
    limit: int = Query(50, ge=1, le=500),
) -> list[VitalSign]:
    stmt = (
        select(VitalSign)
        .where(VitalSign.patient_id == patient_id)
        .order_by(VitalSign.measured_at.desc())
        .limit(limit)
    )
    return list(session.scalars(stmt))


@router.get("/patients/{patient_id}/problems", response_model=list[ProblemListEntryRead])
def patient_problems(
    patient_id: uuid.UUID, session: Session = Depends(get_session)
) -> list[ProblemListEntry]:
    stmt = select(ProblemListEntry).where(ProblemListEntry.patient_id == patient_id)
    return list(session.scalars(stmt))


@router.get("/patients/{patient_id}/lab-orders", response_model=list[LabOrderRead])
def patient_lab_orders(
    patient_id: uuid.UUID, session: Session = Depends(get_session)
) -> list[LabOrder]:
    stmt = (
        select(LabOrder)
        .where(LabOrder.patient_id == patient_id)
        .order_by(LabOrder.ordered_at.desc())
    )
    return list(session.scalars(stmt))


@router.get("/lab-orders/{order_id}/results", response_model=list[LabResultRead])
def lab_order_results(
    order_id: uuid.UUID, session: Session = Depends(get_session)
) -> list[LabResult]:
    stmt = (
        select(LabResult)
        .where(LabResult.lab_order_id == order_id)
        .order_by(LabResult.resulted_at)
    )
    return list(session.scalars(stmt))
