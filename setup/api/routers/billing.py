"""Insurance & billing endpoints."""

from __future__ import annotations

import uuid

from fastapi import APIRouter, Depends, Query
from sqlalchemy import select
from sqlalchemy.orm import Session

from setup.db.postgres import get_session
from setup.models import Claim, PatientCoverage, Payer
from setup.schemas.clinical import ClaimRead
from setup.schemas.common import ORMModel

router = APIRouter(prefix="/billing", tags=["billing"])


class PayerRead(ORMModel):
    id: uuid.UUID
    name: str
    payer_type: str
    electronic_claims_supported: bool


class PatientCoverageRead(ORMModel):
    id: uuid.UUID
    patient_id: uuid.UUID
    insurance_plan_id: uuid.UUID
    subscriber_relationship: str
    member_id: str
    coverage_rank: int


@router.get("/payers", response_model=list[PayerRead])
def list_payers(session: Session = Depends(get_session)) -> list[Payer]:
    return list(session.scalars(select(Payer).order_by(Payer.name)))


@router.get("/patients/{patient_id}/coverages", response_model=list[PatientCoverageRead])
def patient_coverages(
    patient_id: uuid.UUID, session: Session = Depends(get_session)
) -> list[PatientCoverage]:
    stmt = (
        select(PatientCoverage)
        .where(PatientCoverage.patient_id == patient_id)
        .order_by(PatientCoverage.coverage_rank)
    )
    return list(session.scalars(stmt))


@router.get("/claims", response_model=list[ClaimRead])
def list_claims(
    session: Session = Depends(get_session),
    patient_id: uuid.UUID | None = None,
    status: str | None = None,
    limit: int = Query(50, ge=1, le=500),
) -> list[Claim]:
    stmt = select(Claim).order_by(Claim.service_start_date.desc())
    if patient_id:
        stmt = stmt.where(Claim.patient_id == patient_id)
    if status:
        stmt = stmt.where(Claim.status == status)
    return list(session.scalars(stmt.limit(limit)))
