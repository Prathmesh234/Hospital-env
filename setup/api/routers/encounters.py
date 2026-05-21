"""Encounter endpoints (read-mostly)."""

from __future__ import annotations

import uuid

from fastapi import APIRouter, Depends, HTTPException, Query
from sqlalchemy import select
from sqlalchemy.orm import Session

from setup.db.postgres import get_session
from setup.models import Encounter
from setup.schemas.clinical import EncounterRead

router = APIRouter(prefix="/encounters", tags=["encounters"])


@router.get("", response_model=list[EncounterRead])
def list_encounters(
    session: Session = Depends(get_session),
    patient_id: uuid.UUID | None = None,
    status: str | None = None,
    encounter_class: str | None = None,
    limit: int = Query(50, ge=1, le=500),
) -> list[Encounter]:
    stmt = select(Encounter).order_by(Encounter.admitted_at.desc().nullslast())
    if patient_id:
        stmt = stmt.where(Encounter.patient_id == patient_id)
    if status:
        stmt = stmt.where(Encounter.status == status)
    if encounter_class:
        stmt = stmt.where(Encounter.encounter_class == encounter_class)
    return list(session.scalars(stmt.limit(limit)))


@router.get("/{encounter_id}", response_model=EncounterRead)
def get_encounter(encounter_id: uuid.UUID, session: Session = Depends(get_session)) -> Encounter:
    enc = session.get(Encounter, encounter_id)
    if enc is None:
        raise HTTPException(status_code=404, detail="encounter not found")
    return enc
