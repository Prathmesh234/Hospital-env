"""Patient endpoints (read-mostly)."""

from __future__ import annotations

import uuid

from fastapi import APIRouter, Depends, HTTPException, Query
from sqlalchemy import select
from sqlalchemy.orm import Session

from setup.db.postgres import get_session
from setup.models import Patient
from setup.schemas.patients import PatientCreate, PatientRead

router = APIRouter(prefix="/patients", tags=["patients"])


@router.get("", response_model=list[PatientRead])
def list_patients(
    session: Session = Depends(get_session),
    limit: int = Query(50, ge=1, le=500),
    offset: int = Query(0, ge=0),
    mrn: str | None = None,
    last_name: str | None = None,
) -> list[Patient]:
    stmt = select(Patient).order_by(Patient.last_name, Patient.first_name)
    if mrn:
        stmt = stmt.where(Patient.mrn == mrn)
    if last_name:
        stmt = stmt.where(Patient.last_name.ilike(f"{last_name}%"))
    return list(session.scalars(stmt.limit(limit).offset(offset)))


@router.get("/{patient_id}", response_model=PatientRead)
def get_patient(patient_id: uuid.UUID, session: Session = Depends(get_session)) -> Patient:
    patient = session.get(Patient, patient_id)
    if patient is None:
        raise HTTPException(status_code=404, detail="patient not found")
    return patient


@router.post("", response_model=PatientRead, status_code=201)
def create_patient(payload: PatientCreate, session: Session = Depends(get_session)) -> Patient:
    patient = Patient(**payload.model_dump())
    session.add(patient)
    session.commit()
    session.refresh(patient)
    return patient
