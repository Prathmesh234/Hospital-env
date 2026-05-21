"""Provider + organizational hierarchy endpoints."""

from __future__ import annotations

import uuid

from fastapi import APIRouter, Depends, HTTPException, Query
from sqlalchemy import select
from sqlalchemy.orm import Session

from setup.db.postgres import get_session
from setup.models import Department, Location, Provider
from setup.schemas.common import ORMModel

router = APIRouter(prefix="/providers", tags=["providers"])


class ProviderRead(ORMModel):
    id: uuid.UUID
    npi: str | None = None
    first_name: str
    last_name: str
    credentials: str | None = None
    provider_type: str
    is_active: bool


class LocationRead(ORMModel):
    id: uuid.UUID
    name: str
    facility_type: str | None = None
    city: str | None = None
    state: str | None = None


class DepartmentRead(ORMModel):
    id: uuid.UUID
    name: str
    location_id: uuid.UUID
    department_type: str | None = None


@router.get("", response_model=list[ProviderRead])
def list_providers(
    session: Session = Depends(get_session),
    provider_type: str | None = None,
    is_active: bool | None = None,
    limit: int = Query(100, ge=1, le=500),
) -> list[Provider]:
    stmt = select(Provider).order_by(Provider.last_name)
    if provider_type:
        stmt = stmt.where(Provider.provider_type == provider_type)
    if is_active is not None:
        stmt = stmt.where(Provider.is_active == is_active)
    return list(session.scalars(stmt.limit(limit)))


@router.get("/{provider_id}", response_model=ProviderRead)
def get_provider(provider_id: uuid.UUID, session: Session = Depends(get_session)) -> Provider:
    provider = session.get(Provider, provider_id)
    if provider is None:
        raise HTTPException(status_code=404, detail="provider not found")
    return provider


@router.get("/locations/all", response_model=list[LocationRead])
def list_locations(session: Session = Depends(get_session)) -> list[Location]:
    return list(session.scalars(select(Location).order_by(Location.name)))


@router.get("/departments/all", response_model=list[DepartmentRead])
def list_departments(session: Session = Depends(get_session)) -> list[Department]:
    return list(session.scalars(select(Department).order_by(Department.name)))
