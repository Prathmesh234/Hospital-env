"""Medications, prescriptions, MAR, reconciliations, external pharmacies."""

from __future__ import annotations

import uuid
from datetime import date, datetime

from sqlalchemy import Boolean, Date, DateTime, ForeignKey, Integer, Numeric, String, Text
from sqlalchemy.dialects.postgresql import UUID
from sqlalchemy.orm import Mapped, mapped_column

from setup.models.base import Base, TimestampMixin


class RxNormConcept(Base, TimestampMixin):
    __tablename__ = "rxnorm_concepts"

    rxcui: Mapped[str] = mapped_column(String(16), primary_key=True)
    name: Mapped[str] = mapped_column(Text, nullable=False)
    tty: Mapped[str | None] = mapped_column(String(8))
    route: Mapped[str | None] = mapped_column(String(16))
    is_controlled: Mapped[bool] = mapped_column(Boolean, default=False, nullable=False)
    dea_schedule: Mapped[str | None] = mapped_column(String(8))


class Medication(Base, TimestampMixin):
    __tablename__ = "medications"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    rxcui: Mapped[str | None] = mapped_column(
        String(16), ForeignKey("rxnorm_concepts.rxcui", ondelete="SET NULL")
    )
    ndc: Mapped[str | None] = mapped_column(String(16))
    name: Mapped[str] = mapped_column(String(255), nullable=False)
    strength: Mapped[str | None] = mapped_column(String(64))
    dosage_form: Mapped[str | None] = mapped_column(String(32))
    route: Mapped[str | None] = mapped_column(String(16))
    manufacturer: Mapped[str | None] = mapped_column(String(128))
    is_on_formulary: Mapped[bool] = mapped_column(Boolean, default=True, nullable=False)


class Pharmacy(Base, TimestampMixin):
    __tablename__ = "pharmacies"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    name: Mapped[str] = mapped_column(String(255), nullable=False)
    ncpdp_id: Mapped[str | None] = mapped_column(String(16))
    npi: Mapped[str | None] = mapped_column(String(10))
    phone: Mapped[str | None] = mapped_column(String(32))
    fax: Mapped[str | None] = mapped_column(String(32))
    address_line1: Mapped[str | None] = mapped_column(String(255))
    city: Mapped[str | None] = mapped_column(String(120))
    state: Mapped[str | None] = mapped_column(String(8))
    postal_code: Mapped[str | None] = mapped_column(String(16))
    is_mail_order: Mapped[bool] = mapped_column(Boolean, default=False, nullable=False)


class Prescription(Base, TimestampMixin):
    __tablename__ = "prescriptions"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    patient_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("patients.id", ondelete="CASCADE"), nullable=False, index=True
    )
    prescriber_provider_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("providers.id", ondelete="RESTRICT"), nullable=False
    )
    encounter_id: Mapped[uuid.UUID | None] = mapped_column(
        UUID(as_uuid=True), ForeignKey("encounters.id", ondelete="SET NULL")
    )
    medication_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("medications.id", ondelete="RESTRICT"), nullable=False
    )
    dose: Mapped[str] = mapped_column(String(64), nullable=False)
    route: Mapped[str] = mapped_column(String(16), nullable=False)
    frequency: Mapped[str] = mapped_column(String(32), nullable=False)
    duration_days: Mapped[int | None] = mapped_column(Integer)
    quantity: Mapped[float | None] = mapped_column(Numeric(10, 2))
    refills: Mapped[int] = mapped_column(Integer, default=0, nullable=False)
    start_date: Mapped[date] = mapped_column(Date, nullable=False)
    end_date: Mapped[date | None] = mapped_column(Date)
    status: Mapped[str] = mapped_column(String(16), nullable=False)
    indication: Mapped[str | None] = mapped_column(Text)
    prn_reason: Mapped[str | None] = mapped_column(Text)
    pharmacy_id: Mapped[uuid.UUID | None] = mapped_column(
        UUID(as_uuid=True), ForeignKey("pharmacies.id", ondelete="SET NULL")
    )
    is_electronic: Mapped[bool] = mapped_column(Boolean, default=True, nullable=False)


class MedicationAdministration(Base, TimestampMixin):
    __tablename__ = "medication_administrations"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    prescription_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("prescriptions.id", ondelete="CASCADE"), nullable=False
    )
    patient_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("patients.id", ondelete="CASCADE"), nullable=False, index=True
    )
    encounter_id: Mapped[uuid.UUID | None] = mapped_column(
        UUID(as_uuid=True), ForeignKey("encounters.id", ondelete="SET NULL")
    )
    administered_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)
    administered_by_provider_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("providers.id", ondelete="RESTRICT"), nullable=False
    )
    dose_given: Mapped[str] = mapped_column(String(64), nullable=False)
    route: Mapped[str | None] = mapped_column(String(16))
    status: Mapped[str] = mapped_column(String(16), nullable=False)
    not_done_reason: Mapped[str | None] = mapped_column(String(32))
    site: Mapped[str | None] = mapped_column(String(64))


class MedicationReconciliation(Base, TimestampMixin):
    __tablename__ = "medication_reconciliations"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    patient_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("patients.id", ondelete="CASCADE"), nullable=False
    )
    encounter_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("encounters.id", ondelete="CASCADE"), nullable=False
    )
    reconciliation_type: Mapped[str | None] = mapped_column(String(16))
    performed_by_provider_id: Mapped[uuid.UUID | None] = mapped_column(
        UUID(as_uuid=True), ForeignKey("providers.id", ondelete="SET NULL")
    )
    performed_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)
    notes: Mapped[str | None] = mapped_column(Text)
