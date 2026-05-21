"""Patient + demographic ORM models."""

from __future__ import annotations

import uuid
from datetime import date, datetime
from typing import TYPE_CHECKING

from sqlalchemy import Boolean, CheckConstraint, Date, DateTime, ForeignKey, Integer, String, Text
from sqlalchemy.dialects.postgresql import UUID
from sqlalchemy.orm import Mapped, mapped_column, relationship

from setup.models.base import Base, TimestampMixin

if TYPE_CHECKING:
    from setup.models.providers import Provider


class Patient(Base, TimestampMixin):
    __tablename__ = "patients"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    mrn: Mapped[str] = mapped_column(String(64), unique=True, nullable=False, index=True)
    first_name: Mapped[str] = mapped_column(String(120), nullable=False)
    middle_name: Mapped[str | None] = mapped_column(String(120))
    last_name: Mapped[str] = mapped_column(String(120), nullable=False, index=True)
    prefix: Mapped[str | None] = mapped_column(String(16))
    suffix: Mapped[str | None] = mapped_column(String(16))
    date_of_birth: Mapped[date] = mapped_column(Date, nullable=False)
    sex_at_birth: Mapped[str | None] = mapped_column(String(16))
    gender_identity: Mapped[str | None] = mapped_column(String(64))
    pronouns: Mapped[str | None] = mapped_column(String(32))
    race: Mapped[str | None] = mapped_column(String(64))
    ethnicity: Mapped[str | None] = mapped_column(String(64))
    preferred_language: Mapped[str | None] = mapped_column(String(16))
    marital_status: Mapped[str | None] = mapped_column(String(32))
    religion: Mapped[str | None] = mapped_column(String(64))
    ssn_last4: Mapped[str | None] = mapped_column(String(4))
    is_deceased: Mapped[bool] = mapped_column(Boolean, default=False, nullable=False)
    deceased_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))
    primary_provider_id: Mapped[uuid.UUID | None] = mapped_column(
        UUID(as_uuid=True), ForeignKey("providers.id", ondelete="SET NULL")
    )
    vip_status: Mapped[str | None] = mapped_column(String(16), default="none")
    notes: Mapped[str | None] = mapped_column(Text)

    __table_args__ = (
        CheckConstraint(
            "sex_at_birth IN ('male','female','intersex','unknown')",
            name="sex_at_birth_allowed",
        ),
        CheckConstraint(
            "vip_status IN ('none','vip','restricted')",
            name="vip_status_allowed",
        ),
    )

    primary_provider: Mapped["Provider | None"] = relationship(
        "Provider", foreign_keys=[primary_provider_id]
    )
    identifiers: Mapped[list["PatientIdentifier"]] = relationship(
        back_populates="patient", cascade="all, delete-orphan"
    )
    addresses: Mapped[list["PatientAddress"]] = relationship(
        back_populates="patient", cascade="all, delete-orphan"
    )
    contacts: Mapped[list["PatientContact"]] = relationship(
        back_populates="patient", cascade="all, delete-orphan"
    )
    emergency_contacts: Mapped[list["EmergencyContact"]] = relationship(
        back_populates="patient", cascade="all, delete-orphan"
    )
    consents: Mapped[list["PatientConsent"]] = relationship(
        back_populates="patient", cascade="all, delete-orphan"
    )


class PatientIdentifier(Base, TimestampMixin):
    __tablename__ = "patient_identifiers"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    patient_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("patients.id", ondelete="CASCADE"), nullable=False
    )
    identifier_type: Mapped[str] = mapped_column(String(32), nullable=False)
    identifier_value: Mapped[str] = mapped_column(String(128), nullable=False)
    issuing_authority: Mapped[str | None] = mapped_column(String(128))
    valid_from: Mapped[date | None] = mapped_column(Date)
    valid_to: Mapped[date | None] = mapped_column(Date)

    patient: Mapped["Patient"] = relationship(back_populates="identifiers")


class PatientAddress(Base, TimestampMixin):
    __tablename__ = "patient_addresses"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    patient_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("patients.id", ondelete="CASCADE"), nullable=False
    )
    address_use: Mapped[str] = mapped_column(String(16), nullable=False)
    line1: Mapped[str] = mapped_column(String(255), nullable=False)
    line2: Mapped[str | None] = mapped_column(String(255))
    city: Mapped[str] = mapped_column(String(120), nullable=False)
    state: Mapped[str] = mapped_column(String(8), nullable=False)
    postal_code: Mapped[str] = mapped_column(String(16), nullable=False)
    country: Mapped[str] = mapped_column(String(8), default="US", nullable=False)
    is_primary: Mapped[bool] = mapped_column(Boolean, default=False, nullable=False)
    valid_from: Mapped[date | None] = mapped_column(Date)
    valid_to: Mapped[date | None] = mapped_column(Date)

    patient: Mapped["Patient"] = relationship(back_populates="addresses")


class PatientContact(Base, TimestampMixin):
    __tablename__ = "patient_contacts"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    patient_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("patients.id", ondelete="CASCADE"), nullable=False
    )
    contact_system: Mapped[str] = mapped_column(String(16), nullable=False)
    contact_value: Mapped[str] = mapped_column(String(255), nullable=False)
    contact_use: Mapped[str | None] = mapped_column(String(16))
    is_primary: Mapped[bool] = mapped_column(Boolean, default=False, nullable=False)
    consent_to_contact: Mapped[bool] = mapped_column(Boolean, default=True, nullable=False)

    patient: Mapped["Patient"] = relationship(back_populates="contacts")


class EmergencyContact(Base, TimestampMixin):
    __tablename__ = "emergency_contacts"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    patient_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("patients.id", ondelete="CASCADE"), nullable=False
    )
    name: Mapped[str] = mapped_column(String(255), nullable=False)
    relationship_type: Mapped[str] = mapped_column("relationship", String(32), nullable=False)
    phone: Mapped[str] = mapped_column(String(32), nullable=False)
    email: Mapped[str | None] = mapped_column(String(255))
    address_line1: Mapped[str | None] = mapped_column(String(255))
    city: Mapped[str | None] = mapped_column(String(120))
    state: Mapped[str | None] = mapped_column(String(8))
    postal_code: Mapped[str | None] = mapped_column(String(16))
    priority_rank: Mapped[int] = mapped_column(Integer, default=1, nullable=False)
    has_medical_poa: Mapped[bool] = mapped_column(Boolean, default=False, nullable=False)

    patient: Mapped["Patient"] = relationship(back_populates="emergency_contacts")


class PatientConsent(Base, TimestampMixin):
    __tablename__ = "patient_consents"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    patient_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("patients.id", ondelete="CASCADE"), nullable=False
    )
    consent_type: Mapped[str] = mapped_column(String(32), nullable=False)
    granted: Mapped[bool] = mapped_column(Boolean, nullable=False)
    granted_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)
    expires_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))
    witness_name: Mapped[str | None] = mapped_column(String(255))
    document_ref: Mapped[str | None] = mapped_column(Text)

    patient: Mapped["Patient"] = relationship(back_populates="consents")
