"""Insurance & billing models (payers, plans, coverage, auths, claims, payments)."""

from __future__ import annotations

import uuid
from datetime import date, datetime

from sqlalchemy import Boolean, Date, DateTime, ForeignKey, Integer, Numeric, String, Text
from sqlalchemy.dialects.postgresql import UUID
from sqlalchemy.orm import Mapped, mapped_column, relationship

from setup.models.base import Base, TimestampMixin


class Payer(Base, TimestampMixin):
    __tablename__ = "payers"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    name: Mapped[str] = mapped_column(String(255), nullable=False)
    payer_type: Mapped[str] = mapped_column(String(32), nullable=False)
    payer_id_external: Mapped[str | None] = mapped_column(String(64))
    address_line1: Mapped[str | None] = mapped_column(String(255))
    city: Mapped[str | None] = mapped_column(String(120))
    state: Mapped[str | None] = mapped_column(String(8))
    postal_code: Mapped[str | None] = mapped_column(String(16))
    phone: Mapped[str | None] = mapped_column(String(32))
    claims_phone: Mapped[str | None] = mapped_column(String(32))
    claims_fax: Mapped[str | None] = mapped_column(String(32))
    electronic_claims_supported: Mapped[bool] = mapped_column(Boolean, default=True, nullable=False)


class InsurancePlan(Base, TimestampMixin):
    __tablename__ = "insurance_plans"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    payer_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("payers.id", ondelete="CASCADE"), nullable=False
    )
    plan_name: Mapped[str] = mapped_column(String(255), nullable=False)
    plan_type: Mapped[str | None] = mapped_column(String(32))
    metal_tier: Mapped[str | None] = mapped_column(String(16))
    group_number: Mapped[str | None] = mapped_column(String(64))
    effective_date: Mapped[date | None] = mapped_column(Date)
    termination_date: Mapped[date | None] = mapped_column(Date)
    requires_referrals: Mapped[bool] = mapped_column(Boolean, default=False, nullable=False)
    requires_prior_auth_for_imaging: Mapped[bool] = mapped_column(
        Boolean, default=False, nullable=False
    )


class PatientCoverage(Base, TimestampMixin):
    __tablename__ = "patient_coverages"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    patient_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("patients.id", ondelete="CASCADE"), nullable=False, index=True
    )
    insurance_plan_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("insurance_plans.id", ondelete="RESTRICT"), nullable=False
    )
    subscriber_relationship: Mapped[str] = mapped_column(String(16), nullable=False)
    subscriber_name: Mapped[str | None] = mapped_column(String(255))
    subscriber_dob: Mapped[date | None] = mapped_column(Date)
    member_id: Mapped[str] = mapped_column(String(64), nullable=False)
    group_number: Mapped[str | None] = mapped_column(String(64))
    coverage_rank: Mapped[int] = mapped_column(Integer, nullable=False, default=1)
    effective_date: Mapped[date] = mapped_column(Date, nullable=False)
    termination_date: Mapped[date | None] = mapped_column(Date)
    copay_pcp: Mapped[float | None] = mapped_column(Numeric(8, 2))
    copay_specialist: Mapped[float | None] = mapped_column(Numeric(8, 2))
    copay_er: Mapped[float | None] = mapped_column(Numeric(8, 2))
    deductible_individual: Mapped[float | None] = mapped_column(Numeric(10, 2))
    deductible_family: Mapped[float | None] = mapped_column(Numeric(10, 2))
    oop_max_individual: Mapped[float | None] = mapped_column(Numeric(10, 2))
    verified_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))


class Authorization(Base, TimestampMixin):
    __tablename__ = "authorizations"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    patient_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("patients.id", ondelete="CASCADE"), nullable=False
    )
    coverage_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("patient_coverages.id", ondelete="CASCADE"), nullable=False
    )
    auth_number: Mapped[str | None] = mapped_column(String(64), unique=True)
    cpt_code: Mapped[str | None] = mapped_column(
        String(16), ForeignKey("cpt_codes.code", ondelete="SET NULL")
    )
    requested_units: Mapped[int | None] = mapped_column(Integer)
    approved_units: Mapped[int | None] = mapped_column(Integer)
    status: Mapped[str] = mapped_column(String(16), nullable=False)
    effective_date: Mapped[date | None] = mapped_column(Date)
    expiration_date: Mapped[date | None] = mapped_column(Date)
    requested_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))
    decided_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))
    denial_reason: Mapped[str | None] = mapped_column(Text)


class Claim(Base, TimestampMixin):
    __tablename__ = "claims"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    patient_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("patients.id", ondelete="CASCADE"), nullable=False, index=True
    )
    encounter_id: Mapped[uuid.UUID | None] = mapped_column(
        UUID(as_uuid=True), ForeignKey("encounters.id", ondelete="SET NULL")
    )
    coverage_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("patient_coverages.id", ondelete="RESTRICT"), nullable=False
    )
    claim_number: Mapped[str] = mapped_column(String(64), unique=True, nullable=False)
    payer_claim_id: Mapped[str | None] = mapped_column(String(64))
    claim_type: Mapped[str] = mapped_column(String(32), nullable=False)
    status: Mapped[str] = mapped_column(String(32), nullable=False, index=True)
    service_start_date: Mapped[date] = mapped_column(Date, nullable=False)
    service_end_date: Mapped[date | None] = mapped_column(Date)
    total_charge: Mapped[float] = mapped_column(Numeric(12, 2), nullable=False)
    total_allowed: Mapped[float | None] = mapped_column(Numeric(12, 2))
    total_paid: Mapped[float | None] = mapped_column(Numeric(12, 2))
    patient_responsibility: Mapped[float | None] = mapped_column(Numeric(12, 2))
    submitted_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))
    paid_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))
    billing_provider_id: Mapped[uuid.UUID | None] = mapped_column(
        UUID(as_uuid=True), ForeignKey("providers.id", ondelete="SET NULL")
    )

    lines: Mapped[list["ClaimLine"]] = relationship(
        back_populates="claim", cascade="all, delete-orphan"
    )


class ClaimLine(Base, TimestampMixin):
    __tablename__ = "claim_lines"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    claim_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("claims.id", ondelete="CASCADE"), nullable=False
    )
    line_number: Mapped[int] = mapped_column(Integer, nullable=False)
    cpt_code: Mapped[str] = mapped_column(
        String(16), ForeignKey("cpt_codes.code", ondelete="RESTRICT"), nullable=False
    )
    modifier: Mapped[str | None] = mapped_column(String(64))
    icd10_pointer: Mapped[str | None] = mapped_column(String(64))
    service_date: Mapped[date] = mapped_column(Date, nullable=False)
    units: Mapped[float] = mapped_column(Numeric(8, 2), nullable=False, default=1)
    charge_amount: Mapped[float] = mapped_column(Numeric(12, 2), nullable=False)
    allowed_amount: Mapped[float | None] = mapped_column(Numeric(12, 2))
    paid_amount: Mapped[float | None] = mapped_column(Numeric(12, 2))
    adjustment_amount: Mapped[float | None] = mapped_column(Numeric(12, 2))
    denial_code: Mapped[str | None] = mapped_column(String(16))
    place_of_service: Mapped[str | None] = mapped_column(String(8))
    rendering_provider_id: Mapped[uuid.UUID | None] = mapped_column(
        UUID(as_uuid=True), ForeignKey("providers.id", ondelete="SET NULL")
    )

    claim: Mapped["Claim"] = relationship(back_populates="lines")


class Charge(Base, TimestampMixin):
    __tablename__ = "charges"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    encounter_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("encounters.id", ondelete="CASCADE"), nullable=False
    )
    patient_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("patients.id", ondelete="CASCADE"), nullable=False
    )
    cpt_code: Mapped[str] = mapped_column(
        String(16), ForeignKey("cpt_codes.code", ondelete="RESTRICT"), nullable=False
    )
    quantity: Mapped[float] = mapped_column(Numeric(8, 2), default=1, nullable=False)
    charge_amount: Mapped[float] = mapped_column(Numeric(12, 2), nullable=False)
    posted_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)
    posted_by_provider_id: Mapped[uuid.UUID | None] = mapped_column(
        UUID(as_uuid=True), ForeignKey("providers.id", ondelete="SET NULL")
    )
    claim_id: Mapped[uuid.UUID | None] = mapped_column(
        UUID(as_uuid=True), ForeignKey("claims.id", ondelete="SET NULL")
    )


class Payment(Base, TimestampMixin):
    __tablename__ = "payments"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    claim_id: Mapped[uuid.UUID | None] = mapped_column(
        UUID(as_uuid=True), ForeignKey("claims.id", ondelete="SET NULL")
    )
    patient_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("patients.id", ondelete="CASCADE"), nullable=False
    )
    payer_id: Mapped[uuid.UUID | None] = mapped_column(
        UUID(as_uuid=True), ForeignKey("payers.id", ondelete="SET NULL")
    )
    payment_type: Mapped[str] = mapped_column(String(16), nullable=False)
    payment_method: Mapped[str | None] = mapped_column(String(16))
    amount: Mapped[float] = mapped_column(Numeric(12, 2), nullable=False)
    received_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)
    reference_number: Mapped[str | None] = mapped_column(String(64))
    era_835_id: Mapped[str | None] = mapped_column(String(64))


class Adjustment(Base, TimestampMixin):
    __tablename__ = "adjustments"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    claim_line_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("claim_lines.id", ondelete="CASCADE"), nullable=False
    )
    adjustment_group: Mapped[str | None] = mapped_column(String(8))
    reason_code: Mapped[str] = mapped_column(String(16), nullable=False)
    amount: Mapped[float] = mapped_column(Numeric(12, 2), nullable=False)
    note: Mapped[str | None] = mapped_column(Text)


class PatientStatement(Base, TimestampMixin):
    __tablename__ = "patient_statements"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    patient_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("patients.id", ondelete="CASCADE"), nullable=False, index=True
    )
    statement_date: Mapped[date] = mapped_column(Date, nullable=False)
    period_start: Mapped[date | None] = mapped_column(Date)
    period_end: Mapped[date | None] = mapped_column(Date)
    previous_balance: Mapped[float | None] = mapped_column(Numeric(12, 2))
    charges_total: Mapped[float | None] = mapped_column(Numeric(12, 2))
    payments_total: Mapped[float | None] = mapped_column(Numeric(12, 2))
    adjustments_total: Mapped[float | None] = mapped_column(Numeric(12, 2))
    current_balance: Mapped[float] = mapped_column(Numeric(12, 2), nullable=False)
    due_date: Mapped[date | None] = mapped_column(Date)
    status: Mapped[str | None] = mapped_column(String(32))
    delivery_method: Mapped[str | None] = mapped_column(String(16))
