"""Clinical models: code catalogs, problem list, allergies, vitals, observations, care plans."""

from __future__ import annotations

import uuid
from datetime import date, datetime

from sqlalchemy import Boolean, Date, DateTime, ForeignKey, Integer, Numeric, String, Text
from sqlalchemy.dialects.postgresql import UUID
from sqlalchemy.orm import Mapped, mapped_column, relationship

from setup.models.base import Base, TimestampMixin


# ---- Code catalogs ----------------------------------------------------------


class ICD10Code(Base, TimestampMixin):
    __tablename__ = "icd10_codes"

    code: Mapped[str] = mapped_column(String(16), primary_key=True)
    description: Mapped[str] = mapped_column(Text, nullable=False)
    chapter: Mapped[str | None] = mapped_column(String(255))
    is_billable: Mapped[bool] = mapped_column(Boolean, default=True, nullable=False)


class CPTCode(Base, TimestampMixin):
    __tablename__ = "cpt_codes"

    code: Mapped[str] = mapped_column(String(16), primary_key=True)
    description: Mapped[str] = mapped_column(Text, nullable=False)
    category: Mapped[str | None] = mapped_column(String(32))
    default_charge: Mapped[float | None] = mapped_column(Numeric(12, 2))


class LOINCCode(Base, TimestampMixin):
    __tablename__ = "loinc_codes"

    code: Mapped[str] = mapped_column(String(16), primary_key=True)
    display: Mapped[str] = mapped_column(Text, nullable=False)
    loinc_class: Mapped[str | None] = mapped_column("class", String(32))
    system: Mapped[str | None] = mapped_column(String(32))
    units_default: Mapped[str | None] = mapped_column(String(32))


class SNOMEDCode(Base, TimestampMixin):
    __tablename__ = "snomed_codes"

    code: Mapped[str] = mapped_column(String(32), primary_key=True)
    display: Mapped[str] = mapped_column(Text, nullable=False)
    semantic_tag: Mapped[str | None] = mapped_column(String(64))


# ---- Patient-level clinical data --------------------------------------------


class ProblemListEntry(Base, TimestampMixin):
    __tablename__ = "problem_list_entries"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    patient_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("patients.id", ondelete="CASCADE"), nullable=False
    )
    icd10_code: Mapped[str | None] = mapped_column(
        String(16), ForeignKey("icd10_codes.code", ondelete="SET NULL")
    )
    snomed_code: Mapped[str | None] = mapped_column(
        String(32), ForeignKey("snomed_codes.code", ondelete="SET NULL")
    )
    description: Mapped[str] = mapped_column(Text, nullable=False)
    clinical_status: Mapped[str] = mapped_column(String(32), nullable=False)
    verification_status: Mapped[str | None] = mapped_column(String(32))
    severity: Mapped[str | None] = mapped_column(String(16))
    onset_date: Mapped[date | None] = mapped_column(Date)
    resolved_date: Mapped[date | None] = mapped_column(Date)
    recorded_by_provider_id: Mapped[uuid.UUID | None] = mapped_column(
        UUID(as_uuid=True), ForeignKey("providers.id", ondelete="SET NULL")
    )


class Allergy(Base, TimestampMixin):
    __tablename__ = "allergies"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    patient_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("patients.id", ondelete="CASCADE"), nullable=False
    )
    allergen_type: Mapped[str] = mapped_column(String(32), nullable=False)
    allergen_name: Mapped[str] = mapped_column(String(255), nullable=False)
    rxnorm_code: Mapped[str | None] = mapped_column(
        String(16), ForeignKey("rxnorm_concepts.rxcui", ondelete="SET NULL")
    )
    criticality: Mapped[str | None] = mapped_column(String(16))
    clinical_status: Mapped[str | None] = mapped_column(String(16))
    verification_status: Mapped[str | None] = mapped_column(String(16))
    recorded_date: Mapped[date | None] = mapped_column(Date)
    last_occurrence_date: Mapped[date | None] = mapped_column(Date)

    reactions: Mapped[list["AllergyReaction"]] = relationship(
        back_populates="allergy", cascade="all, delete-orphan"
    )


class AllergyReaction(Base, TimestampMixin):
    __tablename__ = "allergy_reactions"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    allergy_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("allergies.id", ondelete="CASCADE"), nullable=False
    )
    manifestation: Mapped[str] = mapped_column(String(64), nullable=False)
    severity: Mapped[str | None] = mapped_column(String(16))
    onset_minutes: Mapped[int | None] = mapped_column(Integer)

    allergy: Mapped["Allergy"] = relationship(back_populates="reactions")


class VitalSign(Base, TimestampMixin):
    __tablename__ = "vital_signs"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    patient_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("patients.id", ondelete="CASCADE"), nullable=False, index=True
    )
    encounter_id: Mapped[uuid.UUID | None] = mapped_column(
        UUID(as_uuid=True), ForeignKey("encounters.id", ondelete="SET NULL")
    )
    measured_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)
    systolic_bp: Mapped[int | None] = mapped_column(Integer)
    diastolic_bp: Mapped[int | None] = mapped_column(Integer)
    heart_rate: Mapped[int | None] = mapped_column(Integer)
    respiratory_rate: Mapped[int | None] = mapped_column(Integer)
    temperature_c: Mapped[float | None] = mapped_column(Numeric(4, 2))
    spo2: Mapped[int | None] = mapped_column(Integer)
    pain_score: Mapped[int | None] = mapped_column(Integer)
    height_cm: Mapped[float | None] = mapped_column(Numeric(5, 2))
    weight_kg: Mapped[float | None] = mapped_column(Numeric(5, 2))
    bmi: Mapped[float | None] = mapped_column(Numeric(4, 2))
    recorded_by_provider_id: Mapped[uuid.UUID | None] = mapped_column(
        UUID(as_uuid=True), ForeignKey("providers.id", ondelete="SET NULL")
    )


class ClinicalObservation(Base, TimestampMixin):
    __tablename__ = "clinical_observations"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    patient_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("patients.id", ondelete="CASCADE"), nullable=False, index=True
    )
    encounter_id: Mapped[uuid.UUID | None] = mapped_column(
        UUID(as_uuid=True), ForeignKey("encounters.id", ondelete="SET NULL")
    )
    loinc_code: Mapped[str] = mapped_column(
        String(16), ForeignKey("loinc_codes.code", ondelete="RESTRICT"), nullable=False
    )
    value_numeric: Mapped[float | None] = mapped_column(Numeric(12, 4))
    value_text: Mapped[str | None] = mapped_column(Text)
    units: Mapped[str | None] = mapped_column(String(32))
    interpretation: Mapped[str | None] = mapped_column(String(8))
    observed_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)
    recorded_by_provider_id: Mapped[uuid.UUID | None] = mapped_column(
        UUID(as_uuid=True), ForeignKey("providers.id", ondelete="SET NULL")
    )


class CarePlan(Base, TimestampMixin):
    __tablename__ = "care_plans"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    patient_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("patients.id", ondelete="CASCADE"), nullable=False
    )
    title: Mapped[str] = mapped_column(String(255), nullable=False)
    category: Mapped[str | None] = mapped_column(String(64))
    status: Mapped[str | None] = mapped_column(String(16))
    start_date: Mapped[date | None] = mapped_column(Date)
    end_date: Mapped[date | None] = mapped_column(Date)
    responsible_provider_id: Mapped[uuid.UUID | None] = mapped_column(
        UUID(as_uuid=True), ForeignKey("providers.id", ondelete="SET NULL")
    )
    description: Mapped[str | None] = mapped_column(Text)

    goals: Mapped[list["CarePlanGoal"]] = relationship(
        back_populates="care_plan", cascade="all, delete-orphan"
    )


class CarePlanGoal(Base, TimestampMixin):
    __tablename__ = "care_plan_goals"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    care_plan_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("care_plans.id", ondelete="CASCADE"), nullable=False
    )
    description: Mapped[str] = mapped_column(Text, nullable=False)
    target_date: Mapped[date | None] = mapped_column(Date)
    achievement_status: Mapped[str | None] = mapped_column(String(32))
    priority: Mapped[str | None] = mapped_column(String(16))

    care_plan: Mapped["CarePlan"] = relationship(back_populates="goals")
