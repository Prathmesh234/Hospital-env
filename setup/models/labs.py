"""Labs (orders, specimens, results) + imaging (orders, studies, reports)."""

from __future__ import annotations

import uuid
from datetime import datetime

from sqlalchemy import Boolean, DateTime, ForeignKey, Integer, Numeric, String, Text
from sqlalchemy.dialects.postgresql import UUID
from sqlalchemy.orm import Mapped, mapped_column, relationship

from setup.models.base import Base, TimestampMixin


class LabOrder(Base, TimestampMixin):
    __tablename__ = "lab_orders"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    patient_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("patients.id", ondelete="CASCADE"), nullable=False, index=True
    )
    encounter_id: Mapped[uuid.UUID | None] = mapped_column(
        UUID(as_uuid=True), ForeignKey("encounters.id", ondelete="SET NULL")
    )
    ordering_provider_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("providers.id", ondelete="RESTRICT"), nullable=False
    )
    ordered_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)
    priority: Mapped[str | None] = mapped_column(String(16))
    status: Mapped[str] = mapped_column(String(16), nullable=False)
    panel_loinc_code: Mapped[str | None] = mapped_column(
        String(16), ForeignKey("loinc_codes.code", ondelete="SET NULL")
    )
    clinical_question: Mapped[str | None] = mapped_column(Text)
    fasting_required: Mapped[bool] = mapped_column(Boolean, default=False, nullable=False)

    specimens: Mapped[list["LabSpecimen"]] = relationship(
        back_populates="lab_order", cascade="all, delete-orphan"
    )
    results: Mapped[list["LabResult"]] = relationship(
        back_populates="lab_order", cascade="all, delete-orphan"
    )


class LabSpecimen(Base, TimestampMixin):
    __tablename__ = "lab_specimens"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    lab_order_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("lab_orders.id", ondelete="CASCADE"), nullable=False
    )
    specimen_type: Mapped[str] = mapped_column(String(32), nullable=False)
    container: Mapped[str | None] = mapped_column(String(32))
    collected_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))
    collected_by_provider_id: Mapped[uuid.UUID | None] = mapped_column(
        UUID(as_uuid=True), ForeignKey("providers.id", ondelete="SET NULL")
    )
    received_in_lab_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))
    volume_ml: Mapped[float | None] = mapped_column(Numeric(6, 2))
    is_rejected: Mapped[bool] = mapped_column(Boolean, default=False, nullable=False)
    rejection_reason: Mapped[str | None] = mapped_column(Text)

    lab_order: Mapped["LabOrder"] = relationship(back_populates="specimens")


class LabResult(Base, TimestampMixin):
    __tablename__ = "lab_results"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    lab_order_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("lab_orders.id", ondelete="CASCADE"), nullable=False
    )
    lab_specimen_id: Mapped[uuid.UUID | None] = mapped_column(
        UUID(as_uuid=True), ForeignKey("lab_specimens.id", ondelete="SET NULL")
    )
    loinc_code: Mapped[str] = mapped_column(
        String(16), ForeignKey("loinc_codes.code", ondelete="RESTRICT"), nullable=False
    )
    value_numeric: Mapped[float | None] = mapped_column(Numeric(14, 4))
    value_text: Mapped[str | None] = mapped_column(Text)
    units: Mapped[str | None] = mapped_column(String(32))
    reference_range_low: Mapped[float | None] = mapped_column(Numeric(14, 4))
    reference_range_high: Mapped[float | None] = mapped_column(Numeric(14, 4))
    interpretation: Mapped[str | None] = mapped_column(String(8))
    is_critical: Mapped[bool] = mapped_column(Boolean, default=False, nullable=False)
    resulted_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)
    verified_by_provider_id: Mapped[uuid.UUID | None] = mapped_column(
        UUID(as_uuid=True), ForeignKey("providers.id", ondelete="SET NULL")
    )
    status: Mapped[str | None] = mapped_column(String(16))

    lab_order: Mapped["LabOrder"] = relationship(back_populates="results")


class ImagingOrder(Base, TimestampMixin):
    __tablename__ = "imaging_orders"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    patient_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("patients.id", ondelete="CASCADE"), nullable=False, index=True
    )
    encounter_id: Mapped[uuid.UUID | None] = mapped_column(
        UUID(as_uuid=True), ForeignKey("encounters.id", ondelete="SET NULL")
    )
    ordering_provider_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("providers.id", ondelete="RESTRICT"), nullable=False
    )
    modality: Mapped[str] = mapped_column(String(8), nullable=False)
    body_part: Mapped[str] = mapped_column(String(64), nullable=False)
    cpt_code: Mapped[str | None] = mapped_column(
        String(16), ForeignKey("cpt_codes.code", ondelete="SET NULL")
    )
    clinical_indication: Mapped[str] = mapped_column(Text, nullable=False)
    priority: Mapped[str | None] = mapped_column(String(16))
    status: Mapped[str] = mapped_column(String(16), nullable=False)
    requires_contrast: Mapped[bool] = mapped_column(Boolean, default=False, nullable=False)
    ordered_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)

    studies: Mapped[list["ImagingStudy"]] = relationship(
        back_populates="imaging_order", cascade="all, delete-orphan"
    )


class ImagingStudy(Base, TimestampMixin):
    __tablename__ = "imaging_studies"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    imaging_order_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("imaging_orders.id", ondelete="CASCADE"), nullable=False
    )
    study_uid: Mapped[str] = mapped_column(String(128), unique=True, nullable=False)
    accession_number: Mapped[str | None] = mapped_column(String(64), unique=True)
    performed_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))
    performed_by_provider_id: Mapped[uuid.UUID | None] = mapped_column(
        UUID(as_uuid=True), ForeignKey("providers.id", ondelete="SET NULL")
    )
    location_id: Mapped[uuid.UUID | None] = mapped_column(
        UUID(as_uuid=True), ForeignKey("locations.id", ondelete="SET NULL")
    )
    series_count: Mapped[int | None] = mapped_column(Integer)
    image_count: Mapped[int | None] = mapped_column(Integer)
    dicom_metadata_doc_id: Mapped[str | None] = mapped_column(String(64))

    imaging_order: Mapped["ImagingOrder"] = relationship(back_populates="studies")
    reports: Mapped[list["ImagingReport"]] = relationship(
        back_populates="imaging_study", cascade="all, delete-orphan"
    )


class ImagingReport(Base, TimestampMixin):
    __tablename__ = "imaging_reports"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    imaging_study_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), ForeignKey("imaging_studies.id", ondelete="CASCADE"), nullable=False
    )
    reading_radiologist_id: Mapped[uuid.UUID | None] = mapped_column(
        UUID(as_uuid=True), ForeignKey("providers.id", ondelete="SET NULL")
    )
    dictated_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))
    signed_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))
    findings: Mapped[str | None] = mapped_column(Text)
    impression: Mapped[str | None] = mapped_column(Text)
    recommendation: Mapped[str | None] = mapped_column(Text)
    status: Mapped[str | None] = mapped_column(String(16))

    imaging_study: Mapped["ImagingStudy"] = relationship(back_populates="reports")
