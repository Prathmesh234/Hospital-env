"""SQLAlchemy ORM models for the hospital schema.

Importing this package registers every model with `Base.metadata`, so
downstream code (`Base.metadata.create_all`, Alembic) sees the full schema.
"""

from setup.models.base import Base, TimestampMixin
from setup.models.billing import (
    Adjustment,
    Authorization,
    Charge,
    Claim,
    ClaimLine,
    InsurancePlan,
    Payer,
    Payment,
    PatientCoverage,
    PatientStatement,
)
from setup.models.clinical import (
    Allergy,
    AllergyReaction,
    CarePlan,
    CarePlanGoal,
    ClinicalObservation,
    CPTCode,
    ICD10Code,
    LOINCCode,
    ProblemListEntry,
    SNOMEDCode,
    VitalSign,
)
from setup.models.communications import (
    AppointmentReminder,
    CallLog,
    ClaimAppeal,
    ClaimDenial,
    InsuranceCorrespondence,
    InterProviderMessage,
    PatientMessage,
    PatientMessageThread,
)
from setup.models.encounters import (
    BedAssignment,
    Encounter,
    EncounterDiagnosis,
    EncounterProcedure,
)
from setup.models.labs import (
    ImagingOrder,
    ImagingReport,
    ImagingStudy,
    LabOrder,
    LabResult,
    LabSpecimen,
)
from setup.models.medications import (
    Medication,
    MedicationAdministration,
    MedicationReconciliation,
    Pharmacy,
    Prescription,
    RxNormConcept,
)
from setup.models.operations import (
    AuditLogSummary,
    Equipment,
    OnCallAssignment,
    PharmacyInventory,
    Shift,
    StaffSchedule,
    Task,
)
from setup.models.patients import (
    EmergencyContact,
    Patient,
    PatientAddress,
    PatientConsent,
    PatientContact,
    PatientIdentifier,
)
from setup.models.providers import (
    Bed,
    Department,
    Location,
    Provider,
    ProviderLicense,
    ProviderSpecialty,
    Room,
    Specialty,
    Unit,
)
from setup.models.scheduling import (
    Appointment,
    AppointmentSlot,
    AppointmentStatusHistory,
    AppointmentType,
    Referral,
    WaitlistEntry,
)

__all__ = [
    "Base",
    "TimestampMixin",
    "Patient", "PatientIdentifier", "PatientAddress", "PatientContact",
    "EmergencyContact", "PatientConsent",
    "Provider", "Specialty", "ProviderSpecialty", "ProviderLicense",
    "Location", "Department", "Unit", "Room", "Bed",
    "AppointmentType", "AppointmentSlot", "Appointment",
    "AppointmentStatusHistory", "Referral", "WaitlistEntry",
    "Encounter", "BedAssignment", "EncounterDiagnosis", "EncounterProcedure",
    "ICD10Code", "CPTCode", "LOINCCode", "SNOMEDCode",
    "ProblemListEntry", "Allergy", "AllergyReaction",
    "VitalSign", "ClinicalObservation", "CarePlan", "CarePlanGoal",
    "RxNormConcept", "Medication", "Prescription",
    "MedicationAdministration", "MedicationReconciliation", "Pharmacy",
    "LabOrder", "LabSpecimen", "LabResult",
    "ImagingOrder", "ImagingStudy", "ImagingReport",
    "Payer", "InsurancePlan", "PatientCoverage", "Authorization",
    "Claim", "ClaimLine", "Charge", "Payment", "Adjustment", "PatientStatement",
    "PatientMessageThread", "PatientMessage", "CallLog", "AppointmentReminder",
    "InsuranceCorrespondence", "ClaimDenial", "ClaimAppeal", "InterProviderMessage",
    "Shift", "StaffSchedule", "OnCallAssignment", "PharmacyInventory",
    "Equipment", "Task", "AuditLogSummary",
]
