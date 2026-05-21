# Hospital-env — Schema Specification

> **Authoritative reference** for the Hospital-env synthetic environment. Use this document when authoring the `.xlsx` seed workbook that `setup/ingest/xlsx_loader.py` will load into the database.

This schema is intentionally **complicated** — it mirrors what a real Hospital Information System (HIS) / Electronic Health Record (EHR) needs to track, modeled on conventions from **HL7 FHIR R4**, **OpenMRS**, **OpenEMR**, and the U.S. revenue-cycle stack (X12 837/835, CMS-1500, UB-04). Standard code systems referenced: **ICD-10-CM**, **ICD-10-PCS**, **CPT/HCPCS**, **LOINC**, **RxNorm**, **NDC**, **SNOMED CT**.

---

## Table of contents

- [Conventions](#conventions)
- [xlsx workbook format](#xlsx-workbook-format)
- [Domain 1 — Patients & Demographics](#domain-1--patients--demographics)
- [Domain 2 — Providers & Organization](#domain-2--providers--organization)
- [Domain 3 — Scheduling](#domain-3--scheduling)
- [Domain 4 — Encounters (ADT)](#domain-4--encounters-adt)
- [Domain 5 — Clinical](#domain-5--clinical)
- [Domain 6 — Medications](#domain-6--medications)
- [Domain 7 — Labs & Imaging](#domain-7--labs--imaging)
- [Domain 8 — Insurance & Billing](#domain-8--insurance--billing)
- [Domain 9 — Communications](#domain-9--communications)
- [Domain 10 — Operations & Audit](#domain-10--operations--audit)
- [MongoDB collections](#mongodb-collections)
- [Consolidated table list](#consolidated-table-list)
- [Load order for the xlsx](#load-order-for-the-xlsx)

---

## Conventions

| Convention | Rule |
| --- | --- |
| **Primary keys** | All tables use a `UUID` PK named `id` unless noted otherwise. Generate UUIDv4 in your xlsx. |
| **Timestamps** | `created_at`, `updated_at` are server-managed (`TIMESTAMPTZ`). Do **not** include them in xlsx sheets. |
| **Foreign keys** | Column named `<entity>_id` references `<entity_plural>.id` (e.g. `patient_id` → `patients.id`). |
| **Soft deletes** | Where present, `deleted_at` (`TIMESTAMPTZ`, nullable). |
| **Enums** | Stored as `TEXT` with a `CHECK` constraint; allowed values listed per column. |
| **Money** | `NUMERIC(12, 2)`, USD assumed. |
| **Code references** | `*_code` columns store the literal code string; `*_system` columns store the code system URI (`http://hl7.org/fhir/sid/icd-10-cm`, `http://loinc.org`, etc.). |
| **JSON** | `JSONB` for semi-structured payloads (FHIR extensions, audit before/after diffs). |
| **Case** | All identifiers `snake_case`, tables plural. |

---

## xlsx workbook format

The seed workbook is a **single `.xlsx` file** with **one sheet per table**. Rules:

1. **Sheet name** = exact table name (e.g. `patients`, `lab_orders`).
2. **Row 1** = header row of column names (exact, case-sensitive).
3. **Rows 2…N** = one row per record. Leave a cell blank for `NULL`.
4. **UUIDs** as strings (`a1b2c3d4-...`); the loader will cast.
5. **Dates** as ISO-8601 strings (`2025-03-14` for dates, `2025-03-14T09:30:00Z` for timestamps).
6. **Booleans** as `true` / `false` (case-insensitive).
7. **JSON columns** as a JSON-encoded string (`{"key": "value"}`); the loader parses.
8. **FK columns** must reference an `id` already present in the workbook — see the [load order](#load-order-for-the-xlsx).
9. **Lookup / catalog sheets** (e.g. `icd10_codes`, `cpt_codes`, `loinc_codes`, `rxnorm_concepts`, `payers`) should be filled first.
10. Sheets the loader doesn't know about are ignored.

---

## Domain 1 — Patients & Demographics

The patient is the central entity. Demographics here align with **FHIR `Patient`** plus **U.S. Core** extensions for race / ethnicity / preferred language.

### `patients`

The patient master record (one per human).

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `mrn` | TEXT UNIQUE NOT NULL | Medical Record Number, facility-issued |
| `first_name` | TEXT NOT NULL | |
| `middle_name` | TEXT | |
| `last_name` | TEXT NOT NULL | |
| `prefix` | TEXT | Mr, Mrs, Dr, … |
| `suffix` | TEXT | Jr, III, … |
| `date_of_birth` | DATE NOT NULL | |
| `sex_at_birth` | TEXT | `male`, `female`, `intersex`, `unknown` |
| `gender_identity` | TEXT | free-form (FHIR US-Core) |
| `pronouns` | TEXT | |
| `race` | TEXT | OMB categories |
| `ethnicity` | TEXT | `hispanic_or_latino`, `not_hispanic_or_latino`, `unknown` |
| `preferred_language` | TEXT | ISO 639-1 (`en`, `es`, …) |
| `marital_status` | TEXT | `single`, `married`, `divorced`, `widowed`, `separated`, `unknown` |
| `religion` | TEXT | |
| `ssn_last4` | TEXT | last 4 digits only |
| `is_deceased` | BOOLEAN DEFAULT false | |
| `deceased_at` | TIMESTAMPTZ | |
| `primary_provider_id` | UUID FK → `providers.id` | PCP |
| `vip_status` | TEXT | `none`, `vip`, `restricted` |
| `notes` | TEXT | |

### `patient_identifiers`

Multi-typed external IDs (driver's license, passport, MBI Medicare, etc.).

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `patient_id` | UUID FK → `patients.id` NOT NULL | |
| `identifier_type` | TEXT NOT NULL | `ssn`, `drivers_license`, `passport`, `medicare_mbi`, `medicaid_id`, `other` |
| `identifier_value` | TEXT NOT NULL | |
| `issuing_authority` | TEXT | e.g. `CA-DMV`, `CMS` |
| `valid_from` | DATE | |
| `valid_to` | DATE | |

### `patient_addresses`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `patient_id` | UUID FK → `patients.id` NOT NULL | |
| `address_use` | TEXT NOT NULL | `home`, `work`, `temp`, `billing`, `old` |
| `line1` | TEXT NOT NULL | |
| `line2` | TEXT | |
| `city` | TEXT NOT NULL | |
| `state` | TEXT NOT NULL | 2-letter |
| `postal_code` | TEXT NOT NULL | |
| `country` | TEXT DEFAULT 'US' | ISO 3166-1 alpha-2 |
| `is_primary` | BOOLEAN DEFAULT false | |
| `valid_from` | DATE | |
| `valid_to` | DATE | |

### `patient_contacts`

Phone / email / fax channels.

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `patient_id` | UUID FK → `patients.id` NOT NULL | |
| `contact_system` | TEXT NOT NULL | `phone`, `mobile`, `email`, `fax`, `pager`, `sms` |
| `contact_value` | TEXT NOT NULL | |
| `contact_use` | TEXT | `home`, `work`, `mobile`, `temp` |
| `is_primary` | BOOLEAN DEFAULT false | |
| `consent_to_contact` | BOOLEAN DEFAULT true | |

### `emergency_contacts`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `patient_id` | UUID FK → `patients.id` NOT NULL | |
| `name` | TEXT NOT NULL | |
| `relationship` | TEXT NOT NULL | `spouse`, `parent`, `child`, `sibling`, `friend`, `other` |
| `phone` | TEXT NOT NULL | |
| `email` | TEXT | |
| `address_line1` | TEXT | |
| `city` | TEXT | |
| `state` | TEXT | |
| `postal_code` | TEXT | |
| `priority_rank` | INT DEFAULT 1 | 1 = first to contact |
| `has_medical_poa` | BOOLEAN DEFAULT false | |

### `patient_consents`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `patient_id` | UUID FK → `patients.id` NOT NULL | |
| `consent_type` | TEXT NOT NULL | `hipaa_notice`, `treatment`, `release_of_info`, `research`, `telehealth`, `financial` |
| `granted` | BOOLEAN NOT NULL | |
| `granted_at` | TIMESTAMPTZ NOT NULL | |
| `expires_at` | TIMESTAMPTZ | |
| `witness_name` | TEXT | |
| `document_ref` | TEXT | path / URL to signed PDF |

---

## Domain 2 — Providers & Organization

### `providers`

Clinical and ancillary staff who can be assigned to encounters / orders.

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `npi` | TEXT UNIQUE | 10-digit National Provider Identifier (CMS) |
| `dea_number` | TEXT | for prescribers |
| `first_name` | TEXT NOT NULL | |
| `last_name` | TEXT NOT NULL | |
| `credentials` | TEXT | `MD`, `DO`, `RN`, `PA-C`, `NP`, `PharmD`, … |
| `provider_type` | TEXT NOT NULL | `physician`, `nurse`, `np`, `pa`, `pharmacist`, `tech`, `therapist`, `admin` |
| `email` | TEXT | |
| `phone` | TEXT | |
| `hire_date` | DATE | |
| `termination_date` | DATE | |
| `is_active` | BOOLEAN DEFAULT true | |
| `primary_department_id` | UUID FK → `departments.id` | |

### `specialties`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `code` | TEXT UNIQUE NOT NULL | NUCC taxonomy code |
| `name` | TEXT NOT NULL | `Internal Medicine`, `Cardiology`, … |
| `category` | TEXT | `physician`, `nursing`, `allied_health` |

### `provider_specialties`

Many-to-many.

| Column | Type | Notes |
| --- | --- | --- |
| `provider_id` | UUID FK → `providers.id` NOT NULL | composite PK |
| `specialty_id` | UUID FK → `specialties.id` NOT NULL | composite PK |
| `is_primary` | BOOLEAN DEFAULT false | |
| `board_certified` | BOOLEAN DEFAULT false | |
| `certification_date` | DATE | |

### `provider_licenses`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `provider_id` | UUID FK → `providers.id` NOT NULL | |
| `license_type` | TEXT NOT NULL | `state_medical`, `dea`, `controlled_substance`, `board_cert` |
| `license_number` | TEXT NOT NULL | |
| `issuing_state` | TEXT | |
| `issue_date` | DATE | |
| `expiration_date` | DATE | |
| `status` | TEXT | `active`, `expired`, `suspended`, `revoked` |

### `locations`

Physical facilities (hospital, clinic, lab draw site).

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `name` | TEXT NOT NULL | |
| `facility_type` | TEXT | `hospital`, `clinic`, `urgent_care`, `lab`, `imaging_center`, `pharmacy` |
| `address_line1` | TEXT | |
| `city` | TEXT | |
| `state` | TEXT | |
| `postal_code` | TEXT | |
| `phone` | TEXT | |
| `npi` | TEXT | facility NPI if applicable |

### `departments`

Organizational units inside a location (Cardiology, ED, ICU, Lab).

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `location_id` | UUID FK → `locations.id` NOT NULL | |
| `name` | TEXT NOT NULL | |
| `code` | TEXT | facility-local |
| `department_type` | TEXT | `inpatient`, `outpatient`, `ed`, `or`, `icu`, `lab`, `pharmacy`, `admin` |
| `phone` | TEXT | |

### `units`

Subdivision of a department (e.g. ICU → MICU/SICU/CCU).

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `department_id` | UUID FK → `departments.id` NOT NULL | |
| `name` | TEXT NOT NULL | |
| `unit_type` | TEXT | `med_surg`, `icu`, `step_down`, `telemetry`, `peds`, `nursery`, `obs` |
| `bed_capacity` | INT | |

### `rooms`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `unit_id` | UUID FK → `units.id` NOT NULL | |
| `room_number` | TEXT NOT NULL | |
| `room_type` | TEXT | `private`, `semi_private`, `isolation`, `procedure` |

### `beds`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `room_id` | UUID FK → `rooms.id` NOT NULL | |
| `bed_label` | TEXT NOT NULL | `A`, `B`, `1`, … |
| `status` | TEXT | `available`, `occupied`, `cleaning`, `out_of_service`, `reserved` |
| `is_monitored` | BOOLEAN DEFAULT false | telemetry-equipped |

---

## Domain 3 — Scheduling

### `appointment_types`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `code` | TEXT UNIQUE NOT NULL | `NP15`, `FU30`, `PHYS`, … |
| `name` | TEXT NOT NULL | |
| `default_duration_minutes` | INT NOT NULL | |
| `requires_referral` | BOOLEAN DEFAULT false | |

### `appointment_slots`

Provider availability blocks the scheduler draws from.

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `provider_id` | UUID FK → `providers.id` NOT NULL | |
| `location_id` | UUID FK → `locations.id` NOT NULL | |
| `slot_start` | TIMESTAMPTZ NOT NULL | |
| `slot_end` | TIMESTAMPTZ NOT NULL | |
| `is_available` | BOOLEAN DEFAULT true | |
| `appointment_type_id` | UUID FK → `appointment_types.id` | optional restriction |

### `appointments`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `patient_id` | UUID FK → `patients.id` NOT NULL | |
| `provider_id` | UUID FK → `providers.id` NOT NULL | |
| `appointment_type_id` | UUID FK → `appointment_types.id` NOT NULL | |
| `location_id` | UUID FK → `locations.id` NOT NULL | |
| `department_id` | UUID FK → `departments.id` | |
| `scheduled_start` | TIMESTAMPTZ NOT NULL | |
| `scheduled_end` | TIMESTAMPTZ NOT NULL | |
| `actual_start` | TIMESTAMPTZ | |
| `actual_end` | TIMESTAMPTZ | |
| `status` | TEXT NOT NULL | `scheduled`, `checked_in`, `in_progress`, `completed`, `cancelled`, `no_show`, `rescheduled` |
| `reason_for_visit` | TEXT | |
| `referral_id` | UUID FK → `referrals.id` | |
| `created_by_provider_id` | UUID FK → `providers.id` | |

### `appointment_status_history`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `appointment_id` | UUID FK → `appointments.id` NOT NULL | |
| `from_status` | TEXT | |
| `to_status` | TEXT NOT NULL | |
| `changed_at` | TIMESTAMPTZ NOT NULL | |
| `changed_by_provider_id` | UUID FK → `providers.id` | |
| `reason` | TEXT | |

### `referrals`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `patient_id` | UUID FK → `patients.id` NOT NULL | |
| `referring_provider_id` | UUID FK → `providers.id` NOT NULL | |
| `referred_to_provider_id` | UUID FK → `providers.id` | nullable, may be specialty only |
| `referred_to_specialty_id` | UUID FK → `specialties.id` | |
| `reason` | TEXT NOT NULL | |
| `priority` | TEXT | `routine`, `urgent`, `stat` |
| `status` | TEXT | `requested`, `scheduled`, `completed`, `cancelled`, `expired` |
| `created_at_date` | DATE NOT NULL | |
| `expires_on` | DATE | |
| `authorization_id` | UUID FK → `authorizations.id` | |

### `waitlist_entries`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `patient_id` | UUID FK → `patients.id` NOT NULL | |
| `provider_id` | UUID FK → `providers.id` | |
| `appointment_type_id` | UUID FK → `appointment_types.id` | |
| `requested_after` | DATE | |
| `requested_before` | DATE | |
| `priority` | TEXT | `routine`, `urgent` |
| `status` | TEXT | `waiting`, `offered`, `accepted`, `declined`, `expired` |

---

## Domain 4 — Encounters (ADT)

The **encounter** ties patient + provider + location + time. Admission/Discharge/Transfer (ADT) movements live alongside.

### `encounters`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `patient_id` | UUID FK → `patients.id` NOT NULL | |
| `encounter_class` | TEXT NOT NULL | `ambulatory`, `inpatient`, `emergency`, `observation`, `home_health`, `virtual`, `inpatient_rehab` |
| `status` | TEXT NOT NULL | `planned`, `arrived`, `in_progress`, `discharged`, `cancelled` |
| `admission_type` | TEXT | `elective`, `urgent`, `emergency`, `newborn`, `trauma` |
| `admission_source` | TEXT | `physician_referral`, `clinic_referral`, `transfer`, `ed`, `court_law` |
| `chief_complaint` | TEXT | |
| `attending_provider_id` | UUID FK → `providers.id` | |
| `admitting_provider_id` | UUID FK → `providers.id` | |
| `location_id` | UUID FK → `locations.id` NOT NULL | |
| `department_id` | UUID FK → `departments.id` | |
| `appointment_id` | UUID FK → `appointments.id` | nullable for walk-ins |
| `admitted_at` | TIMESTAMPTZ | |
| `discharged_at` | TIMESTAMPTZ | |
| `discharge_disposition` | TEXT | `home`, `home_health`, `snf`, `rehab`, `expired`, `ama`, `transfer_acute`, `hospice` |
| `triage_acuity` | INT | ESI 1-5 (1 = most acute) |

### `bed_assignments`

ADT movement history.

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `encounter_id` | UUID FK → `encounters.id` NOT NULL | |
| `bed_id` | UUID FK → `beds.id` NOT NULL | |
| `assigned_at` | TIMESTAMPTZ NOT NULL | |
| `released_at` | TIMESTAMPTZ | NULL = currently occupied |
| `assigned_by_provider_id` | UUID FK → `providers.id` | |
| `reason` | TEXT | `admit`, `transfer`, `discharge` |

### `encounter_diagnoses`

Links encounters to ICD-10 codes with ranking.

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `encounter_id` | UUID FK → `encounters.id` NOT NULL | |
| `icd10_code` | TEXT NOT NULL | FK → `icd10_codes.code` |
| `diagnosis_type` | TEXT NOT NULL | `admitting`, `principal`, `secondary`, `discharge`, `working` |
| `present_on_admission` | TEXT | `Y`, `N`, `U`, `W` |
| `rank` | INT | 1 = principal |
| `documented_by_provider_id` | UUID FK → `providers.id` | |
| `documented_at` | TIMESTAMPTZ | |

### `encounter_procedures`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `encounter_id` | UUID FK → `encounters.id` NOT NULL | |
| `cpt_code` | TEXT NOT NULL | FK → `cpt_codes.code` |
| `performed_at` | TIMESTAMPTZ NOT NULL | |
| `performing_provider_id` | UUID FK → `providers.id` | |
| `assistant_provider_id` | UUID FK → `providers.id` | |
| `location_id` | UUID FK → `locations.id` | |
| `modifier` | TEXT | CPT modifier(s), comma-separated |
| `notes` | TEXT | |

---

## Domain 5 — Clinical

### `icd10_codes` (catalog)

| Column | Type | Notes |
| --- | --- | --- |
| `code` | TEXT PK | e.g. `E11.9` |
| `description` | TEXT NOT NULL | |
| `chapter` | TEXT | |
| `is_billable` | BOOLEAN DEFAULT true | |

### `cpt_codes` (catalog)

| Column | Type | Notes |
| --- | --- | --- |
| `code` | TEXT PK | e.g. `99213` |
| `description` | TEXT NOT NULL | |
| `category` | TEXT | `E/M`, `surgery`, `radiology`, `pathology`, `medicine`, `hcpcs` |
| `default_charge` | NUMERIC(12,2) | |

### `loinc_codes` (catalog)

| Column | Type | Notes |
| --- | --- | --- |
| `code` | TEXT PK | e.g. `718-7` |
| `display` | TEXT NOT NULL | `Hemoglobin [Mass/volume] in Blood` |
| `class` | TEXT | `HEM/BC`, `CHEM`, … |
| `system` | TEXT | `Bld`, `Ser/Plas`, `Urine` |
| `units_default` | TEXT | `g/dL` |

### `snomed_codes` (catalog, optional)

| Column | Type | Notes |
| --- | --- | --- |
| `code` | TEXT PK | |
| `display` | TEXT NOT NULL | |
| `semantic_tag` | TEXT | `disorder`, `finding`, `procedure`, `body structure` |

### `problem_list_entries`

A patient's longitudinal problem list (distinct from encounter diagnoses).

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `patient_id` | UUID FK → `patients.id` NOT NULL | |
| `icd10_code` | TEXT FK → `icd10_codes.code` | |
| `snomed_code` | TEXT FK → `snomed_codes.code` | |
| `description` | TEXT NOT NULL | |
| `clinical_status` | TEXT NOT NULL | `active`, `recurrence`, `relapse`, `inactive`, `remission`, `resolved` |
| `verification_status` | TEXT | `unconfirmed`, `provisional`, `confirmed`, `refuted`, `entered_in_error` |
| `severity` | TEXT | `mild`, `moderate`, `severe` |
| `onset_date` | DATE | |
| `resolved_date` | DATE | |
| `recorded_by_provider_id` | UUID FK → `providers.id` | |

### `allergies`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `patient_id` | UUID FK → `patients.id` NOT NULL | |
| `allergen_type` | TEXT NOT NULL | `medication`, `food`, `environment`, `biologic` |
| `allergen_name` | TEXT NOT NULL | |
| `rxnorm_code` | TEXT FK → `rxnorm_concepts.rxcui` | nullable |
| `criticality` | TEXT | `low`, `high`, `unable_to_assess` |
| `clinical_status` | TEXT | `active`, `inactive`, `resolved` |
| `verification_status` | TEXT | `confirmed`, `unconfirmed`, `refuted` |
| `recorded_date` | DATE | |
| `last_occurrence_date` | DATE | |

### `allergy_reactions`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `allergy_id` | UUID FK → `allergies.id` NOT NULL | |
| `manifestation` | TEXT NOT NULL | `hives`, `anaphylaxis`, `rash`, `nausea`, `swelling`, … |
| `severity` | TEXT | `mild`, `moderate`, `severe` |
| `onset_minutes` | INT | minutes after exposure |

### `vital_signs`

Discrete observations (latest blood pressure, HR, temp, SpO₂, pain).

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `patient_id` | UUID FK → `patients.id` NOT NULL | |
| `encounter_id` | UUID FK → `encounters.id` | |
| `measured_at` | TIMESTAMPTZ NOT NULL | |
| `systolic_bp` | INT | mmHg |
| `diastolic_bp` | INT | mmHg |
| `heart_rate` | INT | bpm |
| `respiratory_rate` | INT | breaths/min |
| `temperature_c` | NUMERIC(4,2) | °C |
| `spo2` | INT | % |
| `pain_score` | INT | 0-10 |
| `height_cm` | NUMERIC(5,2) | |
| `weight_kg` | NUMERIC(5,2) | |
| `bmi` | NUMERIC(4,2) | |
| `recorded_by_provider_id` | UUID FK → `providers.id` | |

### `clinical_observations`

Generic LOINC-coded observations beyond vitals.

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `patient_id` | UUID FK → `patients.id` NOT NULL | |
| `encounter_id` | UUID FK → `encounters.id` | |
| `loinc_code` | TEXT FK → `loinc_codes.code` NOT NULL | |
| `value_numeric` | NUMERIC(12,4) | |
| `value_text` | TEXT | |
| `units` | TEXT | UCUM |
| `interpretation` | TEXT | `N`, `H`, `L`, `A`, `AA` |
| `observed_at` | TIMESTAMPTZ NOT NULL | |
| `recorded_by_provider_id` | UUID FK → `providers.id` | |

### `care_plans`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `patient_id` | UUID FK → `patients.id` NOT NULL | |
| `title` | TEXT NOT NULL | |
| `category` | TEXT | `assess_plan`, `careteam`, `chronic_disease_management`, `discharge` |
| `status` | TEXT | `draft`, `active`, `completed`, `on_hold`, `cancelled` |
| `start_date` | DATE | |
| `end_date` | DATE | |
| `responsible_provider_id` | UUID FK → `providers.id` | |
| `description` | TEXT | |

### `care_plan_goals`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `care_plan_id` | UUID FK → `care_plans.id` NOT NULL | |
| `description` | TEXT NOT NULL | |
| `target_date` | DATE | |
| `achievement_status` | TEXT | `in_progress`, `achieved`, `not_achieved`, `sustaining` |
| `priority` | TEXT | `low`, `medium`, `high` |

---

## Domain 6 — Medications

### `rxnorm_concepts` (catalog)

| Column | Type | Notes |
| --- | --- | --- |
| `rxcui` | TEXT PK | RxNorm CUI |
| `name` | TEXT NOT NULL | |
| `tty` | TEXT | term type — `SCD`, `SBD`, `IN`, `BN` |
| `route` | TEXT | `oral`, `iv`, `im`, `subq`, `topical` |
| `is_controlled` | BOOLEAN DEFAULT false | |
| `dea_schedule` | TEXT | `CII`-`CV` |

### `medications` (formulary)

A drug as the facility stocks / prescribes it.

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `rxcui` | TEXT FK → `rxnorm_concepts.rxcui` | |
| `ndc` | TEXT | National Drug Code, 10/11-digit |
| `name` | TEXT NOT NULL | |
| `strength` | TEXT | `500 mg`, `5 mg/mL` |
| `dosage_form` | TEXT | `tablet`, `capsule`, `injection`, `solution`, `suspension`, `inhaler` |
| `route` | TEXT | |
| `manufacturer` | TEXT | |
| `is_on_formulary` | BOOLEAN DEFAULT true | |

### `prescriptions`

(a.k.a. `medication_orders`).

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `patient_id` | UUID FK → `patients.id` NOT NULL | |
| `prescriber_provider_id` | UUID FK → `providers.id` NOT NULL | |
| `encounter_id` | UUID FK → `encounters.id` | |
| `medication_id` | UUID FK → `medications.id` NOT NULL | |
| `dose` | TEXT NOT NULL | `500 mg` |
| `route` | TEXT NOT NULL | |
| `frequency` | TEXT NOT NULL | `q8h`, `bid`, `prn`, `daily` |
| `duration_days` | INT | |
| `quantity` | NUMERIC(10,2) | |
| `refills` | INT DEFAULT 0 | |
| `start_date` | DATE NOT NULL | |
| `end_date` | DATE | |
| `status` | TEXT NOT NULL | `active`, `completed`, `cancelled`, `stopped`, `on_hold`, `entered_in_error` |
| `indication` | TEXT | reason for use |
| `prn_reason` | TEXT | |
| `pharmacy_id` | UUID FK → `pharmacies.id` | |
| `is_electronic` | BOOLEAN DEFAULT true | e-prescription |

### `medication_administrations`

The MAR — discrete dose events for inpatient/inhospital meds.

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `prescription_id` | UUID FK → `prescriptions.id` NOT NULL | |
| `patient_id` | UUID FK → `patients.id` NOT NULL | denorm for fast read |
| `encounter_id` | UUID FK → `encounters.id` | |
| `administered_at` | TIMESTAMPTZ NOT NULL | |
| `administered_by_provider_id` | UUID FK → `providers.id` NOT NULL | usually RN |
| `dose_given` | TEXT NOT NULL | |
| `route` | TEXT | |
| `status` | TEXT NOT NULL | `completed`, `not_done`, `held`, `refused`, `in_progress` |
| `not_done_reason` | TEXT | `patient_refused`, `npo`, `vomiting`, `unavailable`, `other` |
| `site` | TEXT | `left_deltoid`, `right_gluteus`, `iv_port` |

### `medication_reconciliations`

Periodic verification of home/inpatient meds at admit/discharge.

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `patient_id` | UUID FK → `patients.id` NOT NULL | |
| `encounter_id` | UUID FK → `encounters.id` NOT NULL | |
| `reconciliation_type` | TEXT | `admission`, `transfer`, `discharge` |
| `performed_by_provider_id` | UUID FK → `providers.id` | |
| `performed_at` | TIMESTAMPTZ NOT NULL | |
| `notes` | TEXT | |

### `pharmacies`

External fill pharmacies.

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `name` | TEXT NOT NULL | |
| `ncpdp_id` | TEXT | NCPDP pharmacy ID |
| `npi` | TEXT | |
| `phone` | TEXT | |
| `fax` | TEXT | |
| `address_line1` | TEXT | |
| `city` | TEXT | |
| `state` | TEXT | |
| `postal_code` | TEXT | |
| `is_mail_order` | BOOLEAN DEFAULT false | |

---

## Domain 7 — Labs & Imaging

### `lab_orders`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `patient_id` | UUID FK → `patients.id` NOT NULL | |
| `encounter_id` | UUID FK → `encounters.id` | |
| `ordering_provider_id` | UUID FK → `providers.id` NOT NULL | |
| `ordered_at` | TIMESTAMPTZ NOT NULL | |
| `priority` | TEXT | `routine`, `urgent`, `stat`, `asap` |
| `status` | TEXT NOT NULL | `ordered`, `collected`, `in_lab`, `resulted`, `cancelled` |
| `panel_loinc_code` | TEXT FK → `loinc_codes.code` | for whole panels (CBC, BMP) |
| `clinical_question` | TEXT | |
| `fasting_required` | BOOLEAN DEFAULT false | |

### `lab_specimens`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `lab_order_id` | UUID FK → `lab_orders.id` NOT NULL | |
| `specimen_type` | TEXT NOT NULL | `blood_serum`, `blood_plasma`, `whole_blood`, `urine`, `csf`, `swab`, `tissue` |
| `container` | TEXT | `red_top`, `lavender`, `green`, … |
| `collected_at` | TIMESTAMPTZ | |
| `collected_by_provider_id` | UUID FK → `providers.id` | |
| `received_in_lab_at` | TIMESTAMPTZ | |
| `volume_ml` | NUMERIC(6,2) | |
| `is_rejected` | BOOLEAN DEFAULT false | |
| `rejection_reason` | TEXT | |

### `lab_results`

One row per analyte.

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `lab_order_id` | UUID FK → `lab_orders.id` NOT NULL | |
| `lab_specimen_id` | UUID FK → `lab_specimens.id` | |
| `loinc_code` | TEXT FK → `loinc_codes.code` NOT NULL | |
| `value_numeric` | NUMERIC(14,4) | |
| `value_text` | TEXT | |
| `units` | TEXT | UCUM |
| `reference_range_low` | NUMERIC(14,4) | |
| `reference_range_high` | NUMERIC(14,4) | |
| `interpretation` | TEXT | `N`, `H`, `L`, `HH`, `LL`, `A`, `AA` |
| `is_critical` | BOOLEAN DEFAULT false | |
| `resulted_at` | TIMESTAMPTZ NOT NULL | |
| `verified_by_provider_id` | UUID FK → `providers.id` | |
| `status` | TEXT | `preliminary`, `final`, `corrected`, `cancelled` |

### `imaging_orders`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `patient_id` | UUID FK → `patients.id` NOT NULL | |
| `encounter_id` | UUID FK → `encounters.id` | |
| `ordering_provider_id` | UUID FK → `providers.id` NOT NULL | |
| `modality` | TEXT NOT NULL | `XR`, `CT`, `MR`, `US`, `NM`, `PT`, `MG`, `DX` |
| `body_part` | TEXT NOT NULL | |
| `cpt_code` | TEXT FK → `cpt_codes.code` | |
| `clinical_indication` | TEXT NOT NULL | |
| `priority` | TEXT | `routine`, `urgent`, `stat` |
| `status` | TEXT NOT NULL | `ordered`, `scheduled`, `performed`, `cancelled`, `resulted` |
| `requires_contrast` | BOOLEAN DEFAULT false | |
| `ordered_at` | TIMESTAMPTZ NOT NULL | |

### `imaging_studies`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `imaging_order_id` | UUID FK → `imaging_orders.id` NOT NULL | |
| `study_uid` | TEXT UNIQUE NOT NULL | DICOM Study Instance UID |
| `accession_number` | TEXT UNIQUE | |
| `performed_at` | TIMESTAMPTZ | |
| `performed_by_provider_id` | UUID FK → `providers.id` | technologist |
| `location_id` | UUID FK → `locations.id` | |
| `series_count` | INT | |
| `image_count` | INT | |
| `dicom_metadata_doc_id` | TEXT | reference to Mongo `imaging_metadata._id` |

### `imaging_reports`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `imaging_study_id` | UUID FK → `imaging_studies.id` NOT NULL | |
| `reading_radiologist_id` | UUID FK → `providers.id` | |
| `dictated_at` | TIMESTAMPTZ | |
| `signed_at` | TIMESTAMPTZ | |
| `findings` | TEXT | |
| `impression` | TEXT | |
| `recommendation` | TEXT | |
| `status` | TEXT | `preliminary`, `final`, `amended`, `addendum` |

---

## Domain 8 — Insurance & Billing

### `payers`

Insurance companies / programs.

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `name` | TEXT NOT NULL | |
| `payer_type` | TEXT NOT NULL | `commercial`, `medicare`, `medicaid`, `tricare`, `va`, `workers_comp`, `self_pay`, `auto` |
| `payer_id_external` | TEXT | clearinghouse payer ID |
| `address_line1` | TEXT | |
| `city` | TEXT | |
| `state` | TEXT | |
| `postal_code` | TEXT | |
| `phone` | TEXT | |
| `claims_phone` | TEXT | |
| `claims_fax` | TEXT | |
| `electronic_claims_supported` | BOOLEAN DEFAULT true | |

### `insurance_plans`

A specific product offered by a payer.

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `payer_id` | UUID FK → `payers.id` NOT NULL | |
| `plan_name` | TEXT NOT NULL | |
| `plan_type` | TEXT | `HMO`, `PPO`, `EPO`, `POS`, `HDHP`, `Medicare_Advantage`, `Medigap` |
| `metal_tier` | TEXT | `bronze`, `silver`, `gold`, `platinum`, `catastrophic` (ACA) |
| `group_number` | TEXT | |
| `effective_date` | DATE | |
| `termination_date` | DATE | |
| `requires_referrals` | BOOLEAN DEFAULT false | |
| `requires_prior_auth_for_imaging` | BOOLEAN DEFAULT false | |

### `patient_coverages`

A patient's enrollment in a plan.

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `patient_id` | UUID FK → `patients.id` NOT NULL | |
| `insurance_plan_id` | UUID FK → `insurance_plans.id` NOT NULL | |
| `subscriber_relationship` | TEXT NOT NULL | `self`, `spouse`, `child`, `other` |
| `subscriber_name` | TEXT | |
| `subscriber_dob` | DATE | |
| `member_id` | TEXT NOT NULL | as printed on card |
| `group_number` | TEXT | |
| `coverage_rank` | INT NOT NULL | 1 = primary, 2 = secondary, 3 = tertiary |
| `effective_date` | DATE NOT NULL | |
| `termination_date` | DATE | |
| `copay_pcp` | NUMERIC(8,2) | |
| `copay_specialist` | NUMERIC(8,2) | |
| `copay_er` | NUMERIC(8,2) | |
| `deductible_individual` | NUMERIC(10,2) | |
| `deductible_family` | NUMERIC(10,2) | |
| `oop_max_individual` | NUMERIC(10,2) | |
| `verified_at` | TIMESTAMPTZ | last eligibility check |

### `authorizations`

Pre-authorizations / referrals required by the payer.

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `patient_id` | UUID FK → `patients.id` NOT NULL | |
| `coverage_id` | UUID FK → `patient_coverages.id` NOT NULL | |
| `auth_number` | TEXT UNIQUE | issued by payer |
| `cpt_code` | TEXT FK → `cpt_codes.code` | |
| `requested_units` | INT | visits / procedures authorized |
| `approved_units` | INT | |
| `status` | TEXT NOT NULL | `requested`, `pending`, `approved`, `denied`, `expired`, `cancelled` |
| `effective_date` | DATE | |
| `expiration_date` | DATE | |
| `requested_at` | TIMESTAMPTZ | |
| `decided_at` | TIMESTAMPTZ | |
| `denial_reason` | TEXT | |

### `claims`

The 837-style claim submitted to a payer.

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `patient_id` | UUID FK → `patients.id` NOT NULL | |
| `encounter_id` | UUID FK → `encounters.id` | |
| `coverage_id` | UUID FK → `patient_coverages.id` NOT NULL | |
| `claim_number` | TEXT UNIQUE NOT NULL | internal |
| `payer_claim_id` | TEXT | payer's ICN |
| `claim_type` | TEXT NOT NULL | `professional_837p`, `institutional_837i`, `dental`, `pharmacy` |
| `status` | TEXT NOT NULL | `draft`, `submitted`, `accepted`, `rejected`, `paid`, `partial_paid`, `denied`, `appealed` |
| `service_start_date` | DATE NOT NULL | |
| `service_end_date` | DATE | |
| `total_charge` | NUMERIC(12,2) NOT NULL | |
| `total_allowed` | NUMERIC(12,2) | |
| `total_paid` | NUMERIC(12,2) | |
| `patient_responsibility` | NUMERIC(12,2) | |
| `submitted_at` | TIMESTAMPTZ | |
| `paid_at` | TIMESTAMPTZ | |
| `billing_provider_id` | UUID FK → `providers.id` | |

### `claim_lines`

Service-line detail (one per CPT performed).

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `claim_id` | UUID FK → `claims.id` NOT NULL | |
| `line_number` | INT NOT NULL | |
| `cpt_code` | TEXT FK → `cpt_codes.code` NOT NULL | |
| `modifier` | TEXT | |
| `icd10_pointer` | TEXT | comma-separated diagnosis pointers `1,2` |
| `service_date` | DATE NOT NULL | |
| `units` | NUMERIC(8,2) NOT NULL DEFAULT 1 | |
| `charge_amount` | NUMERIC(12,2) NOT NULL | |
| `allowed_amount` | NUMERIC(12,2) | |
| `paid_amount` | NUMERIC(12,2) | |
| `adjustment_amount` | NUMERIC(12,2) | |
| `denial_code` | TEXT | CARC / RARC |
| `place_of_service` | TEXT | CMS POS code (`11` office, `21` inpatient, `23` ED) |
| `rendering_provider_id` | UUID FK → `providers.id` | |

### `charges`

Pre-claim charge capture (driven by encounter activity).

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `encounter_id` | UUID FK → `encounters.id` NOT NULL | |
| `patient_id` | UUID FK → `patients.id` NOT NULL | |
| `cpt_code` | TEXT FK → `cpt_codes.code` NOT NULL | |
| `quantity` | NUMERIC(8,2) DEFAULT 1 | |
| `charge_amount` | NUMERIC(12,2) NOT NULL | |
| `posted_at` | TIMESTAMPTZ NOT NULL | |
| `posted_by_provider_id` | UUID FK → `providers.id` | |
| `claim_id` | UUID FK → `claims.id` | nullable until bundled |

### `payments`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `claim_id` | UUID FK → `claims.id` | nullable for patient payments not tied to a claim |
| `patient_id` | UUID FK → `patients.id` NOT NULL | |
| `payer_id` | UUID FK → `payers.id` | nullable for patient payments |
| `payment_type` | TEXT NOT NULL | `insurance`, `patient`, `adjustment`, `refund`, `writeoff` |
| `payment_method` | TEXT | `eft`, `check`, `cash`, `credit_card`, `ach` |
| `amount` | NUMERIC(12,2) NOT NULL | |
| `received_at` | TIMESTAMPTZ NOT NULL | |
| `reference_number` | TEXT | check #, EFT trace, transaction ID |
| `era_835_id` | TEXT | for payer remittances |

### `adjustments`

CARC-coded adjustments (write-offs, contractual).

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `claim_line_id` | UUID FK → `claim_lines.id` NOT NULL | |
| `adjustment_group` | TEXT | `CO`, `PR`, `OA`, `PI` |
| `reason_code` | TEXT NOT NULL | CARC (e.g. `45`) |
| `amount` | NUMERIC(12,2) NOT NULL | |
| `note` | TEXT | |

### `patient_statements`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `patient_id` | UUID FK → `patients.id` NOT NULL | |
| `statement_date` | DATE NOT NULL | |
| `period_start` | DATE | |
| `period_end` | DATE | |
| `previous_balance` | NUMERIC(12,2) | |
| `charges_total` | NUMERIC(12,2) | |
| `payments_total` | NUMERIC(12,2) | |
| `adjustments_total` | NUMERIC(12,2) | |
| `current_balance` | NUMERIC(12,2) NOT NULL | |
| `due_date` | DATE | |
| `status` | TEXT | `draft`, `sent`, `paid`, `in_collections` |
| `delivery_method` | TEXT | `paper`, `email`, `portal` |

---

## Domain 9 — Communications

### `patient_message_threads`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `patient_id` | UUID FK → `patients.id` NOT NULL | |
| `subject` | TEXT | |
| `category` | TEXT | `medical_question`, `rx_refill`, `appointment`, `billing`, `lab_result`, `referral`, `other` |
| `status` | TEXT | `open`, `awaiting_provider`, `awaiting_patient`, `closed` |
| `priority` | TEXT | `routine`, `urgent` |
| `assigned_provider_id` | UUID FK → `providers.id` | |
| `last_message_at` | TIMESTAMPTZ | |

### `patient_messages`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `thread_id` | UUID FK → `patient_message_threads.id` NOT NULL | |
| `channel` | TEXT NOT NULL | `portal`, `sms`, `email`, `app_push` |
| `direction` | TEXT NOT NULL | `inbound`, `outbound` |
| `sender_provider_id` | UUID FK → `providers.id` | NULL for patient-sent |
| `sender_patient_id` | UUID FK → `patients.id` | NULL for provider-sent |
| `body` | TEXT NOT NULL | |
| `sent_at` | TIMESTAMPTZ NOT NULL | |
| `read_at` | TIMESTAMPTZ | |
| `attachment_uri` | TEXT | |

### `call_logs`

Phone calls (inbound & outbound) with patients.

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `patient_id` | UUID FK → `patients.id` NOT NULL | |
| `provider_id` | UUID FK → `providers.id` | |
| `direction` | TEXT NOT NULL | `inbound`, `outbound` |
| `phone_number` | TEXT NOT NULL | |
| `started_at` | TIMESTAMPTZ NOT NULL | |
| `duration_seconds` | INT | |
| `outcome` | TEXT | `connected`, `voicemail`, `no_answer`, `wrong_number` |
| `reason` | TEXT | `appointment_reminder`, `result_callback`, `triage`, `rx_question`, `billing` |
| `summary` | TEXT | |
| `follow_up_required` | BOOLEAN DEFAULT false | |

### `appointment_reminders`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `appointment_id` | UUID FK → `appointments.id` NOT NULL | |
| `patient_id` | UUID FK → `patients.id` NOT NULL | |
| `channel` | TEXT NOT NULL | `sms`, `email`, `voice`, `portal` |
| `scheduled_send_at` | TIMESTAMPTZ NOT NULL | |
| `sent_at` | TIMESTAMPTZ | |
| `delivery_status` | TEXT | `queued`, `sent`, `delivered`, `failed`, `bounced` |
| `response` | TEXT | `confirmed`, `cancel_requested`, `reschedule_requested`, `none` |

### `insurance_correspondences`

All payer ↔ provider correspondence (EOBs, denials, authorizations, appeals).

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `payer_id` | UUID FK → `payers.id` NOT NULL | |
| `patient_id` | UUID FK → `patients.id` | |
| `claim_id` | UUID FK → `claims.id` | |
| `authorization_id` | UUID FK → `authorizations.id` | |
| `direction` | TEXT NOT NULL | `inbound`, `outbound` |
| `correspondence_type` | TEXT NOT NULL | `eob`, `denial`, `auth_request`, `auth_response`, `appeal`, `request_for_info`, `eligibility_verification` |
| `channel` | TEXT | `fax`, `mail`, `email`, `portal`, `phone`, `edi_277`, `edi_835` |
| `received_at` | TIMESTAMPTZ | |
| `sent_at` | TIMESTAMPTZ | |
| `subject` | TEXT | |
| `body` | TEXT | |
| `document_uri` | TEXT | |
| `requires_follow_up` | BOOLEAN DEFAULT false | |
| `follow_up_by` | DATE | |

### `claim_denials`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `claim_id` | UUID FK → `claims.id` NOT NULL | |
| `denial_date` | DATE NOT NULL | |
| `carc_code` | TEXT NOT NULL | |
| `rarc_code` | TEXT | |
| `denial_reason` | TEXT | |
| `is_appealable` | BOOLEAN DEFAULT true | |
| `appeal_deadline` | DATE | |
| `worked_by_provider_id` | UUID FK → `providers.id` | denials coordinator |
| `status` | TEXT | `new`, `working`, `appealed`, `resolved`, `written_off` |

### `claim_appeals`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `claim_denial_id` | UUID FK → `claim_denials.id` NOT NULL | |
| `appeal_level` | INT NOT NULL DEFAULT 1 | 1 / 2 / 3 |
| `submitted_at` | TIMESTAMPTZ NOT NULL | |
| `submitted_by_provider_id` | UUID FK → `providers.id` | |
| `narrative` | TEXT | |
| `outcome` | TEXT | `pending`, `overturned`, `upheld`, `partial` |
| `decided_at` | TIMESTAMPTZ | |
| `recovered_amount` | NUMERIC(12,2) | |

### `inter_provider_messages`

Internal consults / handoffs.

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `from_provider_id` | UUID FK → `providers.id` NOT NULL | |
| `to_provider_id` | UUID FK → `providers.id` NOT NULL | |
| `patient_id` | UUID FK → `patients.id` | |
| `encounter_id` | UUID FK → `encounters.id` | |
| `message_type` | TEXT | `consult_request`, `consult_reply`, `handoff`, `curbside`, `fyi` |
| `subject` | TEXT | |
| `body` | TEXT NOT NULL | |
| `sent_at` | TIMESTAMPTZ NOT NULL | |
| `read_at` | TIMESTAMPTZ | |
| `acknowledged_at` | TIMESTAMPTZ | |

---

## Domain 10 — Operations & Audit

### `shifts`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `name` | TEXT NOT NULL | `Day`, `Evening`, `Night`, `Weekend Days` |
| `start_time` | TIME NOT NULL | |
| `end_time` | TIME NOT NULL | |
| `crosses_midnight` | BOOLEAN DEFAULT false | |

### `staff_schedules`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `provider_id` | UUID FK → `providers.id` NOT NULL | |
| `department_id` | UUID FK → `departments.id` | |
| `unit_id` | UUID FK → `units.id` | |
| `shift_id` | UUID FK → `shifts.id` | |
| `work_date` | DATE NOT NULL | |
| `scheduled_start` | TIMESTAMPTZ NOT NULL | |
| `scheduled_end` | TIMESTAMPTZ NOT NULL | |
| `role` | TEXT | `attending`, `resident`, `charge_nurse`, `bedside_rn`, `tech` |
| `status` | TEXT | `scheduled`, `worked`, `pto`, `sick`, `swap`, `cancelled` |

### `on_call_assignments`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `provider_id` | UUID FK → `providers.id` NOT NULL | |
| `department_id` | UUID FK → `departments.id` | |
| `specialty_id` | UUID FK → `specialties.id` | |
| `on_call_start` | TIMESTAMPTZ NOT NULL | |
| `on_call_end` | TIMESTAMPTZ NOT NULL | |
| `pager_number` | TEXT | |

### `pharmacy_inventory`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `medication_id` | UUID FK → `medications.id` NOT NULL | |
| `location_id` | UUID FK → `locations.id` NOT NULL | |
| `lot_number` | TEXT | |
| `expiration_date` | DATE | |
| `quantity_on_hand` | NUMERIC(12,2) NOT NULL | |
| `reorder_level` | NUMERIC(12,2) | |
| `unit_cost` | NUMERIC(10,4) | |
| `last_restocked_at` | TIMESTAMPTZ | |

### `equipment`

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `name` | TEXT NOT NULL | |
| `equipment_type` | TEXT | `monitor`, `ventilator`, `infusion_pump`, `defibrillator`, `imaging`, `wheelchair`, `bed` |
| `manufacturer` | TEXT | |
| `model` | TEXT | |
| `serial_number` | TEXT UNIQUE | |
| `location_id` | UUID FK → `locations.id` | |
| `unit_id` | UUID FK → `units.id` | |
| `status` | TEXT | `in_use`, `available`, `maintenance`, `retired` |
| `last_maintenance_at` | TIMESTAMPTZ | |
| `next_maintenance_due` | DATE | |

### `tasks`

Workflow inbox items (RN tasks, billing follow-ups, denials, etc.).

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `assigned_provider_id` | UUID FK → `providers.id` | |
| `created_by_provider_id` | UUID FK → `providers.id` | |
| `patient_id` | UUID FK → `patients.id` | |
| `encounter_id` | UUID FK → `encounters.id` | |
| `task_type` | TEXT NOT NULL | `lab_review`, `result_callback`, `rx_refill`, `prior_auth`, `denial_followup`, `chart_review`, `signature`, `general` |
| `title` | TEXT NOT NULL | |
| `description` | TEXT | |
| `priority` | TEXT | `low`, `normal`, `high`, `urgent` |
| `status` | TEXT | `open`, `in_progress`, `completed`, `cancelled` |
| `due_at` | TIMESTAMPTZ | |
| `completed_at` | TIMESTAMPTZ | |

### `audit_logs_summary`

Lightweight summary of HIPAA access events; the heavy raw log goes to Mongo `audit_logs`.

| Column | Type | Notes |
| --- | --- | --- |
| `id` | UUID PK | |
| `actor_provider_id` | UUID FK → `providers.id` | |
| `patient_id` | UUID FK → `patients.id` | |
| `action` | TEXT NOT NULL | `view`, `create`, `update`, `delete`, `print`, `export` |
| `resource_type` | TEXT NOT NULL | `patient`, `encounter`, `lab_result`, `note`, … |
| `resource_id` | UUID | |
| `occurred_at` | TIMESTAMPTZ NOT NULL | |
| `success` | BOOLEAN DEFAULT true | |
| `mongo_log_id` | TEXT | reference to `audit_logs._id` |

---

## MongoDB collections

Stored in the database named by `MONGO_DB` (default `hospital_docs`). See `setup/nosql/collections.py` for the index definitions.

### `clinical_notes`

Free-form / SOAP / discharge summaries. Document shape:

```json
{
  "_id": "ObjectId",
  "patient_id": "uuid",
  "encounter_id": "uuid",
  "author_provider_id": "uuid",
  "note_type": "progress | soap | discharge_summary | history_and_physical | consult | nursing | procedure",
  "specialty": "cardiology",
  "written_at": "ISODate",
  "signed_at": "ISODate",
  "status": "draft | signed | amended | addendum",
  "sections": {
    "subjective": "...",
    "objective": "...",
    "assessment": "...",
    "plan": "..."
  },
  "free_text": "...",
  "tags": ["follow_up", "high_risk"],
  "addenda": [
    { "author_provider_id": "uuid", "added_at": "ISODate", "text": "..." }
  ]
}
```

Indexes: `{patient_id: 1, written_at: -1}`, `{encounter_id: 1}`, `{author_provider_id: 1}`, full-text on `free_text`.

### `audit_logs`

HIPAA access logging — verbose, append-only.

```json
{
  "_id": "ObjectId",
  "occurred_at": "ISODate",
  "actor_provider_id": "uuid",
  "actor_ip": "10.20.0.5",
  "actor_session_id": "...",
  "action": "view",
  "resource_type": "lab_result",
  "resource_id": "uuid",
  "patient_id": "uuid",
  "success": true,
  "reason_for_access": "treatment | payment | operations | research",
  "user_agent": "...",
  "before": { "...": "..." },
  "after": { "...": "..." }
}
```

Indexes: `{patient_id: 1, occurred_at: -1}`, `{actor_provider_id: 1, occurred_at: -1}`, TTL optional.

### `imaging_metadata`

DICOM-derived study/series/instance metadata. Document per study:

```json
{
  "_id": "ObjectId",
  "study_uid": "1.2.840.113619.2.55....",
  "patient_id": "uuid",
  "imaging_study_id": "uuid",
  "modality": "CT",
  "study_date": "ISODate",
  "study_description": "CT ABDOMEN W CONTRAST",
  "referring_physician": "Smith, J.",
  "series": [
    {
      "series_uid": "...",
      "series_number": 1,
      "modality": "CT",
      "body_part": "ABDOMEN",
      "instance_count": 200,
      "kvp": 120,
      "slice_thickness": 5.0
    }
  ],
  "dicom_tags": { "(0010,0010)": "DOE^JOHN", "...": "..." }
}
```

Indexes: `{patient_id: 1, study_date: -1}`, `{study_uid: 1}` unique.

### `vitals_streams`

High-frequency telemetry from bedside monitors / wearables.

```json
{
  "_id": "ObjectId",
  "patient_id": "uuid",
  "encounter_id": "uuid",
  "device_id": "MON-ICU-04",
  "metric": "heart_rate | spo2 | etco2 | art_systolic | art_diastolic | resp_rate | temp",
  "unit": "bpm",
  "started_at": "ISODate",
  "ended_at": "ISODate",
  "sample_rate_hz": 1,
  "samples": [
    { "t": "ISODate", "v": 82 },
    { "t": "ISODate", "v": 84 }
  ]
}
```

Indexes: `{patient_id: 1, metric: 1, started_at: -1}`, `{encounter_id: 1}`.

---

## Consolidated table list

**Postgres tables (73):**

```
# Patients & Demographics
patients
patient_identifiers
patient_addresses
patient_contacts
emergency_contacts
patient_consents

# Providers & Organization
providers
specialties
provider_specialties
provider_licenses
locations
departments
units
rooms
beds

# Scheduling
appointment_types
appointment_slots
appointments
appointment_status_history
referrals
waitlist_entries

# Encounters
encounters
bed_assignments
encounter_diagnoses
encounter_procedures

# Clinical (incl. code catalogs)
icd10_codes
cpt_codes
loinc_codes
snomed_codes
problem_list_entries
allergies
allergy_reactions
vital_signs
clinical_observations
care_plans
care_plan_goals

# Medications
rxnorm_concepts
medications
prescriptions
medication_administrations
medication_reconciliations
pharmacies

# Labs & Imaging
lab_orders
lab_specimens
lab_results
imaging_orders
imaging_studies
imaging_reports

# Insurance & Billing
payers
insurance_plans
patient_coverages
authorizations
claims
claim_lines
charges
payments
adjustments
patient_statements

# Communications
patient_message_threads
patient_messages
call_logs
appointment_reminders
insurance_correspondences
claim_denials
claim_appeals
inter_provider_messages

# Operations & Audit
shifts
staff_schedules
on_call_assignments
pharmacy_inventory
equipment
tasks
audit_logs_summary
```

**MongoDB collections (4):**

```
clinical_notes
audit_logs
imaging_metadata
vitals_streams
```

---

## Load order for the xlsx

The xlsx loader respects FK dependencies by loading sheets in this order. Make sure your workbook has all referenced rows in earlier sheets before later ones.

1. **Code catalogs first** (no FKs out):
   `icd10_codes`, `cpt_codes`, `loinc_codes`, `snomed_codes`, `rxnorm_concepts`
2. **Independent organizational** entities:
   `locations`, `specialties`, `shifts`, `appointment_types`, `payers`, `pharmacies`
3. **Org hierarchy**: `departments` → `units` → `rooms` → `beds`
4. **People**: `providers` → `provider_specialties`, `provider_licenses`
5. **Patients core**: `patients` → `patient_identifiers`, `patient_addresses`, `patient_contacts`, `emergency_contacts`, `patient_consents`
6. **Plans & coverage**: `insurance_plans` → `patient_coverages` → `authorizations`
7. **Scheduling**: `appointment_slots`, `referrals`, `appointments` → `appointment_status_history`, `waitlist_entries`, `appointment_reminders`
8. **Encounters**: `encounters` → `bed_assignments`, `encounter_diagnoses`, `encounter_procedures`
9. **Clinical**: `problem_list_entries`, `allergies` → `allergy_reactions`, `vital_signs`, `clinical_observations`, `care_plans` → `care_plan_goals`
10. **Meds**: `medications`, `prescriptions` → `medication_administrations`, `medication_reconciliations`
11. **Labs / Imaging**: `lab_orders` → `lab_specimens` → `lab_results`; `imaging_orders` → `imaging_studies` → `imaging_reports`
12. **Billing**: `claims` → `claim_lines`, `charges`, `payments`, `adjustments`, `claim_denials` → `claim_appeals`, `patient_statements`
13. **Communications**: `patient_message_threads` → `patient_messages`, `call_logs`, `insurance_correspondences`, `inter_provider_messages`
14. **Operations**: `staff_schedules`, `on_call_assignments`, `pharmacy_inventory`, `equipment`, `tasks`, `audit_logs_summary`

If a sheet contains a circular dependency (`appointments.referral_id` ↔ `referrals` — both can predate each other), the loader does a **two-pass insert**: pass 1 inserts rows with `NULL` for cross-cycle FKs, pass 2 fills them.
