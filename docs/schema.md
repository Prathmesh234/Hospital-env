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

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | mrn | first_name | middle_name | last_name | prefix | suffix | date_of_birth | sex_at_birth | gender_identity | pronouns | race | ethnicity | preferred_language | marital_status | religion | ssn_last4 | is_deceased | deceased_at | primary_provider_id | vip_status | notes |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | MRN00012345 | Jane | Marie | Doe | Ms |  | 1985-04-12 | female | female | she/her | White | not_hispanic_or_latino | en | single | None | 1234 | false |  | pr000001-0000-4000-8000-000000000001 | none | Stable, follow-up in 3 months. |

**Column reference:**

- `id` — UUID PK
- `mrn` — TEXT UNIQUE NOT NULL — Medical Record Number, facility-issued
- `first_name` — TEXT NOT NULL
- `middle_name` — TEXT
- `last_name` — TEXT NOT NULL
- `prefix` — TEXT — Mr, Mrs, Dr, …
- `suffix` — TEXT — Jr, III, …
- `date_of_birth` — DATE NOT NULL
- `sex_at_birth` — TEXT — `male`, `female`, `intersex`, `unknown`
- `gender_identity` — TEXT — free-form (FHIR US-Core)
- `pronouns` — TEXT
- `race` — TEXT — OMB categories
- `ethnicity` — TEXT — `hispanic_or_latino`, `not_hispanic_or_latino`, `unknown`
- `preferred_language` — TEXT — ISO 639-1 (`en`, `es`, …)
- `marital_status` — TEXT — `single`, `married`, `divorced`, `widowed`, `separated`, `unknown`
- `religion` — TEXT
- `ssn_last4` — TEXT — last 4 digits only
- `is_deceased` — BOOLEAN DEFAULT false
- `deceased_at` — TIMESTAMPTZ
- `primary_provider_id` — UUID FK → `providers.id` — PCP
- `vip_status` — TEXT — `none`, `vip`, `restricted`
- `notes` — TEXT

### `patient_identifiers`

Multi-typed external IDs (driver's license, passport, MBI Medicare, etc.).

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | patient_id | identifier_type | identifier_value | issuing_authority | valid_from | valid_to |
| --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | p0000001-0000-4000-8000-000000000001 | drivers_license | D1234567 | CA-DMV | 2024-01-01 | 2030-12-31 |

**Column reference:**

- `id` — UUID PK
- `patient_id` — UUID FK → `patients.id` NOT NULL
- `identifier_type` — TEXT NOT NULL — `ssn`, `drivers_license`, `passport`, `medicare_mbi`, `medicaid_id`, `other`
- `identifier_value` — TEXT NOT NULL
- `issuing_authority` — TEXT — e.g. `CA-DMV`, `CMS`
- `valid_from` — DATE
- `valid_to` — DATE

### `patient_addresses`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | patient_id | address_use | line1 | line2 | city | state | postal_code | country | is_primary | valid_from | valid_to |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | p0000001-0000-4000-8000-000000000001 | home | 123 Main St | Apt 4B | Springfield | CA | 94105 | US | true | 2024-01-01 | 2030-12-31 |

**Column reference:**

- `id` — UUID PK
- `patient_id` — UUID FK → `patients.id` NOT NULL
- `address_use` — TEXT NOT NULL — `home`, `work`, `temp`, `billing`, `old`
- `line1` — TEXT NOT NULL
- `line2` — TEXT
- `city` — TEXT NOT NULL
- `state` — TEXT NOT NULL — 2-letter
- `postal_code` — TEXT NOT NULL
- `country` — TEXT DEFAULT 'US' — ISO 3166-1 alpha-2
- `is_primary` — BOOLEAN DEFAULT false
- `valid_from` — DATE
- `valid_to` — DATE

### `patient_contacts`

Phone / email / fax channels.

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | patient_id | contact_system | contact_value | contact_use | is_primary | consent_to_contact |
| --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | p0000001-0000-4000-8000-000000000001 | mobile | 555-123-4567 | mobile | true | true |

**Column reference:**

- `id` — UUID PK
- `patient_id` — UUID FK → `patients.id` NOT NULL
- `contact_system` — TEXT NOT NULL — `phone`, `mobile`, `email`, `fax`, `pager`, `sms`
- `contact_value` — TEXT NOT NULL
- `contact_use` — TEXT — `home`, `work`, `mobile`, `temp`
- `is_primary` — BOOLEAN DEFAULT false
- `consent_to_contact` — BOOLEAN DEFAULT true

### `emergency_contacts`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | patient_id | name | relationship | phone | email | address_line1 | city | state | postal_code | priority_rank | has_medical_poa |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | p0000001-0000-4000-8000-000000000001 | Example Name | spouse | 555-123-4567 | jane.doe@example.com | 123 Main St | Springfield | CA | 94105 | 1 | false |

**Column reference:**

- `id` — UUID PK
- `patient_id` — UUID FK → `patients.id` NOT NULL
- `name` — TEXT NOT NULL
- `relationship` — TEXT NOT NULL — `spouse`, `parent`, `child`, `sibling`, `friend`, `other`
- `phone` — TEXT NOT NULL
- `email` — TEXT
- `address_line1` — TEXT
- `city` — TEXT
- `state` — TEXT
- `postal_code` — TEXT
- `priority_rank` — INT DEFAULT 1 — 1 = first to contact
- `has_medical_poa` — BOOLEAN DEFAULT false

### `patient_consents`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | patient_id | consent_type | granted | granted_at | expires_at | witness_name | document_ref |
| --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | p0000001-0000-4000-8000-000000000001 | hipaa_notice | true | 2024-01-15T10:00:00Z | 2027-01-15T10:00:00Z | Robert Smith, RN | s3://hospital-docs/consents/sample.pdf |

**Column reference:**

- `id` — UUID PK
- `patient_id` — UUID FK → `patients.id` NOT NULL
- `consent_type` — TEXT NOT NULL — `hipaa_notice`, `treatment`, `release_of_info`, `research`, `telehealth`, `financial`
- `granted` — BOOLEAN NOT NULL
- `granted_at` — TIMESTAMPTZ NOT NULL
- `expires_at` — TIMESTAMPTZ
- `witness_name` — TEXT
- `document_ref` — TEXT — path / URL to signed PDF

---

## Domain 2 — Providers & Organization

### `providers`

Clinical and ancillary staff who can be assigned to encounters / orders.

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | npi | dea_number | first_name | last_name | credentials | provider_type | email | phone | hire_date | termination_date | is_active | primary_department_id |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | 1234567893 | BD1234567 | Jane | Doe | MD | physician | jane.doe@example.com | 555-123-4567 | 2018-06-01 |  | true | dp000001-0000-4000-8000-000000000001 |

**Column reference:**

- `id` — UUID PK
- `npi` — TEXT UNIQUE — 10-digit National Provider Identifier (CMS)
- `dea_number` — TEXT — for prescribers
- `first_name` — TEXT NOT NULL
- `last_name` — TEXT NOT NULL
- `credentials` — TEXT — `MD`, `DO`, `RN`, `PA-C`, `NP`, `PharmD`, …
- `provider_type` — TEXT NOT NULL — `physician`, `nurse`, `np`, `pa`, `pharmacist`, `tech`, `therapist`, `admin`
- `email` — TEXT
- `phone` — TEXT
- `hire_date` — DATE
- `termination_date` — DATE
- `is_active` — BOOLEAN DEFAULT true
- `primary_department_id` — UUID FK → `departments.id`

### `specialties`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | code | name | category |
| --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | 207R00000X | Internal Medicine | physician |

**Column reference:**

- `id` — UUID PK
- `code` — TEXT UNIQUE NOT NULL — NUCC taxonomy code
- `name` — TEXT NOT NULL — `Internal Medicine`, `Cardiology`, …
- `category` — TEXT — `physician`, `nursing`, `allied_health`

### `provider_specialties`

Many-to-many.

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| provider_id | specialty_id | is_primary | board_certified | certification_date |
| --- | --- | --- | --- | --- |
| pr000001-0000-4000-8000-000000000001 | sp000001-0000-4000-8000-000000000001 | true | true | 2018-06-01 |

**Column reference:**

- `provider_id` — UUID FK → `providers.id` NOT NULL — composite PK
- `specialty_id` — UUID FK → `specialties.id` NOT NULL — composite PK
- `is_primary` — BOOLEAN DEFAULT false
- `board_certified` — BOOLEAN DEFAULT false
- `certification_date` — DATE

### `provider_licenses`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | provider_id | license_type | license_number | issuing_state | issue_date | expiration_date | status |
| --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | pr000001-0000-4000-8000-000000000001 | state_medical | MD-CA-987654 | CA | 2018-06-01 | 2025-12-31 | active |

**Column reference:**

- `id` — UUID PK
- `provider_id` — UUID FK → `providers.id` NOT NULL
- `license_type` — TEXT NOT NULL — `state_medical`, `dea`, `controlled_substance`, `board_cert`
- `license_number` — TEXT NOT NULL
- `issuing_state` — TEXT
- `issue_date` — DATE
- `expiration_date` — DATE
- `status` — TEXT — `active`, `expired`, `suspended`, `revoked`

### `locations`

Physical facilities (hospital, clinic, lab draw site).

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | name | facility_type | address_line1 | city | state | postal_code | phone | npi |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | Central General Hospital | hospital | 123 Main St | Springfield | CA | 94105 | 555-123-4567 | 1234567893 |

**Column reference:**

- `id` — UUID PK
- `name` — TEXT NOT NULL
- `facility_type` — TEXT — `hospital`, `clinic`, `urgent_care`, `lab`, `imaging_center`, `pharmacy`
- `address_line1` — TEXT
- `city` — TEXT
- `state` — TEXT
- `postal_code` — TEXT
- `phone` — TEXT
- `npi` — TEXT — facility NPI if applicable

### `departments`

Organizational units inside a location (Cardiology, ED, ICU, Lab).

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | location_id | name | code | department_type | phone |
| --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | lo000001-0000-4000-8000-000000000001 | Cardiology | CARD | inpatient | 555-123-4567 |

**Column reference:**

- `id` — UUID PK
- `location_id` — UUID FK → `locations.id` NOT NULL
- `name` — TEXT NOT NULL
- `code` — TEXT — facility-local
- `department_type` — TEXT — `inpatient`, `outpatient`, `ed`, `or`, `icu`, `lab`, `pharmacy`, `admin`
- `phone` — TEXT

### `units`

Subdivision of a department (e.g. ICU → MICU/SICU/CCU).

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | department_id | name | unit_type | bed_capacity |
| --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | dp000001-0000-4000-8000-000000000001 | MICU | icu | 12 |

**Column reference:**

- `id` — UUID PK
- `department_id` — UUID FK → `departments.id` NOT NULL
- `name` — TEXT NOT NULL
- `unit_type` — TEXT — `med_surg`, `icu`, `step_down`, `telemetry`, `peds`, `nursery`, `obs`
- `bed_capacity` — INT

### `rooms`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | unit_id | room_number | room_type |
| --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | un000001-0000-4000-8000-000000000001 | example | private |

**Column reference:**

- `id` — UUID PK
- `unit_id` — UUID FK → `units.id` NOT NULL
- `room_number` — TEXT NOT NULL
- `room_type` — TEXT — `private`, `semi_private`, `isolation`, `procedure`

### `beds`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | room_id | bed_label | status | is_monitored |
| --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | rm000001-0000-4000-8000-000000000001 | A | active | true |

**Column reference:**

- `id` — UUID PK
- `room_id` — UUID FK → `rooms.id` NOT NULL
- `bed_label` — TEXT NOT NULL — `A`, `B`, `1`, …
- `status` — TEXT — `available`, `occupied`, `cleaning`, `out_of_service`, `reserved`
- `is_monitored` — BOOLEAN DEFAULT false — telemetry-equipped

---

## Domain 3 — Scheduling

### `appointment_types`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | code | name | default_duration_minutes | requires_referral |
| --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | FU30 | Follow-up 30 min | 30 | false |

**Column reference:**

- `id` — UUID PK
- `code` — TEXT UNIQUE NOT NULL — `NP15`, `FU30`, `PHYS`, …
- `name` — TEXT NOT NULL
- `default_duration_minutes` — INT NOT NULL
- `requires_referral` — BOOLEAN DEFAULT false

### `appointment_slots`

Provider availability blocks the scheduler draws from.

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | provider_id | location_id | slot_start | slot_end | is_available | appointment_type_id |
| --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | pr000001-0000-4000-8000-000000000001 | lo000001-0000-4000-8000-000000000001 | 2025-03-14T09:00:00Z | 2025-03-14T09:30:00Z | true | at000001-0000-4000-8000-000000000001 |

**Column reference:**

- `id` — UUID PK
- `provider_id` — UUID FK → `providers.id` NOT NULL
- `location_id` — UUID FK → `locations.id` NOT NULL
- `slot_start` — TIMESTAMPTZ NOT NULL
- `slot_end` — TIMESTAMPTZ NOT NULL
- `is_available` — BOOLEAN DEFAULT true
- `appointment_type_id` — UUID FK → `appointment_types.id` — optional restriction

### `appointments`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | patient_id | provider_id | appointment_type_id | location_id | department_id | scheduled_start | scheduled_end | actual_start | actual_end | status | reason_for_visit | referral_id | created_by_provider_id |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | p0000001-0000-4000-8000-000000000001 | pr000001-0000-4000-8000-000000000001 | at000001-0000-4000-8000-000000000001 | lo000001-0000-4000-8000-000000000001 | dp000001-0000-4000-8000-000000000001 | 2025-03-14T09:00:00Z | 2025-03-14T09:30:00Z | 2025-03-14T09:02:00Z | 2025-03-14T09:28:00Z | active | Annual physical | rf000001-0000-4000-8000-000000000001 | pr000001-0000-4000-8000-000000000015 |

**Column reference:**

- `id` — UUID PK
- `patient_id` — UUID FK → `patients.id` NOT NULL
- `provider_id` — UUID FK → `providers.id` NOT NULL
- `appointment_type_id` — UUID FK → `appointment_types.id` NOT NULL
- `location_id` — UUID FK → `locations.id` NOT NULL
- `department_id` — UUID FK → `departments.id`
- `scheduled_start` — TIMESTAMPTZ NOT NULL
- `scheduled_end` — TIMESTAMPTZ NOT NULL
- `actual_start` — TIMESTAMPTZ
- `actual_end` — TIMESTAMPTZ
- `status` — TEXT NOT NULL — `scheduled`, `checked_in`, `in_progress`, `completed`, `cancelled`, `no_show`, `rescheduled`
- `reason_for_visit` — TEXT
- `referral_id` — UUID FK → `referrals.id`
- `created_by_provider_id` — UUID FK → `providers.id`

### `appointment_status_history`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | appointment_id | from_status | to_status | changed_at | changed_by_provider_id | reason |
| --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | ap000001-0000-4000-8000-000000000001 | scheduled | checked_in | 2025-03-14T08:45:00Z | a1b2c3d4-1111-4222-8333-444455556666 | appointment_reminder |

**Column reference:**

- `id` — UUID PK
- `appointment_id` — UUID FK → `appointments.id` NOT NULL
- `from_status` — TEXT
- `to_status` — TEXT NOT NULL
- `changed_at` — TIMESTAMPTZ NOT NULL
- `changed_by_provider_id` — UUID FK → `providers.id`
- `reason` — TEXT

### `referrals`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | patient_id | referring_provider_id | referred_to_provider_id | referred_to_specialty_id | reason | priority | status | created_at_date | expires_on | authorization_id |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | p0000001-0000-4000-8000-000000000001 | pr000001-0000-4000-8000-000000000025 | pr000001-0000-4000-8000-000000000026 | sp000001-0000-4000-8000-000000000002 | appointment_reminder | routine | active | 2025-03-14 | 2025-09-30 | au000001-0000-4000-8000-000000000001 |

**Column reference:**

- `id` — UUID PK
- `patient_id` — UUID FK → `patients.id` NOT NULL
- `referring_provider_id` — UUID FK → `providers.id` NOT NULL
- `referred_to_provider_id` — UUID FK → `providers.id` — nullable, may be specialty only
- `referred_to_specialty_id` — UUID FK → `specialties.id`
- `reason` — TEXT NOT NULL
- `priority` — TEXT — `routine`, `urgent`, `stat`
- `status` — TEXT — `requested`, `scheduled`, `completed`, `cancelled`, `expired`
- `created_at_date` — DATE NOT NULL
- `expires_on` — DATE
- `authorization_id` — UUID FK → `authorizations.id`

### `waitlist_entries`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | patient_id | provider_id | appointment_type_id | requested_after | requested_before | priority | status |
| --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | p0000001-0000-4000-8000-000000000001 | pr000001-0000-4000-8000-000000000001 | at000001-0000-4000-8000-000000000001 | 2025-03-14 | 2025-03-14 | routine | active |

**Column reference:**

- `id` — UUID PK
- `patient_id` — UUID FK → `patients.id` NOT NULL
- `provider_id` — UUID FK → `providers.id`
- `appointment_type_id` — UUID FK → `appointment_types.id`
- `requested_after` — DATE
- `requested_before` — DATE
- `priority` — TEXT — `routine`, `urgent`
- `status` — TEXT — `waiting`, `offered`, `accepted`, `declined`, `expired`

---

## Domain 4 — Encounters (ADT)

The **encounter** ties patient + provider + location + time. Admission/Discharge/Transfer (ADT) movements live alongside.

### `encounters`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | patient_id | encounter_class | status | admission_type | admission_source | chief_complaint | attending_provider_id | admitting_provider_id | location_id | department_id | appointment_id | admitted_at | discharged_at | discharge_disposition | triage_acuity |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | p0000001-0000-4000-8000-000000000001 | ambulatory | active | elective | physician_referral | Chest pain | pr000001-0000-4000-8000-000000000001 | pr000001-0000-4000-8000-000000000002 | lo000001-0000-4000-8000-000000000001 | dp000001-0000-4000-8000-000000000001 | ap000001-0000-4000-8000-000000000001 | 2025-03-14T08:00:00Z | 2025-03-16T11:00:00Z | home | 3 |

**Column reference:**

- `id` — UUID PK
- `patient_id` — UUID FK → `patients.id` NOT NULL
- `encounter_class` — TEXT NOT NULL — `ambulatory`, `inpatient`, `emergency`, `observation`, `home_health`, `virtual`, `inpatient_rehab`
- `status` — TEXT NOT NULL — `planned`, `arrived`, `in_progress`, `discharged`, `cancelled`
- `admission_type` — TEXT — `elective`, `urgent`, `emergency`, `newborn`, `trauma`
- `admission_source` — TEXT — `physician_referral`, `clinic_referral`, `transfer`, `ed`, `court_law`
- `chief_complaint` — TEXT
- `attending_provider_id` — UUID FK → `providers.id`
- `admitting_provider_id` — UUID FK → `providers.id`
- `location_id` — UUID FK → `locations.id` NOT NULL
- `department_id` — UUID FK → `departments.id`
- `appointment_id` — UUID FK → `appointments.id` — nullable for walk-ins
- `admitted_at` — TIMESTAMPTZ
- `discharged_at` — TIMESTAMPTZ
- `discharge_disposition` — TEXT — `home`, `home_health`, `snf`, `rehab`, `expired`, `ama`, `transfer_acute`, `hospice`
- `triage_acuity` — INT — ESI 1-5 (1 = most acute)

### `bed_assignments`

ADT movement history.

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | encounter_id | bed_id | assigned_at | released_at | assigned_by_provider_id | reason |
| --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | e0000001-0000-4000-8000-000000000001 | bd000001-0000-4000-8000-000000000001 | 2025-03-14T08:15:00Z | 2025-03-16T11:00:00Z | pr000001-0000-4000-8000-000000000013 | appointment_reminder |

**Column reference:**

- `id` — UUID PK
- `encounter_id` — UUID FK → `encounters.id` NOT NULL
- `bed_id` — UUID FK → `beds.id` NOT NULL
- `assigned_at` — TIMESTAMPTZ NOT NULL
- `released_at` — TIMESTAMPTZ — NULL = currently occupied
- `assigned_by_provider_id` — UUID FK → `providers.id`
- `reason` — TEXT — `admit`, `transfer`, `discharge`

### `encounter_diagnoses`

Links encounters to ICD-10 codes with ranking.

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | encounter_id | icd10_code | diagnosis_type | present_on_admission | rank | documented_by_provider_id | documented_at |
| --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | e0000001-0000-4000-8000-000000000001 | E11.9 | principal | Y | 1 | pr000001-0000-4000-8000-000000000011 | 2025-03-14T10:30:00Z |

**Column reference:**

- `id` — UUID PK
- `encounter_id` — UUID FK → `encounters.id` NOT NULL
- `icd10_code` — TEXT NOT NULL — FK → `icd10_codes.code`
- `diagnosis_type` — TEXT NOT NULL — `admitting`, `principal`, `secondary`, `discharge`, `working`
- `present_on_admission` — TEXT — `Y`, `N`, `U`, `W`
- `rank` — INT — 1 = principal
- `documented_by_provider_id` — UUID FK → `providers.id`
- `documented_at` — TIMESTAMPTZ

### `encounter_procedures`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | encounter_id | cpt_code | performed_at | performing_provider_id | assistant_provider_id | location_id | modifier | notes |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | e0000001-0000-4000-8000-000000000001 | 99213 | 2025-03-14T11:00:00Z | pr000001-0000-4000-8000-000000000005 | pr000001-0000-4000-8000-000000000006 | lo000001-0000-4000-8000-000000000001 | 26 | Stable, follow-up in 3 months. |

**Column reference:**

- `id` — UUID PK
- `encounter_id` — UUID FK → `encounters.id` NOT NULL
- `cpt_code` — TEXT NOT NULL — FK → `cpt_codes.code`
- `performed_at` — TIMESTAMPTZ NOT NULL
- `performing_provider_id` — UUID FK → `providers.id`
- `assistant_provider_id` — UUID FK → `providers.id`
- `location_id` — UUID FK → `locations.id`
- `modifier` — TEXT — CPT modifier(s), comma-separated
- `notes` — TEXT

---

## Domain 5 — Clinical

### `icd10_codes` (catalog)

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| code | description | chapter | is_billable |
| --- | --- | --- | --- |
| E11.9 | Example description | IV. Endocrine, nutritional and metabolic diseases | true |

**Column reference:**

- `code` — TEXT PK — e.g. `E11.9`
- `description` — TEXT NOT NULL
- `chapter` — TEXT
- `is_billable` — BOOLEAN DEFAULT true

### `cpt_codes` (catalog)

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| code | description | category | default_charge |
| --- | --- | --- | --- |
| 99213 | Example description | general | 125.00 |

**Column reference:**

- `code` — TEXT PK — e.g. `99213`
- `description` — TEXT NOT NULL
- `category` — TEXT — `E/M`, `surgery`, `radiology`, `pathology`, `medicine`, `hcpcs`
- `default_charge` — NUMERIC(12,2)

### `loinc_codes` (catalog)

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| code | display | class | system | units_default |
| --- | --- | --- | --- | --- |
| 718-7 | Hemoglobin [Mass/volume] in Blood | HEM/BC | Ser/Plas | g/dL |

**Column reference:**

- `code` — TEXT PK — e.g. `718-7`
- `display` — TEXT NOT NULL — `Hemoglobin [Mass/volume] in Blood`
- `class` — TEXT — `HEM/BC`, `CHEM`, …
- `system` — TEXT — `Bld`, `Ser/Plas`, `Urine`
- `units_default` — TEXT — `g/dL`

### `snomed_codes` (catalog, optional)

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| code | display | semantic_tag |
| --- | --- | --- |
| 44054006 | Diabetes mellitus (disorder) | disorder |

**Column reference:**

- `code` — TEXT PK
- `display` — TEXT NOT NULL
- `semantic_tag` — TEXT — `disorder`, `finding`, `procedure`, `body structure`

### `problem_list_entries`

A patient's longitudinal problem list (distinct from encounter diagnoses).

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | patient_id | icd10_code | snomed_code | description | clinical_status | verification_status | severity | onset_date | resolved_date | recorded_by_provider_id |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | p0000001-0000-4000-8000-000000000001 | E11.9 | 44054006 | Example description | active | confirmed | moderate | 2022-09-15 |  | pr000001-0000-4000-8000-000000000012 |

**Column reference:**

- `id` — UUID PK
- `patient_id` — UUID FK → `patients.id` NOT NULL
- `icd10_code` — TEXT FK → `icd10_codes.code`
- `snomed_code` — TEXT FK → `snomed_codes.code`
- `description` — TEXT NOT NULL
- `clinical_status` — TEXT NOT NULL — `active`, `recurrence`, `relapse`, `inactive`, `remission`, `resolved`
- `verification_status` — TEXT — `unconfirmed`, `provisional`, `confirmed`, `refuted`, `entered_in_error`
- `severity` — TEXT — `mild`, `moderate`, `severe`
- `onset_date` — DATE
- `resolved_date` — DATE
- `recorded_by_provider_id` — UUID FK → `providers.id`

### `allergies`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | patient_id | allergen_type | allergen_name | rxnorm_code | criticality | clinical_status | verification_status | recorded_date | last_occurrence_date |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | p0000001-0000-4000-8000-000000000001 | medication | Penicillin | 1191 | low | active | confirmed | 2024-11-02 | 2024-11-02 |

**Column reference:**

- `id` — UUID PK
- `patient_id` — UUID FK → `patients.id` NOT NULL
- `allergen_type` — TEXT NOT NULL — `medication`, `food`, `environment`, `biologic`
- `allergen_name` — TEXT NOT NULL
- `rxnorm_code` — TEXT FK → `rxnorm_concepts.rxcui` — nullable
- `criticality` — TEXT — `low`, `high`, `unable_to_assess`
- `clinical_status` — TEXT — `active`, `inactive`, `resolved`
- `verification_status` — TEXT — `confirmed`, `unconfirmed`, `refuted`
- `recorded_date` — DATE
- `last_occurrence_date` — DATE

### `allergy_reactions`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | allergy_id | manifestation | severity | onset_minutes |
| --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | al000001-0000-4000-8000-000000000001 | rash | moderate | 30 |

**Column reference:**

- `id` — UUID PK
- `allergy_id` — UUID FK → `allergies.id` NOT NULL
- `manifestation` — TEXT NOT NULL — `hives`, `anaphylaxis`, `rash`, `nausea`, `swelling`, …
- `severity` — TEXT — `mild`, `moderate`, `severe`
- `onset_minutes` — INT — minutes after exposure

### `vital_signs`

Discrete observations (latest blood pressure, HR, temp, SpO₂, pain).

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | patient_id | encounter_id | measured_at | systolic_bp | diastolic_bp | heart_rate | respiratory_rate | temperature_c | spo2 | pain_score | height_cm | weight_kg | bmi | recorded_by_provider_id |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | p0000001-0000-4000-8000-000000000001 | e0000001-0000-4000-8000-000000000001 | 2025-03-14T08:30:00Z | 118 | 76 | 72 | 16 | 36.80 | 98 | 2 | 168.00 | 65.50 | 23.20 | pr000001-0000-4000-8000-000000000012 |

**Column reference:**

- `id` — UUID PK
- `patient_id` — UUID FK → `patients.id` NOT NULL
- `encounter_id` — UUID FK → `encounters.id`
- `measured_at` — TIMESTAMPTZ NOT NULL
- `systolic_bp` — INT — mmHg
- `diastolic_bp` — INT — mmHg
- `heart_rate` — INT — bpm
- `respiratory_rate` — INT — breaths/min
- `temperature_c` — NUMERIC(4,2) — °C
- `spo2` — INT — %
- `pain_score` — INT — 0-10
- `height_cm` — NUMERIC(5,2)
- `weight_kg` — NUMERIC(5,2)
- `bmi` — NUMERIC(4,2)
- `recorded_by_provider_id` — UUID FK → `providers.id`

### `clinical_observations`

Generic LOINC-coded observations beyond vitals.

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | patient_id | encounter_id | loinc_code | value_numeric | value_text | units | interpretation | observed_at | recorded_by_provider_id |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | p0000001-0000-4000-8000-000000000001 | e0000001-0000-4000-8000-000000000001 | 718-7 | 13.5000 | Negative | g/dL | N | 2025-03-14T08:30:00Z | pr000001-0000-4000-8000-000000000012 |

**Column reference:**

- `id` — UUID PK
- `patient_id` — UUID FK → `patients.id` NOT NULL
- `encounter_id` — UUID FK → `encounters.id`
- `loinc_code` — TEXT FK → `loinc_codes.code` NOT NULL
- `value_numeric` — NUMERIC(12,4)
- `value_text` — TEXT
- `units` — TEXT — UCUM
- `interpretation` — TEXT — `N`, `H`, `L`, `A`, `AA`
- `observed_at` — TIMESTAMPTZ NOT NULL
- `recorded_by_provider_id` — UUID FK → `providers.id`

### `care_plans`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | patient_id | title | category | status | start_date | end_date | responsible_provider_id | description |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | p0000001-0000-4000-8000-000000000001 | Review CBC result | general | active | 2025-01-01 | 2025-04-01 | pr000001-0000-4000-8000-000000000016 | Example description |

**Column reference:**

- `id` — UUID PK
- `patient_id` — UUID FK → `patients.id` NOT NULL
- `title` — TEXT NOT NULL
- `category` — TEXT — `assess_plan`, `careteam`, `chronic_disease_management`, `discharge`
- `status` — TEXT — `draft`, `active`, `completed`, `on_hold`, `cancelled`
- `start_date` — DATE
- `end_date` — DATE
- `responsible_provider_id` — UUID FK → `providers.id`
- `description` — TEXT

### `care_plan_goals`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | care_plan_id | description | target_date | achievement_status | priority |
| --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | cp000001-0000-4000-8000-000000000001 | Example description | 2025-06-01 | in_progress | routine |

**Column reference:**

- `id` — UUID PK
- `care_plan_id` — UUID FK → `care_plans.id` NOT NULL
- `description` — TEXT NOT NULL
- `target_date` — DATE
- `achievement_status` — TEXT — `in_progress`, `achieved`, `not_achieved`, `sustaining`
- `priority` — TEXT — `low`, `medium`, `high`

---

## Domain 6 — Medications

### `rxnorm_concepts` (catalog)

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| rxcui | name | tty | route | is_controlled | dea_schedule |
| --- | --- | --- | --- | --- | --- |
| 1191 | Amoxicillin 500 MG Oral Tablet | SCD | oral | false | CIV |

**Column reference:**

- `rxcui` — TEXT PK — RxNorm CUI
- `name` — TEXT NOT NULL
- `tty` — TEXT — term type — `SCD`, `SBD`, `IN`, `BN`
- `route` — TEXT — `oral`, `iv`, `im`, `subq`, `topical`
- `is_controlled` — BOOLEAN DEFAULT false
- `dea_schedule` — TEXT — `CII`-`CV`

### `medications` (formulary)

A drug as the facility stocks / prescribes it.

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | rxcui | ndc | name | strength | dosage_form | route | manufacturer | is_on_formulary |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | 1191 | 00093-0832-01 | Amoxicillin 500 mg oral tablet | 500 mg | tablet | oral | Acme Pharma | true |

**Column reference:**

- `id` — UUID PK
- `rxcui` — TEXT FK → `rxnorm_concepts.rxcui`
- `ndc` — TEXT — National Drug Code, 10/11-digit
- `name` — TEXT NOT NULL
- `strength` — TEXT — `500 mg`, `5 mg/mL`
- `dosage_form` — TEXT — `tablet`, `capsule`, `injection`, `solution`, `suspension`, `inhaler`
- `route` — TEXT
- `manufacturer` — TEXT
- `is_on_formulary` — BOOLEAN DEFAULT true

### `prescriptions`

(a.k.a. `medication_orders`).

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | patient_id | prescriber_provider_id | encounter_id | medication_id | dose | route | frequency | duration_days | quantity | refills | start_date | end_date | status | indication | prn_reason | pharmacy_id | is_electronic |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | p0000001-0000-4000-8000-000000000001 | pr000001-0000-4000-8000-000000000004 | e0000001-0000-4000-8000-000000000001 | md000001-0000-4000-8000-000000000001 | 500 mg | oral | bid | 10 | 30.00 | 2 | 2025-01-01 | 2025-04-01 | active | Hypertension | For pain greater than 4/10 | ph000001-0000-4000-8000-000000000001 | true |

**Column reference:**

- `id` — UUID PK
- `patient_id` — UUID FK → `patients.id` NOT NULL
- `prescriber_provider_id` — UUID FK → `providers.id` NOT NULL
- `encounter_id` — UUID FK → `encounters.id`
- `medication_id` — UUID FK → `medications.id` NOT NULL
- `dose` — TEXT NOT NULL — `500 mg`
- `route` — TEXT NOT NULL
- `frequency` — TEXT NOT NULL — `q8h`, `bid`, `prn`, `daily`
- `duration_days` — INT
- `quantity` — NUMERIC(10,2)
- `refills` — INT DEFAULT 0
- `start_date` — DATE NOT NULL
- `end_date` — DATE
- `status` — TEXT NOT NULL — `active`, `completed`, `cancelled`, `stopped`, `on_hold`, `entered_in_error`
- `indication` — TEXT — reason for use
- `prn_reason` — TEXT
- `pharmacy_id` — UUID FK → `pharmacies.id`
- `is_electronic` — BOOLEAN DEFAULT true — e-prescription

### `medication_administrations`

The MAR — discrete dose events for inpatient/inhospital meds.

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | prescription_id | patient_id | encounter_id | administered_at | administered_by_provider_id | dose_given | route | status | not_done_reason | site |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | rx000001-0000-4000-8000-000000000001 | p0000001-0000-4000-8000-000000000001 | e0000001-0000-4000-8000-000000000001 | 2025-03-14T09:15:00Z | pr000001-0000-4000-8000-000000000007 | 500 mg | oral | active | patient_refused | left_deltoid |

**Column reference:**

- `id` — UUID PK
- `prescription_id` — UUID FK → `prescriptions.id` NOT NULL
- `patient_id` — UUID FK → `patients.id` NOT NULL — denorm for fast read
- `encounter_id` — UUID FK → `encounters.id`
- `administered_at` — TIMESTAMPTZ NOT NULL
- `administered_by_provider_id` — UUID FK → `providers.id` NOT NULL — usually RN
- `dose_given` — TEXT NOT NULL
- `route` — TEXT
- `status` — TEXT NOT NULL — `completed`, `not_done`, `held`, `refused`, `in_progress`
- `not_done_reason` — TEXT — `patient_refused`, `npo`, `vomiting`, `unavailable`, `other`
- `site` — TEXT — `left_deltoid`, `right_gluteus`, `iv_port`

### `medication_reconciliations`

Periodic verification of home/inpatient meds at admit/discharge.

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | patient_id | encounter_id | reconciliation_type | performed_by_provider_id | performed_at | notes |
| --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | p0000001-0000-4000-8000-000000000001 | e0000001-0000-4000-8000-000000000001 | admission | pr000001-0000-4000-8000-000000000010 | 2025-03-14T11:00:00Z | Stable, follow-up in 3 months. |

**Column reference:**

- `id` — UUID PK
- `patient_id` — UUID FK → `patients.id` NOT NULL
- `encounter_id` — UUID FK → `encounters.id` NOT NULL
- `reconciliation_type` — TEXT — `admission`, `transfer`, `discharge`
- `performed_by_provider_id` — UUID FK → `providers.id`
- `performed_at` — TIMESTAMPTZ NOT NULL
- `notes` — TEXT

### `pharmacies`

External fill pharmacies.

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | name | ncpdp_id | npi | phone | fax | address_line1 | city | state | postal_code | is_mail_order |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | Acme Community Pharmacy | 1234567 | 1234567893 | 555-123-4567 | 555-123-4568 | 123 Main St | Springfield | CA | 94105 | false |

**Column reference:**

- `id` — UUID PK
- `name` — TEXT NOT NULL
- `ncpdp_id` — TEXT — NCPDP pharmacy ID
- `npi` — TEXT
- `phone` — TEXT
- `fax` — TEXT
- `address_line1` — TEXT
- `city` — TEXT
- `state` — TEXT
- `postal_code` — TEXT
- `is_mail_order` — BOOLEAN DEFAULT false

---

## Domain 7 — Labs & Imaging

### `lab_orders`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | patient_id | encounter_id | ordering_provider_id | ordered_at | priority | status | panel_loinc_code | clinical_question | fasting_required |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | p0000001-0000-4000-8000-000000000001 | e0000001-0000-4000-8000-000000000001 | pr000001-0000-4000-8000-000000000003 | 2025-03-14T08:35:00Z | routine | active | 57021-8 | Evaluate anemia | false |

**Column reference:**

- `id` — UUID PK
- `patient_id` — UUID FK → `patients.id` NOT NULL
- `encounter_id` — UUID FK → `encounters.id`
- `ordering_provider_id` — UUID FK → `providers.id` NOT NULL
- `ordered_at` — TIMESTAMPTZ NOT NULL
- `priority` — TEXT — `routine`, `urgent`, `stat`, `asap`
- `status` — TEXT NOT NULL — `ordered`, `collected`, `in_lab`, `resulted`, `cancelled`
- `panel_loinc_code` — TEXT FK → `loinc_codes.code` — for whole panels (CBC, BMP)
- `clinical_question` — TEXT
- `fasting_required` — BOOLEAN DEFAULT false

### `lab_specimens`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | lab_order_id | specimen_type | container | collected_at | collected_by_provider_id | received_in_lab_at | volume_ml | is_rejected | rejection_reason |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | lb000001-0000-4000-8000-000000000001 | blood_serum | red_top | 2025-03-14T08:50:00Z | pr000001-0000-4000-8000-000000000008 | 2025-03-14T09:10:00Z | 5.00 | false | Hemolyzed specimen |

**Column reference:**

- `id` — UUID PK
- `lab_order_id` — UUID FK → `lab_orders.id` NOT NULL
- `specimen_type` — TEXT NOT NULL — `blood_serum`, `blood_plasma`, `whole_blood`, `urine`, `csf`, `swab`, `tissue`
- `container` — TEXT — `red_top`, `lavender`, `green`, …
- `collected_at` — TIMESTAMPTZ
- `collected_by_provider_id` — UUID FK → `providers.id`
- `received_in_lab_at` — TIMESTAMPTZ
- `volume_ml` — NUMERIC(6,2)
- `is_rejected` — BOOLEAN DEFAULT false
- `rejection_reason` — TEXT

### `lab_results`

One row per analyte.

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | lab_order_id | lab_specimen_id | loinc_code | value_numeric | value_text | units | reference_range_low | reference_range_high | interpretation | is_critical | resulted_at | verified_by_provider_id | status |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | lb000001-0000-4000-8000-000000000001 | ls000001-0000-4000-8000-000000000001 | 718-7 | 13.5000 | Negative | g/dL | 12.0000 | 16.0000 | N | false | 2025-03-14T10:45:00Z | pr000001-0000-4000-8000-000000000009 | active |

**Column reference:**

- `id` — UUID PK
- `lab_order_id` — UUID FK → `lab_orders.id` NOT NULL
- `lab_specimen_id` — UUID FK → `lab_specimens.id`
- `loinc_code` — TEXT FK → `loinc_codes.code` NOT NULL
- `value_numeric` — NUMERIC(14,4)
- `value_text` — TEXT
- `units` — TEXT — UCUM
- `reference_range_low` — NUMERIC(14,4)
- `reference_range_high` — NUMERIC(14,4)
- `interpretation` — TEXT — `N`, `H`, `L`, `HH`, `LL`, `A`, `AA`
- `is_critical` — BOOLEAN DEFAULT false
- `resulted_at` — TIMESTAMPTZ NOT NULL
- `verified_by_provider_id` — UUID FK → `providers.id`
- `status` — TEXT — `preliminary`, `final`, `corrected`, `cancelled`

### `imaging_orders`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | patient_id | encounter_id | ordering_provider_id | modality | body_part | cpt_code | clinical_indication | priority | status | requires_contrast | ordered_at |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | p0000001-0000-4000-8000-000000000001 | e0000001-0000-4000-8000-000000000001 | pr000001-0000-4000-8000-000000000003 | CT | Abdomen | 99213 | Rule out PE | routine | active | false | 2025-03-14T08:35:00Z |

**Column reference:**

- `id` — UUID PK
- `patient_id` — UUID FK → `patients.id` NOT NULL
- `encounter_id` — UUID FK → `encounters.id`
- `ordering_provider_id` — UUID FK → `providers.id` NOT NULL
- `modality` — TEXT NOT NULL — `XR`, `CT`, `MR`, `US`, `NM`, `PT`, `MG`, `DX`
- `body_part` — TEXT NOT NULL
- `cpt_code` — TEXT FK → `cpt_codes.code`
- `clinical_indication` — TEXT NOT NULL
- `priority` — TEXT — `routine`, `urgent`, `stat`
- `status` — TEXT NOT NULL — `ordered`, `scheduled`, `performed`, `cancelled`, `resulted`
- `requires_contrast` — BOOLEAN DEFAULT false
- `ordered_at` — TIMESTAMPTZ NOT NULL

### `imaging_studies`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | imaging_order_id | study_uid | accession_number | performed_at | performed_by_provider_id | location_id | series_count | image_count | dicom_metadata_doc_id |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | im000001-0000-4000-8000-000000000001 | 1.2.840.113619.2.55.3.1234567890.123.456789 | ACC202503140001 | 2025-03-14T11:00:00Z | pr000001-0000-4000-8000-000000000010 | lo000001-0000-4000-8000-000000000001 | 3 | 200 | 65f1c2a4d3e4b1a2c3d4e5f7 |

**Column reference:**

- `id` — UUID PK
- `imaging_order_id` — UUID FK → `imaging_orders.id` NOT NULL
- `study_uid` — TEXT UNIQUE NOT NULL — DICOM Study Instance UID
- `accession_number` — TEXT UNIQUE
- `performed_at` — TIMESTAMPTZ
- `performed_by_provider_id` — UUID FK → `providers.id` — technologist
- `location_id` — UUID FK → `locations.id`
- `series_count` — INT
- `image_count` — INT
- `dicom_metadata_doc_id` — TEXT — reference to Mongo `imaging_metadata._id`

### `imaging_reports`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | imaging_study_id | reading_radiologist_id | dictated_at | signed_at | findings | impression | recommendation | status |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | is000001-0000-4000-8000-000000000001 | pr000001-0000-4000-8000-000000000027 | 2025-03-14T12:00:00Z | 2025-03-14T12:15:00Z | No acute cardiopulmonary process. | Unremarkable study. | Clinical correlation recommended. | active |

**Column reference:**

- `id` — UUID PK
- `imaging_study_id` — UUID FK → `imaging_studies.id` NOT NULL
- `reading_radiologist_id` — UUID FK → `providers.id`
- `dictated_at` — TIMESTAMPTZ
- `signed_at` — TIMESTAMPTZ
- `findings` — TEXT
- `impression` — TEXT
- `recommendation` — TEXT
- `status` — TEXT — `preliminary`, `final`, `amended`, `addendum`

---

## Domain 8 — Insurance & Billing

### `payers`

Insurance companies / programs.

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | name | payer_type | payer_id_external | address_line1 | city | state | postal_code | phone | claims_phone | claims_fax | electronic_claims_supported |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | BlueCare Health | commercial | 12345 | 123 Main St | Springfield | CA | 94105 | 555-123-4567 | 800-555-0100 | 800-555-0101 | true |

**Column reference:**

- `id` — UUID PK
- `name` — TEXT NOT NULL
- `payer_type` — TEXT NOT NULL — `commercial`, `medicare`, `medicaid`, `tricare`, `va`, `workers_comp`, `self_pay`, `auto`
- `payer_id_external` — TEXT — clearinghouse payer ID
- `address_line1` — TEXT
- `city` — TEXT
- `state` — TEXT
- `postal_code` — TEXT
- `phone` — TEXT
- `claims_phone` — TEXT
- `claims_fax` — TEXT
- `electronic_claims_supported` — BOOLEAN DEFAULT true

### `insurance_plans`

A specific product offered by a payer.

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | payer_id | plan_name | plan_type | metal_tier | group_number | effective_date | termination_date | requires_referrals | requires_prior_auth_for_imaging |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | py000001-0000-4000-8000-000000000001 | BlueCare Plus PPO | PPO | silver | GRP-7788 | 2025-01-01 |  | false | true |

**Column reference:**

- `id` — UUID PK
- `payer_id` — UUID FK → `payers.id` NOT NULL
- `plan_name` — TEXT NOT NULL
- `plan_type` — TEXT — `HMO`, `PPO`, `EPO`, `POS`, `HDHP`, `Medicare_Advantage`, `Medigap`
- `metal_tier` — TEXT — `bronze`, `silver`, `gold`, `platinum`, `catastrophic` (ACA)
- `group_number` — TEXT
- `effective_date` — DATE
- `termination_date` — DATE
- `requires_referrals` — BOOLEAN DEFAULT false
- `requires_prior_auth_for_imaging` — BOOLEAN DEFAULT false

### `patient_coverages`

A patient's enrollment in a plan.

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | patient_id | insurance_plan_id | subscriber_relationship | subscriber_name | subscriber_dob | member_id | group_number | coverage_rank | effective_date | termination_date | copay_pcp | copay_specialist | copay_er | deductible_individual | deductible_family | oop_max_individual | verified_at |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | p0000001-0000-4000-8000-000000000001 | ip000001-0000-4000-8000-000000000001 | self | Jane Doe | 1985-04-12 | MBR123456789 | GRP-7788 | 1 | 2025-01-01 |  | 20.00 | 40.00 | 150.00 | 1500.00 | 3000.00 | 8000.00 | 2025-03-12T09:00:00Z |

**Column reference:**

- `id` — UUID PK
- `patient_id` — UUID FK → `patients.id` NOT NULL
- `insurance_plan_id` — UUID FK → `insurance_plans.id` NOT NULL
- `subscriber_relationship` — TEXT NOT NULL — `self`, `spouse`, `child`, `other`
- `subscriber_name` — TEXT
- `subscriber_dob` — DATE
- `member_id` — TEXT NOT NULL — as printed on card
- `group_number` — TEXT
- `coverage_rank` — INT NOT NULL — 1 = primary, 2 = secondary, 3 = tertiary
- `effective_date` — DATE NOT NULL
- `termination_date` — DATE
- `copay_pcp` — NUMERIC(8,2)
- `copay_specialist` — NUMERIC(8,2)
- `copay_er` — NUMERIC(8,2)
- `deductible_individual` — NUMERIC(10,2)
- `deductible_family` — NUMERIC(10,2)
- `oop_max_individual` — NUMERIC(10,2)
- `verified_at` — TIMESTAMPTZ — last eligibility check

### `authorizations`

Pre-authorizations / referrals required by the payer.

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | patient_id | coverage_id | auth_number | cpt_code | requested_units | approved_units | status | effective_date | expiration_date | requested_at | decided_at | denial_reason |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | p0000001-0000-4000-8000-000000000001 | cv000001-0000-4000-8000-000000000001 | AUTH-2025-000123 | 99213 | 10 | 10 | active | 2025-01-01 | 2025-12-31 | 2025-03-10T09:00:00Z | 2025-03-20T12:00:00Z | Service not covered under plan |

**Column reference:**

- `id` — UUID PK
- `patient_id` — UUID FK → `patients.id` NOT NULL
- `coverage_id` — UUID FK → `patient_coverages.id` NOT NULL
- `auth_number` — TEXT UNIQUE — issued by payer
- `cpt_code` — TEXT FK → `cpt_codes.code`
- `requested_units` — INT — visits / procedures authorized
- `approved_units` — INT
- `status` — TEXT NOT NULL — `requested`, `pending`, `approved`, `denied`, `expired`, `cancelled`
- `effective_date` — DATE
- `expiration_date` — DATE
- `requested_at` — TIMESTAMPTZ
- `decided_at` — TIMESTAMPTZ
- `denial_reason` — TEXT

### `claims`

The 837-style claim submitted to a payer.

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | patient_id | encounter_id | coverage_id | claim_number | payer_claim_id | claim_type | status | service_start_date | service_end_date | total_charge | total_allowed | total_paid | patient_responsibility | submitted_at | paid_at | billing_provider_id |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | p0000001-0000-4000-8000-000000000001 | e0000001-0000-4000-8000-000000000001 | cv000001-0000-4000-8000-000000000001 | CLM-2025-000456 | PAYER-ICN-789 | professional_837p | active | 2025-03-14 | 2025-03-14 | 350.00 | 280.00 | 240.00 | 20.00 | 2025-03-15T12:00:00Z | 2025-03-29T12:00:00Z | pr000001-0000-4000-8000-000000000017 |

**Column reference:**

- `id` — UUID PK
- `patient_id` — UUID FK → `patients.id` NOT NULL
- `encounter_id` — UUID FK → `encounters.id`
- `coverage_id` — UUID FK → `patient_coverages.id` NOT NULL
- `claim_number` — TEXT UNIQUE NOT NULL — internal
- `payer_claim_id` — TEXT — payer's ICN
- `claim_type` — TEXT NOT NULL — `professional_837p`, `institutional_837i`, `dental`, `pharmacy`
- `status` — TEXT NOT NULL — `draft`, `submitted`, `accepted`, `rejected`, `paid`, `partial_paid`, `denied`, `appealed`
- `service_start_date` — DATE NOT NULL
- `service_end_date` — DATE
- `total_charge` — NUMERIC(12,2) NOT NULL
- `total_allowed` — NUMERIC(12,2)
- `total_paid` — NUMERIC(12,2)
- `patient_responsibility` — NUMERIC(12,2)
- `submitted_at` — TIMESTAMPTZ
- `paid_at` — TIMESTAMPTZ
- `billing_provider_id` — UUID FK → `providers.id`

### `claim_lines`

Service-line detail (one per CPT performed).

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | claim_id | line_number | cpt_code | modifier | icd10_pointer | service_date | units | charge_amount | allowed_amount | paid_amount | adjustment_amount | denial_code | place_of_service | rendering_provider_id |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | cl000001-0000-4000-8000-000000000001 | 1 | 99213 | 26 | 1 | 2025-03-14 | 1.00 | 150.00 | 120.00 | 100.00 | 20.00 | 45 | 11 | pr000001-0000-4000-8000-000000000019 |

**Column reference:**

- `id` — UUID PK
- `claim_id` — UUID FK → `claims.id` NOT NULL
- `line_number` — INT NOT NULL
- `cpt_code` — TEXT FK → `cpt_codes.code` NOT NULL
- `modifier` — TEXT
- `icd10_pointer` — TEXT — comma-separated diagnosis pointers `1,2`
- `service_date` — DATE NOT NULL
- `units` — NUMERIC(8,2) NOT NULL DEFAULT 1
- `charge_amount` — NUMERIC(12,2) NOT NULL
- `allowed_amount` — NUMERIC(12,2)
- `paid_amount` — NUMERIC(12,2)
- `adjustment_amount` — NUMERIC(12,2)
- `denial_code` — TEXT — CARC / RARC
- `place_of_service` — TEXT — CMS POS code (`11` office, `21` inpatient, `23` ED)
- `rendering_provider_id` — UUID FK → `providers.id`

### `charges`

Pre-claim charge capture (driven by encounter activity).

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | encounter_id | patient_id | cpt_code | quantity | charge_amount | posted_at | posted_by_provider_id | claim_id |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | e0000001-0000-4000-8000-000000000001 | p0000001-0000-4000-8000-000000000001 | 99213 | 1.00 | 150.00 | 2025-03-14T13:00:00Z | pr000001-0000-4000-8000-000000000018 | cl000001-0000-4000-8000-000000000001 |

**Column reference:**

- `id` — UUID PK
- `encounter_id` — UUID FK → `encounters.id` NOT NULL
- `patient_id` — UUID FK → `patients.id` NOT NULL
- `cpt_code` — TEXT FK → `cpt_codes.code` NOT NULL
- `quantity` — NUMERIC(8,2) DEFAULT 1
- `charge_amount` — NUMERIC(12,2) NOT NULL
- `posted_at` — TIMESTAMPTZ NOT NULL
- `posted_by_provider_id` — UUID FK → `providers.id`
- `claim_id` — UUID FK → `claims.id` — nullable until bundled

### `payments`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | claim_id | patient_id | payer_id | payment_type | payment_method | amount | received_at | reference_number | era_835_id |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | cl000001-0000-4000-8000-000000000001 | p0000001-0000-4000-8000-000000000001 | py000001-0000-4000-8000-000000000001 | insurance | eft | 120.00 | 2025-03-29T12:00:00Z | CHK-100245 | ERA-2025-03-14-001 |

**Column reference:**

- `id` — UUID PK
- `claim_id` — UUID FK → `claims.id` — nullable for patient payments not tied to a claim
- `patient_id` — UUID FK → `patients.id` NOT NULL
- `payer_id` — UUID FK → `payers.id` — nullable for patient payments
- `payment_type` — TEXT NOT NULL — `insurance`, `patient`, `adjustment`, `refund`, `writeoff`
- `payment_method` — TEXT — `eft`, `check`, `cash`, `credit_card`, `ach`
- `amount` — NUMERIC(12,2) NOT NULL
- `received_at` — TIMESTAMPTZ NOT NULL
- `reference_number` — TEXT — check #, EFT trace, transaction ID
- `era_835_id` — TEXT — for payer remittances

### `adjustments`

CARC-coded adjustments (write-offs, contractual).

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | claim_line_id | adjustment_group | reason_code | amount | note |
| --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | cn000001-0000-4000-8000-000000000001 | CO | 45 | 120.00 | Contractual write-off per fee schedule. |

**Column reference:**

- `id` — UUID PK
- `claim_line_id` — UUID FK → `claim_lines.id` NOT NULL
- `adjustment_group` — TEXT — `CO`, `PR`, `OA`, `PI`
- `reason_code` — TEXT NOT NULL — CARC (e.g. `45`)
- `amount` — NUMERIC(12,2) NOT NULL
- `note` — TEXT

### `patient_statements`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | patient_id | statement_date | period_start | period_end | previous_balance | charges_total | payments_total | adjustments_total | current_balance | due_date | status | delivery_method |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | p0000001-0000-4000-8000-000000000001 | 2025-04-01 | 2025-03-01 | 2025-03-31 | 0.00 | 350.00 | 240.00 | 20.00 | 90.00 | 2025-04-15 | active | portal |

**Column reference:**

- `id` — UUID PK
- `patient_id` — UUID FK → `patients.id` NOT NULL
- `statement_date` — DATE NOT NULL
- `period_start` — DATE
- `period_end` — DATE
- `previous_balance` — NUMERIC(12,2)
- `charges_total` — NUMERIC(12,2)
- `payments_total` — NUMERIC(12,2)
- `adjustments_total` — NUMERIC(12,2)
- `current_balance` — NUMERIC(12,2) NOT NULL
- `due_date` — DATE
- `status` — TEXT — `draft`, `sent`, `paid`, `in_collections`
- `delivery_method` — TEXT — `paper`, `email`, `portal`

---

## Domain 9 — Communications

### `patient_message_threads`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | patient_id | subject | category | status | priority | assigned_provider_id | last_message_at |
| --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | p0000001-0000-4000-8000-000000000001 | Lab results available | general | active | routine | pr000001-0000-4000-8000-000000000014 | 2025-03-14T15:00:00Z |

**Column reference:**

- `id` — UUID PK
- `patient_id` — UUID FK → `patients.id` NOT NULL
- `subject` — TEXT
- `category` — TEXT — `medical_question`, `rx_refill`, `appointment`, `billing`, `lab_result`, `referral`, `other`
- `status` — TEXT — `open`, `awaiting_provider`, `awaiting_patient`, `closed`
- `priority` — TEXT — `routine`, `urgent`
- `assigned_provider_id` — UUID FK → `providers.id`
- `last_message_at` — TIMESTAMPTZ

### `patient_messages`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | thread_id | channel | direction | sender_provider_id | sender_patient_id | body | sent_at | read_at | attachment_uri |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | th000001-0000-4000-8000-000000000001 | portal | outbound | pr000001-0000-4000-8000-000000000022 | p0000001-0000-4000-8000-000000000001 | Your lab results are now available in the portal. | 2025-03-14T15:00:00Z | 2025-03-14T16:00:00Z | s3://hospital-docs/attachments/sample.pdf |

**Column reference:**

- `id` — UUID PK
- `thread_id` — UUID FK → `patient_message_threads.id` NOT NULL
- `channel` — TEXT NOT NULL — `portal`, `sms`, `email`, `app_push`
- `direction` — TEXT NOT NULL — `inbound`, `outbound`
- `sender_provider_id` — UUID FK → `providers.id` — NULL for patient-sent
- `sender_patient_id` — UUID FK → `patients.id` — NULL for provider-sent
- `body` — TEXT NOT NULL
- `sent_at` — TIMESTAMPTZ NOT NULL
- `read_at` — TIMESTAMPTZ
- `attachment_uri` — TEXT

### `call_logs`

Phone calls (inbound & outbound) with patients.

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | patient_id | provider_id | direction | phone_number | started_at | duration_seconds | outcome | reason | summary | follow_up_required |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | p0000001-0000-4000-8000-000000000001 | pr000001-0000-4000-8000-000000000001 | outbound | 555-123-4567 | 2025-03-14T11:00:00Z | 180 | connected | appointment_reminder | Patient reached; reminder confirmed. | false |

**Column reference:**

- `id` — UUID PK
- `patient_id` — UUID FK → `patients.id` NOT NULL
- `provider_id` — UUID FK → `providers.id`
- `direction` — TEXT NOT NULL — `inbound`, `outbound`
- `phone_number` — TEXT NOT NULL
- `started_at` — TIMESTAMPTZ NOT NULL
- `duration_seconds` — INT
- `outcome` — TEXT — `connected`, `voicemail`, `no_answer`, `wrong_number`
- `reason` — TEXT — `appointment_reminder`, `result_callback`, `triage`, `rx_question`, `billing`
- `summary` — TEXT
- `follow_up_required` — BOOLEAN DEFAULT false

### `appointment_reminders`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | appointment_id | patient_id | channel | scheduled_send_at | sent_at | delivery_status | response |
| --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | ap000001-0000-4000-8000-000000000001 | p0000001-0000-4000-8000-000000000001 | portal | 2025-03-13T18:00:00Z | 2025-03-14T15:00:00Z | delivered | confirmed |

**Column reference:**

- `id` — UUID PK
- `appointment_id` — UUID FK → `appointments.id` NOT NULL
- `patient_id` — UUID FK → `patients.id` NOT NULL
- `channel` — TEXT NOT NULL — `sms`, `email`, `voice`, `portal`
- `scheduled_send_at` — TIMESTAMPTZ NOT NULL
- `sent_at` — TIMESTAMPTZ
- `delivery_status` — TEXT — `queued`, `sent`, `delivered`, `failed`, `bounced`
- `response` — TEXT — `confirmed`, `cancel_requested`, `reschedule_requested`, `none`

### `insurance_correspondences`

All payer ↔ provider correspondence (EOBs, denials, authorizations, appeals).

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | payer_id | patient_id | claim_id | authorization_id | direction | correspondence_type | channel | received_at | sent_at | subject | body | document_uri | requires_follow_up | follow_up_by |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | py000001-0000-4000-8000-000000000001 | p0000001-0000-4000-8000-000000000001 | cl000001-0000-4000-8000-000000000001 | au000001-0000-4000-8000-000000000001 | outbound | eob | portal | 2025-03-29T12:00:00Z | 2025-03-14T15:00:00Z | Lab results available | Your lab results are now available in the portal. | s3://hospital-docs/correspondence/sample.pdf | false | 2025-04-15 |

**Column reference:**

- `id` — UUID PK
- `payer_id` — UUID FK → `payers.id` NOT NULL
- `patient_id` — UUID FK → `patients.id`
- `claim_id` — UUID FK → `claims.id`
- `authorization_id` — UUID FK → `authorizations.id`
- `direction` — TEXT NOT NULL — `inbound`, `outbound`
- `correspondence_type` — TEXT NOT NULL — `eob`, `denial`, `auth_request`, `auth_response`, `appeal`, `request_for_info`, `eligibility_verification`
- `channel` — TEXT — `fax`, `mail`, `email`, `portal`, `phone`, `edi_277`, `edi_835`
- `received_at` — TIMESTAMPTZ
- `sent_at` — TIMESTAMPTZ
- `subject` — TEXT
- `body` — TEXT
- `document_uri` — TEXT
- `requires_follow_up` — BOOLEAN DEFAULT false
- `follow_up_by` — DATE

### `claim_denials`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | claim_id | denial_date | carc_code | rarc_code | denial_reason | is_appealable | appeal_deadline | worked_by_provider_id | status |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | cl000001-0000-4000-8000-000000000001 | 2025-04-02 | 45 | N130 | Service not covered under plan | true | 2025-07-02 | pr000001-0000-4000-8000-000000000021 | active |

**Column reference:**

- `id` — UUID PK
- `claim_id` — UUID FK → `claims.id` NOT NULL
- `denial_date` — DATE NOT NULL
- `carc_code` — TEXT NOT NULL
- `rarc_code` — TEXT
- `denial_reason` — TEXT
- `is_appealable` — BOOLEAN DEFAULT true
- `appeal_deadline` — DATE
- `worked_by_provider_id` — UUID FK → `providers.id` — denials coordinator
- `status` — TEXT — `new`, `working`, `appealed`, `resolved`, `written_off`

### `claim_appeals`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | claim_denial_id | appeal_level | submitted_at | submitted_by_provider_id | narrative | outcome | decided_at | recovered_amount |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | cd000001-0000-4000-8000-000000000001 | 1 | 2025-03-15T12:00:00Z | pr000001-0000-4000-8000-000000000020 | Submitting medical records demonstrating medical necessity. | connected | 2025-03-20T12:00:00Z | 150.00 |

**Column reference:**

- `id` — UUID PK
- `claim_denial_id` — UUID FK → `claim_denials.id` NOT NULL
- `appeal_level` — INT NOT NULL DEFAULT 1 — 1 / 2 / 3
- `submitted_at` — TIMESTAMPTZ NOT NULL
- `submitted_by_provider_id` — UUID FK → `providers.id`
- `narrative` — TEXT
- `outcome` — TEXT — `pending`, `overturned`, `upheld`, `partial`
- `decided_at` — TIMESTAMPTZ
- `recovered_amount` — NUMERIC(12,2)

### `inter_provider_messages`

Internal consults / handoffs.

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | from_provider_id | to_provider_id | patient_id | encounter_id | message_type | subject | body | sent_at | read_at | acknowledged_at |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | pr000001-0000-4000-8000-000000000023 | pr000001-0000-4000-8000-000000000024 | p0000001-0000-4000-8000-000000000001 | e0000001-0000-4000-8000-000000000001 | consult_request | Lab results available | Your lab results are now available in the portal. | 2025-03-14T15:00:00Z | 2025-03-14T16:00:00Z | 2025-03-14T16:05:00Z |

**Column reference:**

- `id` — UUID PK
- `from_provider_id` — UUID FK → `providers.id` NOT NULL
- `to_provider_id` — UUID FK → `providers.id` NOT NULL
- `patient_id` — UUID FK → `patients.id`
- `encounter_id` — UUID FK → `encounters.id`
- `message_type` — TEXT — `consult_request`, `consult_reply`, `handoff`, `curbside`, `fyi`
- `subject` — TEXT
- `body` — TEXT NOT NULL
- `sent_at` — TIMESTAMPTZ NOT NULL
- `read_at` — TIMESTAMPTZ
- `acknowledged_at` — TIMESTAMPTZ

---

## Domain 10 — Operations & Audit

### `shifts`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | name | start_time | end_time | crosses_midnight |
| --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | Day | 07:00:00 | 19:00:00 | false |

**Column reference:**

- `id` — UUID PK
- `name` — TEXT NOT NULL — `Day`, `Evening`, `Night`, `Weekend Days`
- `start_time` — TIME NOT NULL
- `end_time` — TIME NOT NULL
- `crosses_midnight` — BOOLEAN DEFAULT false

### `staff_schedules`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | provider_id | department_id | unit_id | shift_id | work_date | scheduled_start | scheduled_end | role | status |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | pr000001-0000-4000-8000-000000000001 | dp000001-0000-4000-8000-000000000001 | un000001-0000-4000-8000-000000000001 | sh000001-0000-4000-8000-000000000001 | 2025-03-14 | 2025-03-14T09:00:00Z | 2025-03-14T09:30:00Z | attending | active |

**Column reference:**

- `id` — UUID PK
- `provider_id` — UUID FK → `providers.id` NOT NULL
- `department_id` — UUID FK → `departments.id`
- `unit_id` — UUID FK → `units.id`
- `shift_id` — UUID FK → `shifts.id`
- `work_date` — DATE NOT NULL
- `scheduled_start` — TIMESTAMPTZ NOT NULL
- `scheduled_end` — TIMESTAMPTZ NOT NULL
- `role` — TEXT — `attending`, `resident`, `charge_nurse`, `bedside_rn`, `tech`
- `status` — TEXT — `scheduled`, `worked`, `pto`, `sick`, `swap`, `cancelled`

### `on_call_assignments`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | provider_id | department_id | specialty_id | on_call_start | on_call_end | pager_number |
| --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | pr000001-0000-4000-8000-000000000001 | dp000001-0000-4000-8000-000000000001 | sp000001-0000-4000-8000-000000000001 | 2025-03-14T19:00:00Z | 2025-03-15T07:00:00Z | 555-987-6543 |

**Column reference:**

- `id` — UUID PK
- `provider_id` — UUID FK → `providers.id` NOT NULL
- `department_id` — UUID FK → `departments.id`
- `specialty_id` — UUID FK → `specialties.id`
- `on_call_start` — TIMESTAMPTZ NOT NULL
- `on_call_end` — TIMESTAMPTZ NOT NULL
- `pager_number` — TEXT

### `pharmacy_inventory`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | medication_id | location_id | lot_number | expiration_date | quantity_on_hand | reorder_level | unit_cost | last_restocked_at |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | md000001-0000-4000-8000-000000000001 | lo000001-0000-4000-8000-000000000001 | LOT-AB1234 | 2025-12-31 | 100.00 | 20.00 | 1.2500 | 2025-03-01T08:00:00Z |

**Column reference:**

- `id` — UUID PK
- `medication_id` — UUID FK → `medications.id` NOT NULL
- `location_id` — UUID FK → `locations.id` NOT NULL
- `lot_number` — TEXT
- `expiration_date` — DATE
- `quantity_on_hand` — NUMERIC(12,2) NOT NULL
- `reorder_level` — NUMERIC(12,2)
- `unit_cost` — NUMERIC(10,4)
- `last_restocked_at` — TIMESTAMPTZ

### `equipment`

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | name | equipment_type | manufacturer | model | serial_number | location_id | unit_id | status | last_maintenance_at | next_maintenance_due |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | Bedside Vitals Monitor | monitor | Acme Pharma | Vitalmon X3 | SN-2024-00042 | lo000001-0000-4000-8000-000000000001 | un000001-0000-4000-8000-000000000001 | active | 2025-01-10T08:00:00Z | 2025-07-10 |

**Column reference:**

- `id` — UUID PK
- `name` — TEXT NOT NULL
- `equipment_type` — TEXT — `monitor`, `ventilator`, `infusion_pump`, `defibrillator`, `imaging`, `wheelchair`, `bed`
- `manufacturer` — TEXT
- `model` — TEXT
- `serial_number` — TEXT UNIQUE
- `location_id` — UUID FK → `locations.id`
- `unit_id` — UUID FK → `units.id`
- `status` — TEXT — `in_use`, `available`, `maintenance`, `retired`
- `last_maintenance_at` — TIMESTAMPTZ
- `next_maintenance_due` — DATE

### `tasks`

Workflow inbox items (RN tasks, billing follow-ups, denials, etc.).

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | assigned_provider_id | created_by_provider_id | patient_id | encounter_id | task_type | title | description | priority | status | due_at | completed_at |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | pr000001-0000-4000-8000-000000000014 | pr000001-0000-4000-8000-000000000015 | p0000001-0000-4000-8000-000000000001 | e0000001-0000-4000-8000-000000000001 | result_callback | Review CBC result | Example description | routine | active | 2025-03-20T17:00:00Z |  |

**Column reference:**

- `id` — UUID PK
- `assigned_provider_id` — UUID FK → `providers.id`
- `created_by_provider_id` — UUID FK → `providers.id`
- `patient_id` — UUID FK → `patients.id`
- `encounter_id` — UUID FK → `encounters.id`
- `task_type` — TEXT NOT NULL — `lab_review`, `result_callback`, `rx_refill`, `prior_auth`, `denial_followup`, `chart_review`, `signature`, `general`
- `title` — TEXT NOT NULL
- `description` — TEXT
- `priority` — TEXT — `low`, `normal`, `high`, `urgent`
- `status` — TEXT — `open`, `in_progress`, `completed`, `cancelled`
- `due_at` — TIMESTAMPTZ
- `completed_at` — TIMESTAMPTZ

### `audit_logs_summary`

Lightweight summary of HIPAA access events; the heavy raw log goes to Mongo `audit_logs`.

**Example row** (columns shown as columns; values are illustrative — replace when authoring the xlsx):

| id | actor_provider_id | patient_id | action | resource_type | resource_id | occurred_at | success | mongo_log_id |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| a1b2c3d4-1111-4222-8333-444455556666 | pr000001-0000-4000-8000-000000000028 | p0000001-0000-4000-8000-000000000001 | view | lab_result | rs000001-0000-4000-8000-000000000001 | 2025-03-14T08:30:00Z | true | 65f1c2a4d3e4b1a2c3d4e5f6 |

**Column reference:**

- `id` — UUID PK
- `actor_provider_id` — UUID FK → `providers.id`
- `patient_id` — UUID FK → `patients.id`
- `action` — TEXT NOT NULL — `view`, `create`, `update`, `delete`, `print`, `export`
- `resource_type` — TEXT NOT NULL — `patient`, `encounter`, `lab_result`, `note`, …
- `resource_id` — UUID
- `occurred_at` — TIMESTAMPTZ NOT NULL
- `success` — BOOLEAN DEFAULT true
- `mongo_log_id` — TEXT — reference to `audit_logs._id`

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
