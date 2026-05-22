# Hospital seed data summary


## 01_catalogs

| Table | Rows |
|---|---:|
| `icd10_codes` | 31 |
| `cpt_codes` | 28 |
| `loinc_codes` | 35 |
| `snomed_codes` | 15 |
| `rxnorm_concepts` | 22 |
| **subtotal** | **131** |

## 02_organization

| Table | Rows |
|---|---:|
| `locations` | 3 |
| `specialties` | 15 |
| `shifts` | 4 |
| `appointment_types` | 10 |
| `payers` | 7 |
| `pharmacies` | 4 |
| `departments` | 10 |
| `units` | 7 |
| `rooms` | 41 |
| `beds` | 61 |
| `providers` | 17 |
| `provider_specialties` | 15 |
| `provider_licenses` | 18 |
| **subtotal** | **212** |

## 03_patients

| Table | Rows |
|---|---:|
| `patients` | 10 |
| `patient_identifiers` | 14 |
| `patient_addresses` | 11 |
| `patient_contacts` | 17 |
| `emergency_contacts` | 12 |
| `patient_consents` | 14 |
| **subtotal** | **78** |

## 04_coverage

| Table | Rows |
|---|---:|
| `insurance_plans` | 8 |
| `patient_coverages` | 12 |
| `authorizations` | 4 |
| **subtotal** | **24** |

## 05_scheduling

| Table | Rows |
|---|---:|
| `appointment_slots` | 41 |
| `referrals` | 6 |
| `appointments` | 10 |
| `appointment_status_history` | 9 |
| `waitlist_entries` | 3 |
| `appointment_reminders` | 6 |
| **subtotal** | **75** |

## 06_encounters

| Table | Rows |
|---|---:|
| `encounters` | 17 |
| `bed_assignments` | 7 |
| `encounter_diagnoses` | 29 |
| `encounter_procedures` | 34 |
| **subtotal** | **87** |

## 07_clinical

| Table | Rows |
|---|---:|
| `problem_list_entries` | 19 |
| `allergies` | 8 |
| `allergy_reactions` | 8 |
| `vital_signs` | 15 |
| `clinical_observations` | 14 |
| `care_plans` | 8 |
| `care_plan_goals` | 13 |
| **subtotal** | **85** |

## 08_medications

| Table | Rows |
|---|---:|
| `medications` | 22 |
| `prescriptions` | 23 |
| `medication_administrations` | 18 |
| `medication_reconciliations` | 8 |
| **subtotal** | **71** |

## 09_labs_imaging

| Table | Rows |
|---|---:|
| `lab_orders` | 12 |
| `lab_specimens` | 14 |
| `lab_results` | 33 |
| `imaging_orders` | 5 |
| `imaging_studies` | 5 |
| `imaging_reports` | 5 |
| **subtotal** | **74** |

## 10_billing

| Table | Rows |
|---|---:|
| `claims` | 12 |
| `claim_lines` | 28 |
| `charges` | 28 |
| `payments` | 11 |
| `adjustments` | 20 |
| `claim_denials` | 1 |
| `claim_appeals` | 1 |
| `patient_statements` | 6 |
| **subtotal** | **107** |

## 11_communications

| Table | Rows |
|---|---:|
| `patient_message_threads` | 6 |
| `patient_messages` | 13 |
| `call_logs` | 6 |
| `insurance_correspondences` | 6 |
| `inter_provider_messages` | 5 |
| **subtotal** | **36** |

## 12_operations

| Table | Rows |
|---|---:|
| `staff_schedules` | 43 |
| `on_call_assignments` | 6 |
| `pharmacy_inventory` | 12 |
| `equipment` | 8 |
| `tasks` | 10 |
| `audit_logs_summary` | 10 |
| **subtotal** | **89** |

**Grand total rows: 1069**
