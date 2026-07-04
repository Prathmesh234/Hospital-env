# Agent interface guide

This document is the operating manual for an **agent working the hospital
environment through a terminal**. It covers the entry point, every tool on the
surface, the shape of the data, and the evaluation protocol.

The fiction: you are an operations/analytics agent at **Saint Raphael Medical
Center (SRMC)**, a synthetic hospital. Every patient, provider, claim, and lab
value is fake — but the schema and workflows are modeled on a real HIS/EHR +
revenue-cycle stack, and questions routinely require chaining 2–4 tables (or
crossing into the document store) to answer.

---

## 1. Entry point

```bash
./agent_entrypoint.sh              # bootstrap everything + print the briefing
./agent_entrypoint.sh HOSP-016     # …and print the assigned task question
```

The script is idempotent. It starts PostgreSQL + MongoDB (Docker), creates the
73-table schema, loads the seed workbook (~1,069 rows), derives the Mongo
document collections deterministically, and starts the read-only REST API on
`http://localhost:8000`.

After bootstrap, every interaction goes through the `./he` launcher (a wrapper
for `uv run hospital-env`):

```bash
./he <command> …
```

---

## 2. Tool surface (all read-only)

### 2.1 Schema discovery

```bash
./he agent tables                    # all 73 tables, domain + live row count
./he agent tables --domain billing   # one domain at a time
./he agent describe claims           # columns, types, PK/FK, indexes, referenced-by
./he agent sample vital_signs        # peek at a few rows (--limit N)
```

Domains: `catalogs`, `organization`, `patients`, `scheduling`, `encounters`,
`clinical`, `medications`, `labs_imaging`, `billing`, `communications`,
`operations`.

### 2.2 SQL (PostgreSQL)

```bash
./he agent sql "SELECT count(*) FROM patients"
./he agent sql "SELECT …" --format json     # or csv; default is a table
./he agent sql "SELECT …" --limit 200       # default 50, cap 1000
```

Constraints enforced by the tool:

- Single statement, must start with `SELECT`, `WITH`, `TABLE`, or `EXPLAIN`.
- Write/DDL keywords are rejected outright; the transaction is additionally
  `READ ONLY` server-side and always rolled back.
- Server-side `statement_timeout` (default 15 s).
- Results are truncated at the row limit and the truncation is reported —
  raise `--limit` rather than trusting a clipped aggregate.

### 2.3 MongoDB (document store)

```bash
./he agent mongo                                             # collections + counts
./he agent mongo clinical_notes --filter '{"note_type": "ed_triage_note"}'
./he agent mongo audit_logs --filter '{"action": "export"}' --count
./he agent mongo vitals_streams --project '{"metric": 1, "device_id": 1}' \
    --sort started_at:desc --limit 5
```

Collections:

| collection | contents | links back to Postgres via |
| --- | --- | --- |
| `clinical_notes` | free-text progress/triage/clinic notes | `patient_id`, `encounter_id`, `author_provider_id` |
| `audit_logs` | HIPAA access events (view/update/print/export) | `patient_id`, `actor_provider_id` |
| `imaging_metadata` | DICOM-ish study metadata + series | `imaging_study_id`, `patient_id`, `study_uid` |
| `vitals_streams` | device telemetry (HR/SpO₂/RR samples) | `patient_id`, `encounter_id` |

All ID references are stored as **strings** of the Postgres UUIDs — cast when
joining back: `… WHERE id = CAST('<uuid>' AS uuid)`.

### 2.4 REST API

```bash
./he agent api /health
./he agent api /patients --param limit=5
./he agent api /patients/SRMC-00010002        # see /docs for the catalog
```

The API is a convenience read surface (pre-joined views, pagination); anything
it can answer, SQL can too.

### 2.5 Interactive shell

```bash
./he agent shell
```

Type SQL terminated by `;`. Meta commands: `\t [domain]` tables, `\d <table>`
describe, `\s <table> [n]` sample, `\m [coll] [filter]` mongo, `\api <path>`,
`\f table|json|csv` output format, `\h` help, `\q` quit.

---

## 3. Data model orientation

The full table-by-table spec is in [`docs/schema.md`](schema.md). The mental
model, in one pass:

- **patients** is the hub (`mrn` like `SRMC-00010001`); demographics satellites
  hang off it (addresses, contacts, identifiers, consents).
- **encounters** is the clinical spine: `encounter_class` ∈ ambulatory /
  emergency / inpatient / virtual, with `admitted_at` / `discharged_at` /
  `discharge_disposition`. ED-to-inpatient admissions appear as *two* rows.
- Clinical facts hang off patient + encounter: `problem_list_entries` (ICD-10),
  `allergies`, `vital_signs`, `prescriptions` → `medication_administrations`,
  `lab_orders` → `lab_specimens` → `lab_results` (LOINC), `imaging_orders` →
  `imaging_studies` → `imaging_reports`.
- **Revenue cycle**: `charges` roll into `claims` (linked to `encounters` and
  `patient_coverages`), which break into `claim_lines`; money lands as
  `payments` and writes off as `adjustments` (groups `CO`/`PR`); disputes live
  in `claim_denials` → `claim_appeals`; patients get `patient_statements`.
- **Coverage**: `payers` → `insurance_plans` → `patient_coverages`
  (`coverage_rank` 1 = primary, 2 = secondary).
- **Operations**: `staff_schedules`, `on_call_assignments`,
  `pharmacy_inventory` (lots + expiration), `equipment`, `tasks`.

Gotchas worth knowing:

- Money on `claims` (`total_charge`, `total_paid`, `patient_responsibility`)
  is denormalized from `claim_lines` — prefer the level the question names.
- `lab_results.interpretation` uses `N`/`H`/`L`; criticality is the separate
  boolean `is_critical`.
- Timestamps are UTC; several questions hinge on `::date` truncation.
- `created_at`/`updated_at` are load-time bookkeeping — never use them to
  answer clinical questions; use the domain timestamps (`admitted_at`,
  `resulted_at`, `posted_at`, …).

---

## 4. Evaluation protocol

Questions are dealt as task IDs (the "query" an agent receives):

```bash
./he task list                      # id / category / difficulty / type
./he task show HOSP-016             # the full question
# …investigate with the tools above…
./he task submit HOSP-016 --answer "59.7"    # exit 0 = correct, 1 = wrong
```

Answer formats (the grader is deliberately forgiving):

| answer_type | submit like | notes |
| --- | --- | --- |
| `number` | `59.7`, `$38,420.00`, `162 mmHg` | units/currency stripped; small tolerance |
| `string` | `Blackwell` | case/whitespace-insensitive |
| `date` | `2025-02-04`, `Feb 4, 2025` | common formats accepted |
| `boolean` | `yes` / `no` / `true` / `false` | |
| `string_set` | `CT, US, XR` | order-free; `,` `;` or newlines |
| `ordered_list` | `Chen, Kim, O'Connor` | order matters |

Batch grading for harnesses:

```bash
./he task grade-file answers.json   # {"HOSP-001": "10", …} or JSONL
```

**Integrity rule:** the dataset file `evals/tasks.jsonl` contains the gold
answers and reference SQL. Agents must not read it (or `evals/`) during a run —
answers must come from the databases. For hard isolation, evaluators can hand
agents only the output of:

```bash
./he task export-questions -o questions.jsonl   # no golds, no reference SQL
```

Maintainers can re-verify every gold answer against the live data at any time:

```bash
./he task check          # re-derives all 31 golds; exits non-zero on drift
```

---

## 5. Resetting the environment

```bash
./he drop-db --yes       # drop schema + collections
./agent_entrypoint.sh    # rebuild from seeds (deterministic, ~30 s)
```

Postgres data survives container restarts (named volume). The Mongo documents
are regenerated from a fixed RNG seed, so gold answers are stable across
rebuilds.
