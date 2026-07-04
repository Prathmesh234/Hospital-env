# AGENTS.md — Hospital-env briefing

You are an operations/analytics agent working the terminal of **Saint Raphael
Medical Center (SRMC)** — a fully synthetic hospital information system
(PostgreSQL + MongoDB + a read-only REST API). Your job is to answer
questions about the hospital's data. Everything you need is reachable through
the commands below; **all access is read-only** and writes are blocked at the
tool layer.

This file is your operating manual. The extended version (with a data-model
walkthrough) is `docs/agent_interface.md`; the table-by-table schema spec is
`docs/schema.md`.

---

## 1. Bootstrap (run this first)

```bash
./agent_entrypoint.sh
```

Idempotent — safe to rerun. It starts the databases (Docker), creates the
73-table schema, loads ~1,069 seed rows, derives 109 Mongo documents, and
starts the REST API on `http://localhost:8000`. If a command later fails with
a connection error, rerun the entrypoint.

Sanity check afterwards:

```bash
./he status                 # both databases reachable?
./he agent tables | head    # 73 tables with row counts
```

> **Always use `./he`** (the repo-root launcher), not `hospital-env` or
> `uv run hospital-env` — the console script can break on iCloud-synced
> macOS checkouts. `./he` always works.

---

## 2. Tools

Every command is one-shot and non-interactive — right for agent harnesses.
**Do not use `./he agent shell`** (it's an interactive REPL and will hang a
non-interactive session).

| command | purpose |
| --- | --- |
| `./he agent tables [--domain billing]` | list tables, domain, live row count |
| `./he agent describe <table>` | columns, types, PK/FK, indexes, referenced-by |
| `./he agent sample <table> [--limit 5]` | peek at rows |
| `./he agent sql "SELECT …" [-f json\|csv] [--limit N]` | one read-only SQL query |
| `./he agent mongo` | list document collections + counts |
| `./he agent mongo <coll> --filter '{…}' [--count] [--project '{…}'] [--sort f:desc]` | read-only find |
| `./he agent api /path [-p key=value]` | GET a REST endpoint (catalog at /docs) |

SQL constraints: a single `SELECT`/`WITH`/`TABLE`/`EXPLAIN` statement; write
and DDL keywords are rejected; 15 s server-side timeout; results truncate at
`--limit` (default 50, cap 1000) **and say so** — if you see "truncated",
raise `--limit` instead of trusting a clipped result. Use `-f csv` or
`-f json` when you want to parse output.

Domains for `--domain`: `catalogs`, `organization`, `patients`, `scheduling`,
`encounters`, `clinical`, `medications`, `labs_imaging`, `billing`,
`communications`, `operations`.

Mongo collections: `clinical_notes`, `audit_logs`, `imaging_metadata`,
`vitals_streams`. Their `patient_id` / `encounter_id` / `*_provider_id`
fields are **string UUIDs** pointing at PostgreSQL rows — join back with
`WHERE id = CAST('<uuid>' AS uuid)`.

---

## 3. Evaluation tasks (the actual assignment)

Questions arrive as task IDs like `HOSP-016`:

```bash
./he task list                              # id / category / difficulty / answer type
./he task show HOSP-016                     # read the question
# …investigate with the section-2 tools…
./he task submit HOSP-016 --answer "59.7"   # exit code 0 = correct, 1 = wrong
```

Answer formats (the grader is forgiving, but stay simple):

| answer_type | example submission | notes |
| --- | --- | --- |
| `number` | `59.7`, `$38,420.00`, `162 mmHg` | currency/units stripped, small tolerance |
| `string` | `Blackwell` | case/whitespace-insensitive; answer exactly what's asked (e.g. last name only) |
| `date` | `2025-02-04` | ISO preferred; common formats accepted |
| `boolean` | `yes` / `no` | |
| `string_set` | `CT, US, XR` | comma-separated, order-free |
| `ordered_list` | `Chen, Kim, O'Connor` | comma-separated, order matters |

Submit only the answer value — no sentences, no explanations, no markdown.

### Integrity rules (non-negotiable during evaluated runs)

*These three rules bind agents being **evaluated** on tasks. A maintainer
session explicitly asked to develop/fix the environment or the eval harness
is exempt.*

1. **Never read anything under `evals/`** — `evals/tasks.jsonl` contains the
   gold answers and reference solutions. Deriving answers from the databases
   is the entire task; reading the answer key is a failed run.
2. Don't edit code, schema, or data. The tools are read-only; don't try to
   route around them (no `docker exec … psql`, no editing files).
3. One submission per task unless told otherwise; check the exit code.

---

## 4. Data model in ten lines

- `patients` is the hub (MRNs like `SRMC-00010001`); `encounters` is the
  clinical spine — `encounter_class` ∈ `ambulatory|emergency|inpatient|virtual`,
  with `admitted_at`/`discharged_at`/`discharge_disposition`. An ED visit that
  turns into an admission is **two** encounter rows.
- Clinical: `problem_list_entries` (ICD-10), `allergies`, `vital_signs`,
  `prescriptions` → `medication_administrations`, `lab_orders` →
  `lab_specimens` → `lab_results` (LOINC), `imaging_orders` →
  `imaging_studies` → `imaging_reports`.
- Revenue cycle: `charges` → `claims` (→ `encounters`, `patient_coverages`)
  → `claim_lines`; money in `payments`, write-offs in `adjustments`
  (groups `CO`/`PR`); disputes in `claim_denials` → `claim_appeals`.
- Coverage: `payers` → `insurance_plans` → `patient_coverages`
  (`coverage_rank` 1 = primary, 2 = secondary).
- Operations: `staff_schedules`, `on_call_assignments`, `pharmacy_inventory`,
  `equipment`, `tasks`.

Gotchas:

- `created_at`/`updated_at` are load-time bookkeeping — **never** use them to
  answer questions; use domain timestamps (`admitted_at`, `resulted_at`,
  `posted_at`, `ordered_at`, …). All timestamps are UTC.
- Claim money is denormalized: `claims.total_charge/total_paid` vs per-line
  amounts on `claim_lines` — use the level the question names.
- `lab_results.interpretation` is `N`/`H`/`L`; critical is the separate
  boolean `is_critical`.
- Names repeat across roles (there is a patient Chen *and* a provider Chen) —
  join through IDs, not names, whenever possible.

---

## 5. Recommended workflow per task

1. `./he task show <ID>` — read carefully; note the answer type and exactly
   what's being asked (full name vs last name, date vs timestamp, etc.).
2. Locate the tables: `./he agent tables --domain <guess>`, then
   `./he agent describe <table>` for join keys (FKs are listed).
3. Build the query incrementally — sample first, then filter, then aggregate.
   For Mongo questions use `--count` or a tight `--filter`; for cross-store
   questions collect the Mongo side first, then map IDs in SQL with `CAST`.
4. Sanity-check the result (row counts, obvious duplicates, NULLs).
5. `./he task submit <ID> --answer "<value>"` and check the exit code.

## 6. Troubleshooting

| symptom | fix |
| --- | --- |
| `connection refused` on SQL/Mongo | `./agent_entrypoint.sh` (docker containers not up) |
| REST API errors | ignore it — SQL/Mongo answer everything; or rerun the entrypoint |
| `ModuleNotFoundError: setup` | you invoked the console script; use `./he` |
| query times out | add filters/aggregation; the 15 s timeout is intentional |
| result says *truncated* | rerun with a bigger `--limit`, or aggregate in SQL |
