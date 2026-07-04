# Hospital-env

A deliberately complex, synthesizer-style **hospital environment** for agentic / RL workloads. It models a realistic Hospital Information System (HIS) backed by **PostgreSQL** (relational core), with an optional **MongoDB** side-store for document-shaped data (clinical notes, audit logs, imaging metadata, IoT vitals streams).

Inspired by the [Prime-Intellect `general-agent`](https://www.primeintellect.ai/blog/general-agent) "task family" pattern — a stateful database, a tool API over it, and a verifier — but scaled out to a realistic, *very* complicated EHR/EMR + revenue-cycle schema.

> ⚠️ **No real PHI.** This repository ships **schemas only**. You bring the synthetic data via an `.xlsx` workbook; the loader writes it into the schema.

---

## What's in the box

| Layer | Purpose |
| --- | --- |
| **PostgreSQL schema** | 73 tables across 10 domains: patients, providers, encounters, clinical, medications, labs, imaging, billing, insurance, communications, scheduling, pharmacy, staff, audit. |
| **MongoDB collections** | 4 document collections for things that don't relationalize cleanly: clinical notes, audit logs, imaging metadata, IoT vitals streams. |
| **SQLAlchemy 2.0 ORM** | `setup/models/` — typed `Mapped[...]` models grouped by domain. |
| **Pydantic v2 schemas** | `setup/schemas/` — request/response models that mirror the ORM. |
| **FastAPI app** | `setup/api/` — REST endpoints for every domain (read-mostly), suitable as an agent tool surface. |
| **Docker Compose** | One-shot `postgres:16` + `mongo:7` spin-up. |
| **xlsx loader** | `setup/ingest/xlsx_loader.py` — one sheet per table, header row = column names, each row inserted. |
| **Mongo seeder** | `setup/nosql/seed_docs.py` — deterministically derives the 4 document collections from the loaded relational rows. |
| **Typer CLI** | `hospital-env init-db`, `hospital-env load <file.xlsx>`, `hospital-env serve`, etc. |
| **Agent toolset** | `hospital-env agent …` — guarded read-only SQL, schema inspection, Mongo find, REST client, interactive shell. |
| **Eval harness** | `evals/` — 31 questions with verified gold answers, a forgiving grader, and `hospital-env task …` to deal/submit/verify. |

The complete table-by-table specification — what the xlsx workbook needs to look like — lives in **[`docs/schema.md`](docs/schema.md)**. Use that file as the authoritative reference when you author the seed workbook.

---

## Quick start

Requires **Python ≥ 3.11**, **[uv](https://docs.astral.sh/uv/)**, and **Docker** (for the databases).

```bash
# 1. clone & install
git clone https://github.com/Prathmesh234/Hospital-env.git
cd Hospital-env
uv sync

# 2. spin up Postgres + Mongo
cp .env.example .env
docker compose up -d

# 3. create the schema (DDL only — no data)
uv run hospital-env init-db

# 4. load your synthetic data workbook
uv run hospital-env load path/to/hospital_seed.xlsx

# 5. serve the API
uv run hospital-env serve
# → http://localhost:8000/docs
```

**Or let the agent entrypoint do all of it:**

```bash
./agent_entrypoint.sh              # docker up + schema + seeds + mongo docs + API
./agent_entrypoint.sh HOSP-016     # …and print an eval task question
```

---

## Agent access (terminal)

> Running an agent harness (opencode, Claude Code, aider, …) against this
> repo? The root [`AGENTS.md`](AGENTS.md) is the drop-in briefing most
> harnesses auto-load — it contains the bootstrap, the tool reference, the
> eval protocol, and the integrity rules.

Agents work the environment through a read-only tool surface (full guide:
[`docs/agent_interface.md`](docs/agent_interface.md)). The `./he` launcher is
equivalent to `uv run hospital-env`, but immune to iCloud/`.pth` weirdness:

```bash
./he agent tables --domain billing        # tables + live row counts
./he agent describe claims                # columns, FKs, indexes
./he agent sql "SELECT count(*) FROM patients"
./he agent mongo audit_logs --filter '{"action": "export"}' --count
./he agent api /patients --param limit=5
./he agent shell                          # interactive REPL
```

The SQL tool accepts a single `SELECT`/`WITH` statement, runs it in a
server-side **read-only** transaction with a statement timeout, and rejects
write/DDL keywords outright.

## Evaluation harness

31 questions across 16 categories and 3 difficulty tiers live in
[`evals/tasks.jsonl`](evals/tasks.jsonl), each with a gold answer plus a
reference solution (`gold_sql`, or a `gold_check` spec for the Mongo side —
including one cross-store question that requires joining Mongo actors back to
PostgreSQL providers).

```bash
./he task list                              # the question bank (no answers)
./he task show HOSP-020                     # deal a question to an agent
./he task submit HOSP-020 -a '$38,420'      # grade (exit 0 = correct)
./he task grade-file answers.json           # batch scoreboard
./he task check                             # re-derive all golds from live DBs
./he task export-questions -o questions.jsonl   # agent-safe split (no golds)
```

Gold answers are **verified, not asserted**: `./he task check` (and
`tests/test_gold_answers.py`) re-executes every reference solution against the
live databases and fails on drift.

---

## Project layout

```
Hospital-env/
├── docker-compose.yml          # postgres + mongo
├── pyproject.toml              # uv-managed deps
├── .env.example                # DB connection strings
├── docs/
│   └── schema.md               # ⭐ THE schema spec (use this to build the xlsx)
└── setup/
    ├── config.py               # pydantic-settings (reads .env)
    ├── cli.py                  # typer entry point
    ├── db/
    │   ├── postgres.py         # SQLAlchemy engine + session
    │   └── mongo.py            # Motor / PyMongo client
    ├── models/                 # SQLAlchemy ORM, one file per domain
    ├── schemas/                # Pydantic v2, one file per domain
    ├── nosql/                  # Mongo collection definitions + indexes
    ├── api/                    # FastAPI routers, one per domain
    └── ingest/
        └── xlsx_loader.py      # workbook → rows → INSERT
```

---

## Schema reference

> 📖 See **[docs/schema.md](docs/schema.md)** for the full table-by-table breakdown, FK graph, code-system references (ICD-10, CPT, LOINC, RxNorm, SNOMED CT), and the exact xlsx sheet/column conventions.

High-level domains:

1. **Patients & Demographics** — patient master, addresses, contacts, emergency contacts, identifiers
2. **Providers & Organization** — physicians, nurses, departments, locations, beds, units
3. **Scheduling** — appointments, resource slots, waitlists, no-shows
4. **Encounters** — visits, admissions, transfers, discharges (ADT)
5. **Clinical** — problem list, diagnoses (ICD-10), procedures (CPT), allergies, vitals
6. **Medications** — formulary, prescriptions, administrations (MAR), reconciliation
7. **Labs & Imaging** — orders, specimens, results (LOINC), imaging studies, reports
8. **Insurance & Billing** — payers, coverage, authorizations, claims, charges, payments
9. **Communications** — patient messages, payer correspondence, inter-provider consults
10. **Operations & Audit** — staff scheduling, pharmacy inventory, HIPAA access logs

---

## License

MIT
