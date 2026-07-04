#!/usr/bin/env bash
# agent_entrypoint.sh — bootstrap the hospital environment for an agent session.
#
# Idempotent: safe to run repeatedly. Brings up the databases, creates the
# schema, loads the relational seed workbook and the derived Mongo documents,
# starts the REST API in the background, then prints the tool reference.
#
# Usage:
#   ./agent_entrypoint.sh              # set up + print the agent briefing
#   ./agent_entrypoint.sh HOSP-016     # …and end with that task's question
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$REPO_ROOT"
HE="$REPO_ROOT/he"
API_LOG="$REPO_ROOT/.hospital-api.log"
API_PID="$REPO_ROOT/.hospital-api.pid"

say()  { printf '\033[1;36m[entrypoint]\033[0m %s\n' "$*"; }
fail() { printf '\033[1;31m[entrypoint]\033[0m %s\n' "$*" >&2; exit 1; }

# ── 0. prerequisites ─────────────────────────────────────────────────────────
command -v docker >/dev/null 2>&1 || fail "docker is required but not on PATH."
[[ -f .env ]] || { say "creating .env from .env.example"; cp .env.example .env; }

if [[ ! -x .venv/bin/python ]]; then
    command -v uv >/dev/null 2>&1 || fail "uv is required to create the venv (https://docs.astral.sh/uv/)."
    say "installing python dependencies (uv sync)…"
    uv sync --quiet
fi

# macOS + iCloud-synced folders: fileproviderd flags .pth files hidden, which
# makes CPython skip them. Harmless elsewhere; ./he sidesteps this anyway.
if command -v chflags >/dev/null 2>&1; then
    chflags nohidden .venv/lib/python*/site-packages/*.pth 2>/dev/null || true
fi

# ── 1. databases ─────────────────────────────────────────────────────────────
say "starting postgres + mongo (docker compose up -d)…"
docker compose up -d --wait 2>/dev/null || docker compose up -d

say "waiting for postgres…"
for _ in $(seq 1 60); do
    if docker exec hospital-postgres pg_isready -U "${POSTGRES_USER:-hospital}" \
        -d "${POSTGRES_DB:-hospital}" >/dev/null 2>&1; then
        break
    fi
    sleep 1
done
docker exec hospital-postgres pg_isready -U "${POSTGRES_USER:-hospital}" \
    -d "${POSTGRES_DB:-hospital}" >/dev/null 2>&1 || fail "postgres never became ready."

# ── 2. schema + data (idempotent) ────────────────────────────────────────────
say "ensuring schema exists…"
"$HE" init-db >/dev/null

PATIENT_COUNT="$("$HE" agent sql "SELECT count(*) FROM patients" --format csv 2>/dev/null | tail -1 || echo 0)"
if [[ "${PATIENT_COUNT:-0}" -eq 0 ]]; then
    say "loading relational seed workbook…"
    "$HE" load seeds/hospital_seed.xlsx >/dev/null
else
    say "relational data already loaded (${PATIENT_COUNT} patients)."
fi

say "seeding mongo document collections (deterministic)…"
"$HE" seed-mongo >/dev/null

# ── 3. REST API in the background ────────────────────────────────────────────
if ! "$HE" agent api /health >/dev/null 2>&1; then
    say "starting REST API on :8000 (log: .hospital-api.log)…"
    nohup .venv/bin/python -m uvicorn setup.api.main:app \
        --host 127.0.0.1 --port 8000 >"$API_LOG" 2>&1 &
    echo $! >"$API_PID"
    for _ in $(seq 1 20); do
        "$HE" agent api /health >/dev/null 2>&1 && break
        sleep 0.5
    done
fi
"$HE" agent api /health >/dev/null 2>&1 && say "REST API is up." \
    || say "WARNING: REST API did not come up — SQL/Mongo tools still work."

# ── 4. briefing ──────────────────────────────────────────────────────────────
cat <<'BRIEFING'

════════════════════════════════════════════════════════════════════════════
  SAINT RAPHAEL MEDICAL CENTER — hospital information system (synthetic)
════════════════════════════════════════════════════════════════════════════
You are working a terminal against a hospital's data stores:
  • PostgreSQL — 73 relational tables (patients, encounters, meds, labs,
    imaging, claims, scheduling, operations…)
  • MongoDB   — 4 document collections (clinical_notes, audit_logs,
    imaging_metadata, vitals_streams)
  • REST API  — read-only FastAPI at http://localhost:8000 (/docs)

All access is READ-ONLY and goes through `./he` (or `uv run hospital-env`):

  ./he agent tables [--domain billing]     list tables + row counts
  ./he agent describe <table>              columns, FKs, indexes
  ./he agent sample <table> [--limit 5]    peek at rows
  ./he agent sql "SELECT …" [-f json|csv]  one read-only query (SELECT/WITH)
  ./he agent mongo                         list document collections
  ./he agent mongo <coll> --filter '{…}'   read-only find (--count for counts)
  ./he agent api /patients                 GET a REST endpoint
  ./he agent shell                         interactive REPL (\h for help)

Evaluation protocol:
  ./he task show <TASK-ID>                 read your assigned question
  ./he task submit <TASK-ID> -a "answer"   submit; exit code 0 = correct

Rules:
  1. Use only the tools above to inspect data. Writes are blocked anyway.
  2. Do NOT read evals/tasks.jsonl, evals/grader.py, or docs under evals/ —
     they contain the answer key. Work the problem from the databases.
  3. Full guide: docs/agent_interface.md   Schema spec: docs/schema.md
════════════════════════════════════════════════════════════════════════════
BRIEFING

# ── 5. optional: print the assigned task ────────────────────────────────────
if [[ $# -ge 1 && -n "${1:-}" ]]; then
    echo
    say "assigned task:"
    "$HE" task show "$1"
fi
