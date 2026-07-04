"""Task loading, gold-answer re-derivation, and submission logging.

The dataset lives in ``evals/tasks.jsonl``. Each record carries either a
``gold_sql`` reference solution (PostgreSQL) or a ``gold_check`` spec for
the document store; :func:`derive_answer` re-executes those against the
live databases so the gold answers can be verified end-to-end at any time
(``hospital-env task check`` / ``tests/test_gold_answers.py``).
"""

from __future__ import annotations

import json
from collections import Counter
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from evals.grader import GradeResult, grade

REPO_ROOT = Path(__file__).resolve().parent.parent
TASKS_PATH = REPO_ROOT / "evals" / "tasks.jsonl"
RESULTS_DIR = REPO_ROOT / "evals" / "results"

CATEGORIES = {
    "demographics", "organization", "clinical", "medications", "labs", "imaging",
    "encounters", "scheduling", "billing", "coverage", "communications",
    "operations", "pharmacy", "documents", "cross-domain", "cross-store",
}
DIFFICULTIES = {"easy", "medium", "hard"}


def load_tasks(path: str | Path | None = None) -> list[dict[str, Any]]:
    tasks_file = Path(path) if path else TASKS_PATH
    tasks: list[dict[str, Any]] = []
    with tasks_file.open() as fh:
        for line_no, line in enumerate(fh, start=1):
            line = line.strip()
            if not line:
                continue
            try:
                tasks.append(json.loads(line))
            except json.JSONDecodeError as exc:
                raise ValueError(f"{tasks_file}:{line_no}: invalid JSON — {exc}") from exc
    return tasks


def get_task(task_id: str, path: str | Path | None = None) -> dict[str, Any]:
    for task in load_tasks(path):
        if task["id"].lower() == task_id.lower():
            return task
    raise KeyError(f"No task with id {task_id!r}. Try: hospital-env task list")


def _sql_rows_to_answer(columns: list[str], rows: list[list[Any]]) -> str:
    """Collapse a reference-query result into a submission-shaped string."""
    if not rows:
        raise ValueError("Reference query returned no rows.")
    if len(rows) == 1 and len(rows[0]) == 1:
        return str(rows[0][0])
    if all(len(row) == 1 for row in rows):
        return ", ".join(str(row[0]) for row in rows)
    if len(rows) == 1:
        return ", ".join(str(v) for v in rows[0])
    raise ValueError(f"Reference query shape {len(rows)}x{len(columns)} is not collapsible.")


def derive_answer(task: dict[str, Any]) -> str:
    """Re-derive the answer for a task from the live databases."""
    gold_sql = task.get("gold_sql")
    if gold_sql:
        from setup.agent.sql_tool import run_sql

        result = run_sql(gold_sql, limit=100)
        return _sql_rows_to_answer(result.columns, result.rows)

    check = task.get("gold_check")
    if not check:
        raise ValueError(f"Task {task['id']} has neither gold_sql nor gold_check.")

    kind = check.get("kind")
    if kind == "mongo_count":
        from setup.agent.mongo_tool import count

        return str(count(check["collection"], check.get("filter") or {}))
    if kind == "mongo_top":
        from setup.agent.mongo_tool import find
        from setup.agent.sql_tool import run_sql

        docs = find(
            check["collection"],
            check.get("filter") or {},
            projection={check["group_field"]: 1},
            limit=200,
        )
        tallies = Counter(doc[check["group_field"]] for doc in docs)
        ranked = tallies.most_common()
        if len(ranked) > 1 and ranked[0][1] == ranked[1][1]:
            raise ValueError(f"Task {task['id']}: tie at the top — gold is ambiguous.")
        top_key = ranked[0][0]
        map_sql = check["map_sql"].replace(":key", f"'{top_key}'")
        result = run_sql(map_sql, limit=5)
        return _sql_rows_to_answer(result.columns, result.rows)
    raise ValueError(f"Task {task['id']}: unknown gold_check kind {kind!r}.")


def check_task(task: dict[str, Any]) -> tuple[str, GradeResult]:
    """Verify one task's stored gold answer against the live databases."""
    derived = derive_answer(task)
    return derived, grade(task, derived)


def log_submission(task_id: str, answer: str, result: GradeResult) -> Path:
    RESULTS_DIR.mkdir(parents=True, exist_ok=True)
    out = RESULTS_DIR / "submissions.jsonl"
    with out.open("a") as fh:
        fh.write(
            json.dumps(
                {
                    "ts": datetime.now(UTC).isoformat(timespec="seconds"),
                    "task_id": task_id,
                    "answer": answer,
                    "correct": result.correct,
                }
            )
            + "\n"
        )
    return out


def public_view(task: dict[str, Any]) -> dict[str, Any]:
    """The agent-facing projection of a task: no gold answer, no reference solution."""
    return {
        "id": task["id"],
        "category": task["category"],
        "difficulty": task["difficulty"],
        "answer_type": task["answer_type"],
        "question": task["question"],
    }
