"""Live verification: every stored gold answer must be re-derivable from the DBs.

Skipped wholesale when PostgreSQL isn't reachable (e.g. CI without Docker).
Run the full environment first:  ./agent_entrypoint.sh
"""

from __future__ import annotations

import pytest
from sqlalchemy import text

from evals.runner import check_task, load_tasks


def _postgres_up() -> bool:
    try:
        from setup.db.postgres import get_engine

        with get_engine().connect() as conn:
            conn.execute(text("SELECT 1"))
        return True
    except Exception:  # noqa: BLE001
        return False


def _mongo_up() -> bool:
    try:
        from setup.db.mongo import get_sync_db

        get_sync_db().command("ping")
        return True
    except Exception:  # noqa: BLE001
        return False


POSTGRES_UP = _postgres_up()
MONGO_UP = _mongo_up()
TASKS = load_tasks()

pytestmark = pytest.mark.skipif(
    not POSTGRES_UP, reason="PostgreSQL not reachable — run ./agent_entrypoint.sh first"
)


def _needs_mongo(task: dict) -> bool:
    return bool(task.get("gold_check"))


@pytest.mark.parametrize("task", TASKS, ids=lambda t: t["id"])
def test_gold_answer_matches_live_data(task):
    if _needs_mongo(task) and not MONGO_UP:
        pytest.skip("MongoDB not reachable")
    derived, result = check_task(task)
    assert result.correct, (
        f"{task['id']}: stored gold {task['gold_answer']!r} but live data derives "
        f"{derived!r} ({result.note})"
    )


@pytest.mark.skipif(not POSTGRES_UP, reason="PostgreSQL not reachable")
def test_seed_row_totals():
    """The workbook load should have produced the documented grand total."""
    from setup.db.postgres import get_engine
    from setup.models import Base

    with get_engine().connect() as conn:
        total = 0
        for table in Base.metadata.sorted_tables:
            total += conn.execute(text(f"SELECT count(*) FROM {table.name}")).scalar_one()
    assert total == 1069, f"expected 1069 seeded rows, found {total}"
