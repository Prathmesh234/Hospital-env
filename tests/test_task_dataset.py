"""Static validation of evals/tasks.jsonl — no database required."""

from __future__ import annotations

import pytest

from evals.grader import ANSWER_TYPES
from evals.runner import CATEGORIES, DIFFICULTIES, load_tasks, public_view

TASKS = load_tasks()
REQUIRED_FIELDS = {"id", "category", "difficulty", "question", "answer_type", "gold_answer"}


def test_dataset_is_nonempty_and_ids_unique():
    ids = [t["id"] for t in TASKS]
    assert len(ids) >= 30
    assert len(ids) == len(set(ids)), "duplicate task ids"


@pytest.mark.parametrize("task", TASKS, ids=lambda t: t["id"])
def test_required_fields(task):
    missing = REQUIRED_FIELDS - task.keys()
    assert not missing, f"{task['id']} missing {missing}"


@pytest.mark.parametrize("task", TASKS, ids=lambda t: t["id"])
def test_enums_valid(task):
    assert task["category"] in CATEGORIES
    assert task["difficulty"] in DIFFICULTIES
    assert task["answer_type"] in ANSWER_TYPES


@pytest.mark.parametrize("task", TASKS, ids=lambda t: t["id"])
def test_has_exactly_one_gold_source(task):
    has_sql = bool(task.get("gold_sql"))
    has_check = bool(task.get("gold_check"))
    assert has_sql != has_check, f"{task['id']} must have gold_sql XOR gold_check"


@pytest.mark.parametrize("task", TASKS, ids=lambda t: t["id"])
def test_gold_sql_is_readonly(task):
    sql = task.get("gold_sql")
    if not sql:
        pytest.skip("mongo-checked task")
    from setup.agent.sql_tool import validate_readonly_sql

    validate_readonly_sql(sql)  # raises SQLGuardError on violation


@pytest.mark.parametrize("task", TASKS, ids=lambda t: t["id"])
def test_collection_answer_types_use_lists(task):
    if task["answer_type"] in {"string_set", "ordered_list"}:
        assert isinstance(task["gold_answer"], list)
        assert len(task["gold_answer"]) >= 2


@pytest.mark.parametrize("task", TASKS, ids=lambda t: t["id"])
def test_public_view_leaks_nothing(task):
    view = public_view(task)
    assert set(view) == {"id", "category", "difficulty", "answer_type", "question"}
    assert "gold" not in str(view).lower()


def test_difficulty_and_category_spread():
    difficulties = {t["difficulty"] for t in TASKS}
    assert difficulties == DIFFICULTIES, "want easy+medium+hard represented"
    categories = {t["category"] for t in TASKS}
    assert len(categories) >= 10, "want broad domain coverage"
    assert "cross-store" in categories, "want at least one Mongo↔Postgres task"
