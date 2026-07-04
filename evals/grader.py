"""Answer normalization and grading.

Submissions arrive as free-form strings (an agent's final answer). Each task
declares an ``answer_type`` that picks the comparison strategy:

    string        case/whitespace-insensitive exact match
    number        numeric compare with tolerance (strips $ , % units)
    date          calendar-date compare across common formats
    boolean       yes/no/true/false
    string_set    unordered collection, split on , ; or newlines
    ordered_list  same, but order matters
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date, datetime
from typing import Any

ANSWER_TYPES = {"string", "number", "date", "boolean", "string_set", "ordered_list"}

DEFAULT_ABS_TOL = 0.01
DEFAULT_REL_TOL = 1e-6

_TRUE_WORDS = {"true", "yes", "y", "t", "1"}
_FALSE_WORDS = {"false", "no", "n", "f", "0"}

_DATE_FORMATS = [
    "%Y-%m-%d",
    "%m/%d/%Y",
    "%m/%d/%y",
    "%d %B %Y",
    "%d %b %Y",
    "%B %d, %Y",
    "%b %d, %Y",
    "%B %d %Y",
    "%b %d %Y",
]


class GradingError(ValueError):
    """A submission (or gold answer) could not be interpreted for its type."""


@dataclass
class GradeResult:
    task_id: str
    correct: bool
    expected: Any
    submitted: str
    note: str = ""


def normalize_string(raw: str) -> str:
    s = str(raw).strip().strip("\"'`").strip()
    s = re.sub(r"\s+", " ", s)
    return s.casefold()


def parse_number(raw: str | int | float) -> float:
    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        return float(raw)
    s = str(raw).strip()
    s = re.sub(r"[\$,%]", "", s)
    s = s.replace(",", "").strip()
    # tolerate trailing units, e.g. "42 mg" or "17 days"
    match = re.match(r"^[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", s)
    if not match:
        raise GradingError(f"Could not parse a number from {raw!r}.")
    return float(match.group(0))


def parse_date(raw: str | date | datetime) -> date:
    if isinstance(raw, datetime):
        return raw.date()
    if isinstance(raw, date):
        return raw
    s = str(raw).strip().strip("\"'`")
    try:
        return datetime.fromisoformat(s).date()
    except ValueError:
        pass
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    raise GradingError(f"Could not parse a date from {raw!r}.")


def parse_boolean(raw: str | bool) -> bool:
    if isinstance(raw, bool):
        return raw
    s = normalize_string(str(raw))
    if s in _TRUE_WORDS:
        return True
    if s in _FALSE_WORDS:
        return False
    raise GradingError(f"Could not parse a boolean from {raw!r}.")


def split_items(raw: str | list[Any]) -> list[str]:
    if isinstance(raw, list):
        return [normalize_string(str(item)) for item in raw]
    s = str(raw).strip()
    # strip list-ish wrappers the model might emit
    s = s.strip("[](){}")
    parts = re.split(r"[,;\n]+", s)
    items = [normalize_string(p) for p in parts if normalize_string(p)]
    return items


def _grade_value(
    answer_type: str,
    gold: Any,
    submitted: str,
    abs_tol: float,
    rel_tol: float,
) -> tuple[bool, str]:
    if answer_type == "string":
        return normalize_string(str(gold)) == normalize_string(submitted), ""
    if answer_type == "number":
        got = parse_number(submitted)
        want = parse_number(gold)
        ok = math.isclose(got, want, abs_tol=abs_tol, rel_tol=rel_tol)
        return ok, "" if ok else f"expected {want}, got {got}"
    if answer_type == "date":
        return parse_date(submitted) == parse_date(gold), ""
    if answer_type == "boolean":
        return parse_boolean(submitted) == parse_boolean(gold), ""
    if answer_type == "string_set":
        got_set = set(split_items(submitted))
        want_set = set(split_items(gold))
        if got_set == want_set:
            return True, ""
        missing = sorted(want_set - got_set)
        extra = sorted(got_set - want_set)
        return False, f"missing={missing} extra={extra}"
    if answer_type == "ordered_list":
        got_list = split_items(submitted)
        want_list = split_items(gold)
        return got_list == want_list, "" if got_list == want_list else f"expected order {want_list}"
    raise GradingError(f"Unknown answer_type {answer_type!r}.")


def grade(task: dict[str, Any], submitted: str) -> GradeResult:
    """Grade one free-form submission against a task record."""
    answer_type = task.get("answer_type", "string")
    if answer_type not in ANSWER_TYPES:
        raise GradingError(f"Task {task.get('id')}: bad answer_type {answer_type!r}.")
    gold = task["gold_answer"]
    abs_tol = float(task.get("abs_tol", DEFAULT_ABS_TOL))
    rel_tol = float(task.get("rel_tol", DEFAULT_REL_TOL))
    try:
        correct, note = _grade_value(answer_type, gold, submitted, abs_tol, rel_tol)
    except GradingError as exc:
        return GradeResult(
            task_id=task.get("id", "?"),
            correct=False,
            expected=gold,
            submitted=submitted,
            note=str(exc),
        )
    return GradeResult(
        task_id=task.get("id", "?"),
        correct=correct,
        expected=gold,
        submitted=submitted,
        note=note,
    )
