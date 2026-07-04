"""Unit tests for the answer grader — no database required."""

from __future__ import annotations

import pytest

from evals.grader import (
    GradingError,
    grade,
    normalize_string,
    parse_boolean,
    parse_date,
    parse_number,
    split_items,
)


def _task(answer_type: str, gold, **extra) -> dict:
    return {"id": "T-TEST", "answer_type": answer_type, "gold_answer": gold, **extra}


class TestNormalization:
    def test_string_case_and_whitespace(self):
        assert normalize_string("  Medi-Cal   Managed Care ") == "medi-cal managed care"

    def test_string_strips_quotes(self):
        assert normalize_string("'Blackwell'") == "blackwell"

    def test_number_currency_and_commas(self):
        assert parse_number("$38,420.00") == pytest.approx(38420.0)

    def test_number_percent_and_units(self):
        assert parse_number("59.7%") == pytest.approx(59.7)
        assert parse_number("162 mmHg") == pytest.approx(162.0)

    def test_number_rejects_garbage(self):
        with pytest.raises(GradingError):
            parse_number("not a number")

    def test_date_formats(self):
        for raw in ["2025-02-04", "02/04/2025", "Feb 4, 2025", "4 February 2025"]:
            assert parse_date(raw).isoformat() == "2025-02-04"

    def test_boolean_words(self):
        assert parse_boolean("Yes") is True
        assert parse_boolean("FALSE") is False
        with pytest.raises(GradingError):
            parse_boolean("maybe")

    def test_split_items_separators(self):
        assert split_items("CT, US; XR") == ["ct", "us", "xr"]
        assert split_items("[Chen, Kim]") == ["chen", "kim"]


class TestGrading:
    def test_number_within_tolerance(self):
        assert grade(_task("number", 59.7, abs_tol=0.05), "59.68").correct

    def test_number_outside_tolerance(self):
        assert not grade(_task("number", 59.7), "42").correct

    def test_string_case_insensitive(self):
        assert grade(_task("string", "Blackwell"), "  blackwell ").correct

    def test_string_wrong(self):
        assert not grade(_task("string", "Blackwell"), "Chen").correct

    def test_date_flexible_format(self):
        assert grade(_task("date", "2025-02-04"), "February 4, 2025").correct

    def test_boolean(self):
        assert grade(_task("boolean", False), "no").correct
        assert not grade(_task("boolean", False), "yes").correct

    def test_string_set_order_free(self):
        task = _task("string_set", ["CT", "US", "XR"])
        assert grade(task, "xr, ct, us").correct

    def test_string_set_missing_item_reported(self):
        result = grade(_task("string_set", ["CT", "US", "XR"]), "CT, US")
        assert not result.correct
        assert "missing" in result.note

    def test_ordered_list_order_matters(self):
        task = _task("ordered_list", ["Chen", "Kim", "O'Connor"])
        assert grade(task, "chen; kim; o'connor").correct
        assert not grade(task, "Kim, Chen, O'Connor").correct

    def test_unparseable_submission_is_incorrect_not_crash(self):
        result = grade(_task("number", 10), "ten-ish")
        assert not result.correct
        assert "Could not parse" in result.note

    def test_unknown_answer_type_raises(self):
        with pytest.raises(GradingError):
            grade(_task("essay", "x"), "x")
