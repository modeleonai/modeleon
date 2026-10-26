# SPDX-License-Identifier: Apache-2.0
"""``IFERROR``, ``ISERROR``, ``ISNUMBER``, ``XOR``, ``IFS`` and ``SWITCH``
compute what the spreadsheet computes, with the formula as the product.

``IFERROR`` catches an error marker before it cascades; ``ISERROR`` and
``ISNUMBER`` name a cell's kind and are never an error themselves;
``XOR`` counts the conditions that hold; ``IFS`` picks the value of the
first condition that holds and ``SWITCH`` the value of the first case
equal to its expression — both lazily, a branch not taken never read.
Each is one call per period through the shared helper: scalars
broadcast, a one-value series broadcasts, other lengths are refused,
tracked operands lift per coordinate, the AST is a call over the
operands. An operand reads as its cell does: an empty text typed into a
line is the blank cell the workbook writes; a formula's is text.
"""

from __future__ import annotations

import datetime as dt

import pytest

import modeleon as mo
from modeleon.core.expr import FuncCall, Literal, VarRef
from modeleon.functions.logical import IFERROR, IFS, ISERROR, ISNUMBER, SWITCH, XOR
from modeleon.functions.registry import (
    admits_in_loop,
    excel_name,
    file_prefix,
    is_elementwise,
    is_passthrough,
    is_range_taking,
    spec,
)

FUNCTIONS = {
    "IFERROR": IFERROR, "ISERROR": ISERROR, "ISNUMBER": ISNUMBER,
    "XOR": XOR, "IFS": IFS, "SWITCH": SWITCH,
}
MARKERS = ["#DIV/0!", "#VALUE!", "#N/A", "#NUM!", "#NAME?", "#REF!", "#NULL!"]


def _model():
    m = mo.Model("m", default_grain="year", default_start="2026", default_periods=3)
    with m:
        m.x = mo.MultiVariable("X")
        m.x.a = mo.Variable([3.0, 8.0, 5.0], display_name="a")
        m.x.z = mo.Variable([1.0, 0.0, 2.0], display_name="z")
        m.x.q = (m.x.a / m.x.z).set_display_name("q")
    return m


def _empty_text_formula(m):
    """A formula line whose every value is the empty text."""
    out = IFS(m.x.a > 0.0, "", True, "x")
    assert out.value == ["", "", ""]
    return out


def _tracked():
    m = mo.Model("t", tracks=mo.Tracks("plan", "actual"),
                 default_grain="month", default_start="2026-01", default_periods=2)
    with m:
        m.x = mo.Variable(plan=[3.0, "#DIV/0!"], actual=[5.0, 8.0])
    return m


# ══════════════════════════════════════════════════════════════════════
# IFERROR
# ══════════════════════════════════════════════════════════════════════


class TestIfError:
    def test_the_fallback_stands_in_for_an_error(self):
        m = _model()
        assert m.x.q.value == [3.0, "#DIV/0!", 2.5]
        assert IFERROR(m.x.q, 0.0).value == [3.0, 0.0, 2.5]
        assert IFERROR(m.x.q, "n/a").value == [3.0, "n/a", 2.5]

    def test_every_marker_is_an_error(self):
        out = IFERROR(mo.Variable([*MARKERS, 1.0]), -1.0)
        assert out.value == [-1.0] * len(MARKERS) + [1.0]

    def test_the_fallback_may_vary_by_period_or_be_an_error_itself(self):
        m = _model()
        assert IFERROR(m.x.q, m.x.a).value == [3.0, 8.0, 2.5]
        assert IFERROR(m.x.q, [9.0, 8.0, 7.0]).value == [3.0, 8.0, 2.5]
        assert IFERROR(m.x.q, mo.Variable(["#N/A"] * 3)).value == [3.0, "#N/A", 2.5]

    def test_a_blank_reads_as_zero(self):
        # A formula that lands on a blank cell shows 0, whichever operand
        # the blank is.
        assert IFERROR(mo.Variable([None, 5.0]), "x").value == [0, 5.0]
        assert IFERROR(mo.Variable(["#N/A", 2.0]), mo.Variable([None, None])).value == [0, 2.0]

    def test_an_empty_text_typed_into_a_line_is_a_blank_a_formula_s_is_text(self):
        m = _model()
        # The workbook writes a typed empty text as nothing: a blank cell.
        assert IFERROR(mo.Variable(["", 5.0]), "x").value == [0, 5.0]
        assert IFERROR(_empty_text_formula(m), "x").value == ["", "", ""]
        assert IFERROR("", 1.0).value == ""

    def test_scalars_only_is_one_value(self):
        assert IFERROR(5.0, 0).value == 5.0 and IFERROR(5.0, 0).var_type == "scalar"
        assert IFERROR(mo.Variable(1.0) / 0, -1.0).value == -1.0

    def test_both_operands_are_precedents_and_the_ast_is_a_call(self):
        m = _model()
        with m:
            m.x.safe = IFERROR(m.x.q, m.x.a)
        assert m.x.safe.dependencies == {m.x.q.id, m.x.a.id}
        call = m.x.safe._expr
        assert isinstance(call, FuncCall) and call.func == "IFERROR"
        assert [type(arg) for arg in call.args] == [VarRef, VarRef]
        lit = IFERROR(m.x.q, 0.0)._expr
        assert isinstance(lit.args[1], Literal) and lit.args[1].value == 0.0

    def test_the_formula_is_the_source(self):
        m = _model()
        out = IFERROR(m.x.q, 0.0)
        assert out.formula == "IFERROR(q, 0.0)"
        assert out._source_code == "IFERROR(q, 0.0)"
        assert IFERROR(m.x.a / m.x.z, "n/a").formula == "IFERROR((a / z), 'n/a')"

    def test_typed_by_the_value_where_it_is_not_an_error(self):
        cost = mo.Variable([3.0, "#DIV/0!"], unit="$")
        assert IFERROR(cost, 0.0).value_type == "float"
        # A text fallback does not make a row of numbers a row of text …
        assert IFERROR(cost, "n/a").value_type == "float"
        assert IFERROR(mo.Variable(["#DIV/0!", 3.0]), "n/a").value_type == "float"
        # … and a row read in with markers among its numbers, typed as
        # text, is cleaned into numbers.
        read_in = mo.Variable(["#N/A", 1.0, "#VALUE!"])
        assert read_in.value_type == "string"
        assert IFERROR(read_in, 0.0).value_type == "float"
        assert IFERROR(mo.Variable(["x", "#N/A"]), "").value_type == "string"
        assert IFERROR(mo.Variable([1, "#N/A"]), 0).value_type == "int"
        assert IFERROR(mo.Variable([dt.date(2026, 1, 1), "#N/A"]), dt.date(2026, 1, 2)) \
            .value_type == "datetime"
        # A value that is an error throughout is typed by the fallback.
        assert IFERROR(mo.Variable(["#N/A", "#N/A"]), "none").value_type == "string"
        assert IFERROR(mo.Variable(["#N/A", "#N/A"]), 0.5).value_type == "float"

    def test_the_unit_the_operands_share(self):
        cost = mo.Variable([3.0, "#DIV/0!"], unit="$")
        assert str(IFERROR(cost, 0.0)._unit) == "$"
        assert str(IFERROR(cost, mo.Variable([1.0, 1.0], unit="$"))._unit) == "$"
        with pytest.raises(ValueError, match="mo.IFERROR takes its values in one unit"):
            IFERROR(cost, mo.Variable([1.0, 1.0], unit="hr"))

    def test_a_tracked_operand_lifts_per_track(self):
        m = _tracked()
        with m:
            m.h = IFERROR(m.x, 0.0)
        assert dict(m.h._value.items()) == {"plan": [3.0, 0.0], "actual": [5.0, 8.0]}
        assert m.h.var_type == "list"


# ══════════════════════════════════════════════════════════════════════
# ISERROR and ISNUMBER
# ══════════════════════════════════════════════════════════════════════


class TestIsError:
    def test_names_an_error_and_is_never_one(self):
        m = _model()
        assert ISERROR(m.x.q).value == [False, True, False]
        assert ISERROR(mo.Variable(MARKERS)).value == [True] * len(MARKERS)
        assert ISERROR(mo.Variable([1.0, "text", None, True, dt.date(2026, 1, 1)])).value == \
            [False, False, False, False, False]

    def test_any_text_starting_with_a_hash_is_a_marker_to_the_engine(self):
        # The convention arithmetic carries: ``'#…'`` is an error cell.
        assert ISERROR(mo.Variable(["#x", "x#"])).value == [True, False]

    def test_scalars_only_is_one_value(self):
        assert ISERROR(1.0).value is False
        assert ISERROR(mo.Variable(1.0) / 0).value is True
        assert ISERROR(2.0).var_type == "scalar"

    def test_a_boolean_without_a_unit(self):
        m = _model()
        out = ISERROR(mo.Variable([3.0, "#DIV/0!"], unit="$"))
        assert out.value_type == "bool" and out._unit is None
        assert ISERROR(m.x.q).formula == "ISERROR(q)" == ISERROR(m.x.q)._source_code

    def test_a_tracked_operand_lifts_per_track(self):
        m = _tracked()
        with m:
            m.e = ISERROR(m.x)
        assert dict(m.e._value.items()) == {"plan": [False, True], "actual": [False, False]}


class TestIsNumber:
    def test_numbers_dates_and_durations_are_numbers(self):
        row = mo.Variable([1, 2.5, dt.date(2026, 1, 1), dt.datetime(2026, 1, 1, 12),
                           dt.timedelta(days=3)])
        assert ISNUMBER(row).value == [True] * 5

    def test_booleans_text_blanks_and_errors_are_not(self):
        # A number written as text is text; a boolean is not a number.
        row = mo.Variable([True, False, "5", "text", "", None, "#DIV/0!", "#N/A"])
        assert ISNUMBER(row).value == [False] * 8

    def test_scalars_only_is_one_value(self):
        assert ISNUMBER(7).value is True and ISNUMBER("7").value is False
        assert ISNUMBER(mo.Variable(1.0) / 0).value is False
        assert ISNUMBER(7).var_type == "scalar"

    def test_a_boolean_without_a_unit(self):
        m = _model()
        out = ISNUMBER(mo.Variable([3.0, 8.0], unit="$"))
        assert out.value_type == "bool" and out._unit is None
        assert ISNUMBER(m.x.a).formula == "ISNUMBER(a)" == ISNUMBER(m.x.a)._source_code


class TestAValueGivenToTheCall:
    """A Python value reaching one of these functions is read as the cell
    it stands for — a loop runs the function again over its rows' plain
    values, where ``None`` is a blank cell and ``'#DIV/0!'`` the error."""

    def test_none_is_a_blank_cell(self):
        assert ISNUMBER(None).value is False
        assert XOR(None, True).value is True
        assert SWITCH(None, "", "empty", "other").value == "empty"
        assert IFS(None, 1.0, True, 2.0).value == 2.0

    def test_a_marker_is_the_error(self):
        assert ISERROR("#N/A").value is True
        assert IFERROR("#DIV/0!", 0).value == 0
        assert IFS(True, "#N/A").value == "#N/A"


# ══════════════════════════════════════════════════════════════════════
# XOR
# ══════════════════════════════════════════════════════════════════════


class TestXor:
    @pytest.mark.parametrize("conditions, expected", [
        ((True,), True),
        ((False,), False),
        ((True, True), False),
        ((True, False), True),
        ((True, True, True), True),
        ((True, True, True, True), False),
        ((1, 0), True),
        ((2, -3), False),
        ((0, 0), False),
    ])
    def test_true_when_an_odd_number_of_conditions_hold(self, conditions, expected):
        out = XOR(*conditions)
        assert out.value is expected and out.var_type == "scalar"

    def test_element_wise_over_series(self):
        m = _model()
        assert XOR(m.x.a > 4.0, True).value == [True, False, False]
        assert XOR(m.x.a > 4.0, m.x.z > 0.0, True).value == [False, False, True]
        assert XOR(mo.Variable([0, 2, -1])).value == [False, True, True]

    def test_a_blank_is_skipped_and_nothing_left_is_value(self):
        # As the spreadsheet skips a blank cell; an empty text typed into
        # a line is written as one.
        out = XOR(mo.Variable([None, 2.0, None]), mo.Variable([True, False, 0]))
        assert out.value == [True, True, False]
        assert XOR(mo.Variable(["", True]), True).value == [True, False]
        assert XOR(mo.Variable([None, None])).value == ["#VALUE!"] * 2

    def test_text_is_refused(self):
        # A text cell the spreadsheet skips; text typed into the formula
        # is #VALUE! — no value matches both, so none is given.
        m = _model()
        with m:
            m.x.answer = mo.Variable(["yes", "no", "yes"])
        with pytest.raises(TypeError, match="XOR counts conditions"):
            XOR(m.x.answer, True)
        with pytest.raises(TypeError, match="XOR counts conditions"):
            XOR(mo.Variable(["abc", True]), True)
        with pytest.raises(TypeError, match="'TRUE'"):
            XOR(True, "TRUE")
        assert XOR(m.x.answer == "yes", True).value == [False, True, False]

    def test_an_empty_text_is_skipped_as_a_blank(self):
        # Typed into a line it is the empty cell the workbook writes; a
        # formula's, in the cell XOR reads, is text the spreadsheet skips.
        m = _model()
        with m:
            m.x.label = _empty_text_formula(m)
        assert XOR(m.x.label, True).value == [True] * 3
        assert XOR(m.x.label).value == ["#VALUE!"] * 3

    def test_the_first_error_in_reading_order_is_the_result(self):
        m = _model()
        assert (m.x.q > 1.0).value == [True, "#DIV/0!", True]
        out = XOR(m.x.q > 1.0, mo.Variable([True, "#N/A", True]))
        assert out.value == [False, "#DIV/0!", False]
        assert XOR(mo.Variable(["#N/A", "#DIV/0!", "#NUM!"]), m.x.q > 1.0).value == \
            ["#N/A", "#DIV/0!", "#NUM!"]

    def test_no_condition_is_refused(self):
        with pytest.raises(TypeError, match="at least one condition"):
            XOR()

    def test_a_boolean_without_a_unit_and_the_formula(self):
        m = _model()
        out = XOR(m.x.a > 4.0, True)
        assert out.value_type == "bool" and out._unit is None
        assert out.formula == "XOR((a > 4.0), True)" == out._source_code
        assert XOR(mo.Variable([1.0, 0.0], unit="$"))._unit is None

    def test_a_tracked_operand_lifts_per_track(self):
        m = _tracked()
        with m:
            m.o = XOR(m.x > 4.0, True)
        assert dict(m.o._value.items()) == {"plan": [True, "#DIV/0!"], "actual": [False, False]}


# ══════════════════════════════════════════════════════════════════════
# IFS
# ══════════════════════════════════════════════════════════════════════


class TestIfs:
    def test_the_value_of_the_first_condition_that_holds(self):
        m = _model()
        out = IFS(m.x.a > 7.0, "big", m.x.a > 4.0, "mid", True, "small")
        assert out.value == ["small", "big", "mid"] and out.var_type == "list"
        assert IFS(m.x.a > 4.0, m.x.a, True, 0.0).value == [0.0, 8.0, 5.0]

    def test_no_condition_holding_is_na(self):
        m = _model()
        assert IFS(m.x.a > 100.0, 1.0).value == ["#N/A"] * 3
        assert IFS(False, 1.0).value == "#N/A"
        assert IFERROR(IFS(m.x.a > 100.0, 1.0), -1.0).value == [-1.0] * 3

    def test_a_branch_not_taken_is_never_read(self):
        m = _model()
        # The error in q never reaches the result where z is 0 …
        assert IFS(m.x.z == 0, 0.0, True, m.x.q).value == [3.0, 0.0, 2.5]
        # … and a condition after the first that holds is not read either.
        assert IFS(True, 1.0, m.x.q > 1.0, 2.0).value == [1.0, 1.0, 1.0]

    def test_a_condition_that_is_an_error_is_the_result(self):
        m = _model()
        assert IFS(m.x.q > 1.0, 1.0, True, 2.0).value == [1.0, "#DIV/0!", 1.0]
        assert IFS(mo.Variable(["#N/A", False]), 1.0, True, 2.0).value == ["#N/A", 2.0]

    def test_how_a_condition_is_read(self):
        # A number is TRUE when nonzero; a blank is FALSE; booleans are
        # themselves.
        assert IFS(mo.Variable([0, 2, -1]), "yes", True, "no").value == ["no", "yes", "yes"]
        assert IFS(mo.Variable([None, "", 1]), "a", True, "b").value == ["b", "b", "a"]
        assert IFS(mo.Variable([True, False]), 1.0, True, 2.0).value == [1.0, 2.0]

    def test_the_text_true_or_false_is_that_boolean_other_text_is_value(self):
        words = mo.Variable(["TRUE", "FALSE", "true", "False"])
        assert IFS(words, 1.0, True, 2.0).value == [1.0, 2.0, 1.0, 2.0]
        assert IFS("TRUE", 1.0, True, 2.0).value == 1.0
        # A number written as text is text, as Excel reads it.
        assert IFS(mo.Variable(["abc", "1"]), 1.0, True, 2.0).value == ["#VALUE!"] * 2

    def test_an_empty_text_is_text_unless_typed_into_a_line(self):
        m = _model()
        assert IFS(_empty_text_formula(m), 1.0, True, 2.0).value == ["#VALUE!"] * 3
        assert IFS("", 1.0, True, 2.0).value == "#VALUE!"
        assert IFS(mo.Variable(["", True]), 1.0, True, 2.0).value == [2.0, 1.0]

    def test_a_picked_blank_shows_zero_and_does_not_type_the_result(self):
        out = IFS(True, mo.Variable([None, 2.5]))
        assert out.value == [0, 2.5] and out.value_type == "float"
        assert IFS(True, mo.Variable(["", 2.5])).value == [0, 2.5]
        out = IFS(mo.Variable([True, False]), mo.Variable([None, 9]), True, -1.0)
        assert out.value == [0, -1.0] and out.value_type == "float"
        assert IFS(False, 1.0, True, None).value == 0

    def test_a_scalar_condition_over_a_series_value_gives_a_series(self):
        m = _model()
        out = IFS(True, m.x.a, False, 0.0)
        assert out.value == [3.0, 8.0, 5.0] and out.var_type == "list"
        assert IFS(m.x.a > 4.0, 1.0, True, 0.0).var_type == "list"

    def test_scalars_only_is_one_value(self):
        out = IFS(7.0 > 5.0, "hi", True, "lo")
        assert out.value == "hi" and out.var_type == "scalar"
        assert IFS(mo.Variable(7.0) > 5.0, "hi", True, "lo").value == "hi"

    def test_pairs_only(self):
        m = _model()
        with pytest.raises(TypeError, match="condition and value pairs"):
            IFS(m.x.a > 1.0)
        with pytest.raises(TypeError, match="got 3 operands"):
            IFS(m.x.a > 1.0, 1.0, True)
        with pytest.raises(TypeError, match="got 0 operands"):
            IFS()

    def test_typed_by_the_branch_taken(self):
        m = _model()
        assert IFS(m.x.a > 4.0, "big", True, "small").value_type == "string"
        assert IFS(m.x.a > 4.0, 1, True, 0).value_type == "int"
        assert IFS(m.x.a > 4.0, True, True, False).value_type == "bool"
        assert IFS(m.x.a > 4.0, dt.date(2026, 1, 1), True, dt.date(2027, 1, 1)) \
            .value_type == "datetime"
        # Typed by the first value that is not an error marker.
        assert IFS(m.x.q > 1.0, "x", True, "y").value_type == "string"

    def test_the_unit_the_values_share(self):
        m = _model()
        cost = mo.Variable([3.0, 8.0, 5.0], unit="$")
        out = IFS(m.x.a > 4.0, cost, True, 0.0)
        assert out.value_type == "float" and str(out._unit) == "$"
        # A condition is not a value the result takes: its unit is not
        # the result's.
        hours = mo.Variable([1.0, 0.0, 2.0], unit="hr")
        out = IFS(hours, cost, True, 0.0)
        assert out.value == [3.0, 0.0, 5.0] and str(out._unit) == "$"
        assert IFS(hours, 1.0, True, 0.0)._unit is None
        assert IFS(cost > 4.0, "x", True, "y")._unit is None

    def test_the_formula_is_the_source_and_the_ast_is_a_call(self):
        m = _model()
        out = IFS(m.x.a > 7.0, "big", m.x.a > 4.0, "mid", True, "small")
        assert out.formula == "IFS((a > 7.0), 'big', (a > 4.0), 'mid', True, 'small')"
        assert out._source_code == out.formula
        call = out._expr
        assert isinstance(call, FuncCall) and call.func == "IFS" and len(call.args) == 6
        assert isinstance(call.args[0], VarRef) and isinstance(call.args[1], Literal)
        with m:
            m.x.hot = m.x.a > 7.0
            m.x.band = IFS(m.x.hot, "big", True, m.x.z)
        assert m.x.band.dependencies == {m.x.hot.id, m.x.z.id}

    def test_a_tracked_operand_lifts_per_track(self):
        m = _tracked()
        with m:
            m.s = IFS(m.x > 4.0, "hi", True, "lo")
        assert dict(m.s._value.items()) == {"plan": ["lo", "#DIV/0!"], "actual": ["hi", "hi"]}
        assert m.s.var_type == "list" and m.s.value_type == "string"


# ══════════════════════════════════════════════════════════════════════
# SWITCH
# ══════════════════════════════════════════════════════════════════════


class TestSwitch:
    def test_the_value_of_the_first_case_equal_to_the_expression(self):
        m = _model()
        out = SWITCH(m.x.a, 3, "three", 8, "eight", "other")
        assert out.value == ["three", "eight", "other"] and out.var_type == "list"
        assert SWITCH(2, 1, "a", 2, "b", "c").value == "b"
        assert SWITCH(2, 1, "a", 2, "b", "c").var_type == "scalar"

    def test_no_case_matching_is_the_default_or_na(self):
        m = _model()
        assert SWITCH(m.x.a, 3, "three").value == ["three", "#N/A", "#N/A"]
        assert SWITCH(9, 1, "one").value == "#N/A"
        assert SWITCH(9, 1, "one", "none").value == "none"

    def test_numbers_are_equal_as_the_spreadsheet_compares_them(self):
        # To within 2**-48 of either — about 15 significant digits — so a
        # computed 0.1 * 3 is the 0.3 typed into a case.
        f = mo.Variable([0.1, 0.2, 0.3])
        assert (f * 3).value != [0.3, 0.6, 0.9]
        assert SWITCH(f * 3, 0.3, "p3", 0.6, "p6", 0.9, "p9", "other").value == \
            ["p3", "p6", "p9"]
        assert SWITCH(f + 0.2, 0.3, "a", 0.4, "b", 0.5, "c", "other").value == ["a", "b", "c"]
        assert SWITCH(1 + 2**-50, 1, "eq", "ne").value == "eq"
        assert SWITCH(1 + 2**-48, 1, "eq", "ne").value == "ne"
        assert SWITCH(123456789012345, 123456789012345.02, "eq", "ne").value == "eq"
        assert SWITCH(-0.1 * 3, -0.3, "eq", "ne").value == "eq"

    def test_zero_equals_only_zero_and_whole_numbers_only_themselves(self):
        assert SWITCH(1e-300, 0, "eq", "ne").value == "ne"
        assert SWITCH(2**52, 2**52 + 1, "eq", "ne").value == "ne"
        assert SWITCH(0.0, -0.0, "eq", "ne").value == "eq"

    def test_text_matches_regardless_of_case(self):
        names = mo.Variable(["Base", "bull", "x"])
        assert SWITCH(names, "base", 0.05, "BULL", 0.08, 0.03).value == [0.05, 0.08, 0.03]

    def test_never_across_kinds(self):
        m = _model()
        # 3 is not "3"; TRUE is not 1; a text is not a number.
        assert SWITCH(m.x.a, "3", "txt", "num").value == ["num"] * 3
        assert SWITCH(mo.Variable(["3", "8"]), 3, "num", "txt").value == ["txt"] * 2
        assert SWITCH(mo.Variable([True, 1]), 1, "one", True, "true").value == ["true", "one"]
        assert SWITCH(mo.Variable([False, 0]), 0, "zero", False, "false").value == \
            ["false", "zero"]

    def test_a_date_by_its_serial_number(self):
        d = mo.Variable(dt.date(2026, 1, 1))
        assert SWITCH(d, 46023, "yes", "no").value == "yes"
        assert SWITCH(d, dt.date(2026, 1, 1), "same", "no").value == "same"
        assert SWITCH(46023, dt.date(2026, 1, 1), "same", "no").value == "same"
        assert SWITCH(d, dt.date(2026, 1, 2), "next", "no").value == "no"

    def test_a_blank_is_zero_empty_text_and_false(self):
        # As the engine's ``==`` reads a blank cell against each kind; an
        # empty text typed into a line is a blank cell.
        blank = mo.Variable([None, 0, ""])
        assert SWITCH(blank, 0, "zero", "other").value == ["zero", "zero", "zero"]
        assert SWITCH(blank, "", "empty", "other").value == ["empty", "other", "empty"]
        assert SWITCH(blank, False, "false", "other").value == ["false", "other", "false"]
        assert SWITCH(mo.Variable([0, 1]), None, "blank", "other").value == ["blank", "other"]

    def test_a_formula_s_empty_text_is_text(self):
        m = _model()
        label = _empty_text_formula(m)
        assert SWITCH(label, 0, "zero", "other").value == ["other"] * 3
        assert SWITCH(label, "", "empty", "other").value == ["empty"] * 3

    def test_an_expression_or_a_reached_case_that_is_an_error_is_the_result(self):
        m = _model()
        assert SWITCH(m.x.q, 3, "three", "other").value == ["three", "#DIV/0!", "other"]
        # The case after the match is never read: period one matches 3
        # before q is compared; period two reaches q's error.
        assert SWITCH(m.x.a, 3, "three", m.x.q, "q", "other").value == \
            ["three", "#DIV/0!", "other"]
        assert SWITCH(5, 5, "#N/A", "other").value == "#N/A"

    def test_a_picked_blank_shows_zero_and_does_not_type_the_result(self):
        m = _model()
        blank = mo.Variable([None, 5.0, None])
        out = SWITCH(m.x.a, 3, blank, 8, blank, "other")
        assert out.value == [0, 5.0, "other"] and out.value_type == "float"
        assert SWITCH(9, 1, "one", None).value == 0

    def test_a_scalar_expression_over_series_values_gives_a_series(self):
        m = _model()
        out = SWITCH(2, 1, m.x.a, 2, m.x.z)
        assert out.value == [1.0, 0.0, 2.0] and out.var_type == "list"
        assert SWITCH(mo.Variable(1), 1, m.x.a, 2, m.x.z).value == [3.0, 8.0, 5.0]

    def test_an_expression_and_at_least_one_pair(self):
        m = _model()
        with pytest.raises(TypeError, match="case and value pairs"):
            SWITCH(m.x.a)
        with pytest.raises(TypeError, match="got 1 operand after the expression"):
            SWITCH(m.x.a, 1)

    def test_typed_by_the_branch_taken(self):
        pick = mo.Variable([1, 2, 3])
        assert SWITCH(pick, 1, "one", 2, "two", "many").value_type == "string"
        assert SWITCH(pick, 1, 10, 2, 20, 0).value_type == "int"
        assert SWITCH(pick, 1, dt.date(2026, 1, 1), dt.date(2027, 1, 1)).value_type == \
            "datetime"

    def test_the_unit_the_values_share(self):
        pick = mo.Variable([1, 2, 3])
        base = mo.Variable([3.0, 8.0, 5.0], unit="$")
        bull = mo.Variable([4.0, 9.0, 6.0], unit="$")
        out = SWITCH(pick, 1, base, 2, bull, 0.0)
        assert out.value == [3.0, 9.0, 0.0]
        assert out.value_type == "float" and str(out._unit) == "$"
        assert SWITCH(pick, 1, "one", 2, "two", "many")._unit is None
        assert SWITCH(base, 3.0, "x", "y")._unit is None
        hours = mo.Variable([1.0, 1.0, 1.0], unit="hr")
        with pytest.raises(ValueError, match="mo.SWITCH takes its values in one unit"):
            SWITCH(pick, 1, base, 2, hours, 0.0)
        # The expression and the cases are compared, never taken: their
        # unit is not the result's.
        shift = mo.Variable([8.0, 1.0, 8.0], unit="hr")
        out = SWITCH(shift, 8.0, base, 0.0)
        assert out.value == [3.0, 0.0, 5.0] and str(out._unit) == "$"
        assert SWITCH(shift, 8.0, 1.0, 0.0)._unit is None

    def test_a_case_in_another_unit_than_the_expression_is_refused(self):
        # Compared as == compares them: hours against a threshold in USD
        # is refused, a plain number or a case in hours goes.
        shift = mo.Variable([8.0, 1.0, 8.0], unit="hr", display_name="Shift")
        full = mo.Variable(8.0, unit="hr", display_name="Full shift")
        limit = mo.Variable(8.0, unit="USD", display_name="Limit")
        assert SWITCH(shift, full, 1.0, 0.0).value == [1.0, 0.0, 1.0]
        assert SWITCH(shift, 8.0, 1.0, 0.0).value == [1.0, 0.0, 1.0]
        with pytest.raises(ValueError, match="mo.SWITCH compares 'Shift' with the case "
                                             "'Limit'; they are in hr and USD"):
            SWITCH(shift, limit, 1.0, 0.0)
        with pytest.raises(ValueError, match="compares 'Shift' with the case 'Limit'"):
            SWITCH(shift, full, 1.0, limit, 2.0, 0.0)   # the second case too

    def test_the_formula_is_the_source_and_the_ast_is_a_call(self):
        m = _model()
        out = SWITCH(m.x.a, 3, "three", 8, "eight", "other")
        assert out.formula == "SWITCH(a, 3, 'three', 8, 'eight', 'other')"
        assert out._source_code == out.formula
        call = out._expr
        assert isinstance(call, FuncCall) and call.func == "SWITCH" and len(call.args) == 6
        assert isinstance(call.args[0], VarRef) and isinstance(call.args[1], Literal)
        with m:
            m.x.pick = SWITCH(m.x.z, 0, m.x.a, 1, m.x.q, -1.0)
        assert m.x.pick.dependencies == {m.x.z.id, m.x.a.id, m.x.q.id}

    def test_a_tracked_operand_lifts_per_track(self):
        m = _tracked()
        with m:
            m.w = SWITCH(m.x, 3.0, "three", 8.0, "eight", "other")
        assert dict(m.w._value.items()) == {"plan": ["three", "#DIV/0!"],
                                            "actual": ["other", "eight"]}


# ══════════════════════════════════════════════════════════════════════
# One call per period — the shared helper's law
# ══════════════════════════════════════════════════════════════════════


class TestOneCallPerPeriod:
    def test_scalars_and_a_one_value_series_broadcast(self):
        assert IFERROR(mo.Variable(["#N/A"]), mo.Variable([1.0, 2.0])).value == [1.0, 2.0]
        assert IFS(mo.Variable([True]), mo.Variable([1.0, 2.0]), True, 0.0).value == [1.0, 2.0]
        assert SWITCH(mo.Variable([2]), 1, "a", mo.Variable([2, 2]), "b").value == ["b", "b"]
        assert XOR(mo.Variable([True]), mo.Variable([True, False])).value == [False, True]

    def test_series_of_different_lengths_are_refused(self):
        a = mo.Variable([1.0, 2.0, 3.0])
        with pytest.raises(ValueError, match="different lengths"):
            IFERROR(a, mo.Variable([1.0, 2.0]))
        with pytest.raises(ValueError, match="different lengths"):
            IFS(a > 1.0, mo.Variable([1.0, 2.0]), True, 0.0)
        with pytest.raises(ValueError, match="different lengths"):
            SWITCH(a, 1.0, mo.Variable([1.0, 2.0]), "other")
        with pytest.raises(ValueError, match="different lengths"):
            XOR(a > 1.0, mo.Variable([True, False]))

    @pytest.mark.parametrize("name", sorted(FUNCTIONS))
    def test_an_axised_operand_is_refused(self, name):
        region = mo.Variable(["N", "S"], display_name="region")
        by_region = mo.Variable([1.0, 2.0], indexed_by=[region])
        args = {
            "IFERROR": (by_region, 0.0), "ISERROR": (by_region,), "ISNUMBER": (by_region,),
            "XOR": (by_region, True), "IFS": (True, by_region), "SWITCH": (by_region, 1.0, "x"),
        }[name]
        with pytest.raises(ValueError, match=name):
            FUNCTIONS[name](*args)


# ══════════════════════════════════════════════════════════════════════
# The registry and the loop
# ══════════════════════════════════════════════════════════════════════


class TestTheRegistry:
    @pytest.mark.parametrize("name", sorted(FUNCTIONS))
    def test_each_is_registered_as_one_call_per_period(self, name):
        assert spec(name) is not None and not spec(name).bare
        assert is_elementwise(name) and is_passthrough(name) and not is_range_taking(name)
        assert admits_in_loop(name, 2) and admits_in_loop(name, 6)
        assert excel_name(name) == name

    def test_the_names_excel_added_after_2007_carry_the_file_prefix(self):
        assert {n: file_prefix(n) for n in FUNCTIONS} == {
            "IFERROR": "", "ISERROR": "", "ISNUMBER": "",
            "XOR": "_xlfn.", "IFS": "_xlfn.", "SWITCH": "_xlfn.",
        }


class TestInsideALoop:
    def test_a_cap_written_as_ifs_runs_period_by_period(self):
        # The loop re-runs a call through the engine's exported function
        # of that name.
        m = mo.Model("m", default_grain="year", default_start="2026", default_periods=3)
        with m:
            m.d = mo.MultiVariable("D", loop=True)
            with m.d as d:
                d.bal = mo.lag(d.r, fill=10.0) * 1.37
                d.r = IFS(d.bal > 15.0, 15.0, True, d.bal)
        assert d.bal.value == pytest.approx([13.7, 18.769, 20.55])
        assert d.r.value == pytest.approx([13.7, 15.0, 15.0])

    def test_an_error_caught_inside_the_loop(self):
        # Inside the loop the call is run again over plain values: the
        # marker of a division by zero reaches IFERROR as text.
        m = mo.Model("m", default_grain="year", default_start="2026", default_periods=3)
        with m:
            m.i = mo.MultiVariable("I")
            m.i.k = mo.Variable([2.0, 0.0, 4.0])
            m.d = mo.MultiVariable("D", loop=True)
            with m.d as d:
                d.bal = mo.lag(d.r, fill=8.0) + 1.0
                d.r = IFERROR(d.bal / m.i.k, 1.0)
        assert d.bal.value == [9.0, 5.5, 2.0]
        assert d.r.value == [4.5, 1.0, 0.5]
