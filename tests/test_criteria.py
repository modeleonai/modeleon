# SPDX-License-Identifier: Apache-2.0
"""Conditional aggregates and the counts: the criteria read as the
spreadsheet reads it, a range a row or a list of scalar lines, one
number each with its formula."""

import datetime as dt

import pytest

import modeleon as mo
from modeleon.core.expr import FuncCall, ListExpr, Literal, VarRef
from modeleon.functions.criteria import (
    AVERAGEIF,
    AVERAGEIFS,
    COUNT,
    COUNTA,
    COUNTIF,
    COUNTIFS,
    MAXIFS,
    MINIFS,
    SUMIF,
    SUMIFS,
    matches,
    parse_criteria,
)
from modeleon.functions.registry import (
    admits_in_loop,
    file_prefix,
    is_passthrough,
    is_range_taking,
    range_positions,
    spec,
)

# ══════════════════════════════════════════════════════════════════════
# The criteria grammar
# ══════════════════════════════════════════════════════════════════════


class TestParseCriteria:
    @pytest.mark.parametrize("criteria, parsed", [
        (">=10", (">=", 10)),
        ("<>0", ("<>", 0)),
        ("<5", ("<", 5)),
        ("<=-2.5", ("<=", -2.5)),
        (">1e3", (">", 1000.0)),
        (">= 10", (">=", 10)),
        ("=text", ("=", "text")),
        ("text", ("=", "text")),
        ("N*", ("=", "N*")),
        ("", ("=", "")),
        ("=", ("=", None)),
        ("<>", ("<>", None)),
        (">=", (">=", None)),
        ("<>#DIV/0!", ("<>", "#DIV/0!")),
        ("5%", ("=", 0.05)),
        ("TRUE", ("=", True)),
        ("=false", ("=", False)),
        (10, ("=", 10)),
        (-3.5, ("=", -3.5)),
        (True, ("=", True)),
        (None, ("=", 0)),
        (dt.date(2026, 1, 1), ("=", dt.date(2026, 1, 1))),
        (dt.timedelta(hours=12), ("=", dt.timedelta(hours=12))),
        # A date written 2026-01-31 is that date, a time of day after it or not.
        (">=2026-01-01", (">=", dt.date(2026, 1, 1))),
        ("2026-1-31", ("=", dt.date(2026, 1, 31))),
        ("<2026-01-31 12:00", ("<", dt.datetime(2026, 1, 31, 12))),
        ("2026-01-31 9:05:30.5", ("=", dt.datetime(2026, 1, 31, 9, 5, 30, 500000))),
    ])
    def test_an_operator_and_a_comparand_read_as_a_cell_entry(self, criteria, parsed):
        assert parse_criteria(criteria) == parsed

    @pytest.mark.parametrize("text", [
        "1_000", "nan", "inf", "x10", "1e3%", "Q1 2026", "10-K", "A-100", "T1", "#5",
        "Jan", "100 USD", "Jan?2026", "10?20", "2026-02-30", "a.b", "(x)",
    ])
    def test_what_does_not_read_as_a_number_is_text(self, text):
        op, comparand = parse_criteria(text)
        assert isinstance(comparand, str) and not isinstance(comparand, bool)

    @pytest.mark.parametrize("text", [
        "1,000", ">$100", "100€", "(5)", "- 5", "5 %", "1,5", "1 000", "12:00", "<1 PM",
        "1/31/2026", "31.01.2026", "1/2", "10-20", "Jan 31, 2026", "31-Jan-2026", "Jan 2026",
        "2026-01-31T12:00", "1899-12-31", "2026-01-31 25:00",
    ])
    def test_an_entry_read_as_a_number_by_locale_is_refused(self, text):
        # The spreadsheet reads these as a number or a date in some locales
        # and as text in others: no answer is the workbook's in every one.
        with pytest.raises(ValueError, match="in some spreadsheets and as text in others"):
            parse_criteria(text)
        with pytest.raises(ValueError, match=r"^COUNTIF\(\): the criteria entry"):
            COUNTIF(mo.Variable([1, 2]), text)


class TestMatches:
    @pytest.mark.parametrize("value, criteria, hit", [
        # numbers
        (5, ">=5", True), (4.9, ">=5", False), (-6, "<-5", True), (-4, "<-5", False),
        (5, "<>5", False), (6, "<>5", True), (5, 5.0, True), (5, "=5", True),
        (0.05, "5%", True),
        # what is not a number never meets a numeric comparison but '<>'
        (None, ">0", False), (None, "<>5", True), ("x", ">0", False), ("x", "<>5", True),
        (True, 1, False), (True, "<>1", True), ("#DIV/0!", ">0", False), ("#DIV/0!", "<>0", True),
        # text that reads as a number meets an equality, never an ordering
        ("10", 10, True), ("10", "=10", True), ("10", "<>10", False), ("ten", 10, False),
        ("10", ">=10", False), ("9.5", "<10", False), ("-1", "<0", False), ("10", ">5", False),
        # two numbers are equal as the spreadsheet compares them
        (0.1 + 0.2, 0.3, True), (0.1 + 0.2, "<>0.3", False), (0.1 + 0.2, ">0.3", False),
        (0.1 + 0.2, ">=0.3", True), (0.7 - 0.4, "<=0.3", True), (0.7 - 0.4, "<0.3", False),
        (1 + 15 * 2.0 ** -52, 1, True), (1 + 16 * 2.0 ** -52, 1, False),
        (-(1 + 16 * 2.0 ** -52), -1, False), (2.0 ** 50 + 2, 2 ** 50, False), (1e-300, 0, False),
        # text, whole cell, case-insensitively
        ("North", "north", True), ("North", "NORTH", True), ("Northern", "North", False),
        ("North", "=North", True), ("North", "<>north", False), ("South", "<>north", True),
        (5, "North", False), (5, "<>North", True), (None, "<>North", True),
        # wildcards and the escape
        ("Northern", "North*", True), ("North", "N?rth", True), ("Nrth", "N?rth", False),
        ("abc", "*", True), ("", "*", True), (5, "*", False), (None, "*", False),
        ("a*b", "a~*b", True), ("axb", "a~*b", False), ("a?b", "a~?b", True),
        ("a~b", "a~~b", True), ("ab", "a~b", True), ("a~b", "a~b", False),
        ("a.b", "a.b", True), ("axb", "a.b", False),
        ("North", "~north", True), ("north~", "north~", True), ("north", "north~", False),
        # text ordered, case-insensitively
        ("b", ">a", True), ("B", ">a", True), ("a", ">a", False), ("a", ">=a", True),
        (5, ">a", False), (None, ">a", False),
        # booleans; the text TRUE meets the criteria written "TRUE"
        (True, True, True), (False, "FALSE", True), (1, True, False), (True, "<>FALSE", True),
        (None, "TRUE", False), ("TRUE", "TRUE", True), ("true", "TRUE", True),
        ("TRUE", "<>TRUE", False), ("TRUE", True, False), (1, "TRUE", False),
        # dates are their serials: a date cell meets a numeric criteria,
        # a number cell a date one
        (dt.date(2026, 1, 1), dt.date(2026, 1, 1), True),
        (dt.datetime(2026, 1, 1, 12), dt.date(2026, 1, 1), False),
        (dt.datetime(2026, 1, 1, 12), ">46023", True),
        (46023, dt.date(2026, 1, 1), True), ("46023", dt.date(2026, 1, 1), True),
        (dt.date(2026, 1, 1), 46023, True), (dt.date(2026, 1, 1), ">45000", True),
        (dt.date(2025, 12, 31), ">=46023", False), (dt.date(2026, 1, 1), "<>", True),
        ("x", dt.date(2026, 1, 1), False), (None, "<>46023", True),
        # a date written 2026-01-31 is a date; the text equal to it meets it too
        (dt.date(2026, 1, 31), "2026-01-31", True), (dt.date(2026, 1, 1), ">=2026-01-01", True),
        (dt.date(2025, 12, 31), ">=2026-01-01", False), (46053, "2026-01-31", True),
        (dt.datetime(2026, 1, 31, 12), "2026-01-31", False),
        (dt.datetime(2026, 1, 31, 12), "2026-01-31 12:00", True),
        ("2026-01-31", "2026-01-31", True), ("2026-01-31", dt.date(2026, 1, 31), False),
        ("2026-01-31", ">=2026-01-01", False),
        # errors
        ("#N/A", "#N/A", True), ("#n/a", "#N/A", True), (5, "#N/A", False),
        ("#N/A", "x", False), ("#N/A", "<>x", True), ("#N/A", "<>#N/A", False),
        # blanks and empty text ('' here is a formula's empty text, a cell;
        # typed into a line it is read as a blank — TestCOUNTIF, TestCOUNTA)
        (None, "", True), ("", "", True), (0, "", False), (None, "=", True), ("", "=", False),
        (None, "<>", False), ("", "<>", True), (0, "<>", True), ("x", "<>", True),
        # an ordering with nothing after it compares text with the empty text
        ("", ">=", True), ("x", ">=", True), (None, ">=", False), (5, ">=", False),
        ("x", ">", True), ("", ">", False), ("", "<=", True), ("x", "<=", False),
        ("", "<", False), (True, ">=", False), ("#N/A", ">=", False),
        # None as the criteria is 0, as a reference to an empty cell reads
        (0, None, True), (0.0, None, True), (None, None, False), ("", None, False),
        ("0", None, True), (False, None, False),
    ])
    def test_one_cell_against_one_criteria(self, value, criteria, hit):
        assert matches(value, criteria) is hit


# ══════════════════════════════════════════════════════════════════════
# The functions, over rows
# ══════════════════════════════════════════════════════════════════════


def _rows():
    m = mo.Model("m", default_grain="month", default_start="2026-01", default_periods=4)
    with m:
        m.d = mo.MultiVariable("D")
        m.d.region = mo.Variable(["North", "South", "North", "West"], display_name="region")
        m.d.amount = mo.Variable([100.0, 200.0, 300.0, 400.0], display_name="amount", unit="USD")
        m.d.units = mo.Variable([1, 2, 3, 4], display_name="units", unit="pcs")
        m.d.year = mo.Variable([2026, 2026, 2027, 2027], display_name="year")
        m.i = mo.MultiVariable("I")
        m.i.floor = mo.Variable(250.0, display_name="floor")
        m.i.pick = mo.Variable("north", display_name="pick")
        m.i.rule = mo.Variable(">=200", display_name="rule")
    return m


class TestSUMIF:
    def test_over_a_sum_range_and_over_the_range_itself(self):
        m = _rows()
        assert SUMIF(m.d.region, "North", m.d.amount).value == 400.0
        assert SUMIF(m.d.amount, ">=200").value == 900.0
        assert SUMIF(m.d.amount, "<>100").value == 900.0

    def test_every_criteria_kind(self):
        m = _rows()
        assert SUMIF(m.d.region, m.i.pick, m.d.amount).value == 400.0
        assert SUMIF(m.d.amount, m.i.rule).value == 900.0
        assert SUMIF(m.d.amount, (">=", m.i.floor)).value == 700.0
        assert SUMIF(m.d.amount, mo.CONCAT(">=", m.i.floor)).value == 700.0
        assert SUMIF(m.d.amount, (">", 250)).value == 700.0
        assert SUMIF(m.d.year, 2027, m.d.amount).value == 700.0
        assert SUMIF(m.d.region, "n*", m.d.amount).value == 400.0

    def test_a_one_cell_range(self):
        m = _rows()
        assert SUMIF(m.i.floor, ">=100").value == 250.0
        assert SUMIF(m.i.floor, "<100").value == 0

    def test_nothing_matched_is_zero(self):
        m = _rows()
        assert SUMIF(m.d.region, "East", m.d.amount).value == 0

    def test_the_sum_range_skips_blanks_text_and_booleans(self):
        flags = mo.Variable(["y", "y", "y", "y", "y"])
        mixed = mo.Variable([10.0, None, "text", True, 5.0])
        assert SUMIF(flags, "y", mixed).value == 15.0

    def test_an_error_among_the_matched_cells_is_the_result(self):
        flags = mo.Variable(["y", "n", "y"])
        assert SUMIF(flags, "y", mo.Variable([1.0, "#DIV/0!", 2.0])).value == 3.0
        assert SUMIF(flags, "y", mo.Variable([1.0, 2.0, "#DIV/0!"])).value == "#DIV/0!"
        assert SUMIF(flags, "y", mo.Variable(["#N/A", 2.0, "#DIV/0!"])).value == "#N/A"

    def test_an_error_in_the_criteria_range_matches_only_its_text(self):
        keys = mo.Variable(["a", "#DIV/0!", "a"])
        amounts = mo.Variable([1.0, 2.0, 3.0])
        assert SUMIF(keys, "a", amounts).value == 4.0
        assert SUMIF(keys, "#DIV/0!", amounts).value == 2.0
        assert SUMIF(keys, "<>a", amounts).value == 2.0

    def test_the_result_is_a_scalar_whose_type_follows_the_summed_range(self):
        m = _rows()
        out = SUMIF(m.d.region, "North", m.d.amount)
        assert out.var_type == "scalar" and out.value_type == "float"
        assert SUMIF(m.d.region, "North", m.d.units).value_type == "int"
        assert SUMIF(m.d.region, "North", m.d.units).value == 4
        # A sum is a number, over a range of text too.
        assert SUMIF(m.d.region, "North").value_type == "float"

    def test_the_unit_is_the_summed_range_s_never_the_criteria_range_s(self):
        from modeleon.core.unit import unit_key
        m = _rows()
        assert unit_key(SUMIF(m.d.region, "North", m.d.amount)._unit) == unit_key(m.d.amount._unit)
        assert unit_key(SUMIF(m.d.units, ">1", m.d.amount)._unit) == unit_key(m.d.amount._unit)
        assert SUMIF(m.d.units, ">1", m.d.year)._unit is None
        assert SUMIF(m.d.amount, ">1", m.d.year)._unit is None

    def test_ranges_of_different_lengths_are_refused(self):
        m = _rows()
        with pytest.raises(ValueError, match="same number of cells"):
            SUMIF(m.d.region, "North", mo.Variable([1.0, 2.0]))

    def test_a_series_criteria_is_refused(self):
        m = _rows()
        with pytest.raises(ValueError, match="a criteria is one value"):
            SUMIF(m.d.amount, m.d.units)
        with pytest.raises(ValueError, match="a criteria is one value"):
            SUMIF(m.d.amount, [1, 2])

    def test_a_bad_pair_is_refused(self):
        m = _rows()
        with pytest.raises(ValueError, match=r"\(operator, value\)"):
            SUMIF(m.d.amount, ("~", 5))
        with pytest.raises(ValueError, match=r"\(operator, value\)"):
            SUMIF(m.d.amount, (">=", 5, 6))

    def test_a_line_with_tracks_is_refused(self):
        m = mo.Model("m", default_grain="month", default_start="2026-01", default_periods=2,
                     tracks=mo.Tracks("plan", "actual"))
        with m:
            m.d = mo.MultiVariable("D")
            m.d.amount = mo.Variable(plan=[1.0, 2.0], actual=[1.5, 2.5])
        with pytest.raises(ValueError, match="tracks"):
            SUMIF(m.d.amount, ">1")

    def test_a_range_that_is_neither_a_line_nor_a_list_is_refused(self):
        with pytest.raises(TypeError, match="a range is a list-valued line"):
            SUMIF(5, ">1")


class TestSUMIFS:
    def test_every_pair_must_hold(self):
        m = _rows()
        assert SUMIFS(m.d.amount, m.d.region, "North", m.d.year, 2027).value == 300.0
        assert SUMIFS(m.d.amount, m.d.region, "North").value == 400.0
        assert SUMIFS(m.d.amount, m.d.region, "<>North", m.d.amount, (">", m.i.floor)).value == 400.0
        assert SUMIFS(m.d.amount, m.d.region, "East").value == 0

    def test_the_pairs_are_checked(self):
        m = _rows()
        with pytest.raises(ValueError, match="in pairs"):
            SUMIFS(m.d.amount)
        with pytest.raises(ValueError, match="in pairs"):
            SUMIFS(m.d.amount, m.d.region)
        with pytest.raises(ValueError, match="same number of cells"):
            SUMIFS(m.d.amount, mo.Variable([1, 2]), 1)

    def test_type_and_unit_follow_the_summed_range(self):
        from modeleon.core.unit import unit_key
        m = _rows()
        out = SUMIFS(m.d.units, m.d.region, "North")
        assert out.value == 4 and out.value_type == "int" and out.var_type == "scalar"
        assert unit_key(out._unit) == unit_key(m.d.units._unit)


class TestCOUNTIF:
    def test_counts_the_cells_that_meet_the_criteria(self):
        m = _rows()
        assert COUNTIF(m.d.region, "North").value == 2
        assert COUNTIF(m.d.region, "n*").value == 2
        assert COUNTIF(m.d.region, "<>North").value == 2
        assert COUNTIF(m.d.amount, ">=200").value == 3
        assert COUNTIF(m.d.amount, (">", m.i.floor)).value == 2
        assert COUNTIF(m.d.year, m.d.year[0]).value == 2

    def test_blanks_numeric_text_and_errors(self):
        row = mo.Variable([None, "", "10", 10, "#N/A", True])
        assert COUNTIF(row, "").value == 2
        assert COUNTIF(row, "<>").value == 4
        assert COUNTIF(row, 10).value == 2
        assert COUNTIF(row, "#N/A").value == 1
        assert COUNTIF(row, "<>10").value == 4
        assert COUNTIF(row, True).value == 1

    def test_empty_text_typed_in_is_a_blank_cell_and_a_formula_s_is_not(self):
        # Typed in, '' is written as nothing: a blank. A formula's '' is
        # a cell the spreadsheet counts (and '' matches it too).
        typed = mo.Variable(["", "x", None])
        assert (COUNTIF(typed, "").value, COUNTIF(typed, "<>").value) == (2, 1)
        assert COUNTIF(typed, "*").value == 1
        formula = mo.IF(mo.Variable([True, False, True]), "", "x")
        assert (COUNTIF(formula, "").value, COUNTIF(formula, "<>").value) == (2, 3)
        assert COUNTIF(formula, "*").value == 3
        # One element of a row is the row's cell: a typed-in row's '' is blank.
        assert COUNTIF([typed[0], typed[1]], "<>").value == 1
        assert COUNTIF([formula[0], formula[1]], "<>").value == 2

    def test_none_is_the_number_zero_typed_or_in_a_blank_cell(self):
        row = mo.Variable([0.0, None, 5.0, "", 0])
        assert COUNTIF(row, None).value == 2
        assert COUNTIF(row, mo.Variable(None)).value == 2
        assert COUNTIF(row, mo.Variable("")).value == 2
        assert SUMIF(row, None, mo.Variable([1, 2, 3, 4, 5])).value == 6
        assert COUNTIF(mo.Variable(["x", None, ""]), mo.Variable(None)).value == 0

    def test_equals_alone_is_a_blank_cell_and_nothing_else(self):
        # '' matches a blank cell and a formula's empty text; '=' only the
        # blank; an ordering alone compares text with the empty text.
        formula = mo.IF(mo.Variable([True, False, True]), "", "x")
        typed = mo.Variable(["", "x", None])
        assert (COUNTIF(formula, "=").value, COUNTIF(formula, "").value) == (0, 2)
        assert (COUNTIF(typed, "=").value, COUNTIF(typed, "").value) == (2, 2)
        assert (COUNTIF(formula, ">=").value, COUNTIF(formula, ">").value) == (3, 1)
        assert (COUNTIF(formula, "<=").value, COUNTIF(formula, "<").value) == (2, 0)

    def test_an_operator_joined_to_a_blank_cell_or_none_is_the_operator_alone(self):
        row = mo.Variable(["a", None, 5.0])
        blank = mo.Variable(None)
        assert COUNTIF(row, ("=", blank)).value == 1
        assert COUNTIF(row, ("<>", blank)).value == 2
        assert COUNTIF(row, (">=", blank)).value == 1
        assert COUNTIF(row, ("=", None)).value == 1

    def test_numbers_are_equal_as_the_spreadsheet_compares_them(self):
        share = mo.Variable([0.1, 0.2, 0.7]) + mo.Variable([0.2, 0.1, -0.4])
        assert share.value[0] != 0.3 and share.value[2] != 0.3
        assert COUNTIF(share, 0.3).value == 3
        assert COUNTIF(share, mo.Variable(0.3)).value == 3
        assert (COUNTIF(share, ">0.3").value, COUNTIF(share, ">=0.3").value) == (0, 3)
        assert COUNTIF(share, "<>0.3").value == 0
        assert SUMIF(share, 0.3, mo.Variable([1.0, 2.0, 4.0])).value == 7.0

    def test_numeric_text_meets_an_equality_and_never_an_ordering(self):
        row = mo.Variable(["10", "3", 7.0, None, "abc", "-1"])
        amounts = mo.Variable([1.0, 2.0, 4.0, 8.0, 16.0, 32.0])
        assert (COUNTIF(row, ">5").value, COUNTIF(row, ">=0").value) == (1, 1)
        assert COUNTIF(row, "<0").value == 0
        assert SUMIF(row, ">5", amounts).value == 4.0
        assert (COUNTIF(row, 10).value, COUNTIF(row, "10").value) == (1, 1)
        assert COUNTIF(row, "<>10").value == 5

    def test_the_text_true_meets_the_criteria_written_true(self):
        row = mo.Variable(["TRUE", True, 1, "x"])
        assert (COUNTIF(row, "TRUE").value, COUNTIF(row, "<>TRUE").value) == (2, 2)
        assert COUNTIF(row, True).value == 1
        assert COUNTIF(row, ("=", True)).value == 2

    def test_a_criteria_cell_holding_an_error_is_the_result(self):
        errs = mo.Variable([1.0, "#DIV/0!", 3.0, "#N/A"])
        flags = mo.Variable(["y", "y", "n", "y"])
        cell = mo.Variable("#DIV/0!")
        assert COUNTIF(errs, cell).value == "#DIV/0!"
        assert COUNTIF(errs, (">", cell)).value == "#DIV/0!"
        assert SUMIF(flags, cell, errs).value == "#DIV/0!"
        assert SUMIFS(errs, flags, "y", flags, mo.Variable("#N/A")).value == "#N/A"
        assert COUNTIFS(flags, "y", errs, cell).value == "#DIV/0!"
        assert AVERAGEIF(flags, cell).value == "#DIV/0!"
        assert MAXIFS(errs, flags, mo.Variable("#N/A")).value == "#N/A"
        # Typed into the call, an error's text is a criteria: it matches the error.
        assert COUNTIF(errs, "#DIV/0!").value == 1 and COUNTIF(errs, "<>#N/A").value == 3
        # Text that only starts with '#' is text.
        assert COUNTIF(mo.Variable(["#1", "#2"]), mo.Variable("#1")).value == 1

    def test_the_count_is_an_int_without_a_unit(self):
        m = _rows()
        out = COUNTIF(m.d.amount, ">0")
        assert out.value_type == "int" and out.var_type == "scalar" and out._unit is None


class TestCOUNTIFS:
    def test_every_pair_must_hold(self):
        m = _rows()
        assert COUNTIFS(m.d.region, "North", m.d.year, 2026).value == 1
        assert COUNTIFS(m.d.region, "<>North", m.d.amount, ">100").value == 2
        assert COUNTIFS(m.d.region, "North").value == 2
        assert COUNTIFS(m.d.region, "East").value == 0

    def test_the_pairs_are_checked(self):
        m = _rows()
        with pytest.raises(ValueError, match="in pairs"):
            COUNTIFS(m.d.region)
        with pytest.raises(ValueError, match="same number of cells"):
            COUNTIFS(m.d.region, "North", mo.Variable([1]), 1)
        assert COUNTIFS(m.d.region, "North", m.d.year, 2026)._unit is None


class TestAVERAGEIF:
    def test_mean_of_the_picked_numbers(self):
        m = _rows()
        assert AVERAGEIF(m.d.region, "North", m.d.amount).value == 200.0
        assert AVERAGEIF(m.d.amount, ">100").value == 300.0
        assert AVERAGEIFS(m.d.amount, m.d.region, "<>North", m.d.year, 2027).value == 400.0

    def test_blanks_and_text_are_not_counted_and_nothing_is_div0(self):
        flags = mo.Variable(["y", "y", "y", "n"])
        assert AVERAGEIF(flags, "y", mo.Variable([10.0, None, "x", 99.0])).value == 10.0
        assert AVERAGEIF(flags, "z", mo.Variable([10.0, 20.0, 30.0, 40.0])).value == "#DIV/0!"
        assert AVERAGEIFS(mo.Variable([None, None]), flags[0:2], "y").value == "#DIV/0!"

    def test_an_error_among_the_matched_cells_is_the_result(self):
        flags = mo.Variable(["y", "n", "y"])
        assert AVERAGEIF(flags, "y", mo.Variable([1.0, "#DIV/0!", 3.0])).value == 2.0
        assert AVERAGEIF(flags, "y", mo.Variable([1.0, 2.0, "#NUM!"])).value == "#NUM!"

    def test_a_float_in_the_averaged_range_s_unit(self):
        from modeleon.core.unit import unit_key
        m = _rows()
        out = AVERAGEIF(m.d.region, "North", m.d.units)
        assert out.value == 2.0 and out.value_type == "float"
        assert unit_key(out._unit) == unit_key(m.d.units._unit)
        assert AVERAGEIFS(m.d.amount, m.d.units, ">1")._unit is not None
        assert AVERAGEIFS(m.d.year, m.d.units, ">1")._unit is None


class TestMAXIFSAndMINIFS:
    def test_the_largest_and_smallest_of_the_picked(self):
        m = _rows()
        assert MAXIFS(m.d.amount, m.d.region, "North").value == 300.0
        assert MINIFS(m.d.amount, m.d.region, "North").value == 100.0
        assert MAXIFS(m.d.amount, m.d.region, "<>North", m.d.year, 2026).value == 200.0
        assert MINIFS(m.d.units, m.d.year, 2027).value == 3

    def test_negatives_rank_and_nothing_is_zero(self):
        flags = mo.Variable(["y", "y", "n"])
        negatives = mo.Variable([-5.0, -3.0, 9.0])
        assert MAXIFS(negatives, flags, "y").value == -3.0
        assert MINIFS(negatives, flags, "y").value == -5.0
        assert MAXIFS(negatives, flags, "z").value == 0
        assert MINIFS(negatives, flags, "z").value == 0

    def test_text_and_blanks_are_skipped_and_an_error_is_the_result(self):
        flags = mo.Variable(["y", "y", "y", "n"])
        assert MAXIFS(mo.Variable(["x", None, 2.0, 9.0]), flags, "y").value == 2.0
        assert MAXIFS(mo.Variable([1.0, "#REF!", 2.0, 9.0]), flags, "y").value == "#REF!"

    def test_type_and_unit_follow_the_ranked_range(self):
        from modeleon.core.unit import unit_key
        m = _rows()
        out = MAXIFS(m.d.units, m.d.region, "North")
        assert out.value == 3 and out.value_type == "int" and out.var_type == "scalar"
        assert unit_key(out._unit) == unit_key(m.d.units._unit)
        # Over text nothing ranks: 0, a number.
        nothing = MAXIFS(m.d.region, m.d.region, "N*")
        assert nothing.value == 0 and nothing.value_type == "float"
        with pytest.raises(ValueError, match="in pairs"):
            MINIFS(m.d.units)


class TestDatesAreNumbers:
    def _dates(self):
        m = mo.Model("m", default_grain="month", default_start="2026-01", default_periods=4)
        with m:
            m.d = mo.MultiVariable("D")
            m.d.region = mo.Variable(["North", "South", "North", "West"], display_name="region")
            m.d.closed = mo.Variable([dt.date(2026, 1, 31), dt.date(2026, 2, 28),
                                      dt.date(2024, 2, 29), None], display_name="closed")
            m.d.amount = mo.Variable([100.0, 200.0, 300.0, 400.0], display_name="amount")
        return m

    def test_a_date_cell_meets_a_numeric_criteria_and_a_number_a_date_one(self):
        m = self._dates()
        assert COUNTIF(m.d.closed, ">45000").value == 3
        assert COUNTIF(m.d.closed, 46053).value == 1
        assert COUNTIF(m.d.closed, (">=", dt.date(2026, 1, 1))).value == 2
        assert COUNTIF(m.d.closed, dt.date(2024, 2, 29)).value == 1
        assert COUNTIF(mo.Variable([46053, 46053.5, "46053"]), dt.date(2026, 1, 31)).value == 2
        assert SUMIF(m.d.closed, (">", dt.date(2026, 1, 31)), m.d.amount).value == 200.0

    def test_the_largest_and_the_smallest_date_come_back_as_dates(self):
        m = self._dates()
        latest = MAXIFS(m.d.closed, m.d.region, "North")
        assert latest.value == dt.date(2026, 1, 31) and latest.value_type == "datetime"
        assert MINIFS(m.d.closed, m.d.region, "North").value == dt.date(2024, 2, 29)
        assert MAXIFS(m.d.closed, m.d.region, "West").value == 0
        assert MAXIFS(m.d.closed, m.d.region, "<>").value == dt.date(2026, 2, 28)

    def test_the_mean_of_dates_is_a_date(self):
        m = self._dates()
        mean = AVERAGEIF(m.d.region, "<>West", m.d.closed)
        assert mean.value == dt.datetime(2025, 6, 20, 8) and mean.value_type == "datetime"
        whole = AVERAGEIFS(m.d.closed, m.d.region, "<>North")
        assert whole.value == dt.date(2026, 2, 28) and whole.value_type == "datetime"
        assert AVERAGEIF(m.d.region, "West", m.d.closed).value == "#DIV/0!"
        assert AVERAGEIF(m.d.region, "North", m.d.amount).value_type == "float"

    def test_a_sum_of_dates_is_a_number_of_days(self):
        m = self._dates()
        total = SUMIF(m.d.region, "North", m.d.closed)
        assert total.value == 46053 + 45351 and total.value_type == "float"
        assert SUMIFS(m.d.closed, m.d.region, "West").value == 0

    def test_a_date_written_2026_01_31_is_that_date(self):
        m = self._dates()
        assert COUNTIF(m.d.closed, ">=2026-01-01").value == 2
        assert COUNTIF(m.d.closed, "2026-01-31").value == 1
        assert SUMIF(m.d.closed, "<2026-02-01", m.d.amount).value == 400.0
        noon = mo.Variable([dt.datetime(2026, 1, 31, 12), dt.date(2026, 1, 31)])
        assert COUNTIF(noon, "2026-01-31 12:00").value == 1
        assert COUNTIF(noon, ">2026-01-31").value == 1

    def test_a_duration_is_its_days(self):
        days = mo.Variable([dt.timedelta(days=1), dt.timedelta(hours=12), None])
        assert COUNTIF(days, ">=0.5").value == 2
        assert SUMIF(days, ">0").value == 1.5
        assert MAXIFS(days, days, ">0").value == dt.timedelta(days=1)


# ══════════════════════════════════════════════════════════════════════
# The counts
# ══════════════════════════════════════════════════════════════════════


class TestCOUNT:
    def test_numbers_and_dates_in_cells_count_and_nothing_else_does(self):
        row = mo.Variable([1, 2.5, None, "x", "3", True, "#N/A", dt.date(2026, 1, 1)])
        assert COUNT(row).value == 3

    def test_values_typed_into_the_call_count_as_the_spreadsheet_counts_them(self):
        # None typed in is written as 0, a number.
        assert COUNT(1, 2.5, "3", True, "x", None).value == 5
        assert COUNT(dt.timedelta(days=1), mo.Variable(dt.timedelta(hours=1))).value == 2
        # Beside a series the call is one per period: the
        # duration typed in counts in every period.
        assert COUNT(dt.timedelta(days=1),
                     mo.Variable([dt.timedelta(hours=1), None])).value == [2, 1]
        assert COUNT([1, 2, "3", True]).value == 2
        assert COUNT(mo.Variable(True), mo.Variable("3")).value == 0

    def test_over_rows_cells_and_lists_of_lines(self):
        m = _rows()
        assert COUNT(m.d.amount).value == 4
        # Two rows are counted period by period, as a + b
        # combines them; their cells over all periods are two counts.
        assert COUNT(m.d.amount, m.d.region).value == [1, 1, 1, 1]
        assert (COUNT(m.d.amount) + COUNT(m.d.region)).value == 4
        assert COUNT(m.i.floor, m.i.pick).value == 1
        assert COUNT([m.i.floor, m.i.pick, m.d.amount[0]]).value == 2

    def test_an_int_without_a_unit_and_at_least_one_argument(self):
        m = _rows()
        out = COUNT(m.d.amount)
        assert out.value_type == "int" and out.var_type == "scalar" and out._unit is None
        with pytest.raises(ValueError, match="at least one argument"):
            COUNT()


class TestCOUNTA:
    def test_whatever_is_not_blank_counts(self):
        row = mo.Variable([1, None, "x", "", True, "#N/A", 0])
        assert COUNTA(row).value == 5
        assert COUNTA(1, None, "", "x").value == 4
        assert COUNTA([mo.Variable(None), mo.Variable("")]).value == 0
        # A formula's empty text is a cell.
        assert COUNTA(mo.IF(mo.Variable([True, False]), "", "x")).value == 2

    def test_over_rows_and_cells(self):
        m = _rows()
        assert COUNTA(m.d.region).value == 4
        # Cells beside a row are counted with each period's cell.
        assert COUNTA(m.i.floor, m.i.pick, m.d.amount).value == [3, 3, 3, 3]
        assert (COUNTA(m.i.floor, m.i.pick) + COUNTA(m.d.amount)).value == 6
        out = COUNTA(m.d.region)
        assert out.value_type == "int" and out.var_type == "scalar" and out._unit is None
        with pytest.raises(ValueError, match="at least one argument"):
            COUNTA()


# ══════════════════════════════════════════════════════════════════════
# A list of scalar lines as a range
# ══════════════════════════════════════════════════════════════════════


class Deal(mo.MultiVariableClass):
    def compute(self, region, amount):
        self.region = mo.Variable(region, display_name="Region")
        self.amount = mo.Variable(amount, display_name="Amount", unit="USD")
        self.flow = mo.Variable([amount * 0.1, amount * 0.2], display_name="Flow")


def _records():
    m = mo.Model("m", default_grain="month", default_start="2026-01", default_periods=2)
    with m:
        m.deals = mo.MultiVariable("Deals", excel_props={"orient": "records"})
        with m.deals as d:
            d.a = Deal(region="North", amount=100.0, display_name="Alpha")
            d.b = Deal(region="South", amount=200.0, display_name="Beta")
            d.c = Deal(region="North", amount=300.0, display_name="Gamma")
    return m


class TestAListOfLinesIsARange:
    def test_the_values_reduce_as_a_row_does(self):
        d = _records().deals
        regions = [d.a.region, d.b.region, d.c.region]
        amounts = [d.a.amount, d.b.amount, d.c.amount]
        assert SUMIF(regions, "North", amounts).value == 400.0
        assert SUMIFS(amounts, regions, "North", amounts, ">100").value == 300.0
        assert COUNTIF(regions, "North").value == 2
        assert COUNTIFS(regions, "North", amounts, ">100").value == 1
        assert AVERAGEIF(regions, "North", amounts).value == 200.0
        assert AVERAGEIFS(amounts, regions, "North").value == 200.0
        assert MAXIFS(amounts, regions, "North").value == 300.0
        assert MINIFS(amounts, regions, "North").value == 100.0
        assert COUNT(amounts).value == 3 and COUNTA(regions).value == 3

    def test_one_element_of_a_row_is_a_line_too(self):
        d = _records().deals
        first = [d.a.flow[0], d.b.flow[0], d.c.flow[0]]
        out = SUMIFS(first, [d.a.region, d.b.region, d.c.region], "North")
        assert out.value == pytest.approx(40.0)
        assert out._source_code == "SUMIFS([a.flow[0], b.flow[0], c.flow[0]], [a.region, b.region, c.region], 'North')"

    def test_a_tuple_a_mixed_list_and_a_literal_list(self):
        d = _records().deals
        assert SUMIF((d.a.region, d.b.region), "North", (d.a.amount, d.b.amount)).value == 100.0
        assert SUMIF([d.a.region, "North"], "North", [d.a.amount, 7.0]).value == 107.0
        assert SUMIF([1, 2, 3], ">1").value == 5

    def test_a_row_inside_the_list_and_an_empty_list_are_refused(self):
        d = _records().deals
        with pytest.raises(ValueError, match=r"a\.flow in the list is a series"):
            SUMIF([d.a.flow, d.b.flow], ">1")
        with pytest.raises(ValueError, match="at least one cell"):
            SUMIF([], ">1")

    def test_the_unit_and_type_are_the_lines_shared_ones(self):
        from modeleon.core.unit import unit_key
        d = _records().deals
        out = SUMIF([d.a.region, d.b.region], "North", [d.a.amount, d.b.amount])
        assert unit_key(out._unit) == unit_key(d.a.amount._unit) and out.value_type == "float"
        # A line without a unit does not vote; a second unit is refused,
        # as SUM refuses it.
        assert SUMIF([d.a.region, d.b.region], "North",
                     [d.a.amount, mo.Variable(5.0)])._unit is not None
        with pytest.raises(ValueError, match="mo.SUMIF takes its values in one unit"):
            SUMIF([d.a.region, d.b.region], "North",
                  [d.a.amount, mo.Variable(5.0, unit="EUR")])


# ══════════════════════════════════════════════════════════════════════
# The AST, the formula and the registry
# ══════════════════════════════════════════════════════════════════════


class TestTheAST:
    def test_a_funccall_over_refs_literals_and_a_list(self):
        m = _rows()
        d = _records().deals
        out = SUMIF(m.d.region, ">=10", m.d.amount)
        assert isinstance(out._expr, FuncCall) and out._expr.func == "SUMIF"
        assert [type(a) for a in out._expr.args] == [VarRef, Literal, VarRef]
        assert out._expr.args[1].value == ">=10"
        assert out.formula == "SUMIF(region, '>=10', amount)"
        assert out._source_code == "SUMIF(region, '>=10', amount)"
        listed = SUMIF([d.a.region, d.b.region], "North", [d.a.amount, d.b.amount])
        assert isinstance(listed._expr.args[0], ListExpr)
        assert [type(i) for i in listed._expr.args[0].items] == [VarRef, VarRef]
        assert listed._source_code == "SUMIF([a.region, b.region], 'North', [a.amount, b.amount])"
        assert sorted(listed.dependencies) == sorted(
            v.id for v in (d.a.region, d.b.region, d.a.amount, d.b.amount))

    def test_a_criteria_cell_and_an_operator_joined_to_one(self):
        m = _rows()
        cell = SUMIF(m.d.region, m.i.pick, m.d.amount)
        assert isinstance(cell._expr.args[1], VarRef) and cell._expr.args[1].var is m.i.pick
        assert cell._source_code == "SUMIF(region, pick, amount)"
        joined = SUMIF(m.d.amount, (">=", m.i.floor))
        node = joined._expr.args[1]
        assert isinstance(node, FuncCall) and node.func == "CONCAT"
        assert isinstance(node.args[0], Literal) and node.args[0].value == ">="
        assert isinstance(node.args[1], VarRef) and node.args[1].var is m.i.floor
        assert joined._source_code == "SUMIF(amount, CONCAT('>=', floor))"
        assert m.i.floor.id in joined.dependencies
        plain = SUMIF(m.d.amount, (">=", 250))
        assert isinstance(plain._expr.args[1], Literal) and plain._expr.args[1].value == ">=250"
        dated = COUNTIF(mo.Variable([dt.date(2026, 1, 1), dt.date(2026, 6, 1)]),
                        (">=", dt.date(2026, 3, 1)))
        assert dated.value == 1 and isinstance(dated._expr.args[1], FuncCall)

    def test_a_pair_is_written_as_the_text_it_joins(self):
        m = _rows()
        alone = SUMIF(m.d.amount, ("=", None))
        assert isinstance(alone._expr.args[1], Literal) and alone._expr.args[1].value == "="
        flags = COUNTIF(mo.Variable([True, "TRUE", False]), ("=", True))
        assert flags._expr.args[1].value == "=TRUE" and flags.value == 2
        # A time of day cannot be written as DATE(...): the serial is.
        noon = dt.datetime(2026, 1, 31, 12)
        cells = mo.Variable([noon, dt.date(2026, 1, 31)])
        typed = COUNTIF(cells, noon)
        assert typed._expr.args[1].value == 46053.5 and typed.value == 1
        later = COUNTIF(cells, (">=", noon))
        assert later._expr.args[1].value == ">=46053.5" and later.value == 1

    def test_the_counts_keep_sum_s_source_shape(self):
        m = _rows()
        assert COUNT(m.d.amount, 5)._source_code == "COUNT(amount, 5)"
        assert COUNTA(m.d.region).formula == "COUNTA(region)"
        assert isinstance(COUNT([1, 2])._expr.args[0], ListExpr)

    def test_every_name_is_registered_as_a_range_function_and_written_as_given(self):
        for name in ("SUMIF", "SUMIFS", "COUNTIF", "COUNTIFS", "AVERAGEIF", "AVERAGEIFS",
                     "MAXIFS", "MINIFS"):
            assert spec(name).shape == "range" and range_positions(name) is None, name
            assert is_range_taking(name) and is_passthrough(name), name
            assert not admits_in_loop(name, 3), name
        assert file_prefix("MAXIFS") == file_prefix("MINIFS") == "_xlfn."
        assert file_prefix("SUMIFS") == file_prefix("COUNTIFS") == ""
        assert spec("COUNTA").shape == "reduction" and is_range_taking("COUNTA")
        assert spec("COUNT").shape == "reduction"

    def test_the_json_renderer_takes_the_list_argument(self):
        d = _records().deals
        out = SUMIF([d.a.region, d.b.region], "North", [d.a.amount, d.b.amount])
        from modeleon.compile.json import to_json
        rendered = to_json(out._expr)
        assert "SUMIF" in str(rendered)
        assert {"kind": "list"}.items() <= rendered["args"][0].items()
        assert len(rendered["args"][0]["items"]) == 2
