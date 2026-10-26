# SPDX-License-Identifier: Apache-2.0
"""Statistics: the middle, the spread, percentiles, quartiles, the k-th
largest and smallest, and ranks — read as the spreadsheet reads them, over
a row or one field down the records of a table, one cell with its
formula. Several series, one call per period: test_statistics_per_period.py."""

import datetime as dt
import math
import statistics

import pytest

import modeleon as mo
from modeleon.core.expr import FuncCall, ListExpr, Literal, VarRef
from modeleon.functions.registry import (
    admits_in_loop,
    excel_name,
    file_prefix,
    is_passthrough,
    per_period_arity,
    range_positions,
    spec,
)
from modeleon.functions.statistics import (
    LARGE,
    MEDIAN,
    PERCENTILE,
    PERCENTILE_EXC,
    PERCENTILE_INC,
    QUARTILE,
    QUARTILE_EXC,
    QUARTILE_INC,
    RANK,
    RANK_AVG,
    RANK_EQ,
    SMALL,
    STDEV,
    STDEV_P,
    STDEV_S,
    STDEVP,
    VAR,
    VAR_P,
    VAR_S,
    VARP,
)

DATA = [10, 20, 30, 40, 50]
TIES = [5, 3, 5, 1, 3, 5]


def _row(values=DATA, **kwargs):
    return mo.Variable(list(values), **kwargs)


class Comp(mo.MultiVariableClass):
    def compute(self, multiple, region="North"):
        self.region = mo.Variable(region, display_name="Region")
        self.multiple = mo.Variable(multiple, display_name="Multiple", unit="x")


def _comps(*multiples):
    m = mo.Model("m", display_name="M", default_grain="year", default_start="2026",
                 default_periods=3)
    with m:
        m.comps = mo.MultiVariable("Comps", excel_props={"tab": True, "orient": "records"})
        with m.comps as c:
            for i, x in enumerate(multiples):
                setattr(c, f"r{i}", Comp(multiple=x, display_name=f"R{i}"))
    return m, [getattr(m.comps, f"r{i}") for i in range(len(multiples))]


# ══════════════════════════════════════════════════════════════════════
# The reductions: MEDIAN and the spreads
# ══════════════════════════════════════════════════════════════════════


class TestMedian:
    @pytest.mark.parametrize("values, middle", [
        (DATA, 30),
        ([40, 10, 30, 20], 25.0),              # an even count: the mean of the middle two
        ([7], 7),
        ([3, 3, 1], 3),
        ([-2.5, 1e-9, 4], 1e-9),
    ])
    def test_the_middle_of_a_row(self, values, middle):
        out = MEDIAN(_row(values))
        assert out.value == middle
        assert out.var_type == "scalar" and out.value_type == "float"

    def test_one_row_or_values_typed_in_are_one_number_two_rows_one_per_period(self):
        assert MEDIAN(_row([1, 9, 5, 7, 3])).value == 5
        assert MEDIAN(1, 2, 3, 4).value == 2.5
        assert MEDIAN(_row([1, 9]), _row([5, 7])).value == [3.0, 8.0]
        assert MEDIAN(_row([1, 9]), 100, _row([5, 7])).value == [5, 9]
        with pytest.raises(ValueError, match="different lengths"):
            MEDIAN(_row([1, 9]), _row([5, 7, 3]))

    def test_a_blank_text_and_a_boolean_in_a_line_are_skipped(self):
        row = _row([3, None, "x", True, 1, 4, 1.5])
        assert MEDIAN(row).value == 2.25
        assert MEDIAN(mo.Variable(True), 4).value == 4          # a cell holding TRUE
        assert MEDIAN(mo.Variable("2"), 4).value == 4           # a cell holding text "2"

    def test_values_typed_into_the_call_count_as_the_spreadsheet_counts_them(self):
        assert MEDIAN(1, "3").value == 2.0                       # numeric text
        assert MEDIAN(1, True, 5).value == 1                     # TRUE is 1
        assert MEDIAN(1, None, 5).value == 1                     # None is the 0 written
        assert MEDIAN(1, "abc").value == "#VALUE!"
        assert MEDIAN(1, "#N/A").value == "#VALUE!"              # typed, it is text

    def test_an_error_is_the_result_the_first_in_reading_order(self):
        assert MEDIAN(_row([1, "#DIV/0!", 3])).value == "#DIV/0!"
        na, div0 = mo.Variable("#N/A"), mo.Variable("#DIV/0!")
        assert MEDIAN([na, mo.Variable(1)], div0).value == "#N/A"
        assert MEDIAN("x", div0).value == "#VALUE!"
        assert MEDIAN(div0, "x").value == "#DIV/0!"
        assert MEDIAN(_row(["#tag", 4])).value == 4              # text that only starts with #
        # Over two rows, each period's first.
        assert MEDIAN(_row(["#N/A", 1]), _row(["#DIV/0!", 2])).value == ["#N/A", 1.5]

    def test_no_number_is_num(self):
        assert MEDIAN(_row(["a", None, "b"])).value == "#NUM!"
        assert MEDIAN(mo.Variable(None)).value == "#NUM!"

    def test_a_number_past_any_cell_is_num(self):
        assert MEDIAN(1e308, 1.7e308).value == "#NUM!"
        assert MEDIAN(_row([1.0, math.inf])).value == "#NUM!"

    def test_dates_give_a_date(self):
        d = _row([dt.date(2026, 1, 31), dt.date(2026, 3, 31), dt.date(2026, 2, 28)])
        out = MEDIAN(d)
        assert out.value == dt.date(2026, 2, 28) and out.value_type == "datetime"
        even = MEDIAN(_row([dt.date(2026, 2, 27), dt.date(2026, 2, 28)]))
        assert even.value == dt.datetime(2026, 2, 27, 12)
        assert MEDIAN(dt.date(2026, 1, 1), dt.date(2026, 1, 3)).value == dt.date(2026, 1, 2)

    def test_the_values_are_in_one_unit_and_the_middle_in_it(self):
        usd = _row([1.0, 2.0], unit="USD")
        assert str(MEDIAN(usd)._unit) == "USD"
        assert str(MEDIAN(usd, 5.0, _row([4.0]))._unit) == "USD"   # a number, an unlabelled line
        with pytest.raises(ValueError, match="mo.MEDIAN takes its values in one unit"):
            MEDIAN(usd, _row([3.0], unit="EUR"))
        with pytest.raises(ValueError, match="mo.MEDIAN takes its values in one unit"):
            MEDIAN([mo.Variable(1.0, unit="USD"), mo.Variable(2.0, unit="hr")])


class TestTheSpreads:
    def test_sample_and_population_over_a_row(self):
        row = _row()
        assert STDEV_S(row).value == pytest.approx(statistics.stdev(DATA), rel=1e-15)
        assert STDEV_P(row).value == pytest.approx(statistics.pstdev(DATA), rel=1e-15)
        assert VAR_S(row).value == 250.0
        assert VAR_P(row).value == 200.0

    def test_the_older_names_are_the_same_functions(self):
        row = _row([2.0, 4.0, 4.0, 4.0, 5.0, 5.0, 7.0, 9.0])
        assert STDEV(row).value == STDEV_S(row).value
        assert STDEVP(row).value == STDEV_P(row).value == 2.0
        assert VAR(row).value == VAR_S(row).value
        assert VARP(row).value == VAR_P(row).value == 4.0

    def test_two_passes_keep_large_numbers_exact(self):
        assert STDEV_S(1e15 + 1, 1e15 + 2, 1e15 + 3).value == 1.0
        assert VAR_P(_row([1e9 + 0.1, 1e9 + 0.3])).value == pytest.approx(0.01, rel=1e-6)

    @pytest.mark.parametrize("func, cells, value", [
        (STDEV_S, [7], "#DIV/0!"),               # a sample needs two numbers
        (VAR_S, [7, "x", None], "#DIV/0!"),
        (STDEV_P, [7], 0.0),                      # a population one
        (VAR_P, ["a", None], "#DIV/0!"),
        (STDEV_S, ["a"], "#DIV/0!"),
    ])
    def test_too_few_numbers_divide_by_zero(self, func, cells, value):
        assert func(_row(cells)).value == value

    def test_lines_skip_what_typed_values_count(self):
        assert STDEV_S(_row([True, 4, 6, "9", None])).value == pytest.approx(math.sqrt(2))
        assert STDEV_S(1, True, 5).value == pytest.approx(statistics.stdev([1, 1, 5]))
        assert STDEV_S(1, "3", 5).value == 2.0
        assert STDEV_S(1, "abc", 5).value == "#VALUE!"

    def test_an_error_is_the_result_the_first_in_reading_order(self):
        assert VAR_S(mo.Variable("#DIV/0!"), mo.Variable("#N/A")).value == "#DIV/0!"
        assert STDEV_P(_row([1.0, "#NUM!"])).value == "#NUM!"
        assert VAR_S(_row(["#DIV/0!", 1.0]), _row(["#N/A", 3.0])).value == ["#DIV/0!", 2.0]

    def test_too_large_a_spread_is_num(self):
        assert VAR_P(-1e200, 1e200).value == "#NUM!"

    def test_a_spread_of_dates_is_a_number_of_days(self):
        d = _row([dt.date(2026, 1, 31), dt.date(2026, 3, 31), dt.date(2026, 2, 28)])
        out = STDEV_S(d)
        assert out.value == pytest.approx(statistics.stdev([46053, 46112, 46081]))
        assert out.value_type == "float"

    def test_a_deviation_keeps_the_unit_and_a_variance_squares_it(self):
        usd = _row([1.0, 2.0, 4.0], unit="USD")
        assert str(STDEV_S(usd)._unit) == "USD"
        assert str(VAR_S(usd)._unit) == "USD^2"
        assert VAR_P(_row([0.1, 0.2], unit="%"))._unit is None    # a label squared is none
        assert VAR_S(_row([1.0, 2.0]))._unit is None
        with pytest.raises(ValueError, match="mo.STDEV_S takes its values in one unit"):
            STDEV_S(usd, _row([1.0], unit="hr"))


class TestTheReductionsArguments:
    def test_at_least_one_and_at_most_255(self):
        with pytest.raises(ValueError, match="at least one argument"):
            MEDIAN()
        assert MEDIAN(*range(255)).value == 127
        with pytest.raises(ValueError, match="at most 255 arguments"):
            STDEV_S(*range(256))

    def test_a_container_or_a_mapping_is_refused(self):
        m, recs = _comps(8.0, 9.0)
        with pytest.raises(TypeError, match="row, a list of lines"):
            MEDIAN(m.comps)
        with pytest.raises(TypeError, match="dict"):
            VAR_S({"a": 1})

    def test_a_line_with_tracks_is_refused(self):
        m = mo.Model("m", tracks=mo.Tracks("plan", "actual"), default_grain="month",
                     default_start="2026-01", default_periods=2)
        with m:
            m.x = mo.Variable(plan=[1.0, 2.0], actual=[3.0, 4.0])
        with pytest.raises(ValueError, match="slice a coordinate first"):
            MEDIAN(m.x)

    def test_a_list_holding_one_series_alone_is_refused(self):
        with pytest.raises(ValueError, match="two or more values"):
            MEDIAN([_row()])
        # With a value beside it, the list is one sample per period.
        assert MEDIAN([_row([1.0, 5.0]), mo.Variable(3.0)]).value == [2.0, 4.0]


# ══════════════════════════════════════════════════════════════════════
# One range: percentiles and quartiles
# ══════════════════════════════════════════════════════════════════════


class TestPercentile:
    @pytest.mark.parametrize("k, value", [
        (0, 10), (1, 50), (0.3, 22.0), (0.5, 30), (0.25, 20), (0.9, 46.0),
        (-0.1, "#NUM!"), (1.1, "#NUM!"),
    ])
    def test_inclusive(self, k, value):
        got = PERCENTILE_INC(_row(), k).value
        assert got == value if isinstance(value, str) else got == pytest.approx(value)

    @pytest.mark.parametrize("k, value", [
        (0.25, 15.0), (0.5, 30), (0.4, 24.0), (0.167, 10.02), (0.833, 49.98),
        (1 / 6, 10), (5 / 6, 50),                # the bounds: rank 1 and rank n
        (0.16, "#NUM!"), (0.84, "#NUM!"), (0, "#NUM!"), (1, "#NUM!"), (-1, "#NUM!"),
    ])
    def test_exclusive_between_the_first_and_the_last_rank(self, k, value):
        got = PERCENTILE_EXC(_row(), k).value
        assert got == value if isinstance(value, str) else got == pytest.approx(value)

    def test_the_cells_need_not_be_sorted(self):
        row = _row([40, 10, 30, 20, 50])
        assert PERCENTILE_INC(row, 0.3).value == 22.0
        assert PERCENTILE_EXC(row, 0.25).value == 15.0

    def test_a_position_within_15_digits_of_a_whole_one_is_that_cell(self):
        assert 100 * 0.29 == 28.999999999999996
        assert PERCENTILE_INC(_row(range(1, 102)), 0.29).value == 30
        assert QUARTILE_INC(_row(range(1, 6)), 2.9999999999999996).value == 4

    def test_one_number_is_every_percentile_of_it(self):
        assert PERCENTILE_INC(_row([7]), 0.3).value == 7
        assert PERCENTILE_EXC(_row([7]), 0.5).value == 7
        assert PERCENTILE_EXC(_row([7]), 0.25).value == "#NUM!"

    def test_k_is_read_as_the_spreadsheet_reads_a_number(self):
        row = _row()
        assert PERCENTILE_INC(row, "0.5").value == 30
        assert PERCENTILE_INC(row, True).value == 50
        assert PERCENTILE_INC(row, mo.Variable(None)).value == 10       # a blank cell is 0
        assert PERCENTILE_INC(row, mo.Variable("")).value == 10         # typed empty text too
        assert PERCENTILE_INC(row, mo.Variable("0.5")).value == 30
        assert PERCENTILE_INC(row, "abc").value == "#VALUE!"
        assert PERCENTILE_INC(row, mo.Variable("#DIV/0!")).value == "#DIV/0!"
        assert PERCENTILE_INC(row, "#DIV/0!").value == "#VALUE!"       # typed, it is text

    def test_k_is_read_before_the_range(self):
        bad = _row([1, "#N/A", 3])
        assert PERCENTILE_INC(bad, 0.5).value == "#N/A"
        assert PERCENTILE_INC(bad, 2).value == "#NUM!"
        assert PERCENTILE_INC(bad, mo.Variable("#DIV/0!")).value == "#DIV/0!"
        assert PERCENTILE_EXC(bad, 0).value == "#NUM!"

    def test_no_number_is_num(self):
        empty = _row(["a", None, "b", False])
        assert PERCENTILE_INC(empty, 0.5).value == "#NUM!"
        assert PERCENTILE_EXC(empty, 0.5).value == "#NUM!"

    def test_a_series_k_is_one_call_per_period(self):
        out = PERCENTILE_INC(_row(), mo.Variable([0, 0.5, 1, 2]))
        assert out.value == [10, 30, 50, "#NUM!"]
        assert out.var_type == "list"

    def test_dates_give_a_date(self):
        d = _row([dt.date(2026, 1, 31), dt.date(2026, 3, 31), dt.date(2026, 2, 28)])
        out = PERCENTILE_INC(d, 0.25)
        assert out.value == dt.date(2026, 2, 14) and out.value_type == "datetime"

    def test_the_unit_is_the_range_s_never_k_s(self):
        out = PERCENTILE_INC(_row([1.0, 2.0], unit="USD"), mo.Variable(0.5, unit="%"))
        assert str(out._unit) == "USD"
        with pytest.raises(ValueError, match="mo.PERCENTILE_INC takes its values in one unit"):
            PERCENTILE_INC([mo.Variable(1.0, unit="USD"), mo.Variable(2.0, unit="EUR")], 0.5)

    def test_the_older_name_is_the_inclusive_one(self):
        assert PERCENTILE(_row(), 0.3).value == PERCENTILE_INC(_row(), 0.3).value
        assert PERCENTILE(_row(), 1.5).value == "#NUM!"


class TestQuartile:
    @pytest.mark.parametrize("quart, value", [
        (0, 10), (1, 20), (2, 30), (3, 40), (4, 50),
        (1.5, 20), (4.5, 50), (-0.5, 10),          # a fraction is truncated
        (5, "#NUM!"), (-1, "#NUM!"),
    ])
    def test_inclusive(self, quart, value):
        assert QUARTILE_INC(_row(), quart).value == value

    @pytest.mark.parametrize("quart, value", [
        (1, 15.0), (2, 30), (3, 45.0), (3.5, 45.0),
        (0, "#NUM!"), (4, "#NUM!"), (0.5, "#NUM!"),
    ])
    def test_exclusive(self, quart, value):
        assert QUARTILE_EXC(_row(), quart).value == value

    def test_too_few_cells_hold_an_exclusive_quartile(self):
        assert QUARTILE_EXC(_row([7]), 1).value == "#NUM!"
        assert QUARTILE_EXC(_row([7]), 2).value == 7
        assert QUARTILE_EXC(_row([10, 20]), 2).value == 15.0
        assert QUARTILE_EXC(_row([10, 20]), 1).value == "#NUM!"

    def test_the_second_quartile_is_the_median(self):
        row = _row([0.1, 0.7, 0.2, 0.9])
        assert QUARTILE_INC(row, 2).value == MEDIAN(row).value
        assert QUARTILE_EXC(row, 2).value == MEDIAN(row).value

    def test_no_number_and_an_error(self):
        assert QUARTILE_INC(_row(["a"]), 1).value == "#NUM!"
        assert QUARTILE_INC(_row([1, "#N/A"]), mo.Variable("#DIV/0!")).value == "#DIV/0!"
        assert QUARTILE_INC(_row([1, "#N/A"]), 9).value == "#NUM!"
        assert QUARTILE_INC(_row([1, "#N/A"]), 1).value == "#N/A"

    def test_the_older_name_is_the_inclusive_one(self):
        assert QUARTILE(_row(), 1).value == 20 and QUARTILE(_row(), 0).value == 10


# ══════════════════════════════════════════════════════════════════════
# One range: the k-th largest and smallest
# ══════════════════════════════════════════════════════════════════════


class TestLargeAndSmall:
    @pytest.mark.parametrize("k, large, small", [
        (1, 50, 10), (2, 40, 20), (5, 10, 50),
        (2.5, 30, 20),                              # LARGE takes k up, SMALL down
        (2.9, 30, 20),
        (0.5, 50, "#NUM!"),
        (5.5, "#NUM!", 50),
        (2.0000000000000004, 40, 20),               # within 15 digits of 2
        (1.9999999999999998, 40, 20),
        (2.000000000001, 30, 20),
        (0, "#NUM!", "#NUM!"), (6, "#NUM!", "#NUM!"), (-1, "#NUM!", "#NUM!"),
    ])
    def test_the_k_th(self, k, large, small):
        row = _row([40, 10, 30, 20, 50])
        assert LARGE(row, k).value == large
        assert SMALL(row, k).value == small

    def test_ties_each_count(self):
        assert LARGE(_row(TIES), 2).value == 5 and LARGE(_row(TIES), 4).value == 3
        assert SMALL(_row(TIES), 2).value == 3

    def test_k_is_read_as_the_spreadsheet_reads_a_number(self):
        assert LARGE(_row(), "2").value == 40
        assert LARGE(_row(), mo.Variable("2")).value == 40      # a cell holding text "2"
        assert SMALL(_row(), mo.Variable("2")).value == 20
        assert LARGE(_row(), True).value == 50
        assert LARGE(_row(), mo.Variable(None)).value == "#NUM!"
        assert LARGE(_row(), "two").value == "#VALUE!"

    def test_k_s_error_then_the_range_s_then_k_against_the_count(self):
        bad = _row([1, "#N/A"])
        assert LARGE(bad, mo.Variable("#DIV/0!")).value == "#DIV/0!"
        assert LARGE(bad, 0).value == "#N/A"
        assert LARGE(bad, 9).value == "#N/A"
        assert SMALL(_row(["a"]), 1).value == "#NUM!"

    def test_the_cell_comes_back_as_it_holds_it(self):
        out = LARGE(_row([3, 1, 2]), 1)
        assert out.value == 3 and isinstance(out.value, int) and out.value_type == "int"
        d = _row([dt.date(2026, 1, 31), dt.date(2026, 3, 31)])
        assert SMALL(d, 1).value == dt.date(2026, 1, 31)
        assert SMALL(d, 1).value_type == "datetime"
        assert str(LARGE(_row([1.0, 2.0], unit="USD"), 1)._unit) == "USD"

    def test_a_series_k_is_one_call_per_period(self):
        assert SMALL(_row(), [1, 2, 3]).value == [10, 20, 30]


# ══════════════════════════════════════════════════════════════════════
# One range: ranks
# ══════════════════════════════════════════════════════════════════════


class TestRank:
    @pytest.mark.parametrize("number, order, eq, avg", [
        (3, 0, 4, 4.5), (3, 1, 2, 2.5),
        (5, 0, 1, 2.0), (5, 1, 4, 5.0),
        (1, 0, 6, 6.0), (1, 1, 1, 1.0),
    ])
    def test_ties_share_the_best_rank_or_their_mean(self, number, order, eq, avg):
        ref = _row(TIES)
        assert RANK_EQ(number, ref, order).value == eq
        assert RANK_AVG(number, ref, order).value == avg

    def test_descending_unless_the_order_is_not_zero(self):
        ref = _row(TIES)
        assert RANK_EQ(3, ref).value == 4
        for order in (True, -1, 0.5, "1", mo.Variable(2)):
            assert RANK_EQ(3, ref, order).value == 2, order
        assert RANK_EQ(3, ref, mo.Variable(None)).value == 4      # a blank cell is 0

    def test_a_number_the_range_lacks_is_na(self):
        assert RANK_EQ(4, _row(TIES)).value == "#N/A"
        assert RANK_EQ(9, _row(TIES)).value == "#N/A"
        assert RANK_EQ(0.3, [mo.Variable(0.1) + mo.Variable(0.2), mo.Variable(0.5)]
                       ).value == "#N/A"                          # equal exactly, or not
        assert RANK_EQ(1, _row(["a", None, "b"])).value == "#N/A"

    def test_the_number_is_read_as_the_spreadsheet_reads_one(self):
        ref = _row(TIES)
        assert RANK_EQ("3", ref).value == 4
        assert RANK_EQ(True, ref).value == 6
        assert RANK_EQ(mo.Variable(None), ref).value == "#N/A"   # a blank is 0
        assert RANK_EQ("abc", ref).value == "#VALUE!"

    def test_the_order_the_range_then_the_number(self):
        div0, na = mo.Variable("#DIV/0!"), mo.Variable("#N/A")
        assert RANK_EQ(div0, _row(DATA)).value == "#DIV/0!"
        assert RANK_EQ(10, _row(DATA), div0).value == "#DIV/0!"
        assert RANK_EQ(div0, _row([1, "#N/A"])).value == "#N/A"
        assert RANK_EQ(na, _row(DATA), div0).value == "#DIV/0!"
        assert RANK_EQ(div0, _row(["a"])).value == "#N/A"

    def test_a_blank_text_and_a_boolean_in_the_range_are_skipped(self):
        assert RANK_EQ(1, _row([3, None, "x", True, 1, 4, 1.5])).value == 4

    def test_dates_rank_by_their_serial(self):
        d = _row([dt.date(2026, 1, 31), dt.date(2026, 3, 31), dt.date(2026, 2, 28)])
        assert RANK_EQ(dt.date(2026, 2, 28), d).value == 2
        assert RANK_EQ(46081, d).value == 2

    def test_a_rank_is_an_int_and_an_average_rank_a_float_without_a_unit(self):
        ref = _row([1.0, 2.0], unit="USD")
        eq, avg = RANK_EQ(mo.Variable(2.0, unit="USD"), ref), RANK_AVG(2.0, ref)
        assert (eq.value, eq.value_type, eq._unit) == (1, "int", None)
        assert (avg.value, avg.value_type, avg._unit) == (1.0, "float", None)

    def test_the_number_and_the_range_are_in_one_unit_the_order_is_no_value(self):
        ref = _row([1.0, 2.0], unit="USD")
        assert RANK_EQ(2.0, ref, mo.Variable(1, unit="flag")).value == 2
        with pytest.raises(ValueError, match="mo.RANK_EQ takes its values in one unit"):
            RANK_EQ(mo.Variable(2.0, unit="hr"), ref)

    def test_a_series_number_ranks_each_period_among_the_row(self):
        row = _row([30, 10, 20])
        out = RANK_EQ(row, row)
        assert out.value == [1, 3, 2] and out.var_type == "list"

    def test_each_record_among_its_peers(self):
        m, recs = _comps(8.0, 12.0, 10.0, 9.0, 10.0)
        peers = [r.multiple for r in recs]
        assert [RANK_EQ(r.multiple, peers).value for r in recs] == [5, 1, 2, 4, 2]
        assert [RANK_AVG(r.multiple, peers).value for r in recs] == [5.0, 1.0, 2.5, 4.0, 2.5]
        assert [RANK_EQ(r.multiple, peers, 1).value for r in recs] == [1, 5, 3, 2, 3]

    def test_the_older_name_is_rank_eq(self):
        assert RANK(3, _row(TIES)).value == RANK_EQ(3, _row(TIES)).value == 4


# ══════════════════════════════════════════════════════════════════════
# The two range kinds and the parameters
# ══════════════════════════════════════════════════════════════════════


class TestTheRangeKinds:
    def test_one_field_down_the_records_reads_as_a_row_does(self):
        m, recs = _comps(8.0, 12.0, 10.0, 9.0)
        field = [r.multiple for r in recs]
        row = _row([8.0, 12.0, 10.0, 9.0])
        for func, args in ((MEDIAN, ()), (STDEV_S, ()), (VAR_P, ())):
            assert func(field, *args).value == func(row, *args).value, func
        for func, k in ((PERCENTILE_INC, 0.25), (PERCENTILE_EXC, 0.4), (QUARTILE_INC, 3),
                        (QUARTILE_EXC, 1), (LARGE, 2), (SMALL, 2)):
            assert func(field, k).value == func(row, k).value, func
        assert str(MEDIAN(field)._unit) == "x"

    def test_a_tuple_one_element_of_a_row_and_typed_values(self):
        row = _row([5.0, 1.0, 3.0])
        assert MEDIAN((row[0], row[1], row[2])).value == 3.0
        assert LARGE([row[0], 7.0, "x", True], 2).value == 5.0    # a list is an array
        assert MEDIAN([1, True, "3", None]).value == 1             # numbers only

    def test_a_range_is_a_row_or_a_list_never_one_typed_value(self):
        with pytest.raises(TypeError, match="a range is"):
            LARGE(5, 1)
        with pytest.raises(ValueError, match="at least one cell"):
            PERCENTILE_INC([], 0.5)

    def test_a_parameter_is_one_value_or_a_line(self):
        with pytest.raises(TypeError, match="k is one value"):
            LARGE(_row(), (1, 2))
        with pytest.raises(TypeError, match="number is one value"):
            RANK_EQ([mo.Variable(1.0)], _row())

    def test_a_parameter_with_tracks_is_refused(self):
        m = mo.Model("m", tracks=mo.Tracks("plan", "actual"), default_grain="month",
                     default_start="2026-01", default_periods=2)
        with m:
            m.k = mo.Variable(plan=[0.1, 0.2], actual=[0.3, 0.4])
        with pytest.raises(ValueError, match="slice a coordinate first"):
            PERCENTILE_INC(_row(), m.k)

    def test_ranges_and_parameters_of_different_lengths_are_fine(self):
        # One range, read whole by every period's call.
        assert LARGE(_row([1, 2, 3, 4, 5, 6]), _row([1, 2])).value == [6, 5]


# ══════════════════════════════════════════════════════════════════════
# The call: the AST, the registry, the loop
# ══════════════════════════════════════════════════════════════════════

ALL = {
    "MEDIAN": MEDIAN, "STDEV_S": STDEV_S, "STDEV_P": STDEV_P, "VAR_S": VAR_S,
    "VAR_P": VAR_P, "STDEV": STDEV, "STDEVP": STDEVP, "VAR": VAR, "VARP": VARP,
    "PERCENTILE_INC": PERCENTILE_INC, "PERCENTILE_EXC": PERCENTILE_EXC,
    "QUARTILE_INC": QUARTILE_INC, "QUARTILE_EXC": QUARTILE_EXC,
    "PERCENTILE": PERCENTILE, "QUARTILE": QUARTILE, "LARGE": LARGE, "SMALL": SMALL,
    "RANK_EQ": RANK_EQ, "RANK_AVG": RANK_AVG, "RANK": RANK,
}
REDUCTIONS = {"MEDIAN", "STDEV_S", "STDEV_P", "VAR_S", "VAR_P", "STDEV", "STDEVP", "VAR", "VARP"}
EXCEL = {
    "STDEV_S": "STDEV.S", "STDEV_P": "STDEV.P", "VAR_S": "VAR.S", "VAR_P": "VAR.P",
    "PERCENTILE_INC": "PERCENTILE.INC", "PERCENTILE_EXC": "PERCENTILE.EXC",
    "QUARTILE_INC": "QUARTILE.INC", "QUARTILE_EXC": "QUARTILE.EXC",
    "RANK_EQ": "RANK.EQ", "RANK_AVG": "RANK.AVG",
}


def _call(name):
    row = mo.Variable([10.0, 20.0, 30.0], display_name="row")
    if name in REDUCTIONS:
        return ALL[name](row, 5)
    if name.startswith("RANK"):
        return ALL[name](mo.Variable(20.0, display_name="x"), row, 1)
    return ALL[name](row, mo.Variable(1, display_name="k"))


class TestTheCall:
    @pytest.mark.parametrize("name", sorted(ALL))
    def test_the_call_is_named_as_the_function_and_holds_its_arguments(self, name):
        out = _call(name)
        assert isinstance(out._expr, FuncCall) and out._expr.func == name
        assert mo.to_json(out._expr)["func"] == name
        args = out._expr.args
        if name in REDUCTIONS:
            assert isinstance(args[0], VarRef) and args[1] == Literal(5)
            assert out.formula == f"{name}(row, 5)"
        elif name.startswith("RANK"):
            assert [type(a) for a in args] == [VarRef, VarRef, Literal]
            assert out.formula == f"{name}(x, row, 1)"
        else:
            assert [type(a) for a in args] == [VarRef, VarRef]
            assert out.formula == f"{name}(row, k)"

    def test_a_list_of_lines_is_one_list_node(self):
        m, recs = _comps(8.0, 12.0)
        out = PERCENTILE_INC([r.multiple for r in recs], 0.5)
        assert isinstance(out._expr.args[0], ListExpr)
        assert out._expr.args[0].items == [VarRef(r.multiple) for r in recs]
        assert out._source_code == "PERCENTILE_INC([r0.multiple, r1.multiple], 0.5)"

    def test_an_omitted_order_is_left_out_of_the_call(self):
        out = RANK_EQ(20.0, _row([10.0, 20.0]))
        assert len(out._expr.args) == 2

    @pytest.mark.parametrize("name", sorted(ALL))
    def test_each_name_is_registered_with_its_shape_and_excel_name(self, name):
        s = spec(name)
        assert s.shape == ("reduction" if name in REDUCTIONS else "range"), name
        assert s.range_args == (None if name in REDUCTIONS
                                else (1,) if name.startswith("RANK") else (0,)), name
        assert excel_name(name) == EXCEL.get(name, name)
        assert file_prefix(name) == ("_xlfn." if name in EXCEL else "")
        assert is_passthrough(name) and not s.bare
        assert range_positions(name) == s.range_args
        # Two values or more, a series among them: one call per period.
        assert per_period_arity(name) == s.per_period_arity == 2
        assert not admits_in_loop(name, 1) and admits_in_loop(name, 2)

    @pytest.mark.parametrize("make", [
        lambda d: MEDIAN(d.bal),
        lambda d: STDEV_S(d.bal),
        lambda d: LARGE(d.bal, 1),
        lambda d: RANK_EQ(d.bal, d.bal),
    ])
    def test_a_loop_container_refuses_a_statistic_over_one_of_its_rows(self, make):
        m = mo.Model("m", default_grain="year", default_start="2026", default_periods=3)
        with pytest.raises(ValueError, match="reads a whole series"):
            with m:
                m.d = mo.MultiVariable("D", loop=True)
                with m.d as d:
                    d.bal = mo.lag(d.x, fill=10.0) * 1.37
                    d.stat = make(d)
                    d.x = d.bal - d.stat
