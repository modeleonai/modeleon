# SPDX-License-Identifier: Apache-2.0
"""Tests for helper functions: recurrence, SUM, cumsum, IF, MAX/MIN/AVERAGE."""

import pytest

import modeleon as mo


class TestRecurrence:
    def test_recurrence_sum_with_scalar_increment(self):
        # period 0 = start (0), then +10 each step
        result = mo.recurrence_sum(start=0, increment=10, periods=5)
        assert result._value == [0, 10, 20, 30, 40]

    def test_recurrence_template_string(self):
        # period 0 = 100 (start), then *1.1 each step
        result = mo.recurrence(100, "{prev} * 1.1", periods=3)
        assert len(result._value) == 3
        assert result._value[0] == 100
        assert result._value[2] == pytest.approx(100 * 1.1 * 1.1)


class TestSum:
    def test_sum_of_list_variable(self):
        v = mo.Variable([10, 20, 30])
        result = mo.SUM(v)
        assert result._value == 60


class TestIF:
    def test_if_scalar_true_branch(self):
        result = mo.IF(mo.Variable(1) > mo.Variable(0), 100, 200)
        assert result._value == 100

    def test_if_scalar_false_branch(self):
        result = mo.IF(mo.Variable(1) < mo.Variable(0), 100, 200)
        assert result._value == 200


class TestMinMax:
    def test_min(self):
        v = mo.Variable([5, 2, 8, 1, 9])
        assert mo.MIN(v)._value == 1

    def test_max(self):
        v = mo.Variable([5, 2, 8, 1, 9])
        assert mo.MAX(v)._value == 9

    def test_average(self):
        v = mo.Variable([10, 20, 30])
        assert mo.AVERAGE(v)._value == 20


class TestVariadic:
    """SUM/MAX/MIN/AVERAGE all accept multiple args, a single list Variable,
    a plain list, or all scalars — matching Excel semantics exactly."""

    def test_sum_multiple_args(self):
        result = mo.SUM(mo.Variable(10), mo.Variable(20), mo.Variable(30))
        assert result._value == 60
        assert result.formula.startswith("SUM(")

    def test_sum_single_list_variable(self):
        assert mo.SUM(mo.Variable([10, 20, 30]))._value == 60

    def test_sum_mixed_variable_and_scalar(self):
        result = mo.SUM(mo.Variable([10, 20]), 100)
        assert result._value == 130

    def test_sum_all_scalars_returns_variable(self):
        # =SUM(1, 2, 3) is a valid Excel formula — so we return a Variable,
        # not a plain int. The formula is the product.
        result = mo.SUM(1, 2, 3)
        assert isinstance(result, mo.Variable)
        assert result._value == 6
        assert result.formula == "SUM(1, 2, 3)"

    def test_max_multiple_scalars(self):
        result = mo.MAX(mo.Variable(5), mo.Variable(9), mo.Variable(7))
        assert result._value == 9

    def test_max_all_literal_returns_variable(self):
        result = mo.MAX(5, 10, 3)
        assert isinstance(result, mo.Variable)
        assert result._value == 10
        assert result.formula == "MAX(5, 10, 3)"

    def test_average_variadic(self):
        result = mo.AVERAGE(mo.Variable(10), mo.Variable(20), mo.Variable(30))
        assert result._value == 20

    def test_average_single_list(self):
        assert mo.AVERAGE(mo.Variable([10, 20, 30]))._value == 20


class TestAlwaysVariable:
    """The library's one rule: every mo.X(...) returns a Variable with a formula."""

    def test_abs_scalar_returns_variable(self):
        result = mo.ABS(-5)
        assert isinstance(result, mo.Variable)
        assert result._value == 5
        assert result.formula == "ABS(-5)"

    def test_round_scalar_returns_variable(self):
        result = mo.ROUND(3.14159, 2)
        assert isinstance(result, mo.Variable)
        assert result._value == 3.14
        assert result.formula == "ROUND(3.14159, 2)"

    def test_int_scalar_returns_variable(self):
        result = mo.INT(-2.5)
        assert isinstance(result, mo.Variable)
        assert result._value == -3  # Excel INT floors toward -inf

    def test_upper_plain_string_returns_variable(self):
        result = mo.UPPER("hello")
        assert isinstance(result, mo.Variable)
        assert result._value == "HELLO"
        assert result.formula == "UPPER('hello')"

    def test_year_plain_date_returns_variable(self):
        from datetime import date
        result = mo.YEAR(date(2025, 3, 14))
        assert isinstance(result, mo.Variable)
        assert result._value == 2025

    def test_concat_plain_strings_returns_variable(self):
        result = mo.CONCAT("Q1 ", "Report")
        assert isinstance(result, mo.Variable)
        assert result._value == "Q1 Report"
        assert result.formula.startswith("CONCAT(")


class TestErrorMessages:
    """Error messages should teach via example, not just describe the failure."""

    def test_sum_empty_mentions_syntax(self):
        with pytest.raises(ValueError, match=r"Syntax: mo\.SUM"):
            mo.SUM()

    def test_max_empty_mentions_syntax(self):
        with pytest.raises(ValueError, match=r"Syntax: mo\.MAX"):
            mo.MAX()

    def test_concat_empty_mentions_syntax(self):
        with pytest.raises(ValueError, match=r"Syntax: mo\.CONCAT"):
            mo.CONCAT()

    def test_cumsum_scalar_suggests_numpy(self):
        # cumsum is a domain helper with no Excel equivalent — stays strict.
        with pytest.raises(TypeError, match=r"numpy\.cumsum"):
            mo.cumsum([1, 2, 3])


class TestRoundIntMod:
    def test_round_scalar(self):
        v = mo.Variable(3.14159)
        assert mo.ROUND(v, 2)._value == 3.14

    def test_round_list(self):
        v = mo.Variable([1.234, 5.678, 9.0])
        assert mo.ROUND(v, 1)._value == [1.2, 5.7, 9.0]

    def test_int_floors_negative(self):
        assert mo.INT(mo.Variable(-2.5))._value == -3

    def test_mod_variable(self):
        assert mo.MOD(mo.Variable(10), 3)._value == 1

    def test_abs_list(self):
        v = mo.Variable([-10, 20, -30])
        assert mo.ABS(v)._value == [10, 20, 30]
