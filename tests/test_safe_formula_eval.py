# SPDX-License-Identifier: Apache-2.0
"""Guardrails for the template-formula evaluator in :mod:`modeleon.functions.recurrence`.

Formula templates like ``"{prev} * (1 - {churn})"`` are parsed with
``ast.parse(mode='eval')`` and walked by :func:`_safe_eval`, which only
honours numeric literals, a fixed set of arithmetic / comparison /
unary / boolean operators, ternary ``x if c else y``, and calls to a
small whitelist of functions. This test suite pins that contract — any
Python construct outside the whitelist must raise, and every legitimate
pattern we support must continue to work.
"""

from __future__ import annotations

import pytest

import modeleon as mo
from modeleon.functions.recurrence import _eval_template


# ─── Legitimate patterns keep working ──────────────────────────────────


class TestAllowedPatterns:
    def test_simple_arithmetic(self):
        # period 0 = start (100), then formula applied to grow each period
        r = _eval_template("{prev} * 1.05", start_value=100, periods=3,
                           var_values_per_period={})
        assert r == [100, 105.0, pytest.approx(110.25)]

    def test_named_variable_substitution(self):
        # period 0 = 1000, period 1 = 1000 * (1-churn[1]) = 1000 * 0.8 = 800
        r = _eval_template("{prev} * (1 - {churn})", start_value=1000, periods=2,
                           var_values_per_period={"churn": [0.1, 0.2]})
        assert r == [1000, 800.0]

    def test_function_call(self):
        r = _eval_template("MAX({prev} * 0.9, 50)", start_value=100, periods=3,
                           var_values_per_period={})
        assert r == [100, 90.0, 81.0]

    def test_nested_function_calls(self):
        r = _eval_template("IF({prev} > 0, ABS({prev}), 0)", start_value=5,
                           periods=3, var_values_per_period={})
        assert r == [5, 5, 5]

    def test_ternary_conditional(self):
        r = _eval_template("{prev} + 1 if {prev} < 3 else 0", start_value=0,
                           periods=5, var_values_per_period={})
        # period 0 = 0; then 0+1=1, 1+1=2, 2+1=3, 3 not <3 so 0
        assert r == [0, 1, 2, 3, 0]

    def test_compound_comparisons(self):
        # periods=1 → just [start], formula not evaluated
        r = _eval_template("IF(0 < {prev} <= 10, {prev} * 2, 0)", start_value=5,
                           periods=2, var_values_per_period={})
        assert r == [5, 10]

    def test_boolean_operators(self):
        r = _eval_template("IF({prev} > 0 and {prev} < 100, {prev}, 0)",
                           start_value=50, periods=2, var_values_per_period={})
        assert r == [50, 50]


class TestDateFunctions:
    """Pure date arithmetic is whitelisted (EDATE/EOMONTH/YEAR/MONTH/DAY)."""

    def test_edate_chains_per_period(self):
        from datetime import date
        # period 0 = start; periods 1+ apply EDATE
        r = _eval_template("EDATE({prev}, 1)", start_value=date(2025, 1, 1),
                           periods=3, var_values_per_period={})
        assert r == [date(2025, 1, 1), date(2025, 2, 1), date(2025, 3, 1)]

    def test_eomonth_returns_last_of_month(self):
        from datetime import date
        r = _eval_template("EOMONTH({prev}, 0)", start_value=date(2025, 1, 15),
                           periods=2, var_values_per_period={})
        assert r == [date(2025, 1, 15), date(2025, 1, 31)]

    def test_year_month_day_extract(self):
        from datetime import date
        # period 0 = start (date), period 1 = formula(start)
        r_y = _eval_template("YEAR({prev})", start_value=date(2025, 6, 15),
                             periods=2, var_values_per_period={})
        r_m = _eval_template("MONTH({prev})", start_value=date(2025, 6, 15),
                             periods=2, var_values_per_period={})
        r_d = _eval_template("DAY({prev})", start_value=date(2025, 6, 15),
                             periods=2, var_values_per_period={})
        assert r_y[1] == 2025
        assert r_m[1] == 6
        assert r_d[1] == 15

    def test_self_ref_emits_edate_chain_in_excel(self, tmp_path):
        from datetime import date
        m = mo.MultiVariable("M")
        m.s = mo.MultiVariable("S")
        m.s.start_date = mo.Variable(date(2025, 1, 1))
        m.s.month = mo.recurrence(
            start=m.s.start_date,
            formula="EDATE({prev}, 1)",
            periods=6,
        )

        m.to_excel(tmp_path / "edate_chain.xlsx")

        from openpyxl import load_workbook
        ws = load_workbook(tmp_path / "edate_chain.xlsx")["S"]
        formulas = [c.value for row in ws.iter_rows() for c in row
                    if isinstance(c.value, str) and c.value.startswith("=EDATE(")]
        # period 0 = start (literal/ref, no EDATE); periods 1..5 emit EDATE
        assert len(formulas) == 5


# ─── Rejection surface — every attack pattern raises ───────────────────


def _eval_raises(formula: str, *, start=1) -> Exception:
    """Run _eval_template and return the raised exception for inspection.

    Uses ``periods=2`` so the formula actually evaluates at period 1
    (under new semantics period 0 is just the start value verbatim, no
    template expansion happens at t=0).
    """
    with pytest.raises(ValueError) as exc_info:
        _eval_template(formula, start_value=start, periods=2,
                       var_values_per_period={
                           k: [None, None] for k in
                           _placeholders_in(formula)
                       })
    return exc_info.value


def _placeholders_in(formula: str) -> list:
    import re
    return re.findall(r"\{([A-Za-z_][A-Za-z0-9_]*)\}", formula)


class TestRejectedPatterns:
    def test_attribute_access_rejected(self):
        # The classic restricted-eval escape vector.
        exc = _eval_raises("(0).__class__")
        assert "not allowed" in str(exc).lower() or "Attribute" in str(exc)

    def test_subscripting_rejected(self):
        exc = _eval_raises("[1, 2, 3][0]")
        assert "not allowed" in str(exc).lower() or "List" in str(exc) or "Subscript" in str(exc)

    def test_import_via_builtins_rejected(self):
        exc = _eval_raises("__import__('os')")
        assert "not allowed" in str(exc).lower() or "__import__" in str(exc)

    def test_unknown_function_rejected(self):
        exc = _eval_raises("exec('print(1)')")
        assert "not allowed" in str(exc).lower() or "exec" in str(exc)

    def test_lambda_rejected(self):
        exc = _eval_raises("(lambda: 1)()")
        # Rejected as "only direct function calls allowed" since the
        # call target isn't an ``ast.Name``.
        msg = str(exc).lower()
        assert "direct function calls" in msg or "not allowed" in msg

    def test_list_comprehension_rejected(self):
        exc = _eval_raises("[x for x in range(10)][0]")
        # Comprehension is rejected; exact wording may vary.
        assert "not allowed" in str(exc).lower() or "ListComp" in str(exc)

    def test_generator_expression_rejected(self):
        exc = _eval_raises("sum(x for x in range(10))")
        assert "not allowed" in str(exc).lower() or "Generator" in str(exc) or "sum" in str(exc).lower()

    def test_walrus_rejected(self):
        exc = _eval_raises("(a := 5)")
        assert "not allowed" in str(exc).lower() or "Walrus" in str(exc) or "Named" in str(exc)

    def test_string_literal_rejected(self):
        exc = _eval_raises("'hello'")
        assert "numeric" in str(exc).lower() or "literal" in str(exc).lower()

    def test_keyword_args_rejected(self):
        exc = _eval_raises("MAX({prev}, key=1)")
        assert "keyword" in str(exc).lower()

    def test_attribute_access_on_known_name_rejected(self):
        exc = _eval_raises("{prev}.real")
        assert "not allowed" in str(exc).lower() or "Attribute" in str(exc)


# ─── End-to-end: self_ref exposes the hardening ────────────────────────


class TestSelfRefIntegration:
    def test_malicious_formula_raises_in_recurrence(self):
        base = mo.Variable(100)
        with pytest.raises(ValueError):
            mo.recurrence(
                start=base,
                formula="{prev}.__class__.__bases__[0].__subclasses__()[0]",
                periods=3,
            )

    def test_legitimate_self_ref_still_works(self):
        base = mo.Variable(100)
        churn = mo.Variable(0.05)
        result = mo.recurrence(
            start=base,
            formula="{prev} * (1 - {churn})",
            churn=churn,
            periods=3,
        )
        # period 0 = start, periods 1+ apply churn
        assert len(result._value) == 3
        assert result._value[0] == 100
        assert result._value[1] == pytest.approx(95.0)
        assert result._value[2] == pytest.approx(90.25)
