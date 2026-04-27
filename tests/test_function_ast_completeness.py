# SPDX-License-Identifier: Apache-2.0
"""Every DSL function must produce an ``_expr`` AST node, not just a value.

If a function returned a Variable with ``_value`` set but ``_expr is None``,
renderers would see a dead-end — no formula to translate, no dependencies
to trace. Same logic applies to future backends (pandas, SQL): each walks
the AST and needs every function call represented as a ``FuncCall`` (or
equivalent) node.

This test pins the invariant across every public function in the
``functions/`` package. See ADR-010.
"""

from __future__ import annotations

import modeleon as mo


class TestFunctionASTCompleteness:
    """Call each public function with Variable inputs; assert the result
    carries a non-``None`` ``_expr``."""

    def test_aggregate_functions(self):
        a = mo.Variable(10)
        b = mo.Variable(20)
        c = mo.Variable(30)
        assert mo.SUM(a, b, c)._expr is not None
        assert mo.MAX(a, b, c)._expr is not None
        assert mo.MIN(a, b, c)._expr is not None
        assert mo.AVERAGE(a, b, c)._expr is not None

    def test_math_functions(self):
        a = mo.Variable(-10)
        b = mo.Variable(3.7)
        c = mo.Variable(100)
        divisor = mo.Variable(7)
        assert mo.ABS(a)._expr is not None
        assert mo.ROUND(b, 1)._expr is not None
        assert mo.INT(b)._expr is not None
        assert mo.MOD(c, divisor)._expr is not None

    def test_conditional_function(self):
        cond = mo.Variable(True)
        a = mo.Variable(1)
        b = mo.Variable(2)
        assert mo.IF(cond, a, b)._expr is not None

    def test_date_functions(self):
        import datetime as dt
        d = mo.Variable(dt.date(2026, 4, 24))
        assert mo.YEAR(d)._expr is not None
        assert mo.MONTH(d)._expr is not None
        assert mo.DAY(d)._expr is not None
        assert mo.EDATE(d, 3)._expr is not None
        assert mo.EOMONTH(d, 0)._expr is not None

    def test_text_functions(self):
        s = mo.Variable("hello")
        t = mo.Variable("world")
        assert mo.LEN(s)._expr is not None
        assert mo.UPPER(s)._expr is not None
        assert mo.LOWER(s)._expr is not None
        assert mo.CONCAT(s, t)._expr is not None

    def test_financial_functions(self):
        cf = mo.Variable([-1000.0, 300.0, 400.0, 500.0])
        rate = mo.Variable(0.1)
        nper = mo.Variable(12)
        pv = mo.Variable(1000)
        pmt = mo.Variable(100)
        assert mo.IRR(cf)._expr is not None
        assert mo.NPV(rate, cf)._expr is not None
        assert mo.PMT(rate, nper, pv)._expr is not None
        assert mo.FV(rate, nper, pmt)._expr is not None
        assert mo.PV(rate, nper, pmt)._expr is not None

    def test_recurrence_functions(self):
        series = mo.Variable([1, 2, 3, 4, 5])
        assert mo.cumsum(series)._expr is not None

    def test_self_ref_produces_ast(self):
        growth = mo.Variable(0.05)
        base = mo.Variable(100)
        result = mo.recurrence(
            start=base,
            formula="{prev} * (1 + {g})",
            g=growth,
            periods=3,
        )
        assert result._expr is not None
