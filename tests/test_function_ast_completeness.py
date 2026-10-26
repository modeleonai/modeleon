# SPDX-License-Identifier: Apache-2.0
"""Every DSL function must produce an ``_expr`` AST node, not just a value.

If a function returned a Variable with ``_value`` set but ``_expr is None``,
renderers would see a dead-end — no formula to translate, no dependencies
to trace. Same logic applies to future backends (pandas, SQL): each walks
the AST and needs every function call represented as a ``FuncCall`` (or
equivalent) node.

This test pins the invariant across every public function in the
``functions/`` package; the last test fails when a function is exported
without a call here.
"""

from __future__ import annotations

import datetime as dt
import re
from pathlib import Path

import modeleon as mo
from modeleon import functions
from modeleon.core.expr import FuncCall


def _calls(results):
    """Each result carries the call itself as its AST root, named as the
    function, and a second renderer walks it."""
    for name, out in results.items():
        assert out._expr is not None, name
        assert isinstance(out._expr, FuncCall) and out._expr.func == name, (name, out._expr)
        node = mo.to_json(out._expr)
        assert node["kind"] == "funccall" and node["func"] == name, (name, node)


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
        _calls({
            "ROUNDUP": mo.ROUNDUP(b, 0),
            "ROUNDDOWN": mo.ROUNDDOWN(b, 0),
            "TRUNC": mo.TRUNC(b),
            "CEILING": mo.CEILING(b, 1),
            "FLOOR": mo.FLOOR(b, 1),
            "POWER": mo.POWER(c, 0.5),
            "SQRT": mo.SQRT(c),
            "EXP": mo.EXP(b),
            "LN": mo.LN(c),
            "LOG": mo.LOG(c, 10),
            "LOG10": mo.LOG10(c),
            "SIGN": mo.SIGN(a),
            "PRODUCT": mo.PRODUCT(a, b, c),
        })

    def test_conditional_function(self):
        cond = mo.Variable(True)
        a = mo.Variable(1)
        b = mo.Variable(2)
        assert mo.IF(cond, a, b)._expr is not None
        keyed = mo.Variable({'a': 1, 'b': 2})
        assert mo.CHOOSE(mo.Variable(1), keyed)._expr is not None
        other = mo.Variable(False)
        _calls({
            "AND": mo.AND(cond, other),
            "OR": mo.OR(cond, other),
            "NOT": mo.NOT(cond),
            "ISBLANK": mo.ISBLANK(a),
        })

    def test_logical_functions(self):
        cond = mo.Variable(True)
        other = mo.Variable(False)
        x = mo.Variable(10.0)
        zero = mo.Variable(0.0)
        n = mo.Variable(12)
        _calls({
            "IFERROR": mo.IFERROR(x / zero, 0),
            "ISERROR": mo.ISERROR(x / zero),
            "ISNUMBER": mo.ISNUMBER(x),
            "XOR": mo.XOR(cond, other),
            "IFS": mo.IFS(cond, x, True, zero),
            "SWITCH": mo.SWITCH(n, 12, "monthly", 4, "quarterly", "other"),
        })

    def test_conditional_aggregates(self):
        row = mo.Variable([5.0, -2.0, 7.0])
        _calls({
            "SUMIF": mo.SUMIF(row, ">0"),
            "SUMIFS": mo.SUMIFS(row, row, ">0"),
            "COUNTIF": mo.COUNTIF(row, ">0"),
            "COUNTIFS": mo.COUNTIFS(row, ">0"),
            "AVERAGEIF": mo.AVERAGEIF(row, ">0"),
            "AVERAGEIFS": mo.AVERAGEIFS(row, row, ">0"),
            "MAXIFS": mo.MAXIFS(row, row, ">0"),
            "MINIFS": mo.MINIFS(row, row, ">0"),
            "COUNT": mo.COUNT(row),
            "COUNTA": mo.COUNTA(row),
        })

    def test_lookup_functions(self):
        years = mo.Variable([2026, 2027, 2028])
        rates = mo.Variable([0.05, 0.04, 0.03])
        year = mo.Variable(2027)
        _calls({
            "MATCH": mo.MATCH(year, years, 0),
            "INDEX": mo.INDEX(rates, 2),
            "XLOOKUP": mo.XLOOKUP(year, years, rates),
            "VLOOKUP": mo.VLOOKUP(2026, [years, rates], 2, False),
            "HLOOKUP": mo.HLOOKUP(year, [years, rates], 2, False),
        })

    def test_date_functions(self):
        d = mo.Variable(dt.date(2026, 4, 24))
        assert mo.YEAR(d)._expr is not None
        assert mo.MONTH(d)._expr is not None
        assert mo.DAY(d)._expr is not None
        assert mo.EDATE(d, 3)._expr is not None
        assert mo.EOMONTH(d, 0)._expr is not None
        end = mo.Variable(dt.date(2027, 1, 15))
        _calls({
            "DATE": mo.DATE(mo.Variable(2026), 1, 31),
            "DAYS360": mo.DAYS360(d, end),
            "YEARFRAC": mo.YEARFRAC(d, end),
            "DAYS": mo.DAYS(end, d),
            "TODAY": mo.TODAY(),
            "DATEDIF": mo.DATEDIF(d, end, "M"),
            "WEEKDAY": mo.WEEKDAY(d),
            "WEEKNUM": mo.WEEKNUM(d),
            "NETWORKDAYS": mo.NETWORKDAYS(d, end),
            "WORKDAY": mo.WORKDAY(d, 10),
        })

    def test_text_functions(self):
        s = mo.Variable("hello")
        t = mo.Variable("world")
        assert mo.LEN(s)._expr is not None
        assert mo.UPPER(s)._expr is not None
        assert mo.LOWER(s)._expr is not None
        assert mo.CONCAT(s, t)._expr is not None
        padded = mo.Variable("  hello world  ")
        _calls({
            "LEFT": mo.LEFT(t, 3),
            "RIGHT": mo.RIGHT(t, 3),
            "MID": mo.MID(t, 2, 3),
            "TRIM": mo.TRIM(padded),
            "FIND": mo.FIND("o", t),
            "SEARCH": mo.SEARCH("O", t),
            "SUBSTITUTE": mo.SUBSTITUTE(t, "o", "0"),
            "VALUE": mo.VALUE(mo.Variable("12.5")),
            "TEXT": mo.TEXT(mo.Variable(3.7), "0.00"),
            "REPT": mo.REPT(t, 2),
            "TEXTJOIN": mo.TEXTJOIN(", ", True, s, t),
        })

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
        dates = mo.Variable([dt.date(2026, 1, 1), dt.date(2026, 7, 1),
                             dt.date(2027, 1, 1), dt.date(2027, 7, 1)])
        monthly = mo.Variable(0.01)
        payment = mo.Variable(-100)
        cost, salvage, life = mo.Variable(10_000), mo.Variable(1_000), mo.Variable(5)
        _calls({
            "XIRR": mo.XIRR(cf, dates),
            "XNPV": mo.XNPV(rate, cf, dates),
            "NPER": mo.NPER(monthly, payment, pv),
            "IPMT": mo.IPMT(monthly, 1, nper, pv),
            "PPMT": mo.PPMT(monthly, 1, nper, pv),
            "RATE": mo.RATE(nper, payment, pv),
            "CUMIPMT": mo.CUMIPMT(monthly, nper, pv, 1, 12, 0),
            "CUMPRINC": mo.CUMPRINC(monthly, nper, pv, 1, 12, 0),
            "SLN": mo.SLN(cost, salvage, life),
            "DB": mo.DB(cost, salvage, life, 1),
            "DDB": mo.DDB(cost, salvage, life, 1),
            "SYD": mo.SYD(cost, salvage, life, 1),
        })

    def test_statistics_functions(self):
        row = mo.Variable([10.0, 40.0, 20.0, 30.0])
        k = mo.Variable(0.25)
        x = mo.Variable(20.0)
        _calls({
            "MEDIAN": mo.MEDIAN(row),
            "STDEV_S": mo.STDEV_S(row),
            "STDEV_P": mo.STDEV_P(row),
            "VAR_S": mo.VAR_S(row),
            "VAR_P": mo.VAR_P(row),
            "PERCENTILE_INC": mo.PERCENTILE_INC(row, k),
            "PERCENTILE_EXC": mo.PERCENTILE_EXC(row, 0.4),
            "QUARTILE_INC": mo.QUARTILE_INC(row, 1),
            "QUARTILE_EXC": mo.QUARTILE_EXC(row, 3),
            "LARGE": mo.LARGE(row, 2),
            "SMALL": mo.SMALL(row, 2),
            "RANK_EQ": mo.RANK_EQ(x, row),
            "RANK_AVG": mo.RANK_AVG(x, row, 1),
            "STDEV": mo.STDEV(row),
            "STDEVP": mo.STDEVP(row),
            "VAR": mo.VAR(row),
            "VARP": mo.VARP(row),
            "PERCENTILE": mo.PERCENTILE(row, k),
            "QUARTILE": mo.QUARTILE(row, 1),
            "RANK": mo.RANK(x, row),
        })

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

    def test_every_exported_function_is_called_here(self):
        called = set(re.findall(r"\bmo\.([A-Z][A-Z0-9_]*)\(", Path(__file__).read_text()))
        exported = {n for n in functions.__all__ if n.isupper()}
        assert exported - called == set(), "call each new export in a test above"
