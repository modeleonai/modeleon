# SPDX-License-Identifier: Apache-2.0
"""``check_compat`` — surface FuncCall nodes that need a fallback path
for a given backend target.

Pairs with the fallback contract in ``test_renderer_fallback.py``:
that test locks in the renderer behavior when a FuncCall declares a
``render_backends`` hint the renderer isn't in; this test locks in the
inspection API that lets callers ask "which nodes will fall back?"
before ever invoking a renderer.
"""

from __future__ import annotations

import modeleon as mo
from modeleon.compile import CompatIssue, check_compat


class TestUniversalIsAlwaysCompatible:
    """Functions with no backend hint are universal and never surface."""

    def test_sum_compatible_with_any_target(self):
        a = mo.Variable(10)
        b = mo.Variable(20)
        total = mo.SUM(a, b)
        assert check_compat(total._expr, "excel") == []
        assert check_compat(total._expr, "pandas") == []
        assert check_compat(total._expr, "made-up-backend") == []


class TestDeclaredBackendsAreEnforced:
    """``render_backends={'excel'}`` hides the call from non-excel targets."""

    def test_excel_only_financial_surfaces_on_pandas(self):
        cf = mo.Variable([-100_000, 30_000, 40_000, 50_000, 60_000])
        project_irr = mo.IRR(cf)
        # Excel target — IRR is native, nothing to report.
        assert check_compat(project_irr, "excel") == []
        # Pandas target — IRR falls back.
        issues = check_compat(project_irr, "pandas")
        assert len(issues) == 1
        assert isinstance(issues[0], CompatIssue)
        assert issues[0].func == "IRR"
        assert issues[0].kind == "render"
        assert issues[0].declared == frozenset({"excel"})
        assert issues[0].target == "pandas"

    def test_excel_only_date_helpers_surface_on_sql(self):
        eom = mo.EOMONTH("2025-01-15", 0)
        assert check_compat(eom, "excel") == []
        issues = check_compat(eom, "sql")
        assert len(issues) == 1
        assert issues[0].func == "EOMONTH"

    def test_nested_funccalls_all_reported(self):
        cf = mo.Variable([-100_000, 30_000, 40_000, 50_000])
        rate = mo.Variable(0.10)
        # NPV + IRR in one expression — both are Excel-only.
        composite = mo.NPV(rate, cf) + mo.IRR(cf)
        funcs = {i.func for i in check_compat(composite, "pandas")}
        assert funcs == {"NPV", "IRR"}


class TestComputeKindIsDistinct:
    """``kind='compute'`` checks the compute hint, not the render hint."""

    def test_render_only_hint_doesnt_affect_compute_check(self):
        cf = mo.Variable([-100, 50, 60, 70])
        project_irr = mo.IRR(cf)
        # We declared render_backends but NOT compute_backends for IRR —
        # it's universally computable in Python. Compute check passes.
        assert check_compat(project_irr, "pandas", kind="compute") == []


class TestInvalidKind:
    """Typos fail loudly — no silent coercion."""

    def test_bad_kind_raises(self):
        v = mo.Variable(1) + mo.Variable(2)
        try:
            check_compat(v, "excel", kind="both")
        except ValueError as e:
            assert "render" in str(e) and "compute" in str(e)
        else:
            raise AssertionError("check_compat should reject unknown kind")
