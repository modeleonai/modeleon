# SPDX-License-Identifier: Apache-2.0
"""Cross-scope-reference diagnostics.

When a formula references a Variable outside the current emission's
layout, the Excel renderer falls back to inlining the value as a
literal. Historically this happened silently, which made two beginner
mistakes hard to spot:

1. Cross-model reference — using ``m1.s.x`` in a formula inside
   ``m2``, then emitting ``m2``.
2. Floating Variable — a Variable created outside any
   ``with MultiVariable(...):`` block, then used in a formula that
   gets emitted.

Both now fire :class:`CrossScopeReferenceWarning` at emission time.
The warning names the orphan Variable so the fix is obvious, and is
filtered as ``'always'`` so it surfaces on every notebook re-run.
"""

from __future__ import annotations

import warnings

import modeleon as mo


class TestCrossModelReference:
    """Using a Variable from one model inside another's formula warns."""

    def test_emits_warning_and_inlines_value(self, tmp_path):
        m1 = mo.MultiVariable("M1")
        m1.s = mo.MultiVariable("S", excel_props={'tab': True})
        with m1.s as s:
                s.a = mo.Variable(10)

        m2 = mo.MultiVariable("M2")
        m2.s2 = mo.MultiVariable("S2", excel_props={'tab': True})
        with m2.s2 as s2:
                # s.a lives in m1, not m2 — the next line makes s2.b
                # depend on a Variable that won't have a cell address
                # when we emit m2.
                s2.b = s.a * 2

        with warnings.catch_warnings(record=True) as caught:
            warnings.simplefilter("always")
            m2.to_excel(tmp_path / "m2.xlsx")

        cross_scope = [
            w for w in caught
            if issubclass(w.category, mo.CrossScopeReferenceWarning)
        ]
        assert len(cross_scope) >= 1, (
            "expected at least one CrossScopeReferenceWarning when m2 "
            "emits a formula that references m1.s.a"
        )
        msg = str(cross_scope[0].message)
        # Message names the orphan and nudges toward the fix.
        assert "a" in msg
        assert "with" in msg.lower() or "model" in msg.lower()


class TestFloatingVariableReference:
    """Variable created outside any with-block, used in a formula, warns."""

    def test_emits_warning(self, tmp_path):
        floating = mo.Variable(42, display_name="Floating")
        model = mo.MultiVariable("M")
        model.s = mo.MultiVariable("S", excel_props={'tab': True})
        with model.s as s:
                s.x = floating * 2

        with warnings.catch_warnings(record=True) as caught:
            warnings.simplefilter("always")
            model.to_excel(tmp_path / "m.xlsx")

        cross_scope = [
            w for w in caught
            if issubclass(w.category, mo.CrossScopeReferenceWarning)
        ]
        assert len(cross_scope) >= 1


class TestInScopeReferenceStaysSilent:
    """Variables correctly attached to the emission tree produce no warning.

    Guards against over-broad matching: a plain Variable used inside
    its own model must not trigger the warning.
    """

    def test_no_warning_for_normal_model(self, tmp_path):
        model = mo.MultiVariable("M")
        model.s = mo.MultiVariable("S", excel_props={'tab': True})
        with model.s as s:
                s.a = mo.Variable(10)
                s.b = mo.Variable(20)
                s.total = s.a + s.b

        with warnings.catch_warnings(record=True) as caught:
            warnings.simplefilter("always")
            model.to_excel(tmp_path / "m.xlsx")

        cross_scope = [
            w for w in caught
            if issubclass(w.category, mo.CrossScopeReferenceWarning)
        ]
        assert cross_scope == [], (
            "CrossScopeReferenceWarning fired on a normal in-scope model; "
            "the fallback path is over-triggering."
        )


class TestPublicSurface:
    """Warning class is reachable from ``import modeleon as mo``."""

    def test_importable_from_top_level(self):
        assert issubclass(mo.CrossScopeReferenceWarning, UserWarning)
