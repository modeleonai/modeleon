# SPDX-License-Identifier: Apache-2.0
"""Tests for the ``.code`` property — the runtime → DSL source
renderer (``display/code.py``).

Mirrors the JSON / HTML renderer test patterns: build a small Model
or MultiVariable, call ``.code``, and check the output text and its
round-trip via ``exec``.
"""

from __future__ import annotations

import modeleon as mo


# ─── Single Variable rendering ───────────────────────────────────


def test_input_variable_self_relative_lhs():
    m = mo.Model("acme")
    m.x = mo.Variable(10)
    assert m.x.code == "x = mo.Variable(10)"


def test_variable_with_unit_kwarg():
    m = mo.Model("acme")
    m.revenue = mo.Variable(100, unit="USD")
    assert m.revenue.code == "revenue = mo.Variable(100, unit='USD')"


def test_variable_with_explicit_display_name():
    m = mo.Model("acme")
    m.metric = mo.Variable(5, display_name="A Custom Label")
    assert "display_name='A Custom Label'" in m.metric.code


def test_variable_skips_auto_humanized_display_name():
    """Adoption doesn't set ``_display_name`` for the Variable itself
    (it falls back to humanized python_name on read), so no display_name
    kwarg should leak into the rendered code unless explicit."""
    m = mo.Model("acme")
    m.tax_rate = mo.Variable(0.25)
    assert m.tax_rate.code == "tax_rate = mo.Variable(0.25)"


def test_leaf_variable_with_out_of_scope_refs_falls_back_to_value():
    """Per the Excel-shaped rule: when a Variable's formula references
    siblings that aren't in the snippet's scope (here: rendering the
    formula leaf alone), emit the value form."""
    m = mo.Model("acme")
    m.a = mo.Variable(10)
    m.b = mo.Variable(20)
    m.total = m.a + m.b
    assert m.total.code == "total = mo.Variable(30)"


# ─── MultiVariable / Model rendering ─────────────────────────────


def test_model_header_uses_mo_model_constructor():
    m = mo.Model("acme")
    code = m.code
    assert code.startswith("acme = mo.Model('acme')")


def test_bare_multivariable_header_uses_mv_constructor():
    m = mo.Model("acme")
    m.metrics = mo.MultiVariable()
    assert "acme.metrics = mo.MultiVariable()" in m.code


def test_model_emits_full_subtree():
    m = mo.Model("acme")
    m.x = mo.Variable(1)
    m.y = mo.Variable(2)
    code = m.code
    lines = code.strip().split("\n")
    assert lines[0] == "acme = mo.Model('acme')"
    assert "acme.x = mo.Variable(1)" in lines
    assert "acme.y = mo.Variable(2)" in lines


def test_formula_in_scope_emits_qualified_formula():
    m = mo.Model("acme")
    m.a = mo.Variable(10)
    m.b = mo.Variable(20)
    m.sum = m.a + m.b
    code = m.code
    assert "acme.sum = mo.Variable(acme.a + acme.b)" in code


def test_nested_mv_qualifies_full_path_in_formulas():
    """Refs inside formulas use the FULL attribute path from the
    snippet root, not just leaf names — so ``co.org.dept.headcount``
    not ``co.headcount``."""
    m = mo.Model("co")
    m.org = mo.MultiVariable()
    m.org.dept = mo.MultiVariable()
    m.org.dept.headcount = mo.Variable(100)
    m.org.dept.salary = mo.Variable(50000)
    m.org.dept.cost = m.org.dept.headcount * m.org.dept.salary
    code = m.code
    assert (
        "co.org.dept.cost = mo.Variable("
        "co.org.dept.headcount * co.org.dept.salary)"
    ) in code


def test_unnamed_temporary_inlined_as_literal():
    """``revenue * mo.Variable(0.6)`` — the 0.6 wrapper has no
    ``python_name`` and isn't part of the model tree; the renderer
    folds it back into the formula as the literal ``0.6``."""
    m = mo.Model("acme")
    m.revenue = mo.Variable(100)
    m.cogs = m.revenue * mo.Variable(0.6)
    code = m.code
    assert "acme.cogs = mo.Variable(acme.revenue * 0.6)" in code


def test_subtree_renders_self_relative():
    """``mv.code`` on a sub-MV emits a snippet rooted at that MV —
    the parent prefix is gone."""
    m = mo.Model("acme")
    m.metrics = mo.MultiVariable()
    m.metrics.cnt = mo.Variable(5)
    code = m.metrics.code
    assert code.startswith("metrics = mo.MultiVariable()")
    assert "metrics.cnt = mo.Variable(5)" in code
    # No "acme." prefix appears in this subtree view.
    assert "acme." not in code


# ─── Round-trip: exec(model.code) reconstructs an equivalent model


def test_exec_round_trip_preserves_values_and_formulas():
    m = mo.Model("acme")
    m.x = mo.Variable(10)
    m.y = mo.Variable(20)
    m.total = m.x + m.y

    ns: dict = {"mo": mo}
    exec(m.code, ns)
    rebuilt = ns["acme"]
    assert rebuilt.x.value == 10
    assert rebuilt.total.value == 30
    assert rebuilt.total.is_formula


def test_exec_round_trip_nested():
    m = mo.Model("co")
    m.org = mo.MultiVariable()
    m.org.dept = mo.MultiVariable()
    m.org.dept.headcount = mo.Variable(100)
    m.org.dept.salary = mo.Variable(50000)
    m.org.dept.cost = m.org.dept.headcount * m.org.dept.salary

    ns: dict = {"mo": mo}
    exec(m.code, ns)
    rebuilt = ns["co"]
    assert rebuilt.org.dept.cost.value == 100 * 50000
    assert rebuilt.org.dept.cost.is_formula
