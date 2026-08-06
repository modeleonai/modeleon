# SPDX-License-Identifier: Apache-2.0
"""Golden-output regression tests — pin exact emitted formulas.

Complements :file:`test_excel_writer.py` (structural checks) and
:file:`test_renderer_fallback.py` (fallback contract). Those verify
invariants; these verify that *this specific* model produces *these
specific* cell contents. A single-character drift in the translator,
layout engine, or rendering pipeline will break one of these tests
immediately — before shipping an .xlsx that's structurally valid but
numerically different.

Each test builds a deterministic model (attribute assignment keeps
cell order stable) and asserts exact coordinates + values. Keep the
models small so a failure points to the regression unambiguously.

Adding a test: construct a model, run it once locally to see the
emitted cells, paste them into ``EXPECTED``. The raw Python dict is
the source of truth — no "roughly equals" assertions.
"""

from __future__ import annotations

import pytest
from openpyxl import load_workbook

import modeleon as mo


def _cells(path, sheet_name: str) -> dict[str, object]:
    """Read every non-empty cell on a sheet as ``{coord: value}``."""
    ws = load_workbook(path)[sheet_name]
    return {
        cell.coordinate: cell.value
        for row in ws.iter_rows()
        for cell in row
        if cell.value is not None
    }


class TestSingleSheetArithmetic:
    """Two inputs + two derived formulas on one tab.

    Exercises the baseline translator path: ``Literal`` (inputs),
    ``BinOp`` over ``VarRef`` (cross-row references using the same
    sheet's addresses).
    """

    def test_pinned_cells(self, tmp_path):
        model = mo.MultiVariable("Acme")
        model.pnl = mo.MultiVariable("PnL", excel_props={'tab': True})
        with model.pnl as pnl:
                pnl.revenue = mo.Variable(1_000_000, display_name='Revenue')
                pnl.cogs_pct = mo.Variable(0.6, display_name='COGS %')
                pnl.cogs = pnl.revenue * pnl.cogs_pct
                pnl.gross = pnl.revenue - pnl.cogs

        model.to_excel(tmp_path / "pnl.xlsx")

        path = tmp_path / "pnl.xlsx"
        assert _cells(path, "PnL") == {
            "A1": "Revenue",    "B1": 1_000_000,
            "A2": "COGS %",     "B2": 0.6,
            "A3": "Cogs",       "B3": "=B1 * B2",
            "A4": "Gross",      "B4": "=B1 - B3",
        }


class TestCrossSheetReferences:
    """A formula variable on one tab that references inputs on another.

    Exercises the translator's sheet-prefix path — the emitted formula
    must qualify addresses from other sheets as ``Inputs!B1``, not bare
    ``B1`` (which would silently reference the wrong cell on the
    current sheet).
    """

    def test_pinned_cells(self, tmp_path):
        model = mo.MultiVariable("Acme")
        model.inputs = mo.MultiVariable("Inputs", excel_props={'tab': True})
        model.inputs.price = mo.Variable(100.0, display_name='Price')
        model.inputs.volume = mo.Variable(50, display_name='Volume')
        model.outputs = mo.MultiVariable("Outputs", excel_props={'tab': True})
        model.outputs.revenue = model.inputs.price * model.inputs.volume

        model.to_excel(tmp_path / "x.xlsx")

        path = tmp_path / "x.xlsx"
        assert _cells(path, "Inputs") == {
            "A1": "Price",      "B1": 100.0,
            "A2": "Volume",     "B2": 50,
        }
        assert _cells(path, "Outputs") == {
            "A1": "Revenue",    "B1": "=Inputs!B1 * Inputs!B2",
        }


class TestDuplicateSheetNameQualification:
    """Two sibling tabs that derive the SAME sheet name must still qualify
    cross-tab references.

    A model and its ``.at(grain=…)`` projection both carry the same
    ``display_name`` (``'Saas'``), so both derive sheet name ``'Saas'``.
    openpyxl de-duplicates the physical tab titles (``'Saas'`` /
    ``'Saas1'``); the writer must use those post-dedup names everywhere so
    the quarterly tab's re-grain formulas keep the ``'Saas'!`` prefix when
    they reference the monthly tab's cells. Without it the prefix is
    silently dropped and the formula points at the wrong cells on its own
    tab. Regression for the projection cross-tab qualification bug.
    """

    def test_regrained_tab_qualifies_cross_tab_refs(self, tmp_path):
        m = mo.Model("SaaS")
        m.default_start = "2024-01"
        m.default_grain = "month"
        m.seats = mo.Variable([10, 20, 30, 40, 50, 60], regrain=mo.up("last"))
        m.price = mo.Variable([5, 5, 5, 5, 5, 5], regrain=mo.up("mean"))
        # A multiplicative flow: Σ(seats·price) ≠ aggSeats·aggPrice, so it
        # carries its own up('sum') rule and re-grains by aggregating itself.
        m.revenue = (m.seats * m.price).set_regrain(mo.up("sum"))

        compare = mo.Model("Compare")
        compare.m = m
        compare.m_q = m.at(grain="quarter")

        # The projection actually aggregates (6 monthly -> 2 quarterly points),
        # not just relabels: Q1 = 50+100+150, Q2 = 200+250+300.
        assert compare.m_q.revenue.value == [300, 750]

        path = tmp_path / "x.xlsx"
        compare.to_excel(path)

        # Two physically distinct tabs despite the shared display_name.
        assert load_workbook(path).sheetnames == ["Saas", "Saas1"]

        # The model declares a window, so row 1 is the timeline header and data
        # starts at row 2. Monthly tab: header + same-sheet arithmetic.
        mo_cells = _cells(path, "Saas")
        assert mo_cells["B1"] == "Jan 2024"            # finance header (default)
        assert mo_cells["B4"] == "=B2 * B3"            # revenue = seats * price

        # Quarterly tab: header in row 1; every re-grain formula references the
        # monthly tab's (shifted) cells AND aggregates over the quarter's months
        # — a SUM/AVERAGE over a range, or a period-end cell — never a single
        # mis-pointed monthly cell.
        q = _cells(path, "Saas1")
        assert q["B1"] == "Q1 2024"
        assert q["C1"] == "Q2 2024"
        assert q["B2"] == "=Saas!D2"               # seats, up('last') -> Q-end
        assert q["B3"] == "=AVERAGE(Saas!B3:D3)"   # price, up('mean')
        assert q["B4"] == "=SUM(Saas!B4:D4)"       # revenue, up('sum') -> Q1
        assert q["C4"] == "=SUM(Saas!E4:G4)"       # revenue, up('sum') -> Q2


class TestLongNameDedupTerminates:
    """Two sibling tabs whose names collide *at the 31-char Excel cap* must still
    emit and de-duplicate without spinning. Regression for an infinite loop in
    the old hand-rolled sheet-name dedup, which re-truncated ``'X'*31 + suffix``
    back to ``'X'*31`` forever; openpyxl now owns the dedup, so this terminates.
    """

    def test_thirtyone_char_collision_emits(self, tmp_path):
        long_name = "X" * 40  # both truncate to 'X'*31 and collide at the cap
        parent = mo.Model("Book")
        a = mo.MultiVariable(long_name, excel_props={"tab": True})
        a.rev = mo.Variable([1, 2, 3], display_name="Rev")
        b = mo.MultiVariable(long_name, excel_props={"tab": True})
        b.cost = mo.Variable([4, 5, 6], display_name="Cost")
        parent.a = a
        parent.b = b

        parent.to_excel(tmp_path / "x.xlsx")  # must return, not hang
        names = load_workbook(tmp_path / "x.xlsx").sheetnames
        assert len(names) == 2 and names[0] != names[1]  # distinct tabs
        assert all(len(n) <= 31 for n in names)          # Excel-legal


class TestFuncCallRendering:
    """Named-function call — ``SUM(a, b, c)`` renders as ``=SUM(B1, B2, B3)``.

    Exercises ``FuncCall`` arg rendering: each ``VarRef`` resolves to
    the referenced Variable's cell address, joined by ``, ``.
    """

    def test_pinned_cells(self, tmp_path):
        model = mo.MultiVariable("M")
        model.s = mo.MultiVariable("S", excel_props={'tab': True})
        with model.s as s:
                s.a = mo.Variable(10.0, display_name='A')
                s.b = mo.Variable(20.0, display_name='B')
                s.c = mo.Variable(30.0, display_name='C')
                s.total = mo.SUM(s.a, s.b, s.c)

        model.to_excel(tmp_path / "sum.xlsx")

        path = tmp_path / "sum.xlsx"
        assert _cells(path, "S") == {
            "A1": "A",          "B1": 10.0,
            "A2": "B",          "B2": 20.0,
            "A3": "C",          "B3": 30.0,
            "A4": "Total",      "B4": "=SUM(B1, B2, B3)",
        }


class TestSelfRefTimeSeries:
    """``recurrence`` — period 0 uses the start value; later periods use the
    recurrence template with ``{prev}`` bound to the previous column.

    This is the regression surface that bit us most often historically:
    period-0 vs. period-N divergence, the ``{prev}`` substitution, and
    the expansion across a time axis driven by a list-valued input.
    """

    def test_pinned_cells(self, tmp_path):
        model = mo.MultiVariable("M")
        model.s = mo.MultiVariable("S", excel_props={'tab': True})
        with model.s as s:
                s.growth = mo.Variable([1.1, 1.1, 1.1], display_name='Growth')
                s.users = mo.recurrence(100, '{prev} * {growth}', growth=s.growth)

        model.to_excel(tmp_path / "t.xlsx")

        # New semantics: period 0 = start (literal 100, no formula);
        # periods 1+ apply the recurrence template.
        path = tmp_path / "t.xlsx"
        assert _cells(path, "S") == {
            "A1": "Growth",     "B1": 1.1,          "C1": 1.1,          "D1": 1.1,
            "A2": "Users",      "B2": "=100",       "C2": "=B2 * C1",   "D2": "=C2 * D1",
        }


class TestBinOpPrecedenceParens:
    """Mixed-precedence binops must keep their parentheses in the
    emitted Excel formula.

    The current renderer (``render_binop`` in
    :file:`compile/excel/renderer.py`) emits ``{left} {op} {right}``
    flat, with no precedence-aware wrapping. ``__truediv__`` /
    ``__rtruediv__`` / ``__neg__`` pass ``parenthesize=True`` to the
    arithmetic chokepoint as a partial workaround — that's why
    ``(a + b) / c`` already comes out correct (covered as a
    counter-example below). But ``__mul__`` / ``__pow__`` / etc.
    don't, so ``a * (1 - b)`` collapses to ``=B1 * 1 - B2`` and Excel
    silently computes the wrong result.

    These tests pin the **correct** emitted formulas. The two known-
    broken shapes are xfail-marked; flip them to pass once the
    renderer compares parent/child operator precedence and wraps
    accordingly.
    """

    def test_subtraction_inside_multiplication(self, tmp_path):
        # ``a * (1 - b)`` — subtraction binds looser than multiplication,
        # so the right-hand subtree needs parens. Multiplication has no
        # ``parenthesize=True`` workaround.
        model = mo.MultiVariable("M")
        model.s = mo.MultiVariable("S", excel_props={'tab': True})
        with model.s as s:
            s.a = mo.Variable(2.0, display_name='A')
            s.b = mo.Variable(0.5, display_name='B')
            s.x = (s.a * (1 - s.b)).set_display_name('X')

        model.to_excel(tmp_path / "p.xlsx")

        cells = _cells(tmp_path / "p.xlsx", "S")
        assert cells["B3"] == "=B1 * (1 - B2)"

    def test_division_and_exponent_groups(self, tmp_path):
        # ``(a / b) ** (1 / c)`` — both children of ``**`` are divisions
        # which bind tighter than ``**`` in Excel's left-associative
        # exponent. Parens needed to preserve grouping.
        model = mo.MultiVariable("M")
        model.s = mo.MultiVariable("S", excel_props={'tab': True})
        with model.s as s:
            s.a = mo.Variable(2.0, display_name='A')
            s.b = mo.Variable(0.5, display_name='B')
            s.c = mo.Variable(3.0, display_name='C')
            s.y = ((s.a / s.b) ** (1 / s.c)).set_display_name('Y')

        model.to_excel(tmp_path / "p.xlsx")

        cells = _cells(tmp_path / "p.xlsx", "S")
        assert cells["B4"] == "=(B1 / B2) ^ (1 / B3)"

    def test_addition_inside_division_keeps_parens(self, tmp_path):
        # Counter-example pinning the existing partial workaround:
        # ``__truediv__`` passes ``parenthesize=True`` so its operands
        # do get wrapped. ``(a + b) / c`` therefore renders correctly
        # today; if the renderer fix lands and removes this hack, the
        # general precedence-aware logic must still produce the same
        # output here.
        model = mo.MultiVariable("M")
        model.s = mo.MultiVariable("S", excel_props={'tab': True})
        with model.s as s:
            s.a = mo.Variable(10.0, display_name='A')
            s.b = mo.Variable(20.0, display_name='B')
            s.c = mo.Variable(5.0, display_name='C')
            s.z = ((s.a + s.b) / s.c).set_display_name('Z')

        model.to_excel(tmp_path / "p.xlsx")

        cells = _cells(tmp_path / "p.xlsx", "S")
        assert cells["B4"] == "=(B1 + B2) / B3"

    def test_same_precedence_does_not_wrap(self, tmp_path):
        # When child and parent share precedence (e.g. both
        # multiplication), left-associativity already gives the right
        # result and no extra parens should appear — the future fix
        # must not over-wrap.
        model = mo.MultiVariable("M")
        model.s = mo.MultiVariable("S", excel_props={'tab': True})
        with model.s as s:
            s.a = mo.Variable(2.0, display_name='A')
            s.b = mo.Variable(3.0, display_name='B')
            s.c = mo.Variable(4.0, display_name='C')
            s.p = (s.a * s.b * s.c).set_display_name('P')

        model.to_excel(tmp_path / "p.xlsx")

        cells = _cells(tmp_path / "p.xlsx", "S")
        assert cells["B4"] == "=B1 * B2 * B3"

    def test_comparison_inside_multiplication_keeps_parens(self, tmp_path):
        # ``a * (b == c)`` — comparison is the lowest precedence, so it
        # MUST be parenthesized: without parens Excel re-reads
        # ``a * b = c`` as ``(a * b) = c`` (a boolean), silently changing
        # the result. Common in flag arithmetic (``prev * (flag = 0)``).
        model = mo.MultiVariable("M")
        model.s = mo.MultiVariable("S", excel_props={'tab': True})
        with model.s as s:
            s.a = mo.Variable(1.0, display_name='A')
            s.b = mo.Variable(0.0, display_name='B')
            s.x = (s.a * (s.b == 0)).set_display_name('X')

        model.to_excel(tmp_path / "p.xlsx")

        cells = _cells(tmp_path / "p.xlsx", "S")
        assert cells["B3"] == "=B1 * (B2 = 0)"
