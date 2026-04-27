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
