"""The coordinate law holds on every address path, not just VarRef.

``mo.lag``, range functions (``mo.SUM``) and the cumsum chain resolve
their operands through their own address paths; read straight from the
address book, inside a ``plan=`` coordinate they would pick up the
operand's live row. These pin each path: same-track row when laid
out, the true value otherwise — never the blend row.
"""

import os
import tempfile

import openpyxl

import modeleon as mo
from modeleon.compile.excel.view import ExcelView


def _book(model, name):
    path = os.path.join(tempfile.mkdtemp(), "t.xlsx")
    model.to_excel(path)
    return openpyxl.load_workbook(path)[name]


def _rows(ws):
    return {ws.cell(r, 1).value: r for r in range(1, ws.max_row + 1)}


def _cells(ws, label):
    r = _rows(ws)[label]
    return [ws.cell(r, c).value for c in range(2, ws.max_column + 1)
            if ws.cell(r, c).value is not None]


def _period_cols(ws, label):
    """Column letters of the row's value cells, in period order."""
    from openpyxl.utils import get_column_letter
    r = _rows(ws)[label]
    return [get_column_letter(c) for c in range(2, ws.max_column + 1)
            if ws.cell(r, c).value is not None]


def _model(mode):
    m = mo.Model("m", display_name="M",
                 tracks=mo.Tracks("plan", "actual",
                                  blend=mo.blend(given="actual", follow="plan",
                                                 until="2026-02")),
                 default_grain="month", default_start="2026-01",
                 default_periods=4)
    m.default_excel_view = ExcelView(tracks=mode)
    m.s = mo.MultiVariable("S", excel_props={"tab": True})
    with m.s as s:
        s.d = mo.Variable(plan=[1.0, 2.0, 3.0, 4.0], actual=[10.0, 20.0, 30.0, 40.0],
                          display_name="D")
        s.rev = mo.Variable(plan=s.d * 2, actual=[5.0, 6.0, 7.0, 8.0],
                            display_name="Rev")
        s.lagged = mo.Variable(plan=mo.lag(s.rev) + 1, display_name="Lagged")
        s.tot = mo.Variable(plan=mo.SUM(s.rev), display_name="Tot")
        s.csd = mo.Variable(plan=mo.cumsum(s.rev) + s.d, display_name="CumRevD")
    return m


def test_lag_reads_the_same_track_row_in_rows_mode():
    ws = _book(_model("rows"), "S")
    rows = _rows(ws)
    rp = rows["Rev · plan"]
    c = _period_cols(ws, "Rev · plan")
    lag = _cells(ws, "Lagged · plan")
    assert lag[1] == f"={c[0]}{rp} + 1" and lag[2] == f"={c[1]}{rp} + 1", lag


def test_lag_inlines_the_track_value_in_blend_mode():
    # Rev's head is a splice (actual authored) — no plan row exists,
    # so the lagged PLAN value is inlined: 2.0 (Jan) + 1 in February.
    ws = _book(_model("blend"), "S")
    lag = _cells(ws, "Lagged")
    assert lag[1] == "=2.0 + 1" and lag[2] == "=4.0 + 1", lag


def test_sum_ranges_over_the_same_track_row():
    ws = _book(_model("rows"), "S")
    rows = _rows(ws)
    rp = rows["Rev · plan"]
    c = _period_cols(ws, "Rev · plan")
    tot = _cells(ws, "Tot · plan")
    assert tot[0] == f"=SUM({c[0]}{rp}:{c[-1]}{rp})", tot


def test_sum_degrades_to_the_value_without_a_coordinate_row():
    # A range over a coordinate nobody lays out cannot be a formula:
    # the cell holds the computed number (2+4+6+8 = 20), not the live
    # row's sum (5+6+7+8 = 26).
    ws = _book(_model("blend"), "S")
    tot = _cells(ws, "Tot")
    assert tot[0] == 20.0, tot


def test_cumsum_plus_series_degrades_to_true_values():
    # ``cumsum(rev) + d`` with d a tracked SERIES cannot fold d into
    # the chain's seed (it varies by period); the composition has no
    # honest chain formula, so the cells hold the true values —
    # never the folded formula that dropped the per-period term.
    ws = _book(_model("rows"), "S")
    c = _cells(ws, "CumRevD · plan")
    assert c == [3.0, 8.0, 15.0, 24.0], c
