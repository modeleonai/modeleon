"""A track authored as a list of REFERENCES keeps them as formulas.

``forecast=[fact, None, None, solved_cell]`` is a sparse row of links — each
milestone cell points at the cell that holds the number. The role stash
keeps the original list, so the track's expression is a ListExpr: a
reference per period where the list holds a Variable, a literal (or a
blank) elsewhere.
"""

import os
import tempfile

import openpyxl

import modeleon as mo
from modeleon.compile.excel.view import ExcelView
from modeleon.core.expr import ListExpr


def _book(model, name):
    path = os.path.join(tempfile.mkdtemp(), "t.xlsx")
    model.to_excel(path)
    return openpyxl.load_workbook(path)[name]


def _rows(ws):
    return {ws.cell(r, 1).value: r for r in range(1, ws.max_row + 1)}


def test_sparse_reference_list_emits_links():
    m = mo.Model("m", display_name="M",
                 tracks=mo.Tracks("forecast", "actual"),
                 default_grain="month", default_start="2026-01",
                 default_periods=4)
    m.default_excel_view = ExcelView(tracks="blend")
    m.s = mo.MultiVariable("S", excel_props={"tab": True})
    with m.s as s:
        s.fact = mo.Variable(280.0, display_name="Fact")
        s.solved = mo.Variable(1000.0, display_name="Solved")
        s.points = mo.Variable(forecast=[s.fact, None, s.solved * 2, s.solved],
                               display_name="Points")
    assert isinstance(m.s.points.track_expr("forecast"), ListExpr)
    ws = _book(m, "S")
    rows = _rows(ws)
    r = rows["Points"]
    cells = [ws.cell(r, c).value for c in range(2, ws.max_column + 1)]
    vals = [v for v in cells if v is not None]   # blanks and the constants column drop out
    assert vals[0] == f"=B{rows['Fact']}", vals
    assert vals[1] == f"=B{rows['Solved']} * 2", vals
    assert vals[2] == f"=B{rows['Solved']}", vals


def test_all_literal_list_stays_data():
    m = mo.Model("m", display_name="M", tracks=mo.Tracks("forecast", "actual"),
                 default_grain="month", default_start="2026-01",
                 default_periods=3)
    m.s = mo.MultiVariable("S", excel_props={"tab": True})
    with m.s as s:
        s.points = mo.Variable(forecast=[1.0, None, 3.0], display_name="Points")
    assert m.s.points.track_expr("forecast") is None
