# SPDX-License-Identifier: Apache-2.0
"""A Context sheet — lines that come from outside this model — is laid
out as a tab of THIS book: it stands on the book's window (start and
grain, never its periods) and wears the book's view, unless it declares
either of its own. Its lines keep no
time of their own, and no subtotal buckets are drawn over them — a
line that is not this model's to re-grain has no rule to sum or carry
by.
"""

from __future__ import annotations

import modeleon as mo
from modeleon.compile.excel.layout import LayoutEngine
from modeleon.compile.excel.writer import _collect_root_sheets
from modeleon.core.time import resolve_default_window


def _model(**view):
    m = mo.Model("Beta", default_grain="month", default_start="2026-01", default_periods=3)
    if view:
        m.default_excel_view = mo.ExcelView(**view)
    m.sales = mo.MultiVariable("Sales")
    m.sales.volume = mo.Variable([10.0, 20.0, 30.0], display_name="Volume")
    return m


def _context():
    # The marker the layout keys on: a container of lines from outside
    # this model.
    ctx = mo.MultiVariable("Context", excel_props={"tab": True})
    ctx._is_context = True
    ctx.alpha = mo.MultiVariable("Alpha")
    ctx.alpha.price = mo.Variable([2.0, 3.0, 4.0], display_name="Price")
    return ctx


def _layout(*tops, **kw):
    roots = [sheet for top in tops for sheet in _collect_root_sheets(top)]
    engine = LayoutEngine(roots, flatten_nested_sheets=True, header_rows=True, **kw)
    addresses = engine.compute_addresses()
    return engine, addresses


def test_a_context_sheet_stands_on_the_books_window_and_its_lines_keep_no_time():
    m, ctx = _model(), _context()
    engine, addresses = _layout(m, ctx)
    assert engine.time_headers_by_sheet["Sales"] == ("2026-01", "month", "finance")
    assert engine.time_headers_by_sheet["Context"] == ("2026-01", "month", "finance")
    # The book's start and grain, never its periods.
    windows = {
        engine.sheet_name_by_id[oid]: w for oid, w in engine.window_by_sheet.items()
    }
    assert windows["Context"].periods is None and windows["Sales"].periods == 3
    # Data under the header, never on row 1.
    price = addresses[ctx.alpha.price.id]
    assert price.values and all(int(a.split("!")[-1][1:]) > 1 for a in price.values), price
    # The lines themselves resolve no window: only the layout stands on it.
    assert resolve_default_window(ctx.alpha.price) is None
    assert getattr(ctx, "default_grain", None) is None


def test_a_context_sheet_wears_the_books_view_and_takes_no_buckets():
    m, ctx = _model(timeline={"label_format": "compact", "totals": ["quarter"]}), _context()
    engine, _ = _layout(m, ctx)
    assert engine.time_headers_by_sheet["Sales"][2] == "compact"
    assert engine.time_headers_by_sheet["Context"][2] == "compact"
    assert engine.totals_by_sheet.get("Sales") is not None
    assert engine.totals_by_sheet.get("Context") is None
    # A totals override handed to the layout: the same answer.
    engine, _ = _layout(_model(), _context(), totals_override=["quarter"])
    assert engine.totals_by_sheet.get("Sales") is not None
    assert engine.totals_by_sheet.get("Context") is None


def test_a_context_that_declares_its_own_window_or_view_keeps_them():
    m, ctx = _model(timeline={"label_format": "compact"}), _context()
    ctx.default_grain = "quarter"
    ctx.default_start = "2026-Q1"
    ctx.default_excel_view = mo.ExcelView(timeline={"label_format": "iso"})
    engine, _ = _layout(m, ctx)
    assert engine.time_headers_by_sheet["Context"] == ("2026-Q1", "quarter", "iso")
    assert engine.time_headers_by_sheet["Sales"] == ("2026-01", "month", "compact")


def test_without_a_windowed_root_the_context_sheet_has_no_header():
    m = mo.MultiVariable("Beta", excel_props={"tab": True})
    m.volume = mo.Variable([10.0, 20.0, 30.0], display_name="Volume")
    engine, addresses = _layout(m, _context())
    assert "Context" not in engine.time_headers_by_sheet
    assert "Beta" not in engine.time_headers_by_sheet
