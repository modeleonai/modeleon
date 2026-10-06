# SPDX-License-Identifier: Apache-2.0
"""The look of a hand-built book, declared once on a view.

Every capability here is off until a view asks for it — a model that
declares none writes the book it always wrote. Pinned: the book's font,
the look of a cell by what it is, formats by what a line is, the rows
repeated at the top of every sheet, a title row, the opening column
("Pre-start") the time rows chain off, full-width rows, column widths,
and the house view an application lays under a model without touching it.
"""

import warnings
from datetime import date

import openpyxl
import pytest

import modeleon as mo
from modeleon.compile.excel.view import resolve_excel_view

GREY = "#7F7F7F"


def deal(view=None):
    """Three sheets like a project-finance book: inputs, series, and
    the time anchor the others repeat at their top."""
    m = mo.Model("m", default_grain="year", default_periods=3)
    with m:
        if view is not None:
            m.default_excel_view = view
        m.ti = mo.MultiVariable("Inputs")
        with m.ti as i:
            i.start = mo.Variable(date(2026, 1, 1), display_name="Model start")
            i.rate = mo.Variable(0.05, display_name="Rate")
        m.default_start = m.ti.start
        m.td = mo.MultiVariable("Series")
        with m.td as s:
            s.volume = mo.Variable([1.0, 2.0, 3.0], display_name="Volume",
                                   excel_props={"row_sum": True})
        m.timing = mo.MultiVariable("Timing")
        with m.timing as t:
            t.ruler = mo.MultiVariable("Timeline")
            with t.ruler as r:
                r.index = mo.time.index.set_display_name("Period number")
                r.index.unit = "#"
                r.start = mo.time.start.set_display_name("Period start")
                r.end = mo.time.end.set_display_name("Period end")
                r.end.unit = "date"
            t.flags = mo.MultiVariable("Flags")
            with t.flags as f:
                f.first = mo.IF(m.timing.ruler.index == 1, 1, 0).set_display_name(
                    "First period flag")
        m.default_header = m.timing.ruler
    return m


def book(model, tmp_path):
    path = tmp_path / "book.xlsx"
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        model.to_excel(str(path))
    return openpyxl.load_workbook(str(path))


def cells(ws):
    return {c.coordinate: c.value for row in ws.iter_rows() for c in row
            if c.value is not None}


def column(ws, caption):
    """The column a header caption heads."""
    return next(c.column for row in ws.iter_rows(max_row=3) for c in row
                if c.value == caption)


LEAD = {"article": "No", "label": "Line item", "unit": "Unit", "constants": "Constant",
        "fields": [("row_sum", "Total")]}


# ══════════════════════════════════════════════════════════════════════
# Off by default
# ══════════════════════════════════════════════════════════════════════


def test_a_model_without_a_view_writes_the_book_it_always_did(tmp_path):
    wb = book(deal(), tmp_path)
    assert wb["Timing"]["A1"].font.name == "Calibri"
    assert wb["Timing"].sheet_view.showGridLines is not False
    assert wb["Inputs"]["A1"].value != "INPUTS"


# ══════════════════════════════════════════════════════════════════════
# Font, cell types, formats
# ══════════════════════════════════════════════════════════════════════


def test_the_view_font_is_the_book_font(tmp_path):
    wb = book(deal(mo.ExcelView(font={"name": "Arial", "size": 10}, meta=LEAD)), tmp_path)
    ws = wb["Timing"]
    # The header captions, the section bands and an empty cell alike.
    for ref in ("A1", "B1", "Z99"):
        assert (ws[ref].font.name, ws[ref].font.sz) == ("Arial", 10.0)
    assert wb._named_styles["Normal"].font.name == "Arial"


def test_an_input_wears_its_type_look_under_its_own(tmp_path):
    view = mo.ExcelView(cell_types={"on": True, "colors": {"input": "#0000FF"},
                                    "styles": {"input": {"bg": "#FFF2CC"}}})
    m = deal(view)
    m.ti.rate.set_style(bg="#FFFFFF")
    ws = book(m, tmp_path)["Inputs"]
    start = next(c for row in ws.iter_rows() for c in row if c.value == date(2026, 1, 1)
                 or getattr(c.value, "date", lambda: None)() == date(2026, 1, 1))
    assert start.fill.fgColor.rgb == "FFFFF2CC" and start.font.color.rgb == "FF0000FF"
    rate = next(c for row in ws.iter_rows() for c in row if c.value == 0.05)
    assert rate.fill.fgColor.rgb == "FFFFFFFF"          # the line's own look wins


def test_formats_follow_what_a_line_is(tmp_path):
    view = mo.ExcelView(meta=LEAD, formats={"date": "dd mmm yyyy", "flag": "0", "#": "0",
                                            "number": "#,##0.0"})
    ws = book(deal(view), tmp_path)["Timing"]
    first = column(ws, "Total") + 1          # the Timing sheet has no constants
    by_label = {ws.cell(r, 2).value: ws.cell(r, first) for r in range(2, ws.max_row + 1)}
    assert by_label["Period end"].number_format == "dd mmm yyyy"      # a date
    assert by_label["Period number"].number_format == "0"              # its unit


# ══════════════════════════════════════════════════════════════════════
# The rows at the top of every sheet
# ══════════════════════════════════════════════════════════════════════


def test_header_rows_wear_the_mirror_band_and_their_unit(tmp_path):
    view = mo.ExcelView(meta=LEAD, bands={"mirror": {"italic": True, "font_color": GREY}},
                        formats={"date": "dd mmm yyyy"})
    ws = book(deal(view), tmp_path)["Series"]
    got = cells(ws)
    row = next(r for r in range(1, 10) if got.get(f"B{r}") == "Period end")
    assert got[f"C{row}"] == "date"
    value = ws.cell(row, 6)      # first period after Constant: No, label, unit, Total, Constant
    assert value.font.i and value.font.color.rgb == "FF7F7F7F"
    assert value.number_format == "dd mmm yyyy"      # the format of the row it repeats


def test_same_columns_keeps_the_constant_column_on_every_sheet(tmp_path):
    view = mo.ExcelView(meta=LEAD, sheet={"same_columns": True})
    wb = book(deal(view), tmp_path)
    starts = {name: next(c.column for c in wb[name][1] if c.value == "Constant")
              for name in ("Inputs", "Series", "Timing")}
    assert len(set(starts.values())) == 1


def test_a_sheet_of_constants_heads_its_columns_without_periods(tmp_path):
    view = mo.ExcelView(meta=LEAD, sheet={"same_columns": True, "title": "upper"})
    m = deal(view)
    m.ti.default_header = None                   # constants: no time rows on top
    wb = book(m, tmp_path)
    inputs, series = wb["Inputs"], wb["Series"]
    captions = ["No", "Line item", "Unit", "Total", "Constant"]
    assert [inputs.cell(2, c).value for c in range(1, 6)] == captions
    # The same words in the same places as on a sheet with periods.
    assert [series.cell(2, c).value for c in range(1, 6)] == captions
    # …and no periods: nothing right of the constants column.
    assert all(inputs.cell(2, c).value is None for c in range(6, 10))
    got = cells(inputs)
    first = min(int(ref[1:]) for ref, v in got.items() if v == "Model start")
    assert first > 2                             # the data stands under the captions
    assert got[f"E{first}"] is not None          # in the constants column
    assert inputs.freeze_panes.endswith("3")     # the captions stay in view
    assert inputs["A2"].font.b


def test_without_same_columns_a_sheet_of_constants_keeps_row_1_for_data(tmp_path):
    m = deal(mo.ExcelView(meta=LEAD))
    m.ti.default_header = None
    got = cells(book(m, tmp_path)["Inputs"])
    assert "Line item" not in got.values()
    assert min(int(ref[1:]) for ref, v in got.items() if v == "Model start") == 1


def test_a_scalar_header_row_sits_in_the_constant_column(tmp_path):
    m = deal(mo.ExcelView(meta=LEAD))
    with m:
        m.timing.flags.status = mo.IF(m.ti.rate > 0, "OK", "ERROR").set_display_name("Status")
    m.td.default_header = [m.timing.ruler.end, m.timing.flags.status]
    ws = book(m, tmp_path)["Series"]
    got = cells(ws)
    assert got["E1"] == "Constant"
    status = next(r for r in range(1, 10) if got.get(f"B{r}") == "Status")
    assert str(got[f"E{status}"]).startswith("=Timing!")


def test_the_anchor_can_repeat_its_own_rows_at_its_top(tmp_path):
    view = mo.ExcelView(timeline={"repeat_on_anchor": True}, meta={"unit": True})
    ws = book(deal(view), tmp_path)["Timing"]
    got = cells(ws)                              # label | unit | 2026 …
    labels = [got.get(f"A{r}") for r in range(2, 5)]
    assert labels == ["Period number", "Period start", "Period end"]
    own = {got.get(f"A{r}"): r for r in range(5, 20)}
    assert got["C4"] == f"=C{own['Period end']}"     # a reference to the row below
    # A formula on the anchor still reads the row itself, not its mirror.
    assert f"C{own['Period number']}" in got[f"C{own['First period flag']}"]
    assert ws.freeze_panes == "C5"               # under the repeated rows


# ══════════════════════════════════════════════════════════════════════
# Title row and the opening column
# ══════════════════════════════════════════════════════════════════════


def test_a_title_row_heads_every_sheet(tmp_path):
    view = mo.ExcelView(meta={**LEAD, "fields": [("description", "Note")]},
                        sheet={"title": "upper"},
                        bands={"title_note": {"italic": True, "font_color": GREY}})
    m = deal(view)
    m.ti.set_description("Constant inputs")
    ws = book(m, tmp_path)["Inputs"]
    assert ws["B1"].value == "INPUTS" and ws["B1"].font.b
    assert ws["D1"].value == "Constant inputs" and ws["D1"].font.i     # the Note column
    assert ws["A2"].value == "No" and ws["B2"].value == "Line item"    # header one row down


def test_the_opening_column_holds_the_day_before_the_start(tmp_path):
    view = mo.ExcelView(meta=LEAD, timeline={"opening": "Pre-start"},
                        sheet={"same_columns": True})
    wb = book(deal(view), tmp_path)
    ws = wb["Timing"]
    got = cells(ws)
    # No | Line item | Unit | Total | Constant | Pre-start | 2026 …
    assert got["E1"] == "Constant" and got["F1"] == "Pre-start"
    rows = {got.get(f"B{r}"): r for r in range(2, 20)}
    idx, start, end = rows["Period number"], rows["Period start"], rows["Period end"]
    assert got[f"F{end}"] == "=Inputs!E6-1"        # the day before the model start
    assert got[f"F{idx}"] == 0
    # One formula across the row, the first period included.
    assert got[f"G{start}"] == f"=F{end}+1" and got[f"H{start}"] == f"=G{end}+1"
    assert got[f"G{idx}"] == f"=F{idx}+1"
    assert f"F{start}" not in got                  # a start has no opening value
    series = cells(wb["Series"])
    mirror = next(r for r in range(1, 10) if series.get(f"B{r}") == "Period end")
    assert series[f"F{mirror}"] == f"=Timing!F{end}"
    assert ws.freeze_panes.startswith("G")         # the opening column stays in view


# ══════════════════════════════════════════════════════════════════════
# Full-width rows, widths, gridlines
# ══════════════════════════════════════════════════════════════════════


def test_a_full_row_spans_the_sheet(tmp_path):
    view = mo.ExcelView(meta=LEAD, bands={"section": {"bg": "#D9E2F3"}},
                        formats={"number": "#,##0.0"},
                        sheet={"full_rows": True})
    m = deal(view)
    m.td.volume.set_style(bold=True, border_top=True)
    wb = book(m, tmp_path)
    ws = wb["Series"]
    got = cells(ws)
    row = next(r for r in range(1, 10) if got.get(f"B{r}") == "Volume")
    for col in range(1, ws.max_column + 1):
        assert ws.cell(row, col).border.top.style == "thin"
    total = ws.cell(row, 4)
    assert total.value.startswith("=SUM(") and total.font.b
    assert total.number_format == "#,##0.0"
    timing = wb["Timing"]
    band = next(r for r in range(1, 10) if timing.cell(r, 2).value == "Flags")
    assert timing.cell(band, 1).fill.fgColor.rgb == "FFD9E2F3"     # from the first column


def test_widths_by_role_and_no_gridlines(tmp_path):
    view = mo.ExcelView(meta=LEAD, sheet={
        "gridlines": False,
        "columns": [("label", {"width": 40}), ("unit", {"width": 9, "font_color": GREY}),
                    ("period", {"width": 12})]})
    ws = book(deal(view), tmp_path)["Series"]
    assert ws.sheet_view.showGridLines is False
    assert ws.column_dimensions["B"].width == 40 and ws.column_dimensions["C"].width == 9
    first = column(ws, "Total") + 1              # no constants on this sheet
    for col in range(first, first + 3):
        assert ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width == 12
    unit = next(c for c in ws["C"] if c.value == "date")
    assert unit.font.color.rgb == "FF7F7F7F"


# ══════════════════════════════════════════════════════════════════════
# The house view
# ══════════════════════════════════════════════════════════════════════


def test_a_house_view_dresses_a_model_that_declares_nothing(tmp_path):
    house = mo.ExcelView(font={"name": "Arial", "size": 10}, sheet={"gridlines": False})
    with mo.house_view(house):
        wb = book(deal(), tmp_path)
    assert wb["Timing"]["A1"].font.name == "Arial"
    assert wb["Timing"].sheet_view.showGridLines is False
    # Outside the block the floor is back.
    assert resolve_excel_view(None).sheet_gridlines is None


def test_the_model_s_own_view_wins_over_the_house(tmp_path):
    house = mo.ExcelView(font={"name": "Arial", "size": 10}, meta=LEAD)
    team = mo.ExcelView(sheet={"gridlines": False})
    m = deal(mo.ExcelView(font={"name": "Georgia", "size": 12}))
    with mo.house_view(house, team):
        view = resolve_excel_view(m.ti.rate)
    assert view.base_font == {"name": "Georgia", "size": 12}     # the model
    assert view.sheet_gridlines is False                          # the team
    assert view.meta_fields == (("row_sum", "Total"),)            # the house, kept


def test_a_nested_view_without_fields_keeps_its_parent_s():
    m = deal(mo.ExcelView(meta=LEAD))
    m.td.default_excel_view = mo.ExcelView(font={"name": "Arial"})
    assert resolve_excel_view(m.td.volume).meta_fields == (("row_sum", "Total"),)


def test_house_view_refuses_what_is_not_a_view():
    with pytest.raises(TypeError, match="ExcelView"):
        with mo.house_view({"font": {"name": "Arial"}}):
            pass


def test_a_model_names_the_rows_every_sheet_opens_on(tmp_path):
    """A header LIST on the model: every sheet shows those rows; the
    section of its first time row still anchors the time."""
    m = deal(mo.ExcelView(timeline={"repeat_on_anchor": True}))
    m.default_header = [m.timing.ruler.end, m.timing.flags.first]
    wb = book(m, tmp_path)
    for name in ("Inputs", "Series", "Timing"):
        got = cells(wb[name])
        assert [got.get(f"A{r}") for r in (2, 3)] == ["Period end", "First period flag"]
    timing = cells(wb["Timing"])
    rows = {timing.get(f"A{r}"): r for r in range(4, 20)}
    # The ruler still chains off the start cell: the anchor is its section.
    assert timing[f"C{rows['Period start']}"] == f"=B{rows['Period end (date)']}+1"


# ══════════════════════════════════════════════════════════════════════
# What a line is: its look, its number, its total
# ══════════════════════════════════════════════════════════════════════


def _linked(view):
    m = deal(view)
    with m:
        m.timing.dates = mo.MultiVariable("Key dates")
        with m.timing.dates as d:
            d.start = m.ti.start.copy().set_display_name("Model start")
            d.close = mo.EDATE(m.timing.dates.start, 6).set_display_name("Financial close")
            d.cod = mo.EDATE(m.timing.dates.close, 18).set_display_name(
                "Commercial operation").set_style(result=True)
    return m


def test_only_a_marked_line_wears_the_result_look(tmp_path):
    view = mo.ExcelView(meta=LEAD, bands={"result": {"bold": True, "border_top": True},
                                          "link": {"font_color": GREY}})
    ws = book(_linked(view), tmp_path)["Timing"]
    got = cells(ws)
    rows = {got.get(f"B{r}"): r for r in range(2, 30)}
    assert not ws[f"B{rows['Financial close']}"].font.b          # computed, not marked
    assert ws[f"B{rows['Commercial operation']}"].font.b          # marked
    assert ws[f"B{rows['Commercial operation']}"].border.top.style == "thin"
    assert ws[f"B{rows['Model start']}"].font.color.rgb == "FF7F7F7F"   # a link


def test_a_link_to_a_result_is_no_result_of_its_own(tmp_path):
    # The next block opens by reading the result: the link carries the
    # source's excel_props, mark included, but it is not this block's result.
    view = mo.ExcelView(meta=LEAD, bands={"result": {"bold": True, "border_top": True}})
    m = _linked(view)
    with m:
        m.timing.use = mo.MultiVariable("Uses")
        with m.timing.use as u:
            u.cod = m.timing.dates.cod.copy().set_display_name("COD (link)")
    ws = book(m, tmp_path)["Timing"]
    got = cells(ws)
    rows = {got.get(f"B{r}"): r for r in range(2, 40)}
    assert ws[f"B{rows['Commercial operation']}"].font.b
    assert not ws[f"B{rows['COD (link)']}"].font.b
    assert ws[f"B{rows['COD (link)']}"].border.top.style is None


def test_lines_are_numbered_by_section_and_links_keep_none(tmp_path):
    view = mo.ExcelView(meta={**LEAD, "numbering": "section"})
    got = cells(book(_linked(view), tmp_path)["Timing"])
    numbers = {got[f"B{r}"]: got.get(f"A{r}") for r in range(2, 30) if f"B{r}" in got}
    assert numbers["Period number"] == "1.1" and numbers["Period end"] == "1.3"
    assert numbers["First period flag"] == "2.1"
    assert numbers["Model start"] is None                         # a link has no number
    assert numbers["Financial close"] == "3.1" and numbers["Commercial operation"] == "3.2"


def test_row_totals_follow_what_a_line_is(tmp_path):
    # A line totals by its re-grain rule (``regrain=mo.up('sum')``); a date
    # line never does.
    view = mo.ExcelView(meta={**LEAD, "row_sum_for": ["sum"]})
    m = deal(view)
    m.td.volume.excel_props.pop("row_sum")
    m.td.volume._regrain = mo.up("sum")
    got = cells(book(m, tmp_path)["Series"])
    rows = {got.get(f"B{r}"): r for r in range(2, 30)}
    assert any(str(v).startswith("=SUM(") for k, v in got.items()
               if k[1:] == str(rows["Volume"]))                        # a sum
    timing = cells(book(m, tmp_path)["Timing"])
    trows = {timing.get(f"B{r}"): r for r in range(2, 30)}
    assert not any(str(v).startswith("=SUM(") for k, v in timing.items()
                   if k[1:] == str(trows["Period end"]))               # a date


def test_a_bar_rows_total_stays_a_number(tmp_path):
    """A row drawn as a bar (``'"X";;'`` marks any positive value) totals
    to a count; the total takes the row's format only when that format
    would show a number."""
    from modeleon.compile.excel.writer import format_shows_a_number

    assert format_shows_a_number('#,##0;[Red](#,##0);')
    assert format_shows_a_number('0.##;-0.##;')
    assert format_shows_a_number("General")
    assert format_shows_a_number('0%')
    assert not format_shows_a_number('"X";;')
    assert not format_shows_a_number('"yes";"no";"-"')

    m = mo.Model("m", default_grain="year", default_start="2026", default_periods=4)
    m.default_excel_view = mo.ExcelView(meta=LEAD)
    m.s = mo.MultiVariable("S")
    with m.s as s:
        s.active = mo.Variable([1, 1, 0, 1], display_name="Active",
                               excel_props={"number_format": '"X";;', "row_sum": True})
        s.cash = mo.Variable([10, 20, 30, 40], display_name="Cash",
                             excel_props={"number_format": "#,##0", "row_sum": True})
    ws = book(m, tmp_path)["S"]
    label, total = column(ws, "Line item"), column(ws, "Total")
    totals = {ws.cell(r, label).value: ws.cell(r, total) for r in range(1, ws.max_row + 1)}
    assert totals["Active"].value.startswith("=SUM(")
    assert totals["Active"].number_format == "General"
    assert totals["Cash"].number_format == "#,##0"
