# SPDX-License-Identifier: Apache-2.0
"""The notebook repr draws a keyed line's caption row as the book does.

A keyed Variable (``mo.Variable({"FY2023": 400, "FY2024": 450})``) stands
under a caption row of its keys in the workbook, off the period header.
The repr lays out through the same engine, and used to skip that row:
the row numbers jumped from 2 to 4 and the keyed values sat under the
period labels. It now paints the row — its number, the keys over the
cells — and consecutive lines with the same keys share it, as on the
sheet.
"""

import re

import modeleon as mo


def _rows(html):
    out = []
    for tr in re.findall(r"<tr>(.*?)</tr>", html, re.S):
        cells = re.findall(r"<t[hd][^>]*>(.*?)</t[hd]>", tr, re.S)
        out.append([re.sub(r"<[^>]+>", "", c).strip() for c in cells])
    return out


def _model():
    m = mo.Model("m", display_name="M", default_grain="year", default_start="2026",
                 default_periods=3)
    with m:
        m.inputs = mo.MultiVariable("Inputs", excel_props={"tab": True})
        with m.inputs as i:
            i.revenue = mo.Variable([1, 2, 3], display_name="Revenue")
            i.hist = mo.Variable({"FY2023": 400, "FY2024": 450}, display_name="Hist")
            i.hist2 = mo.Variable({"FY2023": 40, "FY2024": 45}, display_name="Hist 2")
            i.growth = mo.Variable(0.1, display_name="Growth")
    return m


def test_the_keys_stand_in_their_own_row_over_the_values():
    rows = _rows(_model()._repr_html_())
    assert rows[1:] == [
        ["1", "", "2026", "2027", "2028"],
        ["2", "Revenue", "1", "2", "3"],
        ["3", "", "FY2023", "FY2024", ""],
        ["4", "Hist", "400", "450", ""],
        ["5", "Hist 2", "40", "45", ""],
        ["6", "Growth", "0.10", "", ""],
    ]


def test_the_sheet_repr_draws_the_same_rows():
    assert _rows(_model().inputs._repr_html_())[1:4] == [
        ["1", "", "2026", "2027", "2028"],
        ["2", "Revenue", "1", "2", "3"],
        ["3", "", "FY2023", "FY2024", ""],
    ]


def test_other_keys_get_a_row_of_their_own():
    m = _model()
    with m.inputs as i:
        i.three = mo.Variable({"FY2022": 1, "FY2023": 2, "FY2024": 3}, display_name="Three")
    rows = _rows(m._repr_html_())
    assert rows[6] == ["6", "Growth", "0.10", "", ""]
    assert rows[7] == ["7", "", "FY2022", "FY2023", "FY2024"]
    assert rows[8] == ["8", "Three", "1", "2", "3"]


def test_a_keyed_sheet_without_a_window_starts_with_its_caption_row():
    m = mo.Model("m", display_name="M")
    with m:
        m.hist = mo.MultiVariable("Hist", excel_props={"tab": True})
        m.hist.rev = mo.Variable({"FY2023": 400, "FY2024": 450}, display_name="Revenue")
    rows = _rows(m._repr_html_())
    assert rows[1:] == [
        ["1", "", "FY2023", "FY2024"],
        ["2", "Revenue", "400", "450"],
    ]
