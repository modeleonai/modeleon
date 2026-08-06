# SPDX-License-Identifier: Apache-2.0
"""ExcelView value object + the declared-once cascade resolver."""

import pytest
from openpyxl import load_workbook

import modeleon as mo
from modeleon.compile.excel.view import ExcelView, resolve_excel_view


class TestExcelViewValueObject:
    def test_exported_on_namespace(self):
        assert mo.ExcelView is ExcelView

    def test_defaults_all_none(self):
        snap = ExcelView().snapshot()
        assert snap.base_font is None
        assert snap.band_styles is None
        assert snap.format_by_type is None

    def test_engine_default_is_all_none(self):
        # The byte-safe floor: every field None so the writer keeps its defaults.
        assert ExcelView.engine_default() == ExcelView()

    def test_extend_overrides_named_fields_only(self):
        base = ExcelView(font={"name": "Arial Narrow", "size": 10})
        board = base.extend(cell_types={"on": True}).snapshot()
        assert board.base_font == {"name": "Arial Narrow", "size": 10}  # inherited
        assert board.format_by_type is True                            # overridden
        assert base.snapshot().format_by_type is None                  # original untouched

    def test_extend_rejects_unknown_field(self):
        with pytest.raises(ValueError, match="unknown group"):
            ExcelView().extend(bogus="STRUCTURE")  # not a known sub-view

    def test_is_a_model(self):
        # ExcelView is a model of sub-views: each group is a child sub-model with
        # Variable leaves, navigable like any tree.
        v = ExcelView(timeline={"label_format": "iso"}, font={"name": "Calibri"})
        assert isinstance(v, mo.MultiVariable)
        assert isinstance(v.timeline, mo.MultiVariable)              # sub-view
        assert v.timeline.label_format.value == "iso"
        assert isinstance(v.font, mo.MultiVariable)                 # sub-view
        assert v.font.name.value == "Calibri"

    def test_per_model_view_is_a_slot_not_content(self, tmp_path):
        # Assigning a view to ``default_excel_view`` must NOT pull it into the
        # content tree (it's an MV now) — otherwise it would emit a spurious tab.
        m = mo.Model("m")
        m.default_start, m.default_grain = "2024-01", "month"
        m.default_excel_view = ExcelView(timeline={"label_format": "iso"})
        m.rev = mo.Variable([1, 2, 3], display_name="Rev")
        assert "default_excel_view" not in m._components       # a slot, not content
        m.to_excel(tmp_path / "o.xlsx")
        assert load_workbook(tmp_path / "o.xlsx").sheetnames == ["M"]   # no view tab
        # …and the per-model view is still honored.
        assert resolve_excel_view(m.rev).time_label_format == "iso"


class TestGlobalDefaultView:
    """``mo.default_excel_view`` is the inspectable, globally-overridable house
    default that ``resolve_excel_view`` falls back to and fills partial views
    from."""

    @pytest.fixture(autouse=True)
    def _restore_global(self):
        import modeleon
        saved = modeleon.default_excel_view
        yield
        modeleon.default_excel_view = saved

    def test_house_default_is_populated(self):
        import modeleon
        assert isinstance(modeleon.default_excel_view, ExcelView)
        snap = modeleon.default_excel_view.snapshot()
        assert snap.time_label_format == "finance"
        assert snap.type_colors == {"input": "#0563C1", "reference": "#2E7D32"}

    def test_global_override_changes_resolution(self):
        import modeleon
        m = mo.Model("m")
        m.x = mo.Variable(1)
        modeleon.default_excel_view = ExcelView(timeline={"label_format": "iso"})
        assert resolve_excel_view(m.x).time_label_format == "iso"

    def test_global_override_reaches_partial_per_model_views(self):
        import modeleon
        m = mo.Model("m")
        m.default_excel_view = ExcelView(font={"name": "Calibri"})
        m.y = mo.Variable(1)
        modeleon.default_excel_view = ExcelView(timeline={"label_format": "compact"})
        r = resolve_excel_view(m.y)
        assert r.base_font == {"name": "Calibri"}   # own field kept
        assert r.time_label_format == "compact"     # inherited from the new global


class TestResolveCascade:
    def test_no_view_anywhere_returns_house_default(self):
        # With no pointer up the chain, the resolver returns the global house
        # default (``mo.default_excel_view``), not the all-None floor.
        m = mo.Model("m")
        m.x = mo.Variable(1)
        resolved = resolve_excel_view(m.x)
        assert resolved.time_label_format == "finance"
        assert resolved.type_colors == {"input": "#0563C1", "reference": "#2E7D32"}

    def test_view_on_root_cascades_to_deep_leaf(self):
        m = mo.Model("m")
        m.default_excel_view = ExcelView(font={"name": "Arial Narrow", "size": 10})
        m.section = mo.MultiVariable("Section")
        m.section.x = mo.Variable(1)
        resolved = resolve_excel_view(m.section.x)
        assert resolved.base_font == {"name": "Arial Narrow", "size": 10}

    def test_nearer_pointer_wins_whole_view(self):
        m = mo.Model("m")
        m.default_excel_view = ExcelView(font={"name": "Arial Narrow"})
        m.section = mo.MultiVariable("Section")
        m.section.default_excel_view = ExcelView(font={"name": "Calibri"})
        m.section.x = mo.Variable(1)
        m.y = mo.Variable(2)
        # the leaf under section sees the nearer view; a sibling sees the root view
        assert resolve_excel_view(m.section.x).base_font == {"name": "Calibri"}
        assert resolve_excel_view(m.y).base_font == {"name": "Arial Narrow"}

    def test_cascade_is_per_field_through_intermediate_levels(self):
        # The nearest level wins per FIELD; fields it leaves unset fall back to
        # an INTERMEDIATE ancestor (not skipped), then the global default floor.
        m = mo.Model("m")
        m.default_excel_view = ExcelView(orient="down", timeline={"label_format": "iso"})
        m.sub = mo.MultiVariable("sub")
        m.sub.default_excel_view = ExcelView(timeline={"label_format": "compact"})
        m.sub.x = mo.Variable(1)
        r = resolve_excel_view(m.sub.x)
        assert r.time_label_format == "compact"   # nearest (sub) wins
        assert r.orient == "down"                 # cascades from the model level
        assert r.type_colors == {                 # global default is the floor
            "input": "#0563C1",
            "reference": "#2E7D32",
        }

    def test_pointer_can_reference_a_named_view_child(self):
        # The authoring shape: name a view, then point default at it. The
        # resolved view keeps its own fields and fills the rest from the house
        # default (so it's a filled copy, not the identical object).
        m = mo.Model("m")
        m.fin_view = ExcelView(font={"name": "Arial Narrow", "size": 10})
        m.default_excel_view = m.fin_view
        m.x = mo.Variable(1)
        resolved = resolve_excel_view(m.x)
        assert resolved.base_font == {"name": "Arial Narrow", "size": 10}  # own field kept
        assert resolved.time_label_format == "finance"                    # filled from house default

    def test_extend_for_inherit_and_tweak(self):
        m = mo.Model("m")
        m.base_view = ExcelView(font={"name": "Arial Narrow", "size": 10})
        m.default_excel_view = m.base_view
        m.pnl = mo.MultiVariable("P&L")
        m.pnl.default_excel_view = m.base_view.extend(cell_types={"on": True})
        m.pnl.x = mo.Variable(1)
        r = resolve_excel_view(m.pnl.x)
        assert r.base_font == {"name": "Arial Narrow", "size": 10}  # inherited via extend
        assert r.format_by_type is True


class TestBaseFontWiring:
    """Slice 2: the resolved view's ``base_font`` reaches the emitted cells."""

    @staticmethod
    def _value_cell(ws):
        return next(c for row in ws.iter_rows() for c in row if c.value in (1, 2, 3))

    def test_base_font_applied_from_view(self, tmp_path):
        m = mo.Model("m")
        m.default_excel_view = mo.ExcelView(font={"name": "Arial Narrow", "size": 10})
        m.rev = mo.Variable([1, 2, 3], display_name="Revenue")
        m.to_excel(tmp_path / "out.xlsx")
        ws = load_workbook(tmp_path / "out.xlsx")[load_workbook(tmp_path / "out.xlsx").sheetnames[0]]
        cell = self._value_cell(ws)
        assert cell.font.name == "Arial Narrow"
        assert cell.font.size == 10

    def test_no_view_leaves_default_font(self, tmp_path):
        # Byte-safe floor: with no view, the font is untouched (openpyxl default).
        m = mo.Model("m")
        m.rev = mo.Variable([1, 2, 3], display_name="Revenue")
        m.to_excel(tmp_path / "out.xlsx")
        wb = load_workbook(tmp_path / "out.xlsx")
        cell = self._value_cell(wb[wb.sheetnames[0]])
        assert cell.font.name in (None, "Calibri")

    def test_explicit_variable_font_wins_over_view(self, tmp_path):
        m = mo.Model("m")
        m.default_excel_view = mo.ExcelView(font={"name": "Arial Narrow", "size": 10})
        m.rev = mo.Variable(
            [1, 2, 3], display_name="Revenue",
            excel_props={"font_family": "Times New Roman"},
        )
        m.to_excel(tmp_path / "out.xlsx")
        wb = load_workbook(tmp_path / "out.xlsx")
        cell = self._value_cell(wb[wb.sheetnames[0]])
        assert cell.font.name == "Times New Roman"   # explicit Variable prop wins
        assert cell.font.size == 10                  # size still inherited from the view

    def test_subsection_override_cascades_to_its_cells(self, tmp_path):
        m = mo.Model("m")
        m.default_excel_view = mo.ExcelView(font={"name": "Arial Narrow", "size": 10})
        m.detail = mo.MultiVariable("Detail", excel_props={"tab": True})
        m.detail.a = mo.Variable([1, 2, 3], display_name="A")
        m.detail.default_excel_view = mo.ExcelView(font={"name": "Calibri", "size": 9})
        m.to_excel(tmp_path / "out.xlsx")
        wb = load_workbook(tmp_path / "out.xlsx")
        ws = wb["Detail"]
        cell = self._value_cell(ws)
        assert cell.font.name == "Calibri"           # nearest view wins for these cells
        assert cell.font.size == 9


class TestFormatByTypeField:
    """Slice 3: ``format_by_type`` carried by the view (legacy prop is fallback)."""

    @staticmethod
    def _cell(ws, value):
        return next(c for row in ws.iter_rows() for c in row if c.value == value)

    def test_format_by_type_from_view_colours_inputs(self, tmp_path):
        m = mo.Model("m")
        m.default_excel_view = mo.ExcelView(cell_types={"on": True, "colors": {"input": "#B08C0A"}})
        m.price = mo.Variable(100, display_name="Price")   # literal → 'input'
        m.to_excel(tmp_path / "out.xlsx")
        wb = load_workbook(tmp_path / "out.xlsx")
        cell = self._cell(wb[wb.sheetnames[0]], 100)
        assert "B08C0A" in str(cell.font.color.rgb).upper()

    def test_no_setting_leaves_cells_uncoloured(self, tmp_path):
        m = mo.Model("m")
        m.price = mo.Variable(100, display_name="Price")
        m.to_excel(tmp_path / "out.xlsx")
        wb = load_workbook(tmp_path / "out.xlsx")
        cell = self._cell(wb[wb.sheetnames[0]], 100)
        rgb = cell.font.color.rgb if cell.font and cell.font.color else None
        assert "B08C0A" not in str(rgb).upper()

    def test_legacy_excel_props_format_by_type_still_works(self, tmp_path):
        # View field is None → falls back to the legacy owner→parent walk.
        m = mo.Model("m", excel_props={"format_by_type": {"input": "#B08C0A"}})
        m.price = mo.Variable(100, display_name="Price")
        m.to_excel(tmp_path / "out.xlsx")
        wb = load_workbook(tmp_path / "out.xlsx")
        cell = self._cell(wb[wb.sheetnames[0]], 100)
        assert "B08C0A" in str(cell.font.color.rgb).upper()


class TestBandStylesField:
    """Slice 3: section/subsection band styling carried by the view.

    Role derives from the MV's own props (a ``bg`` fill → ``section``, else
    ``subsection``); the view's ``band_styles[role]`` fills in UNDER the MV's
    own props (so existing models stay byte-identical).
    """

    @staticmethod
    def _header(ws, label):
        return next(c for row in ws.iter_rows() for c in row if c.value == label)

    def test_band_styles_from_view_style_a_bare_section(self, tmp_path):
        m = mo.Model("m")
        m.default_excel_view = mo.ExcelView(
            bands={"subsection": {"bold": True, "bg": "#E9BA0D"}}
        )
        m.tab = mo.MultiVariable("Tab", excel_props={"tab": True})
        m.tab.sec = mo.MultiVariable("Section")          # no bg → role 'subsection'
        m.tab.sec.x = mo.Variable(1, display_name="X")
        m.to_excel(tmp_path / "out.xlsx")
        cell = self._header(load_workbook(tmp_path / "out.xlsx")["Tab"], "Section")
        assert cell.font.bold is True
        assert "E9BA0D" in str(cell.fill.fgColor.rgb).upper()

    def test_mv_own_props_win_over_band_styles(self, tmp_path):
        m = mo.Model("m")
        m.default_excel_view = mo.ExcelView(bands={"section": {"bg": "#00FF00"}})
        m.tab = mo.MultiVariable("Tab", excel_props={"tab": True})
        m.tab.sec = mo.MultiVariable("Section", excel_props={"bg": "#FF0000"})  # own bg wins
        m.tab.sec.x = mo.Variable(1, display_name="X")
        m.to_excel(tmp_path / "out.xlsx")
        cell = self._header(load_workbook(tmp_path / "out.xlsx")["Tab"], "Section")
        assert "FF0000" in str(cell.fill.fgColor.rgb).upper()

    def test_no_view_band_unchanged(self, tmp_path):
        m = mo.Model("m")
        m.tab = mo.MultiVariable("Tab", excel_props={"tab": True})
        m.tab.sec = mo.MultiVariable("Section", excel_props={"bg": "#E9BA0D", "bold": True})
        m.tab.sec.x = mo.Variable(1, display_name="X")
        m.to_excel(tmp_path / "out.xlsx")
        cell = self._header(load_workbook(tmp_path / "out.xlsx")["Tab"], "Section")
        assert "E9BA0D" in str(cell.fill.fgColor.rgb).upper()
        assert cell.font.bold is True


class TestOrientTranspose:
    """Slice 4 (first pivot capability): ``orient='down'`` transposes a flat
    sheet — Variables become columns, periods run down, formulas follow."""

    @staticmethod
    def _ws(tmp_path, m):
        m.to_excel(tmp_path / "out.xlsx")
        wb = load_workbook(tmp_path / "out.xlsx")
        return wb[wb.sheetnames[0]]

    def test_down_transposes_flat_model(self, tmp_path):
        m = mo.Model("m")
        m.default_excel_view = mo.ExcelView(orient="down")
        m.price = mo.Variable([10, 11, 12], display_name="Price")
        m.volume = mo.Variable([100, 120, 140], display_name="Volume")
        ws = self._ws(tmp_path, m)
        # labels become column headers; values run DOWN
        assert ws["A1"].value == "Price"
        assert ws["B1"].value == "Volume"
        assert [ws[f"A{r}"].value for r in (2, 3, 4)] == [10, 11, 12]
        assert [ws[f"B{r}"].value for r in (2, 3, 4)] == [100, 120, 140]

    def test_down_formula_follows_transpose(self, tmp_path):
        m = mo.Model("m")
        m.default_excel_view = mo.ExcelView(orient="down")
        m.price = mo.Variable([10, 11], display_name="Price")
        m.volume = mo.Variable([100, 120], display_name="Volume")
        m.revenue = (m.price * m.volume).set_display_name("Revenue")
        ws = self._ws(tmp_path, m)
        assert ws["C1"].value == "Revenue"
        assert ws["C2"].value == "=A2 * B2"   # references the transposed cells
        assert ws["C3"].value == "=A3 * B3"

    def test_down_recurrence_references_cell_above(self, tmp_path):
        m = mo.Model("m")
        m.default_excel_view = mo.ExcelView(orient="down")
        m.bal = mo.recurrence(100, "{prev} + 10", periods=3, display_name="Balance")
        ws = self._ws(tmp_path, m)
        assert ws["A1"].value == "Balance"
        assert str(ws["A2"].value) in ("100", "=100")   # the seed (period 0)
        assert "A2" in str(ws["A3"].value)              # {prev} = the cell above
        assert "A3" in str(ws["A4"].value)

    def test_default_across_unchanged(self, tmp_path):
        m = mo.Model("m")
        m.price = mo.Variable([10, 11, 12], display_name="Price")
        ws = self._ws(tmp_path, m)
        # label col A, values across B/C/D (today's layout)
        assert ws["A1"].value == "Price"
        assert [ws[f"{c}1"].value for c in ("B", "C", "D")] == [10, 11, 12]

    def test_down_nested_falls_back_to_across(self, tmp_path):
        # A sheet with a sub-section isn't flat → orient='down' falls back safely.
        m = mo.Model("m")
        m.default_excel_view = mo.ExcelView(orient="down")
        m.sec = mo.MultiVariable("Section")
        m.sec.x = mo.Variable([1, 2, 3], display_name="X")
        m.to_excel(tmp_path / "out.xlsx")
        wb = load_workbook(tmp_path / "out.xlsx")
        found = any(
            c.value == "X" for ws in wb.worksheets for row in ws.iter_rows() for c in row
        )
        assert found   # emits without error; nested layout preserved


class TestTimelineHeader:
    """A period-label row at the top of each sheet that declares a time window.

    Default-on but window-gated: a sheet with no resolved window (no
    ``default_grain`` up the tree) renders byte-identically to before. Format
    is configurable via ``ExcelView.time_label_format``.
    """

    def _ws(self, tmp_path, m):
        m.to_excel(tmp_path / "o.xlsx")
        return load_workbook(tmp_path / "o.xlsx").worksheets[0]

    def test_finance_format_is_the_default(self, tmp_path):
        m = mo.Model("plan")
        m.default_start = "2024-01"
        m.default_grain = "month"
        m.rev = mo.Variable([10, 11, 12], display_name="Rev")
        ws = self._ws(tmp_path, m)
        assert ws["B1"].value == "Jan 2024"        # header row
        assert ws["C1"].value == "Feb 2024"
        assert ws["A2"].value == "Rev"             # data shifted below the header
        assert ws["B2"].value == 10
        assert ws.freeze_panes == "B2"             # header + label col pinned

    def test_iso_format_via_view(self, tmp_path):
        m = mo.Model("plan")
        m.default_start = "2024-01"
        m.default_grain = "month"
        m.default_excel_view = mo.ExcelView(timeline={"label_format": "iso"})
        m.rev = mo.Variable([10, 11, 12])
        assert self._ws(tmp_path, m)["B1"].value == "2024-01"

    def test_compact_format_via_view(self, tmp_path):
        m = mo.Model("plan")
        m.default_start = "2024-01"
        m.default_grain = "month"
        m.default_excel_view = mo.ExcelView(timeline={"label_format": "compact"})
        m.rev = mo.Variable([10, 11, 12])
        assert self._ws(tmp_path, m)["B1"].value == "Jan-24"

    def test_quarter_and_year_grains(self, tmp_path):
        mq = mo.Model("q")
        mq.default_start = "2024-Q1"
        mq.default_grain = "quarter"
        mq.rev = mo.Variable([1, 2, 3])
        wsq = self._ws(tmp_path, mq)
        assert [wsq[f"{c}1"].value for c in ("B", "C", "D")] == [
            "Q1 2024", "Q2 2024", "Q3 2024"
        ]
        my = mo.Model("y")
        my.default_start = "2024"
        my.default_grain = "year"
        my.rev = mo.Variable([1, 2])
        wsy = self._ws(tmp_path, my)
        assert [wsy[f"{c}1"].value for c in ("B", "C")] == ["2024", "2025"]

    def test_serializable_by_sheet_projection(self):
        # ``time_headers_by_sheet`` is the name-keyed projection external
        # consumers read (the pro bridge → IDE wire); the id()-keyed map
        # stays the writer's (clone-collision-safe) source. Same tuples.
        from modeleon.compile.excel.layout import LayoutEngine

        m = mo.Model("plan")
        m.default_start = "2024-01"
        m.default_grain = "month"
        m.rev = mo.Variable([10, 11, 12])
        engine = LayoutEngine([m])
        engine.compute_addresses()
        assert engine.time_headers_by_sheet == {
            "Plan": ("2024-01", "month", "finance")
        }
        assert list(engine.time_headers.values()) == [
            ("2024-01", "month", "finance")
        ]

    def test_time_header_false_suppresses(self, tmp_path):
        m = mo.Model("plan")
        m.default_start = "2024-01"
        m.default_grain = "month"
        m.default_excel_view = mo.ExcelView(timeline={"header": False})
        m.rev = mo.Variable([10, 11, 12], display_name="Rev")
        ws = self._ws(tmp_path, m)
        assert ws["A1"].value == "Rev"             # data at row 1, no header
        assert ws["B1"].value == 10
        assert ws.freeze_panes == "B1"

    def test_no_window_is_header_less_and_byte_safe(self, tmp_path):
        # No default_grain anywhere → no window → no header, exactly as before.
        m = mo.Model("plan")
        m.rev = mo.Variable([10, 11, 12], display_name="Rev")
        ws = self._ws(tmp_path, m)
        assert ws["A1"].value == "Rev"
        assert ws["B1"].value == 10
        assert ws.freeze_panes == "B1"

    def test_invalid_format_raises(self, tmp_path):
        m = mo.Model("plan")
        m.default_start = "2024-01"
        m.default_grain = "month"
        m.default_excel_view = mo.ExcelView(timeline={"label_format": "bogus"})
        m.rev = mo.Variable([10, 11, 12])
        with pytest.raises(ValueError, match="time_label_format"):
            m.to_excel(tmp_path / "o.xlsx")


class TestReprHonorsExcelView:
    """The notebook HTML repr mirrors the Excel view's styling — base_font,
    ``format_by_type`` cell-type colours (reusing the writer's decision
    functions), and the variable's own ``excel_props`` — so ``display(model)``
    previews what ``to_excel`` produces. Byte-safe when no view is set.
    """

    def test_format_by_type_colours_inputs(self):
        m = mo.Model("m")
        m.default_excel_view = mo.ExcelView(cell_types={"on": True})
        m.price = mo.Variable([10, 11, 12], display_name="Price")  # input cell
        assert "color:#0563C1" in m._repr_html_()  # input -> the writer's blue

    def test_base_font_family_in_repr(self):
        m = mo.Model("m")
        m.default_excel_view = mo.ExcelView(font={"name": "Arial Narrow"})
        m.x = mo.Variable([1, 2, 3], display_name="X")
        assert "font-family:'Arial Narrow'" in m._repr_html_()

    def test_own_props_styling_in_repr(self):
        m = mo.Model("m")
        m.flag = mo.Variable([1, 2, 3], display_name="Flag",
                             excel_props={"bold": True, "bg": "#E9BA0D"})
        h = m._repr_html_()
        assert "font-weight:600" in h
        assert "background:#E9BA0D" in h

    def test_plain_model_gets_house_default_only(self):
        # A plain model inherits the house default (the explicit font) but adds
        # no type colour (cell_types is off) and no own props of its own.
        from modeleon.display.html import _view_cell_css
        m = mo.Model("m")
        m.x = mo.Variable([1, 2, 3], display_name="X")
        css = _view_cell_css(m.x, None, None)
        assert "font-family:'Calibri'" in css   # the house default font applies
        assert "color:#" not in css             # no type colour (cell_types off)
        assert "font-weight" not in css         # no own props


class TestExcelViewRepr:
    def test_repr_html_is_one_flat_panel_not_per_group_tabs(self):
        # ExcelView is a MultiVariable, so the generic repr would tab each group
        # (font / bands / cell_types / timeline). The custom repr collapses it to
        # ONE flat card — a single <table> (the generic model repr emits one per
        # tab) summarising the effective settings.
        html = mo.default_excel_view._repr_html_()
        assert html.count("<table") == 1
        for token in ("ExcelView", "orient", "across", "Calibri", "cell_types", "finance"):
            assert token in html

    def test_repr_html_shows_unset_groups_as_dash(self):
        # A bare view (no groups set) renders each setting row as an em-dash.
        html = ExcelView()._repr_html_()
        assert "<table" in html and "ExcelView" in html
        assert "—" in html
