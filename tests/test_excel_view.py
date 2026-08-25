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

    def test_band_role_is_depth_first_level_is_a_section(self, tmp_path):
        # The old rule keyed the role on the MV's OWN fill, which made
        # ``bands.section.bg`` unreachable: whoever had a fill kept it,
        # whoever hadn't was a «subsection» — the view could never
        # paint an unstyled раздел (reported live: «I. Доходы» plain
        # under a beige-band view). Role is DEPTH under the tab.
        m = mo.Model("m")
        m.default_excel_view = mo.ExcelView(
            bands={"section": {"bold": True, "bg": "#EFEDE8"},
                   "subsection": {"bold": True}}
        )
        m.tab = mo.MultiVariable("Tab", excel_props={"tab": True})
        m.tab.sec = mo.MultiVariable("Section")          # depth 0 → section
        m.tab.sec.deep = mo.MultiVariable("Deep")        # depth 1 → subsection
        m.tab.sec.deep.x = mo.Variable(1, display_name="X")
        m.to_excel(tmp_path / "out.xlsx")
        ws = load_workbook(tmp_path / "out.xlsx")["Tab"]
        sec = self._header(ws, "Section")
        assert sec.font.bold is True
        assert "EFEDE8" in str(sec.fill.fgColor.rgb).upper()
        deep = next(
            c for row in ws.iter_rows() for c in row
            if c.value and str(c.value).strip() == "Deep"
        )
        assert deep.font.bold is True
        assert "EFEDE8" not in str(deep.fill.fgColor.rgb).upper()

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


class TestNamedViews:
    """A NAMED ExcelView adopted onto a model is an inert recipe —
    selectable presentation metadata, never content."""

    def test_display_name_is_the_first_positional(self):
        v = ExcelView("Quarterly board", grain="quarter", tracks="rows")
        assert v.display_name == "Quarterly board"

    def test_unnamed_view_keeps_the_class_label(self):
        assert ExcelView().display_name == "ExcelView"

    def test_style_kwargs_are_keyword_only(self):
        with pytest.raises(TypeError):
            ExcelView("Name", "down")  # orient must be spelled orient=

    def test_named_view_emits_no_rows_and_no_tabs(self, tmp_path):
        m = mo.Model("m")
        m.x = mo.Variable([1, 2, 3], display_name="X")
        m.board = ExcelView("Quarterly board", grain="quarter")
        path = tmp_path / "named_view.xlsx"
        m.to_excel(str(path))
        from openpyxl import load_workbook
        wb = load_workbook(str(path))
        assert wb.sheetnames == ["M"]
        labels = [c.value for row in wb["M"].iter_rows() for c in row if c.value]
        assert not any("board" in str(v).lower() for v in labels)
        assert not any("quarter" in str(v).lower() for v in labels)

    def test_named_view_does_not_shift_content_addresses(self):
        from modeleon.compile.excel.layout import LayoutEngine
        plain = mo.Model("m1")
        plain.x = mo.Variable([1, 2, 3], display_name="X")
        with_view = mo.Model("m2")
        with_view.x = mo.Variable([1, 2, 3], display_name="X")
        with_view.board = ExcelView("Board", tracks="rows")
        a1 = LayoutEngine([plain]).compute_addresses()["m1.x"]
        a2 = LayoutEngine([with_view]).compute_addresses()["m2.x"]
        assert a1.values == a2.values and a1.name == a2.name

    def test_named_view_stays_out_of_the_notebook_repr(self):
        m = mo.Model("m")
        m.x = mo.Variable([1, 2, 3], display_name="X")
        m.board = ExcelView("Quarterly board", grain="quarter")
        html = m._repr_html_()
        assert "Quarterly board" not in html
        assert "quarter" not in html

    def test_number_format_cascades_from_the_view(self, tmp_path):
        m = mo.Model("m")
        m.default_excel_view = ExcelView(formats={"number": "#,##0;[Red](#,##0)"})
        m.revenue = mo.Variable([1234567.0, -50000.0], display_name="Revenue")
        m.note = mo.Variable("text", display_name="Note")
        path = tmp_path / "fmt.xlsx"
        m.to_excel(str(path))
        from openpyxl import load_workbook
        ws = load_workbook(str(path))["M"]
        cells = {c.value: c for row in ws.iter_rows() for c in row}
        num_cell = next(c for row in ws.iter_rows() for c in row
                        if isinstance(c.value, (int, float)) and c.value == 1234567.0)
        assert num_cell.number_format == "#,##0;[Red](#,##0)"
        text_cell = next(c for row in ws.iter_rows() for c in row
                         if c.value == "text")
        assert text_cell.number_format in ("General",)

    def test_section_number_format_overrides_the_view_floor(self, tmp_path):
        m = mo.Model("m")
        m.default_excel_view = ExcelView(formats={"number": "#,##0"})
        m.rates = mo.MultiVariable(
            display_name="Rates", excel_props={"number_format": "0.0%"}
        )
        m.rates.vat = mo.Variable(0.12, display_name="VAT")
        m.money = mo.Variable(1000.0, display_name="Money")
        path = tmp_path / "fmt2.xlsx"
        m.to_excel(str(path))
        from openpyxl import load_workbook
        wb = load_workbook(str(path))
        cells = [c for ws in wb.worksheets for row in ws.iter_rows() for c in row]
        vat = next(c for c in cells if c.value == 0.12)
        money = next(c for c in cells if c.value == 1000.0)
        assert vat.number_format == "0.0%"
        assert money.number_format == "#,##0"

    def test_untracked_sections_ride_the_expansion_by_reference(self, tmp_path):
        # A tracked model's UNTRACKED sections must keep their literal
        # cells and labels through expand_tracked_tree — re-adoption
        # used to reference-clone them into self-referential formulas.
        m = mo.Model("m", tracks=mo.Tracks("plan", "fact"),
                     default_grain="month", default_start="2026-01",
                     default_periods=3)
        m.params = mo.MultiVariable(display_name="Params")
        m.params.rate = mo.Variable(0.12, display_name="Tax rate")
        m.pnl = mo.MultiVariable(display_name="PnL")
        m.pnl.revenue = mo.Variable(plan=[100.0] * 3, fact=[110.0, None, None])
        path = tmp_path / "ref.xlsx"
        m.to_excel(str(path))
        from openpyxl import load_workbook
        wb = load_workbook(str(path))
        cells = [c for ws in wb.worksheets for row in ws.iter_rows() for c in row]
        assert any(c.value == "Tax rate" for c in cells)
        assert any(c.value == 0.12 for c in cells)
        assert not any(isinstance(c.value, str) and c.value == f"={c.coordinate}"
                       for c in cells), "self-referential clone leaked"

    def test_slot_view_rides_the_expansion(self, tmp_path):
        m = mo.Model("m", tracks=mo.Tracks("plan", "fact"),
                     default_grain="month", default_start="2026-01",
                     default_periods=3)
        m.default_excel_view = ExcelView(formats={"number": "#,##0"})
        m.pnl = mo.MultiVariable(display_name="PnL")
        m.pnl.revenue = mo.Variable(plan=[100.0] * 3, fact=[110.0, None, None])
        path = tmp_path / "slot.xlsx"
        m.to_excel(str(path))
        from openpyxl import load_workbook
        wb = load_workbook(str(path))
        cells = [c for ws in wb.worksheets for row in ws.iter_rows() for c in row]
        rev = next(c for c in cells if c.value == 110.0)
        assert rev.number_format == "#,##0"

    def test_dense_nesting_drops_gaps_and_indents_labels(self, tmp_path):
        m = mo.Model("m")
        m.default_excel_view = ExcelView(nesting={"gap_rows": 0, "indent": 2})
        m.tab = mo.MultiVariable(display_name="Tab", excel_props={"tab": True})
        m.tab.a = mo.MultiVariable(display_name="Section A")
        m.tab.a.x = mo.Variable(1.0, display_name="X")
        m.tab.a.y = mo.Variable(2.0, display_name="Y")
        m.tab.b = mo.MultiVariable(display_name="Section B")
        m.tab.b.z = mo.Variable(3.0, display_name="Z")
        path = tmp_path / "dense.xlsx"
        m.to_excel(str(path))
        from openpyxl import load_workbook
        ws = load_workbook(str(path))["Tab"]
        col_a = [ws.cell(row=r, column=1).value for r in range(1, 8)]
        # No blank separator between Section A's tail and Section B.
        assert col_a[:5] == ["Section A", "  X", "  Y", "Section B", "  Z"]

    def test_default_keeps_the_classic_gap(self, tmp_path):
        m = mo.Model("m")
        m.tab = mo.MultiVariable(display_name="Tab", excel_props={"tab": True})
        m.tab.a = mo.MultiVariable(display_name="Section A")
        m.tab.a.x = mo.Variable(1.0, display_name="X")
        m.tab.b = mo.MultiVariable(display_name="Section B")
        m.tab.b.z = mo.Variable(3.0, display_name="Z")
        path = tmp_path / "classic.xlsx"
        m.to_excel(str(path))
        from openpyxl import load_workbook
        ws = load_workbook(str(path))["Tab"]
        col_a = [ws.cell(row=r, column=1).value for r in range(1, 6)]
        assert col_a == ["Section A", "X", None, "Section B", "Z"]

    def test_header_row_total_prints_on_the_section_header(self, tmp_path):
        # The financial-book form: «Взносы | 886 | 886» on the header
        # line, components under it.
        m = mo.Model("m", default_grain="month", default_start="2026-01",
                     default_periods=2)
        m.tab = mo.MultiVariable(display_name="Tab", excel_props={"tab": True})
        m.tab.взносы = mo.MultiVariable(display_name="1.2 · Взносы")
        m.tab.взносы.со = mo.Variable([5.0, 5.0], display_name="СО")
        m.tab.взносы.осмс = mo.Variable([3.0, 3.0], display_name="ОСМС")
        m.tab.взносы.итого = mo.Variable(
            m.tab.взносы.со + m.tab.взносы.осмс,
            excel_props={"header_row": True},
        )
        path = tmp_path / "hdr.xlsx"
        m.to_excel(str(path))
        from openpyxl import load_workbook
        ws = load_workbook(str(path))["Tab"]
        rows = {}
        for r in range(1, 8):
            a = ws.cell(row=r, column=1).value
            if a:
                rows[a] = [ws.cell(row=r, column=c).value for c in (2, 3)]
        # The header line carries the subtotal formula/values.
        hdr = rows["1.2 · Взносы"]
        assert hdr[0] is not None and hdr[1] is not None
        # Components still have their own rows; no extra «итого» row.
        assert "СО" in "".join(str(k) for k in rows)
        assert not any("итого" in str(k).lower() for k in rows)

    def test_columns_orientation_spreads_instances_across(self, tmp_path):
        # The board form: instances as columns, scalar fields as rows.
        m = mo.Model("m")
        m.board = mo.MultiVariable(
            display_name="Board", excel_props={"tab": True, "orient": "columns"}
        )
        m.board.p1 = mo.MultiVariable(display_name="P1")
        m.board.p1.cost = mo.Variable(50.0, display_name="Cost")
        m.board.p1.term = mo.Variable(12, display_name="Term")
        m.board.p2 = mo.MultiVariable(display_name="P2")
        m.board.p2.cost = mo.Variable(80.0, display_name="Cost")
        m.board.p2.term = mo.Variable(18, display_name="Term")
        path = tmp_path / "board.xlsx"
        m.to_excel(str(path))
        from openpyxl import load_workbook
        ws = load_workbook(str(path))["Board"]
        assert [ws.cell(row=1, column=c).value for c in (2, 3)] == ["P1", "P2"]
        assert ws.cell(row=2, column=1).value == "Cost"
        assert [ws.cell(row=2, column=c).value for c in (2, 3)] == [50.0, 80.0]
        assert [ws.cell(row=3, column=c).value for c in (2, 3)] == [12, 18]

    def test_meta_columns_article_and_unit(self, tmp_path):
        # Their book's form: № | наименование | ед. | values…
        m = mo.Model("m", default_grain="month", default_start="2026-01",
                     default_periods=2)
        m.default_excel_view = ExcelView(meta={"article": True, "unit": True})
        m.tab = mo.MultiVariable(display_name="Tab", excel_props={"tab": True})
        m.tab.fot = mo.Variable([10.0, 11.0], unit="₸",
                                display_name="ФОТ",
                                excel_props={"article": "1.1"})
        path = tmp_path / "meta.xlsx"
        m.to_excel(str(path))
        from openpyxl import load_workbook
        ws = load_workbook(str(path))["Tab"]
        row = next(r for r in range(1, 6)
                   if ws.cell(row=r, column=2).value == "ФОТ")
        assert ws.cell(row=row, column=1).value == "1.1"      # №
        assert ws.cell(row=row, column=3).value == "₸"        # ед.
        assert ws.cell(row=row, column=4).value == 10.0       # values from D
        assert ws.cell(row=1, column=4).value is not None     # header shifts
        assert ws.freeze_panes == "D2"

    def test_sanitized_sheet_name_keeps_meta(self, tmp_path):
        # 'P/L' → worksheet 'P_L'; the meta lookup must survive the
        # rename (identity key), or the book misdates every value.
        m = mo.Model("m", default_grain="month", default_start="2026-01",
                     default_periods=2)
        m.default_excel_view = ExcelView(meta={"article": True, "unit": True})
        m.tab = mo.MultiVariable(display_name="P/L", excel_props={"tab": True})
        m.tab.x = mo.Variable([10.0, 11.0], unit="₸",
                              excel_props={"article": "1.1"})
        path = tmp_path / "sanitized.xlsx"
        m.to_excel(str(path))
        from openpyxl import load_workbook
        wb = load_workbook(str(path))
        ws = wb[wb.sheetnames[0]]
        # Exactly TWO period labels, starting at the shifted column D.
        row1 = [ws.cell(row=1, column=c).value for c in range(1, 7)]
        assert row1[3] == "Jan 2026" and row1[4] == "Feb 2026"
        assert row1[1] is None and row1[5] is None
        assert ws.freeze_panes == "D2"
        r = next(r for r in range(1, 5) if ws.cell(row=r, column=2).value)
        assert ws.cell(row=r, column=1).value == "1.1"

    def test_nested_board_does_not_stretch_the_timeline(self, tmp_path):
        m = mo.Model("m", default_grain="month", default_start="2026-01",
                     default_periods=3)
        m.tab = mo.MultiVariable(display_name="Tab", excel_props={"tab": True})
        m.tab.revenue = mo.Variable([1.0, 2.0, 3.0], display_name="Rev")
        m.tab.projects = mo.MultiVariable(
            display_name="Projects", excel_props={"orient": "columns"})
        for i in range(1, 6):
            inst = mo.MultiVariable(display_name=f"P{i}")
            setattr(m.tab.projects, f"p{i}", inst)
            inst.cost = mo.Variable(float(i), display_name="Cost")
        path = tmp_path / "board_tl.xlsx"
        m.to_excel(str(path))
        from openpyxl import load_workbook
        ws = load_workbook(str(path))["Tab"]
        labels = [ws.cell(row=1, column=c).value for c in range(2, 8)]
        assert labels[:3] == ["Jan 2026", "Feb 2026", "Mar 2026"]
        assert labels[3] is None, "phantom period over instance columns"

    def test_header_row_misplacements_fall_back_to_normal_rows(self, tmp_path):
        m = mo.Model("m")
        m.tab = mo.MultiVariable(display_name="Tab", excel_props={"tab": True})
        m.tab.a = mo.Variable(1.0, display_name="A")
        # Tab-level header_row: no section header to ride — a row, not
        # a silent drop.
        m.tab.total = mo.Variable(
            m.tab.a * 2, display_name="Total",
            excel_props={"header_row": True})
        m.tab.sec = mo.MultiVariable(display_name="Sec")
        m.tab.sec.x = mo.Variable(3.0, display_name="X")
        m.tab.sec.t1 = mo.Variable(
            m.tab.sec.x * 2, excel_props={"header_row": True})
        # Second header_row sibling: only one rides the header.
        m.tab.sec.t2 = mo.Variable(
            m.tab.sec.x * 3, display_name="T2",
            excel_props={"header_row": True})
        path = tmp_path / "hdr_misplaced.xlsx"
        m.to_excel(str(path))
        from openpyxl import load_workbook
        ws = load_workbook(str(path))["Tab"]
        col_a = [ws.cell(row=r, column=1).value for r in range(1, 10)]
        assert "Total" in col_a, "tab-level header_row row dropped"
        assert "T2" in col_a, "second header_row sibling dropped"

    def test_board_keeps_direct_scalar_rows(self, tmp_path):
        m = mo.Model("m")
        m.board = mo.MultiVariable(
            display_name="Board", excel_props={"tab": True, "orient": "columns"})
        m.board.p1 = mo.MultiVariable(display_name="P1")
        m.board.p1.cost = mo.Variable(50.0, display_name="Cost")
        m.board.note = mo.Variable(99.0, display_name="Grand total")
        path = tmp_path / "board_scalar.xlsx"
        m.to_excel(str(path))
        from openpyxl import load_workbook
        ws = load_workbook(str(path))["Board"]
        cells = [c.value for row in ws.iter_rows() for c in row]
        assert "Grand total" in cells and 99.0 in cells

    def test_transposed_sheet_ignores_meta(self, tmp_path):
        m = mo.Model("m", default_grain="month", default_start="2026-01",
                     default_periods=3)
        m.default_excel_view = ExcelView(
            orient="down", meta={"article": True, "unit": True})
        m.tab = mo.MultiVariable(display_name="Tab", excel_props={"tab": True})
        m.tab.rev = mo.Variable([1.0, 2.0, 3.0], unit="₸", display_name="Rev")
        m.tab.cost = mo.Variable([4.0, 5.0, 6.0], unit="₸", display_name="Cost")
        path = tmp_path / "transposed_meta.xlsx"
        m.to_excel(str(path))
        from openpyxl import load_workbook
        ws = load_workbook(str(path))["Tab"]
        row1 = [ws.cell(row=1, column=c).value for c in (1, 2, 3)]
        assert row1 == ["Rev (₸)", "Cost (₸)", None]
        assert ws.freeze_panes in ("B1", "B2")

    def test_scalar_plus_cumsum_rolls_forward_without_compounding(self, tmp_path):
        # остаток = начало + cumsum(поток): the opening balance must
        # live ONLY in the seed cell — later cells are prev + flow.
        m = mo.Model("m", default_grain="month", default_start="2026-01",
                     default_periods=3)
        m.tab = mo.MultiVariable(display_name="Tab", excel_props={"tab": True})
        m.tab.opening = mo.Variable(1000.0, display_name="Opening")
        m.tab.flow = mo.Variable([10.0, 20.0, 30.0], display_name="Flow")
        m.tab.balance = mo.Variable(
            m.tab.opening + mo.cumsum(m.tab.flow), display_name="Balance")
        path = tmp_path / "chain.xlsx"
        m.to_excel(str(path))
        from openpyxl import load_workbook
        ws = load_workbook(str(path))["Tab"]
        row = next(r for r in range(1, 8)
                   if ws.cell(row=r, column=1).value == "Balance")
        # Periods start after the constants column (Opening is one) —
        # read the row rather than assuming column B.
        cells = [ws.cell(row=row, column=c).value for c in range(2, 8)]
        f0, f1 = [c for c in cells if c is not None][:2]
        opening_row = next(r for r in range(1, 8)
                           if ws.cell(row=r, column=1).value == "Opening")
        assert f"B{opening_row}" in f0                     # seed refs opening
        # Later periods: prev + flow ONLY — the opening ref must not
        # reappear (it lives inside prev).
        assert f"B{opening_row}" not in f1, f1

    def test_sum_over_foreign_list_inlines_the_value(self, tmp_path):
        # SUM over an address-less list can't render positionally — the
        # cell must carry the computed VALUE, never a collapsed formula.
        floating = mo.Variable([1.0, 2.0, 3.0])
        m = mo.Model("m")
        m.tab = mo.MultiVariable(display_name="Tab", excel_props={"tab": True})
        m.tab.total = mo.Variable(mo.SUM(floating), display_name="Total")
        path = tmp_path / "lossy.xlsx"
        import warnings as w
        with w.catch_warnings():
            w.simplefilter("ignore")
            m.to_excel(str(path))
        from openpyxl import load_workbook
        ws = load_workbook(str(path))["Tab"]
        row = next(r for r in range(1, 6)
                   if ws.cell(row=r, column=1).value == "Total")
        cell = ws.cell(row=row, column=2).value
        assert cell == 6.0, f"expected the value, got {cell!r}"


def test_meta_flag_is_its_own_column_caption(tmp_path):
    """The period labels deliberately start AFTER the lead metadata
    columns, which left those columns unlabelled — a column nobody
    heads is a column the reader guesses at. A STRING flag heads it;
    a bare ``True`` keeps the historical headerless column, so the
    engine stays language-neutral and the wording lives in the model.
    """
    import openpyxl
    import modeleon as mo

    def build(meta):
        m = mo.Model(
            "m",
            display_name="M",
            default_grain="month",
            default_start="2025-01",
            default_periods=2,
        )
        m.default_excel_view = mo.ExcelView(meta=meta)
        tab = mo.MultiVariable("P&L", excel_props={"tab": True})
        tab.revenue = mo.Variable([10, 20], display_name="Выручка", unit="₸")
        m.pnl = tab
        path = tmp_path / f"{'-'.join(map(str, meta.values()))}.xlsx"
        m.to_excel(str(path))
        return openpyxl.load_workbook(str(path)).active

    named = build({"article": "№", "unit": "Ед. изм."})
    assert named.cell(row=1, column=1).value == "№"
    assert named.cell(row=1, column=3).value == "Ед. изм."
    # The captions never displace the period band.
    assert named.cell(row=1, column=4).value == "Jan 2025"
    # …and the column still carries the unit itself, row by row.
    assert named.cell(row=2, column=3).value == "₸"

    bare = build({"article": True, "unit": True})
    assert bare.cell(row=1, column=1).value is None
    assert bare.cell(row=1, column=3).value is None
    assert bare.cell(row=1, column=4).value == "Jan 2025"
    assert bare.cell(row=2, column=3).value == "₸"


class TestNestedSectionTitleIndent:
    def test_the_title_indents_like_its_rows_do(self, tmp_path):
        # Dense nesting shows hierarchy by indent alone (gap_rows=0) —
        # and only variable labels were indented, so «Договоры» inside
        # «Проект №1» sat flush left and read as its parent's sibling.
        m = mo.Model("m", display_name="M",
                     default_excel_view=ExcelView(
                         nesting={"gap_rows": 0, "indent": 2}))
        m.проекты = mo.MultiVariable("Проекты", excel_props={"tab": True})
        m.проекты.p1 = mo.MultiVariable("Проект №1")
        m.проекты.p1.стоимость = mo.Variable(100, display_name="Стоимость")
        m.проекты.p1.договоры = mo.MultiVariable("Договоры")
        m.проекты.p1.договоры.д1 = mo.Variable(50, display_name="Договор А")
        path = tmp_path / "n.xlsx"
        m.to_excel(str(path))
        col_a = [
            load_workbook(str(path))["Проекты"].cell(r, 1).value
            for r in range(1, 5)
        ]
        assert col_a[0] == "Проект №1"          # level 0 — flush
        assert col_a[2] == "  Договоры"          # level 1 — one indent
        assert col_a[3] == "    Договор А"       # its row, one deeper

    def test_no_indent_declared_no_padding_invented(self, tmp_path):
        m = mo.Model("m", display_name="M")
        m.проекты = mo.MultiVariable("Проекты", excel_props={"tab": True})
        m.проекты.p1 = mo.MultiVariable("Проект №1")
        m.проекты.p1.договоры = mo.MultiVariable("Договоры")
        m.проекты.p1.договоры.д1 = mo.Variable(50, display_name="Договор А")
        path = tmp_path / "n.xlsx"
        m.to_excel(str(path))
        ws = load_workbook(str(path))["Проекты"]
        titles = [ws.cell(r, 1).value for r in range(1, 6)]
        assert "Договоры" in titles              # exact, unpadded


class TestCtorKwargViewIsASlot:
    def test_a_windowed_tracked_model_takes_the_view_kwarg(self, tmp_path):
        # The ctor's registerable loop used to adopt the view as
        # CONTENT — its config leaves (totals=['quarter','year']) hit
        # the extent law under the model's window («'totals' has 2
        # value(s) in a 36-period window», live on a focused
        # Бизнес-план) and the view rendered as a spurious tab.
        m = mo.Model(
            "бп", display_name="БП",
            tracks=mo.Tracks(
                "план", "факт",
                blend=mo.blend(given="факт", follow="план",
                               until="2026-03"),
            ),
            default_grain="month", default_start="2026-01",
            default_periods=36,
            default_excel_view=ExcelView(
                tracks="blend",
                timeline={"totals": ["quarter", "year"]},
            ),
        )
        assert "default_excel_view" not in m._components
        assert resolve_excel_view(m).time_totals == ("quarter", "year")
        m.поток = mo.Variable([1.0] * 36, display_name="Поток",
                              regrain=mo.up("sum"))
        m.to_excel(tmp_path / "v.xlsx")
        assert "default_excel_view" not in load_workbook(
            tmp_path / "v.xlsx"
        ).sheetnames
