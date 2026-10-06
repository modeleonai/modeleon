# SPDX-License-Identifier: Apache-2.0
"""``ExcelView(orient='records')`` — records laid out as a table.

The default stacked layout gives every child MV a header row and a
block of its own. That is right for sections of a model and wrong for a
LIST: a list of records with five fields each becomes one header per
record and five rows under each, repeating the same five labels, and no
one writes that by hand.

Records orientation says the shape out loud, and the fields split by
whether they carry a period axis:

* scalar fields → a table, fields across, records down;
* per-period fields → one block each, titled once, records under it,
  values landing in the sheet's own period columns.

Both halves keep real addresses, so formulas stay live and any renderer
that reads the layout draws the same picture as the .xlsx.

The form is DECLARED, never assumed: a container prints as stacked
sections unless its own view says otherwise. Nothing cascades it — a
model that declared it would turn every section into a record table,
and a section whose children are Variables has no records to lay out,
so its rows would silently vanish. A shape this violent is stated
where it applies.
"""

import openpyxl
import pytest

import modeleon as mo


def _grid(model, sheet, cols=6):
    import os
    import tempfile

    path = os.path.join(tempfile.mkdtemp(), "t.xlsx")
    model.to_excel(path)
    ws = openpyxl.load_workbook(path)[sheet]
    return [
        [ws.cell(r, c).value for c in range(1, cols + 1)]
        for r in range(1, ws.max_row + 1)
    ]


class Employee(mo.MultiVariableClass):
    def compute(self, salary):
        self.salary = mo.Variable(salary, display_name="Salary")
        self.bonus = mo.Variable(salary * 0.1, display_name="Bonus")


def _staff():
    m = mo.Model("m", display_name="M")
    with m:
        m.staff = mo.MultiVariable("Staff", excel_props={"tab": True})
        m.staff.default_excel_view = mo.ExcelView(orient="records")
        with m.staff as c:
            c.s1 = Employee(salary=100, display_name="Employee 1")
            c.s2 = Employee(salary=200, display_name="Employee 2")
    return m


class TestScalarRecords:
    def test_fields_across_records_down(self):
        g = _grid(_staff(), "Staff")
        assert g[0][:3] == [None, "Salary", "Bonus"]
        assert g[1][:3] == ["Employee 1", 100, 10]
        assert g[2][:3] == ["Employee 2", 200, 20]
        assert len(g) == 3  # no per-record sections

    def test_a_field_only_some_records_have_keeps_its_column(self):
        # First-appearance order, and a record that lacks a field
        # leaves the cell empty rather than shifting the row — the
        # failure that turns a table into misaligned nonsense.
        class Custom(mo.MultiVariableClass):
            def compute(self, salary, bonus=None):
                self.salary = mo.Variable(salary, display_name="Salary")
                if bonus is not None:
                    self.bonus = mo.Variable(bonus, display_name="Bonus")

        m = mo.Model("m", display_name="M")
        with m:
            m.staff = mo.MultiVariable(
                "Staff", excel_props={"tab": True, "orient": "records"}
            )
            with m.staff as c:
                c.a = Custom(salary=1, display_name="A")
                c.b = Custom(salary=2, bonus=9, display_name="B")
                c.c = Custom(salary=3, display_name="C")
        g = _grid(m, "Staff")
        assert g[0][:3] == [None, "Salary", "Bonus"]
        assert g[1][:3] == ["A", 1, None]
        assert g[2][:3] == ["B", 2, 9]
        assert g[3][:3] == ["C", 3, None]

    def test_the_first_record_names_the_column(self):
        # Records of one class usually agree on a field's display name.
        # When they don't, the header must be the first record's — a
        # stable rule — not "whoever the walk wrote last".
        class Custom(mo.MultiVariableClass):
            def compute(self, salary, label):
                self.salary = mo.Variable(salary, display_name=label)

        m = mo.Model("m", display_name="M")
        with m:
            m.staff = mo.MultiVariable(
                "Staff", excel_props={"tab": True, "orient": "records"}
            )
            with m.staff as c:
                c.a = Custom(salary=1, label="Salary", display_name="A")
                c.b = Custom(salary=2, label="Pay", display_name="B")
        g = _grid(m, "Staff")
        assert g[0][1] == "Salary"
        assert [row[1] for row in g[1:3]] == [1, 2]

    def test_no_period_header_over_a_scalar_table(self):
        # The sheet's columns are FIELDS here. A period header would
        # name column B "Jan 2026" when B is "Salary" — the same lie
        # ``orient='columns'`` refuses, and it must not cost a
        # reserved blank row either.
        m = _staff()
        m.default_start, m.default_grain, m.default_periods = (
            "2026-01",
            "month",
            3,
        )
        g = _grid(m, "Staff")
        assert g[0][:2] == [None, "Salary"]


class TestPerPeriodRecords:
    """A series cannot be a column of a record table: that table's
    columns are fields and a series' columns are periods."""

    def _model(self):
        m = mo.Model("m", display_name="M")
        with m:
            m.rate = mo.Variable(0.5, display_name="Rate")

            class Project(mo.MultiVariableClass):
                def compute(self, price):
                    self.price = mo.Variable(price, display_name="Price")
                    self.profile = mo.Variable(
                        [0.2, 0.3, 0.5], display_name="Profile"
                    )
                    self.revenue = self.profile * self.price
                    self.share = self.revenue * m.rate

            m.projects = mo.MultiVariable(
                "Projects", excel_props={"tab": True, "orient": "records"}
            )
            with m.projects as c:
                c.p1 = Project(price=100, display_name="Project 1")
                c.p2 = Project(price=200, display_name="Project 2")
        return m

    def test_one_block_per_field_records_under_it(self):
        g = _grid(self._model(), "Projects")
        labels = [row[0] for row in g]
        # Scalar table first, then a block per series field.
        assert labels[:3] == [None, "Project 1", "Project 2"]
        assert [x for x in labels if x in ("Profile", "Revenue", "Share")] == [
            "Profile",
            "Revenue",
            "Share",
        ]
        i = labels.index("Profile")
        assert labels[i + 1 : i + 3] == ["Project 1", "Project 2"]
        # Period values start PAST the fields table (one scalar field
        # here → column C onward), so the month masthead never paints
        # over "Price".
        assert g[i + 1][2:5] == pytest.approx([0.2, 0.3, 0.5])
        assert g[i + 2][2:5] == pytest.approx([0.2, 0.3, 0.5])
        # And the computed block reads ACROSS the blocks: profile row ×
        # the record's own price cell in the table above.
        j = labels.index("Revenue")
        assert g[j + 1][2] == f"=C{i + 2} * B2"

    def test_formulas_stay_live_across_the_blocks(self):
        # The whole reason a list of records cannot be flattened to a
        # value table: "share" references "revenue" in another block and the
        # model's rate on another sheet, and both must survive as cell
        # references, not as numbers.
        g = _grid(self._model(), "Projects")
        labels = [row[0] for row in g]
        i = labels.index("Share")
        first = g[i + 1][2]
        assert isinstance(first, str) and first.startswith("=")
        assert "M!" in first  # the model's rate, by reference

    def test_months_never_paint_over_the_fields_table(self):
        """The defect this geometry exists to prevent: a field header
        ("Price") under "Feb 2026". The month masthead is read back
        from the period blocks' value addresses; those must start past
        the widest scalar column, leaving row 1 over the table empty."""
        m = self._model()
        m.default_start, m.default_grain, m.default_periods = (
            "2026-01", "month", 3,
        )
        import pathlib as _pl
        import tempfile
        with tempfile.TemporaryDirectory() as d:
            path = _pl.Path(d) / "m.xlsx"
            m.to_excel(str(path))
            ws = openpyxl.load_workbook(str(path))["Projects"]
            # Column B is "Price" — the field header owns it, and no
            # month label may stand above it.
            assert ws["B1"].value is None
            # The first month lands past the table, over the blocks.
            row1 = [ws.cell(1, c).value for c in range(1, 8)]
            months = [x for x in row1 if isinstance(x, str) and "2026" in x]
            assert months, row1
            first_month_col = row1.index(months[0]) + 1
            assert first_month_col >= 3

    def test_a_field_named_once_per_block_not_once_per_record(self):
        g = _grid(self._model(), "Projects")
        titles = [row[0] for row in g]
        assert titles.count("Revenue") == 1

    def test_records_name_themselves_in_every_block(self):
        # Rows without an identity are the defect this shape exists to
        # fix: a column of numbers with nothing saying whose.
        g = _grid(self._model(), "Projects")
        assert [row[0] for row in g].count("Project 1") == 4  # table + 3


class TestEdges:
    def test_a_container_with_no_records_lays_out_nothing(self):
        m = mo.Model("m", display_name="M")
        with m:
            m.empty = mo.MultiVariable(
                "Empty", excel_props={"tab": True, "orient": "records"}
            )
            m.x = mo.Variable(1, display_name="X")
        # No crash, and the model's own rows are untouched.
        assert _grid(m, "M")[0][:2] == ["X", 1]

    def test_records_with_no_variable_fields_lay_out_nothing(self):
        # A record built purely from constructor kwargs has no cells to
        # place — the layout must not invent rows for it.
        class Blank(mo.MultiVariableClass):
            def compute(self, note):
                self._note = note

        m = mo.Model("m", display_name="M")
        with m:
            m.x = mo.Variable(1, display_name="X")
            m.roster = mo.MultiVariable(
                "Roster", excel_props={"tab": True, "orient": "records"}
            )
            with m.roster as c:
                c.a = Blank(note=1, display_name="A")
        assert _grid(m, "Roster") in ([], [[None] * 6])


class TestTheFormIsDeclared:
    def test_a_container_without_a_view_stays_stacked(self):
        # The default every container of records prints as: a header
        # per record, its fields beneath. Right for records that carry per-period
        # lines, where the table's own repetition is worse than what it
        # replaces.
        m = mo.Model("m", display_name="M")
        with m:
            m.staff = mo.MultiVariable("Staff", excel_props={"tab": True})
            with m.staff as c:
                c.s1 = Employee(salary=100, display_name="Employee 1")
        g = _grid(m, "Staff")
        assert [row[0] for row in g][:3] == ["Employee 1", "Salary", "Bonus"]

    def test_the_excel_props_spelling_still_wins(self):
        # ``excel_props={'orient': …}`` declares the same form as the
        # view, and a model that uses it must lay out the same way.
        m = mo.Model("m", display_name="M")
        with m:
            m.staff = mo.MultiVariable(
                "Staff", excel_props={"tab": True, "orient": "records"}
            )
            with m.staff as c:
                c.s1 = Employee(salary=100, display_name="Employee 1")
        assert _grid(m, "Staff")[0][:2] == [None, "Salary"]

    def test_an_ancestor_cannot_impose_the_shape(self):
        # ``resolve_excel_view`` cascades ink down the tree. Geometry
        # must not ride it: a model-level ``orient='records'`` would
        # reach a section whose children are Variables — no records to
        # lay out — and its rows would disappear from the sheet.
        m = mo.Model("m", display_name="M")
        m.default_excel_view = mo.ExcelView(orient="records")
        with m:
            m.pnl = mo.MultiVariable("PnL", excel_props={"tab": True})
            with m.pnl as p:
                p.revenue = mo.Variable(100, display_name="Revenue")
                p.cogs = mo.Variable(60, display_name="COGS")
        g = _grid(m, "PnL")
        assert [row[0] for row in g] == ["Revenue", "COGS"]
