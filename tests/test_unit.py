# SPDX-License-Identifier: Apache-2.0
"""Unit propagation through arithmetic and Excel emission."""

import pytest
from openpyxl import load_workbook

import modeleon as mo


class TestBasicConstruction:
    def test_unit_stored_on_variable(self):
        v = mo.Variable(10, unit="$")
        assert str(v._unit) == "$"

    def test_compound_unit_parsed(self):
        v = mo.Variable(10, unit="$/hr")
        assert str(v._unit) == "$/hr"

    def test_no_unit_is_none(self):
        v = mo.Variable(10)
        assert v._unit is None


class TestMultiplication:
    def test_same_unit_composes(self):
        a = mo.Variable(5, unit="hr")
        b = mo.Variable(2, unit="hr")
        assert str((a * b)._unit) == "hr^2"

    def test_different_units_compose(self):
        price = mo.Variable(50, unit="$")
        hours = mo.Variable(10, unit="hr")
        assert str((price * hours)._unit) == "$ * hr"

    def test_scalar_preserves_unit(self):
        h = mo.Variable(10, unit="hr")
        assert str((h * 2)._unit) == "hr"
        assert str((2 * h)._unit) == "hr"


class TestDivision:
    def test_division_simplifies(self):
        revenue = mo.Variable(1000, unit="$")
        hours = mo.Variable(10, unit="hr")
        assert str((revenue / hours)._unit) == "$/hr"

    def test_cancellation_yields_dimensionless(self):
        a = mo.Variable(10, unit="$")
        b = mo.Variable(2, unit="$")
        assert (a / b)._unit is None

    def test_compound_division_cancels(self):
        cost = mo.Variable(100, unit="$")
        rate = mo.Variable(5, unit="$/hr")
        # $ / ($/hr) = hr
        assert str((cost / rate)._unit) == "hr"


class TestUnitReuse:
    """When ``*`` / ``/`` produce a unit one operand already carries, that
    operand's own Unit is the result (as ``+`` / ``-`` do) — equal to what
    the algebra would build, without building it."""

    def test_times_unitless_keeps_the_same_unit(self):
        a = mo.Variable(5, unit="$")
        r = a * mo.Variable(2)
        assert r._unit is a._unit and str(r._unit) == "$"
        assert (mo.Variable(2) * a)._unit is a._unit

    def test_times_dimensionless_keeps_the_same_unit(self):
        a = mo.Variable(5, unit="$")
        pct = mo.Variable(0.2, unit="%")
        assert (a * pct)._unit is a._unit
        assert (pct * a)._unit is a._unit

    def test_divided_by_unitless_keeps_the_same_unit(self):
        a = mo.Variable(12, unit="$/hr")
        assert (a / mo.Variable(4))._unit is a._unit
        assert (a / mo.Variable(4, unit="%"))._unit is a._unit

    def test_unitless_over_a_unit_still_inverts_it(self):
        r = mo.Variable(1) / mo.Variable(4, unit="hr")
        assert str(r._unit) == "1/hr"

    def test_dimensionless_results_stay_none(self):
        pct = mo.Variable(0.2, unit="%")
        assert (pct * pct)._unit is None
        assert (pct * mo.Variable(3))._unit is None
        assert (pct / mo.Variable(3))._unit is None

    def test_no_new_unit_is_built_for_a_reused_result(self, monkeypatch):
        from modeleon.core.unit import Unit

        a = mo.Variable(5, unit="$")
        built = []
        original = Unit.__init__

        def counting(self, *args, **kwargs):
            built.append(1)
            original(self, *args, **kwargs)

        monkeypatch.setattr(Unit, "__init__", counting)
        _ = a * mo.Variable(2) * mo.Variable(0.2, unit="%") / mo.Variable(12)
        # The only Unit built here is the one the "%" literal parses.
        assert len(built) == 1


class TestAddition:
    def test_same_unit_keeps_unit(self):
        a = mo.Variable(10, unit="$")
        b = mo.Variable(20, unit="$")
        assert str((a + b)._unit) == "$"

    def test_mixed_units_raise(self):
        a = mo.Variable(10, unit="$")
        b = mo.Variable(20, unit="hr")
        with pytest.raises(ValueError, match="incompatible units"):
            _ = a + b

    def test_subtract_mixed_raise(self):
        a = mo.Variable(10, unit="$")
        b = mo.Variable(20, unit="hr")
        with pytest.raises(ValueError, match="incompatible units"):
            _ = a - b

    def test_scalar_plus_unitful_keeps_unit(self):
        a = mo.Variable(10, unit="$")
        # 5 is dimensionless/scalar — treat as compatible, keep unit
        assert str((a + 5)._unit) == "$"

    def test_a_rate_per_year_beside_an_amount_points_to_the_year_share(self):
        # 'USD/year' times days / 365 stays per year: the refusal says how
        # to make it an amount, and that way adds.
        m = mo.Model("m", default_grain="month", default_start="2026-01", default_periods=3)
        with m:
            m.rate = mo.Variable(1200.0, unit="USD/year")
            m.amount = mo.Variable([100.0] * 3, unit="USD")
            with pytest.raises(ValueError, match="per 'year'") as err:
                _ = m.amount + m.rate * mo.time.days / 365
            assert "year_share = (mo.time.months / 12).set_unit('year')" in str(err.value)
            year_share = (mo.time.months / 12).set_unit("year")
            total = m.amount + m.rate * year_share
        assert str(total._unit) == "USD"
        assert total.value == pytest.approx([200.0] * 3)

    def test_a_rate_per_another_word_names_it_without_a_year(self):
        a, b = mo.Variable(10.0, unit="USD/seat"), mo.Variable(1.0, unit="USD")
        with pytest.raises(ValueError, match="per 'seat'") as err:
            _ = b + a
        assert "year_share" not in str(err.value)


class TestPower:
    def test_integer_exponent_raises_unit(self):
        t = mo.Variable(3, unit="hr")
        assert str((t ** 2)._unit) == "hr^2"
        assert str((t ** 3)._unit) == "hr^3"

    def test_zero_exponent_dimensionless(self):
        t = mo.Variable(3, unit="hr")
        # hr^0 = dimensionless
        assert (t ** 0)._unit is None

    def test_float_exponent_drops_unit(self):
        t = mo.Variable(3, unit="hr")
        # Fractional/non-int exponents have no clean unit rendering
        assert (t ** 0.5)._unit is None


class TestNegation:
    def test_unary_minus_preserves_unit(self):
        a = mo.Variable(10, unit="$")
        assert str((-a)._unit) == "$"


class TestExcelEmission:
    def test_unit_appears_in_label(self, tmp_path):
        m = mo.MultiVariable("UT")
        m.u = mo.MultiVariable("U", excel_props={'tab': True})
        m.u.revenue = mo.Variable(1000, unit="$")
        m.to_excel(tmp_path / "u.xlsx")
        path = tmp_path / "u.xlsx"
        ws = load_workbook(path)["U"]
        labels = [
            cell.value for row in ws.iter_rows() for cell in row
            if isinstance(cell.value, str)
        ]
        assert any("($)" in label for label in labels)

    def test_no_unit_no_suffix(self, tmp_path):
        m = mo.MultiVariable("UT")
        m.u = mo.MultiVariable("U", excel_props={'tab': True})
        m.u.plain = mo.Variable(1000)
        m.to_excel(tmp_path / "u.xlsx")
        path = tmp_path / "u.xlsx"
        ws = load_workbook(path)["U"]
        labels = [
            cell.value for row in ws.iter_rows() for cell in row
            if isinstance(cell.value, str)
        ]
        assert not any("(" in label for label in labels)


# ══════════════════════════════════════════════════════════════════════
# A word that labels a pure number
# ══════════════════════════════════════════════════════════════════════


FACTOR = mo.Unit.dimensionless("factor")
SHARE = mo.Unit.dimensionless("share")


class TestLabelledPureNumber:
    def test_a_label_prints_and_is_inert(self):
        cash = mo.Variable(100.0, unit="USD m")
        df = mo.Variable(0.9, unit=FACTOR)
        assert str(df.unit) == "factor" and FACTOR.is_dimensionless
        pv = cash * df
        assert pv.unit is cash.unit                     # the label adds nothing
        assert str((cash + pv).unit) == "USD m"         # and so the two add

    def test_a_string_is_always_a_dimension(self):
        cash = mo.Variable(100.0, unit="USD m")
        df = mo.Variable(0.9, unit="factor")
        assert mo.Unit("factor")._unit_components == {"factor": 1}
        with pytest.raises(ValueError, match="incompatible units"):
            cash + cash * df

    def test_pure_numbers_add_the_left_label_prints_a_product_drops_it(self):
        df = mo.Variable(0.9, unit=FACTOR)
        sh = mo.Variable(0.1, unit=SHARE)
        assert str((df + sh).unit) == "factor"
        assert str((sh + df).unit) == "share"
        assert (df * df).unit is None and (df * 2).unit is None
        pct, pp = mo.Variable(0.05, unit="%"), mo.Variable(0.01, unit="pp")
        assert str((pct + pp).unit) == "%"              # the conventional words, unchanged

    def test_the_conventional_words_are_the_same_kind_of_unit(self):
        assert mo.Unit("%") == mo.Unit.dimensionless("%") == mo.Unit.dimensionless()
        assert str(mo.Unit("%")) == "%"

    def test_repr_rebuilds_the_unit(self):
        for u in (FACTOR, mo.Unit("%"), mo.Unit("USD m"), mo.Unit("$/hr"),
                  mo.Unit.dimensionless()):
            back = eval(repr(u), {"Unit": mo.Unit})
            assert back == u and str(back) == str(u), repr(u)
        assert repr(FACTOR) == "Unit.dimensionless('factor')"

    @pytest.mark.parametrize("label", ["", "   ", "a/b", "a * b", "m^2"])
    def test_a_label_is_one_word(self, label):
        with pytest.raises(ValueError):
            mo.Unit.dimensionless(label)

    def test_what_is_not_a_unit_is_refused(self):
        with pytest.raises(TypeError):
            mo.Unit.dimensionless(3)
        with pytest.raises(TypeError):
            mo.Unit("factor", dimensionless=True)       # no stray keywords
        with pytest.raises(TypeError, match="a unit is a string"):
            mo.Variable(1.0, unit=5)
        with pytest.raises(TypeError, match="a unit is a string"):
            mo.Variable(1.0).set_unit(mo.Variable(2.0))

    def test_the_add_error_names_the_word_that_leaked_through_a_product(self):
        cash = mo.Variable(100.0, unit="USD m")
        with pytest.raises(ValueError) as err:
            cash + cash * mo.Variable(0.9, unit="factor")
        msg = str(err.value)
        assert "'factor' is a dimension here" in msg
        assert "mo.Unit.dimensionless('factor')" in msg

    @pytest.mark.parametrize("left, right", [
        ("EUR", "USD"), ("EUR/kWh", "USD/kWh"), ("USD m/yr", "EUR m/yr"),
        ("MWh/yr", "GWh/yr"), ("USD*hr", "USD*day"), ("USD/hr", "USD*hr"),
        ("USD*hr", "USD/hr"), ("USD", "USD*hr*hr"), ("kWh", "kWh*h*h"),
    ])
    def test_the_add_error_stays_plain_where_no_word_leaked(self, left, right):
        """Neither side is the other times words: a real mismatch, no hint."""
        with pytest.raises(ValueError) as err:
            mo.Variable(1.0, unit=left) + mo.Variable(1.0, unit=right)
        assert "dimensionless" not in str(err.value), (left, right)

    @pytest.mark.parametrize("left, right", [
        ("USD m", "USD m * factor"), ("USD m * factor", "USD m"),
        ("EUR/kWh", "EUR * flag/kWh"),
    ])
    def test_the_hint_fires_in_either_order(self, left, right):
        with pytest.raises(ValueError) as err:
            mo.Variable(1.0, unit=left) + mo.Variable(1.0, unit=right)
        assert "mo.Unit.dimensionless(" in str(err.value), (left, right)

    def test_the_add_error_tells_a_dimension_from_a_label_of_the_same_word(self):
        with pytest.raises(ValueError) as err:
            mo.Variable(1.0, unit="factor") + mo.Variable(1.0, unit=FACTOR)
        assert "factor (a dimension) and factor (a labelled pure number)" in str(err.value)

    def test_a_unit_is_a_value_never_a_row(self, tmp_path):
        m = mo.Model("m", default_grain="year", default_start="2026", default_periods=2)
        with m:
            m.F = FACTOR
            m.s = mo.MultiVariable("S")
            with m.s as s:
                s.df = mo.Variable([0.9, 0.8], unit=m.F, display_name="Discount factor")
        assert "F" not in m._component_names and m.F is FACTOR
        m.to_excel(tmp_path / "u.xlsx")
        assert load_workbook(tmp_path / "u.xlsx").sheetnames == ["S"]

    def test_the_label_survives_copy_and_pickle(self):
        import pickle
        df = mo.Variable(0.9, unit=FACTOR)
        assert str(df.copy().unit) == "factor"
        assert str(pickle.loads(pickle.dumps(df)).unit) == "factor"
        assert str(pickle.loads(pickle.dumps(FACTOR))) == "factor"

    def test_the_book_prints_the_label(self, tmp_path):
        m = mo.MultiVariable("UT")
        m.u = mo.MultiVariable("U", excel_props={'tab': True})
        m.u.df = mo.Variable(0.9, unit=FACTOR, display_name="Discount factor")
        m.to_excel(tmp_path / "u.xlsx")
        ws = load_workbook(tmp_path / "u.xlsx")["U"]
        labels = [c.value for row in ws.iter_rows() for c in row if isinstance(c.value, str)]
        assert "Discount factor (factor)" in labels

    def test_code_writes_the_constructor_for_a_label(self):
        df = mo.Variable(0.9, unit=FACTOR)
        assert "unit=mo.Unit.dimensionless('factor')" in df.code
        assert "unit='USD m'" in mo.Variable(1.0, unit="USD m").code
        assert "unit='%'" in mo.Variable(0.1, unit="%").code

    def test_an_aggregate_over_a_dimension_and_its_namesake_names_neither(self):
        dim = mo.Variable(1.0, unit="factor")
        lab = mo.Variable(1.0, unit=FACTOR)
        assert mo.SUM(dim, lab).unit is None
        assert mo.SUM(lab, mo.Variable(1.0, unit=SHARE)).unit is None
        assert str(mo.SUM(lab, lab).unit) == "factor"


class TestDeclaredUnitAcrossViews:
    def test_a_declared_unit_survives_a_grain_change(self):
        m = mo.Model("m", default_grain="month", default_start="2026-01", default_periods=6)
        with m:
            m.s = mo.MultiVariable("S")
            with m.s as s:
                s.a = mo.Variable([0.1] * 6, unit="%", regrain=mo.up("mean"))
                s.b = mo.Variable([0.2] * 6, unit="%", regrain=mo.up("mean"))
                s.f = (s.a + s.b).set_unit(FACTOR)
                s.c1 = mo.Variable([1.0] * 6, unit="USD", regrain=mo.up("sum"))
                s.c2 = (s.c1 + s.c1).set_unit("USD m")
        q = m.at("quarter")
        assert str(q.s.f.unit) == "factor"
        assert str(q.s.c2.unit) == "USD m"

    def test_a_container_in_the_units_place_carries_no_unit(self):
        stand_in = mo.MultiVariable("Elsewhere")
        assert mo.Variable(1.0, unit=stand_in).unit is None

    def test_two_dimensions_written_alike_are_not_called_dimension_and_label(self):
        a = mo.Variable([1.0, 2.0], unit="hr")
        with pytest.raises(ValueError) as err:
            (a * a) + mo.Variable([1.0, 1.0], unit="hr^2")
        assert "labelled pure number" not in str(err.value)
