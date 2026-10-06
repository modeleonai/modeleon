# SPDX-License-Identifier: Apache-2.0
"""The time of a period — ``mo.time`` — and the settings that place it.

Every line of a windowed model lives in periods; ``mo.time.start`` /
``end`` / ``days`` / ``index`` are those periods, readable by any formula
and written by nobody. A window may start at a date Variable (the day a
deal is anchored to), and ``default_header`` names the rows a workbook
shows above its sheets — both are settings of a container, never rows.
"""

import pickle
from datetime import date

import pytest

import modeleon as mo
from modeleon.core.expr import TimeRef
from modeleon.core.time import resolve_default_window


def model(n=3, start='2026-01', grain='month'):
    return mo.Model('m', default_grain=grain, default_start=start, default_periods=n)


# ══════════════════════════════════════════════════════════════════════
# The values
# ══════════════════════════════════════════════════════════════════════


def test_month_fields():
    m = model(3)
    with m:
        m.s, m.e, m.d, m.i = mo.time.start, mo.time.end, mo.time.days, mo.time.index
    assert m.s.value == [date(2026, 1, 1), date(2026, 2, 1), date(2026, 3, 1)]
    assert m.e.value == [date(2026, 1, 31), date(2026, 2, 28), date(2026, 3, 31)]
    assert m.d.value == [31, 28, 31]
    assert m.i.value == [1, 2, 3]


@pytest.mark.parametrize('grain, start, days', [
    ('quarter', '2026-Q1', [90, 91, 92, 92]),
    ('year', '2026', [365, 365, 366]),
    ('day', '2028-02-28', [1, 1, 1]),
])
def test_other_grains(grain, start, days):
    m = model(len(days), start, grain)
    with m:
        m.d = mo.time.days
    assert m.d.value == days


@pytest.mark.parametrize('grain, start, months', [
    ('month', '2026-01', [1, 1, 1]),
    ('quarter', '2026-Q1', [3, 3, 3]),
    ('year', '2026', [12, 12, 12]),
])
def test_months_in_a_period(grain, start, months):
    m = model(3, start, grain)
    with m:
        m.n = mo.time.months
    assert m.n.value == months


def test_an_amount_a_month_times_months_is_the_periods_amount():
    m = model(2, '2026-Q1', 'quarter')
    with m:
        m.rent = mo.Variable(100.0) * mo.time.months
    assert m.rent.value == [300.0, 300.0]


def test_months_regrain_to_the_coarser_period():
    m = model(6, '2026-01')
    with m:
        m.n = mo.time.months
        m.rent = mo.Variable(100.0) * mo.time.months
    q = m.at('quarter')
    assert q.n.value == [3, 3]
    assert q.rent.value == [300.0, 300.0]


def test_a_partial_first_quarter_holds_the_months_the_window_covers():
    m = model(5, '2026-02')      # Feb..Jun: a two-month quarter, then a whole one
    with m:
        m.n = mo.time.months
    assert m.at('quarter').n.value == [2, 3]


def test_months_on_a_day_window_is_refused():
    m = model(2, '2026-03-01', 'day')
    with m:
        with pytest.raises(ValueError, match="A day window has none"):
            m.n = mo.time.months


def test_a_leap_february():
    m = model(2, '2028-02')
    with m:
        m.d = mo.time.days
    assert m.d.value == [29, 31]


def test_a_formula_carries_time_not_numbers():
    m = model(3)
    with m:
        m.rate = mo.Variable(0.12)
        m.accrual = m.rate * mo.time.days / 365
    assert 'time.days' in m.accrual.formula
    assert m.accrual.value == pytest.approx([0.12 * d / 365 for d in (31, 28, 31)])


def test_time_is_not_a_row_of_anything():
    m = model(3)
    with m:
        m.x = mo.Variable(2.0) * mo.time.days
    assert m._component_names == ['x']
    assert not any('time' in dep for dep in m.x.dependencies)


def test_a_nested_window_reads_its_own_periods():
    m = model(12)
    with m:
        m.q = mo.MultiVariable('Q', default_grain='quarter', default_start='2026-Q1',
                               default_periods=2)
        with m.q as q:
            q.d = mo.time.days
    assert m.q.d.value == [90, 91]


def test_a_loop_reads_the_days_of_each_period():
    m = model(3)
    with m:
        m.debt = mo.MultiVariable('Debt', loop=True)
        with m.debt as d:
            d.interest = mo.lag(d.balance, fill=1000.0) * 0.12 * mo.time.days / 365
            d.balance = mo.lag(d.balance, fill=1000.0) + d.interest
    balance, control = 1000.0, []
    for days in (31, 28, 31):
        interest = balance * 0.12 * days / 365
        balance += interest
        control.append(round(interest, 9))
    assert [round(v, 9) for v in m.debt.interest.value] == control


def test_an_earlier_row_reads_time_on_its_own_dates():
    m = model(3)
    with m:
        m.early = mo.Variable([1.0] * 5, start='2025-11', grain='month')
        m.x = m.early * mo.time.days
    assert m.x.value == [30, 31, 31, 28, 31]


def test_reading_time_outside_a_model_says_where_to_read_it():
    import contextvars
    # A fresh context: no open block, no model built yet.
    with pytest.raises(ValueError, match="no model window here"):
        contextvars.Context().run(lambda: mo.time.days)


def test_time_read_for_another_model_is_refused_where_it_lands():
    a = model(3)
    b = model(2, '2027-01')      # built last: an unscoped read borrows it
    days = mo.time.days
    with pytest.raises(ValueError, match="reads mo.time.days for 2 month"):
        a.x = days * 2
    assert b.default_periods == 2


def test_reading_time_before_the_start_names_the_missing_setting():
    m = mo.Model('m', default_grain='month', default_periods=3)
    with m:
        with pytest.raises(ValueError, match="default_start"):
            mo.time.days


def test_time_ref_is_a_value():
    assert TimeRef('days') == TimeRef('days')
    assert pickle.loads(pickle.dumps(TimeRef('end'))) == TimeRef('end')
    assert list(TimeRef('days').iter_refs()) == []


def test_the_json_renderer_names_time():
    from modeleon.compile.json.renderer import JsonRenderer
    from modeleon.compile.renderer import RenderCtx
    from modeleon.compile.walker import Walker
    m = model(3)
    with m:
        m.x = mo.Variable(2.0) * mo.time.days
    renderer = JsonRenderer()
    walker = Walker(renderer)
    rendered = walker.render(m.x._expr, RenderCtx())
    assert {'kind': 'time', 'field': 'days'} in rendered.values()


# ══════════════════════════════════════════════════════════════════════
# A window that starts at a date Variable
# ══════════════════════════════════════════════════════════════════════


def pf(fc=date(2026, 1, 1), n=4):
    m = mo.Model('pf', default_grain='month', default_periods=n)
    with m:
        m.inputs = mo.MultiVariable('Inputs')
        with m.inputs as i:
            i.financial_close = mo.Variable(fc)
            i.capex = mo.Variable([1.0] * n)
            i.grant = mo.Variable([5.0], start='2026-02', grain='month')
        m.default_start = m.inputs.financial_close
    return m


def test_the_start_is_a_setting_not_a_row():
    m = pf()
    assert m._component_names == ['inputs']
    assert m.default_start == '2026-01'
    window = resolve_default_window(m.inputs)
    assert window.start_source is m.inputs.financial_close


def test_inheriting_rows_move_with_the_start_own_rows_stay():
    early, late = pf(date(2026, 1, 1)), pf(date(2026, 3, 1))
    assert early.inputs.capex.time.start == '2026-01'
    assert late.inputs.capex.time.start == '2026-03'
    assert early.inputs.grant.time.start == late.inputs.grant.time.start == '2026-02'


def test_a_start_inside_a_period_is_refused():
    with pytest.raises(ValueError, match="1st of a month"):
        pf(date(2026, 1, 15))


def test_a_string_start_forgets_the_source():
    m = pf()
    m.default_start = '2026-05'
    assert resolve_default_window(m).start_source is None
    assert m.default_start == '2026-05'


def test_the_constructor_takes_a_date_start():
    m = mo.Model('m', default_grain='quarter', default_start=date(2026, 4, 1),
                 default_periods=2)
    with m:
        m.d = mo.time.days
    assert m.default_start == '2026-Q2'
    assert m.d.value == [91, 92]


# ══════════════════════════════════════════════════════════════════════
# default_header — the rows a workbook shows above its sheets
# ══════════════════════════════════════════════════════════════════════


def headed():
    m = model(6)
    with m:
        m.header = mo.MultiVariable('Time')
        with m.header as h:
            h.end = mo.time.end
            h.days = mo.time.days
        m.default_header = m.header
        m.inputs = mo.MultiVariable('Inputs', default_header=None)
        m.inputs.rate = mo.Variable(0.12)
        m.calc = mo.MultiVariable('Calc')
        m.calc.default_header = [m.header.end]
        with m.calc as c:
            c.flow = mo.Variable([10.0] * 6, regrain=mo.up('sum'))
            c.accrual = c.flow * m.inputs.rate * mo.time.days / 365
    return m


def test_the_header_is_a_pointer_not_content():
    m = headed()
    assert m._component_names == ['header', 'inputs', 'calc']
    assert m.default_header is m.header
    assert m.header._python_name == 'header'
    assert m.inputs.default_header is None
    assert m.calc.default_header == [m.header.end]


@pytest.mark.parametrize('bad, error', [
    ([], ValueError),
    ('Time', TypeError),
    ([1, 2], TypeError),
])
def test_a_header_is_rows(bad, error):
    m = model(3)
    with pytest.raises(error, match='default_header'):
        m.default_header = bad


def test_a_loop_container_takes_no_header():
    m = model(3)
    with m:
        m.debt = mo.MultiVariable('Debt', loop=True)
        m.debt.default_header = None
    assert m.debt.default_header is None


# ══════════════════════════════════════════════════════════════════════
# Another grain: time is recomputed, never rolled up
# ══════════════════════════════════════════════════════════════════════


def test_a_quarter_has_the_days_of_a_quarter():
    q = headed().at('quarter')
    assert q.header.days.value == [90, 91]
    assert q.header.end.value == [date(2026, 3, 31), date(2026, 6, 30)]


def test_the_projected_header_is_the_projected_rows():
    q = headed().at('quarter')
    assert q.default_header is q.header
    assert q.calc.default_header == [q.header.end]
    assert q.inputs.default_header is None


def test_a_flow_over_days_sums_its_months():
    m = headed()
    q = m.at('quarter')
    monthly = m.calc.accrual.value
    assert q.calc.accrual.value == pytest.approx([sum(monthly[:3]), sum(monthly[3:])])


def test_a_product_through_an_unnamed_series_sums_its_months():
    # (a * r) is a series with no owner; taken for a constant, the
    # quarter came out as Σa · r · Σg — three times too big here.
    m = model(6)
    with m:
        m.a = mo.Variable([10.0] * 6, regrain=mo.up('sum'))
        m.g = mo.Variable([1.0, 2, 3, 4, 5, 6], regrain=mo.up('sum'))
        m.r = mo.Variable(0.5)
        m.x = m.a * m.r * m.g
    assert m.at('quarter').x.value == [30.0, 75.0]


# ══════════════════════════════════════════════════════════════════════
# Tracked products under a grain change
# ══════════════════════════════════════════════════════════════════════


def tracked():
    m = mo.Model('m', tracks=mo.Tracks('plan', 'actual'), default_grain='month',
                 default_start='2026-01', default_periods=6)
    with m:
        m.rev = mo.Variable(plan=[10., 20, 30, 40, 50, 60],
                            actual=[11., 19, 30, None, None, None], regrain=mo.up('sum'))
        m.qty = mo.Variable(plan=[1., 2, 3, 4, 5, 6], actual=[1., 2, 3, None, None, None],
                            regrain=mo.up('sum'))
        m.fx = mo.Variable([1.0, 1.1, 1.2, 1.3, 1.4, 1.5], regrain=mo.up('mean'))
        m.commission = m.fx * mo.Variable(0.02) * m.rev
        m.named = m.qty * m.rev
    return m


def test_a_tracked_product_sums_each_track_per_quarter():
    m = tracked()
    q = m.at('quarter')
    for row in ('commission', 'named'):
        plan = getattr(m, row).value['plan']
        assert getattr(q, row).value['plan'] == pytest.approx(
            [sum(plan[:3]), sum(plan[3:])]), row
    # The actual track's months are blank in the second quarter: zero there.
    assert q.commission.value['actual'][1] == 0


def test_a_year_window_may_start_at_a_number():
    m = mo.Model('m', default_grain='year', default_start=2026, default_periods=2)
    with m:
        m.d = mo.time.days
    assert m.d.value == [365, 365]


def test_a_class_built_for_another_window_is_refused_where_it_lands():
    class Accrual(mo.MultiVariableClass):
        def compute(self, rate):
            self.acc = rate * mo.time.days / 365

    m = model(6)
    with m:
        m.ops = mo.MultiVariable('Ops', default_grain='month', default_start='2026-04',
                                 default_periods=6)
        with pytest.raises(ValueError, match="reads mo.time.days"):
            m.ops.x = Accrual(rate=0.12)


def test_time_itself_is_read_again_for_the_window_it_joins():
    # Flat statements: no open block names the window.
    m = model(3)
    m.y = mo.MultiVariable('Y', default_grain='year', default_start='2027', default_periods=2)
    m.y.days = mo.time.days
    assert m.y.days.value == [365, 366]


def test_moving_a_window_after_its_lines_read_time_is_refused():
    m = model(3)
    with m:
        m.d = mo.time.days * 1
    m.default_start = '2026-01'                     # the same start: fine
    with pytest.raises(ValueError, match="already read mo.time"):
        m.default_start = '2026-02'


def test_a_window_inside_a_quarter_counts_only_its_own_months():
    m = model(5, '2026-02')
    with m:
        m.d = mo.time.days
        m.acc = mo.Variable(1200.0) * mo.time.days / 365
    q = m.at('quarter')
    assert q.d.value == [28 + 31, 30 + 31 + 30]
    assert q.acc.value == pytest.approx([sum(m.acc.value[:2]), sum(m.acc.value[2:])])


def test_a_start_at_a_time_of_day_is_refused():
    from datetime import datetime
    with pytest.raises(ValueError, match="time of day"):
        mo.Model('m', default_grain='month', default_start=datetime(2026, 1, 1, 12),
                 default_periods=2)


def test_a_variable_holding_a_label_starts_the_window_at_the_label():
    m = mo.Model('m', default_grain='month', default_periods=2)
    with m:
        m.inputs = mo.MultiVariable('Inputs')
        m.inputs.start = mo.Variable('2026-03')
        m.default_start = m.inputs.start
    assert m.default_start == '2026-03'
    assert resolve_default_window(m).start_source is None


def test_a_start_below_the_window_is_refused():
    m = model(3)
    with m:
        m.sub = mo.MultiVariable('Sub')
        with pytest.raises(ValueError, match="declare default_grain"):
            m.sub.default_start = date(2026, 2, 1)


def test_a_coarser_grain_reads_the_start_again_as_a_short_first_period():
    m = pf(date(2026, 2, 1))
    m.default_grain = 'quarter'
    assert (m.default_grain, m.default_start) == ('quarter', '2026-Q1')
    assert m.__dict__['_default_first_day'] == date(2026, 2, 1)


def test_a_grain_the_start_cannot_open_leaves_the_window_as_it_was():
    m = pf(date(2026, 2, 1))
    m.inputs.financial_close._value = date(2026, 2, 15)
    with pytest.raises(ValueError, match="1st of a month"):
        m.default_grain = 'quarter'
    assert (m.default_grain, m.default_start) == ('month', '2026-02')


def test_a_header_lists_rows_not_time_itself():
    m = model(3)
    with pytest.raises(ValueError, match="not a row"):
        m.default_header = [mo.Variable(1.0)]


def test_the_json_renderer_keeps_a_named_time_row_a_reference():
    from modeleon.compile.json.renderer import JsonRenderer
    from modeleon.compile.renderer import RenderCtx
    from modeleon.compile.walker import Walker
    m = model(3)
    with m:
        m.days = mo.time.days
        m.x = m.days * 2
    rendered = Walker(JsonRenderer()).render(m.x._expr, RenderCtx())
    assert {'kind': 'varref', 'path': m.days.id} in rendered.values()



# ══════════════════════════════════════════════════════════════════════
# A section's window is whole, or it is the window above
# ══════════════════════════════════════════════════════════════════════


def test_a_section_naming_only_its_grain_is_refused():
    """It read time off the model's years while its lines stood in
    quarters - now it says what it is missing."""
    m = model(2, '2026', 'year')
    with m:
        m.debt = mo.MultiVariable('Debt', default_grain='quarter')
        with m.debt as d:
            with pytest.raises(ValueError, match="names its grain .quarter. but not its own window"):
                d.days = mo.time.days
            with pytest.raises(ValueError, match="default_start and default_periods"):
                d.x = mo.Variable([1.0] * 8)


def test_a_section_of_the_models_grain_lives_in_its_window():
    m = model(3)
    with m:
        m.s = mo.MultiVariable('Same', default_grain='month')
        with m.s as s:
            s.d = mo.time.days
            s.x = mo.Variable([1.0, 2.0, 3.0])
    assert m.s.d.value == [31, 28, 31]


def test_a_section_may_complete_its_window_before_its_lines():
    m = model(2, '2026', 'year')
    with m:
        m.debt = mo.MultiVariable('Debt', default_grain='quarter')
        m.debt.default_start = '2026-Q1'
        m.debt.default_periods = 8
        with m.debt as d:
            d.days = mo.time.days
    assert len(m.debt.days.value) == 8


def test_the_model_may_name_its_grain_before_its_start():
    m = mo.Model('m', default_grain='month')
    with m:
        m.inputs = mo.MultiVariable('Inputs')
        with m.inputs as i:
            i.close = mo.DATE(2026, 1, 1)
        m.default_start = m.inputs.close
        m.default_periods = 2
        m.s = mo.MultiVariable('S')
        with m.s as s:
            s.d = mo.time.days
    assert m.s.d.value == [31, 28]



def test_what_takes_nothing_from_a_sections_window_joins_it():
    """A scalar, or a line located on its own (an ``.at()`` copy), takes
    nothing from a grain-only section's window - only a series placed by
    it, or a read of ``mo.time``, needs the window."""
    m = model(24)
    with m:
        m.ops = mo.MultiVariable('Ops')
        with m.ops as o:
            o.rev = mo.Variable([10.0] * 24, regrain=mo.up('sum'))
        m.annual = mo.MultiVariable('Annual', default_grain='year')
        with m.annual as y:
            y.rate = mo.Variable(0.2)
            y.rev = m.ops.rev.at('year')
    assert m.annual.rev.value == [120.0, 120.0]


def test_a_section_with_its_start_places_its_series():
    m = model(2, '2026', 'year')
    with m:
        m.debt = mo.MultiVariable('Debt', default_grain='quarter', default_start='2026-Q1')
        with m.debt as d:
            d.x = mo.Variable([1.0] * 8)
    assert len(m.debt.x.value) == 8
