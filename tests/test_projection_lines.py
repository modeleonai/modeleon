# SPDX-License-Identifier: Apache-2.0
"""Projection, line by line.

:func:`project_line` projects one line to a coarser grain through a
per-grain cache shared across calls, and :func:`window_projects` says
whether a model's window can be projected to a grain. Line by line, a
model gives the same numbers, refusals and shapes as
:func:`project_model` — tracked lines included — and the calendar helpers
hand back a fresh answer every time.
"""

from __future__ import annotations

from datetime import date

import pytest

import modeleon as mo
from modeleon.core import time as time_mod
from modeleon.core.projection import project_line, project_model, window_projects
from modeleon.core.tracks import TrackValues
from modeleon.core.variable import Variable


def _model() -> mo.Model:
    m = mo.Model('m', default_grain='month', default_start='2026-01', default_periods=6,
                 tracks=mo.Tracks('plan', 'actual'))
    with m:
        m.inputs = mo.MultiVariable('Inputs')
        with m.inputs as i:
            i.rate = mo.Variable(0.1, display_name='Rate')
            i.price = mo.Variable(plan=[100.0, 110.0, 120.0, 130.0, 140.0, 150.0],
                                  actual=[98.0, 112.0, 118.0, 131.0, 139.0, 152.0],
                                  display_name='Price', regrain=mo.up('mean'))
            i.fx = mo.Variable([1.10, 1.11, 1.12, 1.13, 1.14, 1.15], display_name='FX',
                               regrain=mo.up('last'))
        m.pnl = mo.MultiVariable('PnL')
        with m.pnl as p:
            p.volume = mo.Variable(plan=[10.0, 12.0, 14.0, 16.0, 18.0, 20.0],
                                   actual=[9.0, 13.0, 14.0, 15.0, 19.0, 21.0],
                                   display_name='Volume', regrain=mo.up('sum'))
            p.revenue = mo.Variable(m.inputs.price * p.volume, display_name='Revenue')
            p.usd = mo.Variable(p.revenue / m.inputs.fx, display_name='USD')
            p.bump = mo.Variable(p.volume + mo.Variable(5), display_name='Bump')
    return m


def _values(v: Variable) -> object:
    val = v._value
    if isinstance(val, TrackValues):
        return {r: val[r] for r in val.roles}
    return val


# ─── the calendar helpers remember ────────────────────────────────


def test_the_calendar_helpers_hand_back_fresh_lists() -> None:
    a = time_mod.bucket_ranges('month', 'quarter', '2026-01', 6)
    b = time_mod.bucket_ranges('month', 'quarter', '2026-01', 6)
    assert a == b == [('2026-Q1', 0, 3), ('2026-Q2', 3, 6)]
    assert a is not b
    a.append('poisoned')
    assert time_mod.bucket_ranges('month', 'quarter', '2026-01', 6) == b
    w = time_mod.period_weights('month', '2026-01', 3)
    assert w == [31, 28, 31]
    w[0] = 0
    assert time_mod.period_weights('month', '2026-01', 3) == [31, 28, 31]
    assert time_mod.coarse_first_day('2026-02', 'month', None, 'quarter') == date(2026, 2, 1)
    assert time_mod.coarse_first_day('2026-01', 'month', None, 'quarter') is None
    # A short first period keys separately from a whole one.
    assert time_mod.period_weights('month', '2026-01', 2, date(2026, 1, 15)) != \
        time_mod.period_weights('month', '2026-01', 2)
    with pytest.raises(ValueError):
        time_mod.bucket_ranges('quarter', 'month', '2026-Q1', 2)  # finer: refused, not cached


# ─── window_projects / project_line ───────────────────────────────


def test_window_projects_answers_what_project_model_would_refuse() -> None:
    m = _model()
    assert window_projects(m, 'quarter') and window_projects(m, 'year')
    y = mo.Model('y', default_grain='year', default_start='2026', default_periods=2)
    assert window_projects(y, 'quarter') is False
    with pytest.raises(ValueError):
        project_model(y, 'quarter', on_error='absorb')
    assert window_projects(mo.MultiVariable('bare'), 'quarter')  # no window: nothing to refuse


def test_project_line_is_the_tree_projections_step_for_one_line() -> None:
    m = _model()
    tree = project_model(m, 'quarter', on_error='absorb')
    cache: dict = {}
    for _, line in m.walk():
        if not isinstance(line, Variable):
            continue
        projected = project_line(line, 'quarter', cache, on_error='absorb')
        in_tree = next(v for _, v in tree.walk() if isinstance(v, Variable) and v.id == line.id)
        assert _values(projected) == _values(in_tree), line.id
        assert projected._start == in_tree._start and projected._grain == in_tree._grain
    # The shared cache: an input projected once is the same object its reader saw.
    assert project_line(m.pnl.volume, 'quarter', cache) is cache[id(m.pnl.volume)]


def test_project_line_absorbs_a_refusal_as_the_tree_does() -> None:
    m = mo.Model('m', default_grain='month', default_start='2026-01', default_periods=3)
    with m:
        m.ruleless = mo.Variable([1.0, 2.0, 3.0], display_name='Ruleless')
    with pytest.raises(ValueError, match="has no re-grain rule"):
        project_line(m.ruleless, 'quarter', {})
    row = project_line(m.ruleless, 'quarter', {}, on_error='absorb')
    assert row.value == '#VALUE!' and 'Ruleless' in row._regrain_error
    tree_row = project_model(m, 'quarter', on_error='absorb').ruleless
    assert tree_row.value == '#VALUE!' and tree_row._regrain_error == row._regrain_error


# ─── a tracked line's roles fold as the line would ────────────────


def test_each_role_folds_exactly_as_its_own_series_at_the_grain() -> None:
    m = _model()
    tree = project_model(m, 'quarter')
    for line, rule in ((m.inputs.price, 'mean'), (m.pnl.volume, 'sum')):
        projected = next(v for _, v in tree.walk() if isinstance(v, Variable) and v.id == line.id)
        for role in ('plan', 'actual'):
            own = line.at(track=role).at('quarter')          # the user-facing path
            assert projected._value[role] == own.value, (line._display_name, role, rule)
        assert projected._start == '2026-Q1' and projected._grain == 'quarter'
        assert projected._regrained_tracks is True
    assert tree.pnl.volume._value['plan'] == [36.0, 54.0]
    assert tree.pnl.volume._value['actual'] == [36.0, 55.0]
    # mean is time-weighted: (100*31 + 110*28 + 120*31) / 90
    assert tree.inputs.price._value['plan'][0] == pytest.approx((100 * 31 + 110 * 28 + 120 * 31) / 90)


def test_a_tracked_line_without_a_rule_refuses_by_the_lines_name() -> None:
    m = mo.Model('m', default_grain='month', default_start='2026-01', default_periods=3,
                 tracks=mo.Tracks('plan', 'actual'))
    with m:
        m.heads = mo.Variable(plan=[1.0, 2.0, 3.0], actual=[1.0, None, None], display_name='Heads')
    with pytest.raises(ValueError, match="'Heads' has no re-grain rule and is not a formula"):
        project_model(m, 'quarter')
    row = project_model(m, 'quarter', on_error='absorb').heads
    assert isinstance(row._value, TrackValues)
    assert row._value['plan'] == '#VALUE!' and row._value['actual'] == '#VALUE!'
    assert "'Heads' has no re-grain rule" in row._regrain_error


def test_a_scalar_role_broadcasts_at_every_grain() -> None:
    m = mo.Model('m', default_grain='month', default_start='2026-01', default_periods=3,
                 tracks=mo.Tracks('plan', 'actual'))
    with m:
        m.fee = mo.Variable(plan=[10.0, 20.0, 30.0], actual=7.0, display_name='Fee',
                            regrain=mo.up('sum'))
        m.fixed = mo.Variable(plan=[1.0, 1.0, 1.0], actual=2.0, display_name='Fixed',
                              regrain=mo.up('sum'), grain='month', start='2026-01')
    tree = project_model(m, 'quarter')
    assert tree.fee._value['plan'] == [60.0] and tree.fee._value['actual'] == 7.0
    # A line with its own placement: the scalar role is still one number.
    assert tree.fixed._value['plan'] == [3.0] and tree.fixed._value['actual'] == 2.0


def test_a_frozen_tracked_line_with_a_value_rule_projects_its_values() -> None:
    m = mo.Model('m', default_grain='month', default_start='2026-01', default_periods=6,
                 tracks=mo.Tracks('plan', 'actual'))
    with m:
        m.cash = mo.Variable(plan=[1.0, 2.0, 3.0, 4.0, 5.0, 6.0], actual=[1.5, 2.5, 3.5, 4.5, 5.5, 6.5],
                             display_name='Cash', regrain=mo.frozen(values='last'))
    tree = project_model(m, 'quarter')
    assert tree.cash._value['plan'] == [3.0, 6.0] and tree.cash._value['actual'] == [3.5, 6.5]
    assert tree.cash._value['plan'] == m.cash.at(track='plan').at('quarter').value
