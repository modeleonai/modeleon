# SPDX-License-Identifier: Apache-2.0
"""Rows that read each other across periods — a loop container.

A debt block is the canonical case: interest reads last period's
balance, the balance reads this period's sweep, the sweep reads the
cash left after interest. Row by row that is a circle; period by period
it is a plain order of evaluation, and a hand-built workbook has always
had exactly this shape — a cell standing on the next row's previous
column. In a container built with ``loop=True`` a row may be read before
the line that defines it; these tests pin what that allows, what it
refuses, that nothing changes for any other container, and that the
workbook comes out live.
"""

import shutil
import subprocess
import threading
import warnings
from datetime import date
from pathlib import Path

import pytest
from openpyxl import load_workbook

import modeleon as mo
from modeleon import ForwardReferenceError, SamePeriodCycleError


def series(var):
    return [round(x, 6) for x in var.value]


def model(n=4):
    return mo.Model("M", default_grain="month", default_start="2026-01",
                    default_periods=n)


def debt_control(n, opening=1000.0, rate=0.01, cash=50.0, scheduled=10.0, share=0.5):
    """The debt block by hand: the oracle every loop test compares to."""
    out = {"interest": [], "pot": [], "sweep": [], "balance": []}
    balance = opening
    for _ in range(n):
        prev = balance
        interest = prev * rate
        pot = cash - interest - scheduled
        sweep = (pot if pot > 0 else 0.0) * share
        balance = prev - scheduled - sweep
        for key, value in (("interest", interest), ("pot", pot),
                           ("sweep", sweep), ("balance", balance)):
            out[key].append(round(value, 6))
    return out


def debt_model(n=6):
    m = model(n)
    with m:
        m.inp = mo.MultiVariable("Inputs")
        with m.inp as i:
            i.rate = mo.Variable(0.01)
            i.cash = mo.Variable([50.0] * n, unit="USD")
            i.scheduled = mo.Variable([10.0] * n, unit="USD")
            i.share = mo.Variable(0.5)
            i.opening = mo.Variable(1000.0, unit="USD")
        m.debt = mo.MultiVariable("Debt", loop=True)
        with m.debt as d:
            d.interest = mo.lag(d.balance, fill=m.inp.opening) * m.inp.rate
            d.pot = m.inp.cash - d.interest - m.inp.scheduled
            d.sweep = mo.IF(d.pot > 0, d.pot, 0.0) * m.inp.share
            d.balance = (mo.lag(d.balance, fill=m.inp.opening)
                         - m.inp.scheduled - d.sweep)
    return m


# ══════════════════════════════════════════════════════════════════════
# The values
# ══════════════════════════════════════════════════════════════════════


def test_debt_loop_matches_a_hand_loop():
    m = debt_model(6)
    control = debt_control(6)
    for key in control:
        assert series(getattr(m.debt, key)) == control[key], key


def test_the_rows_carry_units_derived_from_the_closed_loop():
    m = debt_model(4)
    for key in ("interest", "pot", "sweep", "balance"):
        assert str(getattr(m.debt, key)._unit) == "USD", key


def test_the_formula_reads_as_the_author_wrote_it():
    m = debt_model(4)
    assert m.debt.interest.formula == "(balance.shift(1, fill_value=opening)) * rate"
    assert m.debt.balance.is_formula


def test_two_rows_that_read_each_other_across_a_period():
    flow = [10.0, 20.0, 30.0, 40.0, 50.0]
    m = model(5)
    with m:
        m.flow = mo.Variable(flow)
        m.x = mo.MultiVariable("X", loop=True)
        with m.x as x:
            x.b = mo.lag(x.a, fill=100.0) * 0.5 + m.flow
            x.a = mo.lag(x.a, fill=100.0) + x.b
    a, ctrl_a, ctrl_b = 100.0, [], []
    for t in range(5):
        b = a * 0.5 + flow[t]
        a = a + b
        ctrl_a.append(round(a, 6))
        ctrl_b.append(round(b, 6))
    assert series(m.x.a) == ctrl_a
    assert series(m.x.b) == ctrl_b


def test_rows_may_come_in_any_order():
    """In a loop container the engine orders the rows: a same-period read
    before its line is as fine as a lagged one."""
    m = model(6)
    with m:
        m.inp = mo.MultiVariable("Inputs")
        with m.inp as i:
            i.rate = mo.Variable(0.01)
            i.cash = mo.Variable([50.0] * 6)
            i.scheduled = mo.Variable([10.0] * 6)
            i.share = mo.Variable(0.5)
        m.debt = mo.MultiVariable("Debt", loop=True)
        with m.debt as d:
            d.balance = mo.lag(d.balance, fill=1000.0) - m.inp.scheduled - d.sweep
            d.sweep = mo.IF(d.pot > 0, d.pot, 0.0) * m.inp.share
            d.pot = m.inp.cash - d.interest - m.inp.scheduled
            d.interest = mo.lag(d.balance, fill=1000.0) * m.inp.rate
    control = debt_control(6)
    for key in control:
        assert series(getattr(m.debt, key)) == control[key], key


def test_the_loop_waits_for_its_last_row():
    m = model(4)
    with m:
        m.x = mo.MultiVariable("X", loop=True)
        with m.x as x:
            x.a = mo.lag(x.b) + mo.lag(x.c, fill=1.0)
            x.b = x.a + 1.0
            with pytest.raises(ValueError, match="not computed yet"):
                x.b.value
            x.c = x.a * 2.0
    a, b, c = None, 0.0, 1.0
    ctrl = {"a": [], "b": [], "c": []}
    for _ in range(4):
        a = b + c
        b = a + 1.0
        c = a * 2.0
        for key, value in (("a", a), ("b", b), ("c", c)):
            ctrl[key].append(round(value, 6))
    for key in ctrl:
        assert series(getattr(m.x, key)) == ctrl[key], key


def test_a_row_read_ahead_may_turn_out_to_be_an_input():
    m = model(6)
    with m:
        m.x = mo.MultiVariable("X", loop=True)
        with m.x as x:
            x.a = mo.lag(x.b) * 2.0
            x.b = mo.Variable([1.0, 2.0, 3.0, 4.0, 5.0, 6.0])
    assert series(m.x.a) == [0.0, 2.0, 4.0, 6.0, 8.0, 10.0]
    assert m.x.b.is_formula is False


def test_rows_downstream_written_before_the_loop_closes():
    m = model(4)
    with m:
        m.x = mo.MultiVariable("X", loop=True)
        with m.x as x:
            x.a = mo.lag(x.b) + 1.0
            x.total = mo.SUM(x.a)
            x.doubled = x.a * 2.0
            x.running = mo.cumsum(x.a)
            x.ahead = mo.lag(x.a, -1)
            x.b = x.a * 2.0
    assert series(m.x.a) == [1.0, 3.0, 7.0, 15.0]
    assert m.x.total.value == pytest.approx(26.0)
    assert series(m.x.doubled) == [2.0, 6.0, 14.0, 30.0]
    assert series(m.x.running) == [1.0, 4.0, 11.0, 26.0]
    assert series(m.x.ahead) == [3.0, 7.0, 15.0, 0.0]


def test_a_raw_list_operand_inside_the_loop_is_a_series():
    m = model(4)
    with m:
        m.x = mo.MultiVariable("X", loop=True)
        with m.x as x:
            x.a = mo.lag(x.a) + [1.0, 2.0, 3.0, 4.0]
    assert series(m.x.a) == [1.0, 3.0, 6.0, 10.0]


def test_a_loop_of_dates():
    m = model(4)
    with m:
        m.x = mo.MultiVariable("X", loop=True)
        with m.x as x:
            x.pay = mo.EDATE(mo.lag(x.pay, fill=date(2026, 1, 31)), 1)
    assert m.x.pay.value == [date(2026, 2, 28), date(2026, 3, 28),
                             date(2026, 4, 28), date(2026, 5, 28)]
    assert m.x.pay.value_type == "datetime"


def test_the_horizon_is_one_number():
    for n in (48, 60):
        m = debt_model(n)
        assert len(m.debt.balance.value) == n
        assert series(m.debt.balance) == debt_control(n)["balance"]


def test_a_loop_class():
    class Tranche(mo.MultiVariableClass, loop=True):
        def compute(self, opening, rate, repay):
            self.interest = mo.lag(self.balance, fill=opening) * rate
            self.balance = mo.lag(self.balance, fill=opening) - repay + self.interest

    def control(opening, rate, repay):
        balance, out = opening, []
        for _ in range(4):
            balance = balance - repay + balance * rate
            out.append(round(balance, 6))
        return out

    m = model(4)
    with m:
        m.senior = Tranche(opening=1000.0, rate=0.01, repay=100.0)
        m.junior = Tranche(opening=500.0, rate=0.02, repay=50.0)
    assert series(m.senior.balance) == control(1000.0, 0.01, 100.0)
    assert series(m.junior.balance) == control(500.0, 0.02, 50.0)

    # Built without a `with` block: it borrows the horizon of the model
    # given one last.
    other = model(4)
    other.t = Tranche(opening=1000.0, rate=0.01, repay=100.0)
    assert series(other.t.balance) == control(1000.0, 0.01, 100.0)


def test_running_the_cell_again_builds_the_loop_again():
    """A notebook cell run twice: the container is built in the same cell,
    so each run solves the loop from scratch."""
    m = model(4)

    def cell(draw):
        with m:
            m.x = mo.MultiVariable("X", loop=True)
            with m.x as x:
                x.draw = mo.Variable(draw)
                x.interest = mo.lag(x.balance) * 0.1
                x.balance = mo.lag(x.balance) + x.draw + x.interest

    def control(draw):
        balance, out = 0.0, []
        for t in range(4):
            balance = balance + draw[t] + balance * 0.1
            out.append(round(balance, 6))
        return out

    cell([100.0, 50.0, 0.0, 0.0])
    assert series(m.x.balance) == control([100.0, 50.0, 0.0, 0.0])
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")      # the old container is replaced
        cell([200.0, 50.0, 0.0, 0.0])
    assert series(m.x.balance) == control([200.0, 50.0, 0.0, 0.0])


def test_a_loop_is_not_changed_one_row_at_a_time():
    m = debt_model(4)
    with pytest.raises(ValueError, match="cannot be changed one row at a time"):
        with m.debt as d:
            d.interest = mo.lag(d.balance, fill=m.inp.opening) * 0.02
    assert series(m.debt.balance) == debt_control(4)["balance"]


def test_a_probe_does_not_hold_a_name():
    m = model(4)
    with m:
        m.x = mo.MultiVariable("X", loop=True)
        with m.x as x:
            assert hasattr(x, "rate")          # a loop container answers any name
            x.rate = mo.Variable(0.05)         # ... and the name stays free
            x.a = mo.lag(x.b) + x.rate
            x.b = x.a * 2.0
    assert m.x.rate.value == 0.05


def test_a_helper_may_enter_the_same_block_again():
    m = model(4)
    with m:
        m.x = mo.MultiVariable("X", loop=True)

        def fee(block):
            with block as b:
                b.fee = mo.Variable([1.0] * 4)

        with m.x as x:
            x.a = mo.lag(x.b) + 1.0
            fee(x)
            x.b = x.a + x.fee
    assert series(m.x.b) == [2.0, 4.0, 6.0, 8.0]


def test_a_class_built_outside_its_model_joins_one_of_the_same_horizon():
    class Tranche(mo.MultiVariableClass, loop=True):
        def compute(self, opening, repay):
            self.balance = mo.lag(self.balance, fill=opening) - repay

    first = model(4)                                # noqa: F841 — the model built last
    built = Tranche(opening=100.0, repay=10.0)     # borrows its 4 periods
    other = model(6)
    with pytest.raises(ValueError, match="over 4 period.*model of 6"):
        other.t = built


def test_units_wait_for_the_loop_to_close():
    m = model(4)
    with m:
        m.sched = mo.Variable([10.0] * 4, unit="USD")
        m.x = mo.MultiVariable("X", loop=True)
        with m.x as x:
            x.share = 1 + m.sched / mo.lag(x.balance, fill=100.0)
            x.ratio = m.sched / mo.lag(x.balance, fill=100.0)
            x.balance = mo.lag(x.balance, fill=100.0) - m.sched
    assert m.x.balance._unit is not None and str(m.x.balance._unit) == "USD"
    assert m.x.ratio._unit is None
    assert m.x.share._unit is None


def test_a_thrown_away_expression_does_not_stop_the_loop():
    m = model(4)
    with m:
        m.x = mo.MultiVariable("X", loop=True)
        with m.x as x:
            x.a = mo.lag(x.b) + 1.0
            with pytest.raises(TypeError):
                x.a["2026-02"]
            x.b = x.a * 2.0
    assert series(m.x.a) == [1.0, 3.0, 7.0, 15.0]


# ══════════════════════════════════════════════════════════════════════
# What is refused, by name
# ══════════════════════════════════════════════════════════════════════


def test_a_value_is_refused_until_the_loop_closes():
    m = model(4)
    with m:
        m.x = mo.MultiVariable("X", loop=True)
        with m.x as x:
            x.a = mo.lag(x.b) + 1.0
            with pytest.raises(ValueError, match="reads 'b' before it is assigned"):
                x.a.value
            with pytest.raises(ValueError, match="before it is assigned"):
                bool(x.a > 0)
            with pytest.raises(ValueError, match="before it is assigned"):
                x.a[0].value
            assert "<pending>" in repr(x.a)
            x.b = x.a * 2.0
    assert series(m.x.a) == [1.0, 3.0, 7.0, 15.0]


def test_a_cycle_inside_one_period_is_refused_by_name():
    m = model(4)
    with pytest.raises(SamePeriodCycleError, match="WITHIN one period") as err:
        with m:
            m.x = mo.MultiVariable("X", loop=True)
            with m.x as x:
                x.a = x.b + 1.0
                x.b = x.a * 2.0
    assert {v.python_name for v in err.value.rows} == {"a", "b"}
    assert "mo.lag" in str(err.value)


def test_the_cycle_names_rows_not_intermediates():
    m = model(4)
    with pytest.raises(SamePeriodCycleError) as err:
        with m:
            m.x = mo.MultiVariable("X", loop=True)
            with m.x as x:
                x.a = (x.b + 1.0) * 2.0
                x.b = (x.a - 1.0) / 2.0
    assert sorted(v.python_name for v in err.value.rows) == ["a", "b"]
    assert len(err.value.nodes) > len(err.value.rows)


def test_a_misspelt_row_is_reported_with_the_line_that_read_it():
    m = model(4)
    with pytest.raises(ForwardReferenceError, match="'balanse', read at .*test_loops.py:") as err:
        with m:
            m.x = mo.MultiVariable("X", loop=True)
            with m.x as x:
                x.interest = mo.lag(x.balanse) * 0.01
                x.balance = mo.lag(x.balance, fill=100.0) - x.interest
    assert "never assigned" in str(err.value)


def test_an_error_inside_the_block_is_not_masked():
    m = model(4)
    with pytest.raises(ZeroDivisionError):
        with m:
            m.x = mo.MultiVariable("X", loop=True)
            with m.x as x:
                x.interest = mo.lag(x.balance) * 0.01
                1 / 0


def test_a_lead_into_a_loop_row_is_refused():
    m = model(4)
    with pytest.raises(ValueError, match="one period AHEAD"):
        with m:
            m.x = mo.MultiVariable("X", loop=True)
            with m.x as x:
                x.a = mo.lag(x.b, -1) + mo.lag(x.b)
                x.b = mo.lag(x.a) * 2.0


def test_a_reduction_over_a_loop_row_inside_the_loop_is_refused():
    m = model(4)
    with pytest.raises(ValueError, match="MAX reads a whole series") as err:
        with m:
            m.x = mo.MultiVariable("X", loop=True)
            with m.x as x:
                x.a = mo.lag(x.b) + 1.0
                x.floor = mo.MAX(x.a, 0.0)
                x.b = x.a - x.floor
    assert "mo.IF" in str(err.value)
    # The rows it left behind say why they have no value.
    with pytest.raises(ValueError, match="loop could not be computed"):
        m.x.a.value


@pytest.mark.parametrize("value, message", [
    ("container", "assigned a MultiVariable"),
    ("number", "assigned a float"),
    ("single", "lag reads a series"),
])
def test_a_read_row_assigned_something_that_is_not_a_series_is_refused(value, message):
    m = model(4)
    with pytest.raises(ForwardReferenceError, match=message):
        with m:
            m.x = mo.MultiVariable("X", loop=True)
            with m.x as x:
                x.a = mo.lag(x.b) + 1.0
                if value == "container":
                    x.b = mo.MultiVariable("B")
                elif value == "number":
                    x.b = 5.0
                else:
                    x.b = mo.Variable(5.0)


def test_a_placeholder_cannot_be_assigned_under_another_name():
    m = model(4)
    with pytest.raises(ForwardReferenceError, match="'alias' is assigned 'b'"):
        with m:
            m.x = mo.MultiVariable("X", loop=True)
            with m.x as x:
                x.a = mo.lag(x.b) + 1.0
                x.alias = x.b
                x.b = x.a * 2.0


def test_a_fill_assigned_further_down_is_refused_with_its_name():
    m = model(4)
    with pytest.raises(TypeError, match="fill=opening"):
        with m:
            m.x = mo.MultiVariable("X", loop=True)
            with m.x as x:
                x.interest = mo.lag(x.balance, fill=x.opening) * 0.01
                x.opening = mo.Variable(100.0)
                x.balance = mo.lag(x.balance) - x.interest


def test_a_unit_mix_found_when_the_loop_closes_names_the_row():
    m = model(4)
    with m:
        m.fee = mo.Variable([1.0] * 4, unit="EUR")
        m.repay = mo.Variable([1.0] * 4, unit="USD")
        m.x = mo.MultiVariable("X", loop=True)
    with pytest.raises(ValueError, match="cost.*incompatible units"):
        with m.x as x:
            x.cost = mo.lag(x.balance, fill=100.0) + m.fee
            x.balance = mo.lag(x.balance, fill=100.0) - m.repay


def test_python_that_cannot_be_rerun_is_refused_over_an_open_loop():
    m = model(4)
    with m:
        m.x = mo.MultiVariable("X", loop=True)
        with m.x as x:
            x.a = mo.lag(x.b) + 1.0
            with pytest.raises(ValueError, match="pyformula"):
                mo.Variable(pyformula=x.a)
            x.b = x.a * 2.0


def test_a_loop_needs_the_models_horizon():
    root = mo.MultiVariable("root", loop=True)
    with pytest.raises(AttributeError, match="default_periods"):
        with root:
            root.a = mo.lag(root.b) + 1.0


def test_numbers_are_refused_while_the_loop_is_open():
    m = model(4)
    with m:
        m.x = mo.MultiVariable("X", loop=True)
        with m.x as x:
            x.a = mo.lag(x.b, fill=1.0) + 1.0
            for convert in (float, int):
                with pytest.raises(ValueError, match="before it is assigned"):
                    convert(x.a[1])
            with pytest.raises(ValueError, match="before it is assigned"):
                x.a.at("quarter")
            assert m.x.to_dict()["components_data"]["a"]["values"] is None
            x.b = x.a * 2.0
    assert m.x.to_dict()["components_data"]["a"]["values"] == m.x.a.value


def test_rows_reading_a_failed_loop_say_why():
    m = model(4)
    with pytest.raises(ValueError):
        with m:
            m.x = mo.MultiVariable("X", loop=True)
            with m.x as x:
                x.a = mo.lag(x.b) + 1.0
                x.floor = mo.MAX(x.a, 0.0)
                x.b = x.a - x.floor
    m.after = m.x.b + 1.0
    with pytest.raises(ValueError, match="loop could not be computed"):
        m.after.value


def test_the_cycle_error_travels_between_processes():
    import pickle
    m = model(4)
    with pytest.raises(SamePeriodCycleError) as err:
        with m:
            m.x = mo.MultiVariable("X", loop=True)
            with m.x as x:
                x.a = x.b + 1.0
                x.b = x.a * 2.0
    again = pickle.loads(pickle.dumps(err.value))
    assert str(again) == str(err.value)
    assert [v.python_name for v in again.rows] == [v.python_name for v in err.value.rows]


# ══════════════════════════════════════════════════════════════════════
# Nothing changes anywhere else
# ══════════════════════════════════════════════════════════════════════


def test_an_ordinary_container_answers_as_it_always_has():
    m = model(4)
    with m:
        m.x = mo.MultiVariable("X", loop=True)
        with m.x as x:
            x.a = mo.lag(x.b) + 1.0
            # Another container — even inside the loop block.
            assert not hasattr(m, "sub")
            assert getattr(m, "optional", None) is None
            with pytest.raises(AttributeError, match="has no component 'b'"):
                m.b
            x.b = x.a * 2.0
    assert not hasattr(m, "b")


def test_a_row_named_loop_is_a_row():
    x = mo.MultiVariable("Ops", loop=mo.Variable(5.0), other=mo.Variable(1.0))
    assert x._component_names == ["loop", "other"]
    assert not hasattr(x, "missing")          # not a loop container
    with pytest.raises(TypeError, match="loop="):
        type("Bad", (mo.MultiVariableClass,), {}, loop="yes")


def test_the_engines_own_probes_create_no_rows(tmp_path):
    m = debt_model(4)
    m.to_excel(tmp_path / "debt.xlsx")
    with warnings.catch_warnings():
        # The block alone has no cells for the inputs it reads; that is
        # what a sub-block's own rendering always says.
        warnings.simplefilter("ignore")
        m.debt._repr_html_()
    m.debt.to_dict()
    assert m.debt._component_names == ["interest", "pot", "sweep", "balance"]
    assert not m.debt.__dict__.get("_forward")


def test_code_trusted_like_the_engine_probes_without_creating_rows(tmp_path, monkeypatch):
    import importlib.util
    import os

    from modeleon.core import mv_context
    monkeypatch.setattr(mv_context, "_TRUSTED_DIRS", list(mv_context._TRUSTED_DIRS))
    ext = tmp_path / "inspector"
    ext.mkdir()
    source = ext / "probe.py"
    source.write_text("def probe(mv, name):\n    return hasattr(mv, name)\n")
    spec = importlib.util.spec_from_file_location("probe", str(source))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)

    m = model(4)
    m.x = mo.MultiVariable("X", loop=True)
    assert module.probe(m.x, "missing")        # user code: any name answers
    mv_context.trust_code_in(str(ext))
    mv_context.trust_code_in(str(ext))
    assert mv_context._TRUSTED_DIRS.count(os.path.join(str(ext), "")) == 1
    m.y = mo.MultiVariable("Y", loop=True)
    assert not module.probe(m.y, "missing")    # trusted code: a plain miss
    assert not m.y.__dict__.get("_forward")


def test_a_failed_block_leaves_nothing_for_the_next_model():
    m = model(4)
    with pytest.raises(ZeroDivisionError):
        with m:
            m.x = mo.MultiVariable("X", loop=True)
            with m.x as x:
                x.a = mo.lag(x.b) + 1.0
                1 / 0
    fresh = debt_model(4)
    assert series(fresh.debt.balance) == debt_control(4)["balance"]


def test_a_loop_in_another_thread():
    results = {}

    def build():
        results["balance"] = series(debt_model(4).debt.balance)

    t = threading.Thread(target=build)
    t.start()
    t.join()
    assert results["balance"] == debt_control(4)["balance"]


def test_a_loop_is_not_a_cycle_to_the_dependency_check():
    m = debt_model(4)
    assert m.detect_cycles() is None
    m.assert_acyclic()
    assert "M.debt.sweep" in m.debt.balance.dependencies


def test_no_structure_warning_is_raised_when_the_loop_closes():
    with warnings.catch_warnings():
        warnings.simplefilter("error")
        debt_model(4)


def test_the_model_at_another_grain_names_the_loop():
    m = model(6)
    with m:
        m.flow = mo.Variable([1.0, 2.0, 3.0, 4.0, 5.0, 6.0], regrain=mo.up("sum"))
        m.x = mo.MultiVariable("X", loop=True)
        with m.x as x:
            x.interest = mo.lag(x.balance, fill=100.0) * 0.1
            x.balance = mo.lag(x.balance, fill=100.0) + m.flow + x.interest
    with pytest.raises(ValueError, match="form a loop across periods") as err:
        m.at("quarter")
    assert "'interest'" in str(err.value) and "'balance'" in str(err.value)
    m.x.interest.set_regrain(mo.frozen(values="sum"))
    m.x.balance.set_regrain(mo.frozen(values="last"))
    q = m.at("quarter")
    monthly = series(m.x.balance)
    assert [round(v, 6) for v in q.x.balance.value] == [monthly[2], monthly[5]]


# ══════════════════════════════════════════════════════════════════════
# Flat statements, without `with` blocks
# ══════════════════════════════════════════════════════════════════════


def test_flat_statements_need_no_with_block():
    m = model(4)
    m.debt = mo.MultiVariable("Debt", loop=True)
    m.debt.interest = mo.lag(m.debt.balance, fill=1000.0) * 0.01
    m.debt.balance = mo.lag(m.debt.balance, fill=1000.0) - 10.0 - m.debt.interest
    balance, ctrl = 1000.0, []
    for _ in range(4):
        balance = balance - 10.0 - balance * 0.01
        ctrl.append(round(balance, 6))
    assert series(m.debt.balance) == ctrl


def test_flat_statements_report_an_unassigned_row_on_use(tmp_path):
    m = model(4)
    m.debt = mo.MultiVariable("Debt", loop=True)
    m.debt.interest = mo.lag(m.debt.balanse) * 0.01
    with pytest.raises(ValueError, match="reads 'balanse'"):
        m.debt.interest.value
    with pytest.raises(ValueError, match="balanse"):
        m.to_excel(tmp_path / "early.xlsx")


# ══════════════════════════════════════════════════════════════════════
# The workbook
# ══════════════════════════════════════════════════════════════════════


def _rows(ws):
    out = {}
    for row in ws.iter_rows():
        label = row[0].value
        if label:
            out[str(label)] = [c.value for c in row[1:]]
    return out


def test_every_cell_is_a_live_formula_on_the_previous_column(tmp_path):
    m = debt_model(4)
    path = tmp_path / "debt.xlsx"
    m.to_excel(path)
    debt = _rows(load_workbook(path)["Debt"])
    interest, pot, sweep, balance = (
        debt["Interest (USD)"], debt["Pot (USD)"], debt["Sweep (USD)"],
        debt["Balance (USD)"],
    )
    # Interest stands on the balance row (row 5) one column left; the
    # opening balance is the input cell in the head column.
    assert interest[:3] == ["=Inputs!B6 * Inputs!B2", "=B5 * Inputs!B2",
                            "=C5 * Inputs!B2"]
    assert balance[:3] == ["=Inputs!B6 - Inputs!C4 - B4", "=B5 - Inputs!D4 - C4",
                           "=C5 - Inputs!E4 - D4"]
    assert pot[1] == "=Inputs!D3 - C2 - Inputs!D4"
    assert sweep[1] == "=IF(C3 > 0, C3, 0.0) * Inputs!B5"
    for cells in (interest, pot, sweep, balance):
        assert all(str(c).startswith("=") for c in cells[:4])


def test_a_workbook_written_inside_an_open_loop_is_refused(tmp_path):
    m = model(4)
    with m:
        m.x = mo.MultiVariable("X", loop=True)
        with m.x as x:
            x.a = mo.lag(x.b) + 1.0
            with pytest.raises(ValueError, match="before it is assigned"):
                m.to_excel(tmp_path / "early.xlsx")
            x.b = x.a * 2.0
    m.to_excel(tmp_path / "late.xlsx")


def _soffice():
    for candidate in ("soffice", "libreoffice"):
        found = shutil.which(candidate)
        if found:
            return found
    mac = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    return str(mac) if mac.exists() else None


@pytest.mark.slow
def test_the_workbook_recalculates_to_the_python_values(tmp_path):
    soffice = _soffice()
    if soffice is None:
        pytest.skip("LibreOffice not installed")
    m = debt_model(6)
    path = tmp_path / "debt.xlsx"
    m.to_excel(path)
    subprocess.run(
        [soffice, f"-env:UserInstallation={(tmp_path / 'prof').as_uri()}",
         "--headless", "--convert-to", "xlsx", "--outdir",
         str(tmp_path / "calc"), str(path)],
        check=True, capture_output=True,
    )
    debt = _rows(load_workbook(tmp_path / "calc" / "debt.xlsx", data_only=True)["Debt"])
    control = debt_control(6)
    for key, label in (("interest", "Interest (USD)"), ("pot", "Pot (USD)"),
                       ("sweep", "Sweep (USD)"), ("balance", "Balance (USD)")):
        got = [round(v, 6) for v in debt[label][:6]]
        assert got == control[key], key


def test_a_loop_container_mounted_twice_stays_a_loop():
    """A container mounted under a second parent is copied; the copy of a
    ``loop=True`` container lost the flag, so a row read ahead inside the
    copy failed with an unrelated ``AttributeError``."""
    from modeleon.core.mv_context import is_loop_container

    m = debt_model(4)
    with m:
        m.summary = mo.MultiVariable("Summary")
        with m.summary as s:
            s.debt = m.debt
    copy = m.summary.debt
    assert copy is not m.debt
    assert is_loop_container(copy)
    assert series(copy.balance) == series(m.debt.balance)
    for block in (m.debt, copy):
        with block as d:
            d.fee = mo.lag(d.total, fill=0.0) * 0.01
            d.total = d.balance + d.fee
        assert series(block.total) == series(m.debt.total)


def test_the_dependents_of_a_placeholder_are_each_held_once_in_order():
    import pickle

    from modeleon.core.loops import Dependents

    a, b, c = mo.Variable(1.0), mo.Variable(2.0), mo.Variable(3.0)
    held = Dependents()
    for v in (a, b, a, c, b):
        held.append(v)
    assert list(held) == [a, b, c]
    assert b in held and mo.Variable(2.0) not in held
    assert not Dependents() and len(held) == 3
    back = pickle.loads(pickle.dumps(held))
    assert [v.value for v in back] == [1.0, 2.0, 3.0]
    back.append(next(iter(back)))
    assert len(back) == 3


def test_a_block_of_many_rows_reading_each_other_builds_in_time():
    """A hundred loans under one cash sweep, every loan's balance read through
    mo.lag by the sweep: each row that lands handed its dependents to every
    row it waits on with a scan per pair — over ten seconds. Under one now;
    the bound leaves a margin of ten for a slow machine."""
    import time

    n, periods = 100, 20
    started = time.perf_counter()
    m = mo.Model("m", default_grain="year", default_start="2026", default_periods=periods)
    with m:
        m.inp = mo.MultiVariable("Inputs")
        with m.inp as i:
            i.cash = mo.Variable([50.0] * periods)
        m.debt = mo.MultiVariable("Debt", loop=True)
        with m.debt as d:
            for k in range(n):
                setattr(d, f"open_{k}", mo.lag(getattr(d, f"bal_{k}"), fill=100.0))
            d.total_open = sum(getattr(d, f"open_{k}") for k in range(n))
            d.sweep = mo.IF(i.cash < d.total_open, i.cash, d.total_open)
            for k in range(n):
                setattr(d, f"bal_{k}", getattr(d, f"open_{k}")
                        - d.sweep * getattr(d, f"open_{k}") / d.total_open)
    elapsed = time.perf_counter() - started
    assert series(m.debt.bal_0)[:3] == [99.5, 99.0, 98.5]
    assert elapsed < 3.0, elapsed


def test_a_declared_unit_survives_the_closed_loop():
    """A loop's rows are built on placeholders and their units derived
    again when it closes; a unit the author declared is kept. A label on a
    pure number does not survive a product, so it cannot be derived."""
    factor = mo.Unit.dimensionless("factor")
    m = model(4)
    with m:
        m.inp = mo.MultiVariable("Inputs")
        with m.inp as i:
            i.rate = mo.Variable(0.1)
            i.cash = mo.Variable([100.0] * 4, unit="USD")
        m.dcf = mo.MultiVariable("DCF", loop=True)
        with m.dcf as d:
            d.df = (mo.lag(d.step, fill=1.0) / (1 + m.inp.rate)).set_unit(factor)
            d.step = (d.df * 1.0).set_unit(factor)
            d.pv = m.inp.cash * d.df
    assert str(m.dcf.df._unit) == "factor"
    assert str(m.dcf.step._unit) == "factor"
    assert str(m.dcf.pv._unit) == "USD"


def test_a_declared_unit_does_not_excuse_its_formula():
    """A declared unit decides what a loop row prints; an addition of two
    currencies inside its formula is still refused when the loop closes."""
    m = model(4)
    with m:
        m.inp = mo.MultiVariable("Inputs")
        with m.inp as i:
            i.cash = mo.Variable([100.0] * 4, unit="USD")
            i.fx = mo.Variable([0.9] * 4, unit="EUR/USD")
        m.debt = mo.MultiVariable("Debt", loop=True)
        with pytest.raises(ValueError, match="incompatible units"):
            with m.debt as d:
                d.pot = (m.inp.cash + mo.lag(d.eur, fill=0.0)).set_unit("USD")
                d.eur = d.pot * m.inp.fx


def test_an_undeclared_loop_row_prints_the_label_its_formula_gives():
    """Pure numbers add, the left label printing - inside a loop as outside."""
    share = mo.Unit.dimensionless("share")
    factor = mo.Unit.dimensionless("factor")
    m = model(4)
    with m:
        m.inp = mo.MultiVariable("Inputs")
        with m.inp as i:
            i.sh = mo.Variable([0.1] * 4, unit=share)
        m.x = mo.MultiVariable("X", loop=True)
        with m.x as d:
            d.x = mo.lag(d.y, fill=0.0) + m.inp.sh
            d.y = (d.x * 0.5).set_unit(factor)
    assert str(m.x.x._unit) == "factor"
