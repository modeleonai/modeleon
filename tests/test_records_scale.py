# SPDX-License-Identifier: Apache-2.0
"""A table of many records is built and written in time linear in its size.

A container's children live in its ``__dict__``; ``_components`` is a
dict snapshot of them — a copy of every child, fresh on every read. The
adoption check, the path crystallization and every walk of the layout
and the writer read one child at a time through that snapshot, so a
records table of n rows copied the table n times: 10,000 records took
minutes to write. A child is now read by name (``_component``), its
presence asked of a name list whose ``in`` is a set lookup
(``_has_component``), and the snapshot stays what it was for every
other reader — a fresh dict.

The scale tests count the entries the snapshot copies, never the
clock: a table twice as long must cost twice the work, not four times.
"""

import copy
import pickle

import pytest

import modeleon as mo
from modeleon.core.multi_variable import MultiVariableBase
from modeleon.core.mv_context import ModelStructureWarning


class Tx(mo.MultiVariableClass):
    def compute(self, day, amount):
        self.day = mo.Variable(day, display_name="Day")
        self.amount = mo.Variable(amount, display_name="Amount")


def _model() -> mo.Model:
    return mo.Model("m", default_grain="day", default_start="2026-01-01", default_periods=30)


def _fill(m: mo.Model, n: int) -> None:
    with m:
        m.t = mo.MultiVariable("Tx", excel_props={"orient": "records"})
        with m.t as t:
            for k in range(n):
                setattr(t, f"x{k}", Tx(day=k % 30 + 1, amount=float(k % 13),
                                       display_name=f"x{k}"))


def _table(n: int) -> mo.Model:
    m = _model()
    _fill(m, n)
    return m


@pytest.fixture
def copied(monkeypatch):
    """How many children the ``_components`` snapshot has copied, summed
    over every read."""
    count = [0]
    real = MultiVariableBase.__dict__["_components"]

    def read(self):
        snapshot = real.fget(self)
        count[0] += len(snapshot)
        return snapshot

    monkeypatch.setattr(MultiVariableBase, "_components", property(read, real.fset))
    return count


class TestScale:
    @pytest.mark.parametrize("n", [200, 400])
    def test_filling_a_table_copies_each_node_at_most_once(self, copied, n):
        m = _model()
        copied[0] = 0
        _fill(m, n)
        # A record is three nodes: itself and its two fields. A snapshot
        # per child adopted copied the table once per record: n² / 2.
        assert copied[0] <= 3 * n, copied[0]

    def test_writing_a_table_twice_as_long_copies_twice_as_much(self, copied, tmp_path):
        def write(n):
            m = _table(n)
            copied[0] = 0
            m.to_excel(tmp_path / f"t{n}.xlsx")
            return copied[0]

        small, large = write(200), write(400)
        # Linear: about 2. A snapshot per child read: about 4.
        assert large <= 2.5 * small, (small, large)

    def test_a_sheet_of_lines_twice_as_long_with_totals_copies_twice_as_much(
        self, copied, tmp_path
    ):
        def write(n):
            m = mo.Model("m", default_grain="day", default_start="2026-01-01",
                         default_periods=10,
                         default_excel_view=mo.ExcelView(timeline={"totals": ["month"]}))
            with m:
                m.s = mo.MultiVariable("S")
                with m.s as s:
                    for r in range(n):
                        setattr(s, f"r{r}", mo.Variable([float(r + k) for k in range(10)],
                                                        regrain=mo.up("sum")))
            copied[0] = 0
            m.to_excel(tmp_path / f"s{n}.xlsx")
            return copied[0]

        small, large = write(100), write(200)
        assert large <= 2.5 * small, (small, large)

    def test_a_walk_copies_each_child_at_most_once(self, copied):
        m = _table(300)
        copied[0] = 0
        walked = list(m.walk())
        assert len(walked) == 1 + 300 * 3  # the table, its records, their fields
        assert copied[0] <= len(walked)

    def test_the_table_is_written_whole(self, tmp_path):
        import openpyxl

        path = tmp_path / "t.xlsx"
        _table(50).to_excel(path)
        ws = openpyxl.load_workbook(path)["Tx"]
        rows = [[c.value for c in r][:3] for r in ws.iter_rows()]
        assert rows[0] == [None, "Day", "Amount"]
        assert rows[1:] == [[f"x{k}", k % 30 + 1, k % 13] for k in range(50)]


class TestTheSnapshotStaysASnapshot:
    def test_each_read_is_a_fresh_dict_the_container_does_not_share(self):
        mv = mo.MultiVariable(a=mo.Variable(1), b=mo.Variable(2))
        snapshot = mv._components
        assert snapshot is not mv._components
        snapshot["c"] = mo.Variable(3)
        del snapshot["a"]
        assert list(mv._components) == ["a", "b"]
        assert mv._has_component("a") and not mv._has_component("c")

    def test_a_snapshot_iterates_while_the_container_changes(self):
        mv = mo.MultiVariable(**{f"v{i}": mo.Variable(i) for i in range(20)})
        seen = []
        for name, var in mv._components.items():
            seen.append(name)
            mv.remove(name)
            setattr(mv, f"w{name}", mo.Variable(var._value))
        assert seen == [f"v{i}" for i in range(20)]
        assert mv._component_order == [f"wv{i}" for i in range(20)]

    def test_a_walk_reads_each_child_when_its_turn_comes(self):
        mv = mo.MultiVariable(a=mo.Variable(1), b=mo.Variable(2))
        seen = []
        for name, comp in mv.walk():
            seen.append((name, comp._value))
            if name == "a":
                with pytest.warns(ModelStructureWarning):
                    mv.b = mo.Variable(20)  # replaced before its turn
                mv.c = mo.Variable(3)  # joined after the walk began
        assert seen == [("a", 1), ("b", 20)]

    def test_the_helpers_answer_what_the_snapshot_answers(self):
        m = _table(40)
        m.t.note = 5  # a plain attribute, not a child
        for mv in (m, m.t, m.t.x0):
            snapshot = mv._components
            assert list(snapshot) == mv._component_order
            assert all(mv._component(name) is child for name, child in snapshot.items())
            for name in (*snapshot, "missing", "note", "default_grain", "_parent"):
                assert mv._has_component(name) == (name in snapshot)
        with pytest.raises(KeyError):
            m.t._component("note")


class TestChildrenComeAndGo:
    def test_a_replaced_child_moves_to_the_end_once(self):
        mv = mo.MultiVariable(**{f"v{i}": mo.Variable(i) for i in range(20)})
        with pytest.warns(ModelStructureWarning):
            mv.v3 = mo.Variable(33)
        assert mv._component_order == [f"v{i}" for i in range(20) if i != 3] + ["v3"]
        assert mv._component("v3")._value == 33

    def test_a_removed_child_is_gone_and_its_name_free_again(self):
        mv = mo.MultiVariable(**{f"v{i}": mo.Variable(i) for i in range(20)})
        del mv.v5
        gone = mv.remove("v6")
        assert gone._owner is None
        for name in ("v5", "v6"):
            assert not mv._has_component(name) and name not in mv._component_names
            with pytest.raises(KeyError):
                mv._component(name)
        mv.v5 = mo.Variable(55)
        assert mv._component_order.count("v5") == 1 and mv._component_order[-1] == "v5"

    def test_a_child_owned_elsewhere_joins_as_a_copy_an_orphan_as_itself(self):
        a, b, c = mo.MultiVariable("A"), mo.MultiVariable("B"), mo.MultiVariable("C")
        a.rec = Tx(day=1, amount=2.0)
        b.rec = a.rec
        assert b._component("rec") is not a._component("rec")
        assert b._component("rec")._parent is b
        orphan = a.remove("rec")
        assert not a._has_component("rec")
        c.rec = orphan
        assert c._component("rec") is orphan and orphan._parent is c

    def test_a_cloned_table_keeps_its_records(self):
        m = _table(30)
        other = _model()
        with other:
            other.t = m.t  # owned by m: adopted as a clone
        assert other.t is not m.t
        assert other.t._component_order == m.t._component_order
        assert all(other.t._component(n)._parent is other.t for n in other.t._component_order)
        assert other.t.x7.day.id == "m.t.x7.day"

    def test_the_bulk_setter_replaces_the_children(self):
        # Virtual sheets are filled this way: the old set out, the new in.
        mv = mo.MultiVariable(**{f"v{i}": mo.Variable(i) for i in range(20)})
        x, y = mo.Variable(1), mo.Variable(2)
        mv._components = {"x": x, "y": y}
        assert mv._component_order == ["x", "y"]
        assert mv._component("x") is x and mv._has_component("y")
        assert not mv._has_component("v0") and "v0" not in mv.__dict__
        assert [name for name, _ in mv.walk()] == ["x", "y"]

    def test_an_order_without_a_name_drops_it_from_the_children(self):
        mv = mo.MultiVariable(**{f"v{i}": mo.Variable(i) for i in range(20)})
        order = [f"v{i}" for i in reversed(range(1, 20))]
        mv._component_order = order
        assert "v0" not in mv._components and not mv._has_component("v0")
        with pytest.raises(KeyError):
            mv._component("v0")
        assert [name for name, _ in mv.walk()] == order


class TestTheNameList:
    @pytest.fixture
    def names_type(self):
        from modeleon.core.multi_variable import _ComponentNames

        return _ComponentNames

    def test_every_way_a_container_gets_its_names_keeps_the_lookup(self, names_type):
        m = _table(20)
        assert type(m.t._component_names) is names_type
        m.t._component_order = list(reversed(m.t._component_order))
        assert type(m.t._component_names) is names_type
        other = _model()
        with other:
            other.t = m.t
        assert type(other.t._component_names) is names_type
        m.t._components = {"x": mo.Variable(1)}
        assert type(m.t._component_names) is names_type

    @pytest.mark.parametrize("size", [3, 40])  # scanned, looked up
    def test_in_follows_every_change(self, names_type, size):
        names = names_type(f"n{i}" for i in range(size))
        same = list(names)

        def agree():
            for probe in (*same, "absent", "n999"):
                assert (probe in names) == (probe in same), probe
            assert names == same

        changes = [
            lambda s: s.append("a"),
            lambda s: s.extend(["b", "c"]),
            lambda s: s.insert(0, "d"),
            lambda s: s.remove("b"),
            lambda s: s.pop(),
            lambda s: s.pop(0),
            lambda s: s.__setitem__(0, "e"),
            lambda s: s.__setitem__(slice(0, 2), ["f", "g", "h"]),
            lambda s: s.__delitem__(1),
            lambda s: s.__delitem__(slice(0, 2)),
            lambda s: s.__iadd__(["i"]),
            lambda s: s.__imul__(2),
            lambda s: s.sort(),
            lambda s: s.reverse(),
            lambda s: s.clear(),
            lambda s: s.append("j"),
        ]
        agree()
        for change in changes:
            change(names)
            change(same)
            agree()

    def test_a_copy_is_a_list_of_its_own(self, names_type):
        names = names_type(f"n{i}" for i in range(40))
        assert "n0" in names and [1] not in names
        for other in (copy.copy(names), copy.deepcopy(names),
                      pickle.loads(pickle.dumps(names))):
            assert other == names and type(other) is names_type
            other.append("only_here")
            assert "only_here" in other and "only_here" not in names
