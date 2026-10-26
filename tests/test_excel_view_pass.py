# SPDX-License-Identifier: Apache-2.0
"""``view_pass`` serves exactly what a resolve outside it would.

A tree with a view at three levels is resolved, node by node, through a
sequence of the changes a pass must see — a house view entered and left,
the process-wide default replaced, a view edited — once inside one pass
and once with no pass at all. Every answer must be the same, and two
threads, each in its own pass over its own tree, must not see each
other's answers.
"""

from __future__ import annotations

import threading
from contextlib import ExitStack

import pytest

import modeleon as mo
from modeleon.compile.excel.view import ExcelView, resolve_excel_view, view_pass


@pytest.fixture(autouse=True)
def _restore_global():
    saved = mo.default_excel_view
    yield
    mo.default_excel_view = saved


def _tree(font: str = "Arial Narrow"):
    m = mo.Model("m")
    m.default_excel_view = ExcelView(
        font={"name": font}, formats={"number": "#,##0"}, bands={"section": {"bold": True}})
    m.a = mo.MultiVariable("A")
    m.a.default_excel_view = ExcelView(formats={"%": "0.0%"}, nesting={"indent": 2})
    m.a.b = mo.MultiVariable("B")
    m.a.b.default_excel_view = ExcelView(bands={"input": {"font_color": "0000FF"}})
    m.a.b.x = mo.Variable(1)
    m.a.y = mo.Variable(2)
    m.z = mo.Variable(3)
    return m


def _nodes(m):
    return [m.a.b.x, m.a.y, m.z, m.a.b, m.a, m, None]


def _script(m, stack: ExitStack, views: dict) -> list:
    """Resolve every node (twice: the second ask is the memo's), change
    something, again — the answers after each step. The views it installs
    are built before any pass: building a view drops an open pass's memo,
    which would hide whether the pass sees the change by itself."""
    out = []

    def record():
        first = [resolve_excel_view(n) for n in _nodes(m)]
        again = [resolve_excel_view(n) for n in reversed(_nodes(m))][::-1]
        assert again == first
        out.append(first)

    floor = mo.default_excel_view
    record()
    stack.enter_context(mo.house_view(views["house"]))
    record()
    mo.default_excel_view = views["floor"]
    record()
    m.a.default_excel_view.orient = views["orient"]     # an edit of a view
    record()
    stack.close()                                        # the house view ends
    record()
    mo.default_excel_view = floor
    record()
    return out


def _views() -> dict:
    return {
        "house": ExcelView(timeline={"label_format": "iso"}),
        "floor": ExcelView(timeline={"label_format": "compact"}, formats={"number": "0.00"}),
        "orient": mo.Variable("down"),
    }


def test_a_pass_answers_as_no_pass_through_every_change():
    with ExitStack() as stack:
        plain = _script(_tree(), stack, _views())
    tree, views = _tree(), _views()
    with view_pass(), ExitStack() as stack:
        passed = _script(tree, stack, views)
    assert passed == plain
    # Each step did change what the nodes resolve to.
    assert all(a != b for a, b in zip(plain, plain[1:]))


def test_the_pass_ends_with_the_block():
    m = _tree()
    deeper = ExcelView(nesting={"indent": 9})       # built outside the pass
    with view_pass():
        assert resolve_excel_view(m.a.y).nesting_indent == 2
    m.a.default_excel_view = deeper                  # builds no view: drops no memo
    assert resolve_excel_view(m.a.y).nesting_indent == 9


def test_two_threads_in_their_own_passes():
    trees = {"Arial": _tree("Arial"), "Calibri": _tree("Calibri")}
    expected = {k: [resolve_excel_view(n) for n in _nodes(t)] for k, t in trees.items()}
    got: dict = {}
    barrier = threading.Barrier(2, timeout=30)

    def worker(key):
        with view_pass():
            first = [resolve_excel_view(n) for n in _nodes(trees[key])]
            barrier.wait()
            got[key] = (first, [resolve_excel_view(n) for n in _nodes(trees[key])])

    threads = [threading.Thread(target=worker, args=(k,)) for k in trees]
    for t in threads:
        t.start()
    for t in threads:
        t.join()
    for key in trees:
        assert got[key] == (expected[key], expected[key])
    assert expected["Arial"][0].base_font != expected["Calibri"][0].base_font



def _view_model():
    import modeleon as mo

    m = mo.MultiVariable("m")
    m.sec = mo.MultiVariable("sec")
    m.sec.x = mo.Variable(1)
    m.default_excel_view = mo.ExcelView()
    return m


def test_an_edit_inside_a_nested_pass_is_seen_by_the_outer_one():
    """The outer pass resolved a line; an inner pass edited the view. Once
    the inner block ends, the outer pass answers with the edit."""
    import modeleon as mo
    from modeleon.compile.excel.view import resolve_excel_view, view_pass

    m = _view_model()
    before = resolve_excel_view(m.sec.x).orient
    with view_pass():
        assert resolve_excel_view(m.sec.x).orient == before
        with view_pass():
            m.default_excel_view.orient = mo.Variable("down" if before != "down" else "across")
        served = resolve_excel_view(m.sec.x).orient
    assert served != before
    assert served == resolve_excel_view(m.sec.x).orient


def test_an_edit_from_another_thread_is_seen_inside_a_pass():
    """A view edited by another thread while a pass is open: the pass
    answers with the edit, as a resolve without a pass does."""
    import threading

    import modeleon as mo
    from modeleon.compile.excel.view import resolve_excel_view, view_pass

    m = _view_model()
    before = resolve_excel_view(m.sec.x).orient
    with view_pass():
        assert resolve_excel_view(m.sec.x).orient == before
        target = "down" if before != "down" else "across"
        t = threading.Thread(target=lambda: setattr(m.default_excel_view, "orient", mo.Variable(target)))
        t.start()
        t.join()
        served = resolve_excel_view(m.sec.x).orient
    assert served == target
