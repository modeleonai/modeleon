"""The compute-signature memo must not keep an exec'd model alive.

A long-lived process that execs model sources builds new classes on
every run. Memoizing ``inspect.signature`` with a strong key (as
``functools.lru_cache`` does) pinned each class's ``compute`` function —
and through its ``__globals__`` the whole namespace of the model that
built it. A weak key lets the entry die with the class.
"""

from __future__ import annotations

import gc
import weakref

SRC = """
import modeleon as mo

class Rec(mo.MultiVariableClass):
    def compute(self, amount=0.0):
        self.amount = mo.Variable(amount)

m = mo.Model('m')
with m:
    m.r = Rec(amount=1.0)
"""


def _build() -> weakref.ref:
    ns: dict = {}
    exec(compile(SRC, "<model>", "exec"), ns)
    return weakref.ref(ns["Rec"])


def test_an_execd_model_class_is_released_after_use() -> None:
    refs = [_build() for _ in range(3)]
    gc.collect()
    assert all(r() is None for r in refs), "the signature memo pins exec'd model classes"


def test_the_signature_is_still_memoized_while_the_class_lives() -> None:
    from modeleon.core.multi_variable import _signature_of

    ns: dict = {}
    exec(compile(SRC, "<model>", "exec"), ns)
    fn = ns["Rec"].compute
    assert _signature_of(fn) is _signature_of(fn)
    assert list(_signature_of(fn).parameters) == ["self", "amount"]
