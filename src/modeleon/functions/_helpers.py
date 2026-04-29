# SPDX-License-Identifier: Apache-2.0
"""Shared helpers for Variable-aware DSL functions.

The ``mathfn`` / ``text`` / ``dates`` / ``conditional`` / ``aggregate`` modules
each need the same three operations: split a Variable-or-scalar argument
into its raw value plus an AST node, render an argument for the
``_source_code`` string, and assemble a result Variable with a
``=FUNC(args)`` formula. These helpers consolidate that boilerplate.
"""

from __future__ import annotations

import functools
from typing import Any, Callable

from ..core.expr import FuncCall, Literal, VarRef
from ..core.variable import Variable


def val(x: Any) -> Any:
    """Unwrap a Variable to its raw value, or pass through a non-Variable.

    Use in ``MultiVariableClass.compute()`` or any place Python built-ins
    (``range``, ``len``, int comparisons) need a raw value rather than a
    Variable wrapper::

        class MyModel(mo.MultiVariableClass):
            def compute(self, periods=5, rate=0.05):
                for i in range(val(periods)):   # range() needs int
                    ...
                self.growth = self.base * rate  # operators handle Variable fine
    """
    return x._value if isinstance(x, Variable) else x


def operand(value) -> tuple[Any, "Literal | VarRef"]:
    """Split a Variable-or-scalar into ``(raw value, AST node)``.

    Variables become ``VarRef``; scalars become ``Literal`` carrying the
    raw value unchanged. Suits math / date / aggregate inputs where the
    raw-value representation renders cleanly in a formula string.
    """
    if isinstance(value, Variable):
        return value._value, VarRef(value)
    return value, Literal(value)


def text_operand(value) -> tuple[Any, "Literal | VarRef"]:
    """Split a Variable-or-string into ``(raw value, AST node)``.

    The :class:`Literal` stores the raw string; ``Literal.to_string``
    Python-quotes it for ``.formula`` and the translator double-quotes
    it for Excel output — both rendering paths handle quoting, so we
    don't pre-wrap here.
    """
    if isinstance(value, Variable):
        return value._value, VarRef(value)
    return value, Literal(value)


def src(value) -> str:
    """Render a value for the human-readable ``_source_code`` string.

    Returns the Variable's id when it's a Variable, or ``repr(value)``
    for scalars — so list → ``[1, 2, 3]``, string → ``'hello'``.
    """
    return (value.python_name or value.path.leaf) if isinstance(value, Variable) else repr(value)


def make_func_var(
    func_name: str,
    ast_args: list,
    value: Any,
    value_type: str,
    var_type: str = 'scalar',
    source_code: str | None = None,
    render_backends: frozenset | None = None,
    compute_backends: frozenset | None = None,
) -> Variable:
    """Build a Variable wrapping a ``=FUNC(args)`` formula + computed value.

    Args:
        func_name: Excel function name (``"SUM"``, ``"ABS"``, …).
        ast_args: Expr nodes for each argument (built via ``operand`` /
            ``text_operand``). Become the AST under the Variable.
        value: The Python-side computed value (scalar or list).
        value_type: Modeleon value type (``"float"``, ``"int"``, ``"string"``,
            ``"bool"``, ``"datetime"``, …).
        var_type: ``"scalar"`` (default) or ``"list"``.
        source_code: Optional override for ``_source_code``. Helpers that
            want a non-default rendering set this explicitly.
        render_backends: Renderers that emit this function natively
            (``None`` = universal; every renderer handles it). Forwarded
            to :class:`FuncCall`; non-matching renderers fall back to
            inlining ``_value``.
        compute_backends: Compute environments that evaluate this
            function natively at runtime (``None`` = universal). Rarely
            populated today; distinguishes render target from compute
            target when a backend can emit a call but not evaluate it.
    """
    result = Variable(formula=FuncCall(
        func_name, ast_args,
        render_backends=render_backends,
        compute_backends=compute_backends,
    ))
    result._value = value
    result.var_type = var_type
    result.value_type = value_type
    if source_code is not None:
        result._source_code = source_code
    return result


def pyformula(
    fn: Callable[..., Any] | None = None,
    *,
    name: str | None = None,
    value_type: str = 'float',
) -> Callable[..., Any]:
    """Turn an arbitrary Python function into a Variable-returning helper
    with tracked dependencies.

    Wraps a user function so that calling it with Variable arguments
    returns a :class:`Variable` whose AST is a :class:`FuncCall` — the
    graph sees the dep edges, :func:`modeleon.to_json` serializes the
    call, :func:`modeleon.check_compat` can inspect it. Every renderer
    falls back to inlining the computed ``_value`` (``render_backends``
    is the empty set — "native to nothing"), so the Excel cell is a
    literal, not a broken ``=NAME?`` formula.

    Use when the function body is opaque to the DSL — calls into
    ``scipy``, a trained model, an HTTP API, anything that can't be
    expressed as arithmetic on Variables. For ordinary arithmetic,
    just write the expression — operators build the AST automatically.

    Usage (also importable from :mod:`modeleon.extend`, which is the
    stable surface for extension authors)::

        from modeleon.extend import pyformula

        @pyformula
        def minimize(x):
            return scipy.optimize.minimize(x).x

        a = minimize(some_variable)   # Variable with deps on some_variable

        @pyformula(name="SCIPY_MIN", value_type='float')
        def minimize(x):
            ...

    Positional arguments only — :class:`FuncCall` has no kwargs slot.
    ``value_type`` defaults to ``'float'``; override for int / string /
    datetime results. ``var_type`` is inferred (``'list'`` if the
    function returns a list, otherwise ``'scalar'``). ``name`` defaults
    to the wrapped function's ``__name__`` uppercased.
    """
    def deco(inner: Callable[..., Any]) -> Callable[..., Any]:
        func_name = name or inner.__name__.upper()

        @functools.wraps(inner)
        def wrapper(*args: Any) -> Variable:
            raw_args = [a._value if isinstance(a, Variable) else a for a in args]
            ast_args = [
                VarRef(a) if isinstance(a, Variable) else Literal(a) for a in args
            ]
            result_value = inner(*raw_args)
            var_type = 'list' if isinstance(result_value, list) else 'scalar'
            return make_func_var(
                func_name, ast_args, result_value, value_type, var_type,
                render_backends=frozenset(),
            )

        return wrapper

    if fn is not None and callable(fn):
        return deco(fn)
    return deco
