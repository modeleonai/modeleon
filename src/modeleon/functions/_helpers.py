# SPDX-License-Identifier: Apache-2.0
"""Shared helpers for Variable-aware DSL functions.

The ``mathfn`` / ``text`` / ``dates`` / ``conditional`` / ``aggregate`` modules
each need the same three operations: split a Variable-or-scalar argument
into its raw value plus an AST node, render an argument for the
``_source_code`` string, and assemble a result Variable with a
``=FUNC(args)`` formula. These helpers consolidate that boilerplate.

:func:`elementwise` is the one law for a function that is one call per
period — the cap and the floor ``MIN`` / ``MAX`` write in every column,
and the element-wise helpers built on it: scalars broadcast as ``a + b`` does,
a one-value series broadcasts, other length mismatches are refused,
tracked operands lift per coordinate, an error marker is the period's
result, the unit is the function's law (:func:`function_unit`).
"""

from __future__ import annotations

import functools
import inspect
import typing
from typing import Any, Callable, Sequence

from ..core.expr import FuncCall, Literal, VarRef
from ..core.shape import keys_along
from ..core.tracks import TrackValues
from ..core.unit import UNSETTLED
from ..core.variable import Variable
from ..core.variable_ops import _line_name, _on_dates, _VariableArithmetic


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


#: Functions whose result is NOT in the operands' unit — counts, rates,
#: calendar spans, positions, an exponential, a logarithm, a sign, a
#: boolean, text, a date.
_UNITLESS_RESULTS = frozenset({
    'COUNT', 'COUNTA', 'IRR', 'XIRR', 'YEARFRAC', 'MATCH',
    'EXP', 'LN', 'LOG', 'LOG10', 'SIGN',
    'AND', 'OR', 'NOT', 'XOR', 'ISBLANK', 'ISERROR', 'ISNUMBER',
    'DATE', 'EDATE', 'EOMONTH', 'TODAY', 'WORKDAY', 'DAYS', 'DAYS360', 'DATEDIF',
    'DAY', 'MONTH', 'YEAR', 'WEEKDAY', 'WEEKNUM', 'NETWORKDAYS',
    'LEN', 'UPPER', 'LOWER', 'CONCAT', 'LEFT', 'RIGHT', 'MID', 'TRIM', 'FIND',
    'SEARCH', 'SUBSTITUTE', 'VALUE', 'TEXT', 'REPT', 'TEXTJOIN',
})


#: Functions that add, compare or pick their operands, so the operands
#: must share a unit the way ``+`` needs it: the position of the first
#: operand that is a value (an IF's condition, a CHOOSE's index are not).
#: A remainder and a multiple share one unit with what they divide by
#: (``MOD(months, 12)`` is in months, ``MOD(USD, hr)`` is refused).
#: IFS and SWITCH pick every other operand: their laws are declared
#: beside them.
_SAME_UNIT_FROM = {'SUM': 0, 'AVERAGE': 0, 'MIN': 0, 'MAX': 0, 'IF': 1, 'CHOOSE': 1,
                   'IFERROR': 0, 'MOD': 0, 'CEILING': 0, 'FLOOR': 0}

#: A unit law: the operand nodes of a call and the reader of a node's
#: unit → the result's unit.
UnitLaw = Callable[[Sequence[Any], Callable[[Any], Any]], Any]

#: The functions with a law of their own — a power, a product, a
#: rounding, a conditional total, a pick from every other operand, the
#: financial functions — each declared next to the function
#: (``unit_law('POWER', ...)``).
_UNIT_LAWS: dict[str, UnitLaw] = {}


def unit_law(name: str, law: UnitLaw) -> None:
    """Declare the unit law of ``name`` (see :data:`_UNIT_LAWS`)."""
    _UNIT_LAWS[name] = law


def function_unit(func_name: str, args: Sequence[Any], unit_of: Callable[[Any], Any],
                  *, own_law: bool = True) -> Any:
    """The unit ``func_name`` gives over its operand nodes, each one's unit
    read by ``unit_of``; an operand the function cannot take is refused.
    The one law a function's result is built by and a loop container
    derives its rows' units again by: the function's declared law; else
    none for a count, a boolean, text or a date; else the unit the values
    share where the function adds, compares or picks them; else the unit
    every unit-bearing operand shares (``ABS`` of $ is $). A call the
    workbook does not write as a spreadsheet function (``own_law``
    False: a ``pyformula``, whatever its name) is opaque: no unit, unless
    its author declares one (``pyformula(unit=...)``). While a loop settles, ``unit_of`` may read a row as
    :data:`UNSETTLED`, and so may the answer."""
    if not own_law:
        return None
    law = _UNIT_LAWS.get(func_name)
    if law is not None:
        return law(args, unit_of)
    if func_name in _UNITLESS_RESULTS:
        return None
    start = _SAME_UNIT_FROM.get(func_name)
    if start is not None:
        return same_unit(func_name, args[start:], unit_of)
    return _shared_arg_unit(args, unit_of)


def same_unit(func_name: str, nodes: Sequence[Any], unit_of: Callable[[Any], Any]) -> Any:
    """The unit ``nodes`` share where ``func_name`` adds, compares or picks
    them: two different units are refused, as ``+`` refuses them; a
    number or an unlabelled line goes with any. Pure numbers under two
    labels add, and the result names neither. A row its loop has not
    settled yet goes with any too; with no other unit, the answer is
    :data:`UNSETTLED`."""
    from ..core.unit import unit_key
    from ..core.variable_ops import _line_name
    first: Any = None
    keys = set()
    waiting = False
    for node in nodes:
        unit = unit_of(node)
        if unit is UNSETTLED:
            waiting = True
            continue
        if unit is None:
            continue
        if first is None:
            first = (node, unit)
        elif unit != first[1]:
            names = None
            if isinstance(first[0], VarRef) and isinstance(node, VarRef):
                names = (_line_name(first[0].var), _line_name(node.var))
                paths = (str(first[0].var.path), str(node.var.path))
                if names[0] == names[1] and not any(p.startswith('__floating__') for p in paths):
                    names = (repr(paths[0]), repr(paths[1]))   # two records' 'Amount'
            raise mixed_units_error(func_name, first[1], unit, names)
        keys.add(unit_key(unit))
    refuse_counted_shares(func_name, nodes)
    if first is None:
        return UNSETTLED if waiting else None
    return first[1] if len(keys) == 1 else None


def refuse_counted_shares(func_name: str, nodes: Sequence[Any]) -> None:
    """The check ``+`` makes of a bps or pp line holding a count (50)
    beside a % line holding a fraction (0.05), made wherever ``func_name``
    adds, compares or picks the lines (:func:`same_unit`'s callers: a
    total, a pick, a lag's fill, a conditional total): the two cannot be
    read as one. A line its loop has not closed holds placeholders, so the
    loop makes the check when it closes (``core.loops._check_counts``)."""
    from ..core.variable_ops import _on_open_loop, _refuse_counts_among_shares
    lines = [n.var for n in nodes if isinstance(n, VarRef)]
    if len(lines) < 2 or any(_on_open_loop(v) for v in lines):
        return
    _refuse_counts_among_shares(lines, where=f"mo.{func_name}")


def waits_on_loop(var: Any) -> bool:
    """A Variable built on a loop row read before its line: its unit is
    provisional, and the loop derives and checks it when it closes."""
    from ..core.variable_ops import _on_open_loop
    return _on_open_loop(var)


def operand_unit(node: Any) -> Any:
    """The unit of a built call's operand: a line's own; a constant, text
    and a list of lines (a range, whose law reads its lines) have none; a
    call inside it (an operator joined to a cell) gives its own law's."""
    if isinstance(node, VarRef):
        return getattr(node.var, '_unit', None)
    if isinstance(node, FuncCall):
        return function_unit(node.func, node.args, operand_unit)
    return None


def mixed_units_error(func_name: str, first: Any, second: Any,
                      names: Sequence[str] | None = None) -> ValueError:
    """The refusal of two units in one function, worded as ``+`` words
    it (which side is a dimension, which a labelled number, the hint)."""
    from ..core.variable_ops import _incompatible_units
    detail = _incompatible_units(first, second).split(": ", 1)[1]
    who = f"{names[0]} and {names[1]}" if names else "they"
    return ValueError(f"mo.{func_name} takes its values in one unit; {who} are in {detail}")


def refuse_mixed_units(func_name: str, operands: Sequence[Any]) -> None:
    """Two operands in different units where ``func_name`` adds, compares
    or picks them: refused, as ``+`` refuses them. An operand without a
    unit (a number, an unlabelled line) goes with any."""
    from ..core.variable_ops import _line_name
    first = None
    for v in operands:
        u = getattr(v, '_unit', None) if isinstance(v, Variable) else None
        if u is None:
            continue
        if first is None:
            first = (v, u)
        elif u != first[1]:
            raise mixed_units_error(func_name, first[1], u,
                                    (_line_name(first[0]), _line_name(v)))


def _shared_arg_unit(ast_args: Sequence[Any], unit_of: Callable[[Any], Any]) -> Any:
    """The generic law: the single unit shared by every unit-bearing
    Variable operand (top-level VarRef args), else None — never guess;
    unsettled while one of them is."""
    from ..core.unit import unit_key

    units = []
    seen = set()
    for node in ast_args:
        if not isinstance(node, VarRef):
            continue
        u = unit_of(node)
        if u is UNSETTLED:
            return UNSETTLED
        if u is not None and unit_key(u) not in seen:
            seen.add(unit_key(u))
            units.append(u)
    return units[0] if len(units) == 1 else None


def make_func_var(
    func_name: str,
    ast_args: list,
    value: Any,
    value_type: str,
    var_type: str = 'scalar',
    source_code: str | None = None,
    render_backends: frozenset | None = None,
    compute_backends: frozenset | None = None,
    any_value_type: bool = False,
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
        any_value_type: Carry the unit whatever ``value_type`` says — a
            row typed by a first cell that holds an error marker is still
            a row of numbers. Otherwise only a number carries one.

    The unit is the function's law (:func:`function_unit`) — by the name
    the workbook writes; a call native to no renderer (``render_backends``
    empty, as :func:`pyformula` builds) has none unless declared.
    """
    result = Variable(formula=FuncCall(
        func_name, ast_args,
        render_backends=render_backends,
        compute_backends=compute_backends,
    ))
    result._value = value
    result.var_type = var_type
    result.value_type = value_type
    # The function's unit law (:func:`function_unit`) — what a loop
    # container derives the row's unit by again — refuses what the
    # function cannot take; over a loop row not settled yet the units are
    # provisional, and the loop checks them when it closes.
    try:
        unit = function_unit(func_name, ast_args, operand_unit,
                             own_law=render_backends != frozenset())
    except ValueError:
        if not waits_on_loop(result):
            raise
        unit = None
    if unit is not None and (any_value_type or value_type in ('int', 'float')):
        result._unit = unit
    if source_code is not None:
        result._source_code = source_code
    return result


# ─── One call per period ────────────────────────────────────────────


def is_error(value: Any) -> bool:
    """An Excel error token (``'#DIV/0!'``, ``'#N/A'``) as arithmetic
    carries it."""
    return isinstance(value, str) and value.startswith('#')


def is_series(value: Any) -> bool:
    """A period-shaped operand: a list, or a Variable holding one (a row,
    or a row per track)."""
    if isinstance(value, Variable):
        return isinstance(value._value, (list, TrackValues))
    return isinstance(value, list)


def align_keyed(func_name: str, lines: Sequence[Any],
                values: list[Any]) -> tuple[list[Any], list[Any] | None]:
    """The values of a call that is one per period, its lines along keys
    combined key by key as ``a + b`` combines them: a line along the same
    keys in another order is read in the first one's order; other keys,
    and a line along keys against a period series (Bear would meet
    January), are refused as ``a + b`` refuses them. ``lines[i]`` is the
    operand ``values[i]`` came from (a value typed in has none). Returns
    the values and the keys the result runs along — the first keyed
    line's, or None when no line runs along keys."""
    if not any(isinstance(v, (list, TrackValues)) for v in values):
        return values, None            # scalars run along no keys
    operands = [a for a in lines if isinstance(a, Variable)]
    # Only a pair with a side along keys can be refused: a call over many
    # lines none of which runs along keys is read once, not pair by pair.
    along = [keys_along(a) is not None for a in operands]
    if any(along):
        for i, line in enumerate(operands):
            for j in range(i + 1, len(operands)):
                if along[i] or along[j]:
                    _VariableArithmetic._reject_keys_against_periods(line, operands[j])
    first: Any = None
    out = list(values)
    for i, (line, value) in enumerate(zip(lines, values)):
        keys = getattr(line, '_keys', None) if isinstance(line, Variable) else None
        if not keys or not isinstance(value, list):
            continue
        if first is None:
            first = line
            continue
        if list(keys) == list(first._keys):
            continue
        if set(keys) != set(first._keys) or len(keys) != len(value):
            raise ValueError(
                f"mo.{func_name} combines {_line_name(first)} and {_line_name(line)} key "
                f"by key, and they run along other keys "
                f"({', '.join(map(str, first._keys))} and {', '.join(map(str, keys))}): "
                f"give both the same keys."
            )
        at = {key: j for j, key in enumerate(keys)}
        out[i] = [value[at[key]] for key in first._keys]
    return out, (None if first is None else list(first._keys))


def series_node(node: Any) -> bool:
    """An operand node of a built call holding a series: a line holding a
    list (or a row per track), or a list typed in."""
    if isinstance(node, VarRef):
        return is_series(node.var)
    return isinstance(node, Literal) and isinstance(node.value, list)


def shared_unit(args: Sequence[Any]) -> Any:
    """The single unit shared by every unit-bearing Variable operand — an
    aggregation of $ is $. Mixed or absent units → None (never guess)."""
    from ..core.unit import unit_key
    first = None
    keys = set()
    for a in args:
        unit = getattr(a, '_unit', None) if isinstance(a, Variable) else None
        if unit is None:
            continue
        keys.add(unit_key(unit))
        if first is None:
            first = unit
    return first if len(keys) == 1 else None


def first_variable_value_type(args: Sequence[Any], fallback: str = 'float') -> str:
    """The value type of the first Variable operand, else ``fallback``."""
    for a in args:
        if isinstance(a, Variable):
            return a.value_type
    return fallback


def align_located(lines: Sequence[Any], values: list[Any]) -> tuple[list[Any], Any]:
    """The values of a call that is one per period, its series on other
    windows aligned by date as ``a + b`` aligns them (one reading for
    both): each materialized onto the combined window by its own
    ``extend=`` rule, one that does not cover it refused, a first period
    that is short only at the window's start, two grains refused; an
    unnamed formula runs on the window of the lines it reads, a plain
    line beside a tracked one on the tracked line's window, a line along
    keys and a one-value constant on none. ``lines[i]`` is the operand
    ``values[i]`` came from (a value typed in has none). Returns the
    values and the combined window, a ``TimeLoc`` — or the values as given
    and None when fewer than two series run on windows, or all on one
    (the positional zip is then by date)."""
    if not any(isinstance(v, (list, TrackValues)) for v in values):
        return values, None            # scalars run on no window
    return _on_dates(lines, values)


def locate(var: Variable, window: Any) -> None:
    """Put ``var`` on ``window`` (a ``TimeLoc``) explicitly, as an operator
    puts its result on the combined window: downstream alignment then
    knows where it starts, and a renderer that its operands were read by
    date there."""
    var._start, var._grain, var._first_day = window
    var._read_by_date = True


def _typed_blank(raw: Any, fill: Any) -> Any:
    """A ``None`` typed into a call read as ``fill``, in a list typed in
    too — the formula holds the 0 the renderer writes for it."""
    if raw is None:
        return fill
    if isinstance(raw, list):
        return [fill if v is None else v for v in raw]
    return raw


def refuse_other_lengths(values: Sequence[Any]) -> None:
    """Series of other lengths among one call's values, refused as the
    operators refuse them; a one-value series broadcasts. A function with
    a loop of its own asks it once it has read lines on other windows by
    date (:func:`align_located`), where a series with no window of its
    own cannot be paired with the combined window by position."""
    lengths = [len(v) for v in values if isinstance(v, list)]
    n = max(lengths, default=0)
    for k in lengths:
        if k not in (1, n):
            raise ValueError(_other_lengths(k, n))


def _other_lengths(k: int, n: int) -> str:
    return (f"Cannot broadcast lists of different lengths: {k} vs {n}. "
            f"Element-wise ops need equal lengths, or one side of length 1 "
            f"which broadcasts over the other.")


def refuse_unpaired(lines: Sequence[Any], values: Sequence[Any], window: Any) -> None:
    """For a function with a loop of its own, once :func:`align_located`
    has read its ``lines``: the series it cannot pair period by period,
    refused as the operators refuse them. Over a combined ``window`` every
    series must fill it. Otherwise a line whose series is of another
    length than the call's longest is refused — its cells hold nothing
    for the periods it lacks, so the workbook could not read what a
    positional pairing would — while a list typed in keeps its positional
    reading, its last value standing in, as the cell writes it. A line
    refused is named; a one-value series broadcasts."""
    longest = max((len(v) for v in values if isinstance(v, list)), default=None)
    if longest is None:
        return
    for line, v in zip(lines, values):
        if not isinstance(v, list) or len(v) in (1, longest):
            continue
        if isinstance(line, Variable):
            beside = ("" if window is None else
                      f"; the other lines meet on {longest} period(s) from {window.start}")
            raise ValueError(
                f"{_line_name(line)} cannot be paired with the call's other series "
                f"period by period (different lengths: {len(v)} vs {longest}): it runs "
                f"on no window of its own to read by date{beside}. Give it the "
                f"periods of the lines beside it, or put it on the window it runs on "
                f"(a line of a container, or mo.Variable(..., start=..., grain=...))."
            )
        if window is not None:
            raise ValueError(_other_lengths(len(v), longest))


def _per_period(values: list[Any], call: Callable[[list[Any]], Any]) -> Any:
    """``call`` over the operands of one period at a time; a scalar or a
    one-value series stands in every period, as the operators broadcast
    it. Scalars only: one call."""
    lists = [v for v in values if isinstance(v, list)]
    if not lists:
        return call(values)
    n = max(len(v) for v in lists)
    refuse_other_lengths(lists)

    def at(v: Any, i: int) -> Any:
        if isinstance(v, list):
            return v[i] if len(v) == n else v[0]
        return v

    return [call([at(v, i) for v in values]) for i in range(n)]


def _check_arity(func_name: str, pyfunc: Callable[..., Any], n: int) -> None:
    """Refuse a ``pyfunc`` that cannot take the call's ``n`` operands at
    entry — it would otherwise fail in every period and read as
    ``#VALUE!`` throughout. A builtin without a signature (``max``) is
    trusted."""
    try:
        signature = inspect.signature(pyfunc)
    except (TypeError, ValueError):
        return
    try:
        signature.bind(*([None] * n))
    except TypeError:
        raise TypeError(
            f"{func_name}: the function given takes {signature}, but the call "
            f"has {n} operand{'s' if n != 1 else ''}."
        ) from None


def elementwise(
    func_name: str,
    pyfunc: Callable[..., Any],
    args: Sequence[Any],
    *,
    value_type: str | None = None,
    blank: Any = None,
    propagate_errors: bool = True,
    exceptions: typing.Literal['marker', 'raise'] = 'marker',
    source_code: str | None = None,
    render_backends: frozenset | None = None,
    typed_blank: Any = None,
) -> Variable:
    """Build the Variable for ``func_name`` as one call per period.

    ``pyfunc`` takes one period's operand values positionally and gives
    that period's value. Over scalars only it is called once and the
    result is a scalar; a list operand gives a series, scalars and
    one-value series broadcasting into every period and other length
    mismatches refused, as ``a + b`` does; located series on other
    windows are aligned by date as ``a + b`` aligns them
    (:func:`align_located`), the result on the combined window; lines
    along keys combine key by key (:func:`align_keyed`), the result
    along the same keys; a tracked operand lifts per coordinate. The
    AST is a :class:`FuncCall` over the operands.

    An error marker among a period's operands is that period's result —
    the first in reading order — unless ``propagate_errors`` is False,
    when the markers reach ``pyfunc`` (``IFERROR`` reads them). A blank
    operand (``None``) reaches ``pyfunc`` as is, or as ``blank`` when one
    is given (``0`` for arithmetic, ``''`` for text); a ``None`` typed
    into the call itself (or into a list typed in) reads as
    ``typed_blank`` when one is given — the 0 the renderer writes for it
    — while a line's blank cell stays a blank. A call that fails
    is the spreadsheet's error: ``ZeroDivisionError`` → ``#DIV/0!``,
    ``TypeError`` → ``#VALUE!``, ``OverflowError`` → ``#NUM!`` — unless
    ``exceptions='raise'``, when it propagates (``MIN`` / ``MAX`` refuse
    a text operand at build time). A ``pyfunc`` that cannot take the
    call's operands is refused at entry, never read as ``#VALUE!`` in
    every period.

    ``value_type`` defaults to the first Variable operand's, else
    ``'float'``. The unit is the function's law (:func:`function_unit`):
    the one every unit-bearing operand shares, unless the function
    declares another.
    """
    from ..core.shape import reject_axised
    _check_arity(func_name, pyfunc, len(args))
    raws: list[Any] = []
    ast_args: list = []
    for a in args:
        if isinstance(a, Variable):
            reject_axised(a, func_name)
        raw, node = operand(a)
        if typed_blank is not None and not isinstance(a, Variable):
            raw = _typed_blank(raw, typed_blank)
        raws.append(raw)
        ast_args.append(node)
    raws, keys = align_keyed(func_name, args, raws)
    raws, window = align_located(args, raws)

    def call(values: list[Any]) -> Any:
        if propagate_errors:
            for v in values:
                if is_error(v):
                    return v
        if blank is not None:
            values = [blank if v is None else v for v in values]
        if exceptions == 'raise':
            return pyfunc(*values)
        try:
            return pyfunc(*values)
        except ZeroDivisionError:
            return '#DIV/0!'
        except TypeError:
            return '#VALUE!'
        except OverflowError:
            return '#NUM!'

    if any(isinstance(r, TrackValues) for r in raws):
        value: Any = TrackValues.lift(lambda *fibers: _per_period(list(fibers), call), *raws)
        var_type = 'list' if value.time_length else 'scalar'
    else:
        value = _per_period(raws, call)
        var_type = 'list' if isinstance(value, list) else 'scalar'
    if value_type is None:
        value_type = first_variable_value_type(args, 'float')
    if source_code is None:
        source_code = f"{func_name}({', '.join(src(a) for a in args)})"
    out = make_func_var(
        func_name, ast_args, value, value_type, var_type,
        source_code=source_code, render_backends=render_backends, any_value_type=True,
    )
    return placed(out, window, keys)


def placed(out: Variable, window: Any, keys: list[Any] | None) -> Variable:
    """``out`` on the combined ``window`` its operands were aligned onto
    (:func:`align_located`), along the ``keys`` they ran along
    (:func:`align_keyed`), as an operator places its result; either may
    be None."""
    if window is not None:
        locate(out, window)
    if keys is not None and isinstance(out._value, list) and len(out._value) == len(keys):
        out._keys = keys
    return out


def pyformula(
    fn: Callable[..., Any] | None = None,
    *,
    name: str | None = None,
    value_type: str = 'float',
    unit: Any = None,
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

    The engine cannot see inside the function, so the result has no
    unit: a ratio of two USD lines is a pure number, a smoothed USD
    series is still USD, and only the author knows which. Declare it
    with ``unit=`` (``@pyformula(unit='USD')``), as ``set_unit`` would.
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
            out = make_func_var(
                func_name, ast_args, result_value, value_type, var_type,
                render_backends=frozenset(),
            )
            return out.set_unit(unit) if unit is not None else out

        return wrapper

    if fn is not None and callable(fn):
        return deco(fn)
    return deco
