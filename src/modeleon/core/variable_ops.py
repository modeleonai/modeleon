# SPDX-License-Identifier: Apache-2.0
"""Arithmetic and comparison mixin for :class:`Variable`.

Split out from ``variable.py`` to keep the main file focused on
identity, construction, and representation. This module owns:

- Broadcasting (scalar/list, list/list with length-1 fallback)
- Safe-op semantics (Excel-style ``#DIV/0!`` and error propagation)
- Key alignment and propagation for keyed (dict-sourced) Variables
- Unit propagation through arithmetic
- The shared ``_binary_op`` / ``_comparison_op`` scaffolds
- All binary, unary, and comparison dunders

``_VariableArithmetic`` is a **mixin** — it assumes ``self`` is a
:class:`Variable` and relies on ``self._value``, ``self._keys``,
``self._as_operand``, plus the ``expr`` property setter (which updates
``_expr`` and ``_dependency_refs`` together). Variable inherits from
this class so the operator surface is available everywhere Variable
is, while the file on disk stays coherent.

Lazy imports for :class:`Variable` inside each method avoid the
circular import that would arise if we imported at module scope
(``variable.py`` imports this module; this module needs Variable to
construct results).
"""

from __future__ import annotations

from datetime import date, datetime, time, timedelta
from typing import Any, List, Optional, Tuple, TYPE_CHECKING

from .expr import BinOp, Compare, Expr, Literal, UnaryOp, VarRef

if TYPE_CHECKING:
    from .unit import Unit
    from .variable import Variable


# Excel's day 0: a date's serial number counts the days since then.
_EXCEL_EPOCH_ORDINAL = date(1899, 12, 30).toordinal()
_ONE_DAY = timedelta(days=1)
_COMPARISON_SYMBOLS = frozenset({'==', '!=', '<', '<=', '>', '>='})


#: A blank compares as "" against text, as Excel's ``A1=""`` does.
_COMPARISONS = frozenset(('==', '!=', '<', '<=', '>', '>='))

def _to_serial(x: Any) -> Any:
    """A date as Excel's serial number; a timedelta as its days.

    A ``date`` is a whole number of days since 1899-12-30. A ``datetime``
    adds its time of day as a fraction, so it is always a float (noon on
    2024-03-31 is 45382.5). Anything else comes back unchanged.
    """
    if isinstance(x, datetime):
        seconds = x.hour * 3600 + x.minute * 60 + x.second + x.microsecond / 1e6
        return x.toordinal() - _EXCEL_EPOCH_ORDINAL + seconds / 86400
    if isinstance(x, date):
        return x.toordinal() - _EXCEL_EPOCH_ORDINAL
    if isinstance(x, timedelta):
        return x / _ONE_DAY
    return x


def _date_compare(a: Any, b: Any, op: Any) -> Any:
    """Compare the way Excel does when one side is a date.

    A date compares as its serial number, so it can be compared with a
    plain number too. Across types Excel never finds two values equal and
    orders every number (a date included) before any text, and text before
    TRUE/FALSE. The result is a Python ``bool``.
    """
    def key(x: Any) -> Any:
        if isinstance(x, bool):
            return (2, x)
        if isinstance(x, str):
            return (1, x)
        serial = _to_serial(x)
        if isinstance(serial, (int, float)):
            return (0, serial)
        return None

    ka, kb = key(a), key(b)
    if ka is None or kb is None:
        return '#VALUE!'
    return bool(op(ka, kb))


def _date_serial_op(a: Any, b: Any, op: Any, op_symbol: Optional[str]) -> Any:
    """Excel-style arithmetic and comparisons where a date is involved.

    Excel stores a date as its serial number (see ``_to_serial``) and
    every operator works on that number. This mirrors it on Python
    ``date`` / ``datetime`` values, so Python computes what the written
    workbook computes:

    - ``date ± number`` (or ``number + date``) → the date that many days
      later or earlier. A fraction of a day gives a ``datetime``. A
      ``timedelta`` counts as its days, fraction included.
    - ``date - date`` → the days between them: an ``int`` for two pure
      dates, a ``float`` when a ``datetime`` brings a time of day.
    - any other operation (``date * n``, ``n / date``, ``n - date``,
      ``date + date``, ``date % 7``, …) → a plain number computed on
      the serial, as Excel does.
    - comparisons → ``True`` / ``False`` (see ``_date_compare``).
    - text, or a day count that no date can hold → ``#VALUE!``.
    """
    if op_symbol in _COMPARISON_SYMBOLS:
        return _date_compare(a, b, op)
    a_is_date, b_is_date = isinstance(a, date), isinstance(b, date)
    shifts = ((op_symbol == '+' and a_is_date != b_is_date)
              or (op_symbol == '-' and a_is_date and not b_is_date))
    if shifts:
        when, offset = (a, b) if a_is_date else (b, a)
        try:
            if isinstance(offset, timedelta):
                step = offset
            elif isinstance(offset, (int, float)):
                step = timedelta(days=offset)
            else:
                return '#VALUE!'
            if op_symbol == '-':
                step = -step
            # A pure date moved by part of a day gains a time of day.
            if not isinstance(when, datetime) and step % _ONE_DAY:
                when = datetime.combine(when, time())
            return when + step
        except (ValueError, OverflowError):
            return '#VALUE!'

    sa, sb = _to_serial(a), _to_serial(b)
    if not isinstance(sa, (int, float)) or not isinstance(sb, (int, float)):
        return '#VALUE!'
    try:
        return op(sa, sb)
    except ZeroDivisionError:
        # A zero serial (1899-12-30, or an empty timedelta) divides like 0.
        return '#DIV/0!'
    except Exception:
        return '#VALUE!'


def _provisional_unit(unit_fn: Any, left: Any, right: Any, *operands: Any) -> Any:
    """``unit_fn(left, right)`` — except over a row whose loop has not
    closed (see :mod:`modeleon.core.loops`): its unit is not known yet,
    so a mismatch is not an error here; the loop derives every unit
    again when it closes, and reports a real mismatch there."""
    if any(getattr(v, '_awaits', None) or getattr(v, '_forward_of', None) is not None
           for v in operands):
        try:
            return unit_fn(left, right)
        except ValueError:
            return None
    return unit_fn(left, right)


_GRAIN_ORDER = {'day': 0, 'month': 1, 'quarter': 2, 'year': 3}


def _mixed_first_periods(left: Any, lt: Any, right: Any, rt: Any) -> str:
    """The refusal when two series start the same-named periods on
    different days: one window's first period is short (it starts inside
    its quarter or year), the other's is not or starts on another day."""
    def name(v: Any) -> str:
        label = v._display_name or getattr(v, '_python_name', None)
        if label:
            return repr(label)
        formula = getattr(v, 'formula', None)
        return f"the line {formula}" if formula else "an unnamed line"

    def first(v: Any, t: Any) -> str:
        if t.first_day is not None:
            return (f"{name(v)} starts its first period, {t.start}, on "
                    f"{t.first_day.isoformat()}")
        return f"{name(v)} starts with the whole period {t.start}"
    return (
        f"{name(left)} and {name(right)} do not line up: {first(left, lt)}; "
        f"{first(right, rt)}. Periods of the same name would hold different "
        f"days - build both in the same window."
    )


def _mixed_grains(left: Any, left_grain: str, right: Any, right_grain: str) -> str:
    """Two lines of different grains in one formula: name both, their
    grains, and the re-grain that joins them."""
    adjective = {'day': 'daily', 'month': 'monthly', 'quarter': 'quarterly',
                 'year': 'yearly'}

    def describe(v: Any, grain: str) -> str:
        label = (getattr(v, '_display_name', None)
                 or getattr(v, '_python_name', None))
        return f"{label!r} ({grain}s)" if label else f"a {adjective.get(grain, grain)} line"

    finer = left if _GRAIN_ORDER.get(left_grain, 0) < _GRAIN_ORDER.get(right_grain, 0) else right
    coarse = right_grain if finer is left else left_grain
    code = getattr(finer, '_python_name', None)
    hint = f"{code}.at('{coarse}')" if code else f".at('{coarse}')"
    return (
        f"{describe(left, left_grain)} and {describe(right, right_grain)} meet "
        f"in one formula at different grains. Re-grain the finer line first: "
        f"{hint} gives it in {coarse}s (date alignment only aligns within one "
        f"grain)."
    )


#: Unit words read as a year or a month, for the per-period hint.
_YEAR_WORDS = frozenset({"year", "years", "yr", "yrs", "y", "год", "года", "лет"})
_MONTH_WORDS = frozenset({"month", "months", "mo", "mon", "мес", "месяц", "месяца"})


def _incompatible_units(left: Any, right: Any) -> str:
    """Why two units do not add — and, when one side is the other times a
    word or two, that those words are dimensions: a string unit always
    is, and a word that labels a pure number is declared as one."""
    lc, rc = left._unit_components, right._unit_components
    if str(left) == str(right) and bool(lc) != bool(rc):
        def what(u: Any) -> str:
            return "a dimension" if u._unit_components else "a labelled pure number"
        pair = f"{left} ({what(left)}) and {right} ({what(right)})"
    else:
        pair = f"{left} and {right}"
    message = f"Cannot add/subtract values with incompatible units: {pair}."
    # One side is the other times words, each to the first power: what is
    # left of each side once the other's symbols (at the same exponent) are
    # taken away is empty on one side and plain words on the other.
    left_over = {k: v for k, v in lc.items() if rc.get(k) != v}
    right_over = {k: v for k, v in rc.items() if lc.get(k) != v}
    smaller, larger = ((lc, right_over) if not left_over else
                       (rc, left_over) if not right_over else (None, None))
    if (smaller and larger and all(v == 1 for v in larger.values())
            and not set(larger) & set(smaller)):
        words = sorted(larger)
        named = " and ".join(f"'{w}'" for w in words)
        verb = "is a dimension" if len(words) == 1 else "are dimensions"
        example = words[0]
        message += (
            f" {named} {verb} here: a string unit always is. If it labels a "
            f"pure number (a factor, a share, a flag), declare it once - "
            f"{example.upper() if example.isidentifier() else 'NAME'} = "
            f"mo.Unit.dimensionless('{example}') - and set that object as the "
            f"unit of the lines that carry it."
        )
    elif (smaller and larger and all(v == -1 for v in larger.values())
            and not set(larger) & set(smaller)):
        words = sorted(larger)
        named = " and ".join(f"'{w}'" for w in words)
        message += (f" One side is per {named}: it becomes an amount once "
                    f"multiplied by a quantity in {named}")
        word = words[0]
        if len(words) == 1 and word.lower() in _YEAR_WORDS:
            message += (
                f" - the share of the year a period covers: "
                f"year_share = (mo.time.months / 12).set_unit('{word}') for a "
                f"fee or a salary, (mo.time.days / 365) for interest by the "
                f"day - then rate * year_share."
            )
        elif len(words) == 1 and word.lower() in _MONTH_WORDS:
            message += (
                f" - the months a period covers: "
                f"mo.time.months.set_unit('{word}'), then rate * that."
            )
        else:
            message += "."
    return message + (
        " Otherwise convert one side to the other's unit before the "
        "operation, or drop the unit on one side if the arithmetic is "
        "intentional."
    )


class _VariableArithmetic:
    """Operator-overloading mixin for :class:`Variable`.

    The attributes/methods declared below are provided by the concrete
    :class:`Variable` subclass — declared here under ``TYPE_CHECKING``
    so type-checkers know the mixin expects them on ``self``. They have
    no runtime effect.
    """

    # Attributes provided by the Variable subclass — type-only annotations.
    _value: Any
    _expr: Optional[Expr]
    _keys: Optional[List[Any]]
    _unit: Optional["Unit"]
    var_type: str
    value_type: str

    if TYPE_CHECKING:
        def _as_operand(self, parenthesize: bool = False) -> Expr: ...
        def _set_expr(self, expr: Expr) -> None: ...

    # ─── Broadcasting ───────────────────────────────────────────
    #
    # COMPUTE CHOKEPOINT. Every arithmetic value produced by the engine
    # flows through this function (operator overloads in
    # ``_VariableArithmetic`` call it). Do not add parallel compute
    # paths — keep all value arithmetic routed through here.

    @staticmethod
    def _reject_pending_roles(*operands) -> None:
        """Refuse eager compute on a role-kwarg Variable that has not
        materialized. Values compute eagerly, at operator time;
        an unmaterialized operand is ``None`` and would bake a permanent
        ``#VALUE!`` into every dependent — loud beats silently wrong.
        """
        for v in operands:
            if (getattr(v, '_role_kwargs', None) is not None
                    and not getattr(v, '_roles_materialized', False)):
                label = (getattr(v, '_display_name', None)
                         or getattr(v, '_python_name', None) or 'variable')
                raise ValueError(
                    f"{label!r} authors track coordinates that have not "
                    f"materialized yet — attach it to the model (values "
                    f"compute at operator time, so the attach must come "
                    f"BEFORE the formulas that read it)."
                )

    @staticmethod
    def _broadcast_operation(left, right, op, safe_divide: bool = False,
                             op_symbol: Optional[str] = None):
        """Element-wise op with NumPy-style broadcasting.

        Scalar/scalar → scalar. List/scalar → element-wise list.
        Scalar/list → element-wise list. List/list → element-wise;
        length-1 operands broadcast over the longer list.

        Excel-style error markers (``'#DIV/0!'``, ``'#NAME?'``, …)
        propagate through operations. With ``safe_divide=True``, division
        by zero yields ``'#DIV/0!'`` instead of raising.

        ``op_symbol`` names the operator (``'+'``, ``'<'``, …). Dates need
        it: ``date + 5`` is a date, ``date * 5`` is a number and
        ``date < 5`` is a bool, all computed on the date's serial number.
        """
        def _is_excel_error(v) -> bool:
            return isinstance(v, str) and v.startswith('#')

        def safe_op(a, b):
            if _is_excel_error(a):
                return a
            if _is_excel_error(b):
                return b
            # A blank counts as zero, as in Excel: a formula over a
            # month nobody entered computes over 0 (blank + 5 = 5,
            # 5 / blank = #DIV/0!). Only an input cell stays blank;
            # whether an actual was entered is the blend's question,
            # asked of its entries (``core.entries``).
            if a is None:
                a = '' if isinstance(b, str) and op_symbol in _COMPARISONS else 0
            if b is None:
                b = '' if isinstance(a, str) and op_symbol in _COMPARISONS else 0
            if safe_divide and b == 0:
                return '#DIV/0!'
            if isinstance(a, (date, timedelta)) or isinstance(b, (date, timedelta)):
                return _date_serial_op(a, b, op, op_symbol)
            try:
                return op(a, b)
            except Exception:
                return '#VALUE!'

        left_is_list = isinstance(left, list)
        right_is_list = isinstance(right, list)

        if not left_is_list and not right_is_list:
            return safe_op(left, right)
        elif left_is_list and not right_is_list:
            return [safe_op(x, right) for x in left]
        elif not left_is_list and right_is_list:
            return [safe_op(left, x) for x in right]
        else:
            if len(left) != len(right):
                if len(left) == 1:
                    return [safe_op(left[0], x) for x in right]
                elif len(right) == 1:
                    return [safe_op(x, right[0]) for x in left]
                raise ValueError(
                    f"Cannot broadcast lists of different lengths: "
                    f"{len(left)} vs {len(right)}. "
                    f"Element-wise ops need equal lengths, or one side of "
                    f"length 1 which broadcasts over the other. "
                    f"Common fix: align the two list-valued inputs to the "
                    f"same period count."
                )
            return [safe_op(a, b) for a, b in zip(left, right)]

    @staticmethod
    def _align_keyed_values(left, right):
        """Reorder right's list to match left's key order when both are keyed.

        If only one side is keyed, or keys match already, pass through.
        Raises if key sets disagree — silently reordering a mismatch would
        produce a silently wrong result.
        """
        from .variable import Variable
        left_val = left._value if isinstance(left, Variable) else left
        right_val = right._value if isinstance(right, Variable) else right

        left_keys = getattr(left, '_keys', None) if isinstance(left, Variable) else None
        right_keys = getattr(right, '_keys', None) if isinstance(right, Variable) else None

        if (left_keys is not None and right_keys is not None
                and isinstance(right_val, list)):
            if set(left_keys) != set(right_keys):
                raise ValueError(
                    f"Cannot align Variables with different key sets: "
                    f"{left_keys} vs {right_keys}. "
                    f"Element-wise ops on keyed Variables need matching keys "
                    f"so each value has a partner. Either pick the same keys "
                    f"for both sides, or drop keys on one side (pass through "
                    f"``.value``) and use positional broadcasting instead."
                )
            if left_keys != right_keys:
                key_to_idx = {k: i for i, k in enumerate(right_keys)}
                right_val = [right_val[key_to_idx[k]] for k in left_keys]

        return left_val, right_val

    @staticmethod
    def _propagate_keys(left, right):
        """Result keys come from whichever operand has them; left wins if both."""
        from .variable import Variable
        left_keys = getattr(left, '_keys', None) if isinstance(left, Variable) else None
        right_keys = getattr(right, '_keys', None) if isinstance(right, Variable) else None
        if left_keys is not None:
            return list(left_keys)
        if right_keys is not None:
            return list(right_keys)
        return None

    # ─── Units ──────────────────────────────────────────────────
    #
    # A Unit is a value: its algebra always builds a new one and nothing
    # changes a unit once built. So when the result of ``*`` / ``/`` IS
    # one operand's unit — the other side is unitless or dimensionless
    # (``$ * %``, ``$ / 12``) — that operand's own unit is the answer,
    # exactly as ``+`` / ``-`` already return theirs. Returning it avoids
    # building an equal copy (and a dimensionless Unit for the bare side)
    # on every operation; the algebra itself is unchanged.

    @staticmethod
    def _unit_mul(left_unit, right_unit):
        """Resulting unit for multiplication."""
        if left_unit is None and right_unit is None:
            return None
        left_c = left_unit._unit_components if left_unit is not None else None
        right_c = right_unit._unit_components if right_unit is not None else None
        if not right_c:
            return left_unit if left_c else None
        if not left_c:
            return right_unit
        result = left_unit * right_unit
        return result if result._unit_components else None

    @staticmethod
    def _unit_div(left_unit, right_unit):
        """Resulting unit for division."""
        if left_unit is None and right_unit is None:
            return None
        left_c = left_unit._unit_components if left_unit is not None else None
        right_c = right_unit._unit_components if right_unit is not None else None
        if not right_c:
            return left_unit if left_c else None
        from .unit import Unit
        lu = left_unit if left_unit is not None else Unit.dimensionless()
        result = lu / right_unit
        return result if result._unit_components else None

    @staticmethod
    def _unit_add(left_unit, right_unit):
        """Resulting unit for addition/subtraction.

        Both-None → None. One-sided (scalar + unitful) inherits the
        unit. Both-sides-present must agree — otherwise raise, because
        silently returning one side would mask a real modeling bug.
        """
        if left_unit is None and right_unit is None:
            return None
        if left_unit is not None and right_unit is not None:
            if left_unit != right_unit:
                raise ValueError(_incompatible_units(left_unit, right_unit))
            return left_unit
        return left_unit if left_unit is not None else right_unit

    @staticmethod
    def _unit_pow(unit, exponent):
        """Resulting unit for ``x ** n``.

        Only defined for integer exponents (``hr ** 2 → hr^2``). Non-int
        exponents drop the unit — Modeleon doesn't track fractional-power
        units because they rarely appear in financial modeling and the
        rendering gets ugly.
        """
        if unit is None or not isinstance(exponent, int):
            return None
        result = unit ** exponent
        # ``hr ** 0`` → dimensionless; collapse the empty-components
        # Unit to ``None`` to match the ``mul`` / ``div`` convention.
        return result if result._unit_components else None

    def _get_unit(self):
        """Get unit from self, or None for scalars."""
        return getattr(self, '_unit', None)

    # ─── Binary operators ──────────────────────────────────────

    def _binary_op(self, other, op_symbol: str, op_func, unit_fn,
                   safe_divide: bool = False, parenthesize: bool = False,
                   reversed: bool = False) -> "Variable":
        """Shared implementation for all binary arithmetic operators.

        Args:
            other: Right operand — another Variable, a scalar, or a list.
            op_symbol: Text inserted between operands in the formula string.
            op_func: Element-wise function applied for value computation.
            unit_fn: ``(left_unit, right_unit) -> result_unit`` or ``None``.
                     Pass ``None`` for operators that don't track units
                     (``**``, ``%``, and ``//`` when scalar is on the right).
            safe_divide: Route through the safe-division path so 1/0 becomes
                         ``"#DIV/0!"`` instead of raising.
            parenthesize: Wrap compound-unnamed-formula operands in parens
                          to preserve precedence (``(a+b) / (c+d)``).
            reversed: ``True`` for reflected ops (``__rsub__``, etc.).
                      In the formula string, operand order flips; dependency
                      refs still record ``[self, other]`` in that order.
        """
        from .variable import Variable

        self_u = self._get_unit()
        self._reject_pending_roles(self, other)

        if isinstance(other, Variable):
            left, right = (other, self) if reversed else (self, other)
            expr_node = BinOp(op_symbol, left._as_operand(parenthesize),
                              right._as_operand(parenthesize))
            result = Variable()
            result._set_expr(expr_node)
            lv, rv = self._align_keyed_values(left, right)
            expanded = self._expand_axised_operands(left, right, lv, rv)
            if expanded is not None:
                lv, rv = expanded
            aligned = self._align_by_date(left, right, lv, rv)
            if aligned is not None:
                lv, rv, _union_start, _union_grain, _union_first = aligned
                # The result is located at the union window explicitly —
                # downstream alignment then knows where it starts.
                result._start = _union_start
                result._grain = _union_grain
                result._first_day = _union_first
            from .tracks import TrackValues
            if isinstance(lv, TrackValues) or isinstance(rv, TrackValues):
                # The tracks broadcast law: zip per role;
                # a plain operand broadcasts into every track;
                # differing role sets die loudly inside ``combine``.
                result._value = TrackValues.combine(
                    lv, rv,
                    lambda a, b: self._broadcast_operation(
                        a, b, op_func, safe_divide=safe_divide, op_symbol=op_symbol
                    ),
                )
            else:
                result._value = self._broadcast_operation(
                    lv, rv, op_func, safe_divide=safe_divide, op_symbol=op_symbol
                )
            if unit_fn is not None:
                result._unit = _provisional_unit(
                    unit_fn, left._get_unit(), right._get_unit(), left, right)
        else:
            from .tracks import TrackValues

            def _combine(a, b):
                if isinstance(a, TrackValues) or isinstance(b, TrackValues):
                    return TrackValues.combine(
                        a, b,
                        lambda x, y: self._broadcast_operation(
                            x, y, op_func, safe_divide=safe_divide, op_symbol=op_symbol
                        ),
                    )
                return self._broadcast_operation(
                    a, b, op_func, safe_divide=safe_divide, op_symbol=op_symbol
                )

            if reversed:
                expr_node = BinOp(op_symbol, Literal(other), self._as_operand(parenthesize))
                val = _combine(other, self._value)
            else:
                expr_node = BinOp(op_symbol, self._as_operand(parenthesize), Literal(other))
                val = _combine(self._value, other)
            result = Variable()
            result._set_expr(expr_node)
            result._value = val
            if unit_fn is not None:
                # Scalars are treated as dimensionless when computing the output unit
                result._unit = _provisional_unit(
                    unit_fn, *((None, self_u) if reversed else (self_u, None)), self)

        if isinstance(result._value, list) and len(result._value) > 1:
            result.var_type = "list"
        elif type(result._value).__name__ == 'TrackValues' and result._value.time_length:
            result.var_type = "list"
        # A date-valued result (e.g. ``start + 90``, ``end - 1``) must carry
        # the datetime value type so the writer formats the cell as a date.
        sample = (result._value[0] if isinstance(result._value, list) and result._value
                  else result._value)
        if isinstance(sample, (date, datetime)):
            result.value_type = "datetime"
        k_left, k_right = (other, self) if reversed else (self, other)
        result._keys = self._propagate_keys(k_left, k_right)
        result._indexed_by = self._propagate_indexed_by(k_left, k_right)
        return result

    @staticmethod
    def _align_by_date(left: Any, right: Any, lv: Any, rv: Any):
        """Date-aligned zip — calendar alignment of two located operands.

        Engages ONLY when both operands are list-valued Variables with
        resolvable time locations that disagree (different starts or
        different lengths at one grain). Same-window operands keep the
        byte-identical positional path. Different GRAINS refuse — that
        is re-grain territory, not alignment.

        Each side materializes onto the union window; cells outside an
        operand's own extent follow its DECLARED ``extend=`` rule —
        zero (a flow is absent), hold (a rate stays in force; held
        backward from the first value when extended into the past),
        or none/undeclared → a teaching error naming both extents.
        Returns ``(lv', rv', union_start, grain)`` or ``None`` when
        not applicable.
        """
        from .tracks import TrackValues
        from .variable import Variable
        if not isinstance(left, Variable) or not isinstance(right, Variable):
            return None
        lt, rt = left.time, right.time
        if (lt is not None and rt is not None and lt.start is not None
                and rt.start is not None and lt.grain == rt.grain
                and lt.first_day != rt.first_day):
            # One side's first period is short. It may open the combined
            # window, never sit inside it: on the other side the same
            # period starts on another day. Tracked lines too.
            stub, other = (left, right) if lt.first_day is not None else (right, left)
            st, ot = stub.time, other.time
            if ot.first_day is not None or st.start >= ot.start:
                raise ValueError(_mixed_first_periods(left, lt, right, rt))
        if isinstance(lv, TrackValues) or isinstance(rv, TrackValues):
            return None  # tracks share the ambient window by the extent law
        if not isinstance(lv, list) or not isinstance(rv, list):
            return None
        if lt is None or rt is None or lt.start is None or rt.start is None:
            return None
        if lt.start == rt.start and len(lv) == len(rv):
            return None
        if lt.grain != rt.grain:
            raise ValueError(_mixed_grains(left, lt.grain, right, rt.grain))
        from .time import _count, _parse
        grain = lt.grain
        la, ra = _parse(lt.start, grain), _parse(rt.start, grain)
        # Union window in period indices relative to the earlier start.
        # ``_count`` is INCLUSIVE (Jan→Apr = 4 labels); the offset is
        # one less.
        off_l = 0 if lt.start <= rt.start else _count(ra, la, grain) - 1
        off_r = 0 if rt.start <= lt.start else _count(la, ra, grain) - 1
        total = max(off_l + len(lv), off_r + len(rv))
        union_start = min(lt.start, rt.start)
        union_first = lt.first_day if lt.start <= rt.start else rt.first_day

        def project(var: "Variable", vals: list, off: int, other: str) -> list:
            from .expr import TimeRef
            time_expr = getattr(var, '_expr', None)
            if isinstance(time_expr, TimeRef):
                # Time is defined in every period, not only the window's.
                field = time_expr.field
                if field == 'index':
                    return [i - off + 1 for i in range(total)]
                from .time import time_field_series
                return time_field_series(field, union_start, grain, total, union_first)
            out = []
            rule = getattr(var, '_extend', None)
            for i in range(total):
                j = i - off
                if 0 <= j < len(vals):
                    out.append(vals[j])
                elif rule == 'zero':
                    out.append(0.0)
                elif rule == 'hold':
                    out.append(vals[0] if j < 0 else vals[-1])
                else:
                    label = (var._display_name
                             or getattr(var, '_python_name', None)
                             or 'variable')
                    raise ValueError(
                        f"{label!r} does not cover the combined window "
                        f"(it spans {len(vals)} period(s) from "
                        f"{var.time.start}; the other side spans "
                        f"{other}). Say what continues it: "
                        f"extend=mo.zero() (a flow) or extend=mo.hold() "
                        f"(a rate) — never a silent guess."
                    )
            return out

        l_span = f"{len(rv)} period(s) from {rt.start}"
        r_span = f"{len(lv)} period(s) from {lt.start}"
        return (project(left, lv, off_l, l_span),
                project(right, rv, off_r, r_span),
                union_start, grain, union_first)

    @staticmethod
    def _expand_axised_operands(left: Any, right: Any, lv: Any, rv: Any):
        """Identity-aligned expansion of two AXISED operands.

        Engages only when BOTH operands declare ``_indexed_by`` and the
        axis tuples differ (different sets → cross-product; same set in
        a different order → transpose). Each flat row-major value is
        expanded to the union shape (left-first axis order — the same
        order ``_propagate_indexed_by`` emits) by replication, so the
        subsequent element-wise zip computes the true cross-product
        instead of the historical 2-cell diagonal that used to hide
        under a 4-cell label.

        Returns ``(lv', rv')`` or ``None`` when not applicable (either
        side lacks declared axes — scalars broadcast fine and plain
        time-lists keep their positional semantics — or
        the axis tuples are identical, where the plain zip is already
        correct and cheaper).
        """
        la = getattr(left, '_indexed_by', ()) or ()
        ra = getattr(right, '_indexed_by', ()) or ()
        if not la or not ra:
            return None
        if tuple(id(a) for a in la) == tuple(id(a) for a in ra):
            return None
        if lv is None or rv is None:
            return None

        union = _VariableArithmetic._propagate_indexed_by(left, right)
        dims = [
            len(a._value) if isinstance(getattr(a, '_value', None), list) else 1
            for a in union
        ]
        total = 1
        for d in dims:
            total *= d
        pos_by_id = {id(a): i for i, a in enumerate(union)}

        def expand(axes: Tuple[Any, ...], flat: Any) -> list:
            vals = flat if isinstance(flat, list) else [flat]
            pos = [pos_by_id[id(a)] for a in axes]
            own_dims = [dims[p] for p in pos]
            out = []
            for cell in range(total):
                # Union coordinates, row-major (last axis fastest).
                rem, coords = cell, [0] * len(dims)
                for i in range(len(dims) - 1, -1, -1):
                    coords[i] = rem % dims[i]
                    rem //= dims[i]
                # Row-major index into the operand's own axis order.
                idx = 0
                for p, d in zip(pos, own_dims):
                    idx = idx * d + coords[p]
                out.append(vals[idx])
            return out

        return expand(la, lv), expand(ra, rv)

    @staticmethod
    def _propagate_indexed_by(left: Any, right: Any) -> Tuple[Any, ...]:
        """Union the operands' ``_indexed_by`` axes, preserving the
        order of first appearance (left wins ties).

        Operator-built Variables inherit shape from this union — e.g.
        ``revenue * cogs_pct`` of shapes ``(months, scenarios)`` and
        ``()`` produces a result of shape ``(months, scenarios)``;
        ``a * b`` over two distinct axes produces the cross-product
        shape ``(left_axes..., right_axes...)``.

        Compares axes by identity (``id``) — Variable's ``__eq__`` is
        overridden to return a comparison expression, so ``axis in
        seen`` would always go through that operator and produce
        nonsense. Identity is the right relation for "same axis Variable".
        """
        la = getattr(left, '_indexed_by', ()) if left is not None else ()
        ra = getattr(right, '_indexed_by', ()) if right is not None else ()
        seen: list = list(la)
        seen_ids = {id(a) for a in seen}
        for axis in ra:
            if id(axis) not in seen_ids:
                seen.append(axis)
                seen_ids.add(id(axis))
        return tuple(seen)

    # ─── Binary dunders ────────────────────────────────────────

    def __add__(self, other) -> "Variable":
        return self._binary_op(other, "+", lambda a, b: a + b, self._unit_add)

    def __radd__(self, other) -> "Variable":
        return self.__add__(other)

    def __sub__(self, other) -> "Variable":
        return self._binary_op(other, "-", lambda a, b: a - b, self._unit_add)

    def __rsub__(self, other) -> "Variable":
        return self._binary_op(other, "-", lambda a, b: a - b, self._unit_add, reversed=True)

    def __mul__(self, other) -> "Variable":
        return self._binary_op(other, "*", lambda a, b: a * b, self._unit_mul)

    def __rmul__(self, other) -> "Variable":
        return self.__mul__(other)

    def __truediv__(self, other) -> "Variable":
        return self._binary_op(other, "/", lambda a, b: a / b, self._unit_div,
                               safe_divide=True, parenthesize=True)

    def __rtruediv__(self, other) -> "Variable":
        return self._binary_op(other, "/", lambda a, b: a / b, self._unit_div,
                               safe_divide=True, parenthesize=True, reversed=True)

    def __pow__(self, other) -> "Variable":
        result = self._binary_op(other, "**", lambda a, b: a ** b, None)
        # Scalar integer exponent → raise the unit to that power.
        # Variable-on-the-right exponents leave the unit as None.
        from .variable import Variable
        if not isinstance(other, Variable):
            result._unit = self._unit_pow(self._get_unit(), other)
        return result

    def __rpow__(self, other) -> "Variable":
        return self._binary_op(other, "**", lambda a, b: a ** b, None, reversed=True)

    def __mod__(self, other) -> "Variable":
        return self._binary_op(other, "%", lambda a, b: a % b, None)

    def __rmod__(self, other) -> "Variable":
        return self._binary_op(other, "%", lambda a, b: a % b, None, reversed=True)

    def __floordiv__(self, other) -> "Variable":
        return self._binary_op(other, "//", lambda a, b: a // b, self._unit_div)

    def __rfloordiv__(self, other) -> "Variable":
        return self._binary_op(other, "//", lambda a, b: a // b, None, reversed=True)

    # ─── Unary dunders ─────────────────────────────────────────

    def __neg__(self) -> "Variable":
        """-a (unary negation)"""
        from .variable import Variable
        from .tracks import TrackValues
        result = Variable()
        result._set_expr(UnaryOp("-", self._as_operand(parenthesize=True)))

        def neg(x):
            if x is None:
                return 0                # -blank is 0, as in Excel
            if isinstance(x, str) and x.startswith('#'):
                return x
            try:
                return -x
            except TypeError:
                return '#VALUE!'

        def neg_series(v):
            return [neg(x) for x in v] if isinstance(v, list) else neg(v)

        if isinstance(self._value, TrackValues):
            result._value = TrackValues.lift(neg_series, self._value)
        else:
            result._value = neg_series(self._value)
        result._unit = self._get_unit()
        result._keys = list(self._keys) if self._keys is not None else None
        result._indexed_by = getattr(self, '_indexed_by', ())
        if isinstance(result._value, list) and len(result._value) > 1:
            result.var_type = 'list'
        elif type(result._value).__name__ == 'TrackValues' and result._value.time_length:
            result.var_type = 'list'
        return result

    def __pos__(self) -> "Variable":
        """+a is a no-op — return self."""
        return self  # type: ignore[return-value]  # mixin self is the concrete Variable at runtime

    # ─── Comparisons ───────────────────────────────────────────
    #
    # COMPUTE CHOKEPOINT for comparisons. Parallel to
    # ``_broadcast_operation`` for arithmetic — do not add parallel
    # compute paths.

    def _comparison_op(self, other, op_symbol: str, op_func) -> "Variable":
        """Shared implementation for all comparison operators."""
        from .variable import Variable

        # ``self`` is always a :class:`Variable` at runtime — the mixin is
        # only mounted on Variable. The casts below tell mypy that.
        self_var: "Variable" = self  # type: ignore[assignment]
        self._reject_pending_roles(self, other)
        from .tracks import TrackValues

        def _combine(a, b):
            if isinstance(a, TrackValues) or isinstance(b, TrackValues):
                return TrackValues.combine(
                    a, b, lambda x, y: self._broadcast_operation(
                        x, y, op_func, op_symbol=op_symbol
                    )
                )
            return self._broadcast_operation(a, b, op_func, op_symbol=op_symbol)

        if isinstance(other, Variable):
            result = Variable()
            result._set_expr(Compare(op_symbol, VarRef(self_var), VarRef(other)))
            lv, rv = self._align_keyed_values(self, other)
            expanded = self._expand_axised_operands(self, other, lv, rv)
            if expanded is not None:
                lv, rv = expanded
            result._value = _combine(lv, rv)
        else:
            result = Variable()
            result._set_expr(Compare(op_symbol, VarRef(self_var), Literal(other)))
            result._value = _combine(self._value, other)

        if isinstance(result._value, list) and len(result._value) > 1:
            result.var_type = 'list'
        elif type(result._value).__name__ == 'TrackValues' and result._value.time_length:
            result.var_type = 'list'
        result._keys = self._propagate_keys(self, other)
        result._indexed_by = self._propagate_indexed_by(self, other)
        return result

    # ─── Comparison dunders ────────────────────────────────────

    # ``__eq__`` / ``__ne__`` deliberately return a :class:`Variable`
    # (a comparison expression) rather than ``bool`` — the engine treats
    # ``a == b`` as a formula-building operator. The override mismatch
    # against ``object`` is intentional, not a bug.
    def __eq__(self, other) -> "Variable":  # type: ignore[override]
        return self._comparison_op(other, '==', lambda a, b: a == b)

    def __ne__(self, other) -> "Variable":  # type: ignore[override]
        return self._comparison_op(other, '!=', lambda a, b: a != b)

    def __lt__(self, other) -> "Variable":
        return self._comparison_op(other, '<', lambda a, b: a < b)

    def __le__(self, other) -> "Variable":
        return self._comparison_op(other, '<=', lambda a, b: a <= b)

    def __gt__(self, other) -> "Variable":
        return self._comparison_op(other, '>', lambda a, b: a > b)

    def __ge__(self, other) -> "Variable":
        return self._comparison_op(other, '>=', lambda a, b: a >= b)
