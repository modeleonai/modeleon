# SPDX-License-Identifier: Apache-2.0
"""Period shift — read a list Variable ``n`` periods back.

``lag(source, periods=1, fill=0.0)`` builds ``x[t] = source[t - periods]``
with the exposed head cells holding ``fill`` (a number or a scalar
Variable). Negative ``periods`` is a lead. The AST stores a ``shift``
:class:`~modeleon.core.expr.MethodCall`, which the Excel renderer lowers
to a direct reference to the source row's cell ``periods`` columns back —
the native spreadsheet spelling of a lag.

The one-period shift completes roll-forward algebra:
``opening = lag(closing, fill=opening_balance)`` and
``delta = x - lag(x)`` (period-over-period change).
"""

from __future__ import annotations

from datetime import date
from typing import Any, Optional

from ..core.expr import MethodCall, VarRef
from ..core.variable import Variable


def check_fill(fill: Any, spelled: str) -> None:
    """A lag's fill that is a line is kept as a reference to its cell, so it
    is one value assigned above the lag: a fill assigned further down, or a
    series, is refused (``spelled`` is how the call is written)."""
    if not isinstance(fill, Variable):
        return
    if fill._forward_of is not None:
        raise TypeError(
            f"{spelled}{fill._forward_of[1]}): the fill is assigned "
            f"further down — assign it above this line."
        )
    if isinstance(fill._value, list):
        raise TypeError(
            f"{spelled}) must be a number or a SCALAR Variable — "
            f"a list fill has no single head value."
        )


def lag_unit(source: Variable, fill: Any, result: Variable) -> Any:
    """The unit of a lag of ``source`` filled by ``fill`` — ``mo.lag``'s and
    ``.shift``'s, by the law a loop container derives the row by again
    (``core.loops._unit_of``). The source's unit; the fill is the series'
    first cells, so a fill in a unit is in the source's, as the two values of
    an ``IF`` share one (:func:`~modeleon.functions._helpers.same_unit`):
    another dimension is
    refused, two labels of a pure number name neither. A source read
    before its line takes the fill's, and its mismatch waits for the loop
    to close."""
    from ._helpers import operand_unit, same_unit, waits_on_loop
    waiting = waits_on_loop(source) or waits_on_loop(result)
    unit = getattr(source, "_unit", None)
    if unit is None and not waiting:
        return None
    fill_node = [VarRef(fill)] if isinstance(fill, Variable) else []
    try:
        return same_unit("lag", [VarRef(source), *fill_node], operand_unit)
    except ValueError:
        if not waiting:
            raise
        return None


def lag(
    source: Variable,
    periods: int = 1,
    fill: Any = 0.0,
    display_name: Optional[str] = None,
) -> Variable:
    """Shift a list Variable ``periods`` periods back (or forward).

    Args:
        source: A list-valued Variable (a time series).
        periods: How far back to look — ``lag(x)`` at period ``t``
            reads ``x[t-1]``. Negative values read forward (a lead).
        fill: Value for the positions the shift exposes (the first
            ``periods`` cells of a lag, the last of a lead). A number
            or a SCALAR Variable — a Variable keeps the dependency
            live, so editing the fill assumption reflows the model. The
            result is in the source's unit, and a fill in a unit is too:
            one in another is refused.
        display_name: Optional display label for the result.

    Excel: each cell references the source row's cell ``periods``
    columns back; exposed cells hold ``fill`` — a literal, or a
    reference to the fill Variable's own cell.
    """
    from ..core.tracks import TrackValues as _Tracks
    _is_tracked = isinstance(source, Variable) and isinstance(
        source._value, _Tracks
    )
    if not isinstance(source, Variable) or (
        not isinstance(source._value, list) and not _is_tracked
    ):
        raise TypeError(
            "lag(source, ...) needs a list-valued Variable (a time "
            "series) — a scalar has no periods to shift."
        )
    from ..core.shape import reject_axised
    reject_axised(source, "lag")
    reject_axised(fill, "lag(fill=)")
    try:
        n = int(periods)
    except (TypeError, ValueError):
        raise TypeError("lag(..., periods=) must be an int") from None

    fill_is_var = isinstance(fill, Variable)
    check_fill(fill, "lag(..., fill=")
    fill_value = fill._value if fill_is_var else fill

    from ..core.regrain import hole_value
    # A lagged blank reads as 0, as Excel's =previous_cell does - kept
    # blank only in a line of dates or words, which has no zero.
    blank = hole_value(source._value)

    def _shift(values: list) -> list:
        if n > 0:
            k = min(n, len(values))
            out = [fill_value] * k + values[: len(values) - k]
        elif n < 0:
            k = min(-n, len(values))
            out = values[k:] + [fill_value] * k
        else:
            out = list(values)
        return [blank if v is None else v for v in out]

    from ..core.tracks import TrackValues
    shifted: object
    if source._forward_of is not None:
        # A row read before its line (a loop, see ``core.loops``): the
        # shifted values are provisional until the loop closes. Hold the
        # fill in every period so they carry its type — a loop of dates
        # stays dates for the date functions built on it.
        n_periods = len(source._value)
        if isinstance(fill_value, date):
            source._value = [fill_value] * n_periods
        shifted = [fill_value] * n_periods
    elif isinstance(source._value, TrackValues):
        # Rank lifting: the shift applies per track coordinate —
        # plan shifts within plan, actuals within actuals.
        shifted = TrackValues.lift(
            lambda track: _shift(track) if isinstance(track, list) else track,
            source._value,
        )
    else:
        shifted = _shift(list(source._value))

    # ``fill_value`` rides in the AST only when it says something: a
    # VarRef keeps the dependency live (renderer emits the fill cell's
    # address); a non-zero literal must survive to the emitted formula;
    # the 0-default matches the renderer's own default.
    kwargs: dict = {}
    if fill_is_var:
        kwargs["fill_value"] = VarRef(fill)
    elif fill_value not in (0, 0.0):
        kwargs["fill_value"] = fill_value

    result = Variable(formula=MethodCall(VarRef(source), "shift", [n], kwargs))
    result._value = shifted
    result.var_type = "list"
    _sample = shifted
    if isinstance(shifted, TrackValues):
        _sample = next(
            (f for f in shifted.values() if isinstance(f, list)), []
        )
    if any(isinstance(v, date) for v in _sample):
        result.value_type = "datetime"
    elif any(isinstance(v, float) for v in _sample):
        result.value_type = "float"
    else:
        result.value_type = "int"
    unit = lag_unit(source, fill, result)
    if unit is not None:
        result._unit = unit

    label = getattr(source, "python_name", None) or source.path.leaf
    result._source_code = f"lag({label}, {n})"
    if display_name:
        result._display_name = display_name
    return result
