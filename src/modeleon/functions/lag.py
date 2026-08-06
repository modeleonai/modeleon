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
            live, so editing the fill assumption reflows the model.
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
    if fill_is_var and isinstance(fill._value, list):
        raise TypeError(
            "lag(..., fill=) must be a number or a SCALAR Variable — "
            "a list fill has no single head value."
        )
    fill_value = fill._value if fill_is_var else fill

    def _shift(values: list) -> list:
        if n > 0:
            k = min(n, len(values))
            return [fill_value] * k + values[: len(values) - k]
        if n < 0:
            k = min(-n, len(values))
            return values[k:] + [fill_value] * k
        return list(values)

    from ..core.tracks import TrackValues
    if isinstance(source._value, TrackValues):
        # Rank lifting (§15.3.5): the shift applies per coordinate —
        # plan shifts within plan, actuals within actuals.
        shifted: object = TrackValues.lift(
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
    unit = getattr(source, "_unit", None)
    if unit is not None:
        result._unit = unit

    label = getattr(source, "python_name", None) or source.path.leaf
    result._source_code = f"lag({label}, {n})"
    if display_name:
        result._display_name = display_name
    return result
