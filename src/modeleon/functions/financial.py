# SPDX-License-Identifier: Apache-2.0
"""Financial formula helpers — IRR, NPV, XIRR, XNPV, PMT, FV, PV; the
loan schedule NPER, IPMT, PPMT, RATE, CUMIPMT, CUMPRINC; depreciation
SLN, DB, DDB, SYD.

Every function builds a live Excel formula (``=IRR(B2:B6)``, ``=PMT(...)``)
alongside the Python-side value. Pass a Variable (including list-valued
ones) to get the formula; pass a raw list to get a value-only result —
the latter is useful for ad-hoc calculations outside the model tree.

Units: the amounts a function reads together — a loan's present and
future value, a payment and a balance, a cost and its salvage — are in
one unit, and two are refused, as ``+`` refuses them; the result is in
it (``NPV`` and ``XNPV`` in their flows'). A rate, a term, a period and
a type are never the result's unit, and a rate (``IRR``, ``RATE``) or a
count of periods (``NPER``) carries none.
"""

from __future__ import annotations

import decimal
import functools
import logging
import math
from datetime import date as _date_type
from typing import Any, List, Optional, Union

import numpy as np

from ..core.expr import FuncCall, Literal, VarRef
from ..core.variable import Variable
from ._helpers import (
    elementwise, function_unit, is_series, operand_unit, same_unit, unit_law, val,
    waits_on_loop)
from .registry import register
from .dates import _extract_date
from .mathfn import _spreadsheet_round
from .text import _number_of

logger = logging.getLogger(__name__)


__all__ = ['IRR', 'NPV', 'XIRR', 'XNPV', 'PMT', 'FV', 'PV',
           'NPER', 'IPMT', 'PPMT', 'RATE', 'CUMIPMT', 'CUMPRINC',
           'SLN', 'DB', 'DDB', 'SYD']


# ─── Shared helpers ───────────────────────────────────────────────


def _arg(x):
    """A scalar argument's value; a blank cell reads as 0, as in Excel."""
    v = val(x)
    return 0 if v is None else v


def _present(values: List[Any]) -> List[Any]:
    """The flows NPV / IRR read: Excel skips a blank cell in the range, so
    the flows after it close up."""
    return [v for v in values if v is not None]


def _zero_blanks(values: List[Any]) -> List[Any]:
    """The flows XNPV / XIRR read: each sits on its date, so a blank one
    is a 0 on that date."""
    return [0 if v is None else v for v in values]


#: The refusal of a list holding a row, by function: one row has a range
#: in the book, a cell and a row side by side have none.
_ONE_ROW = {
    "IRR": "Give the outlay a period of its own in the row - with a window "
           "that starts at the close, the first: flows = mo.IF(mo.time.index "
           "== 1, -cost, flow).set_regrain(mo.up('sum')) - then mo.IRR(flows).",
    "NPV": "A time-0 outlay stays outside the call, as Excel discounts the "
           "first flow a whole period: outlay + mo.NPV(rate, flows); a later "
           "amount (a terminal value) goes into the row in its own period.",
    "XIRR": "Give the outlay a period of its own - with a window that starts "
            "at the close, the first: flows = mo.IF(mo.time.index == 1, -cost, "
            "flow).set_regrain(mo.up('sum')), dates = mo.IF(mo.time.index == 1, "
            "close, mo.time.end).set_regrain(mo.up('last')) - then "
            "mo.XIRR(flows, dates).",
    "XNPV": "Give the outlay a period of its own - with a window that starts "
            "at the close, the first: flows = mo.IF(mo.time.index == 1, -cost, "
            "flow).set_regrain(mo.up('sum')), dates = mo.IF(mo.time.index == 1, "
            "close, mo.time.end).set_regrain(mo.up('last')) - then "
            "mo.XNPV(rate, flows, dates).",
}


def _is_row(x: Any) -> bool:
    return isinstance(x, Variable) and (isinstance(x._value, list)
                                        or hasattr(x._value, "roles"))


def _refuse_a_row_in_a_list(items: Any, func: str) -> None:
    """A list holding a row (a cell beside a row) has no range in the book:
    the result could not stay live there, and numpy would fail on it."""
    if not isinstance(items, (list, tuple)) or not any(_is_row(x) for x in items):
        return
    if len(items) == 1:
        raise ValueError(
            f"{func} takes the row itself, not a list holding it: drop the "
            f"brackets around it."
        )
    what = ("one row of flows and one row of dates" if func in ("XIRR", "XNPV")
            else "one row of flows")
    raise ValueError(
        f"{func} takes {what}; a list holding a row has no range in the "
        f"book, so the result could not stay live there. {_ONE_ROW[func]}"
    )


def _extract_cashflow_values(cash_flows, func: str) -> List[float]:
    """Pull numeric cash-flow values from a Variable or plain list."""
    if isinstance(cash_flows, Variable):
        values = cash_flows._value
        return list(values) if isinstance(values, list) else [values]
    if isinstance(cash_flows, list):
        _refuse_a_row_in_a_list(cash_flows, func)
        # A cell in the list is its number. Left a Variable, it iterates as
        # a one-item sequence and numpy built an n x n grid of the flows:
        # the NPV of two 50s at 10% read 173.55, not 86.78.
        return [val(x) if isinstance(x, Variable) else x for x in cash_flows]
    raise TypeError(
        f"cash_flows must be a Variable or list — "
        f"got {type(cash_flows).__name__}. Example:\n"
        f"    cf = mo.Variable([-100_000, 30_000, 40_000, 50_000])\n"
        f"    mo.IRR(cf)"
    )


_EXCEL_ONLY = frozenset({'excel'})
"""Render-backend set for Excel-native financial functions — IRR / NPV /
XIRR / XNPV / PMT / FV / PV. Excel emits these natively; other renderers
fall back to inlining ``_value``."""


def _wrap_with_formula(value: float, func_name: str, args: list) -> Variable:
    """Build a scalar-valued result Variable with an Excel FuncCall formula,
    in the unit the function's law gives."""
    result = Variable(value_type='float')
    result._set_expr(FuncCall(
        func_name, args,
        render_backends=_EXCEL_ONLY,
    ))
    result._value = float(value)
    result.var_type = 'scalar'
    result.value_type = 'float'
    try:
        unit = function_unit(func_name, args, operand_unit)
    except ValueError:
        if not waits_on_loop(result):       # provisional until its loop closes
            raise
        unit = None
    if unit is not None:
        result._unit = unit
    return result


def _formula_arg(value) -> "Literal | VarRef":
    """AST node for a function argument — VarRef if it's a Variable, else Literal."""
    return VarRef(value) if isinstance(value, Variable) else Literal(value)


def _range_arg(value: Any) -> Any:
    """AST node for a flows or dates argument: a row is its range; a list
    holding lines is the list of their cells, each its own node, as a
    criteria range is - the book writes their cells and the unit law
    reads each flow's unit; a list of numbers typed in is a literal."""
    from ..core.expr import ListExpr
    if isinstance(value, (list, tuple)) and any(isinstance(x, Variable) for x in value):
        return ListExpr([_formula_arg(x) for x in value])
    return _formula_arg(value)


# ─── Core algorithms ──────────────────────────────────────────────


def _npv(rate: float, cash_flows: np.ndarray) -> float:
    """NPV of a cash-flow series at a flat per-period rate.

    Matches Excel's ``NPV`` convention: the first cash flow is at the
    *end of period 1*, not at time 0. To value a time-0 cash flow,
    keep it outside the call: ``cf0 + NPV(rate, future_flows)``.
    """
    periods = np.arange(1, len(cash_flows) + 1)
    return float(np.sum(cash_flows / ((1 + rate) ** periods)))


def _irr_newton(
    cash_flows: np.ndarray, guess: float, max_iter: int, tol: float,
    times: Optional[np.ndarray] = None,
) -> float:
    """Newton-Raphson root-find for the rate where NPV = 0.

    ``times`` are the exponents the flows are discounted by: whole periods
    ``0, 1, 2, …`` when omitted (IRR), years from the first date for XIRR.

    Raises ``RuntimeError`` when the derivative vanishes or iteration
    fails to converge — callers (IRR / XIRR) translate these into
    user-facing errors.
    """
    periods = np.arange(len(cash_flows)) if times is None else times
    rate = float(guess)
    for _ in range(max_iter):
        # A run-away step overflows the discount factors; the caller falls back
        # to bisection, so the overflow is not news for the user.
        with np.errstate(over="ignore", invalid="ignore", divide="ignore"):
            npv = np.sum(cash_flows / ((1 + rate) ** periods))
            dnpv = np.sum(-periods * cash_flows / ((1 + rate) ** (periods + 1)))
        if not (np.isfinite(npv) and np.isfinite(dnpv)):
            raise RuntimeError("IRR failed: Newton's step ran off the rate axis.")
        if abs(npv) < tol:
            return rate
        if abs(dnpv) < 1e-10:
            raise RuntimeError(
                "IRR failed: derivative near zero. Try a different guess "
                "or verify the cash flows have a well-defined IRR."
            )
        rate -= npv / dnpv
        if rate < -0.99:
            rate = -0.99
    raise RuntimeError(
        "IRR failed to converge within the iteration limit. Try a "
        "different guess closer to the expected rate."
    )


#: Rates a bracket for the fallback is sought on. Below -50% a period the
#: discount factors of a long series overflow; no project IRR lives there.
_IRR_GRID = (-0.5, -0.2, -0.1, -0.05, -0.02, -0.01, 0.0, 0.002, 0.005, 0.01, 0.02,
             0.05, 0.1, 0.2, 0.5, 1.0, 2.0, 5.0, 10.0, 100.0, 1000.0)

#: The same for XIRR, whose rate is ANNUAL. A lost deal lives far below
#: -50% a year (-100 in, +1 back a year later is -99%), and a short span
#: annualises into huge rates (Excel: 1,937.56 = 193,756% for flows six
#: days apart), so the grid reaches further on both sides. Discount
#: factors that overflow at the extremes drop out as non-finite.
_XIRR_GRID = (-0.9999, -0.999, -0.99, -0.95, -0.9, -0.8, -0.7, -0.6, -0.5, -0.4,
              -0.3, -0.2, -0.1, -0.05, -0.02, -0.01, 0.0, 0.01, 0.02, 0.05, 0.1,
              0.2, 0.3, 0.5, 0.75, 1.0, 2.0, 5.0, 10.0, 100.0, 1e3, 1e4, 1e6, 1e9)

#: A year in XIRR and XNPV is 365 days — Excel's documented formula is
#: Σ P_i / (1 + rate)^((d_i − d_1) / 365), in leap years too. Not 365.25,
#: not actual/actual: those land on a different rate.
_DAYS_PER_YEAR = 365


def _npv_from_zero(
    rate: float, cash_flows: np.ndarray, times: Optional[np.ndarray] = None,
) -> float:
    """NPV with the first flow at t=0 — the function IRR finds the root of.
    ``times`` as in :func:`_irr_newton`."""
    periods = np.arange(len(cash_flows)) if times is None else times
    with np.errstate(over="ignore", invalid="ignore", divide="ignore"):
        value = np.sum(cash_flows / ((1 + rate) ** periods))
    return float(value)


def _changes_sign_at(
    cash_flows: np.ndarray, rate: float, times: Optional[np.ndarray] = None,
) -> bool:
    """Whether the NPV crosses zero at ``rate`` — a root, not a rate where the
    NPV is merely small."""
    step = 1e-6 * (1.0 + abs(rate))
    left = _npv_from_zero(rate - step, cash_flows, times)
    right = _npv_from_zero(rate + step, cash_flows, times)
    return np.isfinite(left) and np.isfinite(right) and left * right <= 0.0


def _irr_bisect(
    cash_flows: np.ndarray, guess: float, times: Optional[np.ndarray] = None,
    grid: tuple = _IRR_GRID,
) -> float:
    """Bisection inside the sign change nearest ``guess`` on ``grid``
    (:data:`_IRR_GRID` by default)."""
    points = [(r, _npv_from_zero(r, cash_flows, times)) for r in grid]
    points = [(r, v) for r, v in points if np.isfinite(v)]
    brackets = [
        (a, fa, b) for (a, fa), (b, fb) in zip(points, points[1:]) if fa * fb <= 0.0
    ]
    if not brackets:
        raise RuntimeError(
            "IRR failed: the NPV does not change sign between -50% and 100,000% a "
            "period — verify the cash flows have a well-defined IRR."
        )
    lo, f_lo, hi = min(brackets, key=lambda b: min(abs(b[0] - guess), abs(b[2] - guess)))
    for _ in range(200):
        mid = (lo + hi) / 2
        f_mid = _npv_from_zero(mid, cash_flows, times)
        if f_lo * f_mid <= 0.0:
            hi = mid
        else:
            lo, f_lo = mid, f_mid
        if hi - lo <= 1e-15 * (1.0 + abs(mid)):
            break
    return (lo + hi) / 2


def _irr_solve(
    cash_flows: np.ndarray, guess: float, times: Optional[np.ndarray] = None,
    grid: tuple = _IRR_GRID, tol: float = 1e-6,
) -> float:
    """The rate at which the NPV of ``cash_flows`` is zero — per period for
    IRR, per year for XIRR.

    ``times`` gives each flow's exponent — whole periods when omitted (IRR),
    years since the first date for XIRR — and ``grid`` the rates the
    bisection fallback seeks a bracket on; ``tol`` is Newton's NPV tolerance.

    Leading zero flows are dropped first, and the times re-based on the
    first flow that is not zero. They multiply the NPV by ``(1+r)^-k`` and
    move no root, but they let a large rate shrink the NPV under the
    absolute tolerance without being a root — measured: eleven idle months
    made Newton "converge" at 433% a month where the root is 1.17%.

    Newton from ``guess`` stays the path, so of several roots the one nearest
    the guess is found, as in Excel. Its answer is kept only where the NPV
    really crosses zero; otherwise, and when Newton runs away (a 10% guess on
    monthly flows), bisection inside a bracket finds the root.
    """
    first = int(np.flatnonzero(cash_flows)[0])
    flows = cash_flows[first:]
    if times is not None:
        times = times[first:] - times[first]
    try:
        rate = _irr_newton(flows, guess, max_iter=100, tol=tol, times=times)
    except RuntimeError:
        rate = None
    if rate is not None and _changes_sign_at(flows, rate, times):
        return rate
    return _irr_bisect(flows, guess, times, grid)


def _years_from_first(dates, count: int, func: str, *, earlier_allowed: bool) -> np.ndarray:
    """``(d_i − d_1) / 365`` for each date — the exponent Excel's XIRR and
    XNPV discount a flow by, counted from the FIRST date listed (not the
    earliest). Dates are read as the date functions read them (date,
    datetime truncated to its day, ISO string, or a Variable of those).

    Excel returns ``#NUM!`` when the counts differ, and XNPV does when a
    date precedes the first one; the engine raises instead of showing a
    number the workbook will not. XIRR computes such a schedule (Excel
    does, whatever its help page says), so ``earlier_allowed`` says which.
    """
    raw = dates._value if isinstance(dates, Variable) else dates
    _refuse_a_row_in_a_list(None if isinstance(dates, Variable) else raw, func)
    if not isinstance(raw, (list, tuple)):
        raise TypeError(
            f"{func} dates must be a list or a list-valued Variable, got "
            f"{type(raw).__name__}."
        )
    days = [_extract_date(d) for d in raw]
    if len(days) != count:
        raise ValueError(
            f"{func} needs one date per cash flow — got {len(days)} dates for "
            f"{count} flows (Excel returns #NUM!)."
        )
    first = days[0]
    if not earlier_allowed:
        early = [d for d in days if d < first]
        if early:
            raise ValueError(
                f"{func}: date {early[0].isoformat()} precedes the first date "
                f"{first.isoformat()} — Excel returns #NUM!. The first date "
                f"is the one flows are discounted to; list the earliest first."
            )
    return np.array([(d - first).days for d in days], dtype=float) / _DAYS_PER_YEAR


# ─── Public functions ─────────────────────────────────────────────


# The rate a search starts from is one number for the whole row.
def _one_guess(guess: Any, func: str) -> None:
    if isinstance(val(guess), (list, tuple)):
        raise TypeError(
            f"{func}'s guess is one number — the rate the search starts from — "
            f"got a series. Give one value, or leave it out (0.1)."
        )


def IRR(
    cash_flows: Union[Variable, List[float]],
    guess: float = 0.1,
) -> Variable:
    """Internal Rate of Return — renders ``=IRR(cash_flows_range, guess)``.

    IRR is the per-period rate at which the NPV of ``cash_flows`` equals zero.
    Requires at least one negative and one positive cash flow.

    When ``cash_flows`` is a Variable, the result
    carries an Excel formula referencing the cash-flow range. When it's a
    raw Python list, only the computed value is returned (no formula).

    Example::

        cf = mo.Variable([-100_000, 30_000, 40_000, 50_000, 60_000])
        project_irr = mo.IRR(cf)       # Excel: =IRR(B2:F2, 0.1)
    """
    _one_guess(guess, "IRR")
    cf_values = _present(_extract_cashflow_values(cash_flows, "IRR"))
    if len(cf_values) < 2:
        raise ValueError(
            f"IRR requires at least 2 cash flows, got {len(cf_values)}."
        )
    cf_array = np.array(cf_values, dtype=float)
    if not (np.any(cf_array < 0) and np.any(cf_array > 0)):
        raise ValueError(
            "IRR is undefined when all cash flows have the same sign — "
            "at least one negative (outflow) and one positive (inflow) "
            "value are required."
        )

    irr_value = _irr_solve(cf_array, guess)

    flows = _range_arg(cash_flows)
    if not isinstance(flows, Literal):
        return _wrap_with_formula(irr_value, 'IRR', [flows, Literal(guess)])
    # Raw list input — no cell range to reference; return value-only.
    return Variable(irr_value, value_type='float')


def NPV(
    rate: Union[float, Variable],
    cash_flows: Union[Variable, List[float]],
) -> Variable:
    """Net Present Value — renders ``=NPV(rate, cash_flows_range)``.

    Follows Excel's convention: the first value in ``cash_flows`` is
    discounted as if it occurs at the *end of period 1*. The initial
    (time-0) outlay must stay outside the call.

    Example::

        r        = mo.Variable(0.10)
        initial  = mo.Variable(-100_000)                   # at t=0
        future   = mo.Variable([30_000, 40_000, 50_000])   # t=1, 2, 3
        project_npv = initial + mo.NPV(r, future)
        # Excel: =B1 + NPV(B2, C3:E3)
    """
    rate_value = _arg(rate)
    if isinstance(rate_value, (list, tuple)):
        raise TypeError(
            "NPV takes a single discount rate — got a series. For a rate that "
            "changes over time, discount each flow explicitly."
        )
    cf_values = _present(_extract_cashflow_values(cash_flows, "NPV"))
    npv_value = _npv(rate_value, np.array(cf_values, dtype=float))

    flows = _range_arg(cash_flows)
    if isinstance(rate, Variable) or not isinstance(flows, Literal):
        return _wrap_with_formula(npv_value, 'NPV', [_formula_arg(rate), flows])
    return Variable(npv_value, value_type='float')


#: Newton's NPV tolerance for XIRR, relative to the size of the flows
#: (Σ|P|). An absolute tolerance promises no accuracy of the RATE: 1e-6 is
#: loose on flows written in millions and out of float reach on flows in
#: billions. Relative 1e-12 lands the rate far inside Excel's own
#: "accurate within 0.000001 percent".
_XIRR_NPV_RTOL = 1e-12


def XIRR(
    cash_flows: Union[Variable, List[float]],
    dates: Union[Variable, List[_date_type]],
    guess: Union[float, Variable] = 0.1,
) -> Variable:
    """Extended IRR — renders ``=XIRR(cash_flows_range, dates_range, guess)``.

    The annual rate at which :func:`XNPV` of the flows is zero, Excel's
    formula: each flow discounted by ``(1 + rate) ** ((d_i - d_1) / 365)``,
    365 days a year, counted from the FIRST date listed.

    Solved as :func:`IRR` is — Newton from ``guess``, kept only where the
    NPV really changes sign, else bisection inside a bracket — so leading
    zero flows and long monthly schedules give the true rate, not a false
    root. Where Excel's own iteration succeeds the two agree within Excel's
    documented accuracy (0.000001 percent); the engine returns the root to
    machine precision. Measured in Excel 16.0 (build 20326), Excel is the
    one that fails on two inputs: a FIRST flow of zero makes it return
    2.98e-09 (= 0.1 / 2**25) whatever the rate, and some negative guesses
    return ``#NUM!`` although the root exists. The engine keeps the root.

    Raises ``ValueError`` where Excel returns ``#NUM!`` by rule: flows all
    of one sign, a date count that differs from the flow count, a guess of
    -100% or below. A date earlier than the first is NOT an error — Excel
    computes such a schedule, discounting to the first date listed.

    Example::

        cf = mo.Variable([-100_000, 20_000, 30_000, 60_000])
        d  = mo.Variable([date(2025,1,1), date(2025,3,15), date(2025,7,10), date(2025,12,31)])
        project_xirr = mo.XIRR(cf, d)  # Excel: =XIRR(B2:E2, B3:E3, 0.1)
    """
    _one_guess(guess, "XIRR")
    cf_array = np.array(_zero_blanks(_extract_cashflow_values(cash_flows, "XIRR")), dtype=float)
    years = _years_from_first(dates, len(cf_array), "XIRR", earlier_allowed=True)
    if not (np.any(cf_array < 0) and np.any(cf_array > 0)):
        raise ValueError(
            "XIRR is undefined when all cash flows have the same sign — "
            "at least one negative (outflow) and one positive (inflow) "
            "value are required (Excel returns #NUM!)."
        )
    guess_value = float(val(guess))
    if guess_value <= -1:
        raise ValueError(
            f"XIRR guess must be above -100% — got {guess_value!r}; the "
            f"discount factor (1 + rate) must stay positive (Excel: #NUM!)."
        )

    try:
        rate = _irr_solve(
            cf_array, guess_value, times=years, grid=_XIRR_GRID,
            tol=_XIRR_NPV_RTOL * float(np.abs(cf_array).sum()),
        )
    except RuntimeError as exc:
        raise RuntimeError(
            "XIRR failed: the XNPV of these flows does not change sign for any "
            "annual rate from -99.99% to 1e9 (a hundred billion percent), so "
            "there is no rate to return (Excel returns #NUM!). Verify the "
            "cash flows and their dates."
        ) from exc

    flows, when = _range_arg(cash_flows), _range_arg(dates)
    if not isinstance(flows, Literal) and not isinstance(when, Literal):
        return _wrap_with_formula(rate, 'XIRR', [flows, when, _formula_arg(guess)])
    return Variable(float(rate), value_type='float')


def XNPV(
    rate: Union[float, Variable],
    cash_flows: Union[Variable, List[float]],
    dates: Union[Variable, List[_date_type]],
) -> Variable:
    """Net present value of dated flows — renders ``=XNPV(rate, values, dates)``.

    Excel's formula: ``Σ P_i / (1 + rate) ** ((d_i - d_1) / 365)`` — every
    flow discounted to the FIRST date listed, 365 days a year. Unlike
    :func:`NPV`, the first flow sits at time zero (it is not discounted).

    Raises ``ValueError`` where Excel returns ``#NUM!``: a rate of zero or
    below (Excel's XNPV rejects them although the sum exists — measured in
    Excel 16.0, build 20326), a date earlier than the first date, a date
    count that differs from the flow count. Flows of one sign are fine.

    The formula is live when both ``cash_flows`` and ``dates`` are
    Variables (their cells become the ranges); raw lists give a value-only
    result, as with :func:`IRR` / :func:`XIRR`.

    Example::

        r  = mo.Variable(0.09)
        cf = mo.Variable([-10_000, 2_750, 4_250, 3_250, 2_750])
        d  = mo.Variable([date(2008,1,1), date(2008,3,1), date(2008,10,30),
                          date(2009,2,15), date(2009,4,1)])
        value = mo.XNPV(r, cf, d)   # 2,086.65 — Excel: =XNPV(B1, B2:F2, B3:F3)
    """
    rate_value = val(rate)
    if isinstance(rate_value, (list, tuple)):
        raise TypeError(
            "XNPV takes a single discount rate — got a series. For a rate that "
            "changes over time, discount each flow explicitly."
        )
    rate_value = float(rate_value)
    if rate_value <= 0:
        raise ValueError(
            f"XNPV needs a positive rate — got {rate_value!r}. Excel's XNPV "
            f"returns #NUM! for a rate of zero or below, so the workbook would "
            f"show an error where the engine showed a number."
        )
    cf_array = np.array(_zero_blanks(_extract_cashflow_values(cash_flows, "XNPV")), dtype=float)
    years = _years_from_first(dates, len(cf_array), "XNPV", earlier_allowed=False)
    xnpv_value = float(np.sum(cf_array / (1 + rate_value) ** years))

    flows, when = _range_arg(cash_flows), _range_arg(dates)
    if not isinstance(flows, Literal) and not isinstance(when, Literal):
        return _wrap_with_formula(xnpv_value, 'XNPV', [_formula_arg(rate), flows, when])
    return Variable(xnpv_value, value_type='float')


# The payment, future value and present value of one period's numbers —
# the arithmetic of every call, over scalars or one period of a series,
# as the spreadsheet computes it: ``PMT`` is the payment the schedule
# functions split (``_payment``), ``FV`` the balance they carry
# (``_balance``), and ``PV`` discounts by ``(1 + r) ** -n``. A cell is
# read as the spreadsheet reads it (``_number``) and holds a double, a
# type other than 0 reads as 1 (``_advance``), and the result is
# ``#NUM!`` where the spreadsheet's is: ``PMT`` over a term of 0, at a
# rate below -100% (its power is taken through ``log1p``) or paid at the
# start of a period at -100%, ``PV`` at -100% over a positive term, a
# fractional power of a negative number, a power or a result too large
# for a cell.
def _finite_real(x: Any) -> Any:
    """``x`` as a cell holds it: ``#NUM!`` where it is no finite real
    number, 0 rather than -0."""
    if isinstance(x, complex) or not math.isfinite(x):
        return '#NUM!'
    return x + 0.0


def _pmt(rate: Any, nper: Any, pv: Any, fv: Any = 0, when: Any = 0) -> Any:
    r, n, v, f = (float(_number(x)) for x in (rate, nper, pv, fv))
    w = _advance(when)
    try:
        return _finite_real(_payment(r, n, v, f, w))
    except (ArithmeticError, ValueError):       # ValueError: no log1p below -100%
        return '#NUM!'


def _fv(rate: Any, nper: Any, pmt: Any, pv: Any = 0, when: Any = 0) -> Any:
    r, n, p, v = (float(_number(x)) for x in (rate, nper, pmt, pv))
    w = _advance(when)
    try:
        return _finite_real(_balance(r, n, p, v, w))
    except ArithmeticError:
        return '#NUM!'


def _pv(rate: Any, nper: Any, pmt: Any, fv: Any = 0, when: Any = 0) -> Any:
    r, n, p, f = (float(_number(x)) for x in (rate, nper, pmt, fv))
    w = _advance(when)
    if r == 0:
        return _finite_real(-(f + p * n))
    # Over ``(1 + r) ** -n``: a term too long for ``(1 + r) ** n`` to fit
    # a cell still has a present value.
    try:
        u = (1 + r) ** -n
        if w:
            return _finite_real(-(f * u + p * (1 - (1 + r) ** (1 - n)) / r + p))
        return _finite_real(-(f * u + p * (1 - u) / r))
    except ArithmeticError:
        return '#NUM!'


def _series_among(*args: Any) -> bool:
    return any(is_series(a) for a in args)


# Every call is one per period, as a loop container makes it: over scalars
# once, its text keeping the numbers it was given (``source_code``); with a
# series among the arguments once a period, every argument written as the
# call over scalars writes it, read as ``a + b`` reads its operands (a
# scalar in every period, lines on other windows by date, lines along keys
# key by key). An error an argument holds is that period's result.
def _each_period(func: str, kernel: Any, *args: Any,
                 source_code: str | None = None) -> Variable:
    return elementwise(func, kernel, list(args), blank=0, value_type='float',
                       render_backends=_EXCEL_ONLY, source_code=source_code)


def PMT(
    rate: Union[float, Variable],
    nper: Union[int, float, Variable],
    pv: Union[float, Variable],
    fv: Union[float, Variable] = 0,
    when: Union[int, Variable] = 0,
) -> Variable:
    """Loan/annuity payment per period (renders ``=PMT(rate, nper, pv, fv, when)``).

    Sign convention: payment flow is opposite to ``pv``. A positive loan
    amount ``pv`` yields a negative ``PMT`` (money out of your pocket).
    ``when=0`` → end-of-period (default, matches Excel). ``when=1`` → start-of-period.

    Example::

        # Monthly payment on a $300k 30-year loan at 6% annual
        monthly = mo.PMT(0.06 / 12, 30 * 12, 300_000)  # ≈ -1,798.65

    Any argument may vary by period: a series among them makes the call
    one per period, ``=PMT(C9 / 100, 12, 1000, 0, 0)`` in every column,
    and a loop container makes it one period at a time. The arguments
    read as the spreadsheet reads them — an error a cell holds is the
    result, a ``when`` other than 0 is 1 — and the result is ``#NUM!``
    where the spreadsheet's is: over a term of 0, at a rate below -100%,
    paid at the start of a period at -100%, or where ``(1 + rate) **
    nper`` is too large for a cell. Guard a term that can run out on the
    term itself, ``mo.IF(left > 0, -mo.PMT(rate, left, balance), 0)``:
    paid at the start of a period, the spreadsheet can show a
    meaningless number over a term of 0, which ``IFERROR`` lets through.
    """
    if _series_among(rate, nper, pv, fv, when):
        return _each_period('PMT', _pmt, rate, nper, pv, fv, when)
    r, n, pv_v, fv_v, w = _arg(rate), _arg(nper), _arg(pv), _arg(fv), _arg(when)
    return _each_period('PMT', _pmt, rate, nper, pv, fv, when,
                        source_code=f"PMT({r}, {n}, {pv_v}, {fv_v}, {w})")


def FV(
    rate: Union[float, Variable],
    nper: Union[int, float, Variable],
    pmt: Union[float, Variable],
    pv: Union[float, Variable] = 0,
    when: Union[int, Variable] = 0,
) -> Variable:
    """Future value of an annuity/loan (renders ``=FV(rate, nper, pmt, pv, when)``).

    Example::

        # $500/month into 7%-return account for 30 years, starting from $0
        retirement = mo.FV(0.07 / 12, 30 * 12, -500)  # ≈ 609,985.50

    Any argument may vary by period, one call per period as for
    :func:`PMT`; errors and ``when`` read as there, and the result is
    ``#NUM!`` where the spreadsheet's is.
    """
    if _series_among(rate, nper, pmt, pv, when):
        return _each_period('FV', _fv, rate, nper, pmt, pv, when)
    r, n, p, pv_v, w = _arg(rate), _arg(nper), _arg(pmt), _arg(pv), _arg(when)
    return _each_period('FV', _fv, rate, nper, pmt, pv, when,
                        source_code=f"FV({r}, {n}, {p}, {pv_v}, {w})")


def PV(
    rate: Union[float, Variable],
    nper: Union[int, float, Variable],
    pmt: Union[float, Variable],
    fv: Union[float, Variable] = 0,
    when: Union[int, Variable] = 0,
) -> Variable:
    """Present value of an annuity/loan (renders ``=PV(rate, nper, pmt, fv, when)``).

    Example::

        # PV of $1000/yr for 20 years at 5% discount
        pv = mo.PV(0.05, 20, -1000)  # ≈ 12,462

    Any argument may vary by period, one call per period as for
    :func:`PMT`; errors and ``when`` read as there, and the result is
    ``#NUM!`` where the spreadsheet's is.
    """
    if _series_among(rate, nper, pmt, fv, when):
        return _each_period('PV', _pv, rate, nper, pmt, fv, when)
    r, n, p, fv_v, w = _arg(rate), _arg(nper), _arg(pmt), _arg(fv), _arg(when)
    return _each_period('PV', _pv, rate, nper, pmt, fv, when,
                        source_code=f"PV({r}, {n}, {p}, {fv_v}, {w})")


# ─── Loan schedules and depreciation ──────────────────────────────
#
# ``NPER``, ``IPMT``, ``PPMT``, ``RATE``, ``CUMIPMT``, ``CUMPRINC`` and
# the depreciation functions ``SLN``, ``DB``, ``DDB``, ``SYD`` are each
# one call per period through the shared :func:`elementwise` helper,
# over a kernel that takes one period's numbers: scalars give one
# number, a period-shaped operand (``per=mo.time.index``) a schedule
# row. A kernel reads a cell as the spreadsheet does (``_number``) — a
# blank is 0, a boolean 1 or 0, numeric text its number, other text
# ``#VALUE!``, an error cell passes through — and a number outside the
# function's domain is ``#NUM!``. An optional argument is written only
# when given.

Numeric = Variable | int | float


def _number(v: Any) -> int | float:
    """A cell's number as the spreadsheet reads it, as a plain Python
    number: numpy's float warns and gives ``inf`` where Python raises
    the division by zero or the overflow the helper reads as an error."""
    x = _number_of(v)
    return x if type(x) is int else float(x)


def _whole(v: Any) -> int:
    """A period, a month or a type read whole: the number as a cell shows
    it (15 significant digits), its fraction dropped — 12.999999999999998
    is period 13, as the spreadsheet counts it."""
    x = _number(v)
    if not math.isfinite(x):
        raise OverflowError(f"not a finite number: {v!r}")   # the helper's #NUM!
    return int(_spreadsheet_round(x, 0, decimal.ROUND_DOWN))


def _advance(when: Any) -> int:
    """Excel's ``type``: 1 for a payment at the start of the period, 0
    at its end. Any other number reads as 1, as the spreadsheet reads it."""
    return 1 if _number(when) else 0


def _written(*args: Any) -> list[Any]:
    """The arguments as the cell writes them: an optional one that is
    omitted (None) is dropped from the end; one before a given one is
    its default, 0."""
    out = list(args)
    while out and out[-1] is None:
        out.pop()
    return [0 if a is None else a for a in out]


def _payment(r: Any, n: Any, v: Any, f: Any, w: int) -> Any:
    """The payment ``PMT`` computes — the kernel the schedule functions
    share. ``(1 + r) ** n - 1`` is ``expm1``, as the spreadsheet computes
    it: at a rate of 1e-12 the subtraction would keep four digits."""
    if r == 0:
        return -(v + f) / n
    if r == -1:
        # Every amount gone after a period; over no term, or a negative
        # one, no payment (a division by zero).
        t = 0.0 ** n
        t_less_one = t - 1.0
    else:
        growth = n * math.log1p(r)
        t, t_less_one = math.exp(growth), math.expm1(growth)
    return -(f + v * t) * r / ((1 + r * w) * t_less_one)


def _balance(r: Any, k: Any, payment: Any, v: Any, w: int) -> Any:
    """``FV`` after ``k`` payments: minus the balance the loan still carries."""
    if r == 0:
        return -(v + payment * k)
    t = (1 + r) ** k
    return -(v * t + payment * (1 + r * w) * (t - 1) / r)


def _nper(rate: Any, pmt: Any, pv: Any, fv: Any = 0, when: Any = 0) -> Any:
    r, p, v, f = (_number(x) for x in (rate, pmt, pv, fv))
    w = _advance(when)
    if v + f == 0:
        return 0.0                              # nothing to pay off
    if r == 0:
        return -(v + f) / p + 0.0 if p else '#NUM!'
    if r <= -1:
        return '#NUM!'
    z = p * (1 + r * w) / r
    if v + z == 0:
        return '#NUM!'                          # the payment only ever covers the interest
    ratio = (z - f) / (v + z)
    # ``+ 0.0`` throughout: -0.0 + 0.0 is 0.0, and a cell shows 0, not -0.
    return math.log(ratio) / math.log1p(r) + 0.0 if ratio > 0 else '#NUM!'


def _ipmt(rate: Any, per: Any, nper: Any, pv: Any, fv: Any = 0, when: Any = 0) -> Any:
    r, p, n, v, f = (_number(x) for x in (rate, per, nper, pv, fv))
    w = _advance(when)
    if p < 1 or p > n or r < -1:
        return '#NUM!'
    if r == 0 or (w and p == 1):
        return 0.0
    if p == 1:
        # The interest on the loan itself, the payment not needed: a term
        # too long for (1 + rate) ** nper still has a first period.
        return -v * r + 0.0
    interest = _balance(r, p - 1, _payment(r, n, v, f, w), v, w) * r
    return (interest / (1 + r) if w else interest) + 0.0


def _ppmt(rate: Any, per: Any, nper: Any, pv: Any, fv: Any = 0, when: Any = 0) -> Any:
    interest = _ipmt(rate, per, nper, pv, fv, when)
    if isinstance(interest, str):
        return interest
    r, n, v, f = (_number(x) for x in (rate, nper, pv, fv))
    return _payment(r, n, v, f, _advance(when)) - interest + 0.0


#: ``RATE`` iterates as LibreOffice's source does: Newton's method on the
#: balance from the guess, at most 150 steps, done when a step is under
#: 1e-7 or the balance under 1e-14. With no guess given, 0.1 is tried
#: and then, in turn, 0.2, 0.05, 0.3, 0.1/3, ... 1.0, 0.01, each where
#: the last found no rate. An end that is no rate (-100%, or a point
#: that does not balance the flows) is a start that found none: the
#: spreadsheet stops there and shows it.
_RATE_STEPS = 150
_RATE_STEP_DONE = 1e-7
_RATE_BALANCE_DONE = 1e-14
_RATE_GUESSES = (0.1,) + tuple(g for k in range(2, 11) for g in (0.1 * k, 0.1 / k))


def _rate_iteration(n: float, p: float, v: float, f: float, x: float) -> float | None:
    """The rate Newton's method reaches from ``x`` on the end-of-period
    balance ``f + v (1+x)^n + p ((1+x)^n - 1) / x``, or None: LibreOffice's
    arithmetic and stopping rules, step for step, so it settles where the
    spreadsheet settles and fails where it fails. A whole ``n`` may pass
    below -100% on the way (a whole power of a negative number is real);
    a fractional one may not."""
    whole = float(n).is_integer()
    if not whole:
        x = max(x, -1.0)
    try:
        for _ in range(_RATE_STEPS):
            if whole:
                t1 = (1.0 + x) ** (n - 1.0)
                t = t1 * (1.0 + x)
            else:
                t, t1 = (1.0 + x) ** n, (1.0 + x) ** (n - 1.0)
            if x == 0:
                series, slope = float(n), n * (n - 1.0) / 2.0
            else:
                series = (t - 1.0) / x
                slope = n * t1 / x - series / x
            balance = f + v * t + p * series
            if abs(balance) < _RATE_BALANCE_DONE:
                break
            d_balance = v * n * t1 + p * slope
            x_new = x - balance / d_balance if d_balance else x + 1.1 * _RATE_STEP_DONE
            done = abs(x_new - x) < _RATE_STEP_DONE
            x = x_new
            if not whole and not x >= -1.0:
                return None
            if done:
                break
        else:
            return None
    except (OverflowError, ZeroDivisionError):
        return None                             # the spreadsheet's inf / nan: no root
    return x if not whole or x > -1.0 else None


def _rate_gap(n: float, p: float, v: float, f: float, r: float) -> tuple[float, float, float]:
    """The balance ``f + v (1+r)^n + p ((1+r)^n - 1) / r`` a rate leaves
    (zero at the rate sought), its slope, and the size of its terms;
    infinite where they overflow. log1p / expm1 keep the series exact at
    a small rate, where the slope is read at its limit, n (n-1) / 2."""
    try:
        lg = math.log1p(r)
        t1, t2 = math.exp(n * lg), math.exp((n - 1) * lg)
        series = n if r == 0 else math.expm1(n * lg) / r
    except OverflowError:
        return math.inf, math.inf, math.inf
    slope = (n * t2 - series) / r if abs(r) > 1e-8 else n * (n - 1) / 2
    size = abs(f) + abs(v * t1) + abs(p * series)
    return f + v * t1 + p * series, v * n * t2 + p * slope, size


def _settled_rate(n: float, p: float, v: float, f: float, x: float,
                  at_minus_one: bool) -> float | None:
    """The rate the iteration ended on, polished on the exact series, or
    None where it is no rate: -100% (every amount gone — a root of the
    flows ``at_minus_one``), or a point that does not balance them (an
    iteration that stalled far out)."""
    if not 1.0 + x > (1e-5 if at_minus_one else 0.0):
        return None
    g, dg, size = _rate_gap(n, p, v, f, x)
    if not (math.isfinite(size) and abs(g) <= 1e-6 * size):
        return None
    # The spreadsheet stops within 1e-7; a few exact steps settle the
    # last digits (a rate of 0 is 0, not 1e-11), never past -100%.
    for _ in range(3):
        step = g / dg if dg else 0.0
        if not (abs(step) <= 1e-6 * (1 + abs(x)) and 1.0 + (x - step) > 0.0):
            break
        x -= step
        if abs(step) <= 1e-15 * (1 + abs(x)):
            break
        g, dg, size = _rate_gap(n, p, v, f, x)
    # A rate too small to move 1 + rate is the rate 0, as the cell shows it.
    return x + 0.0 if 1.0 + x != 1.0 else 0.0


def _rate(nper: Any, pmt: Any, pv: Any, fv: Any = 0, when: Any = 0, guess: Any = None) -> Any:
    n, p, v, f = (_number(x) for x in (nper, pmt, pv, fv))
    if n <= 0:
        return '#NUM!'
    w = _advance(when)
    if v + p * w == 0 and f + p * (1 - w) == 0 and (n == 1 or p == 0):
        # Every rate balances flows that are not there, or a single period
        # repaid as it opens: the iteration stops where it starts.
        start = 0.1 if guess is None else _number(guess)
        return start + 0.0 if start > -1 else '#NUM!'
    at_minus_one = f + p * (1 - w) == 0
    # A payment at the start of the period is one made at time zero: it
    # moves from the annuity into the present value and out of the
    # future value, and the iteration is the end-of-period one.
    v, f = v + p * w, f - p * w
    for start in _RATE_GUESSES if guess is None else (_number(guess),):
        x = _rate_iteration(n, p, v, f, start)
        rate = None if x is None else _settled_rate(n, p, v, f, x, at_minus_one)
        if rate is not None:
            return rate
    return '#NUM!'


def _cumulative(kernel: Any, rate: Any, nper: Any, pv: Any, start: Any, end: Any,
                when: Any) -> Any:
    r, n, v = (_number(x) for x in (rate, nper, pv))
    # The periods and the type are read whole, as Excel documents them.
    s, e, w = (_whole(x) for x in (start, end, when))
    if r <= 0 or n <= 0 or v <= 0 or s < 1 or e < s or e > n or w not in (0, 1):
        return '#NUM!'
    return sum(kernel(r, i, n, v, 0, w) for i in range(s, e + 1)) + 0.0


def _sln(cost: Any, salvage: Any, life: Any) -> Any:
    c, s, n = (_number(x) for x in (cost, salvage, life))
    return (c - s) / n + 0.0                    # a life of 0 is the helper's #DIV/0!


def _db(cost: Any, salvage: Any, life: Any, period: Any, month: Any = 12) -> Any:
    c, s, n, p = (_number(x) for x in (cost, salvage, life, period))
    m = _whole(month)
    if c <= 0 or s < 0 or s > c or n <= 0 or p < 1 or p > n + 1 or m < 1 or m > 12:
        return '#NUM!'
    # Excel's documented formula: the rate rounded to three decimals as
    # ROUND rounds it; the first period's charge over ``month`` months;
    # each next period's on the balance; and in a period past ``life``
    # the stub over the months the first year left. The period is read
    # whole: 1.5 is the first period, even past a life shorter than that.
    rate = _spreadsheet_round(1 - (s / c) ** (1 / n), 3)
    first = c * rate * m / 12
    if _whole(p) == 1:
        return first + 0.0
    charge, total = 0.0, first
    for _ in range(2, _whole(min(n, p)) + 1):
        charge = (c - total) * rate
        total += charge
    if p > n:
        charge = (c - total) * rate * (12 - m) / 12
    return charge + 0.0


def _ddb(cost: Any, salvage: Any, life: Any, period: Any, factor: Any = 2) -> Any:
    c, s, n, p, k = (_number(x) for x in (cost, salvage, life, period, factor))
    if c < 0 or s < 0 or s > c or k <= 0 or p < 1 or p > n:
        return '#NUM!'
    # The book value before and after the period at the declining rate,
    # the charge their difference, stopped at the salvage value.
    rate = k / n
    if rate >= 1:
        rate, before = 1.0, (c if p == 1 else 0.0)
    else:
        before = c * (1 - rate) ** (p - 1)
    after = c * (1 - rate) ** p
    charge = before - s if after < s else before - after
    return max(charge, 0.0) + 0.0


def _syd(cost: Any, salvage: Any, life: Any, per: Any) -> Any:
    c, s, n, p = (_number(x) for x in (cost, salvage, life, per))
    if n <= 0 or p < 1 or p > n:
        return '#NUM!'
    return (c - s) * (n - p + 1) * 2 / (n * (n + 1)) + 0.0


def NPER(
    rate: Numeric,
    pmt: Numeric,
    pv: Numeric,
    fv: Numeric | None = None,
    when: Numeric | None = None,
) -> Variable:
    """Number of periods of a loan or an annuity (renders
    ``=NPER(rate, pmt, pv, [fv], [type])``).

    ``NPER(0.01, -100, -1000, 10000)`` is 60.08; a rate of 0 gives
    ``-(pv + fv) / pmt``; ``when=1`` (Excel's ``type``) puts the payment
    at the start of each period. A present and a future value that
    cancel need 0 periods; a payment that never pays the balance down
    (``NPER(0.1, 100, -1000)``), a zero payment at a zero rate, or a
    rate of -100% or below is ``#NUM!``. The result carries no unit; the
    payment and the values are in one. Element-wise over a period-shaped
    operand.
    """
    return elementwise("NPER", _nper, _written(rate, pmt, pv, fv, when),
                       blank=0, value_type='float')


def IPMT(
    rate: Numeric,
    per: Numeric,
    nper: Numeric,
    pv: Numeric,
    fv: Numeric | None = None,
    when: Numeric | None = None,
) -> Variable:
    """Interest paid in period ``per`` (renders
    ``=IPMT(rate, per, nper, pv, [fv], [type])``).

    Signs as ``PMT``: a positive ``pv`` gives a negative payment.
    ``IPMT(0.1, 3, 3, 8000)`` is -292.45. A rate of 0 is 0 interest;
    with ``when=1`` the first period's interest is 0 and every later
    one is discounted a period. ``per`` below 1 or above ``nper``, or a
    rate below -100%, is ``#NUM!``. A series ``per`` — ``mo.time.index``
    — makes the call a schedule row::

        interest = mo.IPMT(inp.rate, mo.time.index, inp.term, inp.loan)
    """
    return elementwise("IPMT", _ipmt, _written(rate, per, nper, pv, fv, when),
                       blank=0, value_type='float')


def PPMT(
    rate: Numeric,
    per: Numeric,
    nper: Numeric,
    pv: Numeric,
    fv: Numeric | None = None,
    when: Numeric | None = None,
) -> Variable:
    """Principal paid in period ``per`` (renders
    ``=PPMT(rate, per, nper, pv, [fv], [type])``).

    The payment less :func:`IPMT`'s interest, with its rules:
    ``PPMT(0.08, 10, 10, 200000)`` is -27,598.05; at a rate of 0 the
    whole payment ``-(pv + fv) / nper``; ``per`` outside 1 .. ``nper``
    is ``#NUM!``. A series ``per`` makes the call a schedule row.
    """
    return elementwise("PPMT", _ppmt, _written(rate, per, nper, pv, fv, when),
                       blank=0, value_type='float')


def RATE(
    nper: Numeric,
    pmt: Numeric,
    pv: Numeric,
    fv: Numeric | None = None,
    when: Numeric | None = None,
    guess: Numeric | None = None,
) -> Variable:
    """The rate per period of a loan or an annuity (renders
    ``=RATE(nper, pmt, pv, [fv], [type], [guess])``).

    ``RATE(48, -200, 8000)`` is 0.77% a month. Found as LibreOffice
    finds it: Newton's iteration from ``guess``, at most 150 steps,
    settled when a step is under 1e-7; with no guess it starts at 0.1
    (Excel's default) and, should that find no rate, at 0.2, 0.05, 0.3,
    0.033, ... 1.0, 0.01 in turn, so of several roots it finds the
    spreadsheet's (``RATE(5, -1000, 1000)`` is 96.6%). The rate settled
    on is refined to full precision; payments that repay the loan exactly
    have a rate of 0. An end at -100% or at a point that does not
    balance the flows is no rate: the spreadsheet shows it, the engine
    tries the next start. ``#NUM!`` when ``nper`` is 0 or below, or when
    no start finds a rate (``RATE(10, 100, 1000)``: every flow the same
    way; a guess the iteration runs away from). With no flow at all
    every rate balances, and the guess is the answer. Excel documents 20
    steps, so a rate that takes more may be ``#NUM!`` in Excel. The
    result carries no unit. Element-wise over a period-shaped operand.
    """
    return elementwise("RATE", _rate, _written(nper, pmt, pv, fv, when, guess),
                       blank=0, value_type='float')


def CUMIPMT(
    rate: Numeric,
    nper: Numeric,
    pv: Numeric,
    start_period: Numeric,
    end_period: Numeric,
    when: Numeric = 0,
) -> Variable:
    """Interest paid between two periods, both included (renders
    ``=CUMIPMT(rate, nper, pv, start_period, end_period, type)``).

    The sum of :func:`IPMT` over the periods — ``CUMIPMT(0.09 / 12,
    360, 125000, 13, 24, 0)`` is -11,135.23, the second year's interest.
    The periods and ``when`` are read whole, as a cell shows them
    (12.999999999999998 is period 13) and as Excel documents it, their
    fractions dropped; LibreOffice refuses a fractional ``when``.
    ``when`` is written always, as Excel requires it. A rate, term or
    present value of 0 or below, a start before period 1 or after the
    end, an end after ``nper``, or a ``when`` other than 0 and 1 is
    ``#NUM!``.
    """
    return elementwise(
        "CUMIPMT", functools.partial(_cumulative, _ipmt),
        [rate, nper, pv, start_period, end_period, 0 if when is None else when],
        blank=0, value_type='float',
    )


def CUMPRINC(
    rate: Numeric,
    nper: Numeric,
    pv: Numeric,
    start_period: Numeric,
    end_period: Numeric,
    when: Numeric = 0,
) -> Variable:
    """Principal paid between two periods, both included (renders
    ``=CUMPRINC(rate, nper, pv, start_period, end_period, type)``).

    The sum of :func:`PPMT` over the periods, with :func:`CUMIPMT`'s
    rules — ``CUMPRINC(0.09 / 12, 360, 125000, 13, 24, 0)`` is -934.11.
    """
    return elementwise(
        "CUMPRINC", functools.partial(_cumulative, _ppmt),
        [rate, nper, pv, start_period, end_period, 0 if when is None else when],
        blank=0, value_type='float',
    )


def SLN(cost: Numeric, salvage: Numeric, life: Numeric) -> Variable:
    """Straight-line depreciation for one period (renders ``=SLN(cost, salvage, life)``).

    ``(cost - salvage) / life``: ``SLN(30000, 7500, 10)`` is 2,250. A
    life of 0 is ``#DIV/0!``; a negative cost or a salvage above the
    cost is computed as written. Element-wise over a period-shaped
    operand; the result keeps the unit the amounts share.
    """
    return elementwise("SLN", _sln, [cost, salvage, life], blank=0, value_type='float')


def DB(
    cost: Numeric,
    salvage: Numeric,
    life: Numeric,
    period: Numeric,
    month: Numeric | None = None,
) -> Variable:
    """Fixed-declining-balance depreciation for ``period`` (renders
    ``=DB(cost, salvage, life, period, [month])``).

    Excel's formula: the rate ``1 - (salvage / cost) ** (1 / life)``
    rounded to three decimals, the first period's charge over ``month``
    months (12 when omitted), each next period's on the remaining
    balance, and in a period past ``life`` the stub over the months the
    first year left. ``DB(1000000, 100000, 6, 1, 7)`` is 186,083.33.
    ``period`` and ``month`` are read whole, as a cell shows them: 1.5 is
    the first period, whatever the life. A cost of 0 or below, a salvage
    below 0 or above the cost, a life of 0 or below, a period below 1 or
    above ``life + 1``, or a month outside 1 .. 12 is ``#NUM!`` (a period
    between 0 and 1 is outside Excel's domain; LibreOffice computes it as
    0). A series ``period`` — ``mo.time.index`` — makes the call a
    schedule row.
    """
    return elementwise("DB", _db, _written(cost, salvage, life, period, month),
                       blank=0, value_type='float')


def DDB(
    cost: Numeric,
    salvage: Numeric,
    life: Numeric,
    period: Numeric,
    factor: Numeric | None = None,
) -> Variable:
    """Double-declining-balance depreciation for ``period`` (renders
    ``=DDB(cost, salvage, life, period, [factor])``).

    The charge at the rate ``factor / life`` (``factor`` 2 when omitted)
    on the book value, never past the salvage value:
    ``DDB(2400, 300, 10, 1)`` is 480, ``DDB(2400, 300, 10, 10)`` is
    22.12 (the balance stops at 300). A fractional period is charged
    along the declining curve, as the spreadsheet charges it. A cost or
    salvage below 0, a salvage above the cost, a factor of 0 or below,
    or a period below 1 or above ``life`` is ``#NUM!``. A series
    ``period`` makes the call a schedule row.
    """
    return elementwise("DDB", _ddb, _written(cost, salvage, life, period, factor),
                       blank=0, value_type='float')


def SYD(cost: Numeric, salvage: Numeric, life: Numeric, per: Numeric) -> Variable:
    """Sum-of-years'-digits depreciation for ``per`` (renders
    ``=SYD(cost, salvage, life, per)``).

    ``(cost - salvage) * (life - per + 1) * 2 / (life * (life + 1))``:
    ``SYD(30000, 7500, 10, 1)`` is 4,090.91, the tenth year 409.09. A
    life of 0 or below, or a period below 1 or above ``life``, is
    ``#NUM!`` — Excel's domain; LibreOffice computes the formula as
    written there (0 the period after the life, negative after that). A
    series ``per`` makes the call a schedule row.
    """
    return elementwise("SYD", _syd, [cost, salvage, life, per], blank=0, value_type='float')


# The cash-flow functions take their flows as a range; a rate or a date
# argument that is one value is written as its cell.
register("IRR", shape="range")
register("NPV", shape="range")
register("XIRR", shape="range")
register("XNPV", shape="range")
# PMT, FV and PV, the schedule and the depreciation functions are one
# call per period.
register("PMT", shape="elementwise")
register("FV", shape="elementwise")
register("PV", shape="elementwise")
register("NPER", shape="elementwise")
register("IPMT", shape="elementwise")
register("PPMT", shape="elementwise")
register("RATE", shape="elementwise")
register("CUMIPMT", shape="elementwise")
register("CUMPRINC", shape="elementwise")
register("SLN", shape="elementwise")
register("DB", shape="elementwise")
register("DDB", shape="elementwise")
register("SYD", shape="elementwise")


# ─── Units ──────────────────────────────────────────────────────────


def _amounts(func: str, first: int, stop: int, *, gives: bool = True) -> Any:
    """The law of a function reading amounts at ``first:stop`` of its
    written arguments: one unit among them, the result's — none for a
    rate or a count (``gives=False``)."""
    from ..core.expr import ListExpr

    def law(args: Any, unit_of: Any) -> Any:
        # A list of lines (``NPV(rate, [y1, y2])``) is read line by line.
        nodes = [n for a in args[first:stop]
                 for n in (a.items if isinstance(a, ListExpr) else (a,))]
        unit = same_unit(func, nodes, unit_of)
        return unit if gives else None
    return law


unit_law("IRR", _amounts("IRR", 0, 1, gives=False))     # (flows, guess)
unit_law("XIRR", _amounts("XIRR", 0, 1, gives=False))   # (flows, dates, guess)
unit_law("NPV", _amounts("NPV", 1, 2))              # (rate, flows)
unit_law("XNPV", _amounts("XNPV", 1, 2))            # (rate, flows, dates)
unit_law("PMT", _amounts("PMT", 2, 4))              # (rate, nper, pv, fv, type)
unit_law("FV", _amounts("FV", 2, 4))                # (rate, nper, pmt, pv, type)
unit_law("PV", _amounts("PV", 2, 4))                # (rate, nper, pmt, fv, type)
unit_law("NPER", _amounts("NPER", 1, 4, gives=False))   # (rate, pmt, pv, fv, type)
unit_law("RATE", _amounts("RATE", 1, 4, gives=False))   # (nper, pmt, pv, fv, type, guess)
unit_law("IPMT", _amounts("IPMT", 3, 5))            # (rate, per, nper, pv, fv, type)
unit_law("PPMT", _amounts("PPMT", 3, 5))
unit_law("CUMIPMT", _amounts("CUMIPMT", 2, 3))      # (rate, nper, pv, start, end, type)
unit_law("CUMPRINC", _amounts("CUMPRINC", 2, 3))
unit_law("SLN", _amounts("SLN", 0, 2))              # (cost, salvage, life)
unit_law("DB", _amounts("DB", 0, 2))                # (cost, salvage, life, period, month)
unit_law("DDB", _amounts("DDB", 0, 2))              # (cost, salvage, life, period, factor)
unit_law("SYD", _amounts("SYD", 0, 2))              # (cost, salvage, life, per)
