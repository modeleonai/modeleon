# SPDX-License-Identifier: Apache-2.0
"""Financial formula helpers — IRR, NPV, XIRR, XNPV, PMT, FV, PV.

Every function builds a live Excel formula (``=IRR(B2:B6)``, ``=PMT(...)``)
alongside the Python-side value. Pass a Variable (including list-valued
ones) to get the formula; pass a raw list to get a value-only result —
the latter is useful for ad-hoc calculations outside the model tree.
"""

from __future__ import annotations

import logging
from datetime import date as _date_type
from typing import Any, List, Optional, Union

import numpy as np

from ..core.expr import FuncCall, Literal, VarRef
from ..core.variable import Variable
from ._helpers import val
from .dates import _extract_date

logger = logging.getLogger(__name__)


__all__ = ['IRR', 'NPV', 'XIRR', 'XNPV', 'PMT', 'FV', 'PV']


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
           "== 1, -cost, flow) - then mo.IRR(flows).",
    "NPV": "A time-0 outlay stays outside the call, as Excel discounts the "
           "first flow a whole period: outlay + mo.NPV(rate, flows); a later "
           "amount (a terminal value) goes into the row in its own period.",
    "XIRR": "Give the outlay a period of its own - with a window that starts "
            "at the close, the first: flows = mo.IF(mo.time.index == 1, -cost, "
            "flow), dates = mo.IF(mo.time.index == 1, close, mo.time.end) - "
            "then mo.XIRR(flows, dates).",
    "XNPV": "Give the outlay a period of its own - with a window that starts "
            "at the close, the first: flows = mo.IF(mo.time.index == 1, -cost, "
            "flow), dates = mo.IF(mo.time.index == 1, close, mo.time.end) - "
            "then mo.XNPV(rate, flows, dates).",
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
    """Build a scalar-valued result Variable with an Excel FuncCall formula."""
    result = Variable(value_type='float')
    result._set_expr(FuncCall(
        func_name, args,
        render_backends=_EXCEL_ONLY,
    ))
    result._value = float(value)
    result.var_type = 'scalar'
    result.value_type = 'float'
    return result


def _formula_arg(value) -> "Literal | VarRef":
    """AST node for a function argument — VarRef if it's a Variable, else Literal."""
    return VarRef(value) if isinstance(value, Variable) else Literal(value)


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

    if isinstance(cash_flows, Variable):
        return _wrap_with_formula(
            irr_value, 'IRR',
            [VarRef(cash_flows), Literal(guess)],
        )
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
    cf_values = _present(_extract_cashflow_values(cash_flows, "NPV"))
    npv_value = _npv(rate_value, np.array(cf_values, dtype=float))

    if isinstance(rate, Variable) or isinstance(cash_flows, Variable):
        return _wrap_with_formula(
            npv_value, 'NPV',
            [_formula_arg(rate), _formula_arg(cash_flows)],
        )
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

    if isinstance(cash_flows, Variable) and isinstance(dates, Variable):
        return _wrap_with_formula(
            rate, 'XIRR',
            [VarRef(cash_flows), VarRef(dates), _formula_arg(guess)],
        )
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

    if isinstance(cash_flows, Variable) and isinstance(dates, Variable):
        return _wrap_with_formula(
            xnpv_value, 'XNPV',
            [_formula_arg(rate), VarRef(cash_flows), VarRef(dates)],
        )
    return Variable(xnpv_value, value_type='float')


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
    """
    r, n, pv_v, fv_v, w = _arg(rate), _arg(nper), _arg(pv), _arg(fv), _arg(when)
    if r == 0:
        pmt_val = -(pv_v + fv_v) / n
    else:
        pmt_val = -(fv_v + pv_v * (1 + r) ** n) * r / ((1 + r * w) * ((1 + r) ** n - 1))

    result = _wrap_with_formula(
        pmt_val, 'PMT',
        [_formula_arg(a) for a in (rate, nper, pv, fv, when)],
    )
    result._source_code = f"PMT({r}, {n}, {pv_v}, {fv_v}, {w})"
    return result


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
        retirement = mo.FV(0.07 / 12, 30 * 12, -500)  # ≈ 611,729
    """
    r, n, p, pv_v, w = _arg(rate), _arg(nper), _arg(pmt), _arg(pv), _arg(when)
    if r == 0:
        fv_val = -(pv_v + p * n)
    else:
        fv_val = -(pv_v * (1 + r) ** n + p * (1 + r * w) * ((1 + r) ** n - 1) / r)

    result = _wrap_with_formula(
        fv_val, 'FV',
        [_formula_arg(a) for a in (rate, nper, pmt, pv, when)],
    )
    result._source_code = f"FV({r}, {n}, {p}, {pv_v}, {w})"
    return result


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
    """
    r, n, p, fv_v, w = _arg(rate), _arg(nper), _arg(pmt), _arg(fv), _arg(when)
    if r == 0:
        pv_val = -(fv_v + p * n)
    else:
        pv_val = -(fv_v + p * (1 + r * w) * ((1 + r) ** n - 1) / r) / (1 + r) ** n

    result = _wrap_with_formula(
        pv_val, 'PV',
        [_formula_arg(a) for a in (rate, nper, pmt, fv, when)],
    )
    result._source_code = f"PV({r}, {n}, {p}, {fv_v}, {w})"
    return result
