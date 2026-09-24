# SPDX-License-Identifier: Apache-2.0
"""Financial formula helpers — IRR, NPV, XIRR, PMT, FV, PV.

Every function builds a live Excel formula (``=IRR(B2:B6)``, ``=PMT(...)``)
alongside the Python-side value. Pass a Variable (including list-valued
ones) to get the formula; pass a raw list to get a value-only result —
the latter is useful for ad-hoc calculations outside the model tree.
"""

from __future__ import annotations

import logging
from datetime import datetime, date as _date_type
from typing import List, Union

import numpy as np

from ..core.expr import FuncCall, Literal, VarRef
from ..core.variable import Variable
from ._helpers import val

logger = logging.getLogger(__name__)


__all__ = ['IRR', 'NPV', 'XIRR', 'PMT', 'FV', 'PV']


# ─── Shared helpers ───────────────────────────────────────────────


def _extract_cashflow_values(cash_flows) -> List[float]:
    """Pull numeric cash-flow values from a Variable or plain list."""
    if isinstance(cash_flows, Variable):
        values = cash_flows._value
        return list(values) if isinstance(values, list) else [values]
    if isinstance(cash_flows, list):
        return list(cash_flows)
    raise TypeError(
        f"cash_flows must be a Variable or list — "
        f"got {type(cash_flows).__name__}. Example:\n"
        f"    cf = mo.Variable([-100_000, 30_000, 40_000, 50_000])\n"
        f"    mo.IRR(cf)"
    )


_EXCEL_ONLY = frozenset({'excel'})
"""Render-backend set for Excel-native financial functions — IRR / NPV /
XIRR / PMT / FV / PV. Excel emits these natively; other renderers fall
back to inlining ``_value``."""


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


def _irr_newton(cash_flows: np.ndarray, guess: float, max_iter: int, tol: float) -> float:
    """Newton-Raphson root-find for the rate where NPV = 0.

    Raises ``RuntimeError`` when the derivative vanishes or iteration
    fails to converge — callers (IRR / XIRR) translate these into
    user-facing errors.
    """
    rate = float(guess)
    for _ in range(max_iter):
        periods = np.arange(len(cash_flows))
        npv = np.sum(cash_flows / ((1 + rate) ** periods))
        dnpv = np.sum(-periods * cash_flows / ((1 + rate) ** (periods + 1)))
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
    cf_values = _extract_cashflow_values(cash_flows)
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

    irr_value = _irr_newton(cf_array, guess, max_iter=100, tol=1e-6)

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
    rate_value = val(rate)
    cf_values = _extract_cashflow_values(cash_flows)
    npv_value = _npv(rate_value, np.array(cf_values, dtype=float))

    if isinstance(rate, Variable) or isinstance(cash_flows, Variable):
        return _wrap_with_formula(
            npv_value, 'NPV',
            [_formula_arg(rate), _formula_arg(cash_flows)],
        )
    return Variable(npv_value, value_type='float')


def XIRR(
    cash_flows: Union[Variable, List[float]],
    dates: Union[Variable, List[_date_type]],
    guess: float = 0.1,
) -> Variable:
    """Extended IRR — renders ``=XIRR(cash_flows_range, dates_range, guess)``.

    Unlike IRR, XIRR honors actual calendar days between cash flows and
    returns an annualized rate.

    Example::

        cf = mo.Variable([-100_000, 20_000, 30_000, 60_000])
        d  = mo.Variable([date(2025,1,1), date(2025,3,15), date(2025,7,10), date(2025,12,31)])
        project_xirr = mo.XIRR(cf, d)  # Excel: =XIRR(B2:E2, B3:E3, 0.1)
    """
    cf_values = _extract_cashflow_values(cash_flows)
    date_values = dates._value if hasattr(dates, '_value') else dates
    if not isinstance(date_values, list):
        raise TypeError(
            f"dates must be a Variable with list value or a list, got "
            f"{type(dates).__name__}."
        )

    date_objects = []
    for d in date_values:
        if isinstance(d, str):
            date_objects.append(datetime.strptime(d, '%Y-%m-%d').date())
        elif isinstance(d, _date_type):
            date_objects.append(d)
        else:
            raise ValueError(f"Invalid date in XIRR dates: {d!r}")

    first = date_objects[0]
    years = np.array(
        [(d - first).days / 365.25 for d in date_objects], dtype=float,
    )
    cf_array = np.array(cf_values, dtype=float)

    rate = float(guess)
    for _ in range(100):
        npv = float(np.sum(cf_array / ((1 + rate) ** years)))
        dnpv = float(np.sum(-years * cf_array / ((1 + rate) ** (years + 1))))
        if abs(npv) < 1e-6:
            break
        if abs(dnpv) < 1e-10:
            raise RuntimeError(
                "XIRR failed: derivative near zero. Try a different guess "
                "or verify the cash flows have a well-defined XIRR."
            )
        rate -= npv / dnpv
        if rate < -0.99:
            rate = -0.99
    else:
        raise RuntimeError(
            "XIRR failed to converge within the iteration limit. Try a "
            "different guess closer to the expected rate."
        )

    if isinstance(cash_flows, Variable) and isinstance(dates, Variable):
        return _wrap_with_formula(
            rate, 'XIRR',
            [VarRef(cash_flows), VarRef(dates), Literal(guess)],
        )
    return Variable(rate, value_type='float')


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
    r, n, pv_v, fv_v, w = val(rate), val(nper), val(pv), val(fv), val(when)
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
    r, n, p, pv_v, w = val(rate), val(nper), val(pmt), val(pv), val(when)
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
    r, n, p, fv_v, w = val(rate), val(nper), val(pmt), val(fv), val(when)
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
