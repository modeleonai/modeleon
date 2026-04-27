# SPDX-License-Identifier: Apache-2.0
"""Multi-sheet 12-month forecast.

Run this, open forecast.xlsx in Excel — it has two tabs. Edit any
input on Assumptions, the P&L tab recalculates across all 12 periods."""
from datetime import date

import modeleon as mo

forecast = mo.MultiVariable("Forecast")

# Each sub-MultiVariable becomes its own Excel tab automatically when
# no explicit `excel_props={'tab': True}` markers exist in the tree.
forecast.assumptions = mo.MultiVariable("Assumptions")

# `n_periods` is a real Variable — it appears as a cell on the
# Assumptions tab and is referenced in formulas that depend on the
# horizon (e.g. the EOMONTH offset below). List-valued Variables still
# bake their length in Python at build time, so we extract `.value`
# when constructing the growth list.
forecast.assumptions.n_periods   = mo.Variable(12, display_name="Periods")
N = forecast.assumptions.n_periods.value

forecast.assumptions.start_date  = mo.Variable(date(2025, 1, 1))
forecast.assumptions.start_rev   = mo.Variable(1_000_000, unit="$",
                                                display_name="Starting Revenue")
forecast.assumptions.growth      = mo.Variable([0.05] * N,
                                                display_name="Monthly Growth %")
forecast.assumptions.cogs_pct    = mo.Variable(0.6, display_name="COGS %")
forecast.assumptions.tax_rate    = mo.Variable(0.25, display_name="Tax Rate")

# Horizon end date — Excel cell: =EOMONTH(B2, B1 - 1). Both inputs
# (start_date and n_periods) are referenced live: change either and
# horizon_end recalculates in the workbook.
forecast.assumptions.horizon_end = mo.EOMONTH(
    forecast.assumptions.start_date,
    forecast.assumptions.n_periods - 1,
).set_display_name("Horizon End")

# `with mv as alias` is plain Python — `pnl is forecast.pnl`. Saves
# typing the path inside the block; attachment still goes through
# attribute assignment.
forecast.pnl = mo.MultiVariable("P&L")
with forecast.pnl as pnl:
    # Period header — one date per column, formula-driven. Period 0 =
    # start_date verbatim; periods 1..N-1 chain via =EDATE(prev, 1).
    # Change start_date in Excel and every header recalculates live.
    pnl.month = mo.recurrence(
        start=forecast.assumptions.start_date,
        formula="EDATE({prev}, 1)",
        periods=N,
        display_name="Month",
    )

    # `recurrence` expresses period-over-period: period 0 = start,
    # periods 1+ apply the formula with `{prev}` bound to the previous
    # period's cell. Change the starting revenue or growth rate on
    # Assumptions and every later cell updates.
    pnl.revenue = mo.recurrence(
        start=forecast.assumptions.start_rev,
        formula="{prev} * (1 + {growth})",
        growth=forecast.assumptions.growth,
        display_name="Revenue",
    )
    # Cross-sheet reference + `set_display_name` overrides the
    # auto-humanized "Cogs" label on a derived Variable.
    pnl.cogs  = (pnl.revenue * forecast.assumptions.cogs_pct).set_display_name("COGS")
    pnl.gross = pnl.revenue - pnl.cogs
    pnl.taxes = pnl.gross * forecast.assumptions.tax_rate
    pnl.net   = pnl.gross - pnl.taxes

forecast.to_excel("forecast.xlsx")
print("Wrote forecast.xlsx")
