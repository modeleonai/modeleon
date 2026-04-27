# SPDX-License-Identifier: Apache-2.0
"""Single-sheet P&L. Run this, open pnl.xlsx in Excel — every formula is live."""
import modeleon as mo

pnl = mo.MultiVariable("P&L")
pnl.revenue  = mo.Variable(1_000_000)
pnl.cogs_pct = mo.Variable(0.6)

pnl.cogs    = pnl.revenue * pnl.cogs_pct
pnl.profit  = pnl.revenue - pnl.cogs

pnl.to_excel("pnl.xlsx")
print("Wrote pnl.xlsx")
