# SPDX-License-Identifier: Apache-2.0
"""
Financial functions: IRR and NPV for cash-flow analysis.

These are pure Python functions that compute on lists of numbers.
Useful for investment decisions, project ROI, etc.
"""

import modeleon as mo

# An investment: upfront cost then five years of returns
cash_flows = [-100_000, 25_000, 35_000, 40_000, 45_000, 50_000]

# Net present value at 10% discount rate
npv_at_10 = mo.NPV(0.10, cash_flows)

# Internal rate of return (where NPV = 0)
irr = mo.IRR(cash_flows)

print(f"Cash flows:  {cash_flows}")
print(f"NPV at 10%:  ${float(npv_at_10):,.2f}")
print(f"IRR:         {float(irr) * 100:.2f}%")
