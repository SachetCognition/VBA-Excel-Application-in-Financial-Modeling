"""
Power BI Loan Amortization — Python companion for DAX calculated tables.

This script provides the loan amortization logic that mirrors
the DAX implementation in Power BI. It can be used:
1. As a Python Script data source in Power BI
2. For standalone validation against the original VBA
3. To generate reference data for DAX formula verification

DAX Approach Summary:
- What-If Parameters: LoanAmount, AnnualRate, LoanYears
- Disconnected tables: PaymentFrequency, PaymentType
- Calculated Table using GENERATESERIES + ADDCOLUMNS with
  closed-form balance formulas (no recursion needed)
"""

import numpy as np
import numpy_financial as npf
import pandas as pd


def dax_amortization_constant(loan, annual_rate_pct, years,
                                frequency="Monthly",
                                payment_type="End of Period"):
    """
    Closed-form constant payment amortization matching DAX approach.
    Uses POWER-based formulas (no recursion) for BegBal, Interest, etc.
    """
    rate = annual_rate_pct / 100.0
    ppy = 12 if frequency == "Monthly" else 1
    r = rate / ppy
    n = int(years * ppy)
    pmt_type = 0 if payment_type == "End of Period" else 1

    pmt = -npf.pmt(r, n, loan, 0, pmt_type)

    rows = []
    for i in range(1, n + 1):
        if pmt_type == 0:
            # End of period: closed-form
            beg_bal = loan * (pow(1 + r, n) - pow(1 + r, i - 1)) / (pow(1 + r, n) - 1)
            interest = beg_bal * r
            principal = pmt - interest
            end_bal = beg_bal - principal
        else:
            # Begin of period
            if i == 1:
                beg_bal = loan
                interest = 0
                principal = pmt
                end_bal = beg_bal - principal
            else:
                beg_bal = loan - pmt * (1 + (1 - pow(1 + r, -(n - (i - 1)))) / r - 1)
                # Iterative fallback for begin-of-period
                interest = beg_bal * r
                principal = pmt - interest
                end_bal = beg_bal - principal

        rows.append({
            "Period": i,
            "BegBal": round(beg_bal, 2),
            "Payment": round(pmt, 2),
            "Interest": round(interest, 2),
            "Principal": round(principal, 2),
            "EndBal": round(end_bal, 2),
        })

    return pd.DataFrame(rows)


def dax_amortization_straight_line(loan, annual_rate_pct, years,
                                    frequency="Monthly",
                                    payment_type="End of Period"):
    """Straight-line amortization with constant principal."""
    rate = annual_rate_pct / 100.0
    ppy = 12 if frequency == "Monthly" else 1
    r = rate / ppy
    n = int(years * ppy)
    const_princ = loan / n

    rows = []
    beg_bal = loan
    for i in range(1, n + 1):
        if payment_type == "Begin of Period" and i == 1:
            interest = 0
        else:
            interest = beg_bal * r
        payment = interest + const_princ
        end_bal = beg_bal - const_princ
        rows.append({
            "Period": i,
            "BegBal": round(beg_bal, 2),
            "Payment": round(payment, 2),
            "Interest": round(interest, 2),
            "Principal": round(const_princ, 2),
            "EndBal": round(end_bal, 2),
        })
        beg_bal = end_bal

    return pd.DataFrame(rows)


# What-If parameter defaults matching Power BI setup
LOAN_AMOUNT = 100000
ANNUAL_RATE = 5.0
LOAN_YEARS = 30
FREQUENCY = "Monthly"
PAYMENT_TYPE = "End of Period"

# Generate both schedules
constant_schedule = dax_amortization_constant(
    LOAN_AMOUNT, ANNUAL_RATE, LOAN_YEARS, FREQUENCY, PAYMENT_TYPE
)
straight_line_schedule = dax_amortization_straight_line(
    LOAN_AMOUNT, ANNUAL_RATE, LOAN_YEARS, FREQUENCY, PAYMENT_TYPE
)

if __name__ == "__main__":
    print("=== Constant Payment Schedule (first 5 periods) ===")
    print(constant_schedule.head().to_string(index=False))
    print(f"\nTotal periods: {len(constant_schedule)}")
    print(f"Monthly payment: ${constant_schedule['Payment'].iloc[0]:,.2f}")
    print(f"Total interest: ${constant_schedule['Interest'].sum():,.2f}")

    print("\n=== Straight-Line Schedule (first 5 periods) ===")
    print(straight_line_schedule.head().to_string(index=False))
    print(f"\nTotal periods: {len(straight_line_schedule)}")
    print(f"First payment: ${straight_line_schedule['Payment'].iloc[0]:,.2f}")
    print(f"Last payment: ${straight_line_schedule['Payment'].iloc[-1]:,.2f}")
    print(f"Total interest: ${straight_line_schedule['Interest'].sum():,.2f}")
