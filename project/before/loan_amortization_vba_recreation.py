"""
Python recreation of the VBA Loan Amortization logic from LoanAmortization.xlsm.

This module faithfully replicates the VBA calculations including:
- Constant Payment Amortization (Part 1 of VBA code)
- Straight-Line Amortization (Part 2 of VBA code)
- Input validation matching VBA robust checks
- Both "End of Period" and "Begin of Period" payment types
- Monthly and Annual payment frequencies

Reference: Loan amortization VBA code.txt (555 lines)
"""

import numpy_financial as npf
import pandas as pd
import numpy as np


def validate_inputs(principal, annual_rate, loan_life_years, frequency="Monthly"):
    """Replicate VBA input robust checks."""
    if principal is None or annual_rate is None or loan_life_years is None:
        raise ValueError("Sorry, there is no input in the userform")
    if not isinstance(principal, (int, float)) or principal <= 0:
        raise ValueError("please enter a positive number for principal")
    if not isinstance(annual_rate, (int, float)):
        raise ValueError("Please enter a number of anual interest")
    if not isinstance(loan_life_years, (int, float)) or loan_life_years <= 0:
        raise ValueError(
            "please enter positive number of years for monthly payment "
            "or positive integer for annual payment"
        )
    if frequency == "Annually":
        if loan_life_years != int(loan_life_years):
            raise ValueError(
                "please enter positive number of years for monthly payment "
                "or positive integer for annual payment"
            )


def constant_payment_schedule(principal, annual_rate_pct, loan_life_years,
                               frequency="Monthly", payment_type="End of Period"):
    """
    Constant Payment Amortization — mirrors VBA Part 1.

    Parameters
    ----------
    principal : float
        Loan amount (loanAmnt in VBA).
    annual_rate_pct : float
        Annual interest rate as percentage (e.g. 5 for 5%).
    loan_life_years : float
        Loan life in years.
    frequency : str
        "Monthly" or "Annually".
    payment_type : str
        "End of Period" or "Begin of Period".

    Returns
    -------
    pd.DataFrame with columns: Period, BegBal, Payment, Interest, Principal, EndBal,
                                InterestPct, PrincipalPct
    """
    validate_inputs(principal, annual_rate_pct, loan_life_years, frequency)

    if frequency == "Monthly":
        rate = annual_rate_pct / 1200.0
        n_periods = int(loan_life_years * 12)
    else:
        rate = annual_rate_pct / 100.0
        n_periods = int(loan_life_years)

    flag = 0 if payment_type == "End of Period" else 1
    pmt = -npf.pmt(rate, n_periods, principal, 0, flag)

    rows = []
    beg_bal = principal

    if payment_type == "End of Period":
        for i in range(1, n_periods + 1):
            interest = beg_bal * rate
            princ = pmt - interest
            end_bal = beg_bal - princ
            rows.append({
                "Period": i,
                "BegBal": beg_bal,
                "Payment": pmt,
                "Interest": interest,
                "Principal": princ,
                "EndBal": end_bal,
                "InterestPct": interest / pmt if pmt != 0 else 0,
                "PrincipalPct": princ / pmt if pmt != 0 else 0,
            })
            beg_bal = end_bal
    else:
        # Begin of Period: first period has zero interest (matching VBA line 178)
        interest = 0.0
        princ = pmt - interest
        end_bal = beg_bal - princ
        rows.append({
            "Period": 1,
            "BegBal": beg_bal,
            "Payment": pmt,
            "Interest": interest,
            "Principal": princ,
            "EndBal": end_bal,
            "InterestPct": interest / pmt if pmt != 0 else 0,
            "PrincipalPct": princ / pmt if pmt != 0 else 0,
        })
        beg_bal = end_bal

        for i in range(2, n_periods + 1):
            interest = beg_bal * rate
            princ = pmt - interest
            end_bal = beg_bal - princ
            rows.append({
                "Period": i,
                "BegBal": beg_bal,
                "Payment": pmt,
                "Interest": interest,
                "Principal": princ,
                "EndBal": end_bal,
                "InterestPct": interest / pmt if pmt != 0 else 0,
                "PrincipalPct": princ / pmt if pmt != 0 else 0,
            })
            beg_bal = end_bal

    return pd.DataFrame(rows)


def straight_line_schedule(principal, annual_rate_pct, loan_life_years,
                           frequency="Monthly", payment_type="End of Period"):
    """
    Straight-Line Amortization — mirrors VBA Part 2.

    Constant principal each period; interest decreases as balance drops.
    """
    validate_inputs(principal, annual_rate_pct, loan_life_years, frequency)

    if frequency == "Monthly":
        rate = annual_rate_pct / 1200.0
        n_periods = int(loan_life_years * 12)
    else:
        rate = annual_rate_pct / 100.0
        n_periods = int(loan_life_years)

    const_princ = principal / n_periods
    rows = []
    beg_bal = principal

    if payment_type == "End of Period":
        for i in range(1, n_periods + 1):
            interest = beg_bal * rate
            payment = interest + const_princ
            end_bal = beg_bal - const_princ
            rows.append({
                "Period": i,
                "BegBal": beg_bal,
                "Payment": payment,
                "Interest": interest,
                "Principal": const_princ,
                "EndBal": end_bal,
                "InterestPct": interest / payment if payment != 0 else 0,
                "PrincipalPct": const_princ / payment if payment != 0 else 0,
            })
            beg_bal = end_bal
    else:
        # Begin of Period: first period zero interest (matching VBA line 229)
        interest = 0.0
        payment = interest + const_princ
        end_bal = beg_bal - const_princ
        rows.append({
            "Period": 1,
            "BegBal": beg_bal,
            "Payment": payment,
            "Interest": interest,
            "Principal": const_princ,
            "EndBal": end_bal,
            "InterestPct": 0.0,
            "PrincipalPct": 1.0,
        })
        beg_bal = end_bal

        for i in range(2, n_periods + 1):
            interest = beg_bal * rate
            payment = interest + const_princ
            end_bal = beg_bal - const_princ
            rows.append({
                "Period": i,
                "BegBal": beg_bal,
                "Payment": payment,
                "Interest": interest,
                "Principal": const_princ,
                "EndBal": end_bal,
                "InterestPct": interest / payment if payment != 0 else 0,
                "PrincipalPct": const_princ / payment if payment != 0 else 0,
            })
            beg_bal = end_bal

    return pd.DataFrame(rows)


def summary_statistics(schedule_df):
    """Compute summary stats matching VBA output cells F4-F7, H6-H7."""
    return {
        "total_periods": len(schedule_df),
        "total_interest": schedule_df["Interest"].sum(),
        "total_payments": schedule_df["Payment"].sum(),
        "total_principal": schedule_df["Principal"].sum(),
    }


if __name__ == "__main__":
    # Test Case 1: $100,000 / 5% / 30yr / Monthly / End of Period
    df = constant_payment_schedule(100000, 5, 30, "Monthly", "End of Period")
    stats = summary_statistics(df)
    print("=== Constant Payment: $100k / 5% / 30yr / Monthly / End ===")
    print(f"Payment: ${df['Payment'].iloc[0]:,.2f}")
    print(f"Total Interest: ${stats['total_interest']:,.2f}")
    print(f"Total Payments: ${stats['total_payments']:,.2f}")
    print(f"Final Balance: ${df['EndBal'].iloc[-1]:.4f}")
    print()

    # Straight-line for same parameters
    df_sl = straight_line_schedule(100000, 5, 30, "Monthly", "End of Period")
    stats_sl = summary_statistics(df_sl)
    print("=== Straight-Line: $100k / 5% / 30yr / Monthly / End ===")
    print(f"First Payment: ${df_sl['Payment'].iloc[0]:,.2f}")
    print(f"Last Payment: ${df_sl['Payment'].iloc[-1]:,.2f}")
    print(f"Total Interest: ${stats_sl['total_interest']:,.2f}")
