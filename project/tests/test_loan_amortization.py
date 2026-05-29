"""
Test suite for Loan Amortization — validates Python recreation against VBA logic.

Test Cases:
  Case 1: $100,000 / 5% / 30yr / Monthly / End of Period
  Case 2: $250,000 / 3.5% / 15yr / Monthly / Begin of Period
  Case 3: $50,000 / 7% / 5yr / Annual / End of Period
"""

import pytest
import numpy as np
import numpy_financial as npf
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "before"))
from loan_amortization_vba_recreation import (
    constant_payment_schedule,
    straight_line_schedule,
    summary_statistics,
    validate_inputs,
)


# ---------------------------------------------------------------------------
# Helper: reference PMT calculation
# ---------------------------------------------------------------------------
def pmt_constant(rate, nper, pv, pmt_type=0):
    return -npf.pmt(rate, nper, pv, 0, pmt_type)


def amortization_schedule(loan, rate, nper, pmt_type=0):
    payment = pmt_constant(rate, nper, loan, pmt_type)
    schedule = []
    balance = loan
    for i in range(1, nper + 1):
        if pmt_type == 1 and i == 1:
            interest = 0
        else:
            interest = balance * rate
        principal = payment - interest
        end_bal = balance - principal
        schedule.append({
            "period": i,
            "beg_bal": balance,
            "interest": interest,
            "principal": principal,
            "end_bal": end_bal,
            "payment": payment,
        })
        balance = end_bal
    return schedule


# ===========================================================================
# Constant Payment Tests
# ===========================================================================
class TestConstantPayment:
    def test_monthly_end_100k_5pct_30yr(self):
        sched = amortization_schedule(100000, 0.05 / 12, 360, 0)
        assert abs(sched[-1]["end_bal"]) < 0.01
        assert abs(sum(r["principal"] for r in sched) - 100000) < 0.01

    def test_monthly_begin_250k_35pct_15yr(self):
        sched = amortization_schedule(250000, 0.035 / 12, 180, 1)
        assert sched[0]["interest"] == 0
        assert abs(sched[-1]["end_bal"]) < 0.01

    def test_annual_end_50k_7pct_5yr(self):
        sched = amortization_schedule(50000, 0.07, 5, 0)
        assert abs(sched[-1]["end_bal"]) < 0.01
        for row in sched:
            assert abs(row["interest"] + row["principal"] - row["payment"]) < 0.001

    def test_payment_matches_excel_pmt(self):
        expected = 536.82  # PMT(0.05/12, 360, -100000)
        actual = pmt_constant(0.05 / 12, 360, 100000, 0)
        assert abs(actual - expected) < 0.01

    def test_vba_recreation_matches_reference(self):
        """Verify our VBA recreation module matches the reference calculation."""
        df = constant_payment_schedule(100000, 5, 30, "Monthly", "End of Period")
        ref = amortization_schedule(100000, 0.05 / 12, 360, 0)
        assert abs(df["EndBal"].iloc[-1]) < 0.01
        assert abs(df["Payment"].iloc[0] - ref[0]["payment"]) < 0.01
        assert abs(df["Interest"].iloc[0] - ref[0]["interest"]) < 0.01

    def test_begin_of_period_zero_first_interest(self):
        """VBA line 178: first period interest is 0 for begin-of-period."""
        df = constant_payment_schedule(100000, 5, 30, "Monthly", "Begin of Period")
        assert df["Interest"].iloc[0] == 0

    def test_total_principal_equals_loan(self):
        df = constant_payment_schedule(100000, 5, 30, "Monthly", "End of Period")
        assert abs(df["Principal"].sum() - 100000) < 0.01

    def test_annual_frequency(self):
        df = constant_payment_schedule(50000, 7, 5, "Annually", "End of Period")
        assert len(df) == 5
        assert abs(df["EndBal"].iloc[-1]) < 0.01


# ===========================================================================
# Straight-Line Tests
# ===========================================================================
class TestStraightLine:
    def test_constant_principal(self):
        df = straight_line_schedule(100000, 5, 30, "Monthly", "End of Period")
        principal_values = df["Principal"].values
        assert np.allclose(principal_values, principal_values[0], atol=0.001)

    def test_decreasing_total_payment(self):
        df = straight_line_schedule(100000, 5, 30, "Monthly", "End of Period")
        payments = df["Payment"].values
        for i in range(len(payments) - 1):
            assert payments[i] >= payments[i + 1] - 0.001

    def test_final_balance_zero(self):
        df = straight_line_schedule(100000, 5, 30, "Monthly", "End of Period")
        assert abs(df["EndBal"].iloc[-1]) < 0.01

    def test_straight_line_annual(self):
        df = straight_line_schedule(50000, 7, 5, "Annually", "End of Period")
        assert len(df) == 5
        expected_princ = 50000 / 5
        assert abs(df["Principal"].iloc[0] - expected_princ) < 0.01
        assert abs(df["EndBal"].iloc[-1]) < 0.01

    def test_begin_period_straight_line(self):
        df = straight_line_schedule(100000, 5, 30, "Monthly", "Begin of Period")
        assert df["Interest"].iloc[0] == 0


# ===========================================================================
# Input Validation Tests
# ===========================================================================
class TestInputValidation:
    def test_negative_loan_rejected(self):
        with pytest.raises(ValueError, match="positive"):
            validate_inputs(-100000, 5, 30)

    def test_zero_loan_rejected(self):
        with pytest.raises(ValueError, match="positive"):
            validate_inputs(0, 5, 30)

    def test_none_input_rejected(self):
        with pytest.raises(ValueError, match="no input"):
            validate_inputs(None, 5, 30)

    def test_non_numeric_rate_rejected(self):
        with pytest.raises(ValueError, match="number"):
            validate_inputs(100000, "abc", 30)

    def test_annual_non_integer_years_rejected(self):
        with pytest.raises(ValueError):
            validate_inputs(100000, 5, 2.5, "Annually")


# ===========================================================================
# Summary Statistics Tests
# ===========================================================================
class TestSummaryStatistics:
    def test_summary_constant_payment(self):
        df = constant_payment_schedule(100000, 5, 30, "Monthly", "End of Period")
        stats = summary_statistics(df)
        assert stats["total_periods"] == 360
        assert abs(stats["total_principal"] - 100000) < 0.01
        assert stats["total_interest"] > 0
        assert abs(stats["total_payments"] - stats["total_interest"] - 100000) < 0.01

    def test_summary_straight_line(self):
        df = straight_line_schedule(100000, 5, 30, "Monthly", "End of Period")
        stats = summary_statistics(df)
        assert stats["total_periods"] == 360
        assert abs(stats["total_principal"] - 100000) < 0.01


# ===========================================================================
# Cross-Validation: Constant vs Straight-Line
# ===========================================================================
class TestCrossValidation:
    def test_same_total_principal(self):
        """Both methods should repay the same total principal."""
        df_const = constant_payment_schedule(100000, 5, 30, "Monthly", "End of Period")
        df_sl = straight_line_schedule(100000, 5, 30, "Monthly", "End of Period")
        assert abs(df_const["Principal"].sum() - df_sl["Principal"].sum()) < 0.01

    def test_straight_line_less_total_interest(self):
        """Straight-line should have less total interest than constant payment."""
        df_const = constant_payment_schedule(100000, 5, 30, "Monthly", "End of Period")
        df_sl = straight_line_schedule(100000, 5, 30, "Monthly", "End of Period")
        assert df_sl["Interest"].sum() < df_const["Interest"].sum()
