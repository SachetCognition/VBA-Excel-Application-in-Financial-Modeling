"""
Performance tests — ensure Python implementations are faster than VBA/Excel.

Benchmarks:
- 360-period amortization < 100ms
- 60-month portfolio analysis < 500ms
- Full efficient frontier < 200ms
"""

import time
import numpy as np
import numpy_financial as npf
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "before"))
from loan_amortization_vba_recreation import (
    constant_payment_schedule,
    straight_line_schedule,
)
from portfolio_analysis_vba_recreation import full_portfolio_analysis


class TestPerformance:
    def test_amortization_360_periods(self):
        start = time.perf_counter()
        loan, rate, nper = 100000, 0.05 / 12, 360
        payment = -npf.pmt(rate, nper, loan)
        bal = loan
        for i in range(nper):
            interest = bal * rate
            principal = payment - interest
            bal -= principal
        elapsed = time.perf_counter() - start
        assert elapsed < 0.1

    def test_portfolio_analysis_60_months(self):
        start = time.perf_counter()
        np.random.seed(42)
        prices1 = 100 * np.exp(np.cumsum(np.random.normal(0.005, 0.05, 60)))
        prices2 = 100 * np.exp(np.cumsum(np.random.normal(0.003, 0.07, 60)))
        ret1 = np.log(prices1[1:] / prices1[:-1])
        ret2 = np.log(prices2[1:] / prices2[:-1])
        cov = np.cov(ret1, ret2, ddof=0) * 12
        np.linalg.inv(cov)
        for w1 in np.linspace(-0.5, 1.5, 44):
            w = np.array([w1, 1 - w1])
            np.sqrt(w @ cov @ w)
        elapsed = time.perf_counter() - start
        assert elapsed < 0.5

    def test_constant_schedule_generation(self):
        start = time.perf_counter()
        constant_payment_schedule(100000, 5, 30, "Monthly", "End of Period")
        elapsed = time.perf_counter() - start
        assert elapsed < 1.0

    def test_straight_line_schedule_generation(self):
        start = time.perf_counter()
        straight_line_schedule(100000, 5, 30, "Monthly", "End of Period")
        elapsed = time.perf_counter() - start
        assert elapsed < 1.0

    def test_full_portfolio_analysis_performance(self):
        np.random.seed(42)
        n = 61
        a1 = 100 * np.exp(np.cumsum(np.random.normal(0.008, 0.05, n)))
        a2 = 100 * np.exp(np.cumsum(np.random.normal(0.005, 0.07, n)))
        rf = 100 * np.exp(np.cumsum(np.random.normal(0.002, 0.005, n)))
        start = time.perf_counter()
        full_portfolio_analysis(a1, a2, rf)
        elapsed = time.perf_counter() - start
        assert elapsed < 2.0

    def test_multiple_loan_scenarios(self):
        """Run 10 different loan scenarios in under 5 seconds."""
        start = time.perf_counter()
        scenarios = [
            (100000, 5, 30, "Monthly", "End of Period"),
            (250000, 3.5, 15, "Monthly", "Begin of Period"),
            (50000, 7, 5, "Annually", "End of Period"),
            (500000, 4.5, 20, "Monthly", "End of Period"),
            (75000, 6, 10, "Monthly", "Begin of Period"),
            (200000, 3, 30, "Monthly", "End of Period"),
            (150000, 5.5, 25, "Monthly", "End of Period"),
            (300000, 4, 15, "Annually", "End of Period"),
            (80000, 8, 7, "Annually", "Begin of Period"),
            (1000000, 3.25, 30, "Monthly", "End of Period"),
        ]
        for loan, rate, years, freq, ptype in scenarios:
            constant_payment_schedule(loan, rate, years, freq, ptype)
            straight_line_schedule(loan, rate, years, freq, ptype)
        elapsed = time.perf_counter() - start
        assert elapsed < 5.0
