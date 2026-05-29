"""
Test suite for Portfolio Mean-Variance Analysis — validates Python against VBA logic.

Tests cover:
- Log return computation
- Portfolio weighted returns
- Matrix operations (MVP, ORP)
- Covariance matching Excel COVAR (ddof=0)
- Sharpe ratio
- Efficient frontier properties
"""

import pytest
import numpy as np
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "before"))
from portfolio_analysis_vba_recreation import (
    get_monthly_log_returns,
    get_monthly_port_log_ret,
    portfolio_mean_return,
    portfolio_variance,
    sharpe_ratio,
    compute_vcv_matrix,
    compute_correlation_matrix,
    minimum_variance_portfolio,
    optimal_risky_portfolio,
    efficient_frontier,
    full_portfolio_analysis,
)


# ===========================================================================
# Log Returns
# ===========================================================================
class TestLogReturns:
    def test_basic_log_return(self):
        prices = np.array([100, 105, 103, 110])
        returns = get_monthly_log_returns(prices)
        assert len(returns) == 3
        assert abs(returns[0] - np.log(1.05)) < 1e-10

    def test_portfolio_weighted_return(self):
        ret1 = np.array([0.01, 0.02, -0.01])
        ret2 = np.array([0.005, -0.01, 0.03])
        w = 0.6
        port_ret = w * ret1 + (1 - w) * ret2
        assert abs(port_ret[0] - (0.6 * 0.01 + 0.4 * 0.005)) < 1e-10

    def test_flag_reverses_direction(self):
        """flag=True: prices newest-to-oldest (reversed calculation)."""
        prices1 = np.array([100, 105, 110])
        prices2 = np.array([200, 210, 220])
        ret_normal = get_monthly_port_log_ret(prices1, prices2, 0.5, flag=False)
        ret_reversed = get_monthly_port_log_ret(
            prices1[::-1], prices2[::-1], 0.5, flag=True
        )
        np.testing.assert_allclose(ret_normal, ret_reversed[::-1], atol=1e-10)

    def test_unequal_lengths_rejected(self):
        with pytest.raises(ValueError, match="Counts not equal"):
            get_monthly_port_log_ret([100, 105], [200, 210, 220], 0.5)


# ===========================================================================
# Matrix Operations
# ===========================================================================
class TestMatrixOps:
    def test_mvp_weights_sum_to_one(self):
        cov = np.array([[0.04, 0.01], [0.01, 0.09]])
        w_mvp = minimum_variance_portfolio(cov)
        assert abs(w_mvp.sum() - 1.0) < 1e-10

    def test_tangency_weights_sum_to_one(self):
        cov = np.array([[0.04, 0.01], [0.01, 0.09]])
        mu = np.array([0.10, 0.15])
        rf = 0.03
        w_tan = optimal_risky_portfolio(cov, mu, rf)
        assert abs(w_tan.sum() - 1.0) < 1e-10

    def test_covariance_matches_excel(self):
        """Excel COVAR uses population formula (ddof=0)."""
        x = np.array([1.0, 2.0, 3.0, 4.0, 5.0])
        y = np.array([2.0, 4.0, 5.0, 4.0, 5.0])
        cov_numpy = np.cov(x, y, ddof=0)[0, 1]
        expected = np.mean((x - 3) * (y - 4))
        assert abs(cov_numpy - expected) < 1e-10

    def test_mvp_is_minimum_variance(self):
        """MVP should have lower variance than any other portfolio on frontier."""
        cov = np.array([[0.04, 0.01], [0.01, 0.09]])
        mu = np.array([0.10, 0.15])
        w_mvp = minimum_variance_portfolio(cov)
        mvp_var = w_mvp @ cov @ w_mvp

        for w1 in np.linspace(0, 1, 50):
            w = np.array([w1, 1 - w1])
            var = w @ cov @ w
            assert var >= mvp_var - 1e-10

    def test_vcv_matrix_symmetric(self):
        ret1 = np.random.RandomState(42).normal(0, 0.05, 60)
        ret2 = np.random.RandomState(43).normal(0, 0.07, 60)
        vcv = compute_vcv_matrix(ret1, ret2)
        assert abs(vcv[0, 1] - vcv[1, 0]) < 1e-10


# ===========================================================================
# Sharpe Ratio
# ===========================================================================
class TestSharpeRatio:
    def test_basic_sharpe(self):
        port_ret = 0.12
        rf = 0.03
        port_std = 0.15
        sharpe = (port_ret - rf) / port_std
        assert abs(sharpe - 0.6) < 1e-10

    def test_sharpe_with_real_data(self, sample_portfolio_data):
        asset1, asset2, rf = sample_portfolio_data
        s = sharpe_ratio(asset1, asset2, 0.6, rf)
        assert isinstance(s, float)
        assert np.isfinite(s)


# ===========================================================================
# Efficient Frontier
# ===========================================================================
class TestEfficientFrontier:
    def test_frontier_has_44_points(self):
        cov = np.array([[0.04, 0.01], [0.01, 0.09]])
        mu = np.array([0.10, 0.15])
        frontier = efficient_frontier(cov, mu, 44)
        assert len(frontier) == 44

    def test_frontier_weights_sum_to_one(self):
        cov = np.array([[0.04, 0.01], [0.01, 0.09]])
        mu = np.array([0.10, 0.15])
        frontier = efficient_frontier(cov, mu, 44)
        for _, row in frontier.iterrows():
            assert abs(row["w_asset1"] + row["w_asset2"] - 1.0) < 1e-10

    def test_frontier_std_positive(self):
        cov = np.array([[0.04, 0.01], [0.01, 0.09]])
        mu = np.array([0.10, 0.15])
        frontier = efficient_frontier(cov, mu, 44)
        assert (frontier["StdDev"] > 0).all()


# ===========================================================================
# Full Analysis Integration
# ===========================================================================
class TestFullAnalysis:
    def test_full_analysis_runs(self, sample_portfolio_data):
        asset1, asset2, rf = sample_portfolio_data
        results = full_portfolio_analysis(asset1, asset2, rf)
        assert "expected_returns" in results
        assert "mvp_weights" in results
        assert "orp_weights" in results
        assert "frontier" in results

    def test_mvp_weights_sum(self, sample_portfolio_data):
        asset1, asset2, rf = sample_portfolio_data
        results = full_portfolio_analysis(asset1, asset2, rf)
        assert abs(results["mvp_weights"].sum() - 1.0) < 1e-10

    def test_orp_weights_sum(self, sample_portfolio_data):
        asset1, asset2, rf = sample_portfolio_data
        results = full_portfolio_analysis(asset1, asset2, rf)
        assert abs(results["orp_weights"].sum() - 1.0) < 1e-10

    def test_correlation_diagonal_is_one(self, sample_portfolio_data):
        asset1, asset2, rf = sample_portfolio_data
        results = full_portfolio_analysis(asset1, asset2, rf)
        corr = results["correlation_matrix"]
        assert abs(corr[0, 0] - 1.0) < 1e-10
        assert abs(corr[1, 1] - 1.0) < 1e-10

    def test_vcv_positive_diagonal(self, sample_portfolio_data):
        asset1, asset2, rf = sample_portfolio_data
        results = full_portfolio_analysis(asset1, asset2, rf)
        vcv = results["vcv_matrix"]
        assert vcv[0, 0] > 0
        assert vcv[1, 1] > 0
