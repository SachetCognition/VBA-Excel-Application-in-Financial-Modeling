"""
Python recreation of the VBA Portfolio Mean-Variance Analysis logic.

This module faithfully replicates the VBA calculations including:
- Monthly log return computation (getMonthlyPortLogRet UDF)
- Portfolio mean return (portfolio_mean_return UDF, annualized)
- Portfolio variance (portfolio_variance UDF, annualized)
- Sharpe ratio calculation (sharpe_ratio UDF)
- Correlation matrix and VCV matrix
- Minimum Variance Portfolio (MVP) via MMULT(MINVERSE(...))
- Optimal Risky Portfolio (ORP / tangency portfolio)
- Efficient frontier (44 data points matching VBA loop rows 23-66)

Reference: Portfolio mean variance analysis VBA code.txt (256 lines)
"""

import numpy as np
import pandas as pd


def get_monthly_log_returns(prices):
    """Compute monthly log returns: Log(P(i+1)/P(i)). Matches VBA line 23."""
    prices = np.asarray(prices, dtype=float)
    return np.log(prices[1:] / prices[:-1])


def get_monthly_port_log_ret(asset1_prices, asset2_prices, w1, flag=False):
    """
    Replicate VBA getMonthlyPortLogRet UDF.

    flag=False: prices sorted oldest to newest (normal)
    flag=True: prices sorted newest to oldest (reversed)
    """
    a1 = np.asarray(asset1_prices, dtype=float)
    a2 = np.asarray(asset2_prices, dtype=float)

    if len(a1) != len(a2):
        raise ValueError("Counts not equal")

    if flag:
        ret1 = np.log(a1[:-1] / a1[1:])
        ret2 = np.log(a2[:-1] / a2[1:])
    else:
        ret1 = np.log(a1[1:] / a1[:-1])
        ret2 = np.log(a2[1:] / a2[:-1])

    return w1 * ret1 + (1 - w1) * ret2


def portfolio_mean_return(asset1_prices, asset2_prices, w1, flag=False):
    """Annualized mean portfolio return. Matches VBA line 41."""
    port_ret = get_monthly_port_log_ret(asset1_prices, asset2_prices, w1, flag)
    return np.mean(port_ret) * 12


def portfolio_variance(asset1_prices, asset2_prices, w1, flag=False):
    """Annualized portfolio variance (population). Matches VBA Var_P * 12."""
    port_ret = get_monthly_port_log_ret(asset1_prices, asset2_prices, w1, flag)
    return np.var(port_ret, ddof=0) * 12


def sharpe_ratio(asset1_prices, asset2_prices, w1, rf_prices, flag=False):
    """
    Sharpe ratio matching VBA sharpe_ratio UDF.
    Uses annualized rf from log returns of risk-free prices.
    portfolio_variance() already returns annualized variance (var * 12),
    so the annualized std is simply sqrt(annualized_var).
    """
    if flag:
        rf_ret = np.log(np.asarray(rf_prices[:-1], dtype=float) /
                        np.asarray(rf_prices[1:], dtype=float))
    else:
        rf_ret = np.log(np.asarray(rf_prices[1:], dtype=float) /
                        np.asarray(rf_prices[:-1], dtype=float))

    rf = np.mean(rf_ret) * 12
    port_mean = portfolio_mean_return(asset1_prices, asset2_prices, w1, flag)
    port_var = portfolio_variance(asset1_prices, asset2_prices, w1, flag)
    port_std = np.sqrt(port_var)

    return (port_mean - rf) / port_std


def compute_vcv_matrix(ret1, ret2):
    """
    Compute annualized Variance-Covariance matrix.
    Uses population covariance (ddof=0) matching Excel COVAR function.
    Matches VBA lines 154-157.
    """
    cov_matrix = np.cov(ret1, ret2, ddof=0) * 12
    return cov_matrix


def compute_correlation_matrix(ret1, ret2):
    """Compute correlation matrix matching VBA lines 144-147."""
    return np.corrcoef(ret1, ret2)


def minimum_variance_portfolio(vcv_matrix):
    """
    Minimum Variance Portfolio weights.
    Matches VBA formula:
    MMULT(MINVERSE(VCV), ones) / MMULT(TRANSPOSE(ones), MMULT(MINVERSE(VCV), ones))
    """
    inv_cov = np.linalg.inv(vcv_matrix)
    ones = np.ones(vcv_matrix.shape[0])
    w_mvp = inv_cov @ ones / (ones @ inv_cov @ ones)
    return w_mvp


def optimal_risky_portfolio(vcv_matrix, expected_returns, rf_rate):
    """
    Optimal Risky Portfolio (tangency portfolio) weights.
    Matches VBA formula:
    MMULT(MINVERSE(VCV), (mu - rf*ones)) /
    MMULT(TRANSPOSE(ones), MMULT(MINVERSE(VCV), (mu - rf*ones)))
    """
    inv_cov = np.linalg.inv(vcv_matrix)
    ones = np.ones(vcv_matrix.shape[0])
    excess = expected_returns - rf_rate * ones
    w_orp = inv_cov @ excess / (ones @ inv_cov @ excess)
    return w_orp


def efficient_frontier(vcv_matrix, expected_returns, n_points=44):
    """
    Generate efficient frontier data matching VBA loop rows 23-66 (44 points).
    Weight of asset 1 ranges from -0.5 to 1.5.
    """
    results = []
    for w1 in np.linspace(-0.5, 1.5, n_points):
        w = np.array([w1, 1 - w1])
        port_ret = w @ expected_returns
        port_std = np.sqrt(w @ vcv_matrix @ w)
        results.append({
            "w_asset1": w1,
            "w_asset2": 1 - w1,
            "Return": port_ret,
            "StdDev": port_std,
            "Variance": w @ vcv_matrix @ w,
        })
    return pd.DataFrame(results)


def full_portfolio_analysis(asset1_prices, asset2_prices, rf_prices, flag=False):
    """
    Run the complete portfolio analysis matching VBA Sub Portfolio_Analysis().
    """
    if flag:
        ret1 = np.log(np.asarray(asset1_prices[:-1], dtype=float) /
                       np.asarray(asset1_prices[1:], dtype=float))
        ret2 = np.log(np.asarray(asset2_prices[:-1], dtype=float) /
                       np.asarray(asset2_prices[1:], dtype=float))
        ret_rf = np.log(np.asarray(rf_prices[:-1], dtype=float) /
                        np.asarray(rf_prices[1:], dtype=float))
    else:
        ret1 = get_monthly_log_returns(asset1_prices)
        ret2 = get_monthly_log_returns(asset2_prices)
        ret_rf = get_monthly_log_returns(rf_prices)

    # Annualized expected returns (VBA lines 136-138)
    mu1 = np.mean(ret1) * 12
    mu2 = np.mean(ret2) * 12
    rf_annual = np.mean(ret_rf) * 12
    mu = np.array([mu1, mu2])

    # Annualized standard deviations (VBA lines 140-141)
    std1 = np.std(ret1, ddof=0) * np.sqrt(12)
    std2 = np.std(ret2, ddof=0) * np.sqrt(12)

    # Correlation matrix (VBA lines 144-147)
    corr = compute_correlation_matrix(ret1, ret2)

    # VCV matrix (VBA lines 154-157)
    vcv = compute_vcv_matrix(ret1, ret2)

    # MVP weights (VBA line 150)
    w_mvp = minimum_variance_portfolio(vcv)

    # ORP weights (VBA line 152)
    w_orp = optimal_risky_portfolio(vcv, mu, rf_annual)

    # Portfolio stats for MVP and ORP
    mvp_ret = w_mvp @ mu
    mvp_std = np.sqrt(w_mvp @ vcv @ w_mvp)
    orp_ret = w_orp @ mu
    orp_std = np.sqrt(w_orp @ vcv @ w_orp)
    orp_sharpe = (orp_ret - rf_annual) / orp_std

    # Efficient frontier (44 points, VBA rows 23-66)
    frontier = efficient_frontier(vcv, mu, 44)

    return {
        "expected_returns": mu,
        "std_devs": np.array([std1, std2]),
        "rf_rate": rf_annual,
        "correlation_matrix": corr,
        "vcv_matrix": vcv,
        "mvp_weights": w_mvp,
        "orp_weights": w_orp,
        "mvp_return": mvp_ret,
        "mvp_std": mvp_std,
        "orp_return": orp_ret,
        "orp_std": orp_std,
        "orp_sharpe": orp_sharpe,
        "frontier": frontier,
    }


if __name__ == "__main__":
    # Generate synthetic 60-month data for demonstration
    np.random.seed(42)
    n_months = 61  # 61 prices -> 60 returns
    asset1 = 100 * np.exp(np.cumsum(np.random.normal(0.008, 0.05, n_months)))
    asset2 = 100 * np.exp(np.cumsum(np.random.normal(0.005, 0.07, n_months)))
    rf = 100 * np.exp(np.cumsum(np.random.normal(0.002, 0.005, n_months)))

    results = full_portfolio_analysis(asset1, asset2, rf)

    print("=== Portfolio Analysis Results ===")
    print(f"Asset 1 - E[R]: {results['expected_returns'][0]:.4%}, "
          f"Std: {results['std_devs'][0]:.4%}")
    print(f"Asset 2 - E[R]: {results['expected_returns'][1]:.4%}, "
          f"Std: {results['std_devs'][1]:.4%}")
    print(f"Risk-Free Rate: {results['rf_rate']:.4%}")
    print(f"\nMVP Weights: [{results['mvp_weights'][0]:.4f}, "
          f"{results['mvp_weights'][1]:.4f}]")
    print(f"MVP Return: {results['mvp_return']:.4%}, "
          f"Std: {results['mvp_std']:.4%}")
    print(f"\nORP Weights: [{results['orp_weights'][0]:.4f}, "
          f"{results['orp_weights'][1]:.4f}]")
    print(f"ORP Return: {results['orp_return']:.4%}, "
          f"Std: {results['orp_std']:.4%}")
    print(f"ORP Sharpe: {results['orp_sharpe']:.4f}")
