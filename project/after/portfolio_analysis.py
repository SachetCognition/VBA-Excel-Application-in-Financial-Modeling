"""
Power BI Python Script Data Source — Portfolio Analysis.

This script replaces the VBA-based portfolio mean-variance analysis
with a pure Python implementation suitable for Power BI's Python
Script data source connector.

When run inside Power BI (Get Data → Python Script), the resulting
DataFrames are automatically imported as Power BI tables.

Can also be run standalone for validation and testing.
"""

import numpy as np
import pandas as pd
import os

# -------------------------------------------------------------------
# Read data — Power BI passes 'dataset' variable, or read from file
# -------------------------------------------------------------------
script_dir = os.path.dirname(os.path.abspath(__file__))
data_path = os.path.join(script_dir, "sample_portfolio_data.csv")

if "dataset" in dir():
    df = dataset  # noqa: F821 — Power BI injects this variable
else:
    df = pd.read_csv(data_path)

# -------------------------------------------------------------------
# Log returns (matching VBA: Log(Price(i) / Price(i-1)))
# -------------------------------------------------------------------
df["ret1"] = np.log(df["Asset1"] / df["Asset1"].shift(1))
df["ret2"] = np.log(df["Asset2"] / df["Asset2"].shift(1))
df["ret_rf"] = np.log(df["RiskFree"] / df["RiskFree"].shift(1))
df = df.dropna()

ret1 = df["ret1"].values
ret2 = df["ret2"].values
ret_rf = df["ret_rf"].values

# -------------------------------------------------------------------
# Annualized statistics (matching VBA lines 136-141)
# -------------------------------------------------------------------
mu = np.array([ret1.mean() * 12, ret2.mean() * 12])
rf = ret_rf.mean() * 12
std_devs = np.array([
    np.std(ret1, ddof=0) * np.sqrt(12),
    np.std(ret2, ddof=0) * np.sqrt(12),
])

# Annualized VCV matrix — ddof=0 matches Excel COVAR (VBA lines 154-157)
cov_matrix = np.cov(ret1, ret2, ddof=0) * 12

# Correlation matrix (VBA lines 144-147)
corr_matrix = np.corrcoef(ret1, ret2)

# -------------------------------------------------------------------
# Inverse covariance
# -------------------------------------------------------------------
inv_cov = np.linalg.inv(cov_matrix)
ones = np.ones(2)

# -------------------------------------------------------------------
# Minimum Variance Portfolio (matching VBA MMULT(MINVERSE(...)) formula)
# -------------------------------------------------------------------
w_mvp = inv_cov @ ones / (ones @ inv_cov @ ones)
mvp_ret = w_mvp @ mu
mvp_std = np.sqrt(w_mvp @ cov_matrix @ w_mvp)

# -------------------------------------------------------------------
# Optimal Risky Portfolio (tangency — matching VBA line 152)
# -------------------------------------------------------------------
excess = mu - rf * ones
w_orp = inv_cov @ excess / (ones @ inv_cov @ excess)
orp_ret = w_orp @ mu
orp_std = np.sqrt(w_orp @ cov_matrix @ w_orp)
orp_sharpe = (orp_ret - rf) / orp_std

# -------------------------------------------------------------------
# Efficient frontier (44 points as in VBA rows 23-66)
# -------------------------------------------------------------------
results = []
for w1 in np.linspace(-0.5, 1.5, 44):
    w_vec = np.array([w1, 1 - w1])
    ret = w_vec @ mu
    std = np.sqrt(w_vec @ cov_matrix @ w_vec)
    var = w_vec @ cov_matrix @ w_vec
    sharpe = (ret - rf) / std if std > 0 else 0
    results.append({
        "w_asset1": round(w1, 4),
        "w_asset2": round(1 - w1, 4),
        "Return": ret,
        "StdDev": std,
        "Variance": var,
        "Sharpe": sharpe,
    })

frontier = pd.DataFrame(results)

# -------------------------------------------------------------------
# Summary table for Power BI cards
# -------------------------------------------------------------------
summary = pd.DataFrame([
    {"Metric": "Asset 1 Expected Return", "Value": mu[0]},
    {"Metric": "Asset 2 Expected Return", "Value": mu[1]},
    {"Metric": "Asset 1 Std Dev", "Value": std_devs[0]},
    {"Metric": "Asset 2 Std Dev", "Value": std_devs[1]},
    {"Metric": "Risk-Free Rate", "Value": rf},
    {"Metric": "Correlation", "Value": corr_matrix[0, 1]},
    {"Metric": "MVP Weight Asset 1", "Value": w_mvp[0]},
    {"Metric": "MVP Weight Asset 2", "Value": w_mvp[1]},
    {"Metric": "MVP Return", "Value": mvp_ret},
    {"Metric": "MVP Std Dev", "Value": mvp_std},
    {"Metric": "ORP Weight Asset 1", "Value": w_orp[0]},
    {"Metric": "ORP Weight Asset 2", "Value": w_orp[1]},
    {"Metric": "ORP Return", "Value": orp_ret},
    {"Metric": "ORP Std Dev", "Value": orp_std},
    {"Metric": "ORP Sharpe Ratio", "Value": orp_sharpe},
])

# -------------------------------------------------------------------
# VCV matrix table for Power BI
# -------------------------------------------------------------------
vcv_table = pd.DataFrame(
    cov_matrix,
    index=["Asset1", "Asset2"],
    columns=["Asset1", "Asset2"],
).reset_index().rename(columns={"index": "Asset"})

# -------------------------------------------------------------------
# Print results when running standalone
# -------------------------------------------------------------------
if __name__ == "__main__":
    print("=" * 60)
    print("PORTFOLIO MEAN-VARIANCE ANALYSIS — Power BI Python Source")
    print("=" * 60)
    print(f"\nAsset 1 — E[R]: {mu[0]:.4%}, Std: {std_devs[0]:.4%}")
    print(f"Asset 2 — E[R]: {mu[1]:.4%}, Std: {std_devs[1]:.4%}")
    print(f"Risk-Free Rate: {rf:.4%}")
    print(f"\nCorrelation Matrix:\n{corr_matrix}")
    print(f"\nVCV Matrix (annualized):\n{cov_matrix}")
    print(f"\nMVP Weights: [{w_mvp[0]:.4f}, {w_mvp[1]:.4f}]")
    print(f"MVP Return: {mvp_ret:.4%}, Std: {mvp_std:.4%}")
    print(f"\nORP Weights: [{w_orp[0]:.4f}, {w_orp[1]:.4f}]")
    print(f"ORP Return: {orp_ret:.4%}, Std: {orp_std:.4%}")
    print(f"ORP Sharpe Ratio: {orp_sharpe:.4f}")
    print(f"\nEfficient Frontier ({len(frontier)} points):")
    print(frontier.to_string(index=False))
