"""Generate sample portfolio data for Power BI Python data source."""

import numpy as np
import pandas as pd

np.random.seed(42)

n_months = 61  # 61 prices -> 60 monthly returns
dates = pd.date_range(start="2019-01-01", periods=n_months, freq="MS")

# Simulate realistic monthly returns
asset1_monthly_ret = np.random.normal(0.008, 0.05, n_months)
asset2_monthly_ret = np.random.normal(0.005, 0.07, n_months)
rf_monthly_ret = np.random.normal(0.002, 0.005, n_months)

# Convert to price series starting at 100
asset1_prices = 100 * np.exp(np.cumsum(asset1_monthly_ret))
asset2_prices = 100 * np.exp(np.cumsum(asset2_monthly_ret))
rf_prices = 100 * np.exp(np.cumsum(rf_monthly_ret))

df = pd.DataFrame({
    "Date": dates,
    "Asset1": np.round(asset1_prices, 4),
    "Asset2": np.round(asset2_prices, 4),
    "RiskFree": np.round(rf_prices, 4),
})

df.to_csv("project/after/sample_portfolio_data.csv", index=False)
print(f"Generated {len(df)} rows of sample portfolio data")
print(df.head(10))
