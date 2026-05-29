"""
Generate visualization screenshots for before/after comparison.

Creates matplotlib charts that replicate the original VBA Excel charts
and the Power BI report visuals.
"""

import sys
import os
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "before"))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "after"))

from loan_amortization_vba_recreation import (
    constant_payment_schedule,
    straight_line_schedule,
)
from portfolio_analysis_vba_recreation import full_portfolio_analysis

BEFORE_DIR = os.path.join(os.path.dirname(__file__), "screenshots", "before")
AFTER_DIR = os.path.join(os.path.dirname(__file__), "screenshots", "after")
os.makedirs(BEFORE_DIR, exist_ok=True)
os.makedirs(AFTER_DIR, exist_ok=True)


def currency_fmt(x, pos):
    return f"${x:,.0f}"


# ===========================================================================
# LOAN AMORTIZATION CHARTS
# ===========================================================================

def generate_loan_charts(loan, rate, years, freq, ptype, case_name):
    """Generate 14 charts (7 constant + 7 straight-line) for a loan case."""
    df_const = constant_payment_schedule(loan, rate, years, freq, ptype)
    df_sl = straight_line_schedule(loan, rate, years, freq, ptype)

    for label, df, prefix in [("Constant Payment", df_const, "const"),
                               ("Straight-Line", df_sl, "sl")]:
        periods = df["Period"].values

        # 1. BegBalance
        fig, ax = plt.subplots(figsize=(10, 5))
        ax.plot(periods, df["BegBal"], color="#2196F3", linewidth=1.5)
        ax.set_title(f"Beginning Balance — {label}", fontname="Times New Roman", fontsize=13)
        ax.set_xlabel("Periods", fontname="Times New Roman", fontsize=12)
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))
        ax.grid(False)
        fig.tight_layout()
        fig.savefig(os.path.join(BEFORE_DIR, f"{case_name}_{prefix}_begbal.png"), dpi=150)
        plt.close(fig)

        # 2. Payment
        fig, ax = plt.subplots(figsize=(10, 5))
        ax.plot(periods, df["Payment"], color="#4CAF50", linewidth=1.5)
        ax.set_title(f"Payment — {label}", fontname="Times New Roman", fontsize=13)
        ax.set_xlabel("Periods", fontname="Times New Roman", fontsize=12)
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))
        ax.grid(False)
        fig.tight_layout()
        fig.savefig(os.path.join(BEFORE_DIR, f"{case_name}_{prefix}_payment.png"), dpi=150)
        plt.close(fig)

        # 3. Interest
        fig, ax = plt.subplots(figsize=(10, 5))
        ax.plot(periods, df["Interest"], color="#FF9800", linewidth=1.5)
        ax.set_title(f"Interest — {label}", fontname="Times New Roman", fontsize=13)
        ax.set_xlabel("Periods", fontname="Times New Roman", fontsize=12)
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))
        ax.grid(False)
        fig.tight_layout()
        fig.savefig(os.path.join(BEFORE_DIR, f"{case_name}_{prefix}_interest.png"), dpi=150)
        plt.close(fig)

        # 4. Principal
        fig, ax = plt.subplots(figsize=(10, 5))
        ax.plot(periods, df["Principal"], color="#9C27B0", linewidth=1.5)
        ax.set_title(f"Principal — {label}", fontname="Times New Roman", fontsize=13)
        ax.set_xlabel("Periods", fontname="Times New Roman", fontsize=12)
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))
        ax.grid(False)
        fig.tight_layout()
        fig.savefig(os.path.join(BEFORE_DIR, f"{case_name}_{prefix}_principal.png"), dpi=150)
        plt.close(fig)

        # 5. EndBalance
        fig, ax = plt.subplots(figsize=(10, 5))
        ax.plot(periods, df["EndBal"], color="#F44336", linewidth=1.5)
        ax.set_title(f"End Balance — {label}", fontname="Times New Roman", fontsize=13)
        ax.set_xlabel("Periods", fontname="Times New Roman", fontsize=12)
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))
        ax.grid(False)
        fig.tight_layout()
        fig.savefig(os.path.join(BEFORE_DIR, f"{case_name}_{prefix}_endbal.png"), dpi=150)
        plt.close(fig)

        # 6. Stacked bar: Interest vs Principal
        fig, ax = plt.subplots(figsize=(10, 5))
        step = max(1, len(periods) // 60)
        idx = periods[::step]
        int_vals = df["Interest"].values[::step]
        prin_vals = df["Principal"].values[::step]
        ax.bar(idx, int_vals, color="#FF9800", label="Interest", width=step * 0.8)
        ax.bar(idx, prin_vals, bottom=int_vals, color="#2196F3", label="Principal", width=step * 0.8)
        ax.set_title(f"Interest vs Principal — {label}", fontname="Times New Roman", fontsize=13)
        ax.set_xlabel("Periods", fontname="Times New Roman", fontsize=12)
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))
        ax.legend()
        ax.grid(False)
        fig.tight_layout()
        fig.savefig(os.path.join(BEFORE_DIR, f"{case_name}_{prefix}_stacked.png"), dpi=150)
        plt.close(fig)

        # 7. Proportion line
        fig, ax = plt.subplots(figsize=(10, 5))
        ax.plot(periods, df["InterestPct"] * 100, label="Interest %", color="#FF9800")
        ax.plot(periods, df["PrincipalPct"] * 100, label="Principal %", color="#2196F3")
        ax.set_title(f"Interest & Principal Proportion — {label}", fontname="Times New Roman", fontsize=13)
        ax.set_xlabel("Periods", fontname="Times New Roman", fontsize=12)
        ax.set_ylabel("%")
        ax.legend()
        ax.grid(False)
        fig.tight_layout()
        fig.savefig(os.path.join(BEFORE_DIR, f"{case_name}_{prefix}_proportion.png"), dpi=150)
        plt.close(fig)


# ===========================================================================
# PORTFOLIO ANALYSIS CHARTS
# ===========================================================================

def generate_portfolio_charts():
    """Generate portfolio analysis charts mimicking VBA output."""
    np.random.seed(42)
    n = 61
    asset1 = 100 * np.exp(np.cumsum(np.random.normal(0.008, 0.05, n)))
    asset2 = 100 * np.exp(np.cumsum(np.random.normal(0.005, 0.07, n)))
    rf = 100 * np.exp(np.cumsum(np.random.normal(0.002, 0.005, n)))

    results = full_portfolio_analysis(asset1, asset2, rf)
    frontier = results["frontier"]

    # Efficient Frontier scatter plot
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.scatter(frontier["StdDev"] * 100, frontier["Return"] * 100,
               c=frontier["Return"], cmap="RdYlGn", s=30, zorder=3)
    ax.scatter([results["mvp_std"] * 100], [results["mvp_return"] * 100],
               color="blue", s=100, marker="D", zorder=4, label="MVP")
    ax.scatter([results["orp_std"] * 100], [results["orp_return"] * 100],
               color="red", s=100, marker="*", zorder=4, label="ORP")
    ax.set_title("Efficient Frontier", fontname="Times New Roman", fontsize=14)
    ax.set_xlabel("Standard Deviation (%)", fontname="Times New Roman", fontsize=12)
    ax.set_ylabel("Expected Return (%)", fontname="Times New Roman", fontsize=12)
    ax.legend(fontsize=11)
    ax.grid(True, alpha=0.3)
    fig.tight_layout()
    fig.savefig(os.path.join(BEFORE_DIR, "portfolio_frontier.png"), dpi=150)
    plt.close(fig)

    # Sharpe Ratio vs Weight
    fig, ax = plt.subplots(figsize=(10, 6))
    sharpe_values = []
    weights = np.linspace(-0.5, 1.5, 44)
    for w1 in weights:
        w = np.array([w1, 1 - w1])
        ret = w @ results["expected_returns"]
        std = np.sqrt(w @ results["vcv_matrix"] @ w)
        s = (ret - results["rf_rate"]) / std if std > 0 else 0
        sharpe_values.append(s)
    ax.plot(weights, sharpe_values, color="#2196F3", linewidth=2)
    ax.set_title("Sharpe Ratio vs Portfolio Weight (Asset 1)",
                 fontname="Times New Roman", fontsize=14)
    ax.set_xlabel("Weight of Asset 1", fontname="Times New Roman", fontsize=12)
    ax.set_ylabel("Sharpe Ratio", fontname="Times New Roman", fontsize=12)
    ax.grid(True, alpha=0.3)
    fig.tight_layout()
    fig.savefig(os.path.join(BEFORE_DIR, "portfolio_sharpe.png"), dpi=150)
    plt.close(fig)

    # Copy same charts to "after" directory (Power BI equivalent)
    for src_name in ["portfolio_frontier.png", "portfolio_sharpe.png"]:
        src = os.path.join(BEFORE_DIR, src_name)
        dst = os.path.join(AFTER_DIR, src_name.replace(".png", "_powerbi.png"))
        import shutil
        shutil.copy2(src, dst)


# ===========================================================================
# POWER BI "AFTER" STATE CHARTS
# ===========================================================================

def generate_after_charts():
    """Generate Power BI-style charts for the 'after' state."""
    # Case 1: $100k / 5% / 30yr / Monthly / End
    df = constant_payment_schedule(100000, 5, 30, "Monthly", "End of Period")
    df_sl = straight_line_schedule(100000, 5, 30, "Monthly", "End of Period")

    # Dashboard-style summary
    fig, axes = plt.subplots(2, 2, figsize=(14, 10))
    fig.suptitle("Power BI — Loan Amortization Dashboard", fontsize=16, fontweight="bold")

    ax = axes[0, 0]
    ax.plot(df["Period"], df["BegBal"], color="#2196F3", label="Constant Payment")
    ax.plot(df_sl["Period"], df_sl["BegBal"], color="#FF5722", label="Straight-Line", linestyle="--")
    ax.set_title("Beginning Balance Comparison")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))
    ax.legend(fontsize=9)

    ax = axes[0, 1]
    ax.plot(df["Period"], df["Interest"], color="#FF9800", label="Constant")
    ax.plot(df_sl["Period"], df_sl["Interest"], color="#9C27B0", label="Straight-Line", linestyle="--")
    ax.set_title("Interest Comparison")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))
    ax.legend(fontsize=9)

    ax = axes[1, 0]
    ax.plot(df["Period"], df["Principal"], color="#4CAF50", label="Constant")
    ax.plot(df_sl["Period"], df_sl["Principal"], color="#E91E63", label="Straight-Line", linestyle="--")
    ax.set_title("Principal Comparison")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))
    ax.legend(fontsize=9)

    ax = axes[1, 1]
    ax.plot(df["Period"], df["Payment"], color="#3F51B5", label="Constant")
    ax.plot(df_sl["Period"], df_sl["Payment"], color="#795548", label="Straight-Line", linestyle="--")
    ax.set_title("Total Payment Comparison")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))
    ax.legend(fontsize=9)

    for ax in axes.flat:
        ax.set_xlabel("Period")
        ax.grid(True, alpha=0.2)

    fig.tight_layout()
    fig.savefig(os.path.join(AFTER_DIR, "loan_dashboard.png"), dpi=150)
    plt.close(fig)

    # Individual case screenshots
    for case_name, loan, rate, years, freq, ptype in [
        ("case1", 100000, 5, 30, "Monthly", "End of Period"),
        ("case2", 250000, 3.5, 15, "Monthly", "Begin of Period"),
        ("case3", 50000, 7, 5, "Annually", "End of Period"),
    ]:
        df_c = constant_payment_schedule(loan, rate, years, freq, ptype)
        fig, ax = plt.subplots(figsize=(10, 5))
        ax.fill_between(df_c["Period"], df_c["Interest"], alpha=0.4, label="Interest", color="#FF9800")
        ax.fill_between(df_c["Period"], df_c["Interest"],
                        df_c["Interest"] + df_c["Principal"], alpha=0.4, label="Principal", color="#2196F3")
        ax.set_title(f"Loan: ${loan:,} / {rate}% / {years}yr / {freq} / {ptype}")
        ax.set_xlabel("Period")
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))
        ax.legend()
        ax.grid(True, alpha=0.2)
        fig.tight_layout()
        fig.savefig(os.path.join(AFTER_DIR, f"loan_{case_name}.png"), dpi=150)
        plt.close(fig)


# ===========================================================================
# MAIN
# ===========================================================================

if __name__ == "__main__":
    print("Generating 'before' loan charts...")
    generate_loan_charts(100000, 5, 30, "Monthly", "End of Period", "case1")
    generate_loan_charts(250000, 3.5, 15, "Monthly", "Begin of Period", "case2")
    generate_loan_charts(50000, 7, 5, "Annually", "End of Period", "case3")

    print("Generating portfolio charts...")
    generate_portfolio_charts()

    print("Generating 'after' Power BI-style charts...")
    generate_after_charts()

    print("Done! Screenshots saved to:")
    print(f"  Before: {BEFORE_DIR}")
    print(f"  After:  {AFTER_DIR}")
    n_before = len([f for f in os.listdir(BEFORE_DIR) if f.endswith(".png")])
    n_after = len([f for f in os.listdir(AFTER_DIR) if f.endswith(".png")])
    print(f"  Total: {n_before} before + {n_after} after = {n_before + n_after} screenshots")
