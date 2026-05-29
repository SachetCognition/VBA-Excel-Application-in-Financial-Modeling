"""
Run all VBA macro recreations for all test cases and generate all 14 chart types.

This script:
1. Runs loan amortization for 3 test cases (constant + straight-line)
2. Runs portfolio mean-variance analysis
3. Generates all 14 VBA chart types as matplotlib images
4. Exports data for Power BI import
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.gridspec import GridSpec

from before.loan_amortization_vba_recreation import (
    constant_payment_schedule, straight_line_schedule, summary_statistics
)
from before.portfolio_analysis_vba_recreation import full_portfolio_analysis
from after.loan_amortization_dax import (
    dax_amortization_constant, dax_amortization_straight_line
)

SCREENSHOT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                              "screenshots", "after")
DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "after")
os.makedirs(SCREENSHOT_DIR, exist_ok=True)

# ============================================================================
# Test Cases
# ============================================================================
TEST_CASES = [
    {"name": "Case1", "principal": 100000, "rate": 5, "years": 30,
     "freq": "Monthly", "ptype": "End of Period",
     "label": "$100K / 5% / 30yr / Monthly / End"},
    {"name": "Case2", "principal": 250000, "rate": 3.5, "years": 15,
     "freq": "Monthly", "ptype": "Begin of Period",
     "label": "$250K / 3.5% / 15yr / Monthly / Begin"},
    {"name": "Case3", "principal": 50000, "rate": 7, "years": 5,
     "freq": "Annually", "ptype": "End of Period",
     "label": "$50K / 7% / 5yr / Annual / End"},
]

CHART_STYLE = {
    'font.family': 'serif',
    'font.serif': ['Times New Roman'],
    'axes.titlesize': 13,
    'axes.labelsize': 12,
    'legend.fontsize': 11,
}
plt.rcParams.update(CHART_STYLE)


def plot_7_constant_charts(df, case_label, prefix):
    """Generate 7 charts matching VBA constant payment charts."""
    figs = []

    # Chart 1: Beginning Balance over Periods (xlXYScatter)
    fig, ax = plt.subplots(figsize=(7, 4))
    ax.scatter(df['Period'], df['BegBal'], s=8, color='#4472C4')
    ax.set_title(f"The Change of Beginning Balance\n({case_label})")
    ax.set_xlabel("Periods")
    ax.set_ylabel("Balance ($)")
    ax.grid(False)
    fig.tight_layout()
    path = os.path.join(SCREENSHOT_DIR, f"{prefix}_chart1_begbal.png")
    fig.savefig(path, dpi=120)
    figs.append(("BegBalance", path))
    plt.close(fig)

    # Chart 2: Payment over Periods (xlXYScatter)
    fig, ax = plt.subplots(figsize=(7, 4))
    ax.scatter(df['Period'], df['Payment'], s=8, color='#ED7D31')
    ax.set_title(f"The Change of Payment\n({case_label})")
    ax.set_xlabel("Periods")
    ax.set_ylabel("Payment ($)")
    ax.grid(False)
    fig.tight_layout()
    path = os.path.join(SCREENSHOT_DIR, f"{prefix}_chart2_payment.png")
    fig.savefig(path, dpi=120)
    figs.append(("Payment", path))
    plt.close(fig)

    # Chart 3: Interest over Periods (xlXYScatter)
    fig, ax = plt.subplots(figsize=(7, 4))
    ax.scatter(df['Period'], df['Interest'], s=8, color='#A5A5A5')
    ax.set_title(f"The Change of Interest\n({case_label})")
    ax.set_xlabel("Periods")
    ax.set_ylabel("Interest ($)")
    ax.grid(False)
    fig.tight_layout()
    path = os.path.join(SCREENSHOT_DIR, f"{prefix}_chart3_interest.png")
    fig.savefig(path, dpi=120)
    figs.append(("Interest", path))
    plt.close(fig)

    # Chart 4: Principal over Periods (xlXYScatter)
    fig, ax = plt.subplots(figsize=(7, 4))
    ax.scatter(df['Period'], df['Principal'], s=8, color='#FFC000')
    ax.set_title(f"The Change of Principal\n({case_label})")
    ax.set_xlabel("Periods")
    ax.set_ylabel("Principal ($)")
    ax.grid(False)
    fig.tight_layout()
    path = os.path.join(SCREENSHOT_DIR, f"{prefix}_chart4_principal.png")
    fig.savefig(path, dpi=120)
    figs.append(("Principal", path))
    plt.close(fig)

    # Chart 5: End Balance over Periods (xlXYScatter)
    fig, ax = plt.subplots(figsize=(7, 4))
    ax.scatter(df['Period'], df['EndBal'], s=8, color='#5B9BD5')
    ax.set_title(f"The Change of End Balance\n({case_label})")
    ax.set_xlabel("Periods")
    ax.set_ylabel("End Balance ($)")
    ax.grid(False)
    fig.tight_layout()
    path = os.path.join(SCREENSHOT_DIR, f"{prefix}_chart5_endbal.png")
    fig.savefig(path, dpi=120)
    figs.append(("EndBalance", path))
    plt.close(fig)

    # Chart 6: Stacked column - Interest vs Principal (xlColumnStacked)
    fig, ax = plt.subplots(figsize=(7, 4))
    ax.bar(df['Period'], df['Interest'], label='Interest Component',
           color='#4472C4', width=1.0)
    ax.bar(df['Period'], df['Principal'], bottom=df['Interest'],
           label='Principal Repaid', color='#ED7D31', width=1.0)
    ax.set_title(f"Payment for Interest and Principal over Periods\n({case_label})")
    ax.set_xlabel("Periods")
    ax.set_ylabel("Amount ($)")
    ax.legend()
    ax.grid(False)
    fig.tight_layout()
    path = os.path.join(SCREENSHOT_DIR, f"{prefix}_chart6_stacked.png")
    fig.savefig(path, dpi=120)
    figs.append(("StackedIntPrincipal", path))
    plt.close(fig)

    # Chart 7: Proportion line chart (xlLine) - InterestPct vs PrincipalPct
    fig, ax = plt.subplots(figsize=(7, 4))
    ax.plot(df['Period'], df['InterestPct'], label='Int', color='#4472C4')
    ax.plot(df['Period'], df['PrincipalPct'], label='Prin', color='#ED7D31')
    ax.set_title(f"Change of Interest and Principal Proportion\n({case_label})")
    ax.set_xlabel("Periods")
    ax.set_ylabel("Proportion")
    ax.legend()
    ax.grid(False)
    fig.tight_layout()
    path = os.path.join(SCREENSHOT_DIR, f"{prefix}_chart7_proportion.png")
    fig.savefig(path, dpi=120)
    figs.append(("Proportion", path))
    plt.close(fig)

    return figs


def plot_portfolio_charts(analysis, prefix="portfolio"):
    """Generate portfolio analysis charts matching VBA visuals."""
    figs = []
    frontier = analysis['frontier']

    # Efficient Frontier scatter
    fig, ax = plt.subplots(figsize=(8, 6))
    sc = ax.scatter(frontier['StdDev'] * 100, frontier['Return'] * 100,
                    c=frontier.get('Sharpe', None),
                    cmap='RdYlGn', s=60, edgecolors='black', linewidth=0.5,
                    zorder=3)
    # Mark MVP
    mvp_w = analysis['mvp_weights']
    mvp_ret = analysis['mvp_return'] * 100
    mvp_std = analysis['mvp_std'] * 100
    ax.scatter([mvp_std], [mvp_ret], c='blue', s=150, marker='*',
               label=f'MVP ({mvp_w[0]:.1%}, {mvp_w[1]:.1%})', zorder=5)
    # Mark ORP
    orp_w = analysis['orp_weights']
    orp_ret = analysis['orp_return'] * 100
    orp_std = analysis['orp_std'] * 100
    ax.scatter([orp_std], [orp_ret], c='red', s=150, marker='D',
               label=f'ORP ({orp_w[0]:.1%}, {orp_w[1]:.1%})', zorder=5)
    ax.set_xlabel('Standard Deviation (%)')
    ax.set_ylabel('Expected Return (%)')
    ax.set_title('Efficient Frontier with MVP & ORP')
    ax.legend(loc='upper left')
    plt.colorbar(sc, label='Sharpe Ratio')
    ax.grid(True, alpha=0.3)
    fig.tight_layout()
    path = os.path.join(SCREENSHOT_DIR, f"{prefix}_efficient_frontier.png")
    fig.savefig(path, dpi=120)
    figs.append(("EfficientFrontier", path))
    plt.close(fig)

    # Sharpe Ratio across weights
    fig, ax = plt.subplots(figsize=(8, 5))
    sharpe_vals = []
    w1_vals = frontier['w_asset1'].values
    for _, row in frontier.iterrows():
        sr = (row['Return'] - analysis['rf_rate']) / row['StdDev'] if row['StdDev'] > 0 else 0
        sharpe_vals.append(sr)
    ax.plot(w1_vals * 100, sharpe_vals, color='#4472C4', linewidth=2)
    ax.axhline(y=0, color='gray', linestyle='--', alpha=0.5)
    best_idx = np.argmax(sharpe_vals)
    ax.scatter([w1_vals[best_idx] * 100], [sharpe_vals[best_idx]],
               c='red', s=100, zorder=5, label=f'Max Sharpe = {sharpe_vals[best_idx]:.4f}')
    ax.set_xlabel('Weight of Asset 1 (%)')
    ax.set_ylabel('Sharpe Ratio')
    ax.set_title('Sharpe Ratio vs Portfolio Weight')
    ax.legend()
    ax.grid(True, alpha=0.3)
    fig.tight_layout()
    path = os.path.join(SCREENSHOT_DIR, f"{prefix}_sharpe_ratio.png")
    fig.savefig(path, dpi=120)
    figs.append(("SharpeRatio", path))
    plt.close(fig)

    # VCV Matrix heatmap
    fig, ax = plt.subplots(figsize=(6, 5))
    vcv = analysis['vcv_matrix']
    im = ax.imshow(vcv, cmap='Blues', aspect='auto')
    ax.set_xticks([0, 1])
    ax.set_yticks([0, 1])
    ax.set_xticklabels(['Asset 1', 'Asset 2'])
    ax.set_yticklabels(['Asset 1', 'Asset 2'])
    for i in range(2):
        for j in range(2):
            ax.text(j, i, f'{vcv[i, j]:.6f}', ha='center', va='center',
                    fontsize=12, fontweight='bold')
    ax.set_title('Annualized Variance-Covariance Matrix')
    plt.colorbar(im)
    fig.tight_layout()
    path = os.path.join(SCREENSHOT_DIR, f"{prefix}_vcv_matrix.png")
    fig.savefig(path, dpi=120)
    figs.append(("VCVMatrix", path))
    plt.close(fig)

    # Portfolio weights comparison
    fig, ax = plt.subplots(figsize=(7, 5))
    x = np.arange(2)
    width = 0.35
    ax.bar(x - width / 2, mvp_w, width, label='MVP', color='#4472C4')
    ax.bar(x + width / 2, orp_w, width, label='ORP', color='#ED7D31')
    ax.set_xticks(x)
    ax.set_xticklabels(['Asset 1', 'Asset 2'])
    ax.set_ylabel('Weight')
    ax.set_title('Portfolio Weights: MVP vs ORP')
    ax.legend()
    ax.axhline(y=0, color='gray', linestyle='-', alpha=0.3)
    for i, (m, o) in enumerate(zip(mvp_w, orp_w)):
        ax.text(i - width / 2, m + 0.02, f'{m:.1%}', ha='center', fontsize=10)
        ax.text(i + width / 2, o + 0.02, f'{o:.1%}', ha='center', fontsize=10)
    fig.tight_layout()
    path = os.path.join(SCREENSHOT_DIR, f"{prefix}_weights_comparison.png")
    fig.savefig(path, dpi=120)
    figs.append(("WeightsComparison", path))
    plt.close(fig)

    return figs


def create_summary_dashboard(all_charts_info, case_results, portfolio_analysis):
    """Create a summary dashboard image combining key visuals."""
    fig = plt.figure(figsize=(20, 16))
    gs = GridSpec(3, 4, figure=fig, hspace=0.35, wspace=0.3)

    # Title
    fig.suptitle('VBA Financial Modeling → Power BI Migration\nComplete Visual Summary',
                 fontsize=18, fontweight='bold', y=0.98)

    # Row 1: Case 1 key charts (constant payment)
    case1 = case_results[0]
    df1_const = case1['constant']
    ax1 = fig.add_subplot(gs[0, 0])
    ax1.scatter(df1_const['Period'], df1_const['BegBal'], s=2, color='#4472C4')
    ax1.set_title('Case 1: Balance', fontsize=10)
    ax1.set_xlabel('Period', fontsize=8)

    ax2 = fig.add_subplot(gs[0, 1])
    ax2.bar(df1_const['Period'], df1_const['Interest'], color='#4472C4', width=1)
    ax2.bar(df1_const['Period'], df1_const['Principal'],
            bottom=df1_const['Interest'], color='#ED7D31', width=1)
    ax2.set_title('Case 1: Int vs Prin', fontsize=10)
    ax2.set_xlabel('Period', fontsize=8)

    # Case 2 key charts
    case2 = case_results[1]
    df2_const = case2['constant']
    ax3 = fig.add_subplot(gs[0, 2])
    ax3.scatter(df2_const['Period'], df2_const['BegBal'], s=2, color='#4472C4')
    ax3.set_title('Case 2: Balance (Begin)', fontsize=10)
    ax3.set_xlabel('Period', fontsize=8)

    ax4 = fig.add_subplot(gs[0, 3])
    ax4.plot(df2_const['Period'], df2_const['InterestPct'], label='Int', color='#4472C4')
    ax4.plot(df2_const['Period'], df2_const['PrincipalPct'], label='Prin', color='#ED7D31')
    ax4.set_title('Case 2: Proportions', fontsize=10)
    ax4.legend(fontsize=8)

    # Row 2: Straight-line comparison + Case 3
    df1_sl = case1['straight_line']
    ax5 = fig.add_subplot(gs[1, 0])
    ax5.scatter(df1_sl['Period'], df1_sl['BegBal'], s=2, color='#A5A5A5')
    ax5.set_title('Case 1 SL: Balance', fontsize=10)
    ax5.set_xlabel('Period', fontsize=8)

    ax6 = fig.add_subplot(gs[1, 1])
    ax6.bar(df1_sl['Period'], df1_sl['Interest'], color='#4472C4', width=1)
    ax6.bar(df1_sl['Period'], df1_sl['Principal'],
            bottom=df1_sl['Interest'], color='#ED7D31', width=1)
    ax6.set_title('Case 1 SL: Int vs Prin', fontsize=10)

    case3 = case_results[2]
    df3 = case3['constant']
    ax7 = fig.add_subplot(gs[1, 2])
    ax7.scatter(df3['Period'], df3['BegBal'], s=30, color='#4472C4')
    ax7.set_title('Case 3: Balance (Annual)', fontsize=10)

    df3_sl = case3['straight_line']
    ax8 = fig.add_subplot(gs[1, 3])
    ax8.scatter(df3_sl['Period'], df3_sl['Payment'], s=30, color='#ED7D31')
    ax8.set_title('Case 3 SL: Payment', fontsize=10)

    # Row 3: Portfolio analysis
    frontier = portfolio_analysis['frontier']
    ax9 = fig.add_subplot(gs[2, 0:2])
    sc = ax9.scatter(frontier['StdDev'] * 100, frontier['Return'] * 100,
                     c=[((r['Return'] - portfolio_analysis['rf_rate']) / r['StdDev']
                         if r['StdDev'] > 0 else 0)
                        for _, r in frontier.iterrows()],
                     cmap='RdYlGn', s=40, edgecolors='black', linewidth=0.3)
    ax9.set_title('Efficient Frontier', fontsize=10)
    ax9.set_xlabel('Std Dev (%)', fontsize=8)
    ax9.set_ylabel('Return (%)', fontsize=8)

    ax10 = fig.add_subplot(gs[2, 2])
    mvp_w = portfolio_analysis['mvp_weights']
    orp_w = portfolio_analysis['orp_weights']
    x = np.arange(2)
    ax10.bar(x - 0.15, mvp_w, 0.3, label='MVP', color='#4472C4')
    ax10.bar(x + 0.15, orp_w, 0.3, label='ORP', color='#ED7D31')
    ax10.set_xticks(x)
    ax10.set_xticklabels(['A1', 'A2'])
    ax10.set_title('Portfolio Weights', fontsize=10)
    ax10.legend(fontsize=8)

    ax11 = fig.add_subplot(gs[2, 3])
    stats_text = (
        f"Asset1 ER: {portfolio_analysis['expected_returns'][0]:.4f}\n"
        f"Asset2 ER: {portfolio_analysis['expected_returns'][1]:.4f}\n"
        f"RF Rate: {portfolio_analysis['rf_rate']:.4f}\n"
        f"MVP: [{mvp_w[0]:.3f}, {mvp_w[1]:.3f}]\n"
        f"ORP: [{orp_w[0]:.3f}, {orp_w[1]:.3f}]\n"
        f"MVP Ret: {portfolio_analysis['mvp_return']:.4f}\n"
        f"ORP Sharpe: {portfolio_analysis['orp_sharpe']:.4f}"
    )
    ax11.text(0.1, 0.5, stats_text, transform=ax11.transAxes,
              fontsize=9, verticalalignment='center', fontfamily='monospace',
              bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))
    ax11.set_title('Key Metrics', fontsize=10)
    ax11.axis('off')

    path = os.path.join(SCREENSHOT_DIR, "summary_dashboard.png")
    fig.savefig(path, dpi=150, bbox_inches='tight')
    plt.close(fig)
    return path


def main():
    print("=" * 70)
    print("RUNNING ALL VBA MACRO RECREATIONS")
    print("=" * 70)

    all_charts = []
    case_results = []

    # ========================================================================
    # LOAN AMORTIZATION - All 3 test cases
    # ========================================================================
    for tc in TEST_CASES:
        print(f"\n{'─' * 50}")
        print(f"  {tc['label']}")
        print(f"{'─' * 50}")

        # Run constant payment schedule (VBA Part 1)
        df_const = constant_payment_schedule(
            tc['principal'], tc['rate'], tc['years'], tc['freq'], tc['ptype']
        )
        stats_const = summary_statistics(df_const)
        print(f"  Constant Payment: ${df_const['Payment'].iloc[0]:,.2f}")
        print(f"  Total Interest:   ${stats_const['total_interest']:,.2f}")
        print(f"  Total Payments:   ${stats_const['total_payments']:,.2f}")
        print(f"  Final Balance:    ${df_const['EndBal'].iloc[-1]:.4f}")

        # Run straight-line schedule (VBA Part 2)
        df_sl = straight_line_schedule(
            tc['principal'], tc['rate'], tc['years'], tc['freq'], tc['ptype']
        )
        stats_sl = summary_statistics(df_sl)
        print(f"  Straight-Line First Payment: ${df_sl['Payment'].iloc[0]:,.2f}")
        print(f"  Straight-Line Last Payment:  ${df_sl['Payment'].iloc[-1]:,.2f}")
        print(f"  Straight-Line Total Interest: ${stats_sl['total_interest']:,.2f}")

        # Cross-validate with DAX companion
        df_dax = dax_amortization_constant(
            tc['principal'], tc['rate'], tc['years'], tc['freq'], tc['ptype']
        )
        max_diff = abs(df_const['EndBal'] - df_dax['EndBal']).max()
        print(f"  DAX Cross-Validation Max Diff: ${max_diff:.6f}")

        # Generate 7 constant payment charts
        const_charts = plot_7_constant_charts(
            df_const, tc['label'], f"{tc['name']}_const"
        )
        all_charts.extend(const_charts)

        # Generate 7 straight-line charts
        sl_charts = plot_7_constant_charts(
            df_sl, f"{tc['label']} (Straight-Line)", f"{tc['name']}_sl"
        )
        all_charts.extend(sl_charts)

        case_results.append({
            'constant': df_const,
            'straight_line': df_sl,
            'stats_const': stats_const,
            'stats_sl': stats_sl,
        })

        # Export data for Power BI
        df_const.to_csv(os.path.join(DATA_DIR, f"{tc['name']}_constant.csv"), index=False)
        df_sl.to_csv(os.path.join(DATA_DIR, f"{tc['name']}_straight_line.csv"), index=False)

    # ========================================================================
    # PORTFOLIO ANALYSIS
    # ========================================================================
    print(f"\n{'═' * 50}")
    print("  PORTFOLIO MEAN-VARIANCE ANALYSIS")
    print(f"{'═' * 50}")

    # Generate sample data
    np.random.seed(42)
    n = 61
    a1 = 100 * np.exp(np.cumsum(np.random.normal(0.008, 0.05, n)))
    a2 = 100 * np.exp(np.cumsum(np.random.normal(0.005, 0.07, n)))
    rf = 100 * np.exp(np.cumsum(np.random.normal(0.002, 0.005, n)))

    analysis = full_portfolio_analysis(a1, a2, rf)
    print(f"  Asset 1 E(R): {analysis['expected_returns'][0]:.6f}")
    print(f"  Asset 2 E(R): {analysis['expected_returns'][1]:.6f}")
    print(f"  Risk-Free:    {analysis['rf_rate']:.6f}")
    print(f"  MVP Weights:  [{analysis['mvp_weights'][0]:.4f}, {analysis['mvp_weights'][1]:.4f}]")
    print(f"  ORP Weights:  [{analysis['orp_weights'][0]:.4f}, {analysis['orp_weights'][1]:.4f}]")
    print(f"  MVP Return:   {analysis['mvp_return']:.6f}")
    print(f"  ORP Sharpe:   {analysis['orp_sharpe']:.6f}")
    print(f"  Frontier:     {len(analysis['frontier'])} points")

    # Generate portfolio charts
    port_charts = plot_portfolio_charts(analysis)
    all_charts.extend(port_charts)

    # Export frontier data
    analysis['frontier'].to_csv(os.path.join(DATA_DIR, "frontier_data.csv"), index=False)

    # ========================================================================
    # SUMMARY DASHBOARD
    # ========================================================================
    print(f"\n{'═' * 50}")
    print("  GENERATING SUMMARY DASHBOARD")
    print(f"{'═' * 50}")

    # Add Sharpe values to frontier for dashboard
    frontier = analysis['frontier']
    if 'Sharpe' not in frontier.columns:
        frontier['Sharpe'] = frontier.apply(
            lambda r: (r['Return'] - analysis['rf_rate']) / r['StdDev']
            if r['StdDev'] > 0 else 0, axis=1
        )

    dashboard_path = create_summary_dashboard(all_charts, case_results, analysis)
    print(f"  Dashboard saved to: {dashboard_path}")

    # Print summary
    print(f"\n{'═' * 70}")
    print(f"  COMPLETE: Generated {len(all_charts)} charts + 1 dashboard")
    print(f"{'═' * 70}")
    for name, path in all_charts:
        print(f"  [{name}] {os.path.basename(path)}")

    return all_charts, case_results, analysis


if __name__ == "__main__":
    main()
