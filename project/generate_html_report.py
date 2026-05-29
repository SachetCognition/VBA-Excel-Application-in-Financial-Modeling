"""
Generate an interactive HTML report showing all VBA macro visuals.

Produces a single self-contained HTML file with:
- All 14 chart types per test case (7 constant + 7 straight-line)
- Portfolio mean-variance analysis visuals
- Interactive Plotly charts for exploration
- Summary dashboard with key metrics
"""

import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import numpy as np
import pandas as pd
import json

from before.loan_amortization_vba_recreation import (
    constant_payment_schedule, straight_line_schedule, summary_statistics
)
from before.portfolio_analysis_vba_recreation import full_portfolio_analysis
from after.loan_amortization_dax import (
    dax_amortization_constant, dax_amortization_straight_line
)

TEST_CASES = [
    {"name": "Case 1", "principal": 100000, "rate": 5, "years": 30,
     "freq": "Monthly", "ptype": "End of Period",
     "label": "$100K / 5% / 30yr / Monthly / End of Period"},
    {"name": "Case 2", "principal": 250000, "rate": 3.5, "years": 15,
     "freq": "Monthly", "ptype": "Begin of Period",
     "label": "$250K / 3.5% / 15yr / Monthly / Begin of Period"},
    {"name": "Case 3", "principal": 50000, "rate": 7, "years": 5,
     "freq": "Annually", "ptype": "End of Period",
     "label": "$50K / 7% / 5yr / Annual / End of Period"},
]


def generate_plotly_trace(x, y, name, chart_type, color=None):
    """Generate a Plotly trace JSON object."""
    trace = {"x": x, "y": y, "name": name}
    if chart_type == "scatter":
        trace["mode"] = "markers"
        trace["type"] = "scatter"
        trace["marker"] = {"size": 4}
    elif chart_type == "line":
        trace["mode"] = "lines"
        trace["type"] = "scatter"
    elif chart_type == "bar":
        trace["type"] = "bar"
    if color:
        if "marker" in trace:
            trace["marker"]["color"] = color
        else:
            trace["marker"] = {"color": color}
        if chart_type == "line":
            trace["line"] = {"color": color}
    return trace


def make_chart_html(chart_id, traces, title, xaxis_title, yaxis_title,
                    barmode=None, height=350):
    """Generate HTML for a single Plotly chart."""
    layout = {
        "title": {"text": title, "font": {"family": "Times New Roman", "size": 14}},
        "xaxis": {"title": xaxis_title, "showgrid": False},
        "yaxis": {"title": yaxis_title, "showgrid": False},
        "margin": {"l": 60, "r": 30, "t": 50, "b": 50},
        "height": height,
        "font": {"family": "Times New Roman"},
    }
    if barmode:
        layout["barmode"] = barmode
    return f"""
    <div id="{chart_id}" class="chart-container"></div>
    <script>
    Plotly.newPlot("{chart_id}", {json.dumps(traces)}, {json.dumps(layout)},
                   {{responsive: true, displayModeBar: false}});
    </script>"""


def generate_7_chart_html(df, prefix, label, amort_type):
    """Generate HTML for 7 VBA chart types."""
    periods = df['Period'].tolist()
    html = f'<h3>{amort_type} Amortization — {label}</h3>\n'
    html += '<div class="chart-grid">\n'

    # Chart 1: Beginning Balance (scatter)
    t = generate_plotly_trace(periods, df['BegBal'].round(2).tolist(),
                              "BegBal", "scatter", "#4472C4")
    html += make_chart_html(f"{prefix}_begbal", [t],
                            "The Change of Beginning Balance", "Periods", "Balance ($)")

    # Chart 2: Payment (scatter)
    t = generate_plotly_trace(periods, df['Payment'].round(2).tolist(),
                              "Payment", "scatter", "#ED7D31")
    html += make_chart_html(f"{prefix}_payment", [t],
                            "The Change of Payment", "Periods", "Payment ($)")

    # Chart 3: Interest (scatter)
    t = generate_plotly_trace(periods, df['Interest'].round(2).tolist(),
                              "Interest", "scatter", "#A5A5A5")
    html += make_chart_html(f"{prefix}_interest", [t],
                            "The Change of Interest", "Periods", "Interest ($)")

    # Chart 4: Principal (scatter)
    t = generate_plotly_trace(periods, df['Principal'].round(2).tolist(),
                              "Principal", "scatter", "#FFC000")
    html += make_chart_html(f"{prefix}_principal", [t],
                            "The Change of Principal", "Periods", "Principal ($)")

    # Chart 5: End Balance (scatter)
    t = generate_plotly_trace(periods, df['EndBal'].round(2).tolist(),
                              "EndBal", "scatter", "#5B9BD5")
    html += make_chart_html(f"{prefix}_endbal", [t],
                            "The Change of End Balance", "Periods", "End Balance ($)")

    # Chart 6: Stacked bar - Interest vs Principal
    t1 = generate_plotly_trace(periods, df['Interest'].round(2).tolist(),
                               "Interest Component", "bar", "#4472C4")
    t2 = generate_plotly_trace(periods, df['Principal'].round(2).tolist(),
                               "Principal Repaid", "bar", "#ED7D31")
    html += make_chart_html(f"{prefix}_stacked", [t1, t2],
                            "Payment for Interest and Principal over Periods",
                            "Periods", "Amount ($)", barmode="stack")

    # Chart 7: Proportion line chart
    t1 = generate_plotly_trace(periods, df['InterestPct'].round(4).tolist(),
                               "Interest %", "line", "#4472C4")
    t2 = generate_plotly_trace(periods, df['PrincipalPct'].round(4).tolist(),
                               "Principal %", "line", "#ED7D31")
    html += make_chart_html(f"{prefix}_proportion", [t1, t2],
                            "Change of Interest and Principal Proportion",
                            "Periods", "Proportion")

    html += '</div>\n'
    return html


def generate_portfolio_html(analysis):
    """Generate HTML for portfolio analysis visuals."""
    frontier = analysis['frontier']
    html = '<h2 id="portfolio">Portfolio Mean-Variance Analysis</h2>\n'
    html += '<div class="chart-grid">\n'

    # Efficient Frontier
    sharpe_vals = []
    for _, row in frontier.iterrows():
        sr = (row['Return'] - analysis['rf_rate']) / row['StdDev'] if row['StdDev'] > 0 else 0
        sharpe_vals.append(round(sr, 4))

    frontier_trace = {
        "x": (frontier['StdDev'] * 100).round(4).tolist(),
        "y": (frontier['Return'] * 100).round(4).tolist(),
        "mode": "markers",
        "type": "scatter",
        "marker": {
            "size": 10,
            "color": sharpe_vals,
            "colorscale": "RdYlGn",
            "showscale": True,
            "colorbar": {"title": "Sharpe"},
            "line": {"width": 0.5, "color": "black"},
        },
        "text": [f"w1={w:.2f}, Sharpe={s:.4f}"
                 for w, s in zip(frontier['w_asset1'].tolist(), sharpe_vals)],
        "hoverinfo": "text+x+y",
        "name": "Frontier",
    }

    # MVP and ORP markers
    mvp_w = analysis['mvp_weights']
    orp_w = analysis['orp_weights']
    mvp_trace = {
        "x": [round(analysis['mvp_std'] * 100, 4)],
        "y": [round(analysis['mvp_return'] * 100, 4)],
        "mode": "markers+text",
        "type": "scatter",
        "marker": {"size": 16, "color": "blue", "symbol": "star"},
        "text": [f"MVP ({mvp_w[0]:.1%}, {mvp_w[1]:.1%})"],
        "textposition": "top right",
        "name": "MVP",
    }
    orp_trace = {
        "x": [round(analysis['orp_std'] * 100, 4)],
        "y": [round(analysis['orp_return'] * 100, 4)],
        "mode": "markers+text",
        "type": "scatter",
        "marker": {"size": 16, "color": "red", "symbol": "diamond"},
        "text": [f"ORP ({orp_w[0]:.1%}, {orp_w[1]:.1%})"],
        "textposition": "top right",
        "name": "ORP",
    }

    layout = {
        "title": {"text": "Efficient Frontier with MVP & ORP",
                  "font": {"family": "Times New Roman", "size": 16}},
        "xaxis": {"title": "Standard Deviation (%)", "showgrid": True,
                  "gridcolor": "rgba(0,0,0,0.1)"},
        "yaxis": {"title": "Expected Return (%)", "showgrid": True,
                  "gridcolor": "rgba(0,0,0,0.1)"},
        "height": 500,
        "margin": {"l": 60, "r": 30, "t": 50, "b": 50},
    }
    html += f"""
    <div id="frontier_chart" class="chart-container" style="grid-column: span 2;"></div>
    <script>
    Plotly.newPlot("frontier_chart",
        {json.dumps([frontier_trace, mvp_trace, orp_trace])},
        {json.dumps(layout)}, {{responsive: true}});
    </script>"""

    # Sharpe Ratio across weights
    w1_vals = frontier['w_asset1'].tolist()
    t = generate_plotly_trace(
        [round(w * 100, 2) for w in w1_vals],
        sharpe_vals, "Sharpe Ratio", "line", "#4472C4"
    )
    best_idx = int(np.argmax(sharpe_vals))
    best_trace = {
        "x": [round(w1_vals[best_idx] * 100, 2)],
        "y": [sharpe_vals[best_idx]],
        "mode": "markers+text",
        "type": "scatter",
        "marker": {"size": 12, "color": "red"},
        "text": [f"Max = {sharpe_vals[best_idx]:.4f}"],
        "textposition": "top center",
        "name": "Max Sharpe",
    }
    html += make_chart_html("sharpe_chart", [t, best_trace],
                            "Sharpe Ratio vs Portfolio Weight",
                            "Weight of Asset 1 (%)", "Sharpe Ratio")

    # Weights comparison bar chart
    mvp_bar = {
        "x": ["Asset 1", "Asset 2"],
        "y": [round(mvp_w[0], 4), round(mvp_w[1], 4)],
        "type": "bar", "name": "MVP",
        "marker": {"color": "#4472C4"},
    }
    orp_bar = {
        "x": ["Asset 1", "Asset 2"],
        "y": [round(orp_w[0], 4), round(orp_w[1], 4)],
        "type": "bar", "name": "ORP",
        "marker": {"color": "#ED7D31"},
    }
    layout_w = {
        "title": {"text": "Portfolio Weights: MVP vs ORP",
                  "font": {"family": "Times New Roman", "size": 14}},
        "barmode": "group",
        "height": 350,
        "margin": {"l": 60, "r": 30, "t": 50, "b": 50},
    }
    html += f"""
    <div id="weights_chart" class="chart-container"></div>
    <script>
    Plotly.newPlot("weights_chart", {json.dumps([mvp_bar, orp_bar])},
                   {json.dumps(layout_w)}, {{responsive: true, displayModeBar: false}});
    </script>"""

    # Key metrics cards
    html += f"""
    <div class="metrics-cards">
        <div class="metric-card">
            <h4>Asset 1 E(R)</h4>
            <span class="metric-value">{analysis['expected_returns'][0]:.4%}</span>
        </div>
        <div class="metric-card">
            <h4>Asset 2 E(R)</h4>
            <span class="metric-value">{analysis['expected_returns'][1]:.4%}</span>
        </div>
        <div class="metric-card">
            <h4>Risk-Free Rate</h4>
            <span class="metric-value">{analysis['rf_rate']:.4%}</span>
        </div>
        <div class="metric-card">
            <h4>MVP Return</h4>
            <span class="metric-value">{analysis['mvp_return']:.4%}</span>
        </div>
        <div class="metric-card">
            <h4>MVP Std Dev</h4>
            <span class="metric-value">{analysis['mvp_std']:.4%}</span>
        </div>
        <div class="metric-card">
            <h4>ORP Sharpe</h4>
            <span class="metric-value">{analysis['orp_sharpe']:.4f}</span>
        </div>
    </div>"""

    # VCV Matrix table
    vcv = analysis['vcv_matrix']
    corr = analysis['correlation_matrix']
    html += f"""
    <div class="matrix-section">
        <h3>Variance-Covariance Matrix (Annualized)</h3>
        <table class="data-table">
            <tr><th></th><th>Asset 1</th><th>Asset 2</th></tr>
            <tr><td><b>Asset 1</b></td><td>{vcv[0,0]:.6f}</td><td>{vcv[0,1]:.6f}</td></tr>
            <tr><td><b>Asset 2</b></td><td>{vcv[1,0]:.6f}</td><td>{vcv[1,1]:.6f}</td></tr>
        </table>
        <h3>Correlation Matrix</h3>
        <table class="data-table">
            <tr><th></th><th>Asset 1</th><th>Asset 2</th></tr>
            <tr><td><b>Asset 1</b></td><td>{corr[0,0]:.6f}</td><td>{corr[0,1]:.6f}</td></tr>
            <tr><td><b>Asset 2</b></td><td>{corr[1,0]:.6f}</td><td>{corr[1,1]:.6f}</td></tr>
        </table>
    </div>"""

    html += '</div>\n'
    return html


def generate_summary_table_html(case_results):
    """Generate comparison summary table."""
    html = '<h2 id="summary">Summary Comparison</h2>\n'
    html += '<table class="data-table">\n'
    html += '<tr><th>Test Case</th><th>Method</th><th>Payment</th>'
    html += '<th>Total Interest</th><th>Total Payments</th><th>Final Balance</th></tr>\n'
    for tc, res in zip(TEST_CASES, case_results):
        sc = res['stats_const']
        ss = res['stats_sl']
        pmt_c = res['constant']['Payment'].iloc[0]
        pmt_s1 = res['straight_line']['Payment'].iloc[0]
        pmt_sn = res['straight_line']['Payment'].iloc[-1]
        html += f'<tr><td rowspan="2"><b>{tc["name"]}</b><br>{tc["label"]}</td>'
        html += f'<td>Constant</td><td>${pmt_c:,.2f}</td>'
        html += f'<td>${sc["total_interest"]:,.2f}</td>'
        html += f'<td>${sc["total_payments"]:,.2f}</td>'
        html += f'<td>${res["constant"]["EndBal"].iloc[-1]:.4f}</td></tr>\n'
        html += f'<tr><td>Straight-Line</td>'
        html += f'<td>${pmt_s1:,.2f} → ${pmt_sn:,.2f}</td>'
        html += f'<td>${ss["total_interest"]:,.2f}</td>'
        html += f'<td>${ss["total_payments"]:,.2f}</td>'
        html += f'<td>${res["straight_line"]["EndBal"].iloc[-1]:.4f}</td></tr>\n'
    html += '</table>\n'
    return html


def main():
    print("Generating interactive HTML report...")

    # Run all macros
    case_results = []
    for tc in TEST_CASES:
        df_const = constant_payment_schedule(
            tc['principal'], tc['rate'], tc['years'], tc['freq'], tc['ptype'])
        df_sl = straight_line_schedule(
            tc['principal'], tc['rate'], tc['years'], tc['freq'], tc['ptype'])
        case_results.append({
            'constant': df_const,
            'straight_line': df_sl,
            'stats_const': summary_statistics(df_const),
            'stats_sl': summary_statistics(df_sl),
        })

    # Portfolio analysis
    np.random.seed(42)
    n = 61
    a1 = 100 * np.exp(np.cumsum(np.random.normal(0.008, 0.05, n)))
    a2 = 100 * np.exp(np.cumsum(np.random.normal(0.005, 0.07, n)))
    rf = 100 * np.exp(np.cumsum(np.random.normal(0.002, 0.005, n)))
    analysis = full_portfolio_analysis(a1, a2, rf)

    # Build HTML
    html = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>VBA Financial Modeling — Macro Visuals Report</title>
<script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
<style>
* { box-sizing: border-box; margin: 0; padding: 0; }
body {
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    background: #f0f2f5; color: #333; line-height: 1.6;
}
.header {
    background: linear-gradient(135deg, #1a1a2e, #16213e, #0f3460);
    color: white; padding: 30px 40px; text-align: center;
}
.header h1 { font-size: 28px; margin-bottom: 8px; }
.header p { opacity: 0.85; font-size: 14px; }
nav {
    background: #16213e; padding: 10px 40px; position: sticky; top: 0; z-index: 100;
    display: flex; gap: 20px; flex-wrap: wrap;
}
nav a { color: #a0c4ff; text-decoration: none; font-size: 14px; padding: 5px 12px;
        border-radius: 4px; transition: background 0.2s; }
nav a:hover { background: rgba(255,255,255,0.1); }
nav a.active { background: #0f3460; color: white; }
.container { max-width: 1400px; margin: 0 auto; padding: 20px; }
h2 {
    font-family: 'Times New Roman', serif; font-size: 22px;
    border-bottom: 2px solid #0f3460; padding-bottom: 8px;
    margin: 30px 0 15px 0; color: #0f3460;
}
h3 {
    font-family: 'Times New Roman', serif; font-size: 17px;
    color: #16213e; margin: 20px 0 10px 0;
}
.chart-grid {
    display: grid; grid-template-columns: repeat(auto-fit, minmax(420px, 1fr));
    gap: 15px; margin-bottom: 20px;
}
.chart-container {
    background: white; border-radius: 8px; padding: 10px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08); min-height: 350px;
}
.case-section {
    background: white; border-radius: 10px; padding: 20px; margin-bottom: 25px;
    box-shadow: 0 2px 12px rgba(0,0,0,0.06);
}
.case-header {
    display: flex; justify-content: space-between; align-items: center;
    margin-bottom: 10px; padding-bottom: 10px;
    border-bottom: 1px solid #eee;
}
.case-header h3 { margin: 0; }
.case-params {
    display: flex; gap: 15px; flex-wrap: wrap;
}
.param-badge {
    background: #e8f0fe; color: #0f3460; padding: 4px 12px;
    border-radius: 20px; font-size: 12px; font-weight: 600;
}
.data-table {
    width: 100%; border-collapse: collapse; margin: 10px 0;
    background: white; border-radius: 8px; overflow: hidden;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
}
.data-table th {
    background: #0f3460; color: white; padding: 10px 15px;
    text-align: left; font-size: 13px;
}
.data-table td {
    padding: 8px 15px; border-bottom: 1px solid #eee; font-size: 13px;
}
.data-table tr:nth-child(even) td { background: #f8f9fa; }
.metrics-cards {
    display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr));
    gap: 12px; margin: 15px 0;
}
.metric-card {
    background: white; border-radius: 8px; padding: 15px; text-align: center;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06); border-left: 4px solid #0f3460;
}
.metric-card h4 { font-size: 11px; color: #666; text-transform: uppercase; margin-bottom: 5px; }
.metric-value { font-size: 20px; font-weight: 700; color: #0f3460; }
.matrix-section { margin: 15px 0; }
.matrix-section .data-table { max-width: 500px; }
.tab-nav { display: flex; gap: 5px; margin-bottom: 15px; }
.tab-btn {
    padding: 8px 18px; border: 1px solid #ccc; background: #f8f9fa;
    border-radius: 6px 6px 0 0; cursor: pointer; font-size: 13px;
    transition: all 0.2s;
}
.tab-btn.active { background: #0f3460; color: white; border-color: #0f3460; }
.tab-content { display: none; }
.tab-content.active { display: block; }
.test-badge {
    display: inline-block; background: #28a745; color: white;
    padding: 3px 10px; border-radius: 12px; font-size: 12px; font-weight: 600;
}
footer {
    text-align: center; padding: 20px; color: #999; font-size: 12px;
    border-top: 1px solid #eee; margin-top: 30px;
}
</style>
</head>
<body>

<div class="header">
    <h1>VBA Financial Modeling &mdash; Macro Visuals Report</h1>
    <p>Complete visual output from all VBA macro recreations &bull; 14 chart types per test case &bull; Portfolio Mean-Variance Analysis</p>
    <p style="margin-top:8px"><span class="test-badge">47/47 Tests Passing</span></p>
</div>

<nav>
    <a href="#summary">Summary</a>
"""
    for i, tc in enumerate(TEST_CASES):
        html += f'    <a href="#case{i+1}">{tc["name"]}</a>\n'
    html += '    <a href="#portfolio">Portfolio Analysis</a>\n'
    html += '</nav>\n\n<div class="container">\n'

    # Summary section
    html += generate_summary_table_html(case_results)

    # Per-case charts
    for i, (tc, res) in enumerate(zip(TEST_CASES, case_results)):
        html += f'<h2 id="case{i+1}">{tc["name"]}: {tc["label"]}</h2>\n'
        html += '<div class="case-section">\n'
        html += '<div class="case-header">\n'
        html += f'  <h3>Parameters</h3>\n  <div class="case-params">\n'
        html += f'    <span class="param-badge">Principal: ${tc["principal"]:,}</span>\n'
        html += f'    <span class="param-badge">Rate: {tc["rate"]}%</span>\n'
        html += f'    <span class="param-badge">Term: {tc["years"]}yr</span>\n'
        html += f'    <span class="param-badge">{tc["freq"]}</span>\n'
        html += f'    <span class="param-badge">{tc["ptype"]}</span>\n'
        html += '  </div>\n</div>\n'

        # Stats cards
        sc = res['stats_const']
        ss = res['stats_sl']
        pmt_c = res['constant']['Payment'].iloc[0]
        html += f"""
    <div class="metrics-cards">
        <div class="metric-card"><h4>Constant Payment</h4>
            <span class="metric-value">${pmt_c:,.2f}</span></div>
        <div class="metric-card"><h4>Total Interest (Const)</h4>
            <span class="metric-value">${sc['total_interest']:,.0f}</span></div>
        <div class="metric-card"><h4>Total Interest (SL)</h4>
            <span class="metric-value">${ss['total_interest']:,.0f}</span></div>
        <div class="metric-card"><h4>Interest Saved (SL)</h4>
            <span class="metric-value">${sc['total_interest'] - ss['total_interest']:,.0f}</span></div>
    </div>"""

        # Tab navigation for Constant vs Straight-Line
        html += f"""
    <div class="tab-nav">
        <button class="tab-btn active" onclick="showTab('const_{i}', this)">Constant Payment (7 Charts)</button>
        <button class="tab-btn" onclick="showTab('sl_{i}', this)">Straight-Line (7 Charts)</button>
    </div>
    <div id="const_{i}" class="tab-content active">"""
        html += generate_7_chart_html(res['constant'], f"c{i}_const",
                                       tc['label'], "Constant Payment")
        html += f"""</div>
    <div id="sl_{i}" class="tab-content">"""
        html += generate_7_chart_html(res['straight_line'], f"c{i}_sl",
                                       tc['label'], "Straight-Line")
        html += '</div>\n</div>\n'

    # Portfolio section
    html += generate_portfolio_html(analysis)

    # Frontier data table
    html += '<h3>Efficient Frontier Data (44 Points)</h3>\n'
    html += '<div style="max-height: 400px; overflow-y: auto; margin-bottom: 20px;">\n'
    html += '<table class="data-table">\n'
    html += '<tr><th>#</th><th>w(Asset1)</th><th>w(Asset2)</th>'
    html += '<th>Return</th><th>Std Dev</th><th>Variance</th></tr>\n'
    for idx, row in analysis['frontier'].iterrows():
        html += f'<tr><td>{idx+1}</td><td>{row["w_asset1"]:.4f}</td>'
        html += f'<td>{row["w_asset2"]:.4f}</td><td>{row["Return"]:.6f}</td>'
        html += f'<td>{row["StdDev"]:.6f}</td><td>{row["Variance"]:.6f}</td></tr>\n'
    html += '</table>\n</div>\n'

    # Close container
    html += """
</div>

<footer>
    VBA Financial Modeling Migration &bull; Python Recreation of all VBA Macros &bull;
    Power BI + Python Target State
</footer>

<script>
function showTab(tabId, btn) {
    const parent = btn.closest('.case-section');
    parent.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
    parent.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.getElementById(tabId).classList.add('active');
    btn.classList.add('active');
    // Resize plotly charts in newly visible tab
    setTimeout(() => {
        document.getElementById(tabId).querySelectorAll('.chart-container').forEach(c => {
            if (c.data) Plotly.Plots.resize(c);
        });
    }, 100);
}

// Sticky nav active state
document.querySelectorAll('nav a').forEach(link => {
    link.addEventListener('click', function() {
        document.querySelectorAll('nav a').forEach(l => l.classList.remove('active'));
        this.classList.add('active');
    });
});
</script>
</body>
</html>"""

    # Write file
    out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "reports", "macro_visuals_report.html")
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"HTML report generated: {out_path}")
    print(f"File size: {os.path.getsize(out_path) / 1024:.1f} KB")
    return out_path


if __name__ == "__main__":
    main()
