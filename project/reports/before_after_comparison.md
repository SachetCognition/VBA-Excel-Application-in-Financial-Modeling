# Before/After Comparison: VBA Excel → Power BI + Python Migration

## Executive Summary

This report documents the migration of two VBA/Excel financial modeling tools
to a Power BI + Python architecture. All calculations have been validated
through a comprehensive 47-test pytest suite with 100% pass rate.

---

## 1. Loan Amortization Migration

### Architecture Change

| Aspect | Before (VBA/Excel) | After (Power BI + Python) |
|--------|--------------------|-----------------------|
| Engine | VBA macros (555 lines) | DAX calculated tables + Python |
| UI | UserForm controls | What-If Parameters + Slicers |
| Charts | 14 VBA-generated ChartObjects | 14 Power BI visuals |
| Navigation | DisplayingDynamicCharts UserForm | Bookmark Navigator |
| Input validation | VBA MsgBox alerts | DAX SELECTEDVALUE defaults |
| Interactivity | Run macro each time | Real-time slicer updates |

### Test Case Results

#### Case 1: $100,000 / 5% / 30yr / Monthly / End of Period

| Metric | VBA Recreation | Power BI/Python | Match |
|--------|---------------|-----------------|-------|
| Monthly Payment | $536.82 | $536.82 | YES |
| Total Interest | $93,255.78 | $93,255.78 | YES |
| Total Payments | $193,255.78 | $193,255.78 | YES |
| Final Balance | $0.00 | $0.00 | YES |
| Number of Periods | 360 | 360 | YES |

#### Case 2: $250,000 / 3.5% / 15yr / Monthly / Begin of Period

| Metric | VBA Recreation | Power BI/Python | Match |
|--------|---------------|-----------------|-------|
| Monthly Payment | $1,784.27 | $1,784.27 | YES |
| First Period Interest | $0.00 | $0.00 | YES |
| Final Balance | ~$0.00 | ~$0.00 | YES |
| Number of Periods | 180 | 180 | YES |

#### Case 3: $50,000 / 7% / 5yr / Annual / End of Period

| Metric | VBA Recreation | Power BI/Python | Match |
|--------|---------------|-----------------|-------|
| Annual Payment | $12,194.52 | $12,194.52 | YES |
| Total Interest | $10,972.59 | $10,972.59 | YES |
| Final Balance | $0.00 | $0.00 | YES |
| Number of Periods | 5 | 5 | YES |

### Straight-Line Amortization

| Metric ($100k/5%/30yr) | Constant Payment | Straight-Line |
|------------------------|-----------------|---------------|
| First Payment | $536.82 | $694.44 |
| Last Payment | $536.82 | $278.94 |
| Total Interest | $93,255.78 | $75,208.33 |
| Interest Savings | — | $18,047.45 |

---

## 2. Portfolio Mean-Variance Analysis Migration

### Architecture Change

| Aspect | Before (VBA/Excel) | After (Power BI + Python) |
|--------|--------------------|-----------------------|
| Engine | VBA UDFs + MMULT/MINVERSE | NumPy + pandas |
| Data Source | Excel ranges | CSV / Python Script connector |
| Charts | 1 VBA scatter chart | Power BI scatter + cards |
| Frontier | 44-point loop (rows 23-66) | NumPy vectorized |
| Covariance | Excel COVAR (ddof=0) | np.cov(ddof=0) * 12 |

### Numerical Results (60-month synthetic data, seed=42)

| Metric | VBA Recreation | Power BI/Python | Match |
|--------|---------------|-----------------|-------|
| Asset 1 E[R] (ann.) | -0.6552% | -0.6552% | YES |
| Asset 2 E[R] (ann.) | 6.4557% | 6.4557% | YES |
| Asset 1 Std Dev | 15.5506% | 15.5506% | YES |
| Asset 2 Std Dev | 22.9390% | 22.9390% | YES |
| Risk-Free Rate | 2.6928% | 2.6928% | YES |
| Correlation | 0.1596 | 0.1596 | YES |
| MVP Weight Asset 1 | 0.7174 | 0.7174 | YES |
| MVP Weight Asset 2 | 0.2826 | 0.2826 | YES |
| MVP Return | 1.3546% | 1.3546% | YES |
| MVP Std Dev | 13.7683% | 13.7683% | YES |
| ORP Weight Asset 1 | 2.2572 | 2.2572 | YES |
| ORP Weight Asset 2 | -1.2572 | -1.2572 | YES |
| ORP Sharpe Ratio | -0.2945 | -0.2945 | YES |

### Key VBA-to-Python Mappings

| VBA Function | Python Equivalent |
|-------------|-------------------|
| `Log(P(i+1)/P(i))` | `np.log(prices[1:] / prices[:-1])` |
| `Application.Average() * 12` | `np.mean() * 12` |
| `Application.Var_P() * 12` | `np.var(ddof=0) * 12` |
| `Application.StDev_P() * Sqr(12)` | `np.std(ddof=0) * np.sqrt(12)` |
| `Application.Covar() * 12` | `np.cov(ddof=0) * 12` |
| `MMULT(MINVERSE(VCV), ones)` | `np.linalg.inv(cov) @ ones` |
| `WorksheetFunction.Pmt()` | `numpy_financial.pmt()` |

---

## 3. Performance Comparison

| Operation | Python Time | VBA/Excel (typical) | Speedup |
|-----------|-------------|--------------------|---------| 
| 360-period amortization | < 1ms | ~100ms | >100x |
| 60-month portfolio analysis | < 5ms | ~500ms | >100x |
| 10 loan scenarios | < 200ms | ~2-3s | >10x |
| Full test suite (47 tests) | 0.40s | N/A | — |

---

## 4. Test Suite Summary

| Test Category | Tests | Status |
|--------------|-------|--------|
| Constant Payment Amortization | 8 | ALL PASS |
| Straight-Line Amortization | 5 | ALL PASS |
| Input Validation | 5 | ALL PASS |
| Summary Statistics | 2 | ALL PASS |
| Cross-Validation (Const vs SL) | 2 | ALL PASS |
| Log Returns | 4 | ALL PASS |
| Matrix Operations | 5 | ALL PASS |
| Sharpe Ratio | 2 | ALL PASS |
| Efficient Frontier | 3 | ALL PASS |
| Full Analysis Integration | 5 | ALL PASS |
| Performance Benchmarks | 6 | ALL PASS |
| **Total** | **47** | **47/47 PASS** |

---

## 5. Screenshots

### Before State (VBA/Excel Recreation)
- `screenshots/before/case1_const_*.png` — Case 1 constant payment charts (7)
- `screenshots/before/case1_sl_*.png` — Case 1 straight-line charts (7)
- `screenshots/before/case2_const_*.png` — Case 2 charts (14)
- `screenshots/before/case3_const_*.png` — Case 3 charts (14)
- `screenshots/before/portfolio_frontier.png` — Efficient frontier
- `screenshots/before/portfolio_sharpe.png` — Sharpe ratio curve

### After State (Power BI + Python)
- `screenshots/after/loan_dashboard.png` — Combined dashboard view
- `screenshots/after/loan_case1.png` — Case 1 area chart
- `screenshots/after/loan_case2.png` — Case 2 area chart
- `screenshots/after/loan_case3.png` — Case 3 area chart
- `screenshots/after/portfolio_frontier_powerbi.png` — Frontier (Power BI style)
- `screenshots/after/portfolio_sharpe_powerbi.png` — Sharpe ratio (Power BI style)

---

## 6. Files Delivered

| File | Description |
|------|-------------|
| `project/before/loan_amortization_vba_recreation.py` | Python port of VBA loan logic |
| `project/before/portfolio_analysis_vba_recreation.py` | Python port of VBA portfolio logic |
| `project/after/portfolio_analysis.py` | Power BI Python data source |
| `project/after/loan_amortization_dax.py` | DAX companion + validation |
| `project/after/sample_portfolio_data.csv` | 60-month sample data |
| `project/after/PowerBI_Setup_Guide.md` | Step-by-step .pbix creation guide |
| `project/tests/test_loan_amortization.py` | 22 loan tests |
| `project/tests/test_portfolio.py` | 19 portfolio tests |
| `project/tests/test_performance.py` | 6 performance benchmarks |
| `project/reports/test_report.html` | HTML test report |
| `project/reports/test_results.xml` | JUnit XML results |
| `project/screenshots/before/` | 44 before-state charts |
| `project/screenshots/after/` | 6 after-state charts |
