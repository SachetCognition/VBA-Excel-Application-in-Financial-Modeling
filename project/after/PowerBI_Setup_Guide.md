# Power BI Financial Modeling Report — Setup Guide

## Overview

This guide describes how to create the `FinancialModeling.pbix` Power BI report
that replaces the original VBA/Excel workbooks:
- `LoanAmortization.xlsm` → DAX-based loan amortization with What-If Parameters
- `Mean variance portfolio analysis Github.xlsm` → Python Script data source

## Page 1: Loan Amortization (Pure DAX)

### Step 1: Create What-If Parameters

In Power BI Desktop: **Modeling → New Parameter**

| Parameter     | Min     | Max       | Increment | Default  |
|---------------|---------|-----------|-----------|----------|
| LoanAmount    | 10,000  | 1,000,000 | 10,000    | 100,000  |
| AnnualRate    | 0.5     | 15.0      | 0.25      | 5.0      |
| LoanYears     | 1       | 30        | 1         | 30       |

### Step 2: Create Disconnected Tables

**Enter Data** to create:

**PaymentFrequency table:**
| Frequency |
|-----------|
| Monthly   |
| Annually  |

**PaymentType table:**
| Type             |
|------------------|
| End of Period    |
| Begin of Period  |

### Step 3: DAX Calculated Table — Constant Payment

```dax
AmortSchedule =
VAR _loan = SELECTEDVALUE('LoanAmount'[LoanAmount], 100000)
VAR _rate = SELECTEDVALUE('AnnualRate'[AnnualRate], 5) / 100
VAR _years = SELECTEDVALUE('LoanYears'[LoanYears], 30)
VAR _freq = SELECTEDVALUE('PaymentFrequency'[Frequency], "Monthly")
VAR _type = SELECTEDVALUE('PaymentType'[Type], "End of Period")
VAR _ppy = IF(_freq = "Monthly", 12, 1)
VAR _r = _rate / _ppy
VAR _n = _years * _ppy
VAR _pmt = IF(_type = "End of Period",
    _loan * _r * POWER(1+_r, _n) / (POWER(1+_r, _n) - 1),
    _loan * _r * POWER(1+_r, _n) / (POWER(1+_r, _n) - 1) / (1+_r)
)
RETURN
ADDCOLUMNS(
    GENERATESERIES(1, _n, 1),
    "Period", [Value],
    "Payment", _pmt,
    "BegBal", _loan * (POWER(1+_r, _n) - POWER(1+_r, [Value]-1)) / (POWER(1+_r, _n) - 1),
    "Interest", _loan * (POWER(1+_r, _n) - POWER(1+_r, [Value]-1)) / (POWER(1+_r, _n) - 1) * _r,
    "Principal", _pmt - _loan * (POWER(1+_r, _n) - POWER(1+_r, [Value]-1)) / (POWER(1+_r, _n) - 1) * _r,
    "EndBal", _loan * (POWER(1+_r, _n) - POWER(1+_r, [Value])) / (POWER(1+_r, _n) - 1)
)
```

### Step 4: DAX Calculated Table — Straight-Line

```dax
StraightLineSchedule =
VAR _loan = SELECTEDVALUE('LoanAmount'[LoanAmount], 100000)
VAR _rate = SELECTEDVALUE('AnnualRate'[AnnualRate], 5) / 100
VAR _years = SELECTEDVALUE('LoanYears'[LoanYears], 30)
VAR _freq = SELECTEDVALUE('PaymentFrequency'[Frequency], "Monthly")
VAR _ppy = IF(_freq = "Monthly", 12, 1)
VAR _r = _rate / _ppy
VAR _n = _years * _ppy
VAR _constPrinc = _loan / _n
RETURN
ADDCOLUMNS(
    GENERATESERIES(1, _n, 1),
    "Period", [Value],
    "BegBal", _loan - _constPrinc * ([Value] - 1),
    "Principal", _constPrinc,
    "Interest", (_loan - _constPrinc * ([Value] - 1)) * _r,
    "Payment", _constPrinc + (_loan - _constPrinc * ([Value] - 1)) * _r,
    "EndBal", _loan - _constPrinc * [Value]
)
```

### Step 5: Create Visuals (14 charts matching original)

**Page 1 — Constant Payment:**
1. BegBalance line chart over Period
2. Payment line chart (constant)
3. Interest line chart (decreasing)
4. Principal line chart (increasing)
5. EndBalance line chart (decreasing)
6. Stacked bar: Interest vs Principal
7. Proportion line: Interest% and Principal% over time

**Page 2 — Straight-Line:**
8-14. Same 7 chart types as above

Add **Bookmark Navigator** with Previous/Next buttons to cycle through charts
(mimics the VBA DisplayingDynamicCharts UserForm).

---

## Page 3: Portfolio Analysis (Python Script Data Source)

### Step 1: Get Data → Python Script

Paste the contents of `portfolio_analysis.py` into the Python script editor.

### Step 2: Select tables to load
- `frontier` — Efficient frontier data (44 points)
- `summary` — Key metrics for cards
- `vcv_table` — Variance-covariance matrix

### Step 3: Create Visuals
- **Scatter plot**: Efficient Frontier (x=StdDev, y=Return)
  - Add data labels for MVP and ORP points
- **Cards**: MVP weights, ORP weights, Sharpe ratio
- **Table**: Full frontier data
- **Matrix**: VCV matrix display

---

## Slicer Setup

Add slicers for:
- LoanAmount (slider)
- AnnualRate (slider)
- LoanYears (slider)
- PaymentFrequency (dropdown)
- PaymentType (dropdown)

All slicers interact with the DAX calculated tables through SELECTEDVALUE().
