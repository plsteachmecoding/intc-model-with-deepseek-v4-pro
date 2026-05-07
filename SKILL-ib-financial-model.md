---
name: ib-financial-model
description: Build investment bank-style financial models in Excel using Python (openpyxl). Covers three-statement linkage (P&L + BS + CF), segment-driven DCF with bottom-up growth drivers, scenario switching, and IB formatting. Use when the user asks to build a financial model, DCF model, valuation model, or create an Excel-based financial analysis from annual reports or 10-K filings.
---

# Investment Bank Financial Model Builder

## Overview

Build programmatic IB-grade Excel financial models via Python `openpyxl`. The model follows a **segment-driven, three-statement linked DCF** architecture with scenario switching and **bottom-up growth driver decomposition**.

## Architecture: 12-Tab Workbook

```
Cover → Key_Summary → Assumptions → Growth_Drivers → Segment_Revenue → Segment_PL
→ Consolidated_PL → BS → Cash_Flow → DCF → Sensitivity → Ratio_Analysis
```

### Data Flow (critical linkage order)

```
Assumptions (scenarios) ──→ Growth_Drivers (bottom-up drivers)
                              ↓
                         Segment_Revenue (pulls from Growth_Drivers)
                              ↓
                         Segment_PL (by division)
                              ↓
                         Consolidated_PL (by cost type)
                              ↓                    ↓
                           BS (balance sheet)    DCF
                              ↓                    ↓
                         Cash_Flow            Sensitivity
                              ↓
                         Ratio_Analysis
```

**Key rule**: Build tabs in dependency order. Store row references in a `dict R = {}` for cross-tab formulas.

## Step 1: Extract Data from Source Documents

### From 10-K / Annual Report PDF

Target these sections (use `Read` with offset/limit for large PDFs):

| Data | Location | Priority |
|------|----------|----------|
| Revenue by segment | Note on Segment Information | P0 |
| Segment operating income | Note on Segment Information | P0 |
| Consolidated P&L | Consolidated Statements of Income | P0 |
| Balance Sheet | Consolidated Balance Sheets | P0 |
| Cash Flow | Consolidated Statements of Cash Flows | P0 |
| Segment cost breakdown | Segment note (emp comp / other) | P1 |
| Share count / EPS | Per-share data in P&L | P1 |
| Equity rollforward | Statements of Stockholders' Equity | P2 |
| **Segment driver data** (units, ASP, share) | Earnings call, 10-Q, industry reports | P1 |

### Searching PDFs

Use `Grep` on the PDF for keywords like "Total assets", "Operating income", "Segment", then `Read` with the right offset. PDF pages ≈ 50-60 lines each.

## Step 2: Model Structure

### Assumptions Tab — The Control Center

```
Row layout: [Parameter | 2023A | 2024A | 2025A | (spacer) | Base | Bull | Bear | Selected | (spacer) | Notes]
```

**CHOOSE/MATCH formula** for scenario switching:
```
=CHOOSE(MATCH($B$2,{"Base","Bull","Bear"},0), F{row}, G{row}, H{row})
```

Every assumption MUST show **historical anchors** (3 years of actuals) so reviewers can judge reasonableness.

**Growth driver assumptions** (expand beyond simple growth rates):

For each business line, add bottom-up driver assumptions:
- **CCG**: PC TAM (M units), Intel PC Unit Share (%), CCG Blended ASP ($)
- **DCAI**: Server TAM (M units), Intel DC Server Share (%), Server ASP ($), AI/ASIC Revenue ($M)
- **Intel Foundry**: Internal Rev per Chip ($), External Foundry Revenue ($M)
- **All Other**: Global LV Production (M units), Mobileye Rev/Vehicle ($), IMS & Other Revenue ($M)

Each driver gets 3 scenario values (Base/Bull/Bear) with CHOOSE/MATCH selection.

### Growth_Drivers Tab — Bottom-Up Revenue Decomposition

This is the **12th tab**, inserted between Assumptions and Segment_Revenue. It decomposes each segment's revenue forecast into granular bottom-up drivers, replacing the simpler `Prior × (1 + growth_rate)` approach.

#### Architecture

```
Growth_Drivers is built AFTER Segment_Revenue pass 1 (historical + structure), BEFORE Segment_Revenue pass 2 (forecast linking).

Build order: Segment_Revenue pass 1 → Growth_Drivers → Segment_Revenue pass 2
```

#### Segment Driver Frameworks

**CCG — Client Computing Group**: Bottom-up, 3 parameters
```
CCG Revenue = PC TAM (M units) × Intel Unit Share (%) × Blended ASP ($)
```
Rows (9 data rows):
1. PC TAM (M units) — historical hardcoded + forecast link to Assumptions
2. Intel PC Unit Share (%) — historical hardcoded + forecast link
3. Intel Implied Units (M) — formula: `=TAM × Share`
4. Intel CCG Blended ASP ($) — historical hardcoded + forecast link
5. Implied CCG Revenue ($M) — formula: `=Units × ASP`
6. Reported CCG Revenue ($M) — green link to Segment_Revenue (historical), mirrors Implied (forecast)
7. Reconciliation ($M) — `=Implied − Reported`
8. Reconciliation (%) — `=Recon / Reported`

**DCAI — Data Center & AI**: Bottom-up, 5 parameters (two-component)
```
DCAI Revenue = Server TAM (M) × Intel DC Share (%) × ASP ($) + AI/ASIC Revenue ($M)
```
Rows (12 data rows):
1. Server TAM (M units)
2. Intel DC Server Unit Share (%)
3. Intel Implied DC Units (M) — formula
4. Intel Server Blended ASP ($)
5. Implied CPU Revenue ($M) — formula
6. AI/ASIC & Other Revenue ($M)
7. Implied DCAI Revenue ($M) — formula: `=CPU Rev + AI/ASIC Rev`
8. Reported DCAI Revenue ($M)
9-10. Reconciliation ($M and %)

**Intel Foundry**: Bottom-up, 2 parameters (two-component)
```
Foundry Revenue = Total Intel Chip Vol (M) × Internal Rev/Chip ($) + External Foundry ($M)
```
- Total Intel Chip Vol links to CCG + DCAI implied units: `=CCG Units + DCAI Units`
- Internal Rev/Chip = internal transfer price (derived, declining with node depreciation)
- External Foundry = 3rd-party wafer customers (growing from small base)

**All Other — Mobileye + IMS + Legacy**: Bottom-up, 3 parameters (two-component)
```
All Other Revenue = Global LV Production (M) × Mobileye Rev/Vehicle ($) + IMS & Other ($M)
```

#### Historical Reconciliation

Every segment MUST show historical reconciliation (Implied − Reported) < 1%. This proves the driver model is well-calibrated.

```python
# Example: CCG Reconciliation
for col in HIST_COLS + FCST_COLS:
    _fval(ws, recon_row, col,
          f"={CL(col)}{implied_rev}-{CL(col)}{reported_rev}")
```

Reconciliation for historical years should be near-zero (within 0.5%). Forecast reconciliation = 0 by construction (Segment_Revenue pulls from Growth_Drivers).

#### Driver Sensitivity Section

At the bottom of Growth_Drivers, add a driver elasticity table:

```
[Driver | FY2026E Value | Change Unit | Revenue Impact ($M) | Impact %]
```

Elasticity formulas:
- **TAM 1% change**: `=ImpliedRev × 0.01`
- **Share 1pp change**: `=TAM × 0.01 × ASP`
- **ASP $1 change**: `=Units` (1:1 per-unit impact)
- **Pass-through items** (AI/ASIC, External Foundry, IMS): hardcoded ±$100M

Highlight **top 3 most sensitive drivers** with analyst commentary in the Notes column.

#### Two-Pass Build Process (Critical)

Growth_Drivers references Segment_Revenue R-keys for historical column links, and Segment_Revenue forecast columns link to Growth_Drivers implied rows. This creates a **circular dependency** that must be broken with two passes:

```python
# Pass 1: Build Segment_Revenue structure + historical, assign R keys
build_segment_revenue(wb, R)
build_segment_revenue_pass1(wb, R)  # leaves forecast columns empty

# Build Growth_Drivers (references segrev R-keys for historical)
build_growth_drivers(wb, R)

# Pass 2: Link Segment_Revenue forecast columns to Growth_Drivers
build_segment_revenue_pass2(wb, R)
# For each segment, forecast cols = Growth_Drivers!{col}{gd_implied_key}
```

```python
# Segment_Revenue pass 2 example
for fc in FCST_COLS:
    f_text = f"=Growth_Drivers!{CL(fc)}{R[gd_key]}"
    ws.cell(row=segrev_row, column=fc, value=f_text)
```

### Segment Revenue — Bottom-Up Build

For each business line:
1. Historical revenue (hardcoded, blue font)
2. **Forecast**: Linked to Growth_Drivers implied revenue (green font, cross-sheet link)
3. YoY % row below each segment
4. Revenue Mix % section at bottom
5. Consolidated Revenue = sum of segments + external Foundry + eliminations

**Important**: With Growth_Drivers, Segment_Revenue forecast columns are pure cross-sheet links — not `Prior × (1+g)` formulas. This makes the entire revenue forecast driver-driven.

### Dual-Layer P&L with Cross-Validation

**Segment_PL**: costs by division (Employee Comp + Other Costs per segment)
**Consolidated_PL**: costs by type (COGS/R&D/S&M/G&A)

Critical: **EBIT in Consolidated_PL = Sum of Segment OI from Segment_PL**

```python
# Consolidated_PL EBIT links to Segment_PL (segment is source of truth)
fval(ws, ebit_row, c, f"=Segment_PL!{CL(c)}{sum_seg_row}")

# Segment_PL cross-check row
fval(ws, diff_row, c, f"={CL(c)}{sum_seg}-{CL(c)}{consol_ebit_link}")
# → must always = 0
```

### Balance Sheet — Cash as Plug

BS forecast logic:

| Line Item | Forecast Method |
|-----------|----------------|
| **Cash** | **PLUG** = Total L&E − NCA − Other CA (ensures A = L+E) |
| Marketable Securities | Grow at stable rate |
| Accounts Receivable | Revenue × AR Days / 365 |
| PP&E | Prior + CapEx − D&A |
| Accounts Payable | COGS × AP Days / 365 |
| Accrued Comp | Revenue × AccComp% |
| Long-term Debt | From Assumptions (absolute) |
| Equity | Prior + NI − Buyback − Dividends + SBC(net) |

**Balance Check row**: `=Total Assets − Total Liabilities − Total Equity` → must = 0

### Cash Flow — WC from BS Delta

```python
# Each WC line links to BS change
# AR change (increase = cash outflow)
f"=-(BS!{CL(c)}{ar_row}-BS!{prev}{ar_row})"
# AP change (increase = cash inflow)
f"=BS!{CL(c)}{ap_row}-BS!{prev}{ap_row}"
```

### DCF: UFCF → Enterprise Value → Equity Value

```
UFCF = NOPAT + D&A − CapEx − ΔNWC
TV = UFCF_terminal × (1+g) / (WACC−g)    [Gordon Growth]
EV = Σ PV(UFCF) + PV(TV)
Equity = EV + Cash − Debt
Price = Equity / Diluted Shares
```

## Step 3: IB Formatting Standards

### Color Coding (must implement + document in legend)

| Style | Meaning | Implementation |
|-------|---------|----------------|
| Blue font (FH) | Historical hardcoded data | `Font(color="305496")` |
| Green font (FL) | Cross-sheet link | `Font(color="548235")` |
| Blue fill + dark font (FI) | Editable assumption | `PatternFill("solid", fgColor="DDEBF7")` |
| Yellow fill (KEY) | Key output row | `PatternFill("solid", fgColor="FFF2CC")` |
| Black font (FN) | Formula / calculated | Default |
| Grey italic (FSM) | Percentage / memo row | `Font(size=9, italic=True, color="808080")` |

### Every Tab Must Have

1. **Title row** (A1, large bold font)
2. **Unit row** (A2, e.g. "USD Million")
3. **Year header** with dark fill + white font
4. **Notes column** (last column) — prediction logic for every row
5. **Data legend** at bottom — color coding explanation
6. **Data sources** at bottom — specific page/note references

### Key Summary Tab (Executive Dashboard)

Links to all other tabs. Contains:
- Core KPIs (Revenue/GP/EBIT/NI + margins + YoY)
- Revenue by segment with mix %
- BS highlights (Total Assets/Equity/Cash)
- ROE / ROA
- Investment highlights + risk warnings (bullet points)

### Ratio Analysis Tab

- Profitability: GPM / OPM / NPM / ROE / ROA
- Efficiency: AR Days / AP Days / CapEx Intensity / CapEx÷D&A
- Leverage: D/A / Current Ratio / Net Cash
- Cash Flow Quality: OCF÷NI / FCF÷Revenue
- Growth: Revenue / NI / EPS YoY
- **Valuation Multiples**: P/E, P/B, P/S, EV/EBITDA (TTM & NTM) with editable stock price cell

## Critical openpyxl Pitfalls

### 1. Percentage Format: Store Decimals, NOT Whole Numbers

**Problem**: Storing 72.0 with format `"0.0%"` displays as "7200.0%" because Excel multiplies by 100. The correct approach is to store 0.72 with `"0.0%"` to display "72.0%".

**Fix**: ALL percentage values (growth rates, margins, shares, tax rates, WACC components) MUST be stored as decimals:
```python
# WRONG — displays as 7200%
_hval(ws, r, col, 72.0, "0.0%")

# RIGHT — displays as 72.0%
_hval(ws, r, col, 0.72, "0.0%")
```

This applies to:
- Segment shares (0.72, not 72.0)
- Growth rates (0.05, not 5.0)
- Operating margins (0.30, not 30.0)
- COGS/R&D/MG&A % of revenue (0.545, not 54.5)
- WACC: Rf (0.0388, not 3.88), ERP (0.055, not 5.50), Kd (0.048, not 4.80)
- D/(D+E) ratio (0.30, not 30.0)
- Tax rate (0.12, not 12.0)

**Impact of getting this wrong**: Contaminates ALL downstream calculations. CCG implied units become 260×72=18,720M instead of 260×0.72=187.2M. NOPAT tax shield becomes EBIT×(1-12.0)=EBIT×(-11) instead of EBIT×(1-0.12)=EBIT×0.88.

### 2. Notes column text starting with `=`

**Problem**: openpyxl treats any cell value starting with `=` as a formula. Text like `"=Revenue×TAC Rate"` becomes an invalid formula → Excel repair deletes it.

**Fix**:
```python
def note(ws, r, text):
    if text and text.startswith('='):
        text = text[1:]  # strip leading '='
    ws.cell(row=r, column=NCOL, value=text)
```

### 3. Notes column overlapping data columns

**Problem**: If Assumptions has 5 forecast years (2026-2030) in cols F-J, and NCOL=10 (J), Notes overwrites 2030E data.

**Fix**: Set `NCOL = 11` (column K) globally. Leave col J free for data.

### 4. CHOOSE/MATCH with array constants

The formula `CHOOSE(MATCH($B$2,{"Base","Bull","Bear"},0),...)` uses inline array `{}`. This works in modern Excel but may warn in older versions. NOT a CSE formula — do NOT wrap in extra braces.

### 5. Cross-sheet formula timing

Build tabs in dependency order. For circular references (e.g., Segment_PL cross-check needs Consolidated_PL EBIT which needs Segment_PL OI), use **two-pass**:
1. First pass: build Segment_PL, leave cross-check link row empty
2. Build Consolidated_PL, store EBIT row number
3. Second pass: go back and fill the cross-check link

```python
# After building Consolidated_PL
ws_seg = wb["Segment_PL"]
for c in range(2, 10):
    ws_seg.cell(row=cross_check_row, column=c,
                value=f"=Consolidated_PL!{CL(c)}{ebit_row}")
```

**Growth_Drivers ↔ Segment_Revenue also requires two-pass** (see Growth_Drivers section above).

### 6. Cash plug formula — subtraction logic

**Wrong**: `f"=...−{CL(c)}{a}+{CL(c)}{b}+{CL(c)}{c_}"` (minus only applies to first term)
**Right**: `f"=...−{CL(c)}{a}−{CL(c)}{b}−{CL(c)}{c_}"` (explicit minus for each)

## Helper Function Pattern

```python
def hval(ws, r, col, val, fmt="#,##0"):
    """Historical value — blue font"""
    c = ws.cell(row=r, column=col, value=val)
    c.number_format = fmt; c.font = Font(color="305496"); c.border = BD

def fval(ws, r, col, formula, fmt="#,##0"):
    """Formula value — black font"""
    c = ws.cell(row=r, column=col, value=formula)
    c.number_format = fmt; c.border = BD

def lval(ws, r, col, formula, fmt="#,##0"):
    """Link from other sheet — green font"""
    c = ws.cell(row=r, column=col, value=formula)
    c.number_format = fmt; c.font = Font(color="548235"); c.border = BD

def arow(ws, r, label, h1, h2, h3, base, bull, bear, fmt, note_text):
    """Assumption row: 3 historical + 3 scenarios + CHOOSE/MATCH selected"""
    # Historical in cols 2-4 (grey fill, blue font)
    # Scenarios in cols 6-8 (blue fill, editable)
    # Selected in col 9 (CHOOSE formula)
    # Note in col NCOL
```

## Verification Checklist

After generating the workbook, verify programmatically:

```python
import openpyxl
wb = openpyxl.load_workbook("output.xlsx")

# 1. 12 tabs present: Cover, Key_Summary, Assumptions, Growth_Drivers,
#    Segment_Revenue, Segment_PL, Consolidated_PL, BS, Cash_Flow,
#    DCF, Sensitivity, Ratio_Analysis
# 2. Growth_Drivers historical reconciliation < 1% for all 4 segments
# 3. Segment_PL cross-check = 0 for historical years
# 4. BS balance check = 0 for historical years
# 5. EBIT link correct: Consolidated_PL!EBIT → Segment_PL!SumOI
# 6. Segment_Revenue forecast = Growth_Drivers implied (all segments, all forecast cols)
# 7. No formula text in Notes column (check XML for <f> tags in Notes col)
# 8. All year columns have data (no overwrites from Notes)
# 9. All "0.0%" format cells contain decimal values (0.xx, not xx.0)
# 10. CHOOSE/MATCH formulas count matches expectation
```

## CapEx Handling

When management provides CapEx guidance (e.g., "$75B per quarter"), use **absolute value inputs** in Assumptions, NOT percentage of revenue. This is more accurate for capital-intensive periods.

```python
# Each forecast year gets its own row with absolute CapEx
arow(ws, 72, "2026E CapEx", hist_23, hist_24, hist_25,
     180000, 185000, 175000, "#,##0", "Q4'25 Earnings: $75B/Q")
```

## External Data Sourcing for Growth Drivers

When building bottom-up drivers, source external data for model inputs:

| Driver | Source | Notes |
|--------|--------|-------|
| PC TAM | Gartner / IDC | Quarterly PC shipment reports |
| Server TAM | IDC / Mercury Research | Server CPU shipment estimates |
| CPU Unit Share | Mercury Research | x86 CPU share by vendor |
| Global LV Production | IHS Markit / S&P Global Mobility | Annual light vehicle production |
| ASIC/AI revenue | Company earnings calls | Segment-level AI disclosures |

**Always cite data sources** in the model Notes column. If primary sources are unavailable (e.g., behind paywalls), use secondary reporting (industry press, earnings transcripts) and clearly flag as estimates.

## Reference Files

- For a complete working example, see [example-structure.md](example-structure.md)
- For a real implementation with all 12 tabs, see [build_intc_model.py](build_intc_model.py)
