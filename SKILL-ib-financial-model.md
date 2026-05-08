---
name: ib-financial-model
description: Build investment bank-style financial models in Excel using Python (openpyxl). Covers three-statement linkage (P&L + BS + CF), segment-driven DCF with bottom-up growth drivers, scenario switching, and IB formatting. Supports both annual-only and quarterly+annual hybrid layouts. Use when the user asks to build a financial model, DCF model, valuation model, or create an Excel-based financial analysis from annual reports or 10-K filings.
---

# Investment Bank Financial Model Builder

## Overview

Build programmatic IB-grade Excel financial models via Python `openpyxl`. The model follows a **segment-driven, three-statement linked DCF** architecture with scenario switching and **bottom-up growth driver decomposition**.

Supports two column layouts:
- **Annual-only** (NCOL=11): All tabs use columns B-K (3 hist + 3 fcst + spacer + Selected + Notes)
- **Quarterly+Annual hybrid** (GD_NCOL=35): Growth_Drivers, Assumptions, Segment_PL use quarterly columns (Q1-Q4 + FY rollup for all 6 years); other tabs stay annual

## Architecture: 11-Tab Workbook (v2 Optimized)

```
Cover → Key_Summary → Assumptions → Growth_Drivers → Segment_PL
→ Consolidated_PL → BS → Cash_Flow → DCF → Sensitivity → Ratio_Analysis
```

**v2 key change**: Segment_Revenue merged into Segment_PL. Revenue + Operating Income are now in ONE tab per segment, simplifying cross-references and reducing tab count from 12 to 11.

### Data Flow (critical linkage order)

```
Assumptions (quarterly, Base/Bull/Bear rows) ──→ Growth_Drivers (quarterly, bottom-up)
                                                      ↓
                                                 Segment_PL (quarterly, 3-pass build)
                                                      ↓ (annual rollup cols)
                                                 Consolidated_PL (annual, by cost type)
                                                      ↓                    ↓
                                                   BS (annual)          DCF (annual)
                                                      ↓                    ↓
                                                 Cash_Flow (annual)   Sensitivity
                                                      ↓
                                                 Ratio_Analysis (annual)
```

**Key rule**: Build tabs in dependency order. Store row references in a `dict R = {}` for cross-tab formulas. For quarterly-to-annual links, use column mapping dicts (SEGPL_COL_MAP, ASP_ANN_COLS).

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
| **Quarterly segment revenue** | 10-Q filings | P1 |

### Searching PDFs

Use `Grep` on the PDF for keywords like "Total assets", "Operating income", "Segment", then `Read` with the right offset. PDF pages ≈ 50-60 lines each.

## Step 2: Column Layout Design

### Option A: Annual-Only (Simpler, for quick models)

```
Col:  B       C       D       E    F       G       H       I         J    K
Use:  2023A   2024A   2025A   sp   2026E   2027E   2028E   Selected  sp   Notes
Idx:  2       3       4       5    6       7       8       9         10   11
```

NCOL = 11. All tabs use this layout. Simpler cross-references, faster build.

### Option B: Quarterly+Annual Hybrid (Recommended for detailed models)

Growth_Drivers, Assumptions, Segment_PL use 35-column quarterly layout:

```
Col:  B-Q1'23 C-Q2'23 D-Q3'23 E-Q4'23 F-FY23  G-Q1'24 ... P-FY25  Q=spacer
      R-Q1'26 S-Q2'26 T-Q3'26 U-Q4'26 V-FY26  W-Q1'27 ... AF-FY28  AG=Sel AH=sp AI=Notes
```

GD_NCOL = 35. Other tabs (Consolidated_PL, BS, CF, DCF, etc.) stay annual (NCOL=11).

**Column mapping for cross-layout references:**
```python
# Annual tab column → Segment_PL quarterly annual rollup column
SEGPL_COL_MAP = {2: 6, 3: 11, 4: 16, 6: 22, 7: 27, 8: 32}

# Annual tab forecast column → Assumptions annual rollup column
ASP_ANN_COLS = {6: 22, 7: 27, 8: 32}

# Unified iterator for all 6 years
GD_ALL_YEARS = [
    (2023, [2,3,4,5], 6, "A"),
    (2024, [7,8,9,10], 11, "A"),
    (2025, [12,13,14,15], 16, "A"),
    (2026, [18,19,20,21], 22, "E"),
    (2027, [23,24,25,26], 27, "E"),
    (2028, [28,29,30,31], 32, "E"),
]
# Each tuple: (year, [q1_col, q2_col, q3_col, q4_col], annual_col, suffix)
```

### Layout-Aware Helpers

```python
GD_NCOL = 35
NCOL = 11

def _is_quarterly(ws):
    """True for tabs using 35-col quarterly layout."""
    return ws.title in ("Growth_Drivers", "Assumptions", "Segment_PL")

def _ncol_for(ws):
    """Return the Notes/width column for this tab."""
    return GD_NCOL if _is_quarterly(ws) else NCOL

def _qheader(ws, r):
    """Write quarterly column headers (Q1'23A ... FY2028E + Selected + Notes)."""
    headers = {
        2:"Q1'23A",3:"Q2'23A",4:"Q3'23A",5:"Q4'23A",6:"FY2023A",
        7:"Q1'24A",8:"Q2'24A",9:"Q3'24A",10:"Q4'24A",11:"FY2024A",
        12:"Q1'25A",13:"Q2'25A",14:"Q3'25A",15:"Q4'25A",16:"FY2025A",
        18:"Q1'26E",19:"Q2'26E",20:"Q3'26E",21:"Q4'26E",22:"FY2026E",
        23:"Q1'27E",24:"Q2'27E",25:"Q3'27E",26:"Q4'27E",27:"FY2027E",
        28:"Q1'28E",29:"Q2'28E",30:"Q3'28E",31:"Q4'28E",32:"FY2028E",
    }
    for col, txt in headers.items():
        c = ws.cell(row=r, column=col, value=txt)
        c.font = FONT_HEADER
        c.fill = FILL_HEADER if col not in {6,11,16,22,27,32} else PatternFill("solid", fgColor="2E75B6")
        c.border = BD; c.alignment = ALIGN_CENTER
    # Selected + Notes columns
    ws.cell(row=r, column=GD_SEL_COL, value="Selected").font = FONT_HEADER
    ws.cell(row=r, column=GD_NCOL, value="Notes").font = FONT_HEADER
```

## Step 3: Model Structure

### Assumptions Tab — The Control Center (Quarterly)

**v2 format**: Parameters as rows, time as columns. Base/Bull/Bear are separate rows (not columns). This allows quarterly adjustment of each scenario independently.

```
Row structure (5 rows per parameter):
  Row N:   [Label]           Q1'23A Q2'23A ... FY2028E [Notes]
  Row N+1:   Base            [quarterly values for Base scenario]
  Row N+2:   Bull            [quarterly values for Bull scenario]
  Row N+3:   Bear            [quarterly values for Bear scenario]
  Row N+4:   Selected        [CHOOSE/MATCH per quarter: picks Base/Bull/Bear]
```

**Quarterly CHOOSE/MATCH** (one per forecast quarter):
```
=CHOOSE(MATCH($B$2,{"Base","Bull","Bear"},0), Base!Q{col}, Bull!Q{col}, Bear!Q{col})
```

**Annual rollup** (for each forecast year):
```
Sum-type (revenue, dollars): =SUM(Q1:Q4)     # row = Selected row
Avg-type (margins, rates):   =AVERAGE(Q1:Q4)  # row = Selected row
```

**Helper for quarterly assumption rows:**
```python
def _aq(label, hist_qvals, base_qvals, bull_qvals, bear_qvals, fmt_val, note_text, driver_type="sum"):
    """Quarterly assumption: 5 rows (label + Base/Bull/Bear/Selected).
    driver_type: 'sum' -> annual=SUM(Q1:Q4), 'avg' -> annual=AVERAGE(Q1:Q4)."""
```

Every assumption MUST show **historical anchors** (quarterly actuals where available) so reviewers can judge reasonableness.

**Growth driver assumptions** (bottom-up, not simple growth rates):

For each business line, add bottom-up driver assumptions:
- **CCG**: PC TAM (M units), Intel PC Unit Share (%), CCG Blended ASP ($)
- **DCAI**: Server TAM (M units), Intel DC Server Share (%), Server ASP ($), AI/ASIC Revenue ($M)
- **Intel Foundry**: Internal Rev per Chip ($), External Foundry Revenue ($M)
- **All Other**: Global LV Production (M units), Mobileye Rev/Vehicle ($), IMS & Other Revenue ($M)

Each driver gets 3 scenario quarterly values (Base/Bull/Bear) with CHOOSE/MATCH selection.

### Growth_Drivers Tab — Bottom-Up Revenue Decomposition (Quarterly)

This tab decomposes each segment's revenue forecast into granular bottom-up drivers. In the quarterly layout, it has 6 years × (4 quarters + 1 annual) = 30 data columns.

#### Architecture

```
Growth_Drivers is built AFTER Segment_PL pass 1 (historical + structure), BEFORE Segment_PL pass 2 (forecast linking).

Build order: Segment_PL pass 1 → Growth_Drivers → Segment_PL pass 2
```

**Critical**: Both Growth_Drivers and Segment_PL use the same quarterly column layout, so they can reference each other at identical column indices. This eliminates column mapping for GD↔SegPL quarterly references.

#### Segment Driver Frameworks

**CCG — Client Computing Group**: Bottom-up, 3 parameters
```
CCG Revenue = PC TAM (M units) × Intel Unit Share (%) × Blended ASP ($)
```
Rows:
1. PC TAM (M units) — hist hardcoded, fcst = Assumptions quarterly / 4
2. Intel PC Unit Share (%) — hist hardcoded, fcst = Assumptions quarterly (avg-type)
3. Intel Implied Units (M) — formula: `=TAM × Share`
4. Intel CCG Blended ASP ($) — hist hardcoded, fcst = Assumptions quarterly
5. Implied CCG Revenue ($M) — formula: `=Units × ASP` (annual = SUM of quarters)
6. Reported CCG Revenue ($M) — hist links to Segment_PL, fcst mirrors Implied

**DCAI — Data Center & AI**: Two-component
```
DCAI Revenue = Server TAM × DC Share × ASP + AI/ASIC Revenue
```

**Intel Foundry**: Two-component
```
Foundry Revenue = Total Intel Chip Vol × Internal Rev/Chip + External Foundry
```
- Total Intel Chip Vol links to CCG + DCAI implied units

**All Other — Mobileye + IMS + Legacy**: Two-component
```
All Other Revenue = LV Production × MBLY Rev/Vehicle + IMS & Other
```

#### Driver Type: Sum vs Average

```python
# Sum-type: annual = SUM(Q1:Q4) — revenue, dollars, units
_fval(ws, r, ann_col, f"=SUM({CL(qcols[0])}{r}:{CL(qcols[3])}{r})", "#,##0")

# Avg-type: annual = AVERAGE(Q1:Q4) — percentages, ratios, ASP
_fval(ws, r, ann_col, f"=AVERAGE({CL(qcols[0])}{r}:{CL(qcols[3])}{r})", "0.0%")
```

#### Historical Reconciliation

Every segment MUST show historical reconciliation (Implied − Reported) < 1%. For quarterly data, reconciliation is at the annual column level.

#### Driver Sensitivity Section

At the bottom, add elasticity table showing revenue impact of ±1% changes in each key driver.

### Segment_PL — Revenue + Operating Income (Quarterly, 3-Pass Build)

**v2**: Revenue and OI are combined in ONE tab. Previously Segment_Revenue and Segment_PL were separate tabs — merging eliminates cross-tab references and reduces tab count.

#### Tab Structure (per segment)

```
[Segment Name]        ← bold label
  Revenue             Q1'23 Q2'23 ... FY2028E  (hist hardcoded, fcst linked to GD)
  YoY%                                        (annual cols only)
  Operating Income    Q1'23 Q2'23 ... FY2028E  (hist=annual/4, fcst=Rev×OPM%)
  OPM%                                        (OI/Revenue)
```

Followed by:
- Corporate Unallocated (annual-only)
- Sum of Segment OI + Corp Unalloc (KEY OUTPUT → links to Consolidated_PL EBIT)
- Cross-check: Sum Seg OI − Consolidated_PL EBIT = 0
- External Foundry Revenue (annual-only)
- Intersegment Eliminations (annual-only)
- TOTAL Consolidated Revenue (annual-only)
- Revenue Mix % (annual-only)

#### Three-Pass Build (Critical for Circular Dependency)

**Pass 1** (`build_segment_pl`): Structure + historical hardcoded values. Forecast quarterly cells get borders only (empty formulas). Assigns all R keys.

**Pass 2** (`build_segment_pl_pass2`): After Growth_Drivers is built, links forecast quarterly revenue to GD implied revenue at identical quarterly columns:
```python
for year, qcols, ann_col, suffix in GD_ALL_YEARS[3:]:  # forecast years only
    for qc in qcols:
        f_text = f"=Growth_Drivers!{CL(qc)}{R[gd_key]}"
    # Annual rollup
    ws.cell(row=r, column=ann_col, value=f"=SUM({CL(qcols[0])}{r}:{CL(qcols[3])}{r})")
```

**Pass 3** (`build_segment_pl_pass3`): After Consolidated_PL is built, fills cross-check EBIT link:
```python
yr_to_ann_col = {2023: 2, 2024: 3, 2025: 4, 2026: 6, 2027: 7, 2028: 8}
for year, qcols, ann_col, suffix in GD_ALL_YEARS:
    ann_tab_col = yr_to_ann_col[year]
    ws.cell(row=R["segpl_xcheck_ebit"], column=ann_col,
            value=f"=Consolidated_PL!{CL(ann_tab_col)}{R['pl_ebit']}")
```

### Dual-Layer P&L with Cross-Validation

**Segment_PL**: costs by division (OI per segment = Revenue × OPM%)
**Consolidated_PL**: costs by type (COGS/R&D/MG&A)

Critical: **EBIT in Consolidated_PL = Sum of Segment OI from Segment_PL**

```python
# Consolidated_PL EBIT links to Segment_PL (segment is source of truth)
# Use SEGPL_COL_MAP because Consolidated_PL is annual, Segment_PL is quarterly
for col in HIST_COLS + FCST_COLS:
    _lval(ws, r, col, f"=Segment_PL!{CL(SEGPL_COL_MAP[col])}{R['segpl_sum_oi']}")

# Segment_PL cross-check row
for year, qcols, ann_col, suffix in GD_ALL_YEARS:
    _fval(ws, diff_row, ann_col, f"={CL(ann_col)}{sum_oi_row}-{CL(ann_col)}{ebit_row}")
# → must always = 0
```

### Consolidated_PL (Annual-Only)

Revenue by cost type. EBIT is segment-driven (linked from Segment_PL sum_oi). All references to Segment_PL use SEGPL_COL_MAP since Consolidated_PL uses annual columns (2,3,4,6,7,8) and Segment_PL uses quarterly annual rollup columns (6,11,16,22,27,32).

```python
# Revenue from Segment_PL consolidated total
for col in HIST_COLS + FCST_COLS:
    _lval(ws, r, col, f"=Segment_PL!{CL(SEGPL_COL_MAP[col])}{R['segpl_consol_rev']}")

# EBIT from Segment_PL sum of segment OI
for col in HIST_COLS + FCST_COLS:
    _lval(ws, r, col, f"=Segment_PL!{CL(SEGPL_COL_MAP[col])}{R['segpl_sum_oi']}")
```

### Balance Sheet — Cash as Plug (Annual)

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

### Cash Flow — WC from BS Delta (Annual)

```python
# AR change (increase = cash outflow)
f"=-(BS!{CL(c)}{ar_row}-BS!{prev}{ar_row})"
# AP change (increase = cash inflow)
f"=BS!{CL(c)}{ap_row}-BS!{prev}{ap_row}"
```

### DCF: UFCF → Enterprise Value → Equity Value (Annual)

```
UFCF = NOPAT + D&A − CapEx − ΔNWC
TV = UFCF_terminal × (1+g) / (WACC−g)    [Gordon Growth]
EV = Σ PV(UFCF) + PV(TV)
Equity = EV + Cash − Debt
Price = Equity / Diluted Shares
```

**v2**: WACC parameters (Rf, Beta, ERP, Kd, D/TC, Tax Rate, Terminal g) live on the DCF tab itself, not in Assumptions. This keeps Assumptions focused on operating drivers.

## Step 4: IB Formatting Standards

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
3. **Column headers** with dark fill + white font (annual headers in dark blue for annual rollup cols)
4. **Notes column** (last column) — prediction logic for every row
5. **Data legend** at bottom — color coding explanation
6. **Data sources** at bottom — specific page/note references

### Key Summary Tab (Executive Dashboard)

Links to all other tabs. Contains:
- Core KPIs (Revenue/GP/EBIT/NI + margins + YoY)
- Revenue by segment with mix %
- BS highlights (Total Assets/Equity/Cash)
- ROE / ROA
- DCF implied share price
- Investment highlights + risk warnings (bullet points)

All Segment_PL references use SEGPL_COL_MAP since Key_Summary is annual.

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

**Fix**: ALL percentage values MUST be stored as decimals:
```python
# WRONG — displays as 7200%
_hval(ws, r, col, 72.0, "0.0%")
# RIGHT — displays as 72.0%
_hval(ws, r, col, 0.72, "0.0%")
```

### 2. Notes column text starting with `=`

**Problem**: openpyxl treats any cell value starting with `=` as a formula. Text like `"=Revenue×TAC Rate"` becomes an invalid formula → Excel repair deletes it.

**Fix**: Strip leading `=` from Notes text.

### 3. Cross-Layout Column Mapping (Quarterly → Annual)

**Problem**: When quarterly tabs (35 cols) and annual tabs (11 cols) reference each other, column indices don't match. Writing `Segment_PL!B{r}` from an annual tab references Q1'23A (quarterly) instead of FY2023A (annual rollup).

**Fix**: Always use column mapping dicts:
```python
# Annual tab → Segment_PL quarterly annual rollup
SEGPL_COL_MAP = {2: 6, 3: 11, 4: 16, 6: 22, 7: 27, 8: 32}
# Annual tab fcst col → Assumptions quarterly annual rollup
ASP_ANN_COLS = {6: 22, 7: 27, 8: 32}
```

### 4. GD → Segment_PL Historical Links

When both tabs use the same quarterly layout, GD historical annual columns (6,11,16) map to the SAME Segment_PL columns (6,11,16) — NOT the old annual layout cols (2,3,4).

**Wrong**: `gd_ann_to_segpl = {6: 2, 11: 3, 16: 4}`
**Right**: `gd_ann_to_segpl = {6: 6, 11: 11, 16: 16}`

### 5. Quarterly Self-Check Column Ranges

When scanning quarterly tabs for formulas, use the quarterly forecast column range (18-32), not the old annual range (6-8):
```python
# WRONG for quarterly tabs
for row in ws.iter_rows(min_col=6, max_col=8): ...
# RIGHT for quarterly tabs
for row in ws.iter_rows(min_col=18, max_col=32): ...
```

### 6. Quarterly Tabs in Self-Check

Add all quarterly tabs to the notes-check exclusion set:
```python
quarterly_tabs = {"Growth_Drivers", "Assumptions", "Segment_PL"}
```

### 7. CHOOSE/MATCH with array constants

The formula `CHOOSE(MATCH($B$2,{"Base","Bull","Bear"},0),...)` uses inline array `{}`. This works in modern Excel. NOT a CSE formula — do NOT wrap in extra braces.

### 8. Cash plug formula — subtraction logic

**Wrong**: `f"=...−{CL(c)}{a}+{CL(c)}{b}+{CL(c)}{c_}"` (minus only applies to first term)
**Right**: `f"=...−{CL(c)}{a}−{CL(c)}{b}−{CL(c)}{c_}"` (explicit minus for each)

## Helper Function Pattern

```python
def _hval(ws, r, col, val, fmt="#,##0"):
    """Historical value — blue font"""
    c = ws.cell(row=r, column=col, value=val)
    c.number_format = fmt; c.font = Font(color="305496"); c.border = BD

def _fval(ws, r, col, formula, fmt="#,##0"):
    """Formula value — black font"""
    c = ws.cell(row=r, column=col, value=formula)
    c.number_format = fmt; c.border = BD

def _lval(ws, r, col, formula, fmt="#,##0"):
    """Cross-sheet link — green font"""
    c = ws.cell(row=r, column=col, value=formula)
    c.number_format = fmt; c.font = Font(color="548235"); c.border = BD

def _key(ws, r, col, formula, fmt="#,##0"):
    """Key output — yellow fill"""
    c = ws.cell(row=r, column=col, value=formula)
    c.number_format = fmt; c.fill = FILL_KEY; c.border = BD

def _pct(ws, r, col, formula):
    """Percentage formula — grey italic"""
    c = ws.cell(row=r, column=col, value=formula)
    c.number_format = "0.0%"; c.font = Font(size=9, italic=True, color="808080"); c.border = BD

def _label(ws, r, text, bold=False):
    """Row label in column A"""
    c = ws.cell(row=r, column=1, value=text)
    c.font = Font(bold=bold, size=10, name="Calibri"); c.border = BD

def _note(ws, r, text):
    """Notes column — strips leading '=' to prevent formula errors"""
    if text and text.startswith('='):
        text = text[1:]
    ws.cell(row=r, column=_ncol_for(ws), value=text)
```

## Multi-Pass Build Pattern

The model requires multiple passes to resolve circular dependencies:

```
Pass 1: Assumptions (all scenario data)
Pass 2: Segment_PL pass 1 (structure + historical, assign R keys)
Pass 3: Growth_Drivers (references Segment_PL R-keys for historical)
Pass 4: Segment_PL pass 2 (links forecast quarterly revenue to Growth_Drivers)
Pass 5: Consolidated_PL (references Segment_PL annual cols via SEGPL_COL_MAP)
Pass 6: Segment_PL pass 3 (fills cross-check EBIT link to Consolidated_PL)
Pass 7: BS → Cash_Flow → DCF → Sensitivity → Ratio_Analysis
Pass 8: Key_Summary (built last, links to everything)
```

## Quarterly Revenue Data (10-Q)

When quarterly segment revenue is available from 10-Q filings, store it for use in historical quarterly columns:

```python
GD_QREV = {
    "ccg": {
        2025: {1: 7629, 2: 8189, 3: 8190, 4: 8220},
        2026: {1: 7727},
    },
    "dcai": { ... },
    "foundry": { ... },
    "allother": { ... },
}
```

Missing quarters fall back to `annual / 4`.

## Verification Checklist

After generating the workbook, verify programmatically:

```python
import openpyxl
wb = openpyxl.load_workbook("output.xlsx")

# 1. Expected tabs present
# 2. Segment_PL cross-check = 0 (annual columns only)
# 3. BS balance check = 0
# 4. EBIT link: Consolidated_PL EBIT → Segment_PL Sum OI
# 5. No formula text in Notes column (quarterly tabs: col 35, annual tabs: col 11)
# 6. CHOOSE/MATCH formulas count matches expectation
# 7. Segment_PL forecast formulas reference Assumptions
# 8. DCF has implied share price
# 9. All "0.0%" format cells contain decimal values (0.xx, not xx.0)
# 10. Quarterly tabs: cross-check scans correct column ranges
```

### Self-Check Implementation Notes

```python
quarterly_tabs = {"Growth_Drivers", "Assumptions", "Segment_PL"}

# Notes column check: tab-specific
notes_col = GD_NCOL if sheet_name in quarterly_tabs else NCOL

# CHOOSE/MATCH: scan quarterly forecast range (18-32) for quarterly tabs
# Segment_PL Assumptions refs: scan cols 18-32 (quarterly forecast)
# Cross-check: scan annual rollup cols (6,11,16,22,27,32) for quarterly tabs
```

## CapEx Handling

When management provides CapEx guidance (e.g., "$75B per quarter"), use **absolute value inputs** in Assumptions, NOT percentage of revenue. This is more accurate for capital-intensive periods.

## External Data Sourcing for Growth Drivers

| Driver | Source | Notes |
|--------|--------|-------|
| PC TAM | Gartner / IDC | Quarterly PC shipment reports |
| Server TAM | IDC / Mercury Research | Server CPU shipment estimates |
| CPU Unit Share | Mercury Research | x86 CPU share by vendor |
| Global LV Production | IHS Markit / S&P Global Mobility | Annual light vehicle production |
| ASIC/AI revenue | Company earnings calls | Segment-level AI disclosures |

**Always cite data sources** in the model Notes column.

## Quick Start Template

For a new company (e.g., AMD, NVDA, QCOM):

1. Copy `build_intc_model.py` as the starting template
2. Replace `GD_QREV` with target company's quarterly segment data
3. Update `build_assumptions()` with target company's drivers:
   - Replace segment names and driver frameworks
   - Adjust Base/Bull/Bear values
4. Update `build_segment_pl()` — change segment list, hist_rev_ann, opm_key values
5. Update `build_growth_drivers()` — adjust driver formulas per segment
6. Update historical hardcoded values in `build_consolidated_pl()`, `build_bs()`, etc.
7. Run `build_model()` → verify self-check passes
8. Open in Excel, switch scenarios, confirm no #REF! errors

## Reference Files

- For a complete working example (quarterly layout), see [build_intc_model.py](build_intc_model.py)
- For the annual-only reference architecture, see [example-structure.md](example-structure.md)
- For a real project README, see [README.md](README.md)
