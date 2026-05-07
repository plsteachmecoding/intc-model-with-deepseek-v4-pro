# Intel (INTC) Investment Bank Financial Model

**Investment Bank-Grade Three-Statement DCF Model**
Segment-driven architecture | Bottom-up growth drivers | Scenario switching (Base / Bull / Bear) | 12 tabs

---

## Quick Start

### Prerequisites
```bash
pip install openpyxl
```

### Generate the Model
```bash
python build_intc_model_final.py
```

This produces `Intel_IB_Model_v2.xlsx` (~913 rows across 12 tabs).

### Open in Excel
1. Open `Intel_IB_Model.xlsx`
2. Navigate to the **Assumptions** tab
3. Use the dropdown in cell **B2** to switch scenarios: `Base` / `Bull` / `Bear`
4. All tabs recalculate automatically
5. Key output: **DCF tab → Implied Share Price** (cell H48)
6. Explore **Growth_Drivers** tab to see bottom-up revenue decomposition

---

## Tab Guide

| Tab | What It Does | Edit? |
|-----|-------------|-------|
| **Cover** | Company info, version, usage instructions | Read only |
| **Key_Summary** | Executive dashboard — all KPIs linked | Read only |
| **Assumptions** | ALL model inputs — growth rates, margins, WACC, CapEx, **bottom-up drivers** | **EDIT HERE** |
| **Growth_Drivers** | Bottom-up revenue decomposition: TAM × Share × ASP per segment | Read only |
| **Segment_Revenue** | Revenue by segment (CCG/DCAI/Foundry/All Other), linked to Growth_Drivers | Read only |
| **Segment_PL** | Operating income by division | Read only |
| **Consolidated_PL** | P&L by cost type → EBIT → NI → EPS | Read only |
| **BS** | Balance sheet (Cash = plug) | Read only |
| **Cash_Flow** | CFO / CFI / FCF | Read only |
| **DCF** | Unlevered DCF → Implied Share Price | Read only |
| **Sensitivity** | WACC × Terminal Growth matrix | Excel Data Table |
| **Ratio_Analysis** | Profitability / efficiency / leverage ratios + Valuation Multiples (TTM & NTM) | Read only* |

## Color Coding (IB Standard)

- **Blue font** = Historical data (10-K actuals)
- **Green font** = Cross-sheet formula link
- **Blue fill (light)** = Editable assumption — **you can change these**
- **Yellow fill** = Key output row
- **Black font** = Calculated formula
- **Grey italic** = Percentage / memo row

## Scenario Switching

The model supports three scenarios:

1. Click cell **B2** in the **Assumptions** tab
2. Select from dropdown: `Base`, `Bull`, or `Bear`
3. All 79 assumption rows switch via `CHOOSE(MATCH(...))` formulas
4. Entire model recalculates — revenue, margins, BS, CF, DCF valuation

### Scenario Summary

| Scenario | Thesis |
|----------|--------|
| **Base** | Moderate PC recovery, DCAI CPU renaissance (+18% FY26), Foundry losses narrowing (−42%→−22% OPM), CapEx $17.7–21.0B |
| **Bull** | Strong DCAI (+25% FY26), faster Foundry external ramp, 18A yields exceed expectations, CapEx scales to $25B |
| **Bear** | PC recession (−3% FY26 CCG), AMD share gains, persistent supply constraints, Foundry losses remain deep (−50%→−40%) |

## Model Architecture

### Revenue Bridge (Important)
Intel Foundry's reported segment revenue (~$17.8B FY2025) is **~99% internal** wafer sales to Intel Products (CCG/DCAI). The model:

- Tracks **Total Foundry Revenue** for Segment_PL OI calculation
- Adds **External Foundry Revenue** row (growing from ~$150M to ~$2.5B by FY2028E)
- **Consolidated Revenue** = CCG + DCAI + All Other + External Foundry + Eliminations
- Historical consolidated revenue hardcoded from 10-K

### Growth_Drivers — Bottom-Up Revenue Decomposition
Each segment's revenue forecast is built from granular drivers, NOT simple growth rates:

| Segment | Driver Formula | Key Parameters |
|---------|---------------|----------------|
| **CCG** | PC TAM × Intel Share × ASP | PC TAM (Gartner), Share (Mercury), ASP (derived) |
| **DCAI** | Server TAM × DC Share × ASP + AI/ASIC | Server TAM (IDC), Share (Mercury), AI/ASIC (earnings calls) |
| **Foundry** | Chip Vol × Rev/Chip + External | Chip Vol = CCG + DCAI units, Rev/Chip (internal transfer price) |
| **All Other** | LV Prod × MBLY Rev/Veh + IMS | LV Prod (IHS/S&P), MBLY Rev/Veh (derived), IMS (earnings calls) |

**Two-pass build**: Segment_Revenue pass 1 (historical + structure) → Growth_Drivers → Segment_Revenue pass 2 (links to Growth_Drivers implied revenue). A **Driver Sensitivity** section at the bottom shows revenue elasticity for each key driver.

### Cash Plug Mechanism
The Balance Sheet uses **Cash as the plug** to ensure A = L+E:
```
Cash = Total L&E − Non-Current Assets − Other Current Assets
```
A balance check row (`Total Assets − Total L&E`) must equal zero — verified programmatically.

### EBIT Linkage
```
Segment_PL (OI by division) → Consolidated_PL EBIT
Sum of Segment OI = CCG_OI + DCAI_OI + Foundry_OI + All_Other_OI
Cross-check row verifies: Sum Seg OI − Consolidated EBIT = 0
```

## Verification

Run the built-in verification:
```bash
python verify_model.py
```

Checks: 12 tabs, cross-checks, BS balance, EBIT link, CF WC links, DCF UFCF, Notes safety, Growth_Drivers reconciliation.

Additional inspection scripts (preserved for audit):
- `_verify_revenue.py` — Revenue bridge validation
- `_verify_bs2.py` — Balance sheet totals and balance check
- `_check_cash.py` — Cash plug formula inspection
- `_compute_outputs.py` — Manual Base case computation

## File Inventory

| File | Purpose |
|------|---------|
| `build_intc_model.py` | Model generator script (~3,210 lines, **12 tabs with Growth_Drivers**) |
| `build_intc_model_final.py` | Earlier version (~2,545 lines, 11-tab) — preserved for reference |
| `Intel_IB_Model_v2.xlsx` | Generated Excel model (12 tabs, 913 rows) |
| `Intel_IB_Model.xlsx` | Earlier 11-tab version — preserved for reference |
| `model_summary.md` | Detailed model summary with assumptions & outputs |
| `progress.md` | Build progress log — all phases documented |
| `verify_model.py` | Automated verification script |
| `intc_financial_data.md` | Extracted 10-K financial data |
| `intc_assumptions.md` | Assumptions design document |
| `SKILL-ib-financial-model.md` | Updated skill definition (12-tab architecture) |
| `example-structure.md` | Reference architecture template |

## Key Limitations

1. **Forecast horizon**: 3 years (FY2026E–FY2028E). Capital-intensive turnaround may require 5+ years for CapEx payback.
2. **DCF sensitivity**: High debt load (~$46.6B FY2025A) means equity value is highly sensitive to WACC and terminal growth assumptions.
3. **External Foundry revenue**: Currently immaterial (~$150M FY2025). Forecast growth to $2.5B by FY2028E is speculative.
4. **Tax rate**: Low ETR (~12%) reflects valuation allowance — may normalize to ~21% as profitability returns.
5. **Non-operating items**: Government agreements (USG Agreement), restructuring charges, and NCI simplifications.

## Regenerate

To regenerate the model from scratch:
```bash
python build_intc_model_final.py
```

To modify assumptions or structure, edit `build_intc_model_final.py` and re-run.

---

*Built with Python openpyxl | Intel 10-K/Q data | May 7, 2026*
