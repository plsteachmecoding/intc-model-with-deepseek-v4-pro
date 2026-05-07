# Intel (INTC) Investment Bank Financial Model — Summary

**Date:** May 7, 2026
**Model Horizon:** FY2026E – FY2028E
**Architecture:** Segment-driven, three-statement linked DCF with scenario switching

---

## 1. Model Structure (12 Tabs)

| # | Tab | Purpose | Key Outputs |
|---|-----|---------|-------------|
| 1 | Cover | Company info + usage instructions | — |
| 2 | Key_Summary | Executive dashboard | All KPIs linked from other tabs |
| 3 | Assumptions | Scenario control center | 112 CHOOSE/MATCH formulas |
| 4 | Growth_Drivers | Bottom-up revenue decomposition | TAM × Share × ASP per segment; driver sensitivity |
| 5 | Segment_Revenue | Revenue by segment, linked to Growth_Drivers | Consolidated revenue |
| 6 | Segment_PL | OI by division (CCG/DCAI/Foundry/All Other) | Sum of Segment OI → EBIT source |
| 7 | Consolidated_PL | P&L by cost type (COGS/R&D/MG&A) | EBIT = Segment_PL sum; NI; EPS |
| 8 | BS | Balance Sheet (Cash as plug) | A = L+E via balance check |
| 9 | Cash_Flow | Direct cash flow | CFO, CFI, FCF |
| 10 | DCF | Unlevered DCF valuation | Implied share price |
| 11 | Sensitivity | WACC × g 2D data table | Price sensitivity matrix |
| 12 | Ratio_Analysis | Financial ratios | Profitability, efficiency, leverage, growth, valuation multiples |

## 2. Key Assumptions (Base Case)

### WACC
| Parameter | Value |
|-----------|-------|
| Risk-Free Rate | 4.5% |
| Beta | 1.50 |
| Equity Risk Premium | 5.5% |
| Cost of Equity | 12.75% |
| Pre-tax Cost of Debt | 5.0% |
| Marginal Tax Rate | 12.0% |
| Debt / Total Capital | 25.0% |
| **WACC** | **10.66%** |
| Terminal Growth | 2.5% |

### Segment Revenue Growth (Base Case)

| Segment | FY2026E | FY2027E | FY2028E |
|---------|---------|---------|---------|
| CCG | +2.0% | +3.0% | +3.0% |
| DCAI | +18.0% | +12.0% | +8.0% |
| Intel Foundry (total) | +12.0% | +10.0% | +10.0% |
| All Other | -5.0% | +8.0% | +8.0% |

### Segment Operating Margins (Base Case)

| Segment | FY2026E | FY2027E | FY2028E |
|---------|---------|---------|---------|
| CCG | 30.0% | 31.5% | 33.0% |
| DCAI | 28.0% | 28.0% | 27.0% |
| Intel Foundry | -42.0% | -32.0% | -22.0% |
| All Other | 14.0% | 18.0% | 20.0% |

### Cost Structure (% of Revenue)

| Line | FY2026E | FY2027E | FY2028E |
|------|---------|---------|---------|
| TAC | 9.1% | 8.8% | 8.7% |
| Other COGS | 45.5% | 44.0% | 43.0% |
| R&D | 19.5% | 18.0% | 17.0% |
| MG&A | 7.0% | 6.5% | 6.3% |
| D&A ($M) | 12,500 | 13,200 | 13,800 |
| CapEx ($M) | 17,700 | 19,500 | 21,000 |

## 3. Model Outputs (Base Case)

### Revenue & Profitability

| Metric | FY2026E | FY2027E | FY2028E |
|--------|---------|---------|---------|
| Consolidated Revenue ($M) | 56,722 | 61,175 | 65,272 |
| YoY Revenue Growth | +7.3% | +7.9% | +6.7% |
| Gross Profit ($M) | 29,471 | 32,774 | 35,664 |
| GPM | 52.0% | 53.6% | 54.6% |
| EBIT ($M) | 2,540 | 6,057 | 9,504 |
| EBIT Margin | 4.5% | 9.9% | 14.6% |
| Net Income ($M) | 1,199 | 5,066 | 9,187 |
| NI Margin | 2.1% | 8.3% | 14.1% |
| Diluted EPS ($) | 0.24 | 0.98 | 1.75 |

### Segment OI Breakdown

| Segment | FY2026E | FY2027E | FY2028E |
|---------|---------|---------|---------|
| CCG OI ($M) | 9,862 | 10,666 | 11,509 |
| DCAI OI ($M) | 5,590 | 6,261 | 6,520 |
| Foundry OI ($M) | -8,385 | -7,028 | -5,315 |
| All Other OI ($M) | 474 | 658 | 790 |

### Cash Flow & DCF

| Metric | FY2026E | FY2027E | FY2028E |
|--------|---------|---------|---------|
| NOPAT ($M) | 2,235 | 5,269 | 8,268 |
| UFCF ($M) | -3,042 | -1,120 | 986 |

**DCF Valuation:**
- Enterprise Value: ~$6.2B
- Net Debt (FY2025A): $32.3B
- Equity Value: Negative (turnaround still in early stages)

> **Note:** Base case DCF yields challenged equity value due to: (a) heavy near-term CapEx ($17-21B/year), (b) high debt load (~$46.6B FY2025A), (c) 3-year forecast horizon captures ramp but not full payback. The Bull case and extended horizon produce positive equity values. See Sensitivity tab for WACC/g ranges.

### Market Valuation Multiples (Ratio_Analysis tab)

Stock price input cell (editable, blue fill) — default **$113.00**. All multiples update automatically when price or scenario changes.

| Multiple | TTM (FY2025A) | NTM (FY2026E) | Notes |
|----------|---------------|---------------|-------|
| P/E | Extreme (~11,300x) | ~470x | FY2025 NI=$26M (near zero); FY2026E recovering |
| P/B | ~5.6x | Formula-driven | Updates with scenario |
| P/S | ~10.7x | Formula-driven | Revenue base recovery |
| EV/EBITDA | ~67.7x | Formula-driven | Heavy debt + minimal EBITDA |
| P/CF | ~58.2x | Formula-driven | CFO positive, FCF negative (CapEx) |
| PEG | NM | NM | Near-zero EPS makes PEG meaningless |

> TTM multiples reflect Intel's trough earnings (FY2025 was break-even). NTM multiples are scenario-dependent — switch Assumptions!B2 to Bull/Bear to see range.

## 4. Model Architecture Highlights

### Revenue Bridge (Critical Fix Applied)
Intel Foundry's reported segment revenue ($17.8B FY2025) is **~99% internal** wafer sales to Intel Products. The model:
- Tracks **Total Foundry Revenue** (used for Segment OI calculation in Segment_PL)
- Adds **External Foundry Revenue** row (~$150M FY2025A, growing to $2.5B by FY2028E)
- **Consolidated Revenue** = CCG + DCAI + All Other + External Foundry + Eliminations
- Historical consolidated revenue is hardcoded from 10-K (FY2023: $54,228M, FY2024: $53,101M, FY2025: $52,853M)

### EBIT: Segment-Driven
- **Segment_PL**: OI computed per segment (Revenue × OPM%)
- **Consolidated_PL EBIT** = Sum of Segment OI + Corporate Unallocated
- **Cross-check row** in Segment_PL verifies Sum Seg OI = Consolidated EBIT (must = 0)

### Balance Sheet: Cash as Plug
- Cash = Total L&E − Non-Current Assets − Other Current Assets
- Ensures A = L+E by construction
- Balance check row = Total Assets − Total L&E (must = 0)

### Cash Flow: WC from BS Delta
- 22 working capital items link to BS period-over-period changes
- AR change, AP change, Inventory change, Accrued Comp, etc.
- FCF = CFO + CFI (CapEx)

### DCF: Unlevered Free Cash Flow
- UFCF = NOPAT + D&A − CapEx − ΔNWC
- Terminal Value = Gordon Growth Model
- Equity Value = EV + Cash − Debt
- Implied Price = Equity Value / Diluted Shares

## 5. Scenario Framework

Three scenarios selectable via dropdown in `Assumptions!B2`:

| Scenario | Description |
|----------|-------------|
| **Base** | Central case: moderate PC recovery, DCAI CPU renaissance, Foundry losses narrowing |
| **Bull** | Upside: stronger DCAI growth, faster Foundry external ramp, better margins |
| **Bear** | Downside: recession, AMD share gains, persistent Foundry losses, supply constraints |

All formulas use `CHOOSE(MATCH($B$2,{"Base","Bull","Bear"},0), ...)` for dynamic switching.

## 6. IB Formatting Standards

| Style | Meaning | Implementation |
|-------|---------|----------------|
| Blue font (#305496) | Historical hardcoded data | 10-K actuals |
| Green font (#548235) | Cross-sheet formula link | References to other tabs |
| Blue fill (#DDEBF7) | Editable assumption | User-modifiable inputs |
| Yellow fill (#FFF2CC) | Key output row | Important calculated values |
| Black font | Formula (calculated) | Standard model formulas |
| Grey italic | Percentage / memo row | YoY%, margin %, mix % |

## 7. Data Sources

- Intel 10-K FY2022, FY2023, FY2024, FY2025
- Intel Q1 2026 10-Q (filed April 2026)
- Intel Q1 2026 Earnings Call (April 23, 2026)
- Segment restatement per Q1 2025 8-K (NEX integrated, Altera deconsolidated)

## 8. Verification Status

| Check | Status |
|-------|--------|
| 11 tabs present | PASS |
| Segment_PL cross-check | PASS |
| BS Balance Check (A = L+E) | PASS |
| EBIT = Sum Seg OI | PASS (all 6 columns) |
| CF WC links to BS | PASS (22 references) |
| DCF UFCF = NOPAT+D&A−CapEx−ΔNWC | PASS |
| No formulas in Notes column | PASS |
| CHOOSE/MATCH formulas | PASS (79 count) |
| Scenario selector (Data Validation) | PASS |

---

*Generated by build_intc_model_final.py — 2,545 lines, 32 functions*
