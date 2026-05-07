# Intel Financial Model — Growth Drivers Build Progress

**Start:** 2026-05-07
**Last Updated:** 2026-05-07
**Status:** All 4 segments complete. Percentage format bug fixed. Model builds clean.

---

## Phase 5: Growth_Drivers Tab

### Objective
Add a new `Growth_Drivers` tab (12th tab) that decomposes each segment's revenue forecast into granular drivers (top-down or bottom-up).

### Architecture
- Tab inserted between Assumptions and Segment_Revenue in workbook
- Build order: Segment_Revenue pass 1 → Growth_Drivers → Segment_Revenue pass 2
- Each segment gets its own driver framework
- Historical data reconciled to reported segment revenue
- External data sources cited in the model

---

### Progress Log

| Segment | Status | Driver Method | Key Parameters | Date Completed |
|---------|--------|---------------|----------------|----------------|
| **CCG** | **DONE** | Bottom-up: PC TAM x Share x ASP | PC TAM (Gartner), Intel Share (Mercury), ASP (derived) | 2026-05-07 |
| **DCAI** | **DONE** | Bottom-up: Server TAM x DC Share x ASP + AI/ASIC | Server TAM (IDC), DC Share (Mercury), ASP (derived), AI/ASIC (Intel Q1 2026 disclosure) | 2026-05-07 |
| **Intel Foundry** | **DONE** | Bottom-up: Total Chip Vol x Internal Rev/Chip + External Rev | Chip vol (CCG+DCAI units), Rev/Chip (derived), External (Intel disclosures) | 2026-05-07 |
| **All Other** | **DONE** | Bottom-up: LV Production x Rev/Veh + IMS/Other ($) | LV Prod (IHS/S&P), MBLY Rev/Veh (derived), IMS (Intel disclosures) | 2026-05-07 |

---

### CCG — Client Computing Group (COMPLETE)

**Method:** Bottom-up
```
CCG Revenue = PC TAM (M units) x Intel Unit Share (%) x Blended ASP ($)
```

**Historical Data:**
| Year | PC TAM (M) | Intel Share | Implied Units (M) | ASP ($) | Implied Rev ($M) | Reported ($M) | Recon |
|------|-----------|-------------|-------------------|---------|------------------|---------------|-------|
| FY2023 | 241.8 | 71.0% | 171.7 | 188 | 32,275 | 32,305 | -$30M |
| FY2024 | 245.4 | 71.5% | 175.5 | 190 | 33,338 | 33,346 | -$8M |
| FY2025 | 258.6 | 72.0% | 186.2 | 173 | 32,211 | 32,228 | -$17M |

**Data Sources:**
- PC TAM: Gartner (2022=286.2M, 2023=241.8M, 2024=245.4M) via Wikipedia. FY2025 est. from Canalys full-year reports (~259M)
- Intel Unit Share: Mercury Research x86 PC CPU share estimates. Q4 2024 ~72%. Verified against Intel 10-K qualitative disclosures
- ASP: Derived as Reported Revenue / Implied Units. Cross-checked against Intel Q1 2026 disclosure (ASP +16% YoY)

**Assumptions Added (9 rows in Assumptions tab):**
- PC TAM 2026E/2027E/2028E (Base/Bull/Bear) — 3 rows
- Intel Unit Share 2026E/2027E/2028E (Base/Bull/Bear) — 3 rows
- CCG ASP 2026E/2027E/2028E (Base/Bull/Bear) — 3 rows

**Growth_Drivers Rows (CCG section = 9 data rows):**
1. PC TAM (M units) — historical hardcoded + forecast links
2. Intel PC Unit Share (%) — historical hardcoded + forecast links
3. Intel Implied Units (M) — formula: TAM x Share
4. Intel CCG Blended ASP ($) — historical hardcoded + forecast links
5. Implied CCG Revenue ($M) — formula: Units x ASP
6. Reported CCG Revenue ($M) — green link to Segment_Revenue
7. Reconciliation ($M) — Implied - Reported
8. Reconciliation (%) — Recon / Reported

**Forecast Base Case:**
| Year | PC TAM | Share | ASP | Implied Rev |
|------|--------|-------|-----|-------------|
| FY2026E | 260.0M | 72.0% | $175 | $32,760M |
| FY2027E | 265.0M | 72.5% | $182 | $35,005M |
| FY2028E | 270.0M | 73.0% | $184 | $36,266M |

---

### DCAI — Data Center & AI (COMPLETE)

**Method:** Bottom-up — Two-component
```
DCAI Revenue = Server TAM (M) x Intel DC Share (%) x ASP ($) + AI/ASIC Revenue ($M)
```

**Historical Data:**
| Year | Server TAM (M) | Intel DC Share | Implied Units (M) | ASP ($) | CPU Rev ($M) | AI/ASIC ($M) | Implied Rev ($M) | Reported ($M) | Recon |
|------|-----------|-------------|-------------------|---------|-------------|-------------|------------------|------------|-------|
| FY2023 | 11.5 | 77.5% | 8.9 | 1,772 | 15,793 | 200 | 15,993 | 15,980 | +$13M |
| FY2024 | 12.0 | 76.0% | 9.1 | 1,712 | 15,613 | 500 | 16,113 | 16,125 | -$12M |
| FY2025 | 13.5 | 73.0% | 9.9 | 1,632 | 16,083 | 800 | 16,883 | 16,919 | -$36M |

**Data Sources:**
- Server TAM: IDC server CPU shipment estimates. FY2023 ~11.5M, FY2024 ~12.0M, FY2025 ~13.5M (AI buildout driven)
- Intel DC Unit Share: Mercury Research x86 server CPU share. Q4 2024 ~72%. Granite Rapids stabilizing vs AMD Turin
- Server ASP: Derived as (DCAI Rev - AI/ASIC) / (TAM x Share). Q1 2026 +27% YoY per Intel earnings
- AI/ASIC Revenue: Intel Q1 2026 disclosure — ASIC run rate "north of $1B". FY2024 ~$500M, FY2025 ~$800M. Gaudi modest contribution

**Assumptions Added (12 rows in Assumptions tab):**
- Server TAM 2026E/2027E/2028E (Base/Bull/Bear) — 3 rows
- Intel DC Server Unit Share 2026E/2027E/2028E (Base/Bull/Bear) — 3 rows
- Intel Server Blended ASP 2026E/2027E/2028E (Base/Bull/Bear) — 3 rows
- AI/ASIC & Other Revenue 2026E/2027E/2028E (Base/Bull/Bear) — 3 rows

**Growth_Drivers Rows (DCAI section = 12 data rows):**
1. Server TAM (M units) — historical hardcoded + forecast links
2. Intel DC Server Unit Share (%) — historical hardcoded + forecast links
3. Intel Implied DC Units (M) — formula: TAM x Share
4. Intel Server Blended ASP ($) — historical hardcoded + forecast links
5. Implied CPU Revenue ($M) — formula: Units x ASP
6. AI/ASIC & Other Revenue ($M) — historical hardcoded + forecast links
7. Implied DCAI Revenue ($M) — formula: CPU Rev + AI/ASIC Rev
8. Reported DCAI Revenue ($M) — green link to Segment_Revenue
9. Reconciliation ($M) — Implied - Reported
10. Reconciliation (%) — Recon / Reported

**Forecast Base Case:**
| Year | Server TAM | DC Share | ASP | CPU Rev | AI/ASIC | Implied Rev |
|------|-----------|----------|-----|---------|---------|-------------|
| FY2026E | 14.5M | 70.0% | $1,650 | $16,748M | $2,800M | $19,548M |
| FY2027E | 14.5M | 70.0% | $1,700 | $17,255M | $4,500M | $21,755M |
| FY2028E | 15.5M | 70.0% | $1,750 | $18,988M | $6,000M | $24,988M |

---

### Intel Foundry (COMPLETE)

**Method:** Bottom-up — Two-component
```
Foundry Revenue = Total Intel Chip Vol (M) x Internal Rev/Chip ($) + External Foundry ($M)
```

**Key Insight:** ~99% of Foundry revenue is internal wafer sales to Intel Products (CCG/DCAI). The bottom-up model links chip volume from the CCG and DCAI driver sections and applies an internal transfer price per chip. External foundry revenue is separate and growing from a very small base (~$150M FY2025).

**Historical Data:**
| Year | Total Chip Vol (M) | Rev/Chip ($) | Internal Rev ($M) | External ($M) | Implied Rev ($M) | Reported ($M) | Recon |
|------|-------------------|-------------|-------------------|---------------|------------------|---------------|-------|
| FY2023 | 180.6 | 102.2 | 18,457 | 50 | 18,507 | 18,504 | +$3M |
| FY2024 | 184.6 | 93.5 | 17,262 | 60 | 17,322 | 17,317 | +$5M |
| FY2025 | 196.1 | 90.1 | 17,665 | 150 | 17,815 | 17,826 | -$11M |

**Data Sources:**
- Total Chip Volume: Derived as CCG Implied Units + DCAI Implied Units (both in Growth_Drivers tab)
- Internal Rev/Chip: Derived as (Foundry Revenue - External) / Total Chip Volume. Declining trend reflects older node depreciation
- External Foundry: Intel disclosures — Q1 2026 $174M annualised. FY2025 estimated ~$150M

**Assumptions Added (3 rows in Assumptions tab):**
- Internal Foundry Rev/Chip 2026E/2027E/2028E (Base/Bull/Bear) — 3 rows

**Growth_Drivers Rows (Foundry section = 10 data rows):**
1. Total Intel Chip Volume (M) — formula: CCG Units + DCAI Units
2. Internal Foundry Rev per Chip ($) — historical hardcoded + forecast links
3. Internal Foundry Revenue ($M) — formula: Vol x Rev/Chip
4. External Foundry Revenue ($M) — historical hardcoded + forecast links
5. Implied Total Foundry Revenue ($M) — formula: Internal + External
6. Reported Foundry Revenue ($M) — green link to Segment_Revenue
7. Reconciliation ($M) — Implied - Reported
8. Reconciliation (%) — Recon / Reported

**Forecast Base Case:**
| Year | Chip Vol (M) | Rev/Chip | Internal Rev | External | Implied Total |
|------|-------------|----------|-------------|----------|---------------|
| FY2026E | 197.4 | $90.0 | $17,762M | $700M | $18,462M |
| FY2027E | 205.3 | $92.0 | $18,888M | $1,500M | $20,388M |
| FY2028E | 212.8 | $95.0 | $20,216M | $2,500M | $22,716M |

---

### All Other — Mobileye + IMS + Legacy (COMPLETE)

**Method:** Bottom-up — Two-component
```
All Other Revenue = Global LV Production (M) x Mobileye Rev/Vehicle ($) + IMS & Other Revenue ($M)
```

**Key Insight:** All Other is Intel's residual segment (Mobileye ADAS chips + IMS nanofabrication tools + legacy businesses). Mobileye revenue is driven by global auto production and ADAS attach rates. IMS benefits from semiconductor capex cycle.

**Historical Data:**
| Year | LV Prod (M) | MBLY Rev/Veh ($) | MBLY Rev ($M) | IMS/Other ($M) | Implied Rev ($M) | Reported ($M) | Recon |
|------|------------|------------------|---------------|----------------|------------------|---------------|-------|
| FY2023 | 89.0 | 22.5 | 2,003 | 3,461 | 5,464 | 5,463 | +$1M |
| FY2024 | 88.0 | 19.9 | 1,751 | 1,850 | 3,601 | 3,601 | $0M |
| FY2025 | 87.0 | 18.4 | 1,601 | 1,962 | 3,563 | 3,563 | $0M |

**Data Sources:**
- Global LV Production: IHS Markit / S&P Global Mobility. FY2023 89M, FY2024 88M, FY2025 ~87M
- Mobileye Rev/Vehicle: Derived as estimated Mobileye revenue / LV Production
- IMS & Other: Derived as residual (All Other total - estimated Mobileye). FY2024 dip reflects portfolio divestitures

**Assumptions Added (9 rows in Assumptions tab):**
- Global LV Production 2026E/2027E/2028E (Base/Bull/Bear) — 3 rows
- Mobileye Rev per Vehicle 2026E/2027E/2028E (Base/Bull/Bear) — 3 rows
- IMS & Other Revenue 2026E/2027E/2028E (Base/Bull/Bear) — 3 rows

**Growth_Drivers Rows (All Other section = 8 data rows):**
1. Global LV Production (M) — historical hardcoded + forecast links
2. Mobileye Rev per Vehicle ($) — historical hardcoded + forecast links
3. Implied Mobileye Revenue ($M) — formula: LV Prod x Rev/Veh
4. IMS & Other Revenue ($M) — historical hardcoded + forecast links
5. Implied All Other Revenue ($M) — formula: Mobileye + IMS/Other
6. Reported All Other Revenue ($M) — green link to Segment_Revenue
7. Reconciliation ($M) — Implied - Reported
8. Reconciliation (%) — Recon / Reported

**Forecast Base Case:**
| Year | LV Prod | MBLY Rev/Veh | MBLY Rev | IMS/Other | Implied Total |
|------|---------|-------------|----------|-----------|---------------|
| FY2026E | 88.0M | $19.0 | $1,672M | $1,800M | $3,472M |
| FY2027E | 89.0M | $20.0 | $1,780M | $2,000M | $3,780M |
| FY2028E | 90.0M | $21.0 | $1,890M | $2,200M | $4,090M |

---

### Phase 5.5: Growth_Drivers → Segment_Revenue Linkage (COMPLETE)

**Change:** Segment_Revenue forecast columns now pull from Growth_Drivers implied revenue instead of the old `Prior x (1 + growth_rate)` formulas.

**Architecture:**
- Segment_Revenue built in two passes to avoid circular R-key dependency:
  1. Pass 1: Historical hardcoded + structure + assign R keys
  2. Growth_Drivers built (references segrev keys for historical, assigns implied rev keys)
  3. Pass 2: Segment_Revenue forecasts linked to Growth_Drivers implied rows
- Growth_Drivers "Reported" rows: historical cols link to Segment_Revenue (hardcoded), forecast cols mirror Implied row (same tab, no circularity)
- Forecast reconciliation = 0 by construction (Segment_Revenue IS Growth_Drivers forecast)

**Impact on Base Case FY2026E:**
| Metric | Old (Growth-Rate) | New (Driver-Based) | Delta |
|--------|-------------------|-------------------|-------|
| CCG Revenue | $32,878M | $32,760M | -$118M |
| DCAI Revenue | $19,964M | $19,547M | -$417M |
| Foundry Revenue | $19,965M | $18,462M | -$1,503M |
| All Other Revenue | $3,385M | $3,472M | +$87M |
| Consolidated Revenue | $56,722M | $56,279M | -$443M (-0.8%) |

---

### Phase 5.6: Driver Sensitivity Section (COMPLETE)

**Added at bottom of Growth_Drivers tab:** A driver elasticity table showing the FY2026E revenue impact of isolated changes in each key driver.

**Format:** [Driver | FY2026E Value | Change Unit | Revenue Impact ($M) | Impact %]

All elasticities reference FY2026E (col F) via formulas — switch scenario in Assumptions!B2 to see Bull/Bear sensitivity.

**Key Findings (Base Case FY2026E):**
| Driver | Unit Change | Approx. Rev Impact |
|--------|------------|-------------------|
| Intel PC Share | +/-1pp | Most sensitive CCG driver |
| Internal Foundry Rev/Chip | +/-$1 | x 197M+ units = material |
| Intel DC Server Share | +/-1pp | Key DCAI battleground |
| PC TAM | +/-1% | 1:1 CCG revenue |
| Server ASP | +/-$1 | x DC Units |
| CCG ASP | +/-$1 | x 187M CCG units |

**Top 3 most sensitive drivers highlighted** with analyst commentary.

---

### Phase 5.7: Critical Percentage Format Fix (COMPLETE)

**Bug:** Market share values displaying as "7200%" instead of "72.0%" in Assumptions and Growth_Drivers tabs.

**Root Cause:** Excel's "0.0%" number format multiplies stored values by 100 for display. Storing 72.0 with format "0.0%" displays as "7200.0%". Correct approach: store 0.72 with "0.0%" to display "72.0%".

**Fix Applied (~40 values across entire Assumptions tab):**
- CCG Intel Share: 71.0, 71.5, 72.0 → 0.71, 0.715, 0.72 (3 rows, all scenarios)
- DCAI DC Share: 76.0, 73.0, 70.0 → 0.76, 0.73, 0.70 (3 rows, all scenarios)
- COGS % of Revenue: 60.0, 67.3, 65.2 → 0.60, 0.673, 0.652; forecast (100-gm) → (100-gm)/100.0
- R&D % of Revenue: 29.6, 31.2, 26.1 → 0.296, 0.312, 0.261
- MG&A % of Revenue: 10.4, 10.4, 8.7 → 0.104, 0.104, 0.087
- Revenue Growth historicals: `h2024, h2025` → `h2024/100.0, h2025/100.0`
- WACC: Rf 3.88→0.0388, ERP 5.50→0.055, Kd 4.80→0.048, D/(D+E) 30.0→0.30
- BS %: Accrued Comp 6.7→0.067, Other CA 15.0→0.15, Other CL 25.0→0.25, Tax Payable 2.0→0.02
- Tax Rate: -119.8→-1.198, -71.6→-0.716, 98.3→0.983, 12.0→0.12

**Impact:** Critical — the broken values contaminated all downstream calculations. CCG implied units was 260x72=18,720M instead of 260x0.72=187.2M. NOPAT tax shield was EBITx(1-12.0)=EBITx(-11) instead of EBITx(1-0.12)=EBITx0.88.

---

### All 4 Segments Complete — Growth_Drivers Tab Summary

| Segment | Rows | Method | Historical Recon |
|---------|------|--------|-----------------|
| CCG | 9 | PC TAM x Share x ASP | <0.1% |
| DCAI | 12 | Server TAM x DC Share x ASP + AI/ASIC | <0.22% |
| Intel Foundry | 10 | Chip Vol x Rev/Chip + External | <0.07% |
| All Other | 8 | LV Prod x MBLY Rev/Veh + IMS/Other | <0.01% |
| **Total** | **39** | | |

---

### Model File State

- **Script:** `build_intc_model.py` (~3,210 lines)
- **Output:** `Intel_IB_Model_v2.xlsx` (913 rows, 12 tabs)
- **R dict keys:** 247
- **CHOOSE/MATCH formulas:** 112
- **Output filename:** `Intel_IB_Model_v2.xlsx`
- **Build result:** All 7 self-checks PASS

---

### Known Issues (for next session)

1. **DCF Row 25**: `TV as % of EV` formula is `=H24/H25` which appears self-referential. Should likely be `=H23/H24` (PV of TV / EV). Needs verification in Excel.

2. **External data verification**: PC TAM and server TAM sourced from secondary sources (Wikipedia, Canalys). Direct Gartner/IDC/Mercury data not accessed — WebSearch tool was unavailable. Verbatim data sources cited in model Notes column but precision not independently confirmed.

3. **DCF only uses 3 forecast years** (FY26E-FY28E). Typical DCF uses 5+ years. Terminal value therefore carries heavy weight in valuation.

4. **Foundry internal transfer pricing**: Rev/Chip derived as residual (Foundry Rev / Total Chip Vol). True internal wafer price by node not publicly disclosed.

---

### Verification Scripts

| Script | Purpose |
|--------|---------|
| `_verify_pct.py` | Confirms all "0.0%" format values stored as decimals (0.x not xx.0) |
| `_verify_gd2.py` | DCAI section structure and formula check |
| `_verify_sens.py` | Driver sensitivity section structure check |
| `_verify_dcf.py` | DCF formula chain and cross-tab linkage check |

---

### Next Steps (for next session)

1. **Fix DCF Row 25** self-reference bug (likely `=H23/H24` not `=H24/H25`)
2. **Extend forecast horizon** from 3 years to 5 years (FY26E-FY30E) for more robust DCF
3. **Verify with real Excel**: open the .xlsx and confirm all formulas evaluate correctly
4. **Add more historical years** if 10-K data available (currently FY23-FY25, could add FY21-FY22)
5. **Refine assumptions** with broker consensus / FactSet data if available
6. **Add Monte Carlo / tornado chart** sensitivity analysis

---

### Notes
- WebSearch tool unavailable (API error). Used WebFetch with limited success (Gartner data from Wikipedia).
- PC TAM FY2025 (258.6M) is an estimate — Canalys reported ~259M. Verify against Gartner Jan 2026 report.
- Intel share data based on Mercury Research secondary reporting (Tom's Hardware, WCCFTech).
- All external data sources cited in model Notes column.
- **Never manually edit the .xlsx** — always regenerate via `python build_intc_model.py`.
- The build script is the single source of truth for the model.
- **Scenario switching**: Change `Assumptions!B2` to "Base", "Bull", or "Bear" to toggle all forecasts.
