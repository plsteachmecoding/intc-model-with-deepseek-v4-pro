# Intel (INTC) Financial Model — Assumptions & Scenario Design

**Model horizon:** FY2026E – FY2028E (3 years)
**Historical anchors:** FY2023A / FY2024A / FY2025A
**Fiscal year:** Ends last Saturday of December
**Date of assumptions:** May 7, 2026 (incorporates Q1 2026 actuals and April 23, 2026 earnings call guidance)
**All values in USD Million unless otherwise noted**

---

## Segment Structure Note

Intel restructured segment reporting in Q1 2025. **NEX** (Network & Edge) was integrated into CCG and DCAI and no longer exists as a separate segment. **Altera** was deconsolidated on September 12, 2025 (51% sold to SLP); the retained 49% stake is accounted for under the equity method and does NOT appear in consolidated revenue.

**Modeling approach:** We use the current 4-segment structure throughout the forecast period:

| Segment | Description | Key Drivers |
|---|---|---|
| **CCG** (Client Computing Group) | Notebook, desktop, edge devices; includes former NEX networking products | PC TAM, 18A ramp (Core Ultra Series 3), pricing, supply availability |
| **DCAI** (Data Center & AI) | Xeon server CPUs, AI accelerators, ASICs, NICs, IPUs; includes former NEX telecom infrastructure | AI CPU demand cycle, hyperscale capex, AMD/ARM competition, ASIC pipeline |
| **Intel Foundry** | Internal + external wafer manufacturing, advanced packaging | 18A/14A ramp, yield improvement, external customer adoption, node transitions |
| **All Other** | Mobileye, IMS Nanofabrication, start-up businesses, divested legacy | Mobileye ADAS cycle, IMS multi-beam mask tools |

Intersegment eliminations reconcile segment revenue to consolidated revenue (Intel Foundry sells wafers to Intel Products at cost+).

---

## 1. Segment Revenue Growth Assumptions

### 1A. CCG — Client Computing Group

**Historical Revenue:**
| | FY2023A | FY2024A | FY2025A | Q1 2025A | Q1 2026A |
|---|---|---|---|---|---|
| CCG Revenue ($M) | 32,305 | 33,346 | 32,228 | 7,629 | 7,727 |
| YoY Growth % | — | +3.2% | −3.4% | — | +1.3% |
| Client ASP change | flat | flat | flat | — | +16% |

**Key Context:**
- FY2025 decline driven by: (a) lower client volume from inventory normalization, (b) customer incentive lapping in H1 2024, (c) Intel 7 node supply constraints in Q4 2025 that limited shipments below market demand
- Q1 2026: ASP +16% (mix shift to premium + pricing actions), volume −13% (supply constrained); market demand exceeds available supply
- Management Q1 2026 call: PC TAM down "low double-digit %" in FY2026, but Intel "not as impacted" due to customer inventory replenishment. Zinsner: "model Q2 level forward for CCG" (flattish from Q2 through year-end)
- 18A Core Ultra Series 3 (Panther Lake): "strongest product launch in five years," volume up 6-7x in Q2 vs Q1
- Supply constraints expected to "persist at least through H1 2026" then ease as wafer starts increase across Intel 7, Intel 3, 18A
- Long-term: PC installed base aging (post-COVID), Windows 10 EOL (Oct 2025), AI PC catalysts

**Scenario Assumptions:**

| CCG Revenue Growth % | FY2026E | FY2027E | FY2028E | Rationale |
|---|---|---|---|---|
| **Base** | +2.0% | +3.0% | +3.0% | FY2026: supply recovery offsets PC TAM decline; FY2027-28: modest PC unit growth + ASP lift from AI PC mix shift. Q2 run-rate implied ~$8.0B/quarter × 4 ≈ $32.9B |
| **Bull** | +5.0% | +6.0% | +5.0% | Stronger PC TAM (flat to +low single), share gains from superior 18A products, faster supply recovery, enterprise refresh cycle accelerates |
| **Bear** | −3.0% | 0.0% | +1.0% | PC TAM down 15%+ (recession), persistent supply constraints through FY2026, AMD share gains in client, pricing pressure from competitive environment |

| CCG Revenue ($M) | FY2025A | FY2026E | FY2027E | FY2028E |
|---|---|---|---|---|
| **Base** | 32,228 | 32,873 | 33,859 | 34,875 |
| **Bull** | 32,228 | 33,839 | 35,870 | 37,664 |
| **Bear** | 32,228 | 31,261 | 31,261 | 31,574 |

---

### 1B. DCAI — Data Center & AI

**Historical Revenue:**
| | FY2023A | FY2024A | FY2025A | Q1 2025A | Q1 2026A |
|---|---|---|---|---|---|
| DCAI Revenue ($M) | 15,980 | 16,125 | 16,919 | 4,126 | 5,052 |
| YoY Growth % | — | +0.9% | +4.9% | — | +22.4% |
| Server ASP change | +11% | −4% | mixed (H1 down, H2 up) | — | +27% |

**Key Context:**
- FY2025: Server volume +9% (hyperscale demand), ASPs down 4% from competitive pricing in H1 partially offset by stronger H2 demand. Q4 2025 constrained by Intel 7/Intel 3 supply.
- Q1 2026: Dramatic acceleration — revenue +22% YoY, ASP +27% (premium mix + pricing), volume −5% (supply constrained despite demand). Operating margin hit 30.5%.
- **This is the engine of Intel's turnaround story.** Management: "strong double-digit unit growth" in FY2026, momentum "extending into 2027." DCAI now >60% of Intel Products mix.
- CPU renaissance thesis: CPU-to-GPU ratio moving from 1:8 (training) → 1:4 (inference) → toward 1:1 (agentic AI). Zinsner: "confidence in sustained growth CPUs driven by the AI infrastructure build-out is growing."
- ASIC business run rate "north of $1B" and growing fast; "nearly doubling year over year" in Q1 2026
- Multi-year LTAs signed (Google, others) — typically 3-5 year agreements with volume + pricing terms
- Supply constraints limiting revenue in near term: "undershipping demand by a B (billions)"
- Q2 2026 guide: DCAI "up double digits" sequentially → implies Q2 revenue ~$5.6B+
- Competitive risks: AMD Turin, ARM-based (Graviton, Axion, NVIDIA Vera), but Intel's captive foundry + advanced packaging + x86 ecosystem create structural advantages

**Scenario Assumptions:**

| DCAI Revenue Growth % | FY2026E | FY2027E | FY2028E | Rationale |
|---|---|---|---|---|
| **Base** | +18.0% | +12.0% | +8.0% | FY2026: supply recovery + hyperscale CPU demand + ASIC ramp + pricing; FY2027-28: moderates as competition responds and base effect kicks in, but x86 CPU demand structural from AI inference shift |
| **Bull** | +25.0% | +18.0% | +12.0% | CPU renaissance fully materializes (ratio reaches 1:1 sooner), Coral Rapids accelerates competitive position, ASIC business grows to $3B+, NVIDIA/Google partnerships expand, supply constraints resolve faster |
| **Bear** | +10.0% | +5.0% | +2.0% | AMD Turin takes meaningful share, ARM CPUs gain hyperscale traction (Graviton/Axion), supply constraints persist longer, hyperscale capex shifts toward accelerators away from CPUs, pricing pressure from competition |

| DCAI Revenue ($M) | FY2025A | FY2026E | FY2027E | FY2028E |
|---|---|---|---|---|
| **Base** | 16,919 | 19,964 | 22,360 | 24,149 |
| **Bull** | 16,919 | 21,149 | 24,956 | 27,950 |
| **Bear** | 16,919 | 18,611 | 19,541 | 19,932 |

---

### 1C. Intel Foundry

**Historical Revenue:**
| | FY2023A | FY2024A | FY2025A | Q1 2025A | Q1 2026A |
|---|---|---|---|---|---|
| Foundry Revenue ($M) | 18,504 | 17,317 | 17,826 | 4,667 | 5,421 |
| YoY Growth % | — | −6.4% | +2.9% | — | +16.2% |
| External Revenue ($M) | — | — | — | ~31 | 174 |

**Key Context:**
- Intel Foundry revenue is **primarily intersegment** (wafer sales to Intel Products at cost+). Of $17.8B FY2025 revenue, the vast majority is internal transfer pricing. External foundry revenue was only $174M in Q1 2026.
- Revenue growth reflects: (a) higher wafer volumes from Intel Products demand, (b) mix shift to higher-ASP EUV nodes (Intel 3, 18A), (c) nascent external customer revenue
- Intel Products is ramping more products on Intel 3 and 18A — these carry higher ASPs (per-wafer pricing) than legacy nodes
- Q1 2026: +20% sequentially driven by "increased EUV wafer mix driven by Intel 3, and significant growth in 18A"
- Advanced packaging: backlog "in billions per year" range, differentiated offering, attractive pricing; Malaysia facility expansion announced; revenue to convert starting 2027
- External customer design commitments expected "late 2026 / early 2027" for 14A; 18A external revenue timeline unclear
- Management expects Intel Foundry revenue to grow as Intel Products volumes shift to leading nodes, AND as external customers begin contributing

**Scenario Assumptions:**

| Intel Foundry Revenue Growth % | FY2026E | FY2027E | FY2028E | Rationale |
|---|---|---|---|---|
| **Base** | +12.0% | +10.0% | +10.0% | Internal product ramp on EUV nodes drives growth; external customer revenue begins contributing in FY2027-28 at modest levels; advanced packaging revenue ramps from 2027 |
| **Bull** | +18.0% | +20.0% | +18.0% | 18A ramp exceeds expectations, external customers commit earlier and at larger scale, advanced packaging backlog converts faster, 14A PDK engagement leads to early revenue |
| **Bear** | +5.0% | +3.0% | +3.0% | Internal product demand weaker (CCG/DCAI down), 18A yield issues slow ramp, external customers delay commitments, competitive foundry market limits external wins |

| Intel Foundry Revenue ($M) | FY2025A | FY2026E | FY2027E | FY2028E |
|---|---|---|---|---|
| **Base** | 17,826 | 19,965 | 21,962 | 24,158 |
| **Bull** | 17,826 | 21,035 | 25,242 | 29,785 |
| **Bear** | 17,826 | 18,717 | 19,279 | 19,857 |

> **Note:** This is TOTAL reported segment revenue (primarily intersegment). For consolidated revenue, intersegment eliminations remove the Intel Products↔Foundry internal transfer. External foundry revenue is tracked separately and is currently immaterial (<$200M/quarter). See intersegment eliminations below.

---

### 1D. All Other (Mobileye + IMS + Other)

**Historical Revenue:**
| | FY2023A | FY2024A | FY2025A | Q1 2025A | Q1 2026A |
|---|---|---|---|---|---|
| All Other Revenue ($M) | 5,463 | 3,601 | 3,563 | 943 | 628 |
| YoY Growth % | — | −34.1% | −1.1% | — | −33.4% |

**Key Context:**
- FY2024 decline driven by Altera inventory correction (−$1.3B YoY) and Mobileye inventory digestion (−$425M)
- FY2025: Altera deconsolidated Sep 12, 2025 (only ~8.5 months included). Mobileye revenue $1.9B (+$240M YoY on EyeQ recovery)
- Q1 2026: $628M — Mobileye revenue impacted by macroeconomic uncertainty and competitive pressure; $3.9B goodwill impairment recognized
- Altera is GONE from All Other from Sep 2025 onward (equity method investment)
- Mobileye: publicly traded (MBLY), Intel owns ~77% economic interest. Global leader in ADAS/autonomous driving. Cyclical automotive semiconductor exposure.
- IMS Nano: multi-beam mask writing tools for EUV; niche, high-value, ~32% owned by external investors

**Scenario Assumptions:**

| All Other Revenue Growth % | FY2026E | FY2027E | FY2028E | Rationale |
|---|---|---|---|---|
| **Base** | −5.0% | +8.0% | +8.0% | FY2026: Mobileye headwinds from auto cycle + Altera full-year deconsolidation effect; FY2027-28: Mobileye recovery as SuperVision/Chauffeur ramp, IMS growth |
| **Bull** | +5.0% | +15.0% | +12.0% | Mobileye wins major OEM programs, ADAS regulatory tailwinds, IMS demand surge from EUV fab buildout |
| **Bear** | −15.0% | −5.0% | +2.0% | Auto recession, Mobileye share loss to Qualcomm/NVIDIA, further impairments |

| All Other Revenue ($M) | FY2025A | FY2026E | FY2027E | FY2028E |
|---|---|---|---|---|
| **Base** | 3,563 | 3,385 | 3,656 | 3,948 |
| **Bull** | 3,563 | 3,741 | 4,302 | 4,819 |
| **Bear** | 3,563 | 3,029 | 2,877 | 2,935 |

---

### 1E. Intersegment Eliminations

Intersegment revenue is primarily Intel Foundry selling wafers to Intel Products (CCG, DCAI). Consolidated revenue = CCG + DCAI + Intel Foundry + All Other − Intersegment Eliminations.

**Historical:**
| ($M) | FY2023A | FY2024A | FY2025A |
|---|---|---|---|
| Intersegment Eliminations (reported) | (205) | (161) | 619 |
| As % of Foundry Revenue | 1.1% | 0.9% | −3.5% |

The FY2025 positive $619M (revenue add) is unusual — it means segment revenue was understated relative to consolidated. This appears related to the U.S. Government Agreement accounting adjustments.

For forecasting, we assume intersegment eliminations net to near zero (Foundry revenue ≈ cost of wafers sold to Products). We set eliminations as a small % of Foundry revenue.

| Intersegment Elim ($M) | FY2026E | FY2027E | FY2028E |
|---|---|---|---|
| **All scenarios** | (200) | (200) | (200) |

This means: **Consolidated Revenue ≈ CCG + DCAI + Foundry + All Other − $200M**

---

### 1F. Consolidated Revenue Summary (Base Case)

| ($M) | FY2025A | FY2026E | FY2027E | FY2028E |
|---|---|---|---|---|
| CCG | 32,228 | 32,873 | 33,859 | 34,875 |
| DCAI | 16,919 | 19,964 | 22,360 | 24,149 |
| Intel Foundry | 17,826 | 19,965 | 21,962 | 24,158 |
| All Other | 3,563 | 3,385 | 3,656 | 3,948 |
| Intersegment Eliminations | 619 | (200) | (200) | (200) |
| **Consolidated Revenue** | **52,853** | **55,987** | **61,637** | **66,930** |
| *YoY Growth %* | *(0.5)%* | *+5.9%* | *+10.1%* | *+8.6%* |

---

## 2. Segment Operating Margin Assumptions

### 2A. CCG Operating Margin

**Historical:**
| | FY2023A | FY2024A | FY2025A | Q1 2026A |
|---|---|---|---|---|
| CCG Rev ($M) | 32,305 | 33,346 | 32,228 | 7,727 |
| CCG OpInc ($M) | 10,128 | 11,594 | 9,317 | 2,516 |
| **CCG OPM %** | **31.4%** | **34.8%** | **28.9%** | **32.6%** |

**Key Drivers:**
- FY2025 margin compression: lower volume + customer incentives in H1, partially offset by cost controls
- Q1 2026: margin recovered to 32.6% on stronger ASP (+16%), improved mix, sales of previously reserved inventory, better 18A yields, lower OpEx
- 18A Core Ultra Series 3 (Panther Lake): Q2 volume 6-7x Q1; product margins improving but "still below corporate average" — this is a headwind to segment margins as mix shifts toward 18A products early in ramp
- Longer-term: as 18A yields mature, product costs decline → margin expansion
- CCG OpEx declining from cost-cutting initiatives (headcount reduction ~15%)
- Pricing actions taken to offset higher input costs (memory, substrates)

| CCG OPM % | FY2026E | FY2027E | FY2028E | Rationale |
|---|---|---|---|---|
| **Base** | 30.0% | 31.5% | 33.0% | Near-term: Q1 32.6% level moderated by 18A ramp costs in Q2-Q4; FY2027-28: gradual expansion as 18A yields mature and cost structure improves |
| **Bull** | 33.0% | 36.0% | 38.0% | 18A yields exceed targets faster, PC TAM supports pricing, OpEx efficiency gains accelerate; returns toward FY2024 peak |
| **Bear** | 27.0% | 27.0% | 28.0% | PC TAM contraction forces pricing concessions, 18A cost headwinds persist, competitive pressure from AMD in client |

---

### 2B. DCAI Operating Margin

**Historical:**
| | FY2023A | FY2024A | FY2025A | Q1 2026A |
|---|---|---|---|---|
| DCAI Rev ($M) | 15,980 | 16,125 | 16,919 | 5,052 |
| DCAI OpInc ($M) | 945 | 1,414 | 3,422 | 1,542 |
| **DCAI OPM %** | **5.9%** | **8.8%** | **20.2%** | **30.5%** |

**Key Drivers:**
- **Remarkable margin trajectory:** 5.9% → 30.5% in 3 years, driven by: (a) higher ASPs (+27% in Q1 2026 from premium mix + pricing), (b) improved product margins, (c) lower OpEx from cost-cutting, (d) ASIC business with attractive margins
- Q1 2026 included benefit of ~$300M from "improved mix and product margins, sales of previously reserved inventory, better 18A yields"
- Q2 2026: sequential double-digit revenue growth; margin likely stays strong
- Multi-year LTAs provide pricing visibility; ASIC business should sustain above-corporate-average margins
- Risk: competitive pressure could compress ASPs; hyperscale customers have bargaining power for volume pricing

| DCAI OPM % | FY2026E | FY2027E | FY2028E | Rationale |
|---|---|---|---|---|
| **Base** | 28.0% | 28.0% | 27.0% | Sustains near Q1 2026 levels; FY2028 slight moderation as competition intensifies (AMD Coral Rapids competitor, ARM inroads). ASIC growth supports margins |
| **Bull** | 32.0% | 33.0% | 32.0% | CPU pricing power sustained, ASIC scales to $3B+ with high margins, 18A-based Xeon products command premium, operating leverage on revenue growth |
| **Bear** | 22.0% | 20.0% | 18.0% | AMD Turin + ARM competition drives ASP erosion, hyperscale consolidation of bargaining power, ASIC growth disappoints |

---

### 2C. Intel Foundry Operating Margin

**Historical:**
| | FY2023A | FY2024A | FY2025A | Q1 2026A |
|---|---|---|---|---|
| Foundry Rev ($M) | 18,504 | 17,317 | 17,826 | 5,421 |
| Foundry OpInc ($M) | (7,083) | (13,291) | (10,318) | (2,437) |
| **Foundry OPM %** | **(38.3)%** | **(76.7)%** | **(57.9)%** | **(45.0)%** |

**Key Drivers:**
- FY2024 massive loss driven by: $3.3B non-cash PP&E impairment + accelerated depreciation (Intel 7 node), higher costs from advanced node ramp, lower intersegment pricing
- FY2025 improvement: lower impairment charges ($950M vs $3.3B), partially offset by excess capacity charges ($493M vs $174M)
- Q1 2026: −45.0% OPM — improved $72M QoQ on better yields across Intel 4, 3, 18A; partially offset by increased 14A R&D investments
- **This is the biggest variable in Intel's financial story.** Foundry will lose money through at least 2027.
- Loss drivers: (a) high D&A from massive fab investments (~$10.8B/year and growing), (b) under-absorption as fabs ramp below full capacity, (c) 18A early-ramp elevated costs (low yields, high per-wafer cost), (d) 14A R&D investment
- Positive drivers: (a) 18A yields "ahead of internal projections" — Lip-Bu's year-end target now expected by mid-year, (b) higher EUV wafer mix improves ASP, (c) cost reduction as volumes scale, (d) external customers will pay market pricing (not cost+)
- Management: "expect Intel Foundry's operating loss to improve through the year as 18A continues to ramp into volume and yields improve further"
- **Breakeven analysis:** At current ~$18B revenue and ~$28B costs, Foundry needs either ~55% revenue growth or ~35% cost reduction to break even. This is a multi-year journey.

| Intel Foundry OPM % | FY2026E | FY2027E | FY2028E | Rationale |
|---|---|---|---|---|
| **Base** | (42.0)% | (32.0)% | (22.0)% | Gradual improvement from 18A yield maturation, higher volume absorption, and cost controls. 14A investment offsets some gains. Still deeply unprofitable through 2028 |
| **Bull** | (35.0)% | (20.0)% | (8.0)% | 18A yields beat targets significantly, external customer revenue at market pricing begins contributing, advanced packaging at attractive margins, 14A costs shared with customers |
| **Bear** | (50.0)% | (45.0)% | (40.0)% | 18A yield improvement stalls, capacity under-absorption persists, no material external customer revenue, additional impairment charges, competitive foundry pricing pressure |

---

### 2D. All Other Operating Margin

**Historical:**
| | FY2023A | FY2024A | FY2025A | Q1 2026A |
|---|---|---|---|---|
| All Other Rev ($M) | 5,463 | 3,601 | 3,563 | 628 |
| All Other OpInc ($M) | 1,507 | (57) | 264 | 102 |
| **All Other OPM %** | **27.6%** | **(1.6)%** | **7.4%** | **16.2%** |

**Key Drivers:**
- FY2023: elevated by Altera profitability (pre-downturn)
- FY2024: Altera losses + Mobileye inventory correction
- FY2025: Mobileye recovery offset Altera deconsolidation; Q1 2026: 16.2% margin
- Mobileye: historically 25-35% OPM; currently depressed by R&D investment in SuperVision/Chauffeur and auto market cyclicality
- IMS: niche high-margin semiconductor equipment; contributes positively

| All Other OPM % | FY2026E | FY2027E | FY2028E | Rationale |
|---|---|---|---|---|
| **Base** | 14.0% | 18.0% | 20.0% | Mobileye gradual margin recovery as auto cycle improves and R&D spend normalizes |
| **Bull** | 20.0% | 25.0% | 28.0% | Mobileye design wins convert to high-margin revenue, IMS benefits from EUV fab buildout cycle |
| **Bear** | 8.0% | 10.0% | 12.0% | Mobileye competitive losses, auto downturn, continued R&D spend without revenue conversion |

---

## 3. Consolidated P&L Cost Structure

### 3A. Gross Margin (Cost of Sales as % of Revenue)

**Historical:**
| | FY2023A | FY2024A | FY2025A | Q1 2026A |
|---|---|---|---|---|
| Net Revenue ($M) | 54,228 | 53,101 | 52,853 | 13,577 |
| Cost of Sales ($M) | 32,517 | 35,756 | 34,478 | 8,230 |
| **Gross Margin %** | **40.0%** | **32.7%** | **34.8%** | **39.4%** |
| GM excluding impairments* | ~41.5% | ~39.1% | ~36.6% | ~39.4% |

*\*Estimated: FY2024 COGS included $3.3B impairment/accelerated depreciation in Foundry; FY2025 included $950M. Adjusting: FY2024 underlying COGS ~$32.5B → GM ~39.1%; FY2025 underlying COGS ~$33.5B → GM ~36.6%.*

**Key Drivers:**
- FY2024-2025 GM depressed by non-cash impairment/accelerated depreciation charges on Intel 7 manufacturing assets ($3.3B and $950M respectively)
- Q1 2026: 39.4% GAAP GM — clean quarter (no major COGS impairments). Q2 2026 guidance: **39% non-GAAP**
- 18A ramp headwind: Q2 2026 Panther Lake volume 6-7x Q1; "while the gross margins are improving in Panther Lake quarter to quarter, it's still below the corporate average" → mix shift toward 18A is a GM headwind
- Input cost headwinds: "some of the materials have gone up in terms of costs, substrates are going up... memory going up" — these offset yield-driven improvements in H2 2026
- Management: Zinsner "hyper focused" on gross margins; "goal is to get the gross margins up quickly"
- Positive structural drivers as 18A matures: (a) yields ↑ → lower per-wafer cost, (b) higher 18A volume → better fixed cost absorption, (c) pricing actions flowing through, (d) lower-cost 18A products replacing higher-cost external foundry wafers
- 2027-2028: as 18A reaches mature yields and 14A begins, GM should expand toward historical 40%+ levels

| Gross Margin % | FY2026E | FY2027E | FY2028E | Rationale |
|---|---|---|---|---|
| **Base** | 38.5% | 41.0% | 43.5% | FY2026: modest expansion from Q1 39.4% level, tempered by 18A mix headwind and input cost pressure in H2; FY2027-28: structural improvement as 18A yields mature, volumes scale, and pricing actions compound |
| **Bull** | 41.0% | 45.0% | 49.0% | 18A yields beat expectations, input costs stabilize, pricing power sustained, external foundry revenue at accretive margins, rapid return toward historical 50%+ levels |
| **Bear** | 36.0% | 37.0% | 38.0% | 18A yield improvement stalls, pricing concessions needed, input cost inflation persists, competitive pressure limits GM recovery |

| COGS as % of Revenue | FY2026E | FY2027E | FY2028E |
|---|---|---|---|
| **Base** | 61.5% | 59.0% | 56.5% |
| **Bull** | 59.0% | 55.0% | 51.0% |
| **Bear** | 64.0% | 63.0% | 62.0% |

---

### 3B. R&D as % of Revenue

**Historical:**
| | FY2023A | FY2024A | FY2025A | Q1 2026A |
|---|---|---|---|---|
| R&D ($M) | 16,046 | 16,546 | 13,774 | 3,375 |
| **R&D % of Rev** | **29.6%** | **31.2%** | **26.1%** | **24.9%** |

**Key Drivers:**
- FY2025 R&D declined significantly ($16.5B → $13.8B) from headcount reductions and program rationalization under 2024/2025 Restructuring Plans
- Q1 2026 annualized: $13.5B — running below FY2025 level
- Management FY2026 OpEx target: "~$16B" total (R&D + MG&A), "likely to be higher due to inflationary pressures, variable compensation, and targeted investments"
- Lip-Bu: investing in 14A R&D ("intentional step up"), CPU architecture refinement, ASIC capabilities
- R&D is predominantly in Intel Foundry (process technology development) and Intel Products (product design)
- Expect R&D to grow modestly in absolute terms but decline as % of revenue as top line grows

| R&D % of Revenue | FY2026E | FY2027E | FY2028E | Rationale |
|---|---|---|---|---|
| **Base** | 26.0% | 25.0% | 24.0% | Modest absolute growth from $13.8B → ~$14.5B → ~$15.5B → ~$16.0B; declining as % of growing revenue |
| **Bull** | 25.0% | 23.0% | 21.0% | Revenue growth outpaces R&D; efficiency gains from restructuring; fewer but more focused R&D programs |
| **Bear** | 28.0% | 28.0% | 27.0% | Revenue stagnation with R&D necessary to stay competitive; 14A investment requirements larger than expected |

---

### 3C. MG&A as % of Revenue

**Historical:**
| | FY2023A | FY2024A | FY2025A | Q1 2026A |
|---|---|---|---|---|
| MG&A ($M) | 5,634 | 5,507 | 4,624 | 1,038 |
| **MG&A % of Rev** | **10.4%** | **10.4%** | **8.7%** | **7.6%** |

**Key Drivers:**
- Declining absolute dollars from headcount reduction (~15% of core Intel workforce eliminated by end of FY2025)
- Q1 2026: $1,038M annualized = $4,152M — significantly below FY2025 level
- Restructuring plans substantially complete; further reductions likely to be incremental
- Expect flattish absolute MG&A (~$4.2-4.5B) → declining as % of revenue

| MG&A % of Revenue | FY2026E | FY2027E | FY2028E | Rationale |
|---|---|---|---|---|
| **Base** | 8.0% | 7.5% | 7.0% | ~$4.5B flat absolute; declining as % of growing revenue |
| **Bull** | 7.5% | 6.5% | 6.0% | Further efficiency gains, revenue growth leverage |
| **Bear** | 9.0% | 9.0% | 8.5% | Revenue declines pressure ratio; inflation in SG&A costs |

---

### 3D. Depreciation & Amortization

**Historical:**
| ($M) | FY2023A | FY2024A | FY2025A | Q1 2026A |
|---|---|---|---|---|
| Depreciation (from CF) | 7,847 | 9,951 | 10,757 | 2,902 |
| Amortization of intangibles (from CF) | 1,755 | 1,428 | 949 | 234 |
| **Total D&A** | **9,602** | **11,379** | **11,706** | **3,136** |
| *Annualized D&A* | — | — | — | *~12,544* |

**Key Drivers:**
- D&A growing as fabs transition from construction (CIP) to in-service assets
- CIP declined from $50.4B (Dec 2024) to $34.5B (Dec 2025) → significant assets placed into service → higher depreciation
- Government incentives reduce gross PP&E (and thus future D&A): $6.7B in FY2025 (48D ITC + CHIPS grants)
- Net PP&E: $96.6B (2023) → $107.9B (2024) → $105.4B (2025) — slight decline in FY2025 from large incentive offset
- Gross PP&E before accum deprec: $211.8B at Dec 2025
- Q1 2026 D&A annualized: ~$12.5B
- Useful lives: Buildings 10-25 years, Equipment 2-5 years (accelerated for process nodes with short life)

| D&A ($M) | FY2026E | FY2027E | FY2028E | Rationale |
|---|---|---|---|---|
| **Base** | 12,500 | 13,200 | 13,800 | Continued growth as assets placed in service; partially offset by govt incentives reducing gross PP&E and some older assets fully depreciating |
| **Bull** | 12,000 | 12,500 | 12,800 | Higher govt incentives, more efficient CapEx deployment, some older assets retired faster |
| **Bear** | 13,000 | 14,000 | 15,000 | Higher CapEx, less incentive offset, assets placed in service faster than expected |

| D&A % of Revenue (Base) | FY2026E | FY2027E | FY2028E |
|---|---|---|---|
| | 22.3% | 21.4% | 20.6% |

> D&A is substantially all in Intel Foundry COGS. The high D&A/Revenue ratio is the primary reason Intel Foundry is deeply unprofitable.

---

### 3E. Restructuring & Other Charges

Restructuring charges are winding down. The 2025 Restructuring Plan ($2.2B total) is substantially complete; the 2024 Restructuring Plan (~$3.2B) completes in 2026. We assume minimal ongoing charges.

| Restructuring ($M) | FY2026E | FY2027E | FY2028E |
|---|---|---|
| **All scenarios** | 500 | 100 | 0 |

---

### 3F. Non-Operating Items

**Interest and Other, Net:**
| ($M) | FY2023A | FY2024A | FY2025A | Q1 2026A |
|---|---|---|---|---|
| Interest and other, net | 629 | 226 | 3,257 | (738) |

FY2025 was distorted by $5.6B Altera gain offset by ($1.8B) Escrowed Shares mark-to-market. Q1 2026 included ($1,090M) Escrowed Shares MTM loss.

**Forecast approach:** Model net interest expense based on debt levels. Escrowed Shares MTM is non-cash and non-recurring (shares release over time).

| Interest & Other ($M) | FY2026E | FY2027E | FY2028E | Rationale |
|---|---|---|---|---|
| **Base** | (800) | (600) | (400) | Net interest expense on ~$50B debt (~5% avg coupon = ~$2.5B gross interest, partially offset by ~$1.5B interest income on ~$35B cash/investments, plus capitalized interest). Escrowed Shares MTM assumed to fade |
| **Bull** | (500) | (300) | (100) | Higher interest income, faster Escrowed Share resolution, debt paydown |
| **Bear** | (1,200) | (1,200) | (1,000) | Higher debt, lower interest income, prolonged Escrowed Shares MTM volatility |

**Gains/Losses on Equity Investments:**
| ($M) | FY2026E | FY2027E | FY2028E |
|---|---|---|---|
| **All scenarios** | 0 | 0 | 0 |

> Unpredictable; set to zero for modeling purposes. Includes Altera equity method income, which is expected to be modest.

---

## 4. Working Capital Assumptions

### 4A. Accounts Receivable (AR Days)

| | FY2023A | FY2024A | FY2025A |
|---|---|---|---|
| AR ($M) | 3,402 | 3,478 | 3,839 |
| Revenue ($M) | 54,228 | 53,101 | 52,853 |
| **AR Days** | **22.9** | **23.9** | **26.5** |

AR Days = AR / Revenue × 365. FY2025 increase likely reflects mix shift and timing; Intel historically runs 22-27 days.

| AR Days | FY2026E | FY2027E | FY2028E |
|---|---|---|---|
| **All scenarios** | 25 | 25 | 25 |

---

### 4B. Accounts Payable (AP Days)

| | FY2023A | FY2024A | FY2025A |
|---|---|---|---|
| AP ($M) | 8,578 | 12,556 | 9,882 |
| COGS ($M) | 32,517 | 35,756 | 34,478 |
| **AP Days** | **96.3** | **128.2** | **104.6** |

FY2024 AP elevated from extended vendor payment terms on CapEx (construction payables). FY2025 normalized toward historical levels.

| AP Days | FY2026E | FY2027E | FY2028E |
|---|---|---|---|
| **All scenarios** | 100 | 100 | 100 |

---

### 4C. Inventory Days

| | FY2023A | FY2024A | FY2025A |
|---|---|---|---|
| Inventories ($M) | 11,127 | 12,198 | 11,618 |
| COGS ($M) | 32,517 | 35,756 | 34,478 |
| **Inventory Days** | **124.9** | **124.5** | **123.0** |

Remarkably stable at ~124 days. Intel carries high WIP inventory (semiconductor manufacturing cycle is 3-4 months).

| Inventory Days | FY2026E | FY2027E | FY2028E |
|---|---|---|---|
| **All scenarios** | 125 | 125 | 125 |

---

### 4D. Accrued Compensation as % of Revenue

| | FY2023A | FY2024A | FY2025A |
|---|---|---|---|
| Accrued Comp ($M) | 3,655 | 3,343 | 3,990 |
| Revenue ($M) | 54,228 | 53,101 | 52,853 |
| **% of Revenue** | **6.7%** | **6.3%** | **7.6%** |

| Accrued Comp % | FY2026E | FY2027E | FY2028E |
|---|---|---|---|
| **All scenarios** | 7.0% | 7.0% | 7.0% |

---

### 4E. Other Working Capital

| ($M or %) | Assumption | Notes |
|---|---|---|
| Other Current Assets % of Revenue | ~15% | Includes refundable tax credits (~$7.5B at Dec 2025) which will convert to cash |
| Other Accrued Liabilities % of Revenue | ~25% | Stable relationship |
| Income Taxes Payable | ~2% of Revenue | Normalized level |

---

## 5. Capital Expenditure (CapEx)

### Historical & Guidance Anchors

| ($M) | FY2022A | FY2023A | FY2024A | FY2025A |
|---|---|---|---|---|
| Cash CapEx | 24,844 | 25,750 | 23,944 | 14,646 |
| Financed CapEx | — | — | 1,178 | 3,026 |
| **Gross CapEx (cash + financed)** | **24,844** | **25,750** | **25,122** | **17,672** |
| Gov't Incentives (capital-related) | ~1,000 | 2,200 | 4,100 | 6,700 |
| **Net CapEx (gross − incentives)** | **~23,844** | **~23,550** | **~21,022** | **~10,972** |
| *Gross CapEx as % of Revenue* | *39.4%* | *47.5%* | *47.3%* | *33.4%* |

### FY2026 Management Guidance (Q1 2026 Earnings Call, April 23, 2026):

1. **Total FY2026 CapEx: "Flat to last year"** (~$17.7B gross), revised up from prior "flat to down"
2. **Tool spending up ~25% YoY** — equipment that directly increases wafer output
3. **Space spending "down materially"** — buildings/shell construction being reduced
4. **Roughly equal across quarters** (~$4.4B/quarter)
5. Q1 2026 actual: $5.0B gross ($3.6B cash + $1.3B financed) — slightly above quarterly run rate
6. **Adjusted FCF positive for FY2026** (excluding Fab 34 buyout)
7. FY2027+: CapEx depends on external foundry customer commitments; "customer signals would be more concrete in the back half of this year and into early next year"

### CapEx Outlook:

Intel is in a multi-year investment cycle. The "space-heavy" phase (2022-2024, ~$25B/year) built the shell capacity. The "tool-heavy" phase (2026+) installs equipment into existing shells — each dollar is more productive for near-term revenue.

Key question for FY2027-2028: Does Intel need ANOTHER wave of space CapEx to support external foundry customers, or can existing shells accommodate both internal and early external demand?

| Gross CapEx ($M) | FY2026E | FY2027E | FY2028E | Rationale |
|---|---|---|---|---|
| **Base** | 17,700 | 19,500 | 21,000 | FY2026: flat per guidance; FY2027-28: modest increase as external foundry commitments begin converting to capacity investments; tool-heavy, efficient spend |
| **Bull** | 17,500 | 22,000 | 25,000 | External customer commitments materialize faster and larger; Intel accelerates 14A capacity buildout; advanced packaging demand requires additional investment |
| **Bear** | 18,000 | 17,000 | 16,000 | External customers delay; Intel prioritizes FCF and moderates spending; focuses only on internal product needs |

**Supporting detail — Gross CapEx buildup (Base Case):**

| ($M) | FY2026E | FY2027E | FY2028E |
|---|---|---|---|
| Cash CapEx (CF investing) | 14,200 | 16,000 | 17,500 |
| Financed CapEx (CF financing) | 3,500 | 3,500 | 3,500 |
| **Gross CapEx** | **17,700** | **19,500** | **21,000** |
| Gov't Incentives (capital) | 5,000 | 4,000 | 3,500 |
| **Net CapEx** | **12,700** | **15,500** | **17,500** |

**Data Sources:** FY2025 10-K p.65 (CF statement); Q1 2026 10-Q p.8 (CF statement); Q1 2026 Earnings Call transcript — Zinsner prepared remarks ¶71 ("capital expenditures in 2026 to be flat to last year") and Q&A ¶103-108 (Ben Reitzes question on CapEx unpack).

---

## 6. Financing Assumptions

### 6A. Debt

**Historical & Current:**
| ($M) | Dec 2023 | Dec 2024 | Dec 2025 |
|---|---|---|---|
| Short-term debt | 2,288 | 3,729 | 2,499 |
| Long-term debt | 46,978 | 46,282 | 44,086 |
| **Total Debt** | **49,266** | **50,011** | **46,585** |

**Key Events:**
- April 2026: $6.5B new term loan for Fab 34 Ireland buyout (closed post-Q1)
- Management commitment: retire all $2.5B maturities in 2026 and $3.8B in 2027
- FY2025: repaid $3.75B debt; no new long-term debt issued (only commercial paper net $0)
- Intel's debt is mostly fixed-rate notes with staggered maturities; weighted average coupon ~4.5-5.0%

| Total Debt ($M) | FY2026E | FY2027E | FY2028E | Rationale |
|---|---|---|---|---|
| **Base** | 49,500 | 47,200 | 46,000 | Q1 2026 BS + $6.5B new term loan − $2.5B maturities repaid − $1.1B other = ~$49.5B; FY2027: −$3.8B maturities repaid; FY2028: modest net reduction. No new debt assumed beyond Fab 34 loan |
| **Bull** | 47,000 | 43,000 | 39,000 | Strong FCF enables accelerated debt paydown beyond maturities |
| **Bear** | 52,000 | 54,000 | 55,000 | Additional debt needed if FCF remains negative; potential for more strategic financing |

---

### 6B. Share Count

**Historical:**
| (Millions) | FY2023A | FY2024A | FY2025A | Q1 2026A |
|---|---|---|---|---|
| Basic shares (period end) | 4,228 | 4,330 | 4,994 | 5,023 |
| Diluted shares (WASO) | 4,212 | 4,280 | 4,530 | 5,083 |

**Key Events:**
- Aug 2025: +275M shares to U.S. Government (CHIPS warrant agreement)
- Aug 2025: +87M shares to SoftBank (private placement @ $23.00)
- Sep 2025: +215M shares to NVIDIA (private placement @ $23.28)
- 159M Escrowed Shares in escrow (released as Secure Enclave milestones met); 3M released by Dec 2025
- 241M warrants held by U.S. Gov (exercisable @ $20.00 if Intel ceases to own ≥51% of Foundry) — NOT in share count yet; treated as contingent dilution
- ESPP/RSU issuance: ~28M shares/quarter (Q1 2026); offset by RSU withholdings (~5M/quarter)
- Net quarterly increase from equity plans: ~23M shares

| Basic Shares (M, period-end) | FY2026E | FY2027E | FY2028E | Rationale |
|---|---|---|---|---|
| **Base** | 5,100 | 5,180 | 5,250 | Q1 2026: 5,023M + ~75M net from ESPP/RSU/Escrowed Share releases through year-end. No additional strategic issuances assumed. Escrowed Shares released gradually |
| **Bull** | 5,080 | 5,120 | 5,150 | Fewer Escrowed Shares released (milestones delayed), lower dilution |
| **Bear** | 5,200 | 5,400 | 5,600 | Additional strategic equity raises if FCF remains negative; accelerated Escrowed Share releases; warrant exercise if Foundry ownership structure changes |

> **Warrant overhang (241M shares):** Exercisable at $20.00 only if Intel ceases to own ≥51% of Intel Foundry. NOT included in base/bull share count. Partial inclusion in bear case if Foundry separation scenario materializes.

---

### 6C. Dividends

Dividends were suspended in FY2025. Management has given no indication of reinstatement. Turnaround priority is debt reduction and CapEx.

| Dividends per Share | FY2026E | FY2027E | FY2028E |
|---|---|---|---|
| **All scenarios** | $0.00 | $0.00 | $0.00 |

---

### 6D. Share-Based Compensation (SBC)

| ($M) | FY2023A | FY2024A | FY2025A | Q1 2026A |
|---|---|---|---|---|
| SBC | 3,229 | 3,410 | 2,434 | 621 |

SBC declining with headcount reduction (~15% core workforce reduction). Q1 2026 annualized: ~$2.5B.

| SBC ($M) | FY2026E | FY2027E | FY2028E |
|---|---|---|
| **All scenarios** | 2,500 | 2,500 | 2,600 |

---

### 6E. Non-Controlling Interests (NCI)

| ($M) | FY2025A | Q1 2026A | FY2026E | FY2027E | FY2028E |
|---|---|---|---|---|---|
| NCI (P&L) | 293 | (553) | ~750 | ~1,100 | ~1,100 |

Key changes:
- **Ireland SCIP (Fab 34):** 49% NCI eliminated in April 2026 (bought out for ~$14.2B). P&L NCI from Ireland SCIP was ~$135M in Q1 2026 → goes to $0 from Q2 2026.
- **Arizona SCIP:** 49% NCI remains; contributed $183M NCI in Q1 2026. Growing as more assets placed in service.
- **Mobileye:** ~23% NCI (Intel owns ~77%). Q1 2026 loss included ($867M) NCI from $3.9B impairment.
- Management guidance: NCI ~$250M/quarter in Q2-Q4 2026; ~$1.1B/year for 2027-2028

| NCI ($M) | FY2026E | FY2027E | FY2028E |
|---|---|---|
| **All scenarios** | 750 | 1,100 | 1,100 |

---

## 7. Tax Assumptions

**Historical Effective Tax Rate:**
| | FY2023A | FY2024A | FY2025A | Q1 2026A |
|---|---|---|---|---|
| Pretax Income ($M) | 762 | (11,210) | 1,557 | (3,946) |
| Tax Provision ($M) | (913) | 8,023 | 1,531 | 335 |
| **Effective Tax Rate** | **(119.8)%** | **(71.6)%** | **98.3%** | **(8.5)%** |

Historical ETRs are meaningless for forecasting — they are distorted by:
- **Domestic valuation allowance:** Intel cannot benefit from U.S. losses due to a full valuation allowance on its U.S. deferred tax assets (recorded in FY2024, ~$9.9B). This means U.S. pretax losses generate NO tax benefit, while foreign profits are taxed.
- Non-deductible goodwill impairment ($3.9B in Q1 2026)
- R&D tax credits and foreign rate differential

**Q2 2026 management guidance: 11% non-GAAP tax rate.** This is the best anchor.

| Effective Tax Rate | FY2026E | FY2027E | FY2028E | Rationale |
|---|---|---|---|---|
| **Base** | 12.0% | 13.0% | 13.0% | Anchored to Q2 2026 guided 11% with slight premium for uncertainty. Reflects foreign profits taxed at local rates + U.S. losses providing no benefit (valuation allowance). As Intel returns to sustained profitability, ETR normalizes but stays low due to tax credit monetization (48D ITC) |
| **Bull** | 11.0% | 12.0% | 12.0% | Management guided rate sustained; higher 48D ITC benefits |
| **Bear** | 15.0% | 15.0% | 15.0% | Higher foreign effective rate; less credit benefit |

---

## 8. DCF Valuation Parameters

### 8A. Weighted Average Cost of Capital (WACC)

**Input Parameters:**

| Parameter | Value | Rationale |
|---|---|---|
| **Risk-Free Rate (Rf)** | 4.50% | 10Y U.S. Treasury yield as of early May 2026. Elevated relative to historical norms due to fiscal/debt concerns |
| **Equity Risk Premium (ERP)** | 5.50% | Damodaran 2026 implied ERP for U.S. market. Semiconductor sector ERP is at or slightly above market |
| **Levered Beta (β)** | 1.50 | Bloomberg 2Y adjusted beta for INTC. Elevated from historical ~1.0 due to: turnaround volatility, high operating leverage (fixed cost-heavy fab model), uncertainty around foundry strategy. Post-restructuring, beta should decline toward 1.2-1.3 |
| **Cost of Equity (Ke)** | **12.75%** | = Rf + β × ERP = 4.50% + 1.50 × 5.50% |
| **Pre-Tax Cost of Debt (Kd)** | 5.00% | Intel's weighted average coupon on outstanding notes; recent $6.5B term loan likely at SOFR + 150-200bps (~5.5-6.0%), but majority of debt is older fixed-rate at lower coupons |
| **Tax Rate** | 12.00% | Marginal tax rate for interest tax shield (aligned with Base ETR) |
| **After-Tax Cost of Debt** | **4.40%** | = Kd × (1 − Tax Rate) = 5.00% × (1 − 0.12) |
| **Target D/(D+E)** | 25.0% | Based on: Enterprise Value estimation. At ~$20/share × 5.1B shares = ~$102B market cap; debt ~$49B → D/EV ≈ 32%. We use 25% as long-term target weight |
| **Target E/(D+E)** | 75.0% | |
| **WACC** | **10.66%** | = Ke × E/(D+E) + Kd(1−t) × D/(D+E) = 12.75% × 0.75 + 4.40% × 0.25 |

**WACC Sensitivity (to support sensitivity analysis):**
| WACC | β=1.2 | β=1.5 | β=1.8 |
|---|---|---|---|
| ERP=5.0% | 9.43% | 10.88% | 12.33% |
| ERP=5.5% | 9.89% | **10.66%** | 12.79% |
| ERP=6.0% | 10.35% | 11.84% | 13.25% |

**Recommended WACC range for DCF: 10.0% – 11.5%**
Central estimate: **10.7%** (round to 10.5% for base case)

---

### 8B. Terminal Value

| Parameter | Value | Rationale |
|---|---|---|
| **Terminal Growth Rate (g)** | 2.5% | Long-term nominal GDP growth + semiconductor industry premium. Semiconductors are secular growth (AI, electrification, digitization) but Intel's mature segments (PC, server) grow at or slightly above GDP |
| **Terminal Value Method** | Gordon Growth Model | TV = UFCF₂₀₂₈ × (1 + g) / (WACC − g) |

**Terminal growth sensitivity:** Bull case 3.0% (Intel captures AI-driven semiconductor growth); Bear case 2.0% (mature, no-growth terminal state).

---

### 8C. Forecast Horizon

| Parameter | Value | Rationale |
|---|---|---|
| **Explicit forecast** | 3 years (FY2026E–FY2028E) | Per user specification. Covers the critical 18A ramp and turnaround period |
| **Terminal year** | FY2028E normalized | FY2028 is assumed to represent a "steady state" where: 18A is mature, Foundry losses have narrowed significantly (but likely not breakeven), DCAI growth has moderated to trend, and restructuring benefits are fully realized |

> **Limitation note:** 3 years is short for a semiconductor company with a 5-7 year foundry investment cycle. By FY2028, Intel Foundry will still be unprofitable (Base case: −22% OPM), meaning the terminal value will embed significant "perpetual improvement" assumptions. Consider extending the explicit forecast to 5 years (through FY2030) for a more robust DCF. Alternatively, use a **normalized terminal margin** approach rather than assuming FY2028 is steady-state.

---

## 9. Scenario Summary

### Base Case Narrative: "Measured Recovery"
Intel executes its turnaround: 18A yields improve on track, supply constraints gradually ease through FY2026, DCAI benefits from AI CPU demand structurally growing 18% in FY2026 and moderating thereafter. CCG is roughly flat as PC TAM declines offset by inventory replenishment and pricing. Foundry losses narrow from −58% to −22% by FY2028 but remain deeply negative. Gross margins recover to 43.5% by FY2028. CapEx stays disciplined at $17.7-21B. No dividend reinstatement. External foundry revenue remains immaterial through FY2028.

### Bull Case Narrative: "CPU Renaissance"
The AI inference shift toward CPUs materializes faster and larger than expected. DCAI grows 25%+ in FY2026 and sustains double-digit growth through FY2028. 18A yields exceed internal targets significantly, enabling Panther Lake and Clearwater Forest to capture premium pricing and share. Intel Foundry attracts multiple external customers on 14A by late FY2027, with advanced packaging backlog converting to meaningful revenue. Gross margins recover toward 49%. FCF turns strongly positive, enabling debt paydown. Stock re-rates as turnaround credibility is established.

### Bear Case Narrative: "Structural Decline"
PC TAM contracts 15%+ in FY2026 (recession), AMD Turin takes meaningful server share, and ARM-based CPUs (AWS Graviton, Google Axion, NVIDIA Vera) erode x86 incumbency. 18A yield improvement stalls, keeping Foundry losses above 40%. External customers stay with TSMC/Samsung. Gross margins remain at 36-38% — below cost of capital. Intel forced into additional equity raises or asset sales. Warrants become relevant if Foundry separation is contemplated.

### Key Scenario Triggers to Monitor:
| Trigger | Bull Signal | Bear Signal |
|---|---|---|
| 18A yield rate | Mid-year target hit in Q2, Q4 target raised | Mid-year target slips to Q4, Q4 target pushed to FY2027 |
| DCAI revenue | Sustained >20% YoY growth through FY2026 | Growth decelerates below 10% in H2 2026 |
| External foundry | Design commitment announced by H2 2026 | No commitments through FY2027 |
| PC TAM | Flat to −5% | −15% or worse |
| Gross margin trajectory | >40% by Q4 2026 | <38% through FY2027 |
| Competitive dynamics | Diamond Rapids/Coral Rapids on schedule | Further roadmap delays; key customer losses |

---

## 10. Data Sources Reference

| Data Point | Source | Location |
|---|---|---|
| Segment revenue & OI (2023-2025) | 2025 10-K | Note 3 (p.76-77) |
| Segment revenue & OI (2022, old structure) | 2024 10-K | Note 3 (p.72-73) |
| Consolidated P&L (2022-2025) | 2024 10-K p.59; 2025 10-K p.27 | Face financial statements |
| Q1 2026 segment & P&L | Q1 2026 10-Q | Note 2 (p.11-12), MD&A (p.30-33) |
| Balance Sheet (2023-2025) | 2025 10-K p.64; 2024 10-K p.61 | Face financial statements |
| Cash Flow (2022-2025) | 2025 10-K p.65; 2024 10-K p.62 | Face financial statements |
| CapEx history + guidance | 2025 10-K p.65; Q1 2026 10-Q p.8; Q1 2026 Earnings Call transcript | CF statements + CFO remarks/Q&A |
| FY2026 guidance (all metrics) | Q1 2026 Earnings Call (April 23, 2026) | Zinsner prepared remarks + Q&A |
| Segment strategy / competitive positioning | Q1 2026 Earnings Call | Lip-Bu Tan prepared remarks + Q&A |
| Restructuring plans | 2025 10-K p.21; Q1 2026 10-Q Note 6 | MD&A + Notes |
| Government incentives | 2025 10-K p.20, Note 3 | MD&A + Notes |
| WACC inputs (Rf, ERP, Beta) | Bloomberg / Damodaran / FRED | Market data as of May 2026 |

---

*End of Assumptions Document. This file feeds directly into the Excel Assumptions Tab (Section 2 of the IB Financial Model architecture).*
