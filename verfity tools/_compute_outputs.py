"""Compute key model outputs by manually walking the formulas for Base case."""
# This script computes what Excel would calculate for the Base case
# We hardcode the assumption values and compute downstream

# === ASSUMPTIONS (Base case) ===
# Segment growth rates
g_ccg = [0.02, 0.03, 0.03]       # FY2026E-2028E
g_dcai = [0.18, 0.12, 0.08]
g_foundry = [0.12, 0.10, 0.10]
g_allother = [-0.05, 0.08, 0.08]

# Segment OPM
opm_ccg = [0.30, 0.315, 0.33]
opm_dcai = [0.28, 0.28, 0.27]
opm_foundry = [-0.42, -0.32, -0.22]
opm_allother = [0.14, 0.18, 0.20]

# External Foundry Revenue ($M)
ext_foundry = [700, 1500, 2500]

# Intersegment eliminations
inter_elim = -200

# Consolidated cost rates
tac_rate = [0.091, 0.088, 0.087]
cogs_other_rate = [0.455, 0.440, 0.430]
rd_rate = [0.195, 0.180, 0.170]
mga_rate = [0.070, 0.065, 0.063]

# WACC
wacc = 0.1066
g_terminal = 0.025

# CapEx ($M)
capex = [17700, 19500, 21000]

# D&A ($M)
da = [12500, 13200, 13800]

# SBC ($M)
sbc = 2500

# Tax rate
etr = [0.12, 0.13, 0.13]

# Interest & Other
interest = [-800, -600, -400]

# NCI ($M)
nci = [750, 1100, 1100]

# Corporate unallocated ($M)
corp_unalloc = [-5000, -4500, -4000]

# Shares (M)
shares = [5100, 5180, 5250]

# Debt ($M)
debt = [49500, 47200, 46000]

# NWC % of revenue change
nwc_pct = 0.02

# === SEGMENT REVENUE ===
fy25 = {"ccg": 32228, "dcai": 16919, "foundry": 17826, "allother": 3563}
fy26_rev = {}
fy27_rev = {}
fy28_rev = {}

for i, yr in enumerate(["fy26", "fy27", "fy28"]):
    fy26_rev["ccg"] = fy25["ccg"] * (1 + g_ccg[0])
    fy26_rev["dcai"] = fy25["dcai"] * (1 + g_dcai[0])
    fy26_rev["foundry"] = fy25["foundry"] * (1 + g_foundry[0])
    fy26_rev["allother"] = fy25["allother"] * (1 + g_allother[0])

fy27_rev["ccg"] = fy26_rev["ccg"] * (1 + g_ccg[1])
fy27_rev["dcai"] = fy26_rev["dcai"] * (1 + g_dcai[1])
fy27_rev["foundry"] = fy26_rev["foundry"] * (1 + g_foundry[1])
fy27_rev["allother"] = fy26_rev["allother"] * (1 + g_allother[1])

fy28_rev["ccg"] = fy27_rev["ccg"] * (1 + g_ccg[2])
fy28_rev["dcai"] = fy27_rev["dcai"] * (1 + g_dcai[2])
fy28_rev["foundry"] = fy27_rev["foundry"] * (1 + g_foundry[2])
fy28_rev["allother"] = fy27_rev["allother"] * (1 + g_allother[2])

# Consolidated Revenue = CCG + DCAI + All Other + External Foundry + Eliminations
consol_rev = {}
consol_rev["fy26"] = fy26_rev["ccg"] + fy26_rev["dcai"] + fy26_rev["allother"] + ext_foundry[0] + inter_elim
consol_rev["fy27"] = fy27_rev["ccg"] + fy27_rev["dcai"] + fy27_rev["allother"] + ext_foundry[1] + inter_elim
consol_rev["fy28"] = fy28_rev["ccg"] + fy28_rev["dcai"] + fy28_rev["allother"] + ext_foundry[2] + inter_elim

print("=" * 70)
print("INTEL (INTC) — MODEL OUTPUT SUMMARY (Base Case)")
print("=" * 70)

print(f"\n{'Metric':<40} {'FY2026E':>10} {'FY2027E':>10} {'FY2028E':>10}")
print("-" * 70)

# Revenue
print(f"{'Consolidated Revenue ($M)':<40} {consol_rev['fy26']:>10,} {consol_rev['fy27']:>10,} {consol_rev['fy28']:>10,}")
print(f"{'  YoY Growth':<40} {'':>10} {consol_rev['fy27']/consol_rev['fy26']-1:>9.1%} {consol_rev['fy28']/consol_rev['fy27']-1:>9.1%}")

# CCG
print(f"{'  CCG Revenue':<40} {fy26_rev['ccg']:>10,} {fy27_rev['ccg']:>10,} {fy28_rev['ccg']:>10,}")
# DCAI
print(f"{'  DCAI Revenue':<40} {fy26_rev['dcai']:>10,} {fy27_rev['dcai']:>10,} {fy28_rev['dcai']:>10,}")
# External Foundry
print(f"{'  External Foundry Revenue':<40} {ext_foundry[0]:>10,} {ext_foundry[1]:>10,} {ext_foundry[2]:>10,}")
# All Other
print(f"{'  All Other Revenue':<40} {fy26_rev['allother']:>10,} {fy27_rev['allother']:>10,} {fy28_rev['allother']:>10,}")

# === CONSOLIDATED P&L ===
# COGS
revs = [consol_rev["fy26"], consol_rev["fy27"], consol_rev["fy28"]]
# TAC = Ad Revenue x TAC rate (Ad rev ~26% of DCAI + 33% of CCG ≈ simplified)
ad_rev = [0.30 * (fy26_rev["ccg"] + fy26_rev["dcai"]),
          0.30 * (fy27_rev["ccg"] + fy27_rev["dcai"]),
          0.30 * (fy28_rev["ccg"] + fy28_rev["dcai"])]
tac = [ad_rev[i] * tac_rate[i] for i in range(3)]
cogs_other = [revs[i] * cogs_other_rate[i] for i in range(3)]
cogs_total = [tac[i] + cogs_other[i] for i in range(3)]
gp = [revs[i] - cogs_total[i] for i in range(3)]
gpm = [gp[i]/revs[i] for i in range(3)]

print(f"{'Gross Profit ($M)':<40} {gp[0]:>10,} {gp[1]:>10,} {gp[2]:>10,}")
print(f"{'  GPM %':<40} {gpm[0]:>9.1%} {gpm[1]:>9.1%} {gpm[2]:>9.1%}")

# OpEx
rd = [revs[i] * rd_rate[i] for i in range(3)]
mga = [revs[i] * mga_rate[i] for i in range(3)]
opex_total = [rd[i] + mga[i] for i in range(3)]

# === SEGMENT OI (source of truth for EBIT) ===
seg_oi_ccg = [fy26_rev["ccg"]*opm_ccg[0], fy27_rev["ccg"]*opm_ccg[1], fy28_rev["ccg"]*opm_ccg[2]]
seg_oi_dcai = [fy26_rev["dcai"]*opm_dcai[0], fy27_rev["dcai"]*opm_dcai[1], fy28_rev["dcai"]*opm_dcai[2]]
seg_oi_foundry = [fy26_rev["foundry"]*opm_foundry[0], fy27_rev["foundry"]*opm_foundry[1], fy28_rev["foundry"]*opm_foundry[2]]
seg_oi_allother = [fy26_rev["allother"]*opm_allother[0], fy27_rev["allother"]*opm_allother[1], fy28_rev["allother"]*opm_allother[2]]
sum_seg_oi = [seg_oi_ccg[i] + seg_oi_dcai[i] + seg_oi_foundry[i] + seg_oi_allother[i] for i in range(3)]
ebit = [sum_seg_oi[i] + corp_unalloc[i] for i in range(3)]
ebit_margin = [ebit[i]/revs[i] for i in range(3)]

print(f"{'EBIT ($M)':<40} {ebit[0]:>10,} {ebit[1]:>10,} {ebit[2]:>10,}")
print(f"{'  EBIT Margin %':<40} {ebit_margin[0]:>9.1%} {ebit_margin[1]:>9.1%} {ebit_margin[2]:>9.1%}")

# Segment OI detail
print(f"{'  CCG OI':<40} {seg_oi_ccg[0]:>10,} {seg_oi_ccg[1]:>10,} {seg_oi_ccg[2]:>10,}")
print(f"{'  DCAI OI':<40} {seg_oi_dcai[0]:>10,} {seg_oi_dcai[1]:>10,} {seg_oi_dcai[2]:>10,}")
print(f"{'  Foundry OI (internal+external)':<40} {seg_oi_foundry[0]:>10,} {seg_oi_foundry[1]:>10,} {seg_oi_foundry[2]:>10,}")
print(f"{'  All Other OI':<40} {seg_oi_allother[0]:>10,} {seg_oi_allother[1]:>10,} {seg_oi_allother[2]:>10,}")

# Below EBIT
pretax = [ebit[i] + interest[i] for i in range(3)]
tax = [-pretax[i] * etr[i] if pretax[i] > 0 else 0 for i in range(3)]  # simplified
ni_before_nci = [pretax[i] - tax[i] for i in range(3)]
ni = [ni_before_nci[i] - nci[i] for i in range(3)]
ni_margin = [ni[i]/revs[i] for i in range(3)]
eps = [ni[i]/shares[i] for i in range(3)]

print(f"{'Net Income ($M)':<40} {ni[0]:>10,} {ni[1]:>10,} {ni[2]:>10,}")
print(f"{'  NI Margin %':<40} {ni_margin[0]:>9.1%} {ni_margin[1]:>9.1%} {ni_margin[2]:>9.1%}")
print(f"{'Diluted EPS ($)':<40} {eps[0]:>10.2f} {eps[1]:>10.2f} {eps[2]:>10.2f}")

# === DCF ===
nopat = [ebit[i] * (1 - etr[i]) for i in range(3)]
nwc_change = [(revs[i] - (52853 if i == 0 else revs[i-1])) * nwc_pct for i in range(3)]  # rough
ufcf = [nopat[i] + da[i] - capex[i] - nwc_change[i] for i in range(3)]

print(f"\n{'UFCF ($M)':<40} {ufcf[0]:>10,} {ufcf[1]:>10,} {ufcf[2]:>10,}")
print(f"{'  NOPAT':<40} {nopat[0]:>10,} {nopat[1]:>10,} {nopat[2]:>10,}")
print(f"{'  D&A':<40} {da[0]:>10,} {da[1]:>10,} {da[2]:>10,}")
print(f"{'  CapEx':<40} {-capex[0]:>10,} {-capex[1]:>10,} {-capex[2]:>10,}")

# TV and DCF
tv = ufcf[2] * (1 + g_terminal) / (wacc - g_terminal)
df = [(1 + wacc)**(-i-1) for i in range(3)]  # simplified - actual model uses mid-year
pv_ufcf = [ufcf[i] * df[i] for i in range(3)]
pv_tv = tv * df[2]
ev = sum(pv_ufcf) + pv_tv

# Equity bridge (simplified - using FY2025 BS cash/debt)
fy25_cash = 14265
fy25_debt = 46585  # approximate
equity_val = ev + fy25_cash - fy25_debt
implied_price = equity_val / shares[2]  # using FY2028E shares

print(f"\n{'DCF VALUATION':-^70}")
print(f"{'Sum PV UFCF ($M)':<40} {sum(pv_ufcf):>10,}")
print(f"{'PV Terminal Value ($M)':<40} {pv_tv:>10,}")
print(f"{'Enterprise Value ($M)':<40} {ev:>10,}")
print(f"{'  + Cash (FY2025A)':<40} {fy25_cash:>10,}")
print(f"{'  - Debt (FY2025A)':<40} {fy25_debt:>10,}")
print(f"{'Equity Value ($M)':<40} {equity_val:>10,}")
print(f"{'Diluted Shares (M)':<40} {shares[2]:>10,}")
print(f"{'IMPLIED SHARE PRICE ($)':<40} {implied_price:>10.2f}")

print(f"\n{'WACC':<40} {wacc:>9.1%}")
print(f"{'Terminal Growth':<40} {g_terminal:>9.1%}")
print(f"{'TV as % of EV':<40} {pv_tv/ev:>9.1%}")
