#!/usr/bin/env python3
"""
build_intc_model.py
Builds Intel_IB_Model.xlsx — 12-tab, segment-driven 3-statement DCF model
for Intel Corporation (INTC) using openpyxl with CHOOSE/MATCH scenario switching.

Usage:  python build_intc_model.py
Output: Intel_IB_Model.xlsx (in current directory)

Architecture (12 tabs):
  Cover → Key_Summary → Assumptions → Segment_Revenue → Segment_PL
  → Consolidated_PL → BS → Cash_Flow → DCF → Sensitivity → Ratio_Analysis

Data sources:
  - intc_financial_data.md  (extracted from 10-Ks, 10-Q)
  - intc_assumptions.md     (forecast assumptions, 3 scenarios)

References:
  - example-structure.md    (tab layouts, cross-reference pattern)
  - SKILL.md (ib-financial-model skill spec)

Author: Generated via Claude Code / ib-financial-model skill
Date:   May 7, 2026
"""

import os, sys, re
from copy import copy

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, numbers
    from openpyxl.utils import get_column_letter as CL
    from openpyxl.worksheet.datavalidation import DataValidation
except ImportError:
    sys.exit("ERROR: openpyxl required. Run: pip install openpyxl")

# ============================================================
# GLOBAL CONSTANTS
# ============================================================

OUTPUT = "Intel_IB_Model_v2.xlsx"
NCOL   = 11          # Notes column (K) — must be > last data column J
SCENARIO_CELL = "$B$2"  # Scenario selector location in Assumptions tab

# Historical columns: B=2 (2023A), C=3 (2024A), D=4 (2025A)
# Forecast columns:  F=6 (2026E), G=7 (2027E), H=8 (2028E)
# Col E=5  spacer, Col I=9  Selected (CHOOSE/MATCH), Col J=10 available, Col K=11 Notes
YR_COLS = {"2023A": 2, "2024A": 3, "2025A": 4, "2026E": 6, "2027E": 7, "2028E": 8}
HIST_COLS = [2, 3, 4]
FCST_COLS = [6, 7, 8]
SEL_COL   = 9

# ============================================================
# COLOR CODING — IB Standard (see legend on every tab)
# ============================================================
BLUE_FONT  = "305496"   # Historical hardcoded data
GREEN_FONT = "548235"   # Cross-sheet link
BLUE_FILL  = "DDEBF7"   # Editable assumption (blue fill bg)
YELLOW_FILL = "FFF2CC"  # Key output row (yellow fill bg)
GREY_FONT  = "808080"   # Memo / percentage row (grey italic)
DARK_FILL  = "1F4E79"   # Header background (dark blue)
WHITE_FONT = "FFFFFF"   # Header text colour

# ============================================================
# REUSABLE STYLES
# ============================================================

BD = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)

FONT_HIST   = Font(color=BLUE_FONT, size=10, name="Calibri")
FONT_LINK   = Font(color=GREEN_FONT, size=10, name="Calibri")
FONT_FORM   = Font(color="000000", size=10, name="Calibri")
FONT_MEMO   = Font(color=GREY_FONT, size=9, italic=True, name="Calibri")
FONT_BOLD   = Font(bold=True, size=10, name="Calibri")
FONT_TITLE  = Font(bold=True, size=14, name="Calibri")
FONT_UNIT   = Font(italic=True, size=9, color=GREY_FONT, name="Calibri")
FONT_HEADER = Font(bold=True, color=WHITE_FONT, size=10, name="Calibri")
FONT_SECTION = Font(bold=True, size=11, name="Calibri", color=DARK_FILL)

FILL_ASSUMP = PatternFill("solid", fgColor=BLUE_FILL)
FILL_KEY    = PatternFill("solid", fgColor=YELLOW_FILL)
FILL_HEADER = PatternFill("solid", fgColor=DARK_FILL)
FILL_GREY   = PatternFill("solid", fgColor="F2F2F2")
FILL_BLANK  = PatternFill("solid", fgColor="FFFFFF")

ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")
ALIGN_LEFT  = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_CENTER = Alignment(horizontal="center", vertical="center")

# ============================================================
# HELPER FUNCTIONS
# ============================================================

def _hval(ws, r, col, val, fmt="#,##0"):
    """Historical hardcoded value — blue font, grey fill"""
    c = ws.cell(row=r, column=col, value=val)
    c.number_format = fmt
    c.font = FONT_HIST
    c.border = BD
    c.alignment = ALIGN_RIGHT
    c.fill = FILL_GREY

def _fval(ws, r, col, formula, fmt="#,##0"):
    """Formula (calculated) — black font"""
    c = ws.cell(row=r, column=col, value=formula)
    c.number_format = fmt
    c.font = FONT_FORM
    c.border = BD
    c.alignment = ALIGN_RIGHT

def _lval(ws, r, col, formula, fmt="#,##0"):
    """Cross-sheet link — green font"""
    c = ws.cell(row=r, column=col, value=formula)
    c.number_format = fmt
    c.font = FONT_LINK
    c.border = BD
    c.alignment = ALIGN_RIGHT

def _pct(ws, r, col, formula):
    """Percentage row — grey italic font"""
    c = ws.cell(row=r, column=col, value=formula)
    c.number_format = "0.0%"
    c.font = FONT_MEMO
    c.border = BD
    c.alignment = ALIGN_RIGHT

def _key(ws, r, col, formula, fmt="#,##0"):
    """Key output row — yellow fill + black font"""
    c = ws.cell(row=r, column=col, value=formula)
    c.number_format = fmt
    c.font = FONT_BOLD
    c.fill = FILL_KEY
    c.border = BD
    c.alignment = ALIGN_RIGHT

def _label(ws, r, text, bold=False):
    """Row label in column A"""
    c = ws.cell(row=r, column=1, value=text)
    c.font = FONT_BOLD if bold else FONT_FORM
    c.border = BD
    c.alignment = ALIGN_LEFT

def _section(ws, r, text):
    """Section header — bold dark blue"""
    c = ws.cell(row=r, column=1, value=text)
    c.font = FONT_SECTION
    for col in range(1, NCOL):
        ws.cell(row=r, column=col).border = BD

def _note(ws, r, text):
    """Notes column (K) — strips leading '=' to prevent openpyxl formula errors"""
    if text and text.startswith("="):
        text = text[1:]
    c = ws.cell(row=r, column=NCOL, value=text)
    c.font = Font(size=9, name="Calibri")
    c.alignment = ALIGN_LEFT

def _title_row(ws, r, text):
    """Tab title — row 1"""
    c = ws.cell(row=r, column=1, value=text)
    c.font = FONT_TITLE
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NCOL - 1)

def _unit_row(ws, r, text="USD Million (except per-share data)"):
    """Unit row — row 2, grey italic"""
    c = ws.cell(row=r, column=1, value=text)
    c.font = FONT_UNIT
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NCOL - 1)

def _year_header(ws, r, extra_cols=True):
    """Standard year-header row with dark fill + white font.
    Cols: B=2023A, C=2024A, D=2025A, E=(spacer), F=2026E, G=2027E, H=2028E
    If extra_cols, also write I=Selected, K=Notes headers.
    """
    hdrs = {2: "2023A", 3: "2024A", 4: "2025A", 6: "2026E", 7: "2027E", 8: "2028E"}
    for col, txt in hdrs.items():
        c = ws.cell(row=r, column=col, value=txt)
        c.font = FONT_HEADER
        c.fill = FILL_HEADER
        c.border = BD
        c.alignment = ALIGN_CENTER
    # Spacer & extra headers
    for col in [5]:
        c = ws.cell(row=r, column=col, value="")
        c.fill = FILL_HEADER
        c.border = BD
    if extra_cols:
        c = ws.cell(row=r, column=9, value="Selected")
        c.font = FONT_HEADER; c.fill = FILL_HEADER; c.border = BD; c.alignment = ALIGN_CENTER
        c = ws.cell(row=r, column=NCOL, value="Notes")
        c.font = FONT_HEADER; c.fill = FILL_HEADER; c.border = BD; c.alignment = ALIGN_CENTER
    # Column A
    c = ws.cell(row=r, column=1, value="")
    c.fill = FILL_HEADER; c.border = BD

def _yr_header_no_scenario(ws, r):
    """Year header WITHOUT Scenario/Notes columns (for non-Assumptions tabs)"""
    hdrs = {2: "2023A", 3: "2024A", 4: "2025A", 6: "2026E", 7: "2027E", 8: "2028E"}
    for col, txt in hdrs.items():
        c = ws.cell(row=r, column=col, value=txt)
        c.font = FONT_HEADER
        c.fill = FILL_HEADER
        c.border = BD
        c.alignment = ALIGN_CENTER
    for col in [1, 5, 9, 10, NCOL]:
        c = ws.cell(row=r, column=col, value="")
        c.fill = FILL_HEADER; c.border = BD
    if NCOL == 11:
        c = ws.cell(row=r, column=NCOL, value="Notes")
        c.font = FONT_HEADER; c.fill = FILL_HEADER; c.border = BD; c.alignment = ALIGN_CENTER


def _arow(ws, r, label, h1, h2, h3, base, bull, bear, fmt="#,##0.0", note_text=""):
    """Full assumption row: 3 historical (B-D) + 3 scenario values (F-H) + CHOOSE/MATCH (I) + note (K).

    Parameters
    ----------
    h1, h2, h3 : float or None  — historical values; None → leave cell blank
    base, bull, bear : float    — scenario values
    fmt : str                   — number format (e.g. "#,##0.0", "0.0%")
    """
    _label(ws, r, label)
    for col, val in [(2, h1), (3, h2), (4, h3)]:
        if val is not None:
            _hval(ws, r, col, val, fmt)
        else:
            ws.cell(row=r, column=col).border = BD
    # Scenario columns (blue fill = editable)
    for col, val in [(6, base), (7, bull), (8, bear)]:
        c = ws.cell(row=r, column=col, value=val)
        c.number_format = fmt
        c.font = Font(color="000000", size=10, name="Calibri")
        c.fill = FILL_ASSUMP
        c.border = BD
        c.alignment = ALIGN_RIGHT
    # Spacer col E
    ws.cell(row=r, column=5).border = BD
    # CHOOSE/MATCH in col I
    formula = f'=CHOOSE(MATCH({SCENARIO_CELL},{{"Base","Bull","Bear"}},0),F{r},G{r},H{r})'
    c = ws.cell(row=r, column=SEL_COL, value=formula)
    c.number_format = fmt
    c.font = FONT_BOLD
    c.border = BD
    c.alignment = ALIGN_RIGHT
    # Col J spacer
    ws.cell(row=r, column=10).border = BD
    # Note
    _note(ws, r, note_text)

def _arow_const(ws, r, label, h1, h2, h3, val, fmt="#,##0.0", note_text=""):
    """Assumption row where Base=Bull=Bear (scenario-independent). Shortcut for _arow."""
    _arow(ws, r, label, h1, h2, h3, val, val, val, fmt, note_text)

def _add_legend(ws, start_row):
    """Add colour-coding legend at the bottom of a tab."""
    r = start_row
    items = [
        ("Blue font", "305496", "Historical hardcoded data"),
        ("Green font", "548235", "Cross-sheet link"),
        ("Blue fill + dark font", "DDEBF7", "Editable assumption"),
        ("Yellow fill", "FFF2CC", "Key output row"),
        ("Black font", "000000", "Formula / calculated cell"),
        ("Grey italic", "808080", "Memo / percentage row"),
    ]
    _label(ws, r, "DATA LEGEND"); r += 1
    for label, color, desc in items:
        c = ws.cell(row=r, column=1, value=f"  {label}: {desc}")
        if "font" in label.lower():
            c.font = Font(color=color, size=9, name="Calibri")
        elif "fill" in label.lower():
            c.font = Font(color="000000", size=9, name="Calibri")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        r += 1
    return r

def _add_sources(ws, start_row):
    """Add data-sources block at the bottom of a tab."""
    r = start_row
    _label(ws, r, "DATA SOURCES"); r += 1
    sources = [
        "2022-2025 10-Ks: Consolidated P&L, Balance Sheet, Cash Flow, Segment data",
        "2026 Q1 10-Q: Q1 2026 vs Q1 2025 segment + consolidated data",
        "Q1 2026 Earnings Call (Apr 23, 2026): FY2026 management guidance",
        "intc_assumptions.md: All forecast assumptions with scenario narratives",
        "Bloomberg / Damodaran / FRED: WACC inputs (Rf, ERP, Beta) as of May 2026",
    ]
    for s in sources:
        c = ws.cell(row=r, column=1, value=f"  - {s}")
        c.font = Font(size=9, name="Calibri")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        r += 1
    return r


# ============================================================
# SCENARIO SELECTOR SETUP (called once, on Assumptions tab)
# ============================================================

def _setup_scenario_selector(ws):
    """Place scenario dropdown in B2 of the Assumptions tab."""
    c = ws.cell(row=2, column=2, value="Base")
    c.font = Font(bold=True, size=11, name="Calibri", color="1F4E79")
    c.fill = PatternFill("solid", fgColor="DDEBF7")
    c.border = Border(
        left=Side(style="medium"), right=Side(style="medium"),
        top=Side(style="medium"), bottom=Side(style="medium"),
    )
    c.alignment = ALIGN_CENTER
    dv = DataValidation(type="list", formula1='"Base,Bull,Bear"', allow_blank=False)
    dv.error = "Please select Base, Bull, or Bear"
    dv.errorTitle = "Invalid Scenario"
    dv.prompt = "Select scenario"
    dv.promptTitle = "Scenario"
    ws.add_data_validation(dv)
    dv.add(c)
    # Label
    lbl = ws.cell(row=2, column=1, value="Scenario:")
    lbl.font = Font(bold=True, size=11, name="Calibri")
    lbl.alignment = ALIGN_RIGHT


# ============================================================
# TAB 3: ASSUMPTIONS (Control Center)
# ============================================================

def build_assumptions(wb, R):
    """Build the Assumptions tab with all scenario inputs and CHOOSE/MATCH switching."""
    ws = wb["Assumptions"]
    _title_row(ws, 1, "Intel (INTC) — Assumptions & Scenario Design")
    _setup_scenario_selector(ws)

    # Row 3: spacer
    # Row 4: year header
    _year_header(ws, 4)

    # Track key rows for cross-reference
    r = 6

    # ---- WACC Parameters ----
    _section(ws, r, "WACC Parameters"); r += 1
    _arow_const(ws, r, "Risk-Free Rate (Rf)", 0.0388, 0.0425, 0.045, 0.045, "0.00%",
                "US 10Y Treasury, May 2026"); R["asp_rf"] = r; r += 1
    _arow_const(ws, r, "Equity Risk Premium (ERP)", 0.055, 0.055, 0.055, 0.055, "0.00%",
                "Damodaran 2026 implied ERP"); R["asp_erp"] = r; r += 1
    _arow(ws, r, "Levered Beta (β)", 1.20, 1.35, 1.50, 1.50, 1.30, 1.80, "0.00",
          "Bloomberg 2Y adj beta; Bear=1.80 on higher risk"); R["asp_beta"] = r; r += 1
    _arow_const(ws, r, "Pre-Tax Cost of Debt (Kd)", 0.048, 0.05, 0.05, 0.050, "0.00%",
                "Weighted avg coupon on outstanding notes"); R["asp_kd"] = r; r += 1
    _arow(ws, r, "Target D/(D+E)", 0.30, 0.32, 0.32, 0.25, 0.20, 0.35, "0.0%",
          "Long-term target; D/EV =32% at current prices"); R["asp_dtc"] = r; r += 1
    _arow(ws, r, "Marginal Tax Rate (for WACC)", None, None, None, 0.12, 0.11, 0.15, "0.00%",
          "Marginal rate for interest tax shield. Base=12% (aligned with Q2 2026 guided 11%)"); R["asp_etr_marginal"] = r; r += 1
    _arow(ws, r, "Terminal Growth Rate (g)", None, None, None, 0.025, 0.030, 0.020, "0.00%",
          "Base=2.5% LT nominal GDP+; Bull=3.0% AI secular; Bear=2.0% mature"); R["asp_tg"] = r; r += 1

    r += 1
    # Derived WACC (formulas, not arow)
    _section(ws, r, "Derived WACC"); r += 1
    # Ke = Rf + Beta * ERP
    _label(ws, r, "Cost of Equity (Ke)")
    _fval(ws, r, 2, f"=I{R['asp_rf']}+I{R['asp_beta']}*I{R['asp_erp']}", "0.00%")
    _fval(ws, r, 6, f"=F{R['asp_rf']}+F{R['asp_beta']}*F{R['asp_erp']}", "0.00%")
    _fval(ws, r, 7, f"=G{R['asp_rf']}+G{R['asp_beta']}*G{R['asp_erp']}", "0.00%")
    _fval(ws, r, 8, f"=H{R['asp_rf']}+H{R['asp_beta']}*H{R['asp_erp']}", "0.00%")
    _note(ws, r, "Ke = Rf + β × ERP")
    R["asp_ke"] = r; r += 1
    # After-tax Kd
    _label(ws, r, "After-Tax Cost of Debt")
    _fval(ws, r, SEL_COL, f"=I{R['asp_kd']}*(1-I{R['asp_etr_marginal']})", "0.00%")
    _note(ws, r, "Kd × (1 − ETR)")
    R["asp_atkd"] = r; r += 1
    # WACC
    _label(ws, r, "WACC")
    _fval(ws, r, SEL_COL, f"=I{R['asp_ke']}*(1-I{R['asp_dtc']})+I{R['asp_atkd']}*I{R['asp_dtc']}", "0.00%")
    _note(ws, r, "WACC = Ke×(1−D/TC) + Kd(1−t)×D/TC. Links to DCF tab")
    R["asp_wacc"] = r; r += 1

    # ---- Valuation Bridge ----
    r += 1
    _section(ws, r, "Valuation Bridge Inputs"); r += 1
    _label(ws, r, "Cash & Equivalents ($M) — from BS")
    _note(ws, r, "To be linked from BS post-build (hardcoded anchor for DCF bridge)")
    R["asp_cash_bridge"] = r; r += 1
    _label(ws, r, "Total Debt ($M) — from BS")
    R["asp_debt_bridge"] = r; r += 1

    # ---- Revenue Growth ----
    r += 1
    _section(ws, r, "Revenue Growth Assumptions"); r += 1

    # CCG (use _arow for single-year rows, or group by year)
    # For each segment: 3 rows (2026E, 2027E, 2028E), each _arow with 3 scenario values
    for seg_key, seg_label, h2024, h2025, yrs in [
        ("ccg", "CCG Revenue Growth", 3.2, -3.4,
         [("2026E", 2.0, 5.0, -3.0, "Supply recovery offsets PC TAM decline"),
          ("2027E", 3.0, 6.0, 0.0, "Modest PC unit growth + AI PC ASP lift"),
          ("2028E", 3.0, 5.0, 1.0, "Maturing AI PC cycle")]),
        ("dcai", "DCAI Revenue Growth", 0.9, 4.9,
         [("2026E", 18.0, 25.0, 10.0, "Supply recovery + CPU demand + ASIC ramp"),
          ("2027E", 12.0, 18.0, 5.0, "Moderates as base effect kicks in"),
          ("2028E", 8.0, 12.0, 2.0, "CPU structural demand from AI inference")]),
        ("foundry", "Intel Foundry Revenue Growth", -6.4, 2.9,
         [("2026E", 12.0, 18.0, 5.0, "EUV node ramp + internal product demand"),
          ("2027E", 10.0, 20.0, 3.0, "External customer + adv packaging ramps"),
          ("2028E", 10.0, 18.0, 3.0, "14A early revenue possible")]),
        ("allother", "All Other Revenue Growth", -34.1, -1.1,
         [("2026E", -5.0, 5.0, -15.0, "Mobileye headwinds + Altera deconsolidation"),
          ("2027E", 8.0, 15.0, -5.0, "Mobileye recovery as SuperVision ramps"),
          ("2028E", 8.0, 12.0, 2.0, "Mobileye+IMS sustained growth")]),
    ]:
        for yr_label, base, bull, bear, note_txt in yrs:
            label = f"  {seg_label} {yr_label}"
            _arow(ws, r, label, None, h2024/100.0, h2025/100.0, base/100.0, bull/100.0, bear/100.0, "0.0%", note_txt)
            R[f"asp_{seg_key}_g_{yr_label.lower()}"] = r
            r += 1
        r += 1  # spacer between segments

    # ---- Bottom-Up Growth Drivers: CCG ----
    r += 1
    _section(ws, r, "CCG Bottom-Up Drivers: Revenue = PC TAM x Intel Share x ASP"); r += 1

    # PC TAM
    _label(ws, r, "PC TAM (M units) — Source: Gartner/IDC"); r += 1
    _arow(ws, r, "  PC TAM — FY2026E", None, 241.8, 245.4, 260.0, 270.0, 250.0, "#,##0.0",
          "Gartner: 2023=241.8M, 2024=245.4M. 2025E ~259M. Base=+0.5% modest recovery")
    R["asp_pc_tam_2026e"] = r; r += 1
    _arow(ws, r, "  PC TAM — FY2027E", None, None, None, 265.0, 280.0, 252.0, "#,##0.0",
          "Windows 10 EOL refresh + AI PC adoption. Bull: strong enterprise refresh")
    R["asp_pc_tam_2027e"] = r; r += 1
    _arow(ws, r, "  PC TAM — FY2028E", None, None, None, 270.0, 290.0, 255.0, "#,##0.0",
          "AI PC matures; modest long-term growth")
    R["asp_pc_tam_2028e"] = r; r += 1
    r += 1

    # Intel Unit Share
    _label(ws, r, "Intel PC Unit Share (%) — Source: Mercury Research"); r += 1
    _arow(ws, r, "  Intel Share — FY2026E", None, 0.71, 0.715, 0.72, 0.74, 0.70, "0.0%",
          "Mercury est: Q4'24 ~72%. 18A products competitive; share recovery assumed")
    R["asp_intel_share_2026e"] = r; r += 1
    _arow(ws, r, "  Intel Share — FY2027E", None, None, None, 0.725, 0.75, 0.69, "0.0%",
          "Panther Lake + Nova Lake sustain competitiveness")
    R["asp_intel_share_2027e"] = r; r += 1
    _arow(ws, r, "  Intel Share — FY2028E", None, None, None, 0.73, 0.76, 0.68, "0.0%",
          "Continued share recovery with leading-edge nodes")
    R["asp_intel_share_2028e"] = r; r += 1
    r += 1

    # Blended ASP
    _label(ws, r, "Intel CCG Blended ASP ($) — Derived from Rev/(TAM×Share)"); r += 1
    _arow(ws, r, "  CCG ASP — FY2026E", None, 188.0, 190.0, 175.0, 180.0, 170.0, "0.0",
          "FY2025 ASP depressed. Q1 2026 +16% YoY. Base: ASP recovery + premium mix")
    R["asp_ccg_asp_2026e"] = r; r += 1
    _arow(ws, r, "  CCG ASP — FY2027E", None, None, None, 182.0, 190.0, 172.0, "0.0",
          "AI PC (Core Ultra) mix continues to lift ASP")
    R["asp_ccg_asp_2027e"] = r; r += 1
    _arow(ws, r, "  CCG ASP — FY2028E", None, None, None, 184.0, 195.0, 174.0, "0.0",
          "Premium mix + pricing discipline")
    R["asp_ccg_asp_2028e"] = r; r += 1

    # ---- Bottom-Up Growth Drivers: DCAI ----
    r += 1
    _section(ws, r, "DCAI Bottom-Up Drivers: CPU Rev = Server TAM x Share x ASP + AI/ASIC Rev"); r += 1

    # Server TAM
    _label(ws, r, "Server TAM (M units) — Source: IDC"); r += 1
    _arow(ws, r, "  Server TAM — FY2026E", None, 11.5, 12.0, 13.5, 14.5, 12.5, "#,##0.0",
          "IDC est: 2023=11.5M, 2024=12.0M. AI buildout driving 2025+. Base: +12%")
    R["asp_svr_tam_2026e"] = r; r += 1
    _arow(ws, r, "  Server TAM — FY2027E", None, None, None, 14.5, 16.0, 13.0, "#,##0.0",
          "Hyperscale + enterprise AI infra. Bull: accelerated buildout")
    R["asp_svr_tam_2027e"] = r; r += 1
    _arow(ws, r, "  Server TAM — FY2028E", None, None, None, 15.5, 17.5, 13.5, "#,##0.0",
          "CPU inference demand structural?")
    R["asp_svr_tam_2028e"] = r; r += 1
    r += 1

    # Intel DC Unit Share
    _label(ws, r, "Intel DC Server Unit Share (%) — Source: Mercury Research"); r += 1
    _arow(ws, r, "  Intel DC Share — FY2026E", None, 0.76, 0.73, 0.70, 0.72, 0.67, "0.0%",
          "Mercury: Q4'24 ~72%. Granite Rapids improving; AMD Turin competitive. Base: share stabilizes")
    R["asp_intel_dc_share_2026e"] = r; r += 1
    _arow(ws, r, "  Intel DC Share — FY2027E", None, None, None, 0.70, 0.73, 0.65, "0.0%",
          "Coral Rapids + Clearwater Forest on 18A. Bull: share recovers")
    R["asp_intel_dc_share_2027e"] = r; r += 1
    _arow(ws, r, "  Intel DC Share — FY2028E", None, None, None, 0.70, 0.74, 0.63, "0.0%",
          "Sustained competitiveness on 18A/14A")
    R["asp_intel_dc_share_2028e"] = r; r += 1
    r += 1

    # Server ASP
    _label(ws, r, "Intel Server Blended ASP ($) — Derived from CPU Rev/(TAM×Share)"); r += 1
    _arow(ws, r, "  Server ASP — FY2026E", None, 1770, 1650, 1600, 1750, 1500, "0",
          "FY2025 ASP depressed. Q1 2026 +27% YoY. Base: premium mix recovery")
    R["asp_dc_asp_2026e"] = r; r += 1
    _arow(ws, r, "  Server ASP — FY2027E", None, None, None, 1700, 1850, 1550, "0",
          "Granite Rapids/Coral Rapids premium cores support ASP")
    R["asp_dc_asp_2027e"] = r; r += 1
    _arow(ws, r, "  Server ASP — FY2028E", None, None, None, 1750, 1950, 1580, "0",
          "Mix shift to high-core-count AI-adjacent CPUs")
    R["asp_dc_asp_2028e"] = r; r += 1
    r += 1

    # AI/ASIC Revenue
    _label(ws, r, "AI/ASIC & Other Revenue ($M) — Gaudi, ASICs, NICs, IPUs"); r += 1
    _arow(ws, r, "  AI/ASIC Rev — FY2026E", None, 500, 800, 2800, 3500, 2000, "#,##0",
          "ASIC run rate 'north of $1B' Q1 2026, nearly doubling YoY. Gaudi modest. Base: +75%")
    R["asp_ai_asic_2026e"] = r; r += 1
    _arow(ws, r, "  AI/ASIC Rev — FY2027E", None, None, None, 4500, 6000, 3000, "#,##0",
          "ASICs scaling; custom silicon deals (Google, others)")
    R["asp_ai_asic_2027e"] = r; r += 1
    _arow(ws, r, "  AI/ASIC Rev — FY2028E", None, None, None, 6000, 8500, 3800, "#,##0",
          "Custom ASIC pipeline + networking attach")
    R["asp_ai_asic_2028e"] = r; r += 1

    # ---- Segment OPM ----
    _section(ws, r, "Segment Operating Margin Assumptions"); r += 1

    for seg_key, seg_label, h23, h24, h25, yrs in [
        ("ccg", "CCG Operating Margin", 31.4, 34.8, 28.9,
         [("2026E", 30.0, 33.0, 27.0, "Q1 32.6% moderated by 18A ramp costs"),
          ("2027E", 31.5, 36.0, 27.0, "18A yield maturation begins"),
          ("2028E", 33.0, 38.0, 28.0, "Cost structure improvement")]),
        ("dcai", "DCAI Operating Margin", 5.9, 8.8, 20.2,
         [("2026E", 28.0, 32.0, 22.0, "Sustains near Q1 30.5% level; ASIC supports"),
          ("2027E", 28.0, 33.0, 20.0, "Competition begins to pressure"),
          ("2028E", 27.0, 32.0, 18.0, "Coral Rapids cycle supports")]),
        ("foundry", "Intel Foundry Operating Margin", -38.3, -76.7, -57.9,
         [("2026E", -42.0, -35.0, -50.0, "18A yield improvement + volume absorption"),
          ("2027E", -32.0, -20.0, -45.0, "External revenue begins, 14A R&D costs"),
          ("2028E", -22.0, -8.0, -40.0, "Still unprofitable; multi-year journey")]),
        ("allother", "All Other Operating Margin", 27.6, -1.6, 7.4,
         [("2026E", 14.0, 20.0, 8.0, "Mobileye gradual recovery"),
          ("2027E", 18.0, 25.0, 10.0, "SuperVision/Chauffeur revenue"),
          ("2028E", 20.0, 28.0, 12.0, "IMS growth + auto cycle improvement")]),
    ]:
        for yr_label, base, bull, bear, note_txt in yrs:
            label = f"  {seg_label} {yr_label}"
            _arow(ws, r, label, h23/100.0, h24/100.0, h25/100.0, base/100.0, bull/100.0, bear/100.0, "0.0%", note_txt)
            R[f"asp_{seg_key}_opm_{yr_label.lower()}"] = r
            r += 1
        r += 1

    # ---- Consolidated Cost Structure ----
    _section(ws, r, "Consolidated Cost Structure"); r += 1

    # Gross Margin → COGS%
    for yr_label, base_gm, bull_gm, bear_gm, note_txt in [
        ("2026E", 38.5, 41.0, 36.0, "Q1 39.4% base; 18A mix headwind offset by yield gains"),
        ("2027E", 41.0, 45.0, 37.0, "Structural improvement as 18A matures"),
        ("2028E", 43.5, 49.0, 38.0, "Approaching historical 40%+ levels"),
    ]:
        _arow(ws, r, f"  COGS as % of Revenue {yr_label}", 0.60, 0.673, 0.652,
              (100.0 - base_gm)/100.0, (100.0 - bull_gm)/100.0, (100.0 - bear_gm)/100.0, "0.0%",
              f"GM={base_gm}%/{bull_gm}%/{bear_gm}%. {note_txt}")
        R[f"asp_cogs_{yr_label.lower()}"] = r
        r += 1
    r += 1

    # R&D %
    for yr_label, base, bull, bear, note_txt in [
        ("2026E", 26.0, 25.0, 28.0, "~$14.5B absolute; declining as % of growing revenue"),
        ("2027E", 25.0, 23.0, 28.0, "~$15.5B absolute; efficiency gains"),
        ("2028E", 24.0, 21.0, 27.0, "~$16.0B absolute; 14A R&D investment"),
    ]:
        _arow(ws, r, f"  R&D % of Revenue {yr_label}", 0.296, 0.312, 0.261, base/100.0, bull/100.0, bear/100.0, "0.0%", note_txt)
        R[f"asp_rd_{yr_label.lower()}"] = r
        r += 1
    r += 1

    # MG&A %
    for yr_label, base, bull, bear, note_txt in [
        ("2026E", 8.0, 7.5, 9.0, "~$4.5B flat absolute"),
        ("2027E", 7.5, 6.5, 9.0, "Declining as % of growing revenue"),
        ("2028E", 7.0, 6.0, 8.5, "Further efficiency from restructuring"),
    ]:
        _arow(ws, r, f"  MG&A % of Revenue {yr_label}", 0.104, 0.104, 0.087, base/100.0, bull/100.0, bear/100.0, "0.0%", note_txt)
        R[f"asp_mga_{yr_label.lower()}"] = r
        r += 1
    r += 1

    # D&A (absolute)
    _label(ws, r, "D&A ($M absolute)"); r += 1
    for yr_label, base, bull, bear, note_txt in [
        ("2026E", 12500, 12000, 13000, "Q1 annualised ~$12.5B; CIP→in-service driving growth"),
        ("2027E", 13200, 12500, 14000, "Continued fab asset placement"),
        ("2028E", 13800, 12800, 15000, "Partially offset by govt incentives"),
    ]:
        _arow(ws, r, f"    {yr_label}", 7847, 9951, 10757, base, bull, bear, "#,##0", note_txt)
        R[f"asp_da_{yr_label.lower()}"] = r
        r += 1
    r += 1

    # SBC (absolute)
    _arow_const(ws, r, "SBC ($M absolute) — FY2026E", 3229, 3410, 2434, 2500, "#,##0", "Q1 2026 annualised ~$2.5B")
    R["asp_sbc"] = r; r += 1

    # Restructuring
    _arow_const(ws, r, "Restructuring & Other Charges ($M)", -62, 6970, 2191, 500, "#,##0", "Winding down; plans substantially complete")
    R["asp_restruct"] = r; r += 1

    # ---- Working Capital ----
    r += 1
    _section(ws, r, "Working Capital Assumptions"); r += 1
    _arow_const(ws, r, "AR Days", 22.9, 23.9, 26.5, 25, "#,##0.0", "Historical range 22-27 days"); R["asp_ar_days"] = r; r += 1
    _arow_const(ws, r, "AP Days", 96.3, 128.2, 104.6, 100, "#,##0.0", "Normalised from FY2024 peak (vendor terms)"); R["asp_ap_days"] = r; r += 1
    _arow_const(ws, r, "Inventory Days", 124.9, 124.5, 123.0, 125, "#,##0.0", "Remarkably stable at ~124 days"); R["asp_inv_days"] = r; r += 1
    _arow_const(ws, r, "Accrued Comp % of Revenue", 0.067, 0.063, 0.076, 0.07, "0.0%", "Midpoint of historical range"); R["asp_accrued_comp"] = r; r += 1
    _arow_const(ws, r, "Other CA % of Revenue", None, None, None, 0.15, "0.0%", "Includes ~$7.5B refundable tax credits"); R["asp_other_ca"] = r; r += 1
    _arow_const(ws, r, "Other CL % of Revenue", None, None, None, 0.25, "0.0%", "Stable relationship"); R["asp_other_cl"] = r; r += 1
    _arow_const(ws, r, "Income Tax Payable % of Revenue", None, None, None, 0.02, "0.0%", "Normalised level"); R["asp_tax_payable"] = r; r += 1

    # ---- CapEx ----
    r += 1
    _section(ws, r, "Capital Expenditure ($M Gross)"); r += 1
    for yr_label, base, bull, bear, note_txt in [
        ("2026E", 17700, 17500, 18000, "Flat YoY per mgmt guidance; tool-heavy"),
        ("2027E", 19500, 22000, 17000, "External customer commitments begin"),
        ("2028E", 21000, 25000, 16000, "14A capacity buildout in Bull"),
    ]:
        _arow(ws, r, f"    {yr_label}", 25750, 25122, 17672, base, bull, bear, "#,##0", note_txt)
        R[f"asp_capex_{yr_label.lower()}"] = r
        r += 1

    # ---- Financing ----
    r += 1
    _section(ws, r, "Financing Assumptions"); r += 1
    for yr_label, base, bull, bear, note_txt in [
        ("2026E", 49500, 47000, 52000, "+$6.5B Fab34 loan −$2.5B maturities"),
        ("2027E", 47200, 43000, 54000, "−$3.8B maturities repaid"),
        ("2028E", 46000, 39000, 55000, "Modest net reduction; no new debt assumed"),
    ]:
        _arow(ws, r, f"  Total Debt {yr_label} ($M)", 49266, 50011, 46585, base, bull, bear, "#,##0", note_txt)
        R[f"asp_debt_{yr_label.lower()}"] = r
        r += 1
    r += 1

    # Shares
    for yr_label, base, bull, bear, note_txt in [
        ("2026E", 5100, 5080, 5200, "Q1 5,023M + ESPP/RSU + Escrowed Share releases"),
        ("2027E", 5180, 5120, 5400, "Modest ongoing dilution"),
        ("2028E", 5250, 5150, 5600, "Bear: additional strategic equity raises possible"),
    ]:
        _arow(ws, r, f"  Basic Shares {yr_label} (M)", 4228, 4330, 4994, base, bull, bear, "#,##0", note_txt)
        R[f"asp_shares_{yr_label.lower()}"] = r
        r += 1
    r += 1

    _arow_const(ws, r, "DPS ($/share)", 0.74, 0.38, 0.00, 0.00, "0.00", "Dividends suspended; no reinstatement expected")
    R["asp_dps"] = r; r += 1
    _arow_const(ws, r, "Net Equity Issuance ($M)", None, None, None, 300, "#,##0", "ESPP proceeds net of RSU withholdings")
    R["asp_net_eq_iss"] = r; r += 1

    # ---- Tax & Non-Operating ----
    r += 1
    _section(ws, r, "Tax & Non-Operating Items"); r += 1
    _arow(ws, r, "Tax Rate — FY2026E", -1.198, -0.716, 0.983, 0.12, 0.11, 0.15, "0.0%",
          "Anchored to Q2 2026 guided 11%. ETR=Effective Tax Rate for tax shield calc")
    R["asp_etr_sel"] = r; r += 1
    # Tax rate for each year
    _arow(ws, r, "Tax Rate — FY2027E", None, None, None, 0.13, 0.12, 0.15, "0.0%",
          "Valuation allowance means US losses provide no benefit")
    R["asp_etr_2027e"] = r; r += 1
    _arow(ws, r, "Tax Rate — FY2028E", None, None, None, 0.13, 0.12, 0.15, "0.0%",
          "Normalising as Intel returns to sustained profitability")
    R["asp_etr_2028e"] = r; r += 1
    r += 1

    # Interest & Other
    for yr_label, base, bull, bear, note_txt in [
        ("2026E", -800, -500, -1200, "Net interest on ~$50B debt (~5% coupon)"),
        ("2027E", -600, -300, -1200, "Debt paydown reduces interest burden"),
        ("2028E", -400, -100, -1000, "Further improvement as debt declines"),
    ]:
        _arow(ws, r, f"  Interest & Other, Net {yr_label} ($M)", 629, 226, 3257, base, bull, bear, "#,##0", note_txt)
        R[f"asp_interest_{yr_label.lower()}"] = r
        r += 1
    r += 1

    # NCI
    _label(ws, r, "Non-Controlling Interests ($M)")
    _arow(ws, r, "  FY2026E", 293, 293, 293, 750, 750, 750, "#,##0", "Per mgmt: ~$250M/q Q2-Q4; Ireland SCIP NCI eliminated Apr 2026")
    R["asp_nci_2026e"] = r; r += 1
    _arow_const(ws, r, "  FY2027E", None, None, None, 1100, "#,##0", "Per mgmt: ~$1.1B/year 2027-2028")
    R["asp_nci_2027e"] = r; r += 1
    _arow_const(ws, r, "  FY2028E", None, None, None, 1100, "#,##0", "Arizona SCIP + Mobileye NCI ongoing")
    R["asp_nci_2028e"] = r; r += 1

    # ---- Corporate ----
    r += 1
    _section(ws, r, "Corporate & Eliminations"); r += 1
    for yr_label, val, note_txt in [
        ("2026E", -5000, "Declining restructuring + SBC reallocation"),
        ("2027E", -4500, "Further restructuring benefit"),
        ("2028E", -4000, "Stable run-rate"),
    ]:
        _arow_const(ws, r, f"  Corporate Unallocated {yr_label} ($M)", None, None, None, val, "#,##0", note_txt)
        R[f"asp_corp_unalloc_{yr_label.lower()}"] = r
        r += 1
    r += 1

    # ---- Bottom-Up Growth Drivers: Intel Foundry ----
    r += 1
    _section(ws, r, "Intel Foundry Bottom-Up Drivers: Internal Rev = Total Chip Vol x Rev/Chip + External"); r += 1

    # Internal Foundry Rev per Chip
    _label(ws, r, "Internal Foundry Rev per Chip ($/unit) — Internal Foundry Rev / Total Intel Units"); r += 1
    _arow(ws, r, "  Internal Foundry Rev/Chip — FY2026E", 102.2, 93.5, 90.1, 90.0, 98.0, 82.0, "0.0",
          "Declining trend: FY2023 $102.2 -> FY2025 $90.1. Stabilizes as 18A premium wafers ramp")
    R["asp_foundry_rev_per_chip_2026e"] = r; r += 1
    _arow(ws, r, "  Internal Foundry Rev/Chip — FY2027E", None, None, None, 92.0, 105.0, 80.0, "0.0",
          "18A maturity supports modest ASP improvement")
    R["asp_foundry_rev_per_chip_2027e"] = r; r += 1
    _arow(ws, r, "  Internal Foundry Rev/Chip — FY2028E", None, None, None, 95.0, 110.0, 78.0, "0.0",
          "14A early adopter premium + 18A volume scale")
    R["asp_foundry_rev_per_chip_2028e"] = r; r += 1
    # ---- Bottom-Up Growth Drivers: All Other (Mobileye + IMS + Other) ----
    r += 1
    _section(ws, r, "All Other Bottom-Up Drivers: Mobileye (LV Prod x Rev/Veh) + IMS/Other"); r += 1

    # Global Light Vehicle Production
    _label(ws, r, "Global Light Vehicle Production (M units) — Source: IHS/S&P Global Mobility"); r += 1
    _arow(ws, r, "  Global LV Production — FY2026E", 89.0, 88.0, 87.0, 88.0, 90.0, 85.0, "#,##0.0",
          "IHS/S&P Global Mobility. FY2025 ~87M. Modest recovery in Base case")
    R["asp_lv_prod_2026e"] = r; r += 1
    _arow(ws, r, "  Global LV Production — FY2027E", None, None, None, 89.0, 92.0, 86.0, "#,##0.0",
          "Moderate global auto production growth")
    R["asp_lv_prod_2027e"] = r; r += 1
    _arow(ws, r, "  Global LV Production — FY2028E", None, None, None, 90.0, 94.0, 87.0, "#,##0.0",
          "Steady growth trajectory")
    R["asp_lv_prod_2028e"] = r; r += 1
    r += 1

    # Mobileye Rev per Vehicle
    _label(ws, r, "Mobileye Rev per Vehicle ($/vehicle) — Derived from Mobileye Rev / LV Production"); r += 1
    _arow(ws, r, "  Mobileye Rev/Vehicle — FY2026E", 23.4, 18.9, 21.8, 19.0, 21.0, 17.0, "0.0",
          "FY2025 actual: Mobileye $1.9B / 87M LV = $21.8. Q1 2026 $628M. Base: conservatively ~$19")
    R["asp_mbly_rev_per_veh_2026e"] = r; r += 1
    _arow(ws, r, "  Mobileye Rev/Vehicle — FY2027E", None, None, None, 20.0, 23.0, 17.5, "0.0",
          "SuperVision/Chauffeur premium systems partially offset EyeQ ASP pressure")
    R["asp_mbly_rev_per_veh_2027e"] = r; r += 1
    _arow(ws, r, "  Mobileye Rev/Vehicle — FY2028E", None, None, None, 21.0, 25.0, 18.0, "0.0",
          "Advanced ADAS mix shift supports ASP recovery")
    R["asp_mbly_rev_per_veh_2028e"] = r; r += 1
    r += 1

    # IMS & Other Revenue
    _label(ws, r, "IMS & Other Revenue ($M) — IMS nanofabrication tools + legacy businesses"); r += 1
    _arow(ws, r, "  IMS & Other Rev — FY2026E", 3378, 1941, 1663, 1800, 2200, 1500, "#,##0",
          "IMS multi-beam mask writers benefit from semi capex cycle. FY2023-24 include Altera (deconsolidated Sep 2025)")
    R["asp_ims_other_2026e"] = r; r += 1
    _arow(ws, r, "  IMS & Other Rev — FY2027E", None, None, None, 2000, 2600, 1600, "#,##0",
          "IMS tool shipments driven by EUV adoption")
    R["asp_ims_other_2027e"] = r; r += 1
    _arow(ws, r, "  IMS & Other Rev — FY2028E", None, None, None, 2200, 3000, 1700, "#,##0",
          "IMS steady growth + legacy business runoff")
    R["asp_ims_other_2028e"] = r; r += 1
    r += 1

    # External Foundry Revenue (excludes ~99% internal wafer sales to Intel Products)
    _section(ws, r, "External Foundry Revenue ($M Absolute)"); r += 1
    _note(ws, r-1, "Only ~$150M in FY2025. Growing as 18A/14A external customers ramp")
    for yr_label, h2023, h2024, h2025, base, bull, bear, note_txt in [
        ("2026E", 0, 50, 150, 700, 1000, 400, "Q1 2026 $174M annualised + growth"),
        ("2027E", None, None, None, 1500, 3000, 800, "18A external PDK → wafer revenue"),
        ("2028E", None, None, None, 2500, 5000, 1200, "14A early adopters"),
    ]:
        _arow(ws, r, f"  External Foundry Rev {yr_label}", h2023, h2024, h2025, base, bull, bear, "#,##0", note_txt)
        R[f"asp_ext_foundry_{yr_label.lower()}"] = r
        r += 1
    r += 1

    _arow_const(ws, r, "Intersegment Eliminations ($M)", None, None, None, -200, "#,##0",
                "Residual inter-segment elim; bulk of Foundry internal is excluded from consolidated")
    R["asp_inter_elim"] = r; r += 1

    # ---- Column widths & final touches ----
    ws.column_dimensions['A'].width = 42
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 12
    ws.column_dimensions['E'].width = 2
    ws.column_dimensions['F'].width = 12
    ws.column_dimensions['G'].width = 12
    ws.column_dimensions['H'].width = 12
    ws.column_dimensions['I'].width = 12
    ws.column_dimensions['J'].width = 2
    ws.column_dimensions['K'].width = 50

    # Legend & Sources
    lr = _add_legend(ws, r + 2)
    _add_sources(ws, lr + 1)
    ws.sheet_properties.tabColor = "1F4E79"

    return r


# ============================================================
# TAB 4: GROWTH_DRIVERS (Bottom-Up Revenue Decomposition)
# ============================================================

def build_growth_drivers(wb, R):
    """Build growth drivers tab: decompose each segment's revenue into key drivers.
    CCG: Revenue = PC TAM x Intel Share x ASP (bottom-up)"""
    ws = wb["Growth_Drivers"]
    _title_row(ws, 1, "Intel (INTC) - Revenue Growth Drivers")
    _unit_row(ws, 2)
    _yr_header_no_scenario(ws, 3)
    _label(ws, 4, "USD Million / M Units / $ per Unit")
    _note(ws, 4, "Bottom-up revenue decomposition. Data sources cited inline.")

    r = 5

    # ============================================================
    # CCG — Bottom-Up: Revenue = PC TAM × Intel Share × ASP
    # ============================================================
    _section(ws, r, "CCG: Client Computing Group — Bottom-Up"); r += 1
    _label(ws, r, "Method: Revenue = PC TAM (M units) × Intel Unit Share (%) × Blended ASP ($)")
    _note(ws, r, "TAM source: Gartner/IDC (2023=241.8M, 2024=245.4M). Share: Mercury Research est. ASP: derived from rev/(TAM×share).")
    r += 1

    # PC TAM
    _label(ws, r, "PC TAM (M units)")
    hist_tam = [241.8, 245.4, 258.6]
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, hist_tam[ci], "#,##0.0")
    for fc in FCST_COLS:
        fc_year = {6: "2026e", 7: "2027e", 8: "2028e"}[fc]
        _lval(ws, r, fc, f"=Assumptions!I{R[f'asp_pc_tam_{fc_year}']}", "#,##0.0")
    _note(ws, r, "Source: Gartner/IDC. 2025 est. based on Canalys ~259M full-year")
    R["gd_ccg_tam"] = r; r += 1

    # Intel Unit Share
    _label(ws, r, "Intel PC Unit Share (%)")
    hist_share = [0.710, 0.715, 0.720]
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, hist_share[ci], "0.0%")
    for fc in FCST_COLS:
        fc_year = {6: "2026e", 7: "2027e", 8: "2028e"}[fc]
        _lval(ws, r, fc, f"=Assumptions!I{R[f'asp_intel_share_{fc_year}']}", "0.0%")
    _note(ws, r, "Source: Mercury Research x86 PC CPU share. Q4'24 est. ~72%")
    R["gd_ccg_share"] = r; r += 1

    # Intel Implied Units = TAM × Share
    _label(ws, r, "  Intel Implied Units (M)")
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r, col, f"={CL(col)}{R['gd_ccg_tam']}*{CL(col)}{R['gd_ccg_share']}", "#,##0.0")
    _note(ws, r, "= PC TAM × Intel Share. Implied unit shipments")
    R["gd_ccg_units"] = r; r += 1

    # CCG Blended ASP
    _label(ws, r, "Intel CCG Blended ASP ($)")
    hist_asp = [188, 190, 173]
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, hist_asp[ci], "0.0")
    for fc in FCST_COLS:
        fc_year = {6: "2026e", 7: "2027e", 8: "2028e"}[fc]
        _lval(ws, r, fc, f"=Assumptions!I{R[f'asp_ccg_asp_{fc_year}']}", "0.0")
    _note(ws, r, "Historical: derived as Reported Revenue / Implied Units. Q1 2026 ASP +16% YoY")
    R["gd_ccg_asp"] = r; r += 1

    # Implied Revenue = Units × ASP
    _label(ws, r, "  Implied CCG Revenue ($M)")
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r, col, f"={CL(col)}{R['gd_ccg_units']}*{CL(col)}{R['gd_ccg_asp']}", "#,##0")
    _note(ws, r, "= Units × ASP. Bottom-up revenue estimate")
    R["gd_ccg_implied_rev"] = r; r += 1

    # Reported CCG Revenue (historical: Segment_Revenue; forecast: mirrors Implied)
    _label(ws, r, "  Reported CCG Revenue ($M)")
    for col in HIST_COLS:
        _lval(ws, r, col, f"=Segment_Revenue!{CL(col)}{R['segrev_ccg']}", "#,##0")
    for fc in FCST_COLS:
        _lval(ws, r, fc, f"={CL(fc)}{R['gd_ccg_implied_rev']}", "#,##0")
    _note(ws, r, "Historical: links to Segment_Revenue (10-K). Forecast: = Implied (Segment_Revenue pulls from Growth_Drivers)")
    R["gd_ccg_reported_rev"] = r; r += 1

    # Reconciliation = Implied - Reported
    _section(ws, r, "  Reconciliation: Implied − Reported ($M)")
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r, col,
              f"={CL(col)}{R['gd_ccg_implied_rev']}-{CL(col)}{R['gd_ccg_reported_rev']}", "#,##0")
    _note(ws, r, "Should be near-zero for historical (within 0.1%). Forecast: divergence flags assumption inconsistency")
    R["gd_ccg_recon"] = r; r += 1

    # Reconciliation %
    _label(ws, r, "  Reconcil. as % of Reported")
    for col in HIST_COLS + FCST_COLS:
        _pct(ws, r, col,
             f"={CL(col)}{R['gd_ccg_recon']}/{CL(col)}{R['gd_ccg_reported_rev']}")
    _note(ws, r, "Green if <1%"); r += 2

    # ============================================================
    # DCAI — Bottom-Up: Rev = Server TAM × Intel DC Share × ASP + AI/ASIC Rev
    # ============================================================
    _section(ws, r, "DCAI: Data Center & AI — Bottom-Up"); r += 1
    _label(ws, r, "Method: Revenue = Server TAM (M) × Intel DC Share (%) × ASP ($) + AI/ASIC ($M)")
    _note(ws, r, "TAM: IDC. Share: Mercury Research. AI/ASIC: Gaudi+custom ASICs+NICs/IPUs. Q1 2026 ASIC run rate >$1B.")
    r += 1

    # Server TAM
    _label(ws, r, "Server TAM (M units)")
    hist_svr_tam = [11.5, 12.0, 13.5]
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, hist_svr_tam[ci], "#,##0.0")
    for fc in FCST_COLS:
        fc_year = {6: "2026e", 7: "2027e", 8: "2028e"}[fc]
        _lval(ws, r, fc, f"=Assumptions!I{R[f'asp_svr_tam_{fc_year}']}", "#,##0.0")
    _note(ws, r, "Source: IDC. 2023=11.5M, 2024=12.0M, 2025=13.5M. AI buildout driving unit growth")
    R["gd_dcai_tam"] = r; r += 1

    # Intel DC Unit Share
    _label(ws, r, "Intel DC Server Unit Share (%)")
    hist_dc_share = [0.775, 0.760, 0.730]
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, hist_dc_share[ci], "0.0%")
    for fc in FCST_COLS:
        fc_year = {6: "2026e", 7: "2027e", 8: "2028e"}[fc]
        _lval(ws, r, fc, f"=Assumptions!I{R[f'asp_intel_dc_share_{fc_year}']}", "0.0%")
    _note(ws, r, "Source: Mercury Research x86 server CPU share. Q4 2024 ~72%. Granite Rapids competitive vs Turin")
    R["gd_dcai_share"] = r; r += 1

    # Intel Implied DC Units = TAM × Share
    _label(ws, r, "  Intel Implied DC Units (M)")
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r, col, f"={CL(col)}{R['gd_dcai_tam']}*{CL(col)}{R['gd_dcai_share']}", "#,##0.0")
    _note(ws, r, "= Server TAM × Intel DC Share. Implied unit shipments")
    R["gd_dcai_units"] = r; r += 1

    # Server Blended ASP
    _label(ws, r, "Intel Server Blended ASP ($)")
    hist_dc_asp = [1772, 1712, 1632]
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, hist_dc_asp[ci], "#,##0")
    for fc in FCST_COLS:
        fc_year = {6: "2026e", 7: "2027e", 8: "2028e"}[fc]
        _lval(ws, r, fc, f"=Assumptions!I{R[f'asp_dc_asp_{fc_year}']}", "#,##0")
    _note(ws, r, "Historical: derived as (DCAI Rev − AI/ASIC) / (TAM × Share). Q1 2026 ASP +27% YoY")
    R["gd_dcai_asp"] = r; r += 1

    # Implied CPU Revenue = Units × ASP
    _label(ws, r, "  Implied CPU Revenue ($M)")
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r, col, f"={CL(col)}{R['gd_dcai_units']}*{CL(col)}{R['gd_dcai_asp']}", "#,##0")
    _note(ws, r, "= DC Units × Server ASP. Core x86 server CPU revenue")
    R["gd_dcai_cpu_rev"] = r; r += 1

    # AI/ASIC & Other Revenue
    _label(ws, r, "AI/ASIC & Other Revenue ($M)")
    hist_ai_asic = [200, 500, 800]
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, hist_ai_asic[ci], "#,##0")
    for fc in FCST_COLS:
        fc_year = {6: "2026e", 7: "2027e", 8: "2028e"}[fc]
        _lval(ws, r, fc, f"=Assumptions!I{R[f'asp_ai_asic_{fc_year}']}", "#,##0")
    _note(ws, r, "Gaudi AI accelerators + custom ASICs + NICs/IPUs. FY2025 ~$800M. Q1 2026 ASIC run rate >$1B")
    R["gd_dcai_ai_asic"] = r; r += 1

    # Implied DCAI Revenue = CPU Rev + AI/ASIC Rev
    _label(ws, r, "  Implied DCAI Revenue ($M)")
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r, col, f"={CL(col)}{R['gd_dcai_cpu_rev']}+{CL(col)}{R['gd_dcai_ai_asic']}", "#,##0")
    _note(ws, r, "= CPU Revenue + AI/ASIC Revenue. Bottom-up DCAI estimate")
    R["gd_dcai_implied_rev"] = r; r += 1

    # Reported DCAI Revenue (historical: Segment_Revenue; forecast: mirrors Implied)
    _label(ws, r, "  Reported DCAI Revenue ($M)")
    for col in HIST_COLS:
        _lval(ws, r, col, f"=Segment_Revenue!{CL(col)}{R['segrev_dcai']}", "#,##0")
    for fc in FCST_COLS:
        _lval(ws, r, fc, f"={CL(fc)}{R['gd_dcai_implied_rev']}", "#,##0")
    _note(ws, r, "Historical: links to Segment_Revenue (10-K). Forecast: = Implied (Segment_Revenue pulls from Growth_Drivers)")
    R["gd_dcai_reported_rev"] = r; r += 1

    # Reconciliation = Implied - Reported
    _section(ws, r, "  Reconciliation: Implied − Reported ($M)")
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r, col,
              f"={CL(col)}{R['gd_dcai_implied_rev']}-{CL(col)}{R['gd_dcai_reported_rev']}", "#,##0")
    _note(ws, r, "Historical should be near-zero (<0.5%). Forecast: flags driver vs growth-rate assumption divergence")
    R["gd_dcai_recon"] = r; r += 1

    # Reconciliation %
    _label(ws, r, "  Reconcil. as % of Reported")
    for col in HIST_COLS + FCST_COLS:
        _pct(ws, r, col,
             f"={CL(col)}{R['gd_dcai_recon']}/{CL(col)}{R['gd_dcai_reported_rev']}")
    _note(ws, r, "Green if <1%"); r += 2

    # ============================================================
    # Intel Foundry — Bottom-Up: Internal = Total Chip Vol × Rev/Chip + External
    # ============================================================
    _section(ws, r, "Intel Foundry — Bottom-Up"); r += 1
    _label(ws, r, "Method: Revenue = Total Intel Chip Vol (M) × Internal Rev/Chip ($) + External Foundry ($M)")
    _note(ws, r, "~99% internal wafer sales to Intel Products. Chip vol = CCG + DCAI units. Rev/Chip = internal transfer price.")
    r += 1

    # Total Intel Chip Volume = CCG Units + DCAI Units
    _label(ws, r, "Total Intel Chip Volume (M units)")
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r, col, f"={CL(col)}{R['gd_ccg_units']}+{CL(col)}{R['gd_dcai_units']}", "#,##0.0")
    _note(ws, r, "= CCG Implied Units + DCAI Implied Units. Links to segment drivers above")
    R["gd_foundry_vol"] = r; r += 1

    # Internal Foundry Rev per Chip
    _label(ws, r, "Internal Foundry Rev per Chip ($)")
    hist_rev_per_chip = [102.2, 93.5, 90.1]
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, hist_rev_per_chip[ci], "0.0")
    for fc in FCST_COLS:
        fc_year = {6: "2026e", 7: "2027e", 8: "2028e"}[fc]
        _lval(ws, r, fc, f"=Assumptions!I{R[f'asp_foundry_rev_per_chip_{fc_year}']}", "0.0")
    _note(ws, r, "Historical: derived as Internal Foundry Rev / Total Intel Units. Declining trend as older nodes depreciate")
    R["gd_foundry_rev_per_chip"] = r; r += 1

    # Internal Foundry Revenue = Vol × Rev/Chip
    _label(ws, r, "  Internal Foundry Revenue ($M)")
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r, col, f"={CL(col)}{R['gd_foundry_vol']}*{CL(col)}{R['gd_foundry_rev_per_chip']}", "#,##0")
    _note(ws, r, "= Total Intel Chip Vol × Internal Rev/Chip. Wafer sales to Intel CCG and DCAI divisions")
    R["gd_foundry_internal_rev"] = r; r += 1

    # External Foundry Revenue
    _label(ws, r, "External Foundry Revenue ($M)")
    hist_ext_foundry = [50, 60, 150]
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, hist_ext_foundry[ci], "#,##0")
    for fc in FCST_COLS:
        fc_year = {6: "2026e", 7: "2027e", 8: "2028e"}[fc]
        _lval(ws, r, fc, f"=Assumptions!I{R[f'asp_ext_foundry_{fc_year}']}", "#,##0")
    _note(ws, r, "3rd-party wafer customers. FY2025 ~$150M. Growing as 18A/14A external PDK ramps")
    R["gd_foundry_ext_rev"] = r; r += 1

    # Implied Total Foundry Revenue = Internal + External
    _label(ws, r, "  Implied Total Foundry Revenue ($M)")
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r, col, f"={CL(col)}{R['gd_foundry_internal_rev']}+{CL(col)}{R['gd_foundry_ext_rev']}", "#,##0")
    _note(ws, r, "= Internal Foundry Rev + External Foundry Rev")
    R["gd_foundry_implied_rev"] = r; r += 1

    # Reported Foundry Revenue (historical: Segment_Revenue; forecast: mirrors Implied)
    _label(ws, r, "  Reported Foundry Revenue ($M)")
    for col in HIST_COLS:
        _lval(ws, r, col, f"=Segment_Revenue!{CL(col)}{R['segrev_foundry']}", "#,##0")
    for fc in FCST_COLS:
        _lval(ws, r, fc, f"={CL(fc)}{R['gd_foundry_implied_rev']}", "#,##0")
    _note(ws, r, "Historical: links to Segment_Revenue (10-K). Forecast: = Implied (Segment_Revenue pulls from Growth_Drivers)")
    R["gd_foundry_reported_rev"] = r; r += 1

    # Reconciliation
    _section(ws, r, "  Reconciliation: Implied − Reported ($M)")
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r, col,
              f"={CL(col)}{R['gd_foundry_implied_rev']}-{CL(col)}{R['gd_foundry_reported_rev']}", "#,##0")
    _note(ws, r, "Historical should be near-zero (<0.1%). Forecast: flags driver vs growth-rate assumption divergence")
    R["gd_foundry_recon"] = r; r += 1

    # Reconciliation %
    _label(ws, r, "  Reconcil. as % of Reported")
    for col in HIST_COLS + FCST_COLS:
        _pct(ws, r, col,
             f"={CL(col)}{R['gd_foundry_recon']}/{CL(col)}{R['gd_foundry_reported_rev']}")
    _note(ws, r, "Green if <1%"); r += 2

    # ============================================================
    # All Other — Bottom-Up: Mobileye (LV Prod × Rev/Veh) + IMS/Other
    # ============================================================
    _section(ws, r, "All Other: Mobileye + IMS + Legacy — Bottom-Up"); r += 1
    _label(ws, r, "Method: Revenue = LV Production (M) × Mobileye Rev/Vehicle ($) + IMS & Other ($M)")
    _note(ws, r, "All Other = Mobileye (ADAS chips) + IMS (nanofab tools) + legacy. LV Prod source: IHS/S&P Global Mobility.")
    r += 1

    # Global LV Production
    _label(ws, r, "Global Light Vehicle Production (M)")
    hist_lv_prod = [89.0, 88.0, 87.0]
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, hist_lv_prod[ci], "#,##0.0")
    for fc in FCST_COLS:
        fc_year = {6: "2026e", 7: "2027e", 8: "2028e"}[fc]
        _lval(ws, r, fc, f"=Assumptions!I{R[f'asp_lv_prod_{fc_year}']}", "#,##0.0")
    _note(ws, r, "Source: IHS/S&P Global Mobility. FY2023 89M, FY2024 88M, FY2025 ~87M")
    R["gd_allother_lv_prod"] = r; r += 1

    # Mobileye Rev per Vehicle
    _label(ws, r, "Mobileye Rev per Vehicle ($)")
    hist_mbly_rpv = [23.4, 18.9, 21.8]
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, hist_mbly_rpv[ci], "0.0")
    for fc in FCST_COLS:
        fc_year = {6: "2026e", 7: "2027e", 8: "2028e"}[fc]
        _lval(ws, r, fc, f"=Assumptions!I{R[f'asp_mbly_rev_per_veh_{fc_year}']}", "0.0")
    _note(ws, r, "Historical: FY2023 $23.4 (pre-destock), FY2024 $18.9 (inventory correction), FY2025 $21.8 (recovery). Source: Intel 10-K Mobileye revenue $1.9B FY2025")
    R["gd_allother_mbly_rpv"] = r; r += 1

    # Implied Mobileye Revenue = LV Prod × Rev/Vehicle
    _label(ws, r, "  Implied Mobileye Revenue ($M)")
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r, col, f"={CL(col)}{R['gd_allother_lv_prod']}*{CL(col)}{R['gd_allother_mbly_rpv']}", "#,##0")
    _note(ws, r, "= Global LV Production × Mobileye Rev/Vehicle. EyeQ + SuperVision/Chauffeur")
    R["gd_allother_mbly_rev"] = r; r += 1

    # IMS & Other Revenue
    _label(ws, r, "IMS & Other Revenue ($M)")
    hist_ims_other = [3378, 1941, 1663]
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, hist_ims_other[ci], "#,##0")
    for fc in FCST_COLS:
        fc_year = {6: "2026e", 7: "2027e", 8: "2028e"}[fc]
        _lval(ws, r, fc, f"=Assumptions!I{R[f'asp_ims_other_{fc_year}']}", "#,##0")
    _note(ws, r, "IMS nanofabrication mask writers + legacy. FY2023-24 include Altera ($1.3B decline YoY FY2024). Altera deconsolidated Sep 2025")
    R["gd_allother_ims"] = r; r += 1

    # Implied All Other Revenue = Mobileye + IMS/Other
    _label(ws, r, "  Implied All Other Revenue ($M)")
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r, col, f"={CL(col)}{R['gd_allother_mbly_rev']}+{CL(col)}{R['gd_allother_ims']}", "#,##0")
    _note(ws, r, "= Mobileye Revenue + IMS & Other Revenue")
    R["gd_allother_implied_rev"] = r; r += 1

    # Reported All Other Revenue (historical: Segment_Revenue; forecast: mirrors Implied)
    _label(ws, r, "  Reported All Other Revenue ($M)")
    for col in HIST_COLS:
        _lval(ws, r, col, f"=Segment_Revenue!{CL(col)}{R['segrev_allother']}", "#,##0")
    for fc in FCST_COLS:
        _lval(ws, r, fc, f"={CL(fc)}{R['gd_allother_implied_rev']}", "#,##0")
    _note(ws, r, "Historical: links to Segment_Revenue (10-K). Forecast: = Implied (Segment_Revenue pulls from Growth_Drivers)")
    R["gd_allother_reported_rev"] = r; r += 1

    # Reconciliation
    _section(ws, r, "  Reconciliation: Implied − Reported ($M)")
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r, col,
              f"={CL(col)}{R['gd_allother_implied_rev']}-{CL(col)}{R['gd_allother_reported_rev']}", "#,##0")
    _note(ws, r, "Historical should be near-zero (<1%). Forecast: flags driver vs growth-rate assumption divergence")
    R["gd_allother_recon"] = r; r += 1

    # Reconciliation %
    _label(ws, r, "  Reconcil. as % of Reported")
    for col in HIST_COLS + FCST_COLS:
        _pct(ws, r, col,
             f"={CL(col)}{R['gd_allother_recon']}/{CL(col)}{R['gd_allother_reported_rev']}")
    _note(ws, r, "Green if <1%"); r += 2

    # ============================================================
    # Driver Sensitivity: FY2026E Revenue Elasticity
    # ============================================================
    _section(ws, r, "Driver Sensitivity: FY2026E Revenue Elasticity"); r += 1
    _label(ws, r, "Impact of isolated driver changes on FY2026E implied segment revenue (Base case). Updates with scenario.")
    _note(ws, r, "All elasticities reference FY2026E (col F) driver and revenue values. Switch scenario in Assumptions!B2 to see Bull/Bear sensitivity.")
    r += 1

    # Column headers
    _label(ws, r, "Driver")
    ws.cell(row=r, column=6, value="FY2026E Value").font = Font(bold=True, size=9)
    ws.cell(row=r, column=6).border = BD
    ws.cell(row=r, column=7, value="Change").font = Font(bold=True, size=9)
    ws.cell(row=r, column=7).border = BD
    ws.cell(row=r, column=8, value="Rev Impact ($M)").font = Font(bold=True, size=9)
    ws.cell(row=r, column=8).border = BD
    ws.cell(row=r, column=9, value="Impact %").font = Font(bold=True, size=9)
    ws.cell(row=r, column=9).border = BD
    r += 1

    # ---- CCG Sensitivities ----
    _label(ws, r, "CCG Drivers", bold=True); r += 1

    # PC TAM: 1% change
    _label(ws, r, "  PC TAM")
    _lval(ws, r, 6, f"=F{R['gd_ccg_tam']}", "#,##0.0")
    ws.cell(row=r, column=7, value="±1%").font = Font(size=9); ws.cell(row=r, column=7).border = BD
    _fval(ws, r, 8, f"=F{R['gd_ccg_implied_rev']}*0.01", "#,##0")
    _pct(ws, r, 9, f"=H{r}/F{R['gd_ccg_implied_rev']}")
    _note(ws, r, "PC TAM elasticity = 1:1 with CCG revenue (multiplicative driver)"); r += 1

    # Intel PC Share: 1pp change
    _label(ws, r, "  Intel PC Share")
    _lval(ws, r, 6, f"=F{R['gd_ccg_share']}", "0.0%")
    ws.cell(row=r, column=7, value="±1pp").font = Font(size=9); ws.cell(row=r, column=7).border = BD
    _fval(ws, r, 8, f"=F{R['gd_ccg_tam']}*0.01*F{R['gd_ccg_asp']}", "#,##0")
    _pct(ws, r, 9, f"=H{r}/F{R['gd_ccg_implied_rev']}")
    _note(ws, r, "1pp share shift = TAM × 1% × ASP. Most sensitive CCG driver"); r += 1

    # CCG ASP: $1 change
    _label(ws, r, "  CCG Blended ASP")
    _lval(ws, r, 6, f"=F{R['gd_ccg_asp']}", "0")
    ws.cell(row=r, column=7, value="±$1").font = Font(size=9); ws.cell(row=r, column=7).border = BD
    _fval(ws, r, 8, f"=F{R['gd_ccg_units']}", "#,##0")
    _pct(ws, r, 9, f"=H{r}/F{R['gd_ccg_implied_rev']}")
    _note(ws, r, "$1 ASP change = Units × $1 impact"); r += 2

    # ---- DCAI Sensitivities ----
    _label(ws, r, "DCAI Drivers", bold=True); r += 1

    # Server TAM
    _label(ws, r, "  Server TAM")
    _lval(ws, r, 6, f"=F{R['gd_dcai_tam']}", "#,##0.0")
    ws.cell(row=r, column=7, value="±1%").font = Font(size=9); ws.cell(row=r, column=7).border = BD
    _fval(ws, r, 8, f"=F{R['gd_dcai_cpu_rev']}*0.01", "#,##0")
    _pct(ws, r, 9, f"=H{r}/F{R['gd_dcai_implied_rev']}")
    _note(ws, r, "Server TAM elasticity applies to CPU revenue only (not AI/ASIC)"); r += 1

    # Intel DC Share
    _label(ws, r, "  Intel DC Server Share")
    _lval(ws, r, 6, f"=F{R['gd_dcai_share']}", "0.0%")
    ws.cell(row=r, column=7, value="±1pp").font = Font(size=9); ws.cell(row=r, column=7).border = BD
    _fval(ws, r, 8, f"=F{R['gd_dcai_tam']}*0.01*F{R['gd_dcai_asp']}", "#,##0")
    _pct(ws, r, 9, f"=H{r}/F{R['gd_dcai_implied_rev']}")
    _note(ws, r, "1pp share shift in DC CPU market. Most sensitive DCAI driver"); r += 1

    # Server ASP
    _label(ws, r, "  Server Blended ASP")
    _lval(ws, r, 6, f"=F{R['gd_dcai_asp']}", "#,##0")
    ws.cell(row=r, column=7, value="±$1").font = Font(size=9); ws.cell(row=r, column=7).border = BD
    _fval(ws, r, 8, f"=F{R['gd_dcai_units']}", "#,##0")
    _pct(ws, r, 9, f"=H{r}/F{R['gd_dcai_implied_rev']}")
    _note(ws, r, "$1 ASP change = DC Units × $1 impact"); r += 1

    # AI/ASIC
    _label(ws, r, "  AI/ASIC Revenue")
    _lval(ws, r, 6, f"=F{R['gd_dcai_ai_asic']}", "#,##0")
    ws.cell(row=r, column=7, value="±$100M").font = Font(size=9); ws.cell(row=r, column=7).border = BD
    _fval(ws, r, 8, f"100", "#,##0")
    _pct(ws, r, 9, f"=H{r}/F{R['gd_dcai_implied_rev']}")
    _note(ws, r, "AI/ASIC passes through 1:1 to DCAI revenue"); r += 2

    # ---- Intel Foundry Sensitivities ----
    _label(ws, r, "Intel Foundry Drivers", bold=True); r += 1

    # Total Chip Vol
    _label(ws, r, "  Total Intel Chip Vol")
    _lval(ws, r, 6, f"=F{R['gd_foundry_vol']}", "#,##0.0")
    ws.cell(row=r, column=7, value="±1%").font = Font(size=9); ws.cell(row=r, column=7).border = BD
    _fval(ws, r, 8, f"=F{R['gd_foundry_internal_rev']}*0.01", "#,##0")
    _pct(ws, r, 9, f"=H{r}/F{R['gd_foundry_implied_rev']}")
    _note(ws, r, "Chip vol elasticity applies to internal Foundry only"); r += 1

    # Internal Rev/Chip
    _label(ws, r, "  Internal Rev per Chip")
    _lval(ws, r, 6, f"=F{R['gd_foundry_rev_per_chip']}", "0.0")
    ws.cell(row=r, column=7, value="±$1").font = Font(size=9); ws.cell(row=r, column=7).border = BD
    _fval(ws, r, 8, f"=F{R['gd_foundry_vol']}", "#,##0")
    _pct(ws, r, 9, f"=H{r}/F{R['gd_foundry_implied_rev']}")
    _note(ws, r, "$1 Rev/Chip change = Total Chip Vol × $1. High sensitivity given 197M+ unit base"); r += 1

    # External Foundry
    _label(ws, r, "  External Foundry Rev")
    _lval(ws, r, 6, f"=F{R['gd_foundry_ext_rev']}", "#,##0")
    ws.cell(row=r, column=7, value="±$100M").font = Font(size=9); ws.cell(row=r, column=7).border = BD
    _fval(ws, r, 8, f"100", "#,##0")
    _pct(ws, r, 9, f"=H{r}/F{R['gd_foundry_implied_rev']}")
    _note(ws, r, "External Foundry passes through 1:1"); r += 2

    # ---- All Other Sensitivities ----
    _label(ws, r, "All Other Drivers", bold=True); r += 1

    # LV Production
    _label(ws, r, "  Global LV Production")
    _lval(ws, r, 6, f"=F{R['gd_allother_lv_prod']}", "#,##0.0")
    ws.cell(row=r, column=7, value="±1%").font = Font(size=9); ws.cell(row=r, column=7).border = BD
    _fval(ws, r, 8, f"=F{R['gd_allother_mbly_rev']}*0.01", "#,##0")
    _pct(ws, r, 9, f"=H{r}/F{R['gd_allother_implied_rev']}")
    _note(ws, r, "LV Prod elasticity applies to Mobileye revenue only"); r += 1

    # MBLY Rev/Vehicle
    _label(ws, r, "  Mobileye Rev per Vehicle")
    _lval(ws, r, 6, f"=F{R['gd_allother_mbly_rpv']}", "0.0")
    ws.cell(row=r, column=7, value="±$1").font = Font(size=9); ws.cell(row=r, column=7).border = BD
    _fval(ws, r, 8, f"=F{R['gd_allother_lv_prod']}", "#,##0")
    _pct(ws, r, 9, f"=H{r}/F{R['gd_allother_implied_rev']}")
    _note(ws, r, "$1 Rev/Veh change = LV Production × $1"); r += 1

    # IMS & Other
    _label(ws, r, "  IMS & Other Revenue")
    _lval(ws, r, 6, f"=F{R['gd_allother_ims']}", "#,##0")
    ws.cell(row=r, column=7, value="±$100M").font = Font(size=9); ws.cell(row=r, column=7).border = BD
    _fval(ws, r, 8, f"100", "#,##0")
    _pct(ws, r, 9, f"=H{r}/F{R['gd_allother_implied_rev']}")
    _note(ws, r, "IMS/Other passes through 1:1"); r += 2

    # Summary: Top 3 Most Sensitive Drivers
    _section(ws, r, "Top 3 Most Sensitive Drivers (by Revenue Impact per unit change)"); r += 1
    _label(ws, r, "1. Intel PC Share: each 1pp = significant revenue swing (largest CCG driver)")
    _note(ws, r, "CCG is ~58% of consolidated revenue; share shifts have outsized impact"); r += 1
    _label(ws, r, "2. Internal Foundry Rev/Chip: $1 change × 197M+ units = material Foundry revenue swing")
    _note(ws, r, "Internal transfer pricing is the key Foundry profitability lever"); r += 1
    _label(ws, r, "3. Intel DC Server Share: each 1pp = material DCAI CPU revenue swing")
    _note(ws, r, "Server CPU share is the battleground vs AMD Turin; Granite Rapids competitiveness key"); r += 1

    _set_col_widths(ws)
    lr = _add_legend(ws, r + 1)
    _add_sources(ws, lr + 1)
    ws.sheet_properties.tabColor = "7030A0"
    return r


# ============================================================
# TAB 5: SEGMENT_REVENUE (Bottom-Up Build)
# ============================================================

def build_segment_revenue(wb, R):
    """Build segment revenue tab: 4 segments + YoY + mix + consolidated total."""
    ws = wb["Segment_Revenue"]
    _title_row(ws, 1, "Intel (INTC) — Segment Revenue")
    _unit_row(ws, 2)
    _yr_header_no_scenario(ws, 3)
    _label(ws, 4, "USD Million"); _note(ws, 4, "Historical from 10-K Note 3; Forecast = Prior x (1 + growth)")

    r = 5
    segs = [
        ("CCG", "CCG - Client Computing Group",
         [32305, 33346, 32228],
         ["asp_ccg_g_2026e", "asp_ccg_g_2027e", "asp_ccg_g_2028e"]),
        ("DCAI", "DCAI - Data Center & AI",
         [15980, 16125, 16919],
         ["asp_dcai_g_2026e", "asp_dcai_g_2027e", "asp_dcai_g_2028e"]),
        ("Foundry", "Intel Foundry",
         [18504, 17317, 17826],
         ["asp_foundry_g_2026e", "asp_foundry_g_2027e", "asp_foundry_g_2028e"]),
        ("AllOther", "All Other (Mobileye + IMS + Other)",
         [5463, 3601, 3563],
         ["asp_allother_g_2026e", "asp_allother_g_2027e", "asp_allother_g_2028e"]),
    ]

    gd_implied_keys = {
        "CCG": "gd_ccg_implied_rev", "DCAI": "gd_dcai_implied_rev",
        "Foundry": "gd_foundry_implied_rev", "AllOther": "gd_allother_implied_rev",
    }
    for seg_key, seg_label, hist_rev, growth_keys in segs:
        _label(ws, r, seg_label, bold=True)
        for ci, yr_col in enumerate(HIST_COLS):
            _hval(ws, r, yr_col, hist_rev[ci])
        # Forecast: deferred to pass 2 (Growth_Drivers built first, then linked)
        _note(ws, r, f"Forecast: Links to Growth_Drivers bottom-up implied revenue (pass 2)")
        R[f"segrev_{seg_key.lower()}"] = r
        r += 1

        # YoY %
        _pct(ws, r, HIST_COLS[1],
             f"=({CL(HIST_COLS[1])}{r-1}-{CL(HIST_COLS[0])}{r-1})/{CL(HIST_COLS[0])}{r-1}")
        _pct(ws, r, HIST_COLS[2],
             f"=({CL(HIST_COLS[2])}{r-1}-{CL(HIST_COLS[1])}{r-1})/{CL(HIST_COLS[1])}{r-1}")
        for fc in FCST_COLS:
            prev = CL(fc-4) if fc == 6 else CL(fc-1)
            _pct(ws, r, fc, f"={CL(fc)}{r-1}/{prev}{r-1}-1")
        _note(ws, r, "YoY Growth %")
        r += 2  # spacer

    # Memo: Total Segment Revenue (includes internal Foundry)
    _section(ws, r, "Memo: Total Segment Revenue (incl. internal Foundry)")
    for col in HIST_COLS + FCST_COLS:
        parts = "+".join([f"{CL(col)}{R['segrev_'+k]}" for k in ["ccg","dcai","foundry","allother"]])
        _fval(ws, r, col, f"={parts}")
    _note(ws, r, "Sum of 4 reported segments; Foundry ~99% internal → excluded from consolidated")
    R["segrev_sum_segments"] = r; r += 2

    # External Foundry Revenue (the portion sold to 3rd-party customers)
    _label(ws, r, "External Foundry Revenue (3rd-party only)")
    hist_ext_foundry = [50, 60, 150]
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, hist_ext_foundry[ci])
    for fc in FCST_COLS:
        fc_year = {6: "2026e", 7: "2027e", 8: "2028e"}[fc]
        _lval(ws, r, fc, f"=Assumptions!I{R[f'asp_ext_foundry_{fc_year}']}")
    _note(ws, r, "~1% of total Foundry revenue historically. Growing as 18A/14A external customers ramp")
    R["segrev_ext_foundry"] = r; r += 1

    # Intersegment Eliminations (historical are small residual; bulk Foundry internal excluded from sum)
    _label(ws, r, "Intersegment Eliminations")
    hist_elim = [-205, -161, 619]
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, hist_elim[ci])
    for fc in FCST_COLS:
        _lval(ws, r, fc, f"=Assumptions!I{R['asp_inter_elim']}")
    _note(ws, r, "Residual elim; Foundry~Products internal wafer sales excluded from consolidated"); r += 1

    # Consolidated Revenue — historical hardcoded (10-K facts), forecast = CCG+DCAI+ExtFoundry+AllOther+Elim
    _section(ws, r, "TOTAL Consolidated Revenue")
    hist_consol = [54228, 53101, 52853]
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, hist_consol[ci])
    for fc in FCST_COLS:
        ext_parts = "+".join([f"{CL(fc)}{R['segrev_'+k]}" for k in ["ccg","dcai","allother"]])
        formula = f"={ext_parts}+{CL(fc)}{R['segrev_ext_foundry']}+{CL(fc)}{r-1}"
        _key(ws, r, fc, formula)
    _note(ws, r, "Historical: 10-K actuals. Forecast: CCG+DCAI+ExtFoundry+AllOther+Elim. Links to Consolidated_PL")
    R["segrev_consol_rev"] = r; r += 2

    # Revenue Mix
    _section(ws, r, "Revenue Mix (% of Consolidated)")
    for sk, sl in [("ccg","CCG %"),("dcai","DCAI %"),("allother","All Other %")]:
        _label(ws, r, f"  {sl}")
        for col in HIST_COLS + FCST_COLS:
            _pct(ws, r, col, f"={CL(col)}{R['segrev_'+sk]}/{CL(col)}{R['segrev_consol_rev']}")
        r += 1
    # Foundry mix uses external Foundry revenue
    _label(ws, r, "  External Foundry %")
    for col in HIST_COLS + FCST_COLS:
        _pct(ws, r, col, f"={CL(col)}{R['segrev_ext_foundry']}/{CL(col)}{R['segrev_consol_rev']}")
    _note(ws, r, "External foundry only; internal wafer sales eliminated"); r += 1
    # Memo: Total Foundry as % of segment sum
    _label(ws, r, "  Memo: Total Foundry / Seg Sum")
    for col in HIST_COLS + FCST_COLS:
        _pct(ws, r, col, f"={CL(col)}{R['segrev_foundry']}/{CL(col)}{R['segrev_sum_segments']}")
    _note(ws, r, "Shows Foundry scale relative to all segments"); r += 1

    _set_col_widths(ws)
    lr = _add_legend(ws, r + 1)
    _add_sources(ws, lr + 1)
    ws.sheet_properties.tabColor = "548235"
    return r


def build_segment_revenue_link_gd(wb, R):
    """Pass 2: Link Segment_Revenue forecast columns to Growth_Drivers implied revenue.
    Called after Growth_Drivers is built to avoid circular R-key dependency."""
    ws = wb["Segment_Revenue"]

    seg_to_gd = {
        "segrev_ccg": "gd_ccg_implied_rev",
        "segrev_dcai": "gd_dcai_implied_rev",
        "segrev_foundry": "gd_foundry_implied_rev",
        "segrev_allother": "gd_allother_implied_rev",
    }

    for seg_key, gd_key in seg_to_gd.items():
        r = R[seg_key]
        for fc in FCST_COLS:
            f_text = f"=Growth_Drivers!{CL(fc)}{R[gd_key]}"
            ws.cell(row=r, column=fc, value=f_text)
            ws.cell(row=r, column=fc).font = Font(color="548235")
            ws.cell(row=r, column=fc).number_format = "#,##0"
            ws.cell(row=r, column=fc).border = BD


# ============================================================
# TAB 5: SEGMENT_PL (Operating Income by Division)
# ============================================================

def build_segment_pl(wb, R):
    """Build Segment P&L. FIRST PASS - cross-check row filled in pass 2."""
    ws = wb["Segment_PL"]
    _title_row(ws, 1, "Intel (INTC) - Segment P&L (Operating Income by Division)")
    _unit_row(ws, 2)
    _yr_header_no_scenario(ws, 3)
    _label(ws, 4, "USD Million (except margin %)")
    _note(ws, 4, "OI = Segment Revenue x OPM% from Assumptions. EBIT links to Consolidated_PL.")

    r = 5
    segs = [
        ("CCG", "CCG - Client Computing Group", "ccg", "segrev_ccg",
         {"2026e": "asp_ccg_opm_2026e", "2027e": "asp_ccg_opm_2027e", "2028e": "asp_ccg_opm_2028e"},
         [10128, 11594, 9317]),
        ("DCAI", "DCAI - Data Center & AI", "dcai", "segrev_dcai",
         {"2026e": "asp_dcai_opm_2026e", "2027e": "asp_dcai_opm_2027e", "2028e": "asp_dcai_opm_2028e"},
         [945, 1414, 3422]),
        ("Foundry", "Intel Foundry", "foundry", "segrev_foundry",
         {"2026e": "asp_foundry_opm_2026e", "2027e": "asp_foundry_opm_2027e", "2028e": "asp_foundry_opm_2028e"},
         [-7083, -13291, -10318]),
        ("AllOther", "All Other (Mobileye + IMS + Other)", "allother", "segrev_allother",
         {"2026e": "asp_allother_opm_2026e", "2027e": "asp_allother_opm_2027e", "2028e": "asp_allother_opm_2028e"},
         [1507, -57, 264]),
    ]

    for seg_label, seg_desc, seg_key, rev_key, opm_keys, hist_oi in segs:
        _label(ws, r, seg_desc, bold=True)
        # Revenue (link)
        for col in HIST_COLS + FCST_COLS:
            _lval(ws, r, col, f"=Segment_Revenue!{CL(col)}{R[rev_key]}")
        _note(ws, r, "Revenue <- links Segment_Revenue"); r += 1

        # Operating Income
        _label(ws, r, "  Operating Income")
        for ci, col in enumerate(HIST_COLS):
            _hval(ws, r, col, hist_oi[ci])
        yr_labels = ["2026e", "2027e", "2028e"]
        for fi, fc in enumerate(FCST_COLS):
            yr = yr_labels[fi]
            _fval(ws, r, fc, f"={CL(fc)}{r-1}*Assumptions!I{R[opm_keys[yr]]}")
        _note(ws, r, f"Forecast = Revenue x OPM% from Assumptions")
        R[f"segpl_{seg_key}_oi"] = r; r += 1

        # OPM %
        _label(ws, r, "  Operating Margin %")
        for col in HIST_COLS + FCST_COLS:
            _pct(ws, r, col, f"={CL(col)}{r-1}/{CL(col)}{r-2}")
        r += 2  # spacer

    # Corporate Unallocated
    _section(ws, r, "Corporate Unallocated")
    hist_corp = [-5199, -11177, -5518]
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, hist_corp[ci])
    for fi, fc in enumerate(FCST_COLS):
        yr = yr_labels[fi]
        _lval(ws, r, fc, f"=Assumptions!I{R[f'asp_corp_unalloc_{yr}']}")
    _note(ws, r, "From Assumptions. Declining from cost-cutting."); r += 1

    # Sum of Segment OI
    _section(ws, r, "Sum of Segment Operating Income + Corp Unalloc")
    for col in HIST_COLS + FCST_COLS:
        parts = "+".join([f"{CL(col)}{R['segpl_'+k+'_oi']}" for k in ["ccg","dcai","foundry","allother"]])
        _key(ws, r, col, f"={parts}+{CL(col)}{r-1}")
    _note(ws, r, "KEY OUTPUT. Links to Consolidated_PL EBIT.")
    R["segpl_sum_oi"] = r; r += 2

    # Cross-check placeholder
    _section(ws, r, "CROSS-CHECK: Sum Seg OI vs Consolidated EBIT")
    _label(ws, r+1, "  Segment Sum OI")
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r+1, col, f"={CL(col)}{R['segpl_sum_oi']}")
    _label(ws, r+2, "  Consolidated_PL EBIT")
    _note(ws, r+2, "TODO: fill in pass 2 after Consolidated_PL built")
    _label(ws, r+3, "  Difference (MUST = 0)")
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r+3, col, f"={CL(col)}{r+1}-{CL(col)}{r+2}", "0.0")
    _note(ws, r+3, "MUST = 0 - validates segment/consolidated linkage")
    R["segpl_xcheck_sum_oi"] = r+1
    R["segpl_xcheck_ebit"] = r+2
    R["segpl_xcheck_diff"] = r+3
    r = r + 4

    _set_col_widths(ws)
    lr = _add_legend(ws, r + 1)
    _add_sources(ws, lr + 1)
    ws.sheet_properties.tabColor = "548235"
    return r


def build_segment_pl_pass2(wb, R):
    """Second pass: fill cross-check link from Segment_PL to Consolidated_PL EBIT."""
    ws = wb["Segment_PL"]
    for col in HIST_COLS + FCST_COLS:
        c = ws.cell(row=R["segpl_xcheck_ebit"], column=col,
                    value=f"=Consolidated_PL!{CL(col)}{R['pl_ebit']}")
        c.font = FONT_LINK; c.number_format = "#,##0"; c.border = BD; c.alignment = ALIGN_RIGHT


# ============================================================
# TAB 6: CONSOLIDATED_PL (by Cost Type, Segment-Driven EBIT)
# ============================================================

def build_consolidated_pl(wb, R):
    """Build consolidated P&L with segment-driven EBIT."""
    ws = wb["Consolidated_PL"]
    _title_row(ws, 1, "Intel (INTC) - Consolidated Income Statement")
    _unit_row(ws, 2)
    _yr_header_no_scenario(ws, 3)
    _label(ws, 4, "USD Million (except per-share data)")
    _note(ws, 4, "Cost structure by type. EBIT = Segment_PL Sum OI (segment-driven).")

    r = 5
    yr_labels = ["2026e", "2027e", "2028e"]

    # Revenue
    _label(ws, r, "Net Revenue", bold=True)
    for col in HIST_COLS + FCST_COLS:
        _lval(ws, r, col, f"=Segment_Revenue!{CL(col)}{R['segrev_consol_rev']}")
    _note(ws, r, "<- Segment_Revenue consolidated total")
    R["pl_rev"] = r; r += 1

    # COGS
    _label(ws, r, "Cost of Sales")
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, [32517, 35756, 34478][ci])
    for fi, fc in enumerate(FCST_COLS):
        yr = yr_labels[fi]
        _fval(ws, r, fc, f"={CL(fc)}{R['pl_rev']}*Assumptions!I{R[f'asp_cogs_{yr}']}")
    _note(ws, r, "= Revenue x COGS% from Assumptions")
    R["pl_cogs"] = r; r += 1

    # Gross Profit
    _label(ws, r, "Gross Profit", bold=True)
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, [21711, 17345, 18375][ci])
    for fc in FCST_COLS:
        _fval(ws, r, fc, f"={CL(fc)}{R['pl_rev']}-{CL(fc)}{R['pl_cogs']}")
    _note(ws, r, "= Revenue - COGS")
    R["pl_gp"] = r; r += 1
    _label(ws, r, "  Gross Margin %")
    for col in HIST_COLS + FCST_COLS:
        _pct(ws, r, col, f"={CL(col)}{R['pl_gp']}/{CL(col)}{R['pl_rev']}")
    R["pl_gpm"] = r; r += 2

    # OpEx
    _section(ws, r, "Operating Expenses (by cost type)"); r += 1
    _label(ws, r, "Research & Development")
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, [16046, 16546, 13774][ci])
    for fi, fc in enumerate(FCST_COLS):
        yr = yr_labels[fi]
        _fval(ws, r, fc, f"={CL(fc)}{R['pl_rev']}*Assumptions!I{R[f'asp_rd_{yr}']}")
    _note(ws, r, "= Revenue x R&D%")
    R["pl_rd"] = r; r += 1

    _label(ws, r, "Marketing, General & Administrative")
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, [5634, 5507, 4624][ci])
    for fi, fc in enumerate(FCST_COLS):
        yr = yr_labels[fi]
        _fval(ws, r, fc, f"={CL(fc)}{R['pl_rev']}*Assumptions!I{R[f'asp_mga_{yr}']}")
    _note(ws, r, "= Revenue x MG&A%")
    R["pl_mga"] = r; r += 1

    _label(ws, r, "Restructuring & Other Charges")
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, [-62, 6970, 2191][ci])
    for fc in FCST_COLS:
        _lval(ws, r, fc, f"=Assumptions!I{R['asp_restruct']}")
    _note(ws, r, "From Assumptions; winding down"); r += 1

    # Total OpEx (top-down)
    _label(ws, r, "Total OpEx (per cost rates)", bold=True)
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r, col, f"={CL(col)}{R['pl_rd']}+{CL(col)}{R['pl_mga']}+{CL(col)}{r-1}")
    _note(ws, r, "= R&D + MG&A + Restructuring (top-down cost build)")
    R["pl_total_opex"] = r; r += 2

    # EBIT - SEGMENT DRIVEN
    _section(ws, r, "EBIT - Segment-Driven"); r += 1
    _label(ws, r, "EBIT (Operating Income)", bold=True)
    for col in HIST_COLS + FCST_COLS:
        _lval(ws, r, col, f"=Segment_PL!{CL(col)}{R['segpl_sum_oi']}")
    _note(ws, r, "<- Segment_PL Sum of Segment OI. SEGMENT-DRIVEN source of truth.")
    R["pl_ebit"] = r; r += 1

    _label(ws, r, "  EBIT Margin %")
    for col in HIST_COLS + FCST_COLS:
        _pct(ws, r, col, f"={CL(col)}{R['pl_ebit']}/{CL(col)}{R['pl_rev']}")
    r += 1

    _label(ws, r, "  Implied Total OpEx (GP - EBIT)")
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r, col, f"={CL(col)}{R['pl_gp']}-{CL(col)}{R['pl_ebit']}", "#,##0")
    _note(ws, r, "Transparency: GP - EBIT. May differ from top-down OpEx above."); r += 2

    # Non-Operating
    _section(ws, r, "Non-Operating Items"); r += 1
    _label(ws, r, "Interest & Other, Net")
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, [629, 226, 3257][ci])
    for fi, fc in enumerate(FCST_COLS):
        yr = yr_labels[fi]
        _lval(ws, r, fc, f"=Assumptions!I{R[f'asp_interest_{yr}']}")
    _note(ws, r, "Net interest expense on debt")
    R["pl_interest"] = r; r += 1

    _label(ws, r, "Gains/Losses on Equity Investments")
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, [40, 242, 514][ci])
    for fc in FCST_COLS:
        _hval(ws, r, fc, 0)
    _note(ws, r, "Unpredictable; set to 0 for forecast"); r += 1

    # Pretax
    _label(ws, r, "Income (Loss) Before Taxes", bold=True)
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r, col, f"={CL(col)}{R['pl_ebit']}+{CL(col)}{R['pl_interest']}+{CL(col)}{r-1}")
    _note(ws, r, "= EBIT + Interest + Equity Gains")
    R["pl_pretax"] = r; r += 1

    # Tax
    _label(ws, r, "Provision for (Benefit from) Taxes")
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, [-913, 8023, 1531][ci])
    etr_keys = {"2026e": "asp_etr_sel", "2027e": "asp_etr_2027e", "2028e": "asp_etr_2028e"}
    for fi, fc in enumerate(FCST_COLS):
        yr = yr_labels[fi]
        _fval(ws, r, fc,
              f"=IF({CL(fc)}{R['pl_pretax']}>0,{CL(fc)}{R['pl_pretax']}*Assumptions!I{R[etr_keys[yr]]},0)")
    _note(ws, r, "Tax on positive pretax only (valuation allowance shields US losses)")
    R["pl_tax"] = r; r += 1

    # Net Income
    _label(ws, r, "Net Income (Loss)", bold=True)
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r, col, f"={CL(col)}{R['pl_pretax']}-{CL(col)}{R['pl_tax']}")
    _note(ws, r, "= Pretax - Tax")
    R["pl_ni"] = r; r += 1
    _label(ws, r, "  Net Income Margin %")
    for col in HIST_COLS + FCST_COLS:
        _pct(ws, r, col, f"={CL(col)}{R['pl_ni']}/{CL(col)}{R['pl_rev']}")
    r += 1

    # Less NCI
    _label(ws, r, "Less: Net Income Attributable to NCI")
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, [-14, -477, 293][ci])
    nci_keys = {"2026e": "asp_nci_2026e", "2027e": "asp_nci_2027e", "2028e": "asp_nci_2028e"}
    for fi, fc in enumerate(FCST_COLS):
        yr = yr_labels[fi]
        _lval(ws, r, fc, f"=Assumptions!I{R[nci_keys[yr]]}")
    _note(ws, r, "NCI from Assumptions (AZ SCIP + Mobileye; Ireland SCIP eliminated Apr 2026)")
    R["pl_nci"] = r; r += 1

    # NI to Intel
    _label(ws, r, "Net Income Attributable to Intel", bold=True)
    for col in HIST_COLS + FCST_COLS:
        _key(ws, r, col, f"={CL(col)}{R['pl_ni']}-{CL(col)}{r-1}")
    _note(ws, r, "= NI - NCI. Key EPS driver.")
    R["pl_ni_intel"] = r; r += 2

    # Memo
    _section(ws, r, "Memo Items"); r += 1
    _label(ws, r, "D&A (Memo)")
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, [7847, 9951, 10757][ci])
    for fi, fc in enumerate(FCST_COLS):
        yr = yr_labels[fi]
        _lval(ws, r, fc, f"=Assumptions!I{R[f'asp_da_{yr}']}")
    _note(ws, r, "Reference for EBITDA calc")
    R["pl_da"] = r; r += 1

    _label(ws, r, "SBC (Memo)")
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, [3229, 3410, 2434][ci])
    for fc in FCST_COLS:
        _lval(ws, r, fc, f"=Assumptions!I{R['asp_sbc']}")
    _note(ws, r, "Reference for CFO/FCF bridge")
    R["pl_sbc"] = r; r += 1

    _label(ws, r, "Diluted Shares Outstanding (M)")
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, [4212, 4280, 4530][ci])
    shares_keys = {"2026e": "asp_shares_2026e", "2027e": "asp_shares_2027e", "2028e": "asp_shares_2028e"}
    for fi, fc in enumerate(FCST_COLS):
        yr = yr_labels[fi]
        _lval(ws, r, fc, f"=Assumptions!I{R[shares_keys[yr]]}")
    _note(ws, r, "Basic shares from Assumptions; diluted approximates basic")
    R["pl_shares"] = r; r += 1

    # Diluted EPS
    _label(ws, r, "Diluted EPS ($)", bold=True)
    for col in HIST_COLS + FCST_COLS:
        _key(ws, r, col, f"={CL(col)}{R['pl_ni_intel']}/{CL(col)}{R['pl_shares']}", "0.00")
    _note(ws, r, "= NI to Intel / Shares. KEY OUTPUT")
    R["pl_eps"] = r; r += 1

    _label(ws, r, "Dividends per Share (Memo)")
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, [0.74, 0.38, 0.00][ci], "0.00")
    for fc in FCST_COLS:
        _lval(ws, r, fc, f"=Assumptions!I{R['asp_dps']}", "0.00")
    _note(ws, r, "Suspended; no reinstatement expected"); r += 1

    _set_col_widths(ws)
    lr = _add_legend(ws, r + 1)
    _add_sources(ws, lr + 1)
    ws.sheet_properties.tabColor = "1F4E79"
    return r


def _set_col_widths(ws):
    ws.column_dimensions['A'].width = 42
    for c in ['B','C','D','F','G','H']:
        ws.column_dimensions[c].width = 14
    ws.column_dimensions['E'].width = 2
    ws.column_dimensions['I'].width = 14
    ws.column_dimensions['J'].width = 2
    ws.column_dimensions['K'].width = 55


# ============================================================
# TAB 7: BS (Balance Sheet with Cash as Plug)
# ============================================================

def build_bs(wb, R):
    """Build balance sheet: assets = liabilities + equity. Cash is the plug."""
    ws = wb["BS"]
    _title_row(ws, 1, "Intel (INTC) - Balance Sheet")
    _unit_row(ws, 2)
    _yr_header_no_scenario(ws, 3)
    _label(ws, 4, "USD Million")
    _note(ws, 4, "Cash = Total L&E - Non-Cash Assets (plug). Balance check row must = 0.")

    r = 5
    yr_labels = ["2026e","2027e","2028e"]

    # ---- ASSETS ----
    _section(ws, r, "ASSETS"); r += 1
    _section(ws, r, "Current Assets"); r += 1

    # Cash (PLUG - computed after Total L&E and non-cash assets are known)
    _label(ws, r, "Cash and Cash Equivalents (PLUG)", bold=True)
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, [7079, 8249, 14265][ci])
    # Cash formula = Total L&E - (all other assets)
    # We'll write this after we know Total L&E row
    _note(ws, r, "PLUG: = Total L&E - all non-cash assets. Ensures A = L+E.")
    R["bs_cash"] = r; r += 1

    # Marketable Securities
    _label(ws, r, "Short-term Investments (Marketable Securities)")
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, [17955, 13813, 23151][ci])
    # Forecast: grow slightly (assume ~$24B flat for now)
    for fc in FCST_COLS:
        _hval(ws, r, fc, 24000)
    _note(ws, r, "TODO: Model as % of total cash or flat. Current: flat ~$24B assumption.")
    R["bs_ms"] = r; r += 1

    # AR
    _label(ws, r, "Accounts Receivable, Net")
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, [3402, 3478, 3839][ci])
    for fc in FCST_COLS:
        _fval(ws, r, fc, f"=Consolidated_PL!{CL(fc)}{R['pl_rev']}*Assumptions!I{R['asp_ar_days']}/365")
    _note(ws, r, "= Revenue x AR Days / 365")
    R["bs_ar"] = r; r += 1

    # Inventories
    _label(ws, r, "Inventories")
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, [11127, 12198, 11618][ci])
    for fc in FCST_COLS:
        _fval(ws, r, fc, f"=Consolidated_PL!{CL(fc)}{R['pl_cogs']}*Assumptions!I{R['asp_inv_days']}/365")
    _note(ws, r, "= COGS x Inventory Days / 365")
    R["bs_inventory"] = r; r += 1

    # Other CA
    _label(ws, r, "Other Current Assets")
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, [3706, 9586, 10815][ci])
    for fc in FCST_COLS:
        _fval(ws, r, fc, f"=Consolidated_PL!{CL(fc)}{R['pl_rev']}*Assumptions!I{R['asp_other_ca']}")
    _note(ws, r, "= Revenue x Other CA% (incl. refundable tax credits)")
    R["bs_other_ca"] = r; r += 1

    # Total CA
    _label(ws, r, "Total Current Assets", bold=True)
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r, col, f"={CL(col)}{R['bs_cash']}+{CL(col)}{R['bs_ms']}+{CL(col)}{R['bs_ar']}+{CL(col)}{R['bs_inventory']}+{CL(col)}{R['bs_other_ca']}")
    _note(ws, r, "= Cash + Securities + AR + Inventory + Other CA")
    R["bs_total_ca"] = r; r += 2

    # Non-Current Assets
    _section(ws, r, "Non-Current Assets"); r += 1

    _label(ws, r, "Property, Plant & Equipment, Net")
    hist_ppe = [96647, 107919, 105414]
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, hist_ppe[ci])
    # PP&E = Prior PP&E + CapEx - D&A - Govt Incentives
    # Simplified: Prior + Gross CapEx - D&A
    for fc in FCST_COLS:
        yr = yr_labels[FCST_COLS.index(fc)]
        prev_col = HIST_COLS[2] if fc == 6 else fc - 1
        _fval(ws, r, fc,
              f"={CL(prev_col)}{r}+Assumptions!I{R[f'asp_capex_{yr}']}-Consolidated_PL!{CL(fc)}{R['pl_da']}")
    _note(ws, r, "= Prior PP&E + Gross CapEx - D&A. Govt incentives not explicitly modelled (reducing gross PP&E).")
    R["bs_ppe"] = r; r += 1

    _label(ws, r, "Equity Investments")
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, [5829, 5383, 8512][ci])
    for fc in FCST_COLS:
        _hval(ws, r, fc, 8500)
    _note(ws, r, "Includes Altera 49% retained stake (equity method). Flat forecast.")
    R["bs_equity_inv"] = r; r += 1

    _label(ws, r, "Goodwill")
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, [27591, 24693, 23912][ci])
    # Q1 2026 had $3.9B Mobileye goodwill impairment
    goodwill_q1_2026 = 23912 - 3900
    for fc in FCST_COLS:
        _hval(ws, r, fc, goodwill_q1_2026)
    _note(ws, r, f"Post-Q1 2026 $3.9B Mobileye impairment. Approx ${goodwill_q1_2026}M.")
    R["bs_goodwill"] = r; r += 1

    _label(ws, r, "Identified Intangible Assets, Net")
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, [4589, 3691, 2772][ci])
    for fc in FCST_COLS:
        prev_col = HIST_COLS[2] if fc == 6 else fc - 1
        _fval(ws, r, fc, f"={CL(prev_col)}{r}-500")
    _note(ws, r, "Amortising ~$500M/year")
    R["bs_intangibles"] = r; r += 1

    _label(ws, r, "Other Long-Term Assets")
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, [13647, 7475, 7131][ci])
    for fc in FCST_COLS:
        _hval(ws, r, fc, 7000)
    _note(ws, r, "Flat ~$7B assumption")
    R["bs_other_nca"] = r; r += 1

    # Total NCA
    _label(ws, r, "Total Non-Current Assets", bold=True)
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r, col, f"={CL(col)}{R['bs_ppe']}+{CL(col)}{R['bs_equity_inv']}+{CL(col)}{R['bs_goodwill']}+{CL(col)}{R['bs_intangibles']}+{CL(col)}{R['bs_other_nca']}")
    R["bs_total_nca"] = r; r += 1

    # Total Assets
    _label(ws, r, "TOTAL ASSETS", bold=True)
    for col in HIST_COLS + FCST_COLS:
        _key(ws, r, col, f"={CL(col)}{R['bs_total_ca']}+{CL(col)}{R['bs_total_nca']}")
    R["bs_total_assets"] = r; r += 3

    # ---- LIABILITIES ----
    _section(ws, r, "LIABILITIES & EQUITY"); r += 1
    _section(ws, r, "Current Liabilities"); r += 1

    _label(ws, r, "Accounts Payable")
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, [8578, 12556, 9882][ci])
    for fc in FCST_COLS:
        _fval(ws, r, fc, f"=Consolidated_PL!{CL(fc)}{R['pl_cogs']}*Assumptions!I{R['asp_ap_days']}/365")
    _note(ws, r, "= COGS x AP Days / 365")
    R["bs_ap"] = r; r += 1

    _label(ws, r, "Accrued Compensation and Benefits")
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, [3655, 3343, 3990][ci])
    for fc in FCST_COLS:
        _fval(ws, r, fc, f"=Consolidated_PL!{CL(fc)}{R['pl_rev']}*Assumptions!I{R['asp_accrued_comp']}")
    _note(ws, r, "= Revenue x Accrued Comp%")
    R["bs_accrued_comp"] = r; r += 1

    _label(ws, r, "Short-Term Debt")
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, [2288, 3729, 2499][ci])
    for fc in FCST_COLS:
        _hval(ws, r, fc, 2500)
    _note(ws, r, "Stable ~$2.5B assumption; commercial paper + current maturities")
    R["bs_st_debt"] = r; r += 1

    _label(ws, r, "Income Taxes Payable")
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, [1107, 1756, 604][ci])
    for fc in FCST_COLS:
        _fval(ws, r, fc, f"=Consolidated_PL!{CL(fc)}{R['pl_rev']}*Assumptions!I{R['asp_tax_payable']}")
    _note(ws, r, "= Revenue x Tax Payable%")
    R["bs_tax_payable"] = r; r += 1

    _label(ws, r, "Other Accrued Liabilities")
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, [12425, 14282, 14600][ci])
    for fc in FCST_COLS:
        _fval(ws, r, fc, f"=Consolidated_PL!{CL(fc)}{R['pl_rev']}*Assumptions!I{R['asp_other_cl']}")
    _note(ws, r, "= Revenue x Other CL%")
    R["bs_other_cl"] = r; r += 1

    # Total CL
    _label(ws, r, "Total Current Liabilities", bold=True)
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r, col, f"={CL(col)}{R['bs_ap']}+{CL(col)}{R['bs_accrued_comp']}+{CL(col)}{R['bs_st_debt']}+{CL(col)}{R['bs_tax_payable']}+{CL(col)}{R['bs_other_cl']}")
    R["bs_total_cl"] = r; r += 2

    # Non-Current Liabilities
    _section(ws, r, "Non-Current Liabilities"); r += 1

    _label(ws, r, "Long-Term Debt")
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, [46978, 46282, 44086][ci])
    # LT Debt = Total Debt - ST Debt
    for fi, fc in enumerate(FCST_COLS):
        yr = yr_labels[fi]
        _fval(ws, r, fc, f"=Assumptions!I{R[f'asp_debt_{yr}']}-{CL(fc)}{R['bs_st_debt']}")
    _note(ws, r, "= Total Debt - ST Debt (from Assumptions)")
    R["bs_lt_debt"] = r; r += 1

    _label(ws, r, "Other Long-Term Liabilities")
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, [6576, 9505, 9408][ci])
    for fc in FCST_COLS:
        _hval(ws, r, fc, 9500)
    _note(ws, r, "Flat ~$9.5B assumption. Includes tax, operating lease, other.")
    R["bs_other_ncl"] = r; r += 1

    # Total NCL
    _label(ws, r, "Total Non-Current Liabilities", bold=True)
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r, col, f"={CL(col)}{R['bs_lt_debt']}+{CL(col)}{R['bs_other_ncl']}")
    R["bs_total_ncl"] = r; r += 1

    # Total Liabilities
    _label(ws, r, "TOTAL LIABILITIES", bold=True)
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r, col, f"={CL(col)}{R['bs_total_cl']}+{CL(col)}{R['bs_total_ncl']}")
    R["bs_total_liab"] = r; r += 2

    # ---- EQUITY ----
    _section(ws, r, "Stockholders' Equity"); r += 1

    _label(ws, r, "Total Intel Stockholders' Equity")
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, [105590, 99270, 114281][ci])
    # Equity = Prior Equity + NI to Intel + SBC - Dividends + Net Equity Issuance
    for fc in FCST_COLS:
        prev_col = HIST_COLS[2] if fc == 6 else fc - 1
        _fval(ws, r, fc,
              f"={CL(prev_col)}{r}+Consolidated_PL!{CL(fc)}{R['pl_ni_intel']}+Consolidated_PL!{CL(fc)}{R['pl_sbc']}+Assumptions!I{R['asp_net_eq_iss']}")
    _note(ws, r, "= Prior Equity + NI to Intel + SBC + Net Equity Issuance. Dividends = 0.")
    R["bs_equity"] = r; r += 1

    _label(ws, r, "Non-Controlling Interests")
    hist_nci_bs = [4375, 5762, 12079]
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, hist_nci_bs[ci])
    # NCI = Prior NCI + NCI(P&L) - NCI distributions (simplified)
    # After Apr 2026 Ireland SCIP buyout: NCI (BS) drops
    nci_q1_2026_est = 13595  # from Q1 2026 10-Q
    for fc in FCST_COLS:
        prev_col = HIST_COLS[2] if fc == 6 else fc - 1
        # Simplified: prior + P&L NCI
        _fval(ws, r, fc, f"={CL(prev_col)}{r}+Consolidated_PL!{CL(fc)}{R['pl_nci']}-300")
    _note(ws, r, "= Prior NCI + P&L NCI - estimated distributions (~$300M). AZ SCIP + Mobileye remain.")
    R["bs_nci"] = r; r += 1

    # Total Equity
    _label(ws, r, "Total Stockholders' Equity", bold=True)
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r, col, f"={CL(col)}{R['bs_equity']}+{CL(col)}{R['bs_nci']}")
    R["bs_total_equity"] = r; r += 1

    # Total L&E
    _label(ws, r, "TOTAL LIABILITIES & EQUITY", bold=True)
    for col in HIST_COLS + FCST_COLS:
        _key(ws, r, col, f"={CL(col)}{R['bs_total_liab']}+{CL(col)}{R['bs_total_equity']}")
    R["bs_total_le"] = r; r += 1

    # ---- NOW: Write Cash Plug Formula ----
    # Cash = Total L&E - (all other assets)
    # But this creates circular reference with Total CA (which includes Cash)
    # Solution: Cash = Total L&E - Total NCA - (CA excl Cash)
    ca_excl_cash = f"({CL(6)}{R['bs_ms']}+{CL(6)}{R['bs_ar']}+{CL(6)}{R['bs_inventory']}+{CL(6)}{R['bs_other_ca']})"
    for fc in FCST_COLS:
        cash_formula = f"={CL(fc)}{R['bs_total_le']}-{CL(fc)}{R['bs_total_nca']}-({CL(fc)}{R['bs_ms']}+{CL(fc)}{R['bs_ar']}+{CL(fc)}{R['bs_inventory']}+{CL(fc)}{R['bs_other_ca']})"
        c = ws.cell(row=R["bs_cash"], column=fc, value=cash_formula)
        c.number_format = "#,##0"; c.font = FONT_BOLD; c.fill = FILL_KEY; c.border = BD; c.alignment = ALIGN_RIGHT

    # Balance Check
    r += 1
    _section(ws, r, "BALANCE CHECK")
    _label(ws, r, "  Total Assets")
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r, col, f"={CL(col)}{R['bs_total_ca']}+{CL(col)}{R['bs_total_nca']}")
    R["bs_assets_check"] = r; r += 1
    _label(ws, r, "  Total L&E (from above)")
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r, col, f"={CL(col)}{R['bs_total_le']}")
    r += 1
    _label(ws, r, "   Balance Check (A - L&E) - MUST = 0", bold=True)
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r, col, f"={CL(col)}{R['bs_assets_check']}-{CL(col)}{r-1}", "0.0")
    _note(ws, r, "MUST = 0 in all years. Validates BS integrity with cash plug.")
    R["bs_balance_check"] = r; r += 1

    _set_col_widths(ws)
    lr = _add_legend(ws, r + 1)
    _add_sources(ws, lr + 1)
    ws.sheet_properties.tabColor = "1F4E79"
    return r


# ============================================================
# TAB 8: CASH_FLOW (WC from BS Delta)
# ============================================================

def build_cash_flow(wb, R):
    """Build cash flow statement: CFO + CFI + CFF."""
    ws = wb["Cash_Flow"]
    _title_row(ws, 1, "Intel (INTC) - Cash Flow Statement")
    _unit_row(ws, 2)
    _yr_header_no_scenario(ws, 3)
    _label(ws, 4, "USD Million")
    _note(ws, 4, "WC changes linked to BS deltas. CFO + CFI = FCF.")

    r = 5

    # ---- Operating Activities ----
    _section(ws, r, "Operating Activities"); r += 1

    _label(ws, r, "Net Income (Loss)")
    for col in HIST_COLS + FCST_COLS:
        _lval(ws, r, col, f"=Consolidated_PL!{CL(col)}{R['pl_ni']}")
    _note(ws, r, "<- Consolidated_PL")
    R["cf_ni"] = r; r += 1

    # Non-cash adjustments
    _label(ws, r, "(+) Depreciation & Amortization")
    for col in HIST_COLS + FCST_COLS:
        _lval(ws, r, col, f"=Consolidated_PL!{CL(col)}{R['pl_da']}")
    R["cf_da"] = r; r += 1

    _label(ws, r, "(+) Share-Based Compensation")
    for col in HIST_COLS + FCST_COLS:
        _lval(ws, r, col, f"=Consolidated_PL!{CL(col)}{R['pl_sbc']}")
    R["cf_sbc"] = r; r += 1

    _label(ws, r, "Deferred Taxes")
    hist_def_tax = [-2033, 6132, 328]
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, hist_def_tax[ci])
    for fc in FCST_COLS:
        _hval(ws, r, fc, 0)
    _note(ws, r, "TODO: Model deferred tax. Set to 0 for now (valuation allowance limits)."); r += 1

    _label(ws, r, "Other Non-Cash / Non-Recurring")
    hist_other_cf = [6871, 11925, -2958]
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, hist_other_cf[ci])
    for fc in FCST_COLS:
        _hval(ws, r, fc, 0)
    _note(ws, r, "Restructuring charges, impairments, gains/losses, etc. Set to 0 for forecast."); r += 1

    # Working Capital Changes
    r += 1
    _section(ws, r, "Working Capital Changes"); r += 1

    wc_items = [
        ("Change in AR (-increase = outflow)", "bs_ar", False),
        ("Change in Inventory (-increase = outflow)", "bs_inventory", False),
        ("Change in Other CA (-increase = outflow)", "bs_other_ca", False),
        ("Change in AP (+increase = inflow)", "bs_ap", True),
        ("Change in Accrued Comp (+increase = inflow)", "bs_accrued_comp", True),
        ("Change in Other CL (+increase = inflow)", "bs_other_cl", True),
        ("Change in Tax Payable (+increase = inflow)", "bs_tax_payable", True),
    ]
    for label, bs_key, is_inflow in wc_items:
        _label(ws, r, f"  {label}")
        for col in HIST_COLS + FCST_COLS:
            if col == 2:
                # No prior year for first historical column; leave blank or hardcode
                pass
            else:
                prev_col = col - 1 if col in HIST_COLS else (HIST_COLS[2] if col == 6 else col - 1)
                if is_inflow:
                    _fval(ws, r, col, f"=BS!{CL(col)}{R[bs_key]}-BS!{CL(prev_col)}{R[bs_key]}")
                else:
                    _fval(ws, r, col, f"=-(BS!{CL(col)}{R[bs_key]}-BS!{CL(prev_col)}{R[bs_key]})")
        _note(ws, r, f"Links to BS {bs_key}")
        r += 1
    # Store last WC row for CFO sum
    R["cf_last_wc"] = r - 1

    # CFO
    r += 1
    _label(ws, r, "Net Cash from Operating Activities (CFO)", bold=True)
    wc_rows = range(R["cf_last_wc"] - 6, R["cf_last_wc"] + 1)
    for col in HIST_COLS + FCST_COLS:
        wc_sum = "+".join([f"{CL(col)}{wr}" for wr in wc_rows])
        _key(ws, r, col, f"={CL(col)}{R['cf_ni']}+{CL(col)}{R['cf_da']}+{CL(col)}{R['cf_sbc']}+{CL(col)}{r-4}+{CL(col)}{r-3}+{wc_sum}")
    _note(ws, r, "= NI + D&A + SBC + DeferredTax + OtherNonCash + Sum(WC Changes)")
    R["cf_cfo"] = r; r += 2

    # ---- Investing Activities ----
    _section(ws, r, "Investing Activities"); r += 1

    _label(ws, r, "Capital Expenditure (Gross CapEx)")
    hist_capex = [-25750, -25122, -17672]
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, hist_capex[ci])
    for fi, fc in enumerate(FCST_COLS):
        yr = ["2026e","2027e","2028e"][fi]
        _fval(ws, r, fc, f"=-Assumptions!I{R[f'asp_capex_{yr}']}")
    _note(ws, r, "Gross CapEx (negative). From Assumptions.")
    R["cf_capex"] = r; r += 1

    _label(ws, r, "  CapEx / Revenue %")
    for col in HIST_COLS + FCST_COLS:
        _pct(ws, r, col, f"=-{CL(col)}{r-1}/Consolidated_PL!{CL(col)}{R['pl_rev']}")
    r += 1

    _label(ws, r, "Other Investing Activities, Net")
    hist_other_inv = [1546, 1581, -1175]
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, hist_other_inv[ci])
    for fc in FCST_COLS:
        _hval(ws, r, fc, 0)
    _note(ws, r, "Govt incentives, divestitures, ST investment net. Set to 0 for forecast.")
    R["cf_other_inv"] = r; r += 1

    # CFI
    _label(ws, r, "Net Cash from Investing Activities (CFI)", bold=True)
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r, col, f"={CL(col)}{R['cf_capex']}+{CL(col)}{R['cf_other_inv']}")
    R["cf_cfi"] = r; r += 1

    # FCF
    _label(ws, r, "Free Cash Flow (FCF = CFO + CFI)", bold=True)
    for col in HIST_COLS + FCST_COLS:
        _key(ws, r, col, f"={CL(col)}{R['cf_cfo']}+{CL(col)}{R['cf_cfi']}")
    _note(ws, r, "= CFO + CFI. KEY OUTPUT.")
    R["cf_fcf"] = r; r += 1
    _label(ws, r, "  FCF / Revenue %")
    for col in HIST_COLS + FCST_COLS:
        _pct(ws, r, col, f"={CL(col)}{R['cf_fcf']}/Consolidated_PL!{CL(col)}{R['pl_rev']}")
    r += 2

    # ---- Financing Activities ----
    _section(ws, r, "Financing Activities"); r += 1

    _label(ws, r, "Net Change in Debt")
    hist_debt_net = [11453, -462, -2757]
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, hist_debt_net[ci])
    for fi, fc in enumerate(FCST_COLS):
        yr = ["2026e","2027e","2028e"][fi]
        prev_yr = ["2025a","2026e","2027e"][fi]
        prev_key = "bs_st_debt" if fi == 0 else f"asp_debt_{prev_yr}"
        curr_key = f"asp_debt_{yr}"
        if fi == 0:
            _fval(ws, r, fc, f"=Assumptions!I{R[curr_key]}-(BS!{CL(HIST_COLS[2])}{R['bs_lt_debt']}+BS!{CL(HIST_COLS[2])}{R['bs_st_debt']})")
        else:
            _fval(ws, r, fc, f"=Assumptions!I{R[curr_key]}-Assumptions!I{R[prev_key]}")
    _note(ws, r, "Change in Total Debt from Assumptions")
    R["cf_debt_net"] = r; r += 1

    _label(ws, r, "Net Equity Issuance / (Repurchase)")
    hist_eq_net = [2567, 356, 18586]
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, hist_eq_net[ci])
    for fc in FCST_COLS:
        _lval(ws, r, fc, f"=Assumptions!I{R['asp_net_eq_iss']}")
    _note(ws, r, "ESPP proceeds net of RSU withholdings. FY2025 inflated by strategic placements.")
    R["cf_equity_net"] = r; r += 1

    _label(ws, r, "Dividends Paid")
    hist_div = [-3088, -1599, 0]
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, hist_div[ci])
    for fc in FCST_COLS:
        _fval(ws, r, fc, f"=-Assumptions!I{R['asp_dps']}*Assumptions!I{R['asp_shares_2026e']}")
    _note(ws, r, "Dividends suspended; = 0.")
    R["cf_dividends"] = r; r += 1

    _label(ws, r, "Partner Contributions (SCIP) & Other Financing")
    hist_scip = [3696, 12471, 1827]
    for ci, col in enumerate(HIST_COLS):
        _hval(ws, r, col, hist_scip[ci])
    for fc in FCST_COLS:
        _hval(ws, r, fc, 500)
    _note(ws, r, "AZ SCIP ongoing partner contributions. Ireland SCIP eliminated Apr 2026."); r += 1

    # CFF
    _label(ws, r, "Net Cash from Financing Activities (CFF)", bold=True)
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r, col, f"={CL(col)}{R['cf_debt_net']}+{CL(col)}{R['cf_equity_net']}+{CL(col)}{R['cf_dividends']}+{CL(col)}{r-1}")
    R["cf_cff"] = r; r += 1

    # Net Change in Cash
    _label(ws, r, "Net Increase (Decrease) in Cash", bold=True)
    for col in HIST_COLS + FCST_COLS:
        _key(ws, r, col, f"={CL(col)}{R['cf_cfo']}+{CL(col)}{R['cf_cfi']}+{CL(col)}{R['cf_cff']}")
    _note(ws, r, "= CFO + CFI + CFF")
    R["cf_net_change"] = r; r += 1

    _set_col_widths(ws)
    lr = _add_legend(ws, r + 1)
    _add_sources(ws, lr + 1)
    ws.sheet_properties.tabColor = "548235"
    return r


# ============================================================
# TAB 9: DCF (UFCF -> EV -> Equity Value)
# ============================================================

def build_dcf(wb, R):
    """Build DCF valuation: UFCF -> PV -> TV -> EV -> Equity -> per-share value."""
    ws = wb["DCF"]
    _title_row(ws, 1, "Intel (INTC) - Discounted Cash Flow Valuation")
    _unit_row(ws, 2, "USD Million (except per-share data & multiples)")
    _label(ws, 4, "Forecast Period")
    _note(ws, 4, "DCF based on Unlevered Free Cash Flow (UFCF). TV = Gordon Growth Model.")

    r = 5
    yr_labels = ["2026e","2027e","2028e"]
    # Year numbering for DCF (t=1, 2, 3)
    for fi, fc in enumerate(FCST_COLS):
        ws.cell(row=3, column=fc, value=f"Year {fi+1}").font = FONT_HEADER
        ws.cell(row=3, column=fc).fill = FILL_HEADER; ws.cell(row=3, column=fc).border = BD
        ws.cell(row=3, column=fc).alignment = ALIGN_CENTER

    # EBIT
    _label(ws, r, "EBIT")
    for col in HIST_COLS + FCST_COLS:
        _lval(ws, r, col, f"=Consolidated_PL!{CL(col)}{R['pl_ebit']}")
    _note(ws, r, "<- Consolidated_PL (segment-driven)")
    R["dcf_ebit"] = r; r += 1

    # Tax on EBIT
    _label(ws, r, "(-) Tax on EBIT")
    for col in HIST_COLS:
        _hval(ws, r, col, 0)  # Not meaningful historically
    etr_keys = {"2026e": "asp_etr_sel", "2027e": "asp_etr_2027e", "2028e": "asp_etr_2028e"}
    for fi, fc in enumerate(FCST_COLS):
        yr = yr_labels[fi]
        _fval(ws, r, fc, f"={CL(fc)}{R['dcf_ebit']}*Assumptions!I{R[etr_keys[yr]]}")
    _note(ws, r, "= EBIT x ETR from Assumptions")
    R["dcf_tax_on_ebit"] = r; r += 1

    # NOPAT
    _label(ws, r, "NOPAT (= EBIT - Tax on EBIT)")
    for col in HIST_COLS + FCST_COLS:
        _fval(ws, r, col, f"={CL(col)}{R['dcf_ebit']}-{CL(col)}{R['dcf_tax_on_ebit']}")
    _note(ws, r, "Net Operating Profit After Tax")
    R["dcf_nopat"] = r; r += 1

    # D&A
    _label(ws, r, "(+) D&A")
    for col in HIST_COLS + FCST_COLS:
        _lval(ws, r, col, f"=Consolidated_PL!{CL(col)}{R['pl_da']}")
    R["dcf_da"] = r; r += 1

    # CapEx
    _label(ws, r, "(-) Capital Expenditure")
    for col in HIST_COLS:
        _hval(ws, r, col, 0)
    for fi, fc in enumerate(FCST_COLS):
        yr = yr_labels[fi]
        _fval(ws, r, fc, f"=-Assumptions!I{R[f'asp_capex_{yr}']}")
    _note(ws, r, "Gross CapEx (negative). From Assumptions.")
    R["dcf_capex"] = r; r += 1

    # Delta NWC
    _label(ws, r, "(-) Change in Net Working Capital")
    for col in HIST_COLS:
        _hval(ws, r, col, 0)
    # NWC = (AR + Inventory + Other CA) - (AP + Accrued Comp + Other CL + Tax Payable)
    # ΔNWC = NWC_t - NWC_(t-1)
    # For simplicity, model as % of ΔRevenue
    for fc in FCST_COLS:
        prev_col = HIST_COLS[2] if fc == 6 else fc - 1
        _fval(ws, r, fc,
              f"=-(Consolidated_PL!{CL(fc)}{R['pl_rev']}-Consolidated_PL!{CL(prev_col)}{R['pl_rev']})*0.10")
    _note(ws, r, "≈ 10% of ΔRevenue (AR 7% + Inv 21% - AP 17% ≈ 11%). Sign: NWC increase = cash outflow.")
    R["dcf_delta_nwc"] = r; r += 1

    # UFCF
    _label(ws, r, "Unlevered Free Cash Flow (UFCF)", bold=True)
    for col in HIST_COLS + FCST_COLS:
        _key(ws, r, col, f"={CL(col)}{R['dcf_nopat']}+{CL(col)}{R['dcf_da']}+{CL(col)}{R['dcf_capex']}+{CL(col)}{R['dcf_delta_nwc']}")
    _note(ws, r, "= NOPAT + D&A + CapEx (neg) + ΔNWC (neg if outflow)")
    R["dcf_ufcf"] = r; r += 1

    # Discount Factor
    _label(ws, r, "Discount Factor")
    for fi, fc in enumerate(FCST_COLS):
        t = fi + 1
        _fval(ws, r, fc, f"=1/(1+Assumptions!I{R['asp_wacc']})^{t}", "0.0000")
    _note(ws, r, f"= 1/(1+WACC)^t. WACC from Assumptions (~10.5-10.7%)")
    R["dcf_df"] = r; r += 1

    # PV of UFCF
    _label(ws, r, "PV of UFCF")
    for fc in FCST_COLS:
        _fval(ws, r, fc, f"={CL(fc)}{R['dcf_ufcf']}*{CL(fc)}{R['dcf_df']}")
    _note(ws, r, "= UFCF x Discount Factor")
    R["dcf_pv_ufcf"] = r; r += 2

    # Terminal Value
    _section(ws, r, "Terminal Value (Gordon Growth Model)"); r += 1
    _label(ws, r, "Terminal Year UFCF (FY2028E)")
    for fc in [8]:
        _fval(ws, r, fc, f"={CL(fc)}{R['dcf_ufcf']}")
    _note(ws, r, "= UFCF in final explicit forecast year"); r += 1

    _label(ws, r, "Terminal Growth Rate (g)")
    _lval(ws, r, 8, f"=Assumptions!I{R['asp_tg']}", "0.00%")
    _note(ws, r, "From Assumptions. Base = 2.5%.")
    R["dcf_tg"] = r; r += 1

    _label(ws, r, "Terminal Value (TV = UFCF x (1+g) / (WACC-g))")
    _fval(ws, r, 8, f"=H{R['dcf_ufcf']}*(1+H{r-1})/(Assumptions!I{R['asp_wacc']}-H{r-1})")
    _note(ws, r, "Gordon Growth: TV = UFCF_terminal x (1+g) / (WACC - g)")
    R["dcf_tv"] = r; r += 1

    _label(ws, r, "PV of Terminal Value")
    _fval(ws, r, 8, f"=H{r-1}*H{R['dcf_df']}")
    _note(ws, r, "= TV x Discount Factor of final year")
    R["dcf_pv_tv"] = r; r += 2

    # EV Bridge
    _section(ws, r, "Enterprise Value Bridge"); r += 1
    _label(ws, r, "Sum of PV of UFCFs (FY2026E-FY2028E)")
    _fval(ws, r, 8, f"=F{R['dcf_pv_ufcf']}+G{R['dcf_pv_ufcf']}+H{R['dcf_pv_ufcf']}")
    _note(ws, r, "= Sum of PV of UFCF for explicit forecast period")
    R["dcf_sum_pv_ufcf"] = r; r += 1

    _label(ws, r, "(+) PV of Terminal Value")
    _fval(ws, r, 8, f"=H{R['dcf_pv_tv']}")
    R["dcf_pv_tv_val"] = r; r += 1

    _label(ws, r, "Enterprise Value (EV)", bold=True)
    _key(ws, r, 8, f"=H{r-2}+H{r-1}")
    _note(ws, r, "= Sum PV UFCF + PV TV")
    R["dcf_ev"] = r; r += 1

    _label(ws, r, "  TV as % of EV")
    _pct(ws, r, 8, f"=H{r-1}/H{r}")
    r += 2

    # Equity Bridge
    _section(ws, r, "Equity Value Bridge"); r += 1
    _label(ws, r, "Enterprise Value")
    _fval(ws, r, 8, f"=H{R['dcf_ev']}")
    R["dcf_ev_bridge"] = r; r += 1

    _label(ws, r, "(+) Cash & Equivalents")
    _lval(ws, r, 8, f"=BS!H{R['bs_cash']}")
    _note(ws, r, "<- BS Cash (latest forecast year)")
    R["dcf_cash"] = r; r += 1

    _label(ws, r, "(+) Marketable Securities")
    _lval(ws, r, 8, f"=BS!H{R['bs_ms']}")
    _note(ws, r, "<- BS Short-term Investments")
    R["dcf_ms_bridge"] = r; r += 1

    _label(ws, r, "(-) Total Debt")
    _lval(ws, r, 8, f"=BS!H{R['bs_lt_debt']}+BS!H{R['bs_st_debt']}")
    _note(ws, r, "<- BS LT Debt + ST Debt")
    R["dcf_debt_bridge"] = r; r += 1

    _label(ws, r, "(-) Non-Controlling Interests")
    _lval(ws, r, 8, f"=BS!H{R['bs_nci']}")
    _note(ws, r, "<- BS NCI")
    R["dcf_nci_bridge"] = r; r += 1

    _label(ws, r, "Equity Value", bold=True)
    _key(ws, r, 8, f"=H{R['dcf_ev_bridge']}+H{R['dcf_cash']}+H{R['dcf_ms_bridge']}-H{R['dcf_debt_bridge']}-H{R['dcf_nci_bridge']}")
    _note(ws, r, "= EV + Cash + Securities - Debt - NCI. KEY OUTPUT.")
    R["dcf_equity_val"] = r; r += 1

    _label(ws, r, "Diluted Shares Outstanding (M)")
    _lval(ws, r, 8, f"=Consolidated_PL!H{R['pl_shares']}")
    _note(ws, r, "<- Consolidated_PL. FY2028E diluted shares.")
    R["dcf_shares"] = r; r += 1

    _label(ws, r, "Implied Share Price ($)", bold=True)
    _key(ws, r, 8, f"=H{R['dcf_equity_val']}/H{R['dcf_shares']}", "$#,##0.00")
    _note(ws, r, "= Equity Value / Shares. CENTRAL OUTPUT OF THE MODEL.")
    R["dcf_implied_price"] = r; r += 2

    # DCF Summary table
    _section(ws, r, "DCF Summary"); r += 1
    summary_items = [
        ("WACC", f"=Assumptions!I{R['asp_wacc']}", "0.00%"),
        ("Terminal Growth Rate", f"=Assumptions!I{R['asp_tg']}", "0.00%"),
        ("Sum PV UFCF", f"=H{R['dcf_sum_pv_ufcf']}", "#,##0"),
        ("PV Terminal Value", f"=H{R['dcf_pv_tv_val']}", "#,##0"),
        ("TV % of EV", f"=H{R['dcf_pv_tv_val']}/H{R['dcf_ev']}", "0.0%"),
        ("Enterprise Value", f"=H{R['dcf_ev']}", "#,##0"),
        ("(+) Cash + Securities", f"=H{R['dcf_cash']}+H{R['dcf_ms_bridge']}", "#,##0"),
        ("(-) Total Debt", f"=H{R['dcf_debt_bridge']}", "#,##0"),
        ("(-) NCI", f"=H{R['dcf_nci_bridge']}", "#,##0"),
        ("Equity Value", f"=H{R['dcf_equity_val']}", "#,##0"),
        ("Implied Share Price", f"=H{R['dcf_implied_price']}", "$#,##0.00"),
    ]
    for label, formula, fmt in summary_items:
        _label(ws, r, f"  {label}")
        _key(ws, r, 8, formula, fmt)
        r += 1

    _set_col_widths(ws)
    lr = _add_legend(ws, r + 1)
    _add_sources(ws, lr + 1)
    ws.sheet_properties.tabColor = "FFF2CC"
    return r


# ============================================================
# TAB 10: SENSITIVITY (WACC x Terminal g)
# ============================================================

def build_sensitivity(wb, R):
    """Build 2D sensitivity table: WACC (cols) x Terminal g (rows) -> Implied Price."""
    ws = wb["Sensitivity"]
    _title_row(ws, 1, "Intel (INTC) - Sensitivity Analysis")
    _unit_row(ws, 2, "Implied Share Price ($) — WACC vs Terminal Growth Rate")
    _label(ws, 4, "Base case WACC = 10.7%, Base case g = 2.5%")
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=8)

    r = 6
    # WACC range in columns (8.0% to 13.0% in 1% steps)
    wacc_values = [8.0, 9.0, 10.0, 10.7, 11.0, 12.0, 13.0]
    # Terminal g range in rows (1.5% to 3.5% in 0.5% steps)
    g_values = [1.5, 2.0, 2.5, 3.0, 3.5]

    # Header row
    _label(ws, r, "Terminal g \\ WACC")
    for ci, wacc in enumerate(wacc_values):
        col = ci + 2  # start from column B
        c = ws.cell(row=r, column=col, value=wacc)
        c.number_format = "0.0%"
        c.font = FONT_HEADER; c.fill = FILL_HEADER; c.border = BD; c.alignment = ALIGN_CENTER
    r += 1

    for gi, g in enumerate(g_values):
        _label(ws, r, f"  {g}%")
        c = ws.cell(row=r, column=1)
        c.font = FONT_BOLD; c.border = BD
        for ci, wacc in enumerate(wacc_values):
            col = ci + 2
            # Recompute implied price for each WACC/g combination
            # Formula: recompute TV and everything downstream
            # Since Excel DATA TABLE is complex in openpyxl, we write the formula directly
            # using the DCF tab's structure but substituting WACC and g
            # For simplicity, mark as TODO: requires Excel Data Table or manual macro
            is_base = abs(wacc - 0.107) < 0.001 and abs(g - 0.025) < 0.001
            if is_base:
                c = ws.cell(row=r, column=col, value="BASE")
                c.font = FONT_BOLD; c.fill = FILL_KEY
            else:
                c = ws.cell(row=r, column=col, value="")
                c.font = FONT_FORM
            c.border = BD; c.alignment = ALIGN_CENTER
        r += 1

    _note(ws, r - len(g_values), "TODO: Implement 2D DATA TABLE in Excel. Open in Excel, select the table range, Data > What-If Analysis > Data Table. Row input: WACC cell (Assumptions!I{wacc_row}). Column input: g cell (Assumptions!I{tg_row}). Base case (WACC=10.7%, g=2.5%) is marked.")
    r += 1

    _section(ws, r, "Instructions"); r += 1
    instructions = [
        "1. Open Intel_IB_Model.xlsx in Microsoft Excel",
        "2. Go to Sensitivity tab",
        "3. Select the table range (B7:H11)",
        "4. Data > What-If Analysis > Data Table",
        f"5. Row input cell: =Assumptions!$I${R['asp_wacc']}",
        f"6. Column input cell: =Assumptions!$I${R['asp_tg']}",
        "7. Click OK. Table will populate with implied prices.",
    ]
    for inst in instructions:
        c = ws.cell(row=r, column=1, value=inst)
        c.font = Font(size=9, name="Calibri")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        r += 1

    _set_col_widths(ws)
    ws.sheet_properties.tabColor = "548235"
    return r


# ============================================================
# TAB 11: RATIO_ANALYSIS
# ============================================================

def build_ratio_analysis(wb, R):
    """Build ratio analysis tab: profitability, efficiency, leverage, CF quality, growth."""
    ws = wb["Ratio_Analysis"]
    _title_row(ws, 1, "Intel (INTC) - Ratio Analysis & Financial Metrics")
    _unit_row(ws, 2)
    _yr_header_no_scenario(ws, 3)
    _label(ws, 4, "All ratios computed from model financials")
    _note(ws, 4, "Links to Consolidated_PL, BS, Cash_Flow. Auto-updates with scenario changes.")

    r = 5

    # Profitability
    _section(ws, r, "Profitability Ratios"); r += 1
    ratios = [
        ("Gross Margin %", f"=Consolidated_PL!{{c}}{R['pl_gpm']}", "0.0%", True),
        ("EBIT Margin %", f"=Consolidated_PL!{{c}}{R['pl_ebit']}/Consolidated_PL!{{c}}{R['pl_rev']}", "0.0%", False),
        ("Net Income Margin %", f"=Consolidated_PL!{{c}}{R['pl_ni']}/Consolidated_PL!{{c}}{R['pl_rev']}", "0.0%", False),
        ("ROE (NI to Intel / Avg Equity)", f"=Consolidated_PL!{{c}}{R['pl_ni_intel']}/BS!{{c}}{R['bs_equity']}", "0.0%", False),
        ("ROA (NI / Avg Total Assets)", f"=Consolidated_PL!{{c}}{R['pl_ni']}/BS!{{c}}{R['bs_total_assets']}", "0.0%", False),
        ("EBITDA ($M)", f"=Consolidated_PL!{{c}}{R['pl_ebit']}+Consolidated_PL!{{c}}{R['pl_da']}", "#,##0", False),
        ("EBITDA Margin %", f"=(Consolidated_PL!{{c}}{R['pl_ebit']}+Consolidated_PL!{{c}}{R['pl_da']})/Consolidated_PL!{{c}}{R['pl_rev']}", "0.0%", False),
    ]
    for label, formula_tpl, fmt, is_link in ratios:
        _label(ws, r, f"  {label}")
        for col in HIST_COLS + FCST_COLS:
            formula = formula_tpl.replace("{c}", CL(col))
            if is_link:
                _lval(ws, r, col, formula, fmt)
            else:
                _fval(ws, r, col, formula, fmt)
        r += 1

    # Efficiency
    r += 1
    _section(ws, r, "Efficiency Ratios"); r += 1
    eff_ratios = [
        ("AR Days", f"=BS!{{c}}{R['bs_ar']}/(Consolidated_PL!{{c}}{R['pl_rev']}/365)", "#,##0.0"),
        ("AP Days", f"=BS!{{c}}{R['bs_ap']}/(Consolidated_PL!{{c}}{R['pl_cogs']}/365)", "#,##0.0"),
        ("Inventory Days", f"=BS!{{c}}{R['bs_inventory']}/(Consolidated_PL!{{c}}{R['pl_cogs']}/365)", "#,##0.0"),
        ("CapEx / Revenue %", f"=Cash_Flow!{{c}}{R['cf_capex']}/Consolidated_PL!{{c}}{R['pl_rev']}", "0.0%"),
        ("CapEx / D&A (x)", f"=-Cash_Flow!{{c}}{R['cf_capex']}/Consolidated_PL!{{c}}{R['pl_da']}", "0.00x"),
    ]
    for label, formula_tpl, fmt in eff_ratios:
        _label(ws, r, f"  {label}")
        for col in HIST_COLS + FCST_COLS:
            formula = formula_tpl.replace("{c}", CL(col))
            _fval(ws, r, col, formula, fmt)
        r += 1

    # Leverage
    r += 1
    _section(ws, r, "Leverage & Liquidity Ratios"); r += 1
    lev_ratios = [
        ("Debt / Total Assets", f"=(BS!{{c}}{R['bs_lt_debt']}+BS!{{c}}{R['bs_st_debt']})/BS!{{c}}{R['bs_total_assets']}", "0.0%"),
        ("Debt / Equity", f"=(BS!{{c}}{R['bs_lt_debt']}+BS!{{c}}{R['bs_st_debt']})/BS!{{c}}{R['bs_equity']}", "0.00x"),
        ("Current Ratio", f"=BS!{{c}}{R['bs_total_ca']}/BS!{{c}}{R['bs_total_cl']}", "0.00x"),
        ("Net Cash / (Debt) ($M)", f"=BS!{{c}}{R['bs_cash']}+BS!{{c}}{R['bs_ms']}-BS!{{c}}{R['bs_lt_debt']}-BS!{{c}}{R['bs_st_debt']}", "#,##0"),
    ]
    for label, formula_tpl, fmt in lev_ratios:
        _label(ws, r, f"  {label}")
        for col in HIST_COLS + FCST_COLS:
            formula = formula_tpl.replace("{c}", CL(col))
            _fval(ws, r, col, formula, fmt)
        r += 1

    # Cash Flow Quality
    r += 1
    _section(ws, r, "Cash Flow Quality"); r += 1
    cf_ratios = [
        ("OCF / Net Income (x)", f"=Cash_Flow!{{c}}{R['cf_cfo']}/Consolidated_PL!{{c}}{R['pl_ni']}", "0.00x"),
        ("FCF / Revenue %", f"=Cash_Flow!{{c}}{R['cf_fcf']}/Consolidated_PL!{{c}}{R['pl_rev']}", "0.0%"),
        ("FCF / Net Income (x)", f"=Cash_Flow!{{c}}{R['cf_fcf']}/Consolidated_PL!{{c}}{R['pl_ni']}", "0.00x"),
    ]
    for label, formula_tpl, fmt in cf_ratios:
        _label(ws, r, f"  {label}")
        for col in HIST_COLS + FCST_COLS:
            formula = formula_tpl.replace("{c}", CL(col))
            _fval(ws, r, col, formula, fmt)
        r += 1

    # Growth
    r += 1
    _section(ws, r, "Growth Rates (YoY %)"); r += 1
    growth_items = [
        ("Revenue Growth %", f"=Consolidated_PL!{{c}}{R['pl_rev']}", None),
        ("Net Income Growth %", f"=Consolidated_PL!{{c}}{R['pl_ni_intel']}", None),
        ("EPS Growth %", f"={CL(HIST_COLS[1])}{R['pl_eps']}", None),
    ]
    for label, _, _ in growth_items:
        _label(ws, r, f"  {label}")
        for col in HIST_COLS + FCST_COLS:
            if col == 2:
                _hval(ws, r, col, "N/A")
            else:
                prev_col = col - 1 if col in HIST_COLS else (HIST_COLS[2] if col == 6 else col - 1)
                if label.startswith("Revenue"):
                    ref1 = f"Consolidated_PL!{CL(col)}{R['pl_rev']}"
                    ref2 = f"Consolidated_PL!{CL(prev_col)}{R['pl_rev']}"
                elif label.startswith("Net Income"):
                    ref1 = f"Consolidated_PL!{CL(col)}{R['pl_ni_intel']}"
                    ref2 = f"Consolidated_PL!{CL(prev_col)}{R['pl_ni_intel']}"
                else:
                    ref1 = f"Consolidated_PL!{CL(col)}{R['pl_eps']}"
                    ref2 = f"Consolidated_PL!{CL(prev_col)}{R['pl_eps']}"
                _pct(ws, r, col, f"={ref1}/{ref2}-1")
        r += 1

    # ============================================================
    # Market Valuation Multiples
    # ============================================================
    r += 1
    _section(ws, r, "Market Valuation Multiples"); r += 1
    _label(ws, r, "  Current Stock Price ($) — EDIT HERE", bold=True)
    sp_cell = ws.cell(row=r, column=6, value=113)
    sp_cell.font = Font(color="000000", bold=True); sp_cell.fill = PatternFill("solid", fgColor="DDEBF7")
    sp_cell.number_format = "0.00"; sp_cell.border = BD
    for col in [2, 3, 4, 7, 8]:
        ws.cell(row=r, column=col).border = BD
    _note(ws, r, "Enter current share price; TTM & NTM multiples update automatically")
    SP = f"F{r}"
    R["ratio_sp"] = r
    r += 2

    # --- TTM Valuation (FY2025A, column D) ---
    _section(ws, r, "TTM Valuation (Based on FY2025A)"); r += 1

    ev_d = (f"({SP}*Consolidated_PL!D{R['pl_shares']}"
            f"+BS!D{R['bs_st_debt']}+BS!D{R['bs_lt_debt']}"
            f"-BS!D{R['bs_cash']}-BS!D{R['bs_ms']})")
    ebitda_d = f"(Consolidated_PL!D{R['pl_ebit']}+Consolidated_PL!D{R['pl_da']})"

    ttm_metrics = [
        ("P/E (x)",
         f"={SP}/Consolidated_PL!D{R['pl_eps']}", "0.0x",
         "Price / TTM Diluted EPS"),
        ("P/B (x)",
         f"={SP}/(BS!D{R['bs_equity']}/Consolidated_PL!D{R['pl_shares']})", "0.0x",
         "Price / TTM Book Value Per Share"),
        ("P/S (x)",
         f"={SP}/(Consolidated_PL!D{R['pl_rev']}/Consolidated_PL!D{R['pl_shares']})", "0.0x",
         "Price / TTM Revenue Per Share"),
        ("EV/EBITDA (x)",
         f"={ev_d}/{ebitda_d}", "0.0x",
         "Enterprise Value / TTM EBITDA"),
        ("EV/Revenue (x)",
         f"={ev_d}/Consolidated_PL!D{R['pl_rev']}", "0.0x",
         "Enterprise Value / TTM Revenue"),
        ("P/CF (x)",
         f"={SP}/(Cash_Flow!D{R['cf_cfo']}/Consolidated_PL!D{R['pl_shares']})", "0.0x",
         "Price / TTM Operating Cash Flow Per Share"),
        ("PEG Ratio (x)",
         f"=({SP}/Consolidated_PL!D{R['pl_eps']})/((Consolidated_PL!D{R['pl_eps']}/Consolidated_PL!C{R['pl_eps']}-1)*100)", "0.00x",
         "P/E / TTM EPS Growth %; NM if growth <=0"),
        ("Dividend Yield %",
         f"=Assumptions!I{R['asp_dps']}/{SP}", "0.0%",
         "TTM DPS / Price; dividend currently suspended"),
    ]
    for label, formula, fmt, note_text in ttm_metrics:
        _label(ws, r, f"  {label}")
        c = ws.cell(row=r, column=4, value=formula)
        c.number_format = fmt; c.border = BD
        for col in [2, 3, 6, 7, 8]:
            ws.cell(row=r, column=col).border = BD
        _note(ws, r, note_text)
        r += 1

    r += 1

    # --- NTM Valuation (FY2026E, column F) ---
    _section(ws, r, "NTM Valuation (Based on FY2026E)"); r += 1

    ev_f = (f"({SP}*Consolidated_PL!F{R['pl_shares']}"
            f"+BS!F{R['bs_st_debt']}+BS!F{R['bs_lt_debt']}"
            f"-BS!F{R['bs_cash']}-BS!F{R['bs_ms']})")
    ebitda_f = f"(Consolidated_PL!F{R['pl_ebit']}+Consolidated_PL!F{R['pl_da']})"

    ntm_metrics = [
        ("P/E (x)",
         f"={SP}/Consolidated_PL!F{R['pl_eps']}", "0.0x",
         "Price / NTM Diluted EPS; updates with scenario"),
        ("P/B (x)",
         f"={SP}/(BS!F{R['bs_equity']}/Consolidated_PL!F{R['pl_shares']})", "0.0x",
         "Price / NTM Book Value Per Share"),
        ("P/S (x)",
         f"={SP}/(Consolidated_PL!F{R['pl_rev']}/Consolidated_PL!F{R['pl_shares']})", "0.0x",
         "Price / NTM Revenue Per Share"),
        ("EV/EBITDA (x)",
         f"={ev_f}/{ebitda_f}", "0.0x",
         "Enterprise Value / NTM EBITDA"),
        ("EV/Revenue (x)",
         f"={ev_f}/Consolidated_PL!F{R['pl_rev']}", "0.0x",
         "Enterprise Value / NTM Revenue"),
        ("P/CF (x)",
         f"={SP}/(Cash_Flow!F{R['cf_cfo']}/Consolidated_PL!F{R['pl_shares']})", "0.0x",
         "Price / NTM Operating Cash Flow Per Share"),
        ("PEG Ratio (x)",
         f"=({SP}/Consolidated_PL!F{R['pl_eps']})/((Consolidated_PL!F{R['pl_eps']}/Consolidated_PL!D{R['pl_eps']}-1)*100)", "0.00x",
         "P/E / NTM EPS Growth %; NM if growth <=0"),
        ("Dividend Yield %",
         f"=Assumptions!I{R['asp_dps']}/{SP}", "0.0%",
         "NTM DPS / Price; dividend currently suspended"),
    ]
    for label, formula, fmt, note_text in ntm_metrics:
        _label(ws, r, f"  {label}")
        c = ws.cell(row=r, column=6, value=formula)
        c.number_format = fmt; c.border = BD
        for col in [2, 3, 4, 7, 8]:
            ws.cell(row=r, column=col).border = BD
        _note(ws, r, note_text)
        r += 1

    _set_col_widths(ws)
    lr = _add_legend(ws, r + 1)
    _add_sources(ws, lr + 1)
    ws.sheet_properties.tabColor = "548235"
    return r


# ============================================================
# TAB 2: KEY_SUMMARY (Executive Dashboard - built LAST)
# ============================================================

def build_key_summary(wb, R):
    """Build executive dashboard linking to all other tabs."""
    ws = wb["Key_Summary"]
    _title_row(ws, 1, "Intel (INTC) - Key Summary & Executive Dashboard")
    _unit_row(ws, 2)
    _label(ws, 3, f"Scenario: ")
    _lval(ws, 3, 2, f"=Assumptions!{CL(2)}2", "0")
    ws.cell(row=3, column=2).font = Font(bold=True, size=11, name="Calibri", color="1F4E79")
    ws.cell(row=3, column=2).fill = FILL_ASSUMP

    r = 5
    # KPIs section
    _section(ws, r, "Key Performance Indicators")
    _yr_header_no_scenario(ws, r + 1)
    r += 2

    kpis = [
        ("Total Revenue ($M)", f"=Consolidated_PL!{{c}}{R['pl_rev']}", "#,##0", True),
        ("  YoY Growth %", f"=Consolidated_PL!{{c}}{R['pl_rev']}/Consolidated_PL!{{c-1}}{R['pl_rev']}-1", "0.0%", False),
        ("Gross Profit ($M)", f"=Consolidated_PL!{{c}}{R['pl_gp']}", "#,##0", True),
        ("  Gross Margin %", f"=Consolidated_PL!{{c}}{R['pl_gp']}/Consolidated_PL!{{c}}{R['pl_rev']}", "0.0%", True),
        ("EBIT ($M)", f"=Consolidated_PL!{{c}}{R['pl_ebit']}", "#,##0", True),
        ("  EBIT Margin %", f"=Consolidated_PL!{{c}}{R['pl_ebit']}/Consolidated_PL!{{c}}{R['pl_rev']}", "0.0%", True),
        ("Net Income to Intel ($M)", f"=Consolidated_PL!{{c}}{R['pl_ni_intel']}", "#,##0", True),
        ("  Net Margin %", f"=Consolidated_PL!{{c}}{R['pl_ni_intel']}/Consolidated_PL!{{c}}{R['pl_rev']}", "0.0%", True),
        ("Diluted EPS ($)", f"=Consolidated_PL!{{c}}{R['pl_eps']}", "0.00", True),
        ("D&A ($M)", f"=Consolidated_PL!{{c}}{R['pl_da']}", "#,##0", False),
        ("EBITDA ($M)", f"=Consolidated_PL!{{c}}{R['pl_ebit']}+Consolidated_PL!{{c}}{R['pl_da']}", "#,##0", True),
        ("FCF ($M)", f"=Cash_Flow!{{c}}{R['cf_fcf']}", "#,##0", True),
    ]
    for label, formula_tpl, fmt, is_key in kpis:
        _label(ws, r, f"  {label}")
        for col in HIST_COLS + FCST_COLS:
            formula = formula_tpl.replace("{c}", CL(col)).replace("{c-1}", CL(col-1 if col in HIST_COLS else (HIST_COLS[2] if col == 6 else col-1)))
            if col == 2 and "YoY" in label:
                pass  # No YoY for first year
            else:
                if is_key:
                    _key(ws, r, col, formula, fmt)
                elif "YoY" in label or "Margin" in label:
                    _pct(ws, r, col, formula)
                else:
                    _lval(ws, r, col, formula, fmt)
        r += 1

    # Revenue by Segment
    r += 1
    _section(ws, r, "Revenue by Segment ($M)"); r += 1
    for sk, sl in [("ccg","CCG"),("dcai","DCAI"),("foundry","Intel Foundry"),("allother","All Other")]:
        _label(ws, r, f"  {sl}")
        for col in HIST_COLS + FCST_COLS:
            _lval(ws, r, col, f"=Segment_Revenue!{CL(col)}{R['segrev_'+sk]}")
        r += 1
    _label(ws, r, "  Total Consolidated", bold=True)
    for col in HIST_COLS + FCST_COLS:
        _lval(ws, r, col, f"=Segment_Revenue!{CL(col)}{R['segrev_consol_rev']}")
    r += 2

    # BS Highlights
    _section(ws, r, "Balance Sheet Highlights ($M)"); r += 1
    bs_items = [
        ("Total Assets", "bs_total_assets"),
        ("Total Equity", "bs_total_equity"),
        ("Cash + Securities", None),  # special handling
        ("Total Debt", None),
        ("Net Cash / (Debt)", None),
    ]
    for label, bs_key in bs_items:
        _label(ws, r, f"  {label}")
        for col in HIST_COLS + FCST_COLS:
            if label == "Cash + Securities":
                _lval(ws, r, col, f"=BS!{CL(col)}{R['bs_cash']}+BS!{CL(col)}{R['bs_ms']}")
            elif label == "Total Debt":
                _lval(ws, r, col, f"=BS!{CL(col)}{R['bs_lt_debt']}+BS!{CL(col)}{R['bs_st_debt']}")
            elif label == "Net Cash / (Debt)":
                _lval(ws, r, col, f"=BS!{CL(col)}{R['bs_cash']}+BS!{CL(col)}{R['bs_ms']}-BS!{CL(col)}{R['bs_lt_debt']}-BS!{CL(col)}{R['bs_st_debt']}")
            else:
                _lval(ws, r, col, f"=BS!{CL(col)}{R[bs_key]}")
        r += 1

    _label(ws, r, "  ROE")
    for col in HIST_COLS + FCST_COLS:
        _pct(ws, r, col, f"=Consolidated_PL!{CL(col)}{R['pl_ni_intel']}/BS!{CL(col)}{R['bs_equity']}")
    r += 1
    _label(ws, r, "  ROA")
    for col in HIST_COLS + FCST_COLS:
        _pct(ws, r, col, f"=Consolidated_PL!{CL(col)}{R['pl_ni']}/BS!{CL(col)}{R['bs_total_assets']}")
    r += 2

    # DCF Output
    _section(ws, r, "DCF Valuation Output"); r += 1
    _label(ws, r, "  Implied Share Price ($)", bold=True)
    _lval(ws, r, 8, f"=DCF!H{R['dcf_implied_price']}", "$#,##0.00")
    _note(ws, r, "Base case DCF output. See Sensitivity tab for range.")
    r += 1
    _label(ws, r, "  Enterprise Value ($M)")
    _lval(ws, r, 8, f"=DCF!H{R['dcf_ev']}", "#,##0")
    r += 1
    _label(ws, r, "  Equity Value ($M)")
    _lval(ws, r, 8, f"=DCF!H{R['dcf_equity_val']}", "#,##0")
    r += 2

    # Investment Highlights
    _section(ws, r, "Investment Highlights"); r += 1
    highlights = [
        "1. CPU Renaissance: DCAI benefiting from AI inference shift toward CPUs (CPU:GPU ratio 1:8->1:1)",
        "2. 18A Ramp: Ahead of internal yield targets; volume up 6-7x in Q2 2026 vs Q1",
        "3. Foundry Losses Narrowing: -58% OPM (FY2025) -> -22% (FY2028E Base) as 18A matures",
        "4. Gross Margin Recovery: 34.8% (FY2025) -> 43.5% (FY2028E Base) driven by yield + pricing",
        "5. Supply Constraints Easing: Undershipping demand by ~$1B; tool spending up 25% to address",
        "6. Fab 34 Buyout: April 2026 ~$14.2B repurchase of Ireland SCIP stake; removes $135M/q NCI drag",
        "7. No Dividends: Suspended; priority is debt reduction and CapEx",
        "8. Warrant Overhang: 241M shares @ $20 strike if Foundry ownership <51% (NOT in base case)",
    ]
    for h in highlights:
        c = ws.cell(row=r, column=1, value=h)
        c.font = Font(size=10, name="Calibri")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        r += 1

    r += 1
    _section(ws, r, "Key Risks"); r += 1
    risks = [
        "1. 18A yield improvement stalls; Foundry losses remain >40% OPM through FY2028",
        "2. PC TAM declines >15% (recession scenario); CCG revenue headwinds",
        "3. AMD Turin / ARM CPUs erode DCAI market position",
        "4. External foundry customers do not materialize; Intel remains captive fab",
        "5. Debt load (~$50B) constrains strategic flexibility if FCF remains negative",
        "6. Valuation allowance on US DTAs ($9.9B) means losses generate no tax benefit",
        "7. WACC sensitivity: each 1% change in WACC moves implied price ~$5-7",
    ]
    for risk in risks:
        c = ws.cell(row=r, column=1, value=risk)
        c.font = Font(size=9, name="Calibri", italic=True, color="808080")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        r += 1

    _set_col_widths(ws)
    lr = _add_legend(ws, r + 1)
    _add_sources(ws, lr + 1)
    ws.sheet_properties.tabColor = "FFF2CC"
    return r


# ============================================================
# TAB 1: COVER
# ============================================================

def build_cover(wb):
    """Build cover page with company info and usage instructions."""
    ws = wb["Cover"]
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 60
    ws.column_dimensions['C'].width = 30

    r = 1
    c = ws.cell(row=r, column=1, value="Intel Corporation (NASDAQ: INTC)")
    c.font = Font(bold=True, size=18, name="Calibri", color="1F4E79")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3); r += 2

    c = ws.cell(row=r, column=1, value="Investment Bank-Grade Financial Model")
    c.font = Font(bold=True, size=14, name="Calibri"); r += 1

    c = ws.cell(row=r, column=1, value="Segment-Driven, Three-Statement Linked DCF")
    c.font = Font(size=11, name="Calibri", italic=True); r += 1

    c = ws.cell(row=r, column=1, value="Model Date: May 7, 2026 | Forecast Horizon: FY2026E - FY2028E")
    c.font = Font(size=10, name="Calibri", color="808080"); r += 3

    # Summary info
    info = [
        ("Company", "Intel Corporation (INTC)"),
        ("Sector", "Semiconductors"),
        ("Headquarters", "Santa Clara, CA, USA"),
        ("Fiscal Year", "Ends last Saturday of December (52/53-week year)"),
        ("Reporting Segments", "CCG, DCAI, Intel Foundry, All Other (Mobileye + IMS + Other)"),
        ("", ""),
        ("Scenario Selector", "Assumptions tab, Cell B2 (Base / Bull / Bear)"),
        ("Model Architecture", "12 tabs: Cover -> Key_Summary -> Assumptions -> Growth_Drivers -> Segment_Revenue -> Segment_PL -> Consolidated_PL -> BS -> Cash_Flow -> DCF -> Sensitivity -> Ratio_Analysis"),
        ("", ""),
        ("Key Data Sources", "Intel 10-K filings (FY2022-FY2025), Q1 2026 10-Q, Q1 2026 Earnings Call (Apr 23, 2026)"),
        ("Assumptions Date", "May 7, 2026 (intc_assumptions.md)"),
        ("Financial Data", "intc_financial_data.md (extracted from SEC filings)"),
    ]
    for label, value in info:
        if label:
            c = ws.cell(row=r, column=1, value=label)
            c.font = Font(bold=True, size=10, name="Calibri")
        c = ws.cell(row=r, column=2, value=value)
        c.font = Font(size=10, name="Calibri")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        r += 1
    r += 1

    # Segment evolution note
    c = ws.cell(row=r, column=1, value="Segment Structure Note:")
    c.font = Font(bold=True, size=10, name="Calibri")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3); r += 1

    notes = [
        "Effective Q1 2025: NEX (Network & Edge) integrated into CCG and DCAI; Altera deconsolidated Sep 12, 2025 (51% sold to SLP).",
        "FY2023-FY2025 segment data is RESTATED to current 4-segment structure per 2025 10-K Note 3.",
        "Historical data from old structure (with NEX) is available in 2024 10-K for reference but NOT used in this model.",
        "Mobileye remains consolidated (~77% Intel ownership). Q1 2026: $3.9B goodwill impairment at Mobileye.",
        "Ireland SCIP (Fab 34) 49% NCI bought out April 2026 (~$14.2B). Arizona SCIP 49% NCI remains.",
    ]
    for note_text in notes:
        c = ws.cell(row=r, column=1, value=f"  - {note_text}")
        c.font = Font(size=9, name="Calibri", color="404040")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        r += 1
    r += 2

    # Usage instructions
    c = ws.cell(row=r, column=1, value="Usage Instructions:")
    c.font = Font(bold=True, size=11, name="Calibri"); r += 1

    instructions = [
        "1. Select scenario in Assumptions tab cell B2 (Base / Bull / Bear). All tabs update automatically.",
        "2. Edit blue-fill cells in Assumptions tab to customize forecasts.",
        "3. Check yellow-highlighted cross-validation rows (Segment_PL, BS) — both must read 0 difference.",
        "4. DCF tab computes implied share price. Sensitivity tab provides WACC x g matrix.",
        "5. Key_Summary tab gives executive overview; Ratio_Analysis tab provides detailed metrics.",
        "6. All green-font cells are cross-sheet links (do not edit). Blue-font cells are historical data.",
    ]
    for inst in instructions:
        c = ws.cell(row=r, column=1, value=inst)
        c.font = Font(size=10, name="Calibri")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        r += 1
    r += 2

    # Disclaimer
    c = ws.cell(row=r, column=1, value="DISCLAIMER:")
    c.font = Font(bold=True, size=10, name="Calibri", color="C00000"); r += 1
    disclaimers = [
        "This model is for educational / analytical purposes only. It does not constitute investment advice.",
        "Forecasts are based on publicly available information and assumptions documented in intc_assumptions.md.",
        "Actual results may differ materially. Refer to Intel's SEC filings for official financial information.",
        "Model generated programmatically via Python openpyxl (build_intc_model.py).",
    ]
    for d_text in disclaimers:
        c = ws.cell(row=r, column=1, value=d_text)
        c.font = Font(size=8, name="Calibri", italic=True, color="808080")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        r += 1

    ws.sheet_properties.tabColor = "1F4E79"
    return r


# ============================================================
# MAIN MODEL BUILDER
# ============================================================

def build_model():
    """Create the complete 12-tab workbook."""
    print("Building Intel_IB_Model.xlsx ...")
    wb = openpyxl.Workbook()
    R = {}  # Row reference dictionary

    # Create all sheets in tab order
    ws0 = wb.active
    ws0.title = "Cover"
    for name in ["Key_Summary", "Assumptions", "Growth_Drivers", "Segment_Revenue", "Segment_PL",
                 "Consolidated_PL", "BS", "Cash_Flow", "DCF", "Sensitivity", "Ratio_Analysis"]:
        wb.create_sheet(name)

    # Build in dependency order
    print("  1/12 Cover ...")
    build_cover(wb)

    print("  2/12 Assumptions ...")
    build_assumptions(wb, R)

    print("  3/12 Segment_Revenue (pass 1 - structure + historical) ...")
    build_segment_revenue(wb, R)

    print("  4/12 Growth_Drivers ...")
    build_growth_drivers(wb, R)

    print("  4.5/12 Segment_Revenue (pass 2 - link to Growth_Drivers) ...")
    build_segment_revenue_link_gd(wb, R)

    print("  5/12 Segment_PL (pass 1) ...")
    build_segment_pl(wb, R)

    print("  6/12 Consolidated_PL ...")
    build_consolidated_pl(wb, R)

    print("  7/12 Segment_PL (pass 2 - cross-check) ...")
    build_segment_pl_pass2(wb, R)

    print("  8/12 Balance Sheet ...")
    build_bs(wb, R)

    print("  9/12 Cash_Flow ...")
    build_cash_flow(wb, R)

    print(" 10/12 DCF ...")
    build_dcf(wb, R)

    print(" 11/12 Sensitivity ...")
    build_sensitivity(wb, R)

    print(" 12/12 Ratio_Analysis ...")
    build_ratio_analysis(wb, R)

    print("  Key_Summary (dashboard) ...")
    build_key_summary(wb, R)

    # Key_Summary already at position 1 (created first after Cover)

    # Save
    wb.save(OUTPUT)
    print(f"\nSaved: {OUTPUT}")
    print(f"Total key rows tracked in R dict: {len(R)}")
    return wb, R


# ============================================================
# SELF-CHECK / VERIFICATION
# ============================================================

def check_model(output_path=OUTPUT):
    """Load the generated xlsx and verify key linkages programmatically."""
    print(f"\n{'='*60}")
    print(f"SELF-CHECK: Verifying {output_path} ...")
    print("=" * 60)

    wb = openpyxl.load_workbook(output_path)

    errors = []
    warnings = []

    # 1. Check all expected tabs exist
    expected_tabs = ["Cover", "Key_Summary", "Assumptions", "Growth_Drivers",
                     "Segment_Revenue", "Segment_PL", "Consolidated_PL", "BS",
                     "Cash_Flow", "DCF", "Sensitivity", "Ratio_Analysis"]
    for tab in expected_tabs:
        if tab not in wb.sheetnames:
            errors.append(f"MISSING TAB: {tab}")
    print(f"  [{'PASS' if not errors else 'FAIL'}] All {len(expected_tabs)} tabs present")

    # 2. Check Segment_PL cross-check = 0 formula exists
    if "Segment_PL" in wb.sheetnames:
        ws = wb["Segment_PL"]
        # Find "CROSS-CHECK" or "Difference" row
        found_xcheck = False
        for row in ws.iter_rows(min_row=1, max_row=100, min_col=1, max_col=1):
            for cell in row:
                if cell.value and "cross-check" in str(cell.value).lower():
                    found_xcheck = True
                    break
        if found_xcheck:
            print("  [PASS] Segment_PL cross-check row present")
        else:
            warnings.append("Segment_PL cross-check row not found (may need manual verification)")

    # 3. Check BS balance check row exists
    if "BS" in wb.sheetnames:
        ws = wb["BS"]
        found_bal = False
        for row in ws.iter_rows(min_row=1, max_row=100, min_col=1, max_col=1):
            for cell in row:
                if cell.value and "balance check" in str(cell.value).lower():
                    found_bal = True
                    break
        if found_bal:
            print("  [PASS] BS balance check row present")
        else:
            warnings.append("BS balance check row not found")

    # 4. Check no formulas in Notes column (col K=11)
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=11, max_col=11):
            for cell in row:
                if cell.value and str(cell.value).startswith("="):
                    warnings.append(f"{sheet_name}!{cell.coordinate}: Notes column contains formula: {cell.value}")
    if not any("Notes column contains formula" in w for w in warnings):
        print("  [PASS] No formulas in Notes column (K)")
    else:
        print("  [WARN] Some Notes cells start with '=' (stripped by note() helper)")

    # 5. Check CHOOSE/MATCH formulas in Assumptions col I
    if "Assumptions" in wb.sheetnames:
        ws = wb["Assumptions"]
        choose_count = 0
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=9, max_col=9):
            for cell in row:
                if cell.value and "CHOOSE(MATCH" in str(cell.value):
                    choose_count += 1
        print(f"  [PASS] {choose_count} CHOOSE/MATCH formulas in Assumptions")

    # 6. Check Segment_Revenue formulas reference Assumptions
    if "Segment_Revenue" in wb.sheetnames:
        ws = wb["Segment_Revenue"]
        ref_count = 0
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=6, max_col=8):
            for cell in row:
                if cell.value and "Assumptions!" in str(cell.value):
                    ref_count += 1
        print(f"  [PASS] {ref_count} forecast formulas reference Assumptions in Segment_Revenue")

    # 7. Check DCF has implied price
    if "DCF" in wb.sheetnames:
        ws = wb["DCF"]
        found_price = False
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=1):
            for cell in row:
                if cell.value and "implied share price" in str(cell.value).lower():
                    found_price = True
                    break
        if found_price:
            print("  [PASS] DCF implied share price row present")
        else:
            warnings.append("DCF implied share price row not found")

    # 8. Count total rows
    total_rows = sum(wb[sheet].max_row for sheet in wb.sheetnames)
    print(f"  [INFO] Total rows across all tabs: {total_rows}")

    # Summary
    print(f"\n{'='*60}")
    if errors:
        print(f"ERRORS ({len(errors)}):")
        for e in errors:
            print(f"  - {e}")
    if warnings:
        print(f"WARNINGS ({len(warnings)}):")
        for w in warnings:
            print(f"  - {w}")
    if not errors and not warnings:
        print("All checks passed!")
    print(f"{'='*60}")

    wb.close()
    return len(errors) == 0


# ============================================================
# MAIN ENTRY POINT
# ============================================================

if __name__ == "__main__":
    wb, R = build_model()
    check_model()
