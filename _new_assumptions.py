def build_assumptions(wb, R):
    """Build the Assumptions tab with quarterly columns, Base/Bull/Bear rows,
    and CHOOSE/MATCH Selected row. User adjusts each quarter independently.
    Annual = SUM or AVERAGE of quarterly values."""
    ws = wb["Assumptions"]
    _title_row(ws, 1, "Intel (INTC) - Assumptions & Scenario Design (Quarterly)")
    _setup_scenario_selector(ws)

    # Row 3: spacer, Row 4: column headers
    r = 4
    # Build quarterly header matching Growth_Drivers
    hdrs = {
        2: "Q1'23A", 3: "Q2'23A", 4: "Q3'23A", 5: "Q4'23A", 6: "FY2023A",
        7: "Q1'24A", 8: "Q2'24A", 9: "Q3'24A", 10: "Q4'24A", 11: "FY2024A",
        12: "Q1'25A", 13: "Q2'25A", 14: "Q3'25A", 15: "Q4'25A", 16: "FY2025A",
        18: "Q1'26E", 19: "Q2'26E", 20: "Q3'26E", 21: "Q4'26E", 22: "FY2026E",
        23: "Q1'27E", 24: "Q2'27E", 25: "Q3'27E", 26: "Q4'27E", 27: "FY2027E",
        28: "Q1'28E", 29: "Q2'28E", 30: "Q3'28E", 31: "Q4'28E", 32: "FY2028E",
    }
    ANN_COLS_SET = {6, 11, 16, 22, 27, 32}
    for col, txt in hdrs.items():
        c = ws.cell(row=r, column=col, value=txt)
        c.font = FONT_HEADER
        c.fill = PatternFill("solid", fgColor="2E75B6") if col in ANN_COLS_SET else FILL_HEADER
        c.border = BD; c.alignment = ALIGN_CENTER
    for col in [1, 17, 33, 34]:
        c = ws.cell(row=r, column=col, value="")
        c.fill = FILL_HEADER; c.border = BD
    c = ws.cell(row=r, column=GD_NCOL, value="Notes")
    c.font = FONT_HEADER; c.fill = FILL_HEADER; c.border = BD; c.alignment = ALIGN_CENTER

    # All 6 years iteration
    all_years = [
        (2023, [2,3,4,5], 6), (2024, [7,8,9,10], 11), (2025, [12,13,14,15], 16),
        (2026, [18,19,20,21], 22), (2027, [23,24,25,26], 27), (2028, [28,29,30,31], 32),
    ]
    hist_years = all_years[:3]
    fcst_years = all_years[3:]

    r = 6

    # ============================================================
    # Helper: write quarterly assumption with Base/Bull/Bear/Selected rows
    # ============================================================
    def _aq(label, hist_qvals, base_qvals, bull_qvals, bear_qvals, fmt_val, note_text, driver_type="sum"):
        """Quarterly assumption: 5 rows (label + Base/Bull/Bear/Selected).
        hist_qvals: dict {year: [q1,q2,q3,q4]} or None (uses annual/4 of first hist year).
        base_qvals: dict {year: [q1,q2,q3,q4]} for forecast quarters.
        driver_type: 'sum' -> annual=SUM(Q1:Q4), 'avg' -> annual=AVERAGE(Q1:Q4).
        Returns: (next_row, selected_row_num)"""
        nonlocal r
        _label(ws, r, label)
        r_label = r; r += 1

        base_r = r; r += 1
        bull_r = r; r += 1
        bear_r = r; r += 1
        sel_r = r; r += 1

        # Column indices for iteration
        all_qcols = []
        for yr, qcols, ann_col in all_years:
            all_qcols.append((yr, qcols, ann_col))
        hist_qcols = [(yr, qcols, ann_col) for yr, qcols, ann_col in hist_years]
        fcst_qcols = [(yr, qcols, ann_col) for yr, qcols, ann_col in fcst_years]

        ann_fn = "SUM" if driver_type == "sum" else "AVERAGE"

        for scenario, srow in [("Base", base_r), ("Bull", bull_r), ("Bear", bear_r)]:
            ws.cell(row=srow, column=1, value=f"  {scenario}").font = Font(bold=True, size=10, name="Calibri")
            ws.cell(row=srow, column=1).border = BD
            # Historical: same across scenarios, put in Base row only (leave Bull/Bear blank for clarity)
            if scenario == "Base":
                for yr, qcols, ann_col in hist_years:
                    qv = hist_qvals.get(yr) if hist_qvals else None
                    for qi, qc in enumerate(qcols):
                        if qv:
                            _hval(ws, srow, qc, qv[qi], fmt_val)
                        else:
                            ws.cell(row=srow, column=qc).border = BD
                    if qv:
                        if driver_type == "sum":
                            _fval(ws, srow, ann_col, f"=SUM({CL(qcols[0])}{srow}:{CL(qcols[3])}{srow})", fmt_val)
                        else:
                            _fval(ws, srow, ann_col, f"=AVERAGE({CL(qcols[0])}{srow}:{CL(qcols[3])}{srow})", fmt_val)
                    else:
                        ws.cell(row=srow, column=ann_col).border = BD
            else:
                for yr, qcols, ann_col in hist_years:
                    for qc in qcols:
                        ws.cell(row=srow, column=qc).border = BD
                    ws.cell(row=srow, column=ann_col).border = BD
            # Forecast: editable quarterly values (blue fill)
            for yr, qcols, ann_col in fcst_years:
                qv = {"Base": base_qvals, "Bull": bull_qvals, "Bear": bear_qvals}[scenario].get(yr)
                for qi, qc in enumerate(qcols):
                    if qv:
                        c = ws.cell(row=srow, column=qc, value=qv[qi])
                        c.number_format = fmt_val
                        c.font = Font(color="000000", size=10, name="Calibri")
                        c.fill = FILL_ASSUMP
                        c.border = BD
                        c.alignment = ALIGN_RIGHT
                    else:
                        ws.cell(row=srow, column=qc).border = BD
                # Annual = SUM/AVERAGE of quarters
                if qv:
                    if driver_type == "sum":
                        _fval(ws, srow, ann_col, f"=SUM({CL(qcols[0])}{srow}:{CL(qcols[3])}{srow})", fmt_val)
                    else:
                        _fval(ws, srow, ann_col, f"=AVERAGE({CL(qcols[0])}{srow}:{CL(qcols[3])}{srow})", fmt_val)
                else:
                    ws.cell(row=srow, column=ann_col).border = BD
            # Spacer + note cols
            for sc in [17, 33, 34]:
                ws.cell(row=srow, column=sc).border = BD

        # Selected row
        ws.cell(row=sel_r, column=1, value="  SELECTED").font = Font(bold=True, size=10, name="Calibri", color="1F4E79")
        ws.cell(row=sel_r, column=1).border = BD
        # Historical: hardcoded (same as Base)
        for yr, qcols, ann_col in hist_years:
            qv = hist_qvals.get(yr) if hist_qvals else None
            for qi, qc in enumerate(qcols):
                if qv:
                    _hval(ws, sel_r, qc, qv[qi], fmt_val)
                else:
                    ws.cell(row=sel_r, column=qc).border = BD
            if qv:
                if driver_type == "sum":
                    _fval(ws, sel_r, ann_col, f"=SUM({CL(qcols[0])}{sel_r}:{CL(qcols[3])}{sel_r})", fmt_val)
                else:
                    _fval(ws, sel_r, ann_col, f"=AVERAGE({CL(qcols[0])}{sel_r}:{CL(qcols[3])}{sel_r})", fmt_val)
            else:
                ws.cell(row=sel_r, column=ann_col).border = BD
        # Forecast: CHOOSE/MATCH per quarter
        for yr, qcols, ann_col in fcst_years:
            for qc in qcols:
                cl = CL(qc)
                formula = f'=CHOOSE(MATCH({SCENARIO_CELL},{{"Base","Bull","Bear"}},0),{cl}{base_r},{cl}{bull_r},{cl}{bear_r})'
                c = ws.cell(row=sel_r, column=qc, value=formula)
                c.number_format = fmt_val
                c.font = FONT_BOLD
                c.border = BD
                c.alignment = ALIGN_RIGHT
            if driver_type == "sum":
                _fval(ws, sel_r, ann_col, f"=SUM({CL(qcols[0])}{sel_r}:{CL(qcols[3])}{sel_r})", fmt_val)
            else:
                _fval(ws, sel_r, ann_col, f"=AVERAGE({CL(qcols[0])}{sel_r}:{CL(qcols[3])}{sel_r})", fmt_val)
        for sc in [17, 33, 34]:
            ws.cell(row=sel_r, column=sc).border = BD
        # Notes
        c = ws.cell(row=sel_r, column=GD_NCOL, value=note_text)
        c.font = Font(size=9, name="Calibri"); c.alignment = ALIGN_LEFT

        # Fill key cells with yellow for Selected row annual cols
        for yr, qcols, ann_col in fcst_years:
            ws.cell(row=sel_r, column=ann_col).fill = FILL_KEY

        r += 1  # blank row after each driver group
        return sel_r

    # ============================================================
    # Helper: annual-only parameter (still uses Base/Bull/Bear rows but only annual columns)
    # ============================================================
    def _a_annual(label, hist_vals, base_vals, bull_vals, bear_vals, fmt_val, note_text):
        """Annual parameter: Base/Bull/Bear rows, values only in FY columns (6,11,16,22,27,32)."""
        nonlocal r
        _label(ws, r, label)
        r += 1
        base_r = r; r += 1
        bull_r = r; r += 1
        bear_r = r; r += 1
        sel_r = r; r += 2

        ann_hist = [6, 11, 16]
        ann_fcst = [22, 27, 32]

        for scenario, srow, vals in [("Base", base_r, base_vals), ("Bull", bull_r, bull_vals), ("Bear", bear_r, bear_vals)]:
            ws.cell(row=srow, column=1, value=f"  {scenario}").font = Font(bold=True, size=10, name="Calibri")
            ws.cell(row=srow, column=1).border = BD
            # Historical (same across all, only in Base)
            if scenario == "Base":
                for ci, ac in enumerate(ann_hist):
                    if hist_vals and ci < len(hist_vals) and hist_vals[ci] is not None:
                        _hval(ws, srow, ac, hist_vals[ci], fmt_val)
                    else:
                        ws.cell(row=srow, column=ac).border = BD
            else:
                for ac in ann_hist:
                    ws.cell(row=srow, column=ac).border = BD
            # Forecast
            for fi, ac in enumerate(ann_fcst):
                if vals and fi < len(vals):
                    c = ws.cell(row=srow, column=ac, value=vals[fi])
                    c.number_format = fmt_val
                    c.font = Font(color="000000", size=10, name="Calibri")
                    c.fill = FILL_ASSUMP
                    c.border = BD
                    c.alignment = ALIGN_RIGHT
                else:
                    ws.cell(row=srow, column=ac).border = BD
            for sc in [17, 33, 34]:
                ws.cell(row=srow, column=sc).border = BD

        # Selected row
        ws.cell(row=sel_r, column=1, value="  SELECTED").font = Font(bold=True, size=10, name="Calibri", color="1F4E79")
        ws.cell(row=sel_r, column=1).border = BD
        for ci, ac in enumerate(ann_hist):
            if hist_vals and ci < len(hist_vals) and hist_vals[ci] is not None:
                _hval(ws, sel_r, ac, hist_vals[ci], fmt_val)
            else:
                ws.cell(row=sel_r, column=ac).border = BD
        for ac in ann_fcst:
            cl = CL(ac)
            formula = f'=CHOOSE(MATCH({SCENARIO_CELL},{{"Base","Bull","Bear"}},0),{cl}{base_r},{cl}{bull_r},{cl}{bear_r})'
            c = ws.cell(row=sel_r, column=ac, value=formula)
            c.number_format = fmt_val
            c.font = FONT_BOLD
            c.fill = FILL_KEY
            c.border = BD
            c.alignment = ALIGN_RIGHT
        for sc in [17, 33, 34]:
            ws.cell(row=sel_r, column=sc).border = BD
        _note(ws, sel_r, note_text)
        return sel_r

    # ============================================================
    # BOTTOM-UP GROWTH DRIVERS: CCG
    # ============================================================
    _section(ws, r, "CCG Bottom-Up Drivers: Revenue = PC TAM x Intel Share x ASP"); r += 1

    # PC TAM — sum-type (annual split equally across quarters)
    R["asp_pc_tam"] = _aq(
        "PC TAM (M units) - Source: Gartner/IDC",
        {2023: [60.45, 60.45, 60.45, 60.45], 2024: [61.35, 61.35, 61.35, 61.35], 2025: [64.65, 64.65, 64.65, 64.65]},
        {2026: [65.0, 65.0, 65.0, 65.0], 2027: [66.25, 66.25, 66.25, 66.25], 2028: [67.5, 67.5, 67.5, 67.5]},
        {2026: [67.5, 67.5, 67.5, 67.5], 2027: [70.0, 70.0, 70.0, 70.0], 2028: [72.5, 72.5, 72.5, 72.5]},
        {2026: [62.5, 62.5, 62.5, 62.5], 2027: [63.0, 63.0, 63.0, 63.0], 2028: [63.75, 63.75, 63.75, 63.75]},
        "#,##0.0",
        "Annual TAM split equally across 4 quarters (edit individual quarters as needed). Gartner: 2023=241.8M, 2024=245.4M, 2025~259M.",
        "sum"
    )

    # Intel PC Unit Share — avg-type (same % each quarter)
    R["asp_intel_share"] = _aq(
        "Intel PC Unit Share (%) - Source: Mercury Research",
        {2023: [0.710, 0.710, 0.710, 0.710], 2024: [0.715, 0.715, 0.715, 0.715], 2025: [0.720, 0.720, 0.720, 0.720]},
        {2026: [0.72, 0.72, 0.72, 0.72], 2027: [0.725, 0.725, 0.725, 0.725], 2028: [0.73, 0.73, 0.73, 0.73]},
        {2026: [0.74, 0.74, 0.74, 0.74], 2027: [0.75, 0.75, 0.75, 0.75], 2028: [0.76, 0.76, 0.76, 0.76]},
        {2026: [0.70, 0.70, 0.70, 0.70], 2027: [0.69, 0.69, 0.69, 0.69], 2028: [0.68, 0.68, 0.68, 0.68]},
        "0.0%",
        "Mercury est: ~72%. 18A products competitive; share recovery assumed. Edit quarterly for seasonal patterns.",
        "avg"
    )

    # CCG Blended ASP — avg-type
    R["asp_ccg_asp"] = _aq(
        "Intel CCG Blended ASP ($) - Derived from Rev/(TAMxShare)",
        {2023: [188, 188, 188, 188], 2024: [190, 190, 190, 190], 2025: [173, 173, 173, 173]},
        {2026: [175, 175, 175, 175], 2027: [182, 182, 182, 182], 2028: [184, 184, 184, 184]},
        {2026: [180, 180, 180, 180], 2027: [190, 190, 190, 190], 2028: [195, 195, 195, 195]},
        {2026: [170, 170, 170, 170], 2027: [172, 172, 172, 172], 2028: [174, 174, 174, 174]},
        "0.0",
        "FY2025 ASP depressed. Q1 2026 +16% YoY. Base: ASP recovery + premium mix. Edit quarterly for product launch timing.",
        "avg"
    )

    # ============================================================
    # DCAI Bottom-Up Drivers
    # ============================================================
    _section(ws, r, "DCAI Bottom-Up Drivers: CPU Rev = Server TAM x Share x ASP + AI/ASIC Rev"); r += 1

    R["asp_svr_tam"] = _aq(
        "Server TAM (M units) - Source: IDC",
        {2023: [2.875, 2.875, 2.875, 2.875], 2024: [3.0, 3.0, 3.0, 3.0], 2025: [3.375, 3.375, 3.375, 3.375]},
        {2026: [3.375, 3.375, 3.375, 3.375], 2027: [3.625, 3.625, 3.625, 3.625], 2028: [3.875, 3.875, 3.875, 3.875]},
        {2026: [3.625, 3.625, 3.625, 3.625], 2027: [4.0, 4.0, 4.0, 4.0], 2028: [4.375, 4.375, 4.375, 4.375]},
        {2026: [3.125, 3.125, 3.125, 3.125], 2027: [3.25, 3.25, 3.25, 3.25], 2028: [3.375, 3.375, 3.375, 3.375]},
        "#,##0.0",
        "IDC: 2023=11.5M, 2024=12.0M, 2025~13.5M. AI buildout driving growth. Equal quarterly allocation (edit for seasonality).",
        "sum"
    )

    R["asp_intel_dc_share"] = _aq(
        "Intel DC Server Unit Share (%) - Source: Mercury Research",
        {2023: [0.775, 0.775, 0.775, 0.775], 2024: [0.760, 0.760, 0.760, 0.760], 2025: [0.730, 0.730, 0.730, 0.730]},
        {2026: [0.70, 0.70, 0.70, 0.70], 2027: [0.70, 0.70, 0.70, 0.70], 2028: [0.70, 0.70, 0.70, 0.70]},
        {2026: [0.72, 0.72, 0.72, 0.72], 2027: [0.73, 0.73, 0.73, 0.73], 2028: [0.74, 0.74, 0.74, 0.74]},
        {2026: [0.67, 0.67, 0.67, 0.67], 2027: [0.65, 0.65, 0.65, 0.65], 2028: [0.63, 0.63, 0.63, 0.63]},
        "0.0%",
        "Mercury: Q4'24 ~72%. Granite Rapids improving; AMD Turin competitive. Base: share stabilizes.",
        "avg"
    )

    R["asp_dc_asp"] = _aq(
        "Intel Server Blended ASP ($) - Derived from CPU Rev/(TAMxShare)",
        {2023: [1772, 1772, 1772, 1772], 2024: [1712, 1712, 1712, 1712], 2025: [1632, 1632, 1632, 1632]},
        {2026: [1600, 1600, 1600, 1600], 2027: [1700, 1700, 1700, 1700], 2028: [1750, 1750, 1750, 1750]},
        {2026: [1750, 1750, 1750, 1750], 2027: [1850, 1850, 1850, 1850], 2028: [1950, 1950, 1950, 1950]},
        {2026: [1500, 1500, 1500, 1500], 2027: [1550, 1550, 1550, 1550], 2028: [1580, 1580, 1580, 1580]},
        "#,##0",
        "FY2025 ASP depressed. Q1 2026 +27% YoY. Base: premium mix recovery. Edit quarterly for product cycle effects.",
        "avg"
    )

    R["asp_ai_asic"] = _aq(
        "AI/ASIC & Other Revenue ($M) - Gaudi, ASICs, NICs, IPUs",
        {2023: [50, 50, 50, 50], 2024: [125, 125, 125, 125], 2025: [200, 200, 200, 200]},
        {2026: [700, 700, 700, 700], 2027: [1125, 1125, 1125, 1125], 2028: [1500, 1500, 1500, 1500]},
        {2026: [875, 875, 875, 875], 2027: [1500, 1500, 1500, 1500], 2028: [2125, 2125, 2125, 2125]},
        {2026: [500, 500, 500, 500], 2027: [750, 750, 750, 750], 2028: [950, 950, 950, 950]},
        "#,##0",
        "ASIC run rate 'north of $1B' Q1 2026. Gaudi modest. Edit quarterly for customer ramp timing.",
        "sum"
    )

    # ============================================================
    # Intel Foundry Bottom-Up Drivers
    # ============================================================
    _section(ws, r, "Intel Foundry Bottom-Up Drivers: Internal Rev = Total Chip Vol x Rev/Chip + External"); r += 1

    R["asp_foundry_rev_per_chip"] = _aq(
        "Internal Foundry Rev per Chip ($/unit) - Internal Foundry Rev / Total Intel Units",
        {2023: [102.2, 102.2, 102.2, 102.2], 2024: [93.5, 93.5, 93.5, 93.5], 2025: [90.1, 90.1, 90.1, 90.1]},
        {2026: [90, 90, 90, 90], 2027: [92, 92, 92, 92], 2028: [95, 95, 95, 95]},
        {2026: [98, 98, 98, 98], 2027: [105, 105, 105, 105], 2028: [110, 110, 110, 110]},
        {2026: [82, 82, 82, 82], 2027: [80, 80, 80, 80], 2028: [78, 78, 78, 78]},
        "0.0",
        "Declining trend stabilizes as 18A premium wafers ramp. Edit quarterly for node transition effects.",
        "avg"
    )

    R["asp_ext_foundry"] = _aq(
        "External Foundry Revenue ($M) - 3rd-party wafer customers",
        {2023: [12.5, 12.5, 12.5, 12.5], 2024: [15, 15, 15, 15], 2025: [37.5, 37.5, 37.5, 37.5]},
        {2026: [175, 175, 175, 175], 2027: [375, 375, 375, 375], 2028: [625, 625, 625, 625]},
        {2026: [250, 250, 250, 250], 2027: [750, 750, 750, 750], 2028: [1250, 1250, 1250, 1250]},
        {2026: [100, 100, 100, 100], 2027: [200, 200, 200, 200], 2028: [300, 300, 300, 300]},
        "#,##0",
        "~$150M FY2025. Growing as 18A/14A external PDK ramps. Edit quarterly for customer win timing.",
        "sum"
    )

    # ============================================================
    # All Other Bottom-Up Drivers
    # ============================================================
    _section(ws, r, "All Other Bottom-Up Drivers: Mobileye (LV Prod x Rev/Veh) + IMS/Other"); r += 1

    R["asp_lv_prod"] = _aq(
        "Global Light Vehicle Production (M units) - Source: IHS/S&P Global Mobility",
        {2023: [22.25, 22.25, 22.25, 22.25], 2024: [22.0, 22.0, 22.0, 22.0], 2025: [21.75, 21.75, 21.75, 21.75]},
        {2026: [22.0, 22.0, 22.0, 22.0], 2027: [22.25, 22.25, 22.25, 22.25], 2028: [22.5, 22.5, 22.5, 22.5]},
        {2026: [22.5, 22.5, 22.5, 22.5], 2027: [23.0, 23.0, 23.0, 23.0], 2028: [23.5, 23.5, 23.5, 23.5]},
        {2026: [21.25, 21.25, 21.25, 21.25], 2027: [21.5, 21.5, 21.5, 21.5], 2028: [21.75, 21.75, 21.75, 21.75]},
        "#,##0.0",
        "IHS/S&P Global Mobility. FY2025 ~87M. Modest recovery in Base case.",
        "sum"
    )

    R["asp_mbly_rev_per_veh"] = _aq(
        "Mobileye Rev per Vehicle ($/vehicle) - Derived from Mobileye Rev / LV Production",
        {2023: [23.4, 23.4, 23.4, 23.4], 2024: [18.9, 18.9, 18.9, 18.9], 2025: [21.8, 21.8, 21.8, 21.8]},
        {2026: [19, 19, 19, 19], 2027: [20, 20, 20, 20], 2028: [21, 21, 21, 21]},
        {2026: [21, 21, 21, 21], 2027: [23, 23, 23, 23], 2028: [25, 25, 25, 25]},
        {2026: [17, 17, 17, 17], 2027: [17.5, 17.5, 17.5, 17.5], 2028: [18, 18, 18, 18]},
        "0.0",
        "FY2025 Mobileye $1.9B / 87M = $21.8. Q1 2026 $628M. Base: conservatively ~$19.",
        "avg"
    )

    R["asp_ims_other"] = _aq(
        "IMS & Other Revenue ($M) - IMS nanofabrication tools + legacy businesses",
        {2023: [844.5, 844.5, 844.5, 844.5], 2024: [485.25, 485.25, 485.25, 485.25], 2025: [415.75, 415.75, 415.75, 415.75]},
        {2026: [450, 450, 450, 450], 2027: [500, 500, 500, 500], 2028: [550, 550, 550, 550]},
        {2026: [550, 550, 550, 550], 2027: [650, 650, 650, 650], 2028: [750, 750, 750, 750]},
        {2026: [375, 375, 375, 375], 2027: [400, 400, 400, 400], 2028: [425, 425, 425, 425]},
        "#,##0",
        "IMS multi-beam mask writers. FY2023-24 include Altera (deconsolidated Sep 2025).",
        "sum"
    )

    # ============================================================
    # SEGMENT OPERATING MARGINS — avg-type (same each quarter)
    # ============================================================
    _section(ws, r, "Segment Operating Margin Assumptions"); r += 1

    R["asp_ccg_opm"] = _aq(
        "CCG Operating Margin (%)",
        {2023: [0.314, 0.314, 0.314, 0.314], 2024: [0.348, 0.348, 0.348, 0.348], 2025: [0.289, 0.289, 0.289, 0.289]},
        {2026: [0.30, 0.30, 0.30, 0.30], 2027: [0.315, 0.315, 0.315, 0.315], 2028: [0.33, 0.33, 0.33, 0.33]},
        {2026: [0.33, 0.33, 0.33, 0.33], 2027: [0.36, 0.36, 0.36, 0.36], 2028: [0.38, 0.38, 0.38, 0.38]},
        {2026: [0.27, 0.27, 0.27, 0.27], 2027: [0.27, 0.27, 0.27, 0.27], 2028: [0.28, 0.28, 0.28, 0.28]},
        "0.0%",
        "Q1 32.6% moderated by 18A ramp costs. Edit quarterly for product launch margin impact.",
        "avg"
    )

    R["asp_dcai_opm"] = _aq(
        "DCAI Operating Margin (%)",
        {2023: [0.059, 0.059, 0.059, 0.059], 2024: [0.088, 0.088, 0.088, 0.088], 2025: [0.202, 0.202, 0.202, 0.202]},
        {2026: [0.28, 0.28, 0.28, 0.28], 2027: [0.28, 0.28, 0.28, 0.28], 2028: [0.27, 0.27, 0.27, 0.27]},
        {2026: [0.32, 0.32, 0.32, 0.32], 2027: [0.33, 0.33, 0.33, 0.33], 2028: [0.32, 0.32, 0.32, 0.32]},
        {2026: [0.22, 0.22, 0.22, 0.22], 2027: [0.20, 0.20, 0.20, 0.20], 2028: [0.18, 0.18, 0.18, 0.18]},
        "0.0%",
        "Sustains near Q1 30.5% level; ASIC supports margins.",
        "avg"
    )

    R["asp_foundry_opm"] = _aq(
        "Intel Foundry Operating Margin (%)",
        {2023: [-0.383, -0.383, -0.383, -0.383], 2024: [-0.767, -0.767, -0.767, -0.767], 2025: [-0.579, -0.579, -0.579, -0.579]},
        {2026: [-0.42, -0.42, -0.42, -0.42], 2027: [-0.32, -0.32, -0.32, -0.32], 2028: [-0.22, -0.22, -0.22, -0.22]},
        {2026: [-0.35, -0.35, -0.35, -0.35], 2027: [-0.20, -0.20, -0.20, -0.20], 2028: [-0.08, -0.08, -0.08, -0.08]},
        {2026: [-0.50, -0.50, -0.50, -0.50], 2027: [-0.45, -0.45, -0.45, -0.45], 2028: [-0.40, -0.40, -0.40, -0.40]},
        "0.0%",
        "18A yield improvement + volume absorption. Still unprofitable; multi-year journey.",
        "avg"
    )

    R["asp_allother_opm"] = _aq(
        "All Other Operating Margin (%)",
        {2023: [0.276, 0.276, 0.276, 0.276], 2024: [-0.016, -0.016, -0.016, -0.016], 2025: [0.074, 0.074, 0.074, 0.074]},
        {2026: [0.14, 0.14, 0.14, 0.14], 2027: [0.18, 0.18, 0.18, 0.18], 2028: [0.20, 0.20, 0.20, 0.20]},
        {2026: [0.20, 0.20, 0.20, 0.20], 2027: [0.25, 0.25, 0.25, 0.25], 2028: [0.28, 0.28, 0.28, 0.28]},
        {2026: [0.08, 0.08, 0.08, 0.08], 2027: [0.10, 0.10, 0.10, 0.10], 2028: [0.12, 0.12, 0.12, 0.12]},
        "0.0%",
        "Mobileye gradual recovery + IMS growth.",
        "avg"
    )

    # ============================================================
    # CONSOLIDATED COST STRUCTURE — avg-type
    # ============================================================
    _section(ws, r, "Consolidated Cost Structure (% of Revenue)"); r += 1

    R["asp_cogs"] = _aq(
        "COGS as % of Revenue (=> GM% = 1 - COGS%)",
        {2023: [0.600, 0.600, 0.600, 0.600], 2024: [0.673, 0.673, 0.673, 0.673], 2025: [0.652, 0.652, 0.652, 0.652]},
        {2026: [0.615, 0.615, 0.615, 0.615], 2027: [0.590, 0.590, 0.590, 0.590], 2028: [0.565, 0.565, 0.565, 0.565]},
        {2026: [0.590, 0.590, 0.590, 0.590], 2027: [0.550, 0.550, 0.550, 0.550], 2028: [0.510, 0.510, 0.510, 0.510]},
        {2026: [0.640, 0.640, 0.640, 0.640], 2027: [0.630, 0.630, 0.630, 0.630], 2028: [0.620, 0.620, 0.620, 0.620]},
        "0.0%",
        "Base GM: 38.5%/41.0%/43.5%. Q1 39.4% base; 18A mix headwind offset by yield gains.",
        "avg"
    )

    R["asp_rd"] = _aq(
        "R&D as % of Revenue",
        {2023: [0.296, 0.296, 0.296, 0.296], 2024: [0.312, 0.312, 0.312, 0.312], 2025: [0.261, 0.261, 0.261, 0.261]},
        {2026: [0.260, 0.260, 0.260, 0.260], 2027: [0.250, 0.250, 0.250, 0.250], 2028: [0.240, 0.240, 0.240, 0.240]},
        {2026: [0.250, 0.250, 0.250, 0.250], 2027: [0.230, 0.230, 0.230, 0.230], 2028: [0.210, 0.210, 0.210, 0.210]},
        {2026: [0.280, 0.280, 0.280, 0.280], 2027: [0.280, 0.280, 0.280, 0.280], 2028: [0.270, 0.270, 0.270, 0.270]},
        "0.0%",
        "~$14.5B absolute in FY2026E; declining as % of growing revenue.",
        "avg"
    )

    R["asp_mga"] = _aq(
        "MG&A as % of Revenue",
        {2023: [0.104, 0.104, 0.104, 0.104], 2024: [0.104, 0.104, 0.104, 0.104], 2025: [0.087, 0.087, 0.087, 0.087]},
        {2026: [0.080, 0.080, 0.080, 0.080], 2027: [0.075, 0.075, 0.075, 0.075], 2028: [0.070, 0.070, 0.070, 0.070]},
        {2026: [0.075, 0.075, 0.075, 0.075], 2027: [0.065, 0.065, 0.065, 0.065], 2028: [0.060, 0.060, 0.060, 0.060]},
        {2026: [0.090, 0.090, 0.090, 0.090], 2027: [0.090, 0.090, 0.090, 0.090], 2028: [0.085, 0.085, 0.085, 0.085]},
        "0.0%",
        "~$4.5B flat absolute in FY2026E; declining as % of revenue.",
        "avg"
    )

    # ============================================================
    # ANNUAL PARAMETERS — use annual columns only
    # ============================================================
    _section(ws, r, "Annual Parameters (Non-Quarterly)"); r += 1

    R["asp_da"] = _a_annual(
        "D&A ($M absolute)",
        [7847, 9951, 10757],
        [12500, 13200, 13800],
        [12000, 12500, 12800],
        [13000, 14000, 15000],
        "#,##0",
        "Q1 annualised ~$12.5B; CIP->in-service driving growth."
    )

    R["asp_sbc"] = _a_annual(
        "SBC ($M absolute)",
        [3229, 3410, 2434],
        [2500, 2500, 2500],
        [2500, 2500, 2500],
        [2500, 2500, 2500],
        "#,##0",
        "Q1 2026 annualised ~$2.5B. Scenario-independent."
    )

    R["asp_restruct"] = _a_annual(
        "Restructuring & Other Charges ($M)",
        [-62, 6970, 2191],
        [500, 500, 500],
        [500, 500, 500],
        [500, 500, 500],
        "#,##0",
        "Winding down; plans substantially complete."
    )

    # CapEx
    R["asp_capex"] = _a_annual(
        "Capital Expenditure ($M Gross)",
        [25750, 25122, 17672],
        [17700, 19500, 21000],
        [17500, 22000, 25000],
        [18000, 17000, 16000],
        "#,##0",
        "FY2026E flat YoY per mgmt guidance; tool-heavy."
    )

    # Working Capital
    R["asp_ar_days"] = _a_annual(
        "AR Days",
        [22.9, 23.9, 26.5],
        [25, 25, 25],
        [25, 25, 25],
        [25, 25, 25],
        "#,##0.0",
        "Historical range 22-27 days."
    )

    R["asp_ap_days"] = _a_annual(
        "AP Days",
        [96.3, 128.2, 104.6],
        [100, 100, 100],
        [100, 100, 100],
        [100, 100, 100],
        "#,##0.0",
        "Normalised from FY2024 peak (vendor terms)."
    )

    R["asp_inv_days"] = _a_annual(
        "Inventory Days",
        [124.9, 124.5, 123.0],
        [125, 125, 125],
        [125, 125, 125],
        [125, 125, 125],
        "#,##0.0",
        "Remarkably stable at ~124 days."
    )

    R["asp_accrued_comp"] = _a_annual(
        "Accrued Comp % of Revenue",
        [0.067, 0.063, 0.076],
        [0.07, 0.07, 0.07],
        [0.07, 0.07, 0.07],
        [0.07, 0.07, 0.07],
        "0.0%",
        "Midpoint of historical range."
    )

    R["asp_other_ca"] = _a_annual(
        "Other CA % of Revenue",
        [None, None, None],
        [0.15, 0.15, 0.15],
        [0.15, 0.15, 0.15],
        [0.15, 0.15, 0.15],
        "0.0%",
        "Includes ~$7.5B refundable tax credits."
    )

    R["asp_other_cl"] = _a_annual(
        "Other CL % of Revenue",
        [None, None, None],
        [0.25, 0.25, 0.25],
        [0.25, 0.25, 0.25],
        [0.25, 0.25, 0.25],
        "0.0%",
        "Stable relationship."
    )

    R["asp_tax_payable"] = _a_annual(
        "Income Tax Payable % of Revenue",
        [None, None, None],
        [0.02, 0.02, 0.02],
        [0.02, 0.02, 0.02],
        [0.02, 0.02, 0.02],
        "0.0%",
        "Normalised level."
    )

    # Financing
    R["asp_debt"] = _a_annual(
        "Total Debt ($M)",
        [49266, 50011, 46585],
        [49500, 47200, 46000],
        [47000, 43000, 39000],
        [52000, 54000, 55000],
        "#,##0",
        "FY2026E: +$6.5B Fab34 loan -$2.5B maturities."
    )

    R["asp_shares"] = _a_annual(
        "Basic Shares (M)",
        [4228, 4330, 4994],
        [5100, 5180, 5250],
        [5080, 5120, 5150],
        [5200, 5400, 5600],
        "#,##0",
        "Q1 5,023M + ESPP/RSU + Escrowed Share releases."
    )

    R["asp_dps"] = _a_annual(
        "DPS ($/share)",
        [0.74, 0.38, 0.00],
        [0.00, 0.00, 0.00],
        [0.00, 0.00, 0.00],
        [0.00, 0.00, 0.00],
        "0.00",
        "Dividends suspended; no reinstatement expected."
    )

    R["asp_net_eq_iss"] = _a_annual(
        "Net Equity Issuance ($M)",
        [None, None, None],
        [300, 300, 300],
        [300, 300, 300],
        [300, 300, 300],
        "#,##0",
        "ESPP proceeds net of RSU withholdings."
    )

    # Tax & Non-Operating
    R["asp_etr"] = _a_annual(
        "Effective Tax Rate",
        [-1.198, -0.716, 0.983],
        [0.12, 0.13, 0.13],
        [0.11, 0.12, 0.12],
        [0.15, 0.15, 0.15],
        "0.0%",
        "Anchored to Q2 2026 guided 11%. FY2026E/FY2027E/FY2028E."
    )

    R["asp_interest"] = _a_annual(
        "Interest & Other, Net ($M)",
        [629, 226, 3257],
        [-800, -600, -400],
        [-500, -300, -100],
        [-1200, -1200, -1000],
        "#,##0",
        "Net interest on ~$50B debt (~5% coupon)."
    )

    R["asp_nci"] = _a_annual(
        "Non-Controlling Interests ($M)",
        [293, 293, 293],
        [750, 1100, 1100],
        [750, 1100, 1100],
        [750, 1100, 1100],
        "#,##0",
        "Per mgmt: ~$250M/q Q2-Q4 FY2026E; Ireland SCIP NCI eliminated Apr 2026."
    )

    # Corporate
    R["asp_corp_unalloc"] = _a_annual(
        "Corporate Unallocated ($M)",
        [-5199, -11177, -5518],
        [-5000, -4500, -4000],
        [-5000, -4500, -4000],
        [-5000, -4500, -4000],
        "#,##0",
        "Declining restructuring + SBC reallocation."
    )

    R["asp_inter_elim"] = _a_annual(
        "Intersegment Eliminations ($M)",
        [-205, -161, 619],
        [-200, -200, -200],
        [-200, -200, -200],
        [-200, -200, -200],
        "#,##0",
        "Residual inter-segment elim."
    )

    # ---- Column widths ----
    ws.column_dimensions['A'].width = 44
    for c in range(2, GD_NCOL + 1):
        cl = CL(c)
        if c in [17, 33, 34]:
            ws.column_dimensions[cl].width = 2
        elif c in [6, 11, 16, 22, 27, 32]:
            ws.column_dimensions[cl].width = 15
        elif c == GD_NCOL:
            ws.column_dimensions[cl].width = 55
        else:
            ws.column_dimensions[cl].width = 12

    # Legend & Sources
    lr = _add_legend(ws, r + 2)
    _add_sources(ws, lr + 1)
    ws.sheet_properties.tabColor = "1F4E79"
    return r
