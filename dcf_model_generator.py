"""Generate a professional DCF valuation workbook in Excel.

Running this script creates:
    DCF_Valuation_Model.xlsx
"""

from datetime import date
import xlsxwriter

OUTPUT_FILE = "DCF_Valuation_Model.xlsx"


def set_column_widths(ws, widths):
    """Apply a list of (first_col, last_col, width) settings."""
    for first_col, last_col, width in widths:
        ws.set_column(first_col, last_col, width)


def write_section_title(ws, row, col, title, fmt):
    ws.write(row, col, title, fmt)


def build_formats(workbook):
    """Centralized format definitions for consistent professional styling."""
    return {
        "title": workbook.add_format({
            "bold": True,
            "font_size": 16,
            "font_color": "#1F4E78",
            "align": "left",
            "valign": "vcenter",
        }),
        "subtitle": workbook.add_format({
            "bold": True,
            "font_size": 11,
            "font_color": "#1F4E78",
            "bottom": 1,
            "bottom_color": "#1F4E78",
        }),
        "header": workbook.add_format({
            "bold": True,
            "bg_color": "#D9E1F2",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
        }),
        "label": workbook.add_format({
            "bold": True,
            "border": 1,
            "bg_color": "#F2F2F2",
        }),
        "cell": workbook.add_format({"border": 1}),
        "input": workbook.add_format({
            "border": 1,
            "bg_color": "#FFF2CC",
            "num_format": "0.00%",
        }),
        "input_num": workbook.add_format({
            "border": 1,
            "bg_color": "#FFF2CC",
            "num_format": "0.00",
        }),
        "input_int": workbook.add_format({
            "border": 1,
            "bg_color": "#FFF2CC",
            "num_format": "0",
        }),
        "pct": workbook.add_format({"border": 1, "num_format": "0.00%"}),
        "currency": workbook.add_format({"border": 1, "num_format": "$#,##0"}),
        "currency_2": workbook.add_format({"border": 1, "num_format": "$#,##0.00"}),
        "number": workbook.add_format({"border": 1, "num_format": "#,##0"}),
        "total_label": workbook.add_format({
            "bold": True,
            "border": 1,
            "bg_color": "#BDD7EE",
        }),
        "total_currency": workbook.add_format({
            "bold": True,
            "border": 1,
            "bg_color": "#BDD7EE",
            "num_format": "$#,##0",
        }),
        "total_currency_2": workbook.add_format({
            "bold": True,
            "border": 1,
            "bg_color": "#BDD7EE",
            "num_format": "$#,##0.00",
        }),
        "total_pct": workbook.add_format({
            "bold": True,
            "border": 1,
            "bg_color": "#BDD7EE",
            "num_format": "0.00%",
        }),
        "note": workbook.add_format({"italic": True, "font_color": "#666666"}),
    }


def create_assumptions_sheet(workbook, formats):
    ws = workbook.add_worksheet("Assumptions")
    ws.freeze_panes(4, 1)
    set_column_widths(ws, [(0, 0, 46), (1, 1, 22), (2, 2, 18)])

    write_section_title(ws, 0, 0, "Key Model Assumptions", formats["title"])
    ws.write(2, 0, "Assumption", formats["header"])
    ws.write(2, 1, "Value", formats["header"])
    ws.write(2, 2, "Notes", formats["header"])

    assumptions = [
        ("Revenue Growth Rate", 0.07, "Annual growth used for projection period", "pct"),
        ("EBIT Margin", 0.18, "EBIT as % of revenue", "pct"),
        ("Tax Rate", 0.25, "Effective cash tax rate", "pct"),
        ("Capex % of Revenue", 0.05, "Capital expenditures as % of revenue", "pct"),
        ("Depreciation % of Revenue", 0.03, "Depreciation as % of revenue", "pct"),
        ("Change in Working Capital %", 0.02, "Incremental working capital need", "pct"),
        ("Risk Free Rate", 0.04, "10Y sovereign yield proxy", "pct"),
        ("Market Risk Premium", 0.055, "Long-term equity risk premium", "pct"),
        ("Beta", 1.10, "Levered beta", "num"),
        ("Cost of Debt", 0.06, "Pre-tax borrowing rate", "pct"),
        ("Terminal Growth Rate", 0.025, "Perpetual growth in terminal period", "pct"),
        ("Projection Years", 5, "Supported range: 5-10 years", "int"),
        ("Current Market Price", 72.50, "Current share price for comparison", "currency"),
        ("Net Debt", 2500, "Debt minus cash", "currency"),
        ("Shares Outstanding", 1000, "Diluted shares outstanding (mm)", "number"),
        ("Equity Weight (E/V)", 0.70, "Capital structure weight", "pct"),
        ("Debt Weight (D/V)", 0.30, "Capital structure weight", "pct"),
    ]

    row = 3
    for name, value, note, typ in assumptions:
        ws.write(row, 0, name, formats["label"])
        if typ == "pct":
            ws.write(row, 1, value, formats["input"])
        elif typ == "num":
            ws.write(row, 1, value, formats["input_num"])
        elif typ == "int":
            ws.write(row, 1, value, formats["input_int"])
        elif typ == "currency":
            ws.write(row, 1, value, workbook.add_format({"border": 1, "bg_color": "#FFF2CC", "num_format": "$#,##0.00"}))
        else:
            ws.write(row, 1, value, workbook.add_format({"border": 1, "bg_color": "#FFF2CC", "num_format": "#,##0"}))
        ws.write(row, 2, note, formats["cell"])
        row += 1

    ws.write(row + 1, 0, "Yellow cells are user-editable inputs.", formats["note"])
    return ws


def create_historical_sheet(workbook, formats):
    ws = workbook.add_worksheet("Historical Financials")
    ws.freeze_panes(4, 1)
    set_column_widths(ws, [(0, 0, 34), (1, 5, 14)])

    write_section_title(ws, 0, 0, "Historical Financials (USD mm)", formats["title"])
    ws.write(2, 0, "Line Item", formats["header"])
    for c, year in enumerate(["Y-5", "Y-4", "Y-3", "Y-2", "Y-1"], start=1):
        ws.write(2, c, year, formats["header"])

    metrics = [
        ("Revenue", [10000, 10600, 11350, 11900, 12500]),
        ("EBIT", [1600, 1740, 1910, 2020, 2180]),
        ("EBITDA", [1900, 2070, 2270, 2390, 2570]),
        ("Depreciation", [300, 330, 360, 370, 390]),
        ("Capex", [450, 470, 510, 540, 560]),
        ("Working Capital", [900, 945, 1005, 1070, 1125]),
        ("Tax", [400, 435, 478, 505, 545]),
        ("Free Cash Flow", [1050, 1165, 1282, 1345, 1465]),
    ]

    for r, (metric, values) in enumerate(metrics, start=3):
        ws.write(r, 0, metric, formats["label"])
        for c, val in enumerate(values, start=1):
            ws.write(r, c, val, formats["currency"])
    return ws


def create_projections_sheet(workbook, formats):
    ws = workbook.add_worksheet("Projections")
    ws.freeze_panes(4, 1)
    set_column_widths(ws, [(0, 0, 34), (1, 10, 12)])

    write_section_title(ws, 0, 0, "Financial Projections (USD mm)", formats["title"])
    ws.write(2, 0, "Line Item", formats["header"])
    for c in range(1, 11):
        ws.write(2, c, f"Year {c}", formats["header"])

    line_items = [
        "Revenue",
        "EBIT",
        "Taxes",
        "NOPAT",
        "Depreciation",
        "Capex",
        "Change in Working Capital",
    ]

    for r, item in enumerate(line_items, start=3):
        ws.write(r, 0, item, formats["label"])

    # Revenue
    ws.write_formula(3, 1, "='Historical Financials'!F4*(1+Assumptions!B4)", formats["currency"])
    for c in range(2, 11):
        ws.write_formula(3, c, f"={xlsxwriter.utility.xl_rowcol_to_cell(3, c-1)}*(1+Assumptions!B4)", formats["currency"])

    # EBIT, Taxes, NOPAT, D&A, Capex, ΔNWC
    for c in range(1, 11):
        cell_rev = xlsxwriter.utility.xl_rowcol_to_cell(3, c)
        ws.write_formula(4, c, f"={cell_rev}*Assumptions!B5", formats["currency"])
        ws.write_formula(5, c, f"={xlsxwriter.utility.xl_rowcol_to_cell(4, c)}*Assumptions!B6", formats["currency"])
        ws.write_formula(6, c, f"={xlsxwriter.utility.xl_rowcol_to_cell(4, c)}-{xlsxwriter.utility.xl_rowcol_to_cell(5, c)}", formats["currency"])
        ws.write_formula(7, c, f"={cell_rev}*Assumptions!B8", formats["currency"])
        ws.write_formula(8, c, f"={cell_rev}*Assumptions!B7", formats["currency"])
        if c == 1:
            ws.write_formula(9, c, f"=({cell_rev}-'Historical Financials'!F4)*Assumptions!B9", formats["currency"])
        else:
            prev_rev = xlsxwriter.utility.xl_rowcol_to_cell(3, c-1)
            ws.write_formula(9, c, f"=({cell_rev}-{prev_rev})*Assumptions!B9", formats["currency"])

    ws.write(11, 0, "Note: Model displays 10 projection years; valuation uses first N years per Assumptions!B15.", formats["note"])
    return ws


def create_fcf_sheet(workbook, formats):
    ws = workbook.add_worksheet("Free Cash Flow Calculation")
    ws.freeze_panes(4, 1)
    set_column_widths(ws, [(0, 0, 36), (1, 10, 12)])

    write_section_title(ws, 0, 0, "Free Cash Flow Build (USD mm)", formats["title"])
    ws.write(2, 0, "Line Item", formats["header"])
    for c in range(1, 11):
        ws.write(2, c, f"Year {c}", formats["header"])

    line_items = ["NOPAT", "Depreciation", "Capex", "Change in Working Capital", "Free Cash Flow"]
    for r, item in enumerate(line_items, start=3):
        ws.write(r, 0, item, formats["label"])

    for c in range(1, 11):
        ws.write_formula(3, c, f"='Projections'!{xlsxwriter.utility.xl_col_to_name(c)}7", formats["currency"])
        ws.write_formula(4, c, f"='Projections'!{xlsxwriter.utility.xl_col_to_name(c)}8", formats["currency"])
        ws.write_formula(5, c, f"='Projections'!{xlsxwriter.utility.xl_col_to_name(c)}9", formats["currency"])
        ws.write_formula(6, c, f"='Projections'!{xlsxwriter.utility.xl_col_to_name(c)}10", formats["currency"])
        ws.write_formula(7, c, f"={xlsxwriter.utility.xl_rowcol_to_cell(3, c)}+{xlsxwriter.utility.xl_rowcol_to_cell(4, c)}-{xlsxwriter.utility.xl_rowcol_to_cell(5, c)}-{xlsxwriter.utility.xl_rowcol_to_cell(6, c)}", formats["total_currency"])
    return ws


def create_wacc_sheet(workbook, formats):
    ws = workbook.add_worksheet("WACC Calculation")
    ws.freeze_panes(4, 1)
    set_column_widths(ws, [(0, 0, 44), (1, 1, 22), (2, 2, 18)])

    write_section_title(ws, 0, 0, "Weighted Average Cost of Capital", formats["title"])
    ws.write_row(2, 0, ["Metric", "Value", "Formula"], formats["header"])

    metrics = [
        ("Risk Free Rate", "=Assumptions!B10", "Input from assumptions", "pct"),
        ("Beta", "=Assumptions!B12", "Input from assumptions", "num"),
        ("Market Risk Premium", "=Assumptions!B11", "Input from assumptions", "pct"),
        ("Cost of Equity (CAPM)", "=B4+B5*B6", "Rf + Beta × MRP", "total_pct"),
        ("Cost of Debt", "=Assumptions!B13", "Pre-tax cost of debt", "pct"),
        ("Tax Rate", "=Assumptions!B6", "Effective tax rate", "pct"),
        ("Equity Weight (E/V)", "=Assumptions!B19", "Capital structure weight", "pct"),
        ("Debt Weight (D/V)", "=Assumptions!B20", "Capital structure weight", "pct"),
        ("WACC", "=B10*B7+B11*B8*(1-B9)", "(E/V×Ke)+(D/V×Kd×(1-T))", "total_pct"),
    ]

    for i, (name, formula, note, fmt_key) in enumerate(metrics, start=3):
        ws.write(i, 0, name, formats["label"] if "total" not in fmt_key else formats["total_label"])
        ws.write_formula(i, 1, formula, formats[fmt_key] if fmt_key in formats else formats["number"])
        ws.write(i, 2, note, formats["cell"])
    return ws


def create_dcf_sheet(workbook, formats):
    ws = workbook.add_worksheet("Discounted Cash Flow Valuation")
    ws.freeze_panes(4, 1)
    set_column_widths(ws, [(0, 0, 46), (1, 10, 12)])

    write_section_title(ws, 0, 0, "DCF Valuation (USD mm, except per share)", formats["title"])
    ws.write_row(2, 0, ["Line Item"] + [f"Year {i}" for i in range(1, 11)], formats["header"])

    rows = {
        "fcf": 3,
        "discount_factor": 4,
        "pv_fcf": 5,
        "terminal_fcf": 7,
        "terminal_value": 8,
        "pv_terminal": 9,
    }

    ws.write(rows["fcf"], 0, "Free Cash Flow", formats["label"])
    ws.write(rows["discount_factor"], 0, "Discount Factor", formats["label"])
    ws.write(rows["pv_fcf"], 0, "Present Value of FCF", formats["label"])

    for c in range(1, 11):
        ws.write_formula(rows["fcf"], c, f"='Free Cash Flow Calculation'!{xlsxwriter.utility.xl_col_to_name(c)}8", formats["currency"])
        ws.write_formula(rows["discount_factor"], c, f"=1/(1+'WACC Calculation'!B12)^{c}", formats["number"])
        ws.write_formula(rows["pv_fcf"], c, f"={xlsxwriter.utility.xl_rowcol_to_cell(rows['fcf'], c)}*{xlsxwriter.utility.xl_rowcol_to_cell(rows['discount_factor'], c)}", formats["currency"])

    ws.write(rows["terminal_fcf"], 0, "Terminal Year FCF", formats["label"])
    ws.write(rows["terminal_value"], 0, "Terminal Value", formats["label"])
    ws.write(rows["pv_terminal"], 0, "Present Value of Terminal Value", formats["label"])

    ws.write_formula(rows["terminal_fcf"], 1, "=INDEX(B4:K4,Assumptions!B15)", formats["currency"])
    ws.write_formula(rows["terminal_value"], 1, "=B8*(1+Assumptions!B14)/('WACC Calculation'!B12-Assumptions!B14)", formats["total_currency"])
    ws.write_formula(rows["pv_terminal"], 1, "=B9/(1+'WACC Calculation'!B12)^Assumptions!B15", formats["total_currency"])

    summary_start = 12
    summary_items = [
        ("PV of Projected FCFs", "=SUM(B6:INDEX(B6:K6,Assumptions!B15))", "total_currency"),
        ("PV of Terminal Value", "=B10", "total_currency"),
        ("Enterprise Value", "=B13+B14", "total_currency"),
        ("Less: Net Debt", "=Assumptions!B17", "currency"),
        ("Equity Value", "=B15-B16", "total_currency"),
        ("Shares Outstanding (mm)", "=Assumptions!B18", "number"),
        ("Intrinsic Share Price", "=B17/B18", "total_currency_2"),
    ]

    for i, (label, formula, fmt) in enumerate(summary_items):
        row = summary_start + i
        label_fmt = formats["total_label"] if "total" in fmt or label in {"Enterprise Value", "Equity Value", "Intrinsic Share Price"} else formats["label"]
        ws.write(row, 0, label, label_fmt)
        ws.write_formula(row, 1, formula, formats[fmt])

    return ws


def create_sensitivity_sheet(workbook, formats):
    ws = workbook.add_worksheet("Sensitivity Analysis")
    ws.freeze_panes(5, 2)
    set_column_widths(ws, [(0, 0, 26), (1, 7, 14)])

    write_section_title(ws, 0, 0, "Sensitivity: WACC vs Terminal Growth Rate", formats["title"])
    ws.write(2, 0, "Base Intrinsic Share Price", formats["label"])
    ws.write_formula(2, 1, "='Discounted Cash Flow Valuation'!B19", formats["total_currency_2"])

    g_rates = [0.015, 0.020, 0.025, 0.030, 0.035, 0.040]
    waccs = [0.070, 0.075, 0.080, 0.085, 0.090, 0.095]

    ws.write(4, 0, "WACC \\ g", formats["header"])
    for idx, g in enumerate(g_rates, start=1):
        ws.write(4, idx, g, formats["header"])

    for r_idx, w in enumerate(waccs, start=5):
        ws.write(r_idx, 0, w, formats["header"])
        for c_idx, g in enumerate(g_rates, start=1):
            col_name = xlsxwriter.utility.xl_col_to_name(c_idx + 1)
            # Price = (PV of explicit FCFs + PV of TV - net debt) / shares
            formula = (
                f"=((SUM('Discounted Cash Flow Valuation'!B6:INDEX('Discounted Cash Flow Valuation'!B6:K6,Assumptions!B15))"
                f"+((INDEX('Discounted Cash Flow Valuation'!B4:K4,Assumptions!B15)*(1+{col_name}$5)/($A{r_idx}-{col_name}$5))"
                f"/(1+$A{r_idx})^Assumptions!B15)-Assumptions!B17)/Assumptions!B18)"
            )
            ws.write_formula(r_idx, c_idx, formula, formats["currency_2"])

    ws.conditional_format(5, 1, 10, 6, {"type": "3_color_scale"})
    return ws


def create_cover_sheet(workbook, formats):
    ws = workbook.add_worksheet("Cover / Summary")
    ws.freeze_panes(8, 0)
    set_column_widths(ws, [(0, 0, 38), (1, 1, 28), (2, 2, 20)])

    write_section_title(ws, 0, 0, "Discounted Cash Flow Valuation Model", formats["title"])
    ws.write(2, 0, "Company Name", formats["label"])
    ws.write(2, 1, "Sample Company Inc.", formats["cell"])
    ws.write(3, 0, "Valuation Date", formats["label"])
    ws.write(3, 1, str(date.today()), formats["cell"])
    ws.write(4, 0, "Analyst Name", formats["label"])
    ws.write(4, 1, "Automated Python Model", formats["cell"])

    ws.write(6, 0, "Valuation Snapshot", formats["subtitle"])
    ws.write(7, 0, "Intrinsic Share Price", formats["total_label"])
    ws.write_formula(7, 1, "='Discounted Cash Flow Valuation'!B19", formats["total_currency_2"])
    ws.write(8, 0, "Current Market Price", formats["label"])
    ws.write_formula(8, 1, "=Assumptions!B16", formats["currency_2"])
    ws.write(9, 0, "Upside / Downside %", formats["total_label"])
    ws.write_formula(9, 1, "=B8/B9-1", formats["total_pct"])

    ws.write(11, 0, "Navigation", formats["subtitle"])
    links = [
        "Assumptions",
        "Historical Financials",
        "Projections",
        "Free Cash Flow Calculation",
        "WACC Calculation",
        "Discounted Cash Flow Valuation",
        "Sensitivity Analysis",
        "Charts / Valuation Summary",
    ]
    for i, name in enumerate(links, start=12):
        ws.write_url(i, 0, f"internal:'{name}'!A1", string=f"Go to {name}")

    return ws


def create_charts_sheet(workbook, formats):
    ws = workbook.add_worksheet("Charts / Valuation Summary")
    set_column_widths(ws, [(0, 8, 18)])
    write_section_title(ws, 0, 0, "Charts & Valuation Summary", formats["title"])

    # Revenue forecast chart
    rev_chart = workbook.add_chart({"type": "line"})
    rev_chart.add_series({
        "name": "Revenue Forecast",
        "categories": "=Projections!$B$3:$K$3",
        "values": "=Projections!$B$4:$K$4",
        "line": {"color": "#1F77B4", "width": 2.25},
    })
    rev_chart.set_title({"name": "Revenue Forecast"})
    rev_chart.set_y_axis({"num_format": "$#,##0"})
    rev_chart.set_legend({"none": True})

    # FCF forecast chart
    fcf_chart = workbook.add_chart({"type": "column"})
    fcf_chart.add_series({
        "name": "FCF Forecast",
        "categories": "='Free Cash Flow Calculation'!$B$3:$K$3",
        "values": "='Free Cash Flow Calculation'!$B$8:$K$8",
        "fill": {"color": "#2CA02C"},
        "border": {"color": "#1C7C1C"},
    })
    fcf_chart.set_title({"name": "Free Cash Flow Forecast"})
    fcf_chart.set_y_axis({"num_format": "$#,##0"})
    fcf_chart.set_legend({"none": True})

    # Valuation bridge chart data
    ws.write_row(2, 0, ["Bridge Item", "Value"], formats["header"])
    bridge_data = [
        ("PV of Projected FCFs", "='Discounted Cash Flow Valuation'!B13"),
        ("PV of Terminal Value", "='Discounted Cash Flow Valuation'!B14"),
        ("Less: Net Debt", "=-'Discounted Cash Flow Valuation'!B16"),
        ("Equity Value", "='Discounted Cash Flow Valuation'!B17"),
    ]
    for i, (label, formula) in enumerate(bridge_data, start=3):
        ws.write(i, 0, label, formats["label"])
        ws.write_formula(i, 1, formula, formats["currency"])

    bridge_chart = workbook.add_chart({"type": "column"})
    bridge_chart.add_series({
        "name": "DCF Valuation Bridge",
        "categories": "='Charts / Valuation Summary'!$A$4:$A$7",
        "values": "='Charts / Valuation Summary'!$B$4:$B$7",
        "fill": {"color": "#9467BD"},
    })
    bridge_chart.set_title({"name": "DCF Valuation Bridge"})
    bridge_chart.set_legend({"none": True})
    bridge_chart.set_y_axis({"num_format": "$#,##0"})

    ws.insert_chart("D2", rev_chart, {"x_scale": 1.1, "y_scale": 1.2})
    ws.insert_chart("D20", fcf_chart, {"x_scale": 1.1, "y_scale": 1.2})
    ws.insert_chart("D38", bridge_chart, {"x_scale": 1.1, "y_scale": 1.2})

    ws.write(56, 0, "All charts are linked to dynamic model outputs.", formats["note"])
    return ws


def generate_dcf_workbook(filename=OUTPUT_FILE):
    workbook = xlsxwriter.Workbook(filename)
    formats = build_formats(workbook)

    create_cover_sheet(workbook, formats)
    create_assumptions_sheet(workbook, formats)
    create_historical_sheet(workbook, formats)
    create_projections_sheet(workbook, formats)
    create_fcf_sheet(workbook, formats)
    create_wacc_sheet(workbook, formats)
    create_dcf_sheet(workbook, formats)
    create_sensitivity_sheet(workbook, formats)
    create_charts_sheet(workbook, formats)

    workbook.close()


if __name__ == "__main__":
    generate_dcf_workbook()
    print(f"Workbook generated: {OUTPUT_FILE}")
