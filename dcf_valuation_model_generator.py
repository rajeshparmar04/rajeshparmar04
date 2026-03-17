"""Generate a professional DCF valuation model workbook in Excel.

This script creates `DCF_Valuation_Model.xlsx` with the following sheets:
1. Cover / Summary
2. Assumptions
3. Historical Financials
4. Projections
5. Free Cash Flow Calculation
6. WACC Calculation
7. Discounted Cash Flow Valuation
8. Sensitivity Analysis
9. Charts / Valuation Summary
"""

from datetime import date
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell


OUTPUT_FILE = "DCF_Valuation_Model.xlsx"
COMPANY_NAME = "Sample Company Inc."
ANALYST_NAME = "Analyst Name"
BASE_CURRENCY = "$"


def create_formats(workbook):
    """Create reusable workbook formats for professional styling."""
    return {
        "title": workbook.add_format(
            {
                "bold": True,
                "font_size": 16,
                "font_color": "#FFFFFF",
                "bg_color": "#1F4E78",
                "align": "left",
                "valign": "vcenter",
                "border": 1,
            }
        ),
        "section": workbook.add_format(
            {
                "bold": True,
                "font_size": 11,
                "font_color": "#FFFFFF",
                "bg_color": "#2F75B5",
                "border": 1,
                "align": "left",
                "valign": "vcenter",
            }
        ),
        "header": workbook.add_format(
            {
                "bold": True,
                "bg_color": "#D9E1F2",
                "border": 1,
                "align": "center",
                "valign": "vcenter",
            }
        ),
        "label": workbook.add_format({"bold": True, "border": 1, "bg_color": "#F2F2F2"}),
        "text": workbook.add_format({"border": 1}),
        "input": workbook.add_format(
            {
                "border": 1,
                "bg_color": "#FFF2CC",
                "font_color": "#7F6000",
                "num_format": "0.00%",
            }
        ),
        "input_number": workbook.add_format(
            {
                "border": 1,
                "bg_color": "#FFF2CC",
                "font_color": "#7F6000",
                "num_format": "0",
            }
        ),
        "currency": workbook.add_format({"border": 1, "num_format": f'{BASE_CURRENCY}#,##0'}),
        "currency_1dp": workbook.add_format({"border": 1, "num_format": f'{BASE_CURRENCY}#,##0.0'}),
        "percent": workbook.add_format({"border": 1, "num_format": "0.00%"}),
        "number": workbook.add_format({"border": 1, "num_format": "#,##0"}),
        "decimal": workbook.add_format({"border": 1, "num_format": "0.00"}),
        "total_label": workbook.add_format(
            {
                "bold": True,
                "border": 1,
                "top": 2,
                "bg_color": "#E2F0D9",
            }
        ),
        "total_currency": workbook.add_format(
            {
                "bold": True,
                "border": 1,
                "top": 2,
                "bg_color": "#E2F0D9",
                "num_format": f'{BASE_CURRENCY}#,##0',
            }
        ),
        "total_percent": workbook.add_format(
            {
                "bold": True,
                "border": 1,
                "top": 2,
                "bg_color": "#E2F0D9",
                "num_format": "0.00%",
            }
        ),
        "highlight": workbook.add_format(
            {
                "bold": True,
                "border": 1,
                "bg_color": "#FFD966",
                "num_format": f'{BASE_CURRENCY}#,##0.00',
            }
        ),
        "highlight_percent": workbook.add_format(
            {
                "bold": True,
                "border": 1,
                "bg_color": "#FFD966",
                "num_format": "0.00%",
            }
        ),
        "date": workbook.add_format({"border": 1, "num_format": "yyyy-mm-dd"}),
    }


def setup_sheet(worksheet, widths):
    """Apply common worksheet setup."""
    for column, width in widths.items():
        worksheet.set_column(column, column, width)
    worksheet.freeze_panes(2, 1)
    worksheet.hide_gridlines(2)


def write_assumptions_sheet(workbook, formats):
    ws = workbook.add_worksheet("Assumptions")
    setup_sheet(ws, {0: 42, 1: 20, 2: 48})

    ws.merge_range("A1:C1", "DCF ASSUMPTIONS", formats["title"])
    ws.write_row("A2", ["Input", "Value", "Notes"], formats["header"])

    assumptions = [
        ("Revenue Growth Rate", 0.08, "Annual top-line growth assumption for forecast period", "percent"),
        ("EBIT Margin", 0.18, "Operating profitability assumption", "percent"),
        ("Tax Rate", 0.25, "Effective corporate tax rate", "percent"),
        ("Capex % of Revenue", 0.04, "Capital expenditures as % of revenue", "percent"),
        ("Depreciation % of Revenue", 0.03, "Depreciation as % of revenue", "percent"),
        ("Change in Working Capital %", 0.01, "Incremental net working capital investment", "percent"),
        ("Risk Free Rate", 0.04, "10Y government bond proxy", "percent"),
        ("Market Risk Premium", 0.055, "Expected equity risk premium", "percent"),
        ("Beta", 1.10, "Levered beta estimate", "decimal"),
        ("Cost of Debt", 0.06, "Pre-tax borrowing cost", "percent"),
        ("Terminal Growth Rate", 0.025, "Perpetual growth rate", "percent"),
        ("Projection Years", 5, "Editable between 5 and 10 years", "number"),
        ("Net Debt", 1500, "Debt less cash (same units as financials)", "currency"),
        ("Diluted Shares Outstanding", 500, "Shares outstanding in millions", "number"),
        ("Current Market Price", 22.5, "Current traded share price", "currency_1dp"),
        ("Equity Weight (E/V)", 0.70, "Capital structure weight for equity", "percent"),
    ]

    for idx, (label, value, note, fmt_key) in enumerate(assumptions, start=2):
        ws.write(idx, 0, label, formats["label"])
        input_format = formats["input"] if fmt_key == "percent" else formats["input_number"]
        if fmt_key not in {"percent", "number"}:
            input_format = formats["text"]
        if fmt_key == "currency":
            input_format = workbook.add_format({"border": 1, "bg_color": "#FFF2CC", "num_format": f'{BASE_CURRENCY}#,##0'})
        elif fmt_key == "currency_1dp":
            input_format = workbook.add_format({"border": 1, "bg_color": "#FFF2CC", "num_format": f'{BASE_CURRENCY}#,##0.0'})
        elif fmt_key == "decimal":
            input_format = workbook.add_format({"border": 1, "bg_color": "#FFF2CC", "num_format": "0.00"})

        ws.write(idx, 1, value, input_format)
        ws.write(idx, 2, note, formats["text"])

    ws.data_validation("B13", {"validate": "integer", "criteria": "between", "minimum": 5, "maximum": 10})

    return ws


def write_historical_sheet(workbook, formats):
    ws = workbook.add_worksheet("Historical Financials")
    setup_sheet(ws, {0: 34, 1: 14, 2: 14, 3: 14, 4: 14, 5: 14, 6: 14})

    ws.merge_range("A1:G1", "HISTORICAL FINANCIALS (LAST 5 YEARS)", formats["title"])
    ws.write("A2", "Line Item", formats["header"])

    historical_years = [2020, 2021, 2022, 2023, 2024]
    for col, year in enumerate(historical_years, start=1):
        ws.write(1, col, year, formats["header"])

    line_items = [
        "Revenue",
        "EBIT",
        "EBITDA",
        "Depreciation",
        "Capex",
        "Working Capital",
        "Tax",
        "Free Cash Flow",
    ]

    start_row = 2
    for row_offset, item in enumerate(line_items):
        ws.write(start_row + row_offset, 0, item, formats["label"])

    base_revenue = [9000, 9600, 10350, 11000, 11800]
    for col, rev in enumerate(base_revenue, start=1):
        r = start_row
        ws.write(r, col, rev, formats["currency"])
        ws.write_formula(r + 1, col, f"={xl_rowcol_to_cell(r, col)}*Assumptions!$B$4", formats["currency"])
        ws.write_formula(r + 3, col, f"={xl_rowcol_to_cell(r, col)}*Assumptions!$B$6", formats["currency"])
        ws.write_formula(r + 2, col, f"={xl_rowcol_to_cell(r+1, col)}+{xl_rowcol_to_cell(r+3, col)}", formats["currency"])
        ws.write_formula(r + 4, col, f"={xl_rowcol_to_cell(r, col)}*Assumptions!$B$5", formats["currency"])
        ws.write_formula(r + 5, col, f"={xl_rowcol_to_cell(r, col)}*12%", formats["currency"])
        ws.write_formula(r + 6, col, f"={xl_rowcol_to_cell(r+1, col)}*Assumptions!$B$5", formats["currency"])
        ws.write_formula(
            r + 7,
            col,
            f"={xl_rowcol_to_cell(r+1, col)}-{xl_rowcol_to_cell(r+6, col)}+{xl_rowcol_to_cell(r+3, col)}-{xl_rowcol_to_cell(r+4, col)}",
            formats["currency"],
        )

    return ws


def write_projections_sheet(workbook, formats):
    ws = workbook.add_worksheet("Projections")
    setup_sheet(ws, {0: 34, 1: 14, 2: 14, 3: 14, 4: 14, 5: 14, 6: 14})

    ws.merge_range("A1:G1", "FINANCIAL PROJECTIONS", formats["title"])
    ws.write("A2", "Line Item", formats["header"])

    projection_years = [2025, 2026, 2027, 2028, 2029]
    for col, year in enumerate(projection_years, start=1):
        ws.write(1, col, year, formats["header"])

    rows = {
        "Revenue": 2,
        "EBIT": 3,
        "Taxes": 4,
        "NOPAT": 5,
        "Depreciation": 6,
        "Capex": 7,
        "Change in Working Capital": 8,
    }

    for name, row in rows.items():
        ws.write(row, 0, name, formats["label"])

    for col in range(1, 6):
        rev_cell = xl_rowcol_to_cell(rows["Revenue"], col)
        if col == 1:
            ws.write_formula(
                rows["Revenue"],
                col,
                "='Historical Financials'!F3*(1+Assumptions!$B$3)",
                formats["currency"],
            )
        else:
            prev_rev = xl_rowcol_to_cell(rows["Revenue"], col - 1)
            ws.write_formula(rows["Revenue"], col, f"={prev_rev}*(1+Assumptions!$B$3)", formats["currency"])

        ws.write_formula(rows["EBIT"], col, f"={rev_cell}*Assumptions!$B$4", formats["currency"])
        ws.write_formula(rows["Taxes"], col, f"={xl_rowcol_to_cell(rows['EBIT'], col)}*Assumptions!$B$5", formats["currency"])
        ws.write_formula(
            rows["NOPAT"],
            col,
            f"={xl_rowcol_to_cell(rows['EBIT'], col)}-{xl_rowcol_to_cell(rows['Taxes'], col)}",
            formats["currency"],
        )
        ws.write_formula(rows["Depreciation"], col, f"={rev_cell}*Assumptions!$B$6", formats["currency"])
        ws.write_formula(rows["Capex"], col, f"={rev_cell}*Assumptions!$B$5", formats["currency"])
        ws.write_formula(rows["Change in Working Capital"], col, f"={rev_cell}*Assumptions!$B$7", formats["currency"])

    return ws


def write_fcf_sheet(workbook, formats):
    ws = workbook.add_worksheet("Free Cash Flow Calculation")
    setup_sheet(ws, {0: 38, 1: 14, 2: 14, 3: 14, 4: 14, 5: 14})

    ws.merge_range("A1:F1", "FREE CASH FLOW BUILD-UP", formats["title"])
    ws.write("A2", "Line Item", formats["header"])
    for col, year in enumerate([2025, 2026, 2027, 2028, 2029], start=1):
        ws.write(1, col, year, formats["header"])

    line_rows = {
        "NOPAT": 2,
        "Depreciation": 3,
        "Capex": 4,
        "Change in Working Capital": 5,
        "Free Cash Flow": 6,
    }

    for name, row in line_rows.items():
        line_format = formats["total_label"] if name == "Free Cash Flow" else formats["label"]
        ws.write(row, 0, name, line_format)

    for col in range(1, 6):
        ws.write_formula(line_rows["NOPAT"], col, f"=Projections!{xl_rowcol_to_cell(5, col)}", formats["currency"])
        ws.write_formula(line_rows["Depreciation"], col, f"=Projections!{xl_rowcol_to_cell(6, col)}", formats["currency"])
        ws.write_formula(line_rows["Capex"], col, f"=Projections!{xl_rowcol_to_cell(7, col)}", formats["currency"])
        ws.write_formula(line_rows["Change in Working Capital"], col, f"=Projections!{xl_rowcol_to_cell(8, col)}", formats["currency"])
        ws.write_formula(
            line_rows["Free Cash Flow"],
            col,
            f"={xl_rowcol_to_cell(2, col)}+{xl_rowcol_to_cell(3, col)}-{xl_rowcol_to_cell(4, col)}-{xl_rowcol_to_cell(5, col)}",
            formats["total_currency"],
        )

    return ws


def write_wacc_sheet(workbook, formats):
    ws = workbook.add_worksheet("WACC Calculation")
    setup_sheet(ws, {0: 44, 1: 20, 2: 36})

    ws.merge_range("A1:C1", "WEIGHTED AVERAGE COST OF CAPITAL", formats["title"])
    ws.write_row("A2", ["Metric", "Value", "Formula / Source"], formats["header"])

    inputs = [
        ("Risk Free Rate", "=Assumptions!B8", "Input from assumptions", "percent"),
        ("Beta", "=Assumptions!B10", "Input from assumptions", "decimal"),
        ("Market Risk Premium", "=Assumptions!B9", "Input from assumptions", "percent"),
        ("Cost of Equity", "=B3+B4*B5", "CAPM = Rf + Beta * ERP", "percent"),
        ("Pre-tax Cost of Debt", "=Assumptions!B11", "Input from assumptions", "percent"),
        ("Tax Rate", "=Assumptions!B5", "Input from assumptions", "percent"),
        ("Equity Weight (E/V)", "=Assumptions!B17", "Input from assumptions", "percent"),
        ("Debt Weight (D/V)", "=1-B9", "1 - Equity Weight", "percent"),
        ("WACC", "=B9*B6+B10*B7*(1-B8)", "(E/V*Ke)+(D/V*Kd*(1-T))", "percent"),
    ]

    start_row = 2
    for i, (metric, formula, note, fmt_key) in enumerate(inputs):
        row = start_row + i
        ws.write(row, 0, metric, formats["label"])
        ws.write_formula(row, 1, formula, formats[fmt_key] if fmt_key in formats else formats["text"])
        ws.write(row, 2, note, formats["text"])

    ws.write("A12", "Blended WACC", formats["total_label"])
    ws.write_formula("B12", "=B11", formats["total_percent"])
    ws.write("C12", "Used in DCF discounting", formats["total_label"])

    return ws


def write_dcf_sheet(workbook, formats):
    ws = workbook.add_worksheet("Discounted Cash Flow Valuation")
    setup_sheet(ws, {0: 42, 1: 14, 2: 14, 3: 14, 4: 14, 5: 14, 6: 18})

    ws.merge_range("A1:G1", "DISCOUNTED CASH FLOW VALUATION", formats["title"])
    ws.write_row("A2", ["Line Item", "2025", "2026", "2027", "2028", "2029", "Notes"], formats["header"])

    ws.write("A3", "Free Cash Flow", formats["label"])
    ws.write("A4", "Discount Factor", formats["label"])
    ws.write("A5", "Present Value of FCF", formats["label"])

    for col in range(1, 6):
        ws.write_formula(2, col, f"='Free Cash Flow Calculation'!{xl_rowcol_to_cell(6, col)}", formats["currency"])
        ws.write_formula(col + 1, 0, "", formats["text"])
        ws.write_formula(3, col, f"=1/(1+'WACC Calculation'!$B$12)^{col}", formats["decimal"])
        ws.write_formula(4, col, f"={xl_rowcol_to_cell(2, col)}*{xl_rowcol_to_cell(3, col)}", formats["currency"])

    ws.write("G3", "Linked from FCF sheet", formats["text"])
    ws.write("G4", "1 / (1 + WACC)^t", formats["text"])
    ws.write("G5", "Discounted projected cash flows", formats["text"])

    ws.write("A7", "Terminal Value", formats["label"])
    ws.write_formula("B7", "=F3*(1+Assumptions!$B$12)/('WACC Calculation'!$B$12-Assumptions!$B$12)", formats["currency"])
    ws.write("G7", "Gordon Growth Method", formats["text"])

    ws.write("A8", "PV of Terminal Value", formats["label"])
    ws.write_formula("B8", "=B7*F4", formats["currency"])
    ws.write("G8", "Discounted to present value", formats["text"])

    ws.write("A10", "Sum of PV of Explicit FCF", formats["label"])
    ws.write_formula("B10", "=SUM(B5:F5)", formats["currency"])
    ws.write("A11", "Enterprise Value", formats["total_label"])
    ws.write_formula("B11", "=B10+B8", formats["total_currency"])

    ws.write("A12", "Less: Net Debt", formats["label"])
    ws.write_formula("B12", "=Assumptions!$B$14", formats["currency"])
    ws.write("A13", "Equity Value", formats["total_label"])
    ws.write_formula("B13", "=B11-B12", formats["total_currency"])

    ws.write("A14", "Diluted Shares Outstanding", formats["label"])
    ws.write_formula("B14", "=Assumptions!$B$15", formats["number"])
    ws.write("A15", "Intrinsic Share Price", formats["total_label"])
    ws.write_formula("B15", "=B13/B14", formats["highlight"])

    ws.write("A16", "Current Market Price", formats["label"])
    ws.write_formula("B16", "=Assumptions!$B$16", formats["currency"])
    ws.write("A17", "Upside / Downside", formats["total_label"])
    ws.write_formula("B17", "=B15/B16-1", formats["highlight_percent"])

    return ws


def write_sensitivity_sheet(workbook, formats):
    ws = workbook.add_worksheet("Sensitivity Analysis")
    setup_sheet(ws, {0: 20, 1: 14, 2: 14, 3: 14, 4: 14, 5: 14, 6: 14})

    ws.merge_range("A1:G1", "SENSITIVITY: IMPLIED SHARE PRICE (WACC vs TERMINAL GROWTH)", formats["title"])

    tg_values = [0.015, 0.020, 0.025, 0.030, 0.035]
    wacc_values = [0.075, 0.080, 0.085, 0.090, 0.095]

    ws.write("A3", "WACC \ g", formats["header"])
    for j, tg in enumerate(tg_values, start=1):
        ws.write(2, j, tg, formats["header"])

    for i, wacc in enumerate(wacc_values, start=3):
        ws.write(i, 0, wacc, formats["header"])
        for j, _ in enumerate(tg_values, start=1):
            tg_cell = xl_rowcol_to_cell(2, j)
            wacc_cell = xl_rowcol_to_cell(i, 0)
            formula = (
                f"=((('Free Cash Flow Calculation'!F7*(1+{tg_cell})/({wacc_cell}-{tg_cell}))"
                f"/(1+{wacc_cell})^5)+SUMPRODUCT('Free Cash Flow Calculation'!B7:F7,"
                f"1/(1+{wacc_cell})^{{1,2,3,4,5}})-Assumptions!$B$14)/Assumptions!$B$15"
            )
            ws.write_formula(i, j, f"={formula}", formats["currency_1dp"])

    ws.conditional_format("B4:F8", {"type": "3_color_scale"})
    ws.write("A10", "Base Case Share Price", formats["label"])
    ws.write_formula("B10", "='Discounted Cash Flow Valuation'!B15", formats["highlight"])

    return ws


def write_cover_sheet(workbook, formats):
    ws = workbook.add_worksheet("Cover / Summary")
    setup_sheet(ws, {0: 30, 1: 40, 2: 24})

    ws.merge_range("A1:C1", "DISCOUNTED CASH FLOW VALUATION MODEL", formats["title"])
    ws.write("A3", "Company Name", formats["label"])
    ws.write("B3", COMPANY_NAME, formats["text"])

    ws.write("A4", "Valuation Date", formats["label"])
    ws.write_datetime("B4", date.today(), formats["date"])

    ws.write("A5", "Analyst Name", formats["label"])
    ws.write("B5", ANALYST_NAME, formats["text"])

    ws.write("A7", "Intrinsic Share Price", formats["total_label"])
    ws.write_formula("B7", "='Discounted Cash Flow Valuation'!B15", formats["highlight"])

    ws.write("A8", "Current Market Price", formats["label"])
    ws.write_formula("B8", "=Assumptions!B16", formats["currency"])

    ws.write("A9", "Upside / Downside %", formats["total_label"])
    ws.write_formula("B9", "=B7/B8-1", formats["highlight_percent"])

    ws.write("A11", "Model Navigation", formats["section"])
    nav = [
        "Assumptions",
        "Historical Financials",
        "Projections",
        "Free Cash Flow Calculation",
        "WACC Calculation",
        "Discounted Cash Flow Valuation",
        "Sensitivity Analysis",
        "Charts / Valuation Summary",
    ]
    for i, item in enumerate(nav, start=12):
        ws.write(i, 0, item, formats["label"])

    return ws


def write_charts_sheet(workbook, formats):
    ws = workbook.add_worksheet("Charts / Valuation Summary")
    setup_sheet(ws, {0: 32, 1: 14, 2: 14, 3: 14, 4: 14, 5: 14, 6: 20})

    ws.merge_range("A1:G1", "VISUAL VALUATION DASHBOARD", formats["title"])

    ws.write_row("A3", ["Metric", "2025", "2026", "2027", "2028", "2029", "Commentary"], formats["header"])
    ws.write("A4", "Revenue", formats["label"])
    ws.write("A5", "Free Cash Flow", formats["label"])

    for col in range(1, 6):
        ws.write_formula(3, col, f"=Projections!{xl_rowcol_to_cell(2, col)}", formats["currency"])
        ws.write_formula(4, col, f"='Free Cash Flow Calculation'!{xl_rowcol_to_cell(6, col)}", formats["currency"])

    ws.write("A7", "Valuation Bridge", formats["section"])
    ws.write("A8", "Enterprise Value", formats["label"])
    ws.write_formula("B8", "='Discounted Cash Flow Valuation'!B11", formats["currency"])
    ws.write("A9", "Net Debt", formats["label"])
    ws.write_formula("B9", "='Discounted Cash Flow Valuation'!B12", formats["currency"])
    ws.write("A10", "Equity Value", formats["total_label"])
    ws.write_formula("B10", "='Discounted Cash Flow Valuation'!B13", formats["total_currency"])

    revenue_chart = workbook.add_chart({"type": "line"})
    revenue_chart.add_series(
        {
            "name": "Revenue Forecast",
            "categories": "='Charts / Valuation Summary'!$B$3:$F$3",
            "values": "='Charts / Valuation Summary'!$B$4:$F$4",
            "line": {"color": "#2F75B5", "width": 2.25},
        }
    )
    revenue_chart.set_title({"name": "Revenue Forecast"})
    revenue_chart.set_y_axis({"num_format": f'{BASE_CURRENCY}#,##0'})
    revenue_chart.set_legend({"none": True})

    fcf_chart = workbook.add_chart({"type": "column"})
    fcf_chart.add_series(
        {
            "name": "Free Cash Flow Forecast",
            "categories": "='Charts / Valuation Summary'!$B$3:$F$3",
            "values": "='Charts / Valuation Summary'!$B$5:$F$5",
            "fill": {"color": "#70AD47"},
            "border": {"color": "#548235"},
        }
    )
    fcf_chart.set_title({"name": "Free Cash Flow Forecast"})
    fcf_chart.set_y_axis({"num_format": f'{BASE_CURRENCY}#,##0'})
    fcf_chart.set_legend({"none": True})

    bridge_chart = workbook.add_chart({"type": "column"})
    bridge_chart.add_series(
        {
            "name": "Valuation Bridge",
            "categories": "='Charts / Valuation Summary'!$A$8:$A$10",
            "values": "='Charts / Valuation Summary'!$B$8:$B$10",
            "fill": {"color": "#4472C4"},
        }
    )
    bridge_chart.set_title({"name": "DCF Valuation Bridge"})
    bridge_chart.set_legend({"none": True})

    ws.insert_chart("A12", revenue_chart, {"x_scale": 1.15, "y_scale": 1.15})
    ws.insert_chart("D12", fcf_chart, {"x_scale": 1.1, "y_scale": 1.15})
    ws.insert_chart("A30", bridge_chart, {"x_scale": 1.1, "y_scale": 1.15})

    return ws


def build_dcf_model(output_path=OUTPUT_FILE):
    """Build the complete DCF model workbook."""
    workbook = xlsxwriter.Workbook(output_path)
    formats = create_formats(workbook)

    # Build sheets in requested order.
    write_cover_sheet(workbook, formats)
    write_assumptions_sheet(workbook, formats)
    write_historical_sheet(workbook, formats)
    write_projections_sheet(workbook, formats)
    write_fcf_sheet(workbook, formats)
    write_wacc_sheet(workbook, formats)
    write_dcf_sheet(workbook, formats)
    write_sensitivity_sheet(workbook, formats)
    write_charts_sheet(workbook, formats)

    workbook.close()


if __name__ == "__main__":
    build_dcf_model()
    print(f"Workbook created successfully: {OUTPUT_FILE}")
