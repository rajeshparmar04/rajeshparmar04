# TATA Steel Valuation Model — Workbook Notes

## Overview

`TATA_Steel_Valuation_Model.xlsx` is the companion Excel workbook for the
[TATA Steel Equity Valuation Report](TATA_Steel_Valuation_Report.md). It
contains every data table, assumption, and calculation referenced in the
report, laid out across dedicated worksheet tabs so that inputs can be
modified and outputs recalculated.

A ZIP download (`TATA_Steel_Valuation_Model.zip`) is provided as a fallback
for environments where GitHub cannot preview `.xlsx` files directly.

## Workbook Structure

| # | Sheet | Contents |
|---|---|---|
| 1 | **Cover** | Ticker, report metadata, CMP, recommendation |
| 2 | **Income Statement** | Consolidated P&L (FY2021–FY2025) |
| 3 | **Cash Flows** | OCF, Capex, FCF, FCF Margin (FY2021–FY2025) |
| 4 | **Balance Sheet** | Snapshot as of Mar 31, 2025 |
| 5 | **Key Ratios** | ROE, ROCE, Debt/Equity, Coverage, etc. |
| 6 | **Assumptions** | Revenue growth, EBITDA margin, Capex, tax rate, terminal growth; projected Revenue & EBITDA (FY2026E–FY2030E) |
| 7 | **DCF Model** | WACC build-up, FCFF projections, Terminal Value, PV table, Equity Value bridge |
| 8 | **Monte Carlo** | Simulation input distributions and output percentiles (10,000 iterations) |
| 9 | **EVA** | NOPAT, Invested Capital, Capital Charge, EVA forecast, Normalized EVA adjustment |
| 10 | **Relative Valuation** | Peer comparison table, EV/EBITDA scenario matrices (FY2026E and FY2028E) |
| 11 | **DDM** | Two-stage Dividend Discount Model |
| 12 | **Sensitivity** | WACC vs Terminal Growth matrix, EBITDA Margin sensitivity |
| 13 | **Football Field** | Valuation range summary across all methodologies, weighted average fair value |

## How to Use

1. **Review assumptions** — Open the *Assumptions* sheet and adjust growth
   rates, margins, or capex to reflect your own outlook.
2. **Trace the DCF** — The *DCF Model* sheet walks through WACC → FCFF →
   Terminal Value → Equity Value step by step.
3. **Compare methods** — The *Football Field* sheet consolidates bear / base /
   bull ranges from every valuation approach.
4. **Stress-test** — Use the *Sensitivity* sheet to see how the intrinsic
   value responds to changes in WACC and terminal growth.

## Disclaimer

> This workbook is prepared purely for academic purposes as part of a
> Final Year MBA project. It does not constitute investment advice. All data
> and projections are based on publicly available information and academic
> estimates. See the full disclaimer in the valuation report.

---

*Report Date: March 9, 2026 | Prepared by: Rajesh Parmar | For Academic Use Only*
