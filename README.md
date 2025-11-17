# trade-analysis
# Trade Analysis — Excel Workbook Generator

This repository contains:
- `scripts/generate_workbook.py` — Python script to generate an Excel workbook (`Trade_Analysis_Full_Workbook.xlsx`) with:
  - Raw Data sheet
  - Cleaned Data sheet (formula-driven)
  - Lookup Tables
  - Year/HSN/Model/Supplier summaries (SUMIFS/AVERAGEIFS)
  - Charts

- `src/utils/parseTradeData.ts` — TypeScript parser (supplementary only). **Do not** use for final assignment deliverables (assignment requires Excel-only formulas and PivotTables).

## How to use

1. Place your `Sample Data 2.xlsx` (or `Sheet1.csv`) in `sample_data/`.
2. Install Python dependencies:
