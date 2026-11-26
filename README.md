# Financial Management Project – Migros (MGROS) vs. BİM (BIMAS)

Group project for the Financial Management course (Istanbul Medipol University).

**Group members:** Muhammed Oğuzhan Sağmak, Seher Naz Yomluer, Buse Çalış, Elif Aldemir

| | |
|---|---|
| Main company | Migros Ticaret A.Ş. (MGROS) |
| Rival company | BİM Birleşik Mağazalar A.Ş. (BIMAS) |
| Sector | Consumer trade (food retail) |
| Periods | 2025 Q4 (31.12.2025, 12 months) and 2026 Q2 (30.06.2026, 6 months) |
| Data source | Consolidated financial reports published on KAP (Public Disclosure Platform) |

## Contents

| File / folder | Description |
|---|---|
| [`FM_Project_MGROS_BIMAS.xlsx`](FM_Project_MGROS_BIMAS.xlsx) | Shared workbook: raw financial statement data, liquidity and activity ratios, comparison tables and data checks |
| [`kap/`](kap/) | The four KAP financial reports (English and Turkish PDFs) and links to the original disclosures |
| [`DEVIR_NOTU.md`](DEVIR_NOTU.md) | Handover note for the group members working on the next parts (in Turkish) |
| [`teams_bildirim.txt`](teams_bildirim.txt) | Group and company notification sent to the instructor |
| [`scripts/build_excel.py`](scripts/build_excel.py) | Script that builds the workbook from `kap/veri.json` |

## Workbook structure

| Sheet | Content |
|---|---|
| `Raw_Data` | Balance sheet and income statement items for both companies and both periods (TRY thousand), with the source page for every figure |
| `Parameters` | Days in a year (365) and the annualization factor (2025Q4 × 1, 2026Q2 × 2) |
| `Liquidity` | Current ratio, acid-test ratio, cash ratio, net working capital |
| `Activity` | Inventory, receivables and payables turnover and days, cash conversion cycle, fixed asset turnover, total asset turnover |
| `Comparison` | 6-month change for each company and the MGROS − BIMAS difference for each ratio |
| `Checks` | Accounting identity checks (all OK) |

Every ratio shows its numerator, denominator and result, and all cells are live formulas linked to `Raw_Data`.

## Method notes

- Each period uses the current-period column of its own report. The 2026 Q2 income statement covers 1 January – 30 June 2026.
- Income statement figures for 2026 Q2 are annualized (× 2) when they are divided by balance sheet items.
- Period-end balances are used for balance sheet items.
- All reports apply TAS 29 (inflation accounting), so 2025 Q4 figures are stated in December 2025 purchasing power and 2026 Q2 figures in June 2026 purchasing power.
- All figures were cross-checked against both the English and the Turkish versions of the reports.

## Project status

| Part | Scope | Status |
|---|---|---|
| 1 | Data collection, liquidity and activity ratios | Completed |
| 2 | Leverage, profitability and market ratios | In progress |
| 3 | DuPont analysis and 6-month performance evaluation | In progress |
| 4 | Written interpretation, final report and presentation | In progress |

## Use of AI

AI was used to download the financial statements from KAP, organize the data in Excel, and build the liquidity and activity ratio calculation tables. All figures were checked against the original reports. The interpretations belong to the group members.
