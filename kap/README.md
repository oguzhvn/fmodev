# KAP reports

Consolidated "Financial Report" disclosures downloaded from KAP (both the English and the Turkish PDF attachments):

| Period | English file | Turkish file | KAP disclosure | Published |
|---|---|---|---|---|
| MGROS 2025Q4 | `MGROS_2025Q4_EN.pdf` | `MGROS_2025Q4.pdf` | https://www.kap.org.tr/tr/Bildirim/1566129 | 05.03.2026 |
| MGROS 2026Q2 | `MGROS_2026Q2_EN.pdf` | `MGROS_2026Q2.pdf` | https://www.kap.org.tr/tr/Bildirim/1648240 | 11.08.2026 |
| BIMAS 2025Q4 | `BIMAS_2025Q4_EN.pdf` | `BIMAS_2025Q4.pdf` | https://www.kap.org.tr/tr/Bildirim/1570150 | 11.03.2026 |
| BIMAS 2026Q2 | `BIMAS_2026Q2_EN.pdf` | `BIMAS_2026Q2.pdf` | https://www.kap.org.tr/tr/Bildirim/1651656 | 17.08.2026 |

`veri.json` holds the raw figures used in the Excel file (TRY thousand) together with their sources (file and page in the English report).
The figures are identical in the English and Turkish versions.

To rebuild the Excel file: `python3 scripts/build_excel.py`
