"""Builds FM_Project_MGROS_BIMAS.xlsx.

Raw figures are read from kap/veri.json (values taken from the KAP financial reports).
If the file or a value is missing, the cell is left empty and highlighted in yellow.
"""
import json, os, re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_PATH = os.path.join(ROOT, "kap", "veri.json")
OUT = os.path.join(ROOT, "FM_Project_MGROS_BIMAS.xlsx")

COLS = ["MGROS 2025Q4", "MGROS 2026Q2", "BIMAS 2025Q4", "BIMAS 2026Q2"]
VC = ["B", "C", "D", "E"]          # Raw_Data value columns
SC = ["F", "G", "H", "I"]          # Raw_Data source columns

# (key, label)  key "H" = section heading
ITEMS = [
    ("H", "STATEMENT OF FINANCIAL POSITION (TRY thousand)"),
    ("nakit", "Cash and cash equivalents"),
    ("kvfy", "Short-term financial investments"),
    ("ticalacak", "Trade receivables (current)"),
    ("stok", "Inventories"),
    ("donen", "Total current assets"),
    ("mdv", "Property, plant and equipment"),
    ("duran", "Total non-current assets"),
    ("toplamvar", "Total assets"),
    ("ticborc", "Trade payables (current)"),
    ("kvyk", "Total current liabilities"),
    ("uvyk", "Total non-current liabilities"),
    ("toplamyuk", "Total liabilities"),
    ("ozkaynak", "Total equity"),
    ("anaozk", "Equity attributable to owners of the parent"),
    ("sermaye", "Paid-in capital (TRY thousand)"),
    ("paysayisi", "Number of shares (units)"),
    ("H", "STATEMENT OF PROFIT OR LOSS (TRY thousand; expenses shown as negative, as reported)"),
    ("hasilat", "Revenue"),
    ("smm", "Cost of sales (-)"),
    ("brut", "Gross profit"),
    ("efk", "Operating profit / (loss)"),
    ("fingid", "Finance expenses (-)"),
    ("vok", "Profit before tax"),
    ("netkar", "Net profit for the period"),
    ("ananetkar", "Net profit attributable to owners of the parent"),
    ("H", "PERIOD INFORMATION"),
    ("ay", "Income statement period length (months)"),
]

F = "Arial"
bold = Font(name=F, bold=True)
norm = Font(name=F)
blue = Font(name=F, color="0000FF")
green = Font(name=F, color="008000")
title = Font(name=F, bold=True, size=13)
ital = Font(name=F, italic=True, color="555555")
hdrfill = PatternFill("solid", fgColor="1F3864")
hdrfont = Font(name=F, bold=True, color="FFFFFF")
secfill = PatternFill("solid", fgColor="D9E1F2")
yellow = PatternFill("solid", fgColor="FFFF00")
resfill = PatternFill("solid", fgColor="E2EFDA")
thin = Side(style="thin", color="BFBFBF")
box = Border(top=thin, bottom=thin, left=thin, right=thin)
TL = '#,##0;(#,##0);"-"'
RAT = '0.00;(0.00);"-"'
DAY = '#,##0.0;(#,##0.0);"-"'
PCT = '0.0%;(0.0%);"-"'

data = {}
if os.path.exists(DATA_PATH):
    with open(DATA_PATH, encoding="utf-8") as f:
        data = json.load(f)

wb = Workbook()

# ---------------- Raw_Data ----------------
hv = wb.active
hv.title = "Raw_Data"
hv["A1"] = "Raw Data – Migros (MGROS) and BİM (BIMAS), consolidated financial reports published on KAP"
hv["A1"].font = title
hv["A2"] = ("All amounts are in TRY thousand, as presented in the reports (except the number of shares). "
            "The reports apply TAS 29 (inflation accounting). For each period, the current-period column of that "
            "period's own report is used. Blue = figure entered from KAP; yellow = not entered yet.")
hv["A2"].font = ital
hdr = ["Item"] + COLS + [f"Source – {c}" for c in COLS]
for i, h in enumerate(hdr):
    c = hv.cell(row=4, column=i + 1, value=h)
    c.font, c.fill, c.alignment, c.border = hdrfont, hdrfill, Alignment(horizontal="center", wrap_text=True), box
ROW = {}
_r = 5
for _k, _l in ITEMS:
    if _k != "H":
        ROW[_k] = _r
    _r += 1
r = 5
for key, label in ITEMS:
    if key == "H":
        hv.cell(row=r, column=1, value=label).font = bold
        for col in range(1, 10):
            hv.cell(row=r, column=col).fill = secfill
        r += 1
        continue
    hv.cell(row=r, column=1, value=label).font = norm
    for j, col in enumerate(COLS):
        cell = hv[f"{VC[j]}{r}"]
        src = hv[f"{SC[j]}{r}"]
        if key == "ay":
            cell.value = 12 if col.endswith("Q4") else 6
            src.value = ("1 Jan – 31 Dec 2025 (12 months)" if col.endswith("Q4")
                         else "1 Jan – 30 Jun 2026 (6 months)")
            cell.font = blue
        else:
            d = data.get(col, {}).get(key)
            if d is not None and d.get("formul"):
                fml = d["formul"]
                for k2 in sorted(ROW, key=len, reverse=True):
                    fml = fml.replace(k2, f"{VC[j]}{ROW[k2]}")
                cell.value = fml
                src.value = d.get("kaynak", "")
            elif d is not None and d.get("deger") is not None:
                cell.value = d["deger"]
                cell.font = blue
                src.value = d.get("kaynak", "")
            else:
                cell.fill = yellow
        cell.number_format = "0" if key == "ay" else TL
        cell.border = box
        src.font = Font(name=F, size=9)
        src.border = box
    r += 1
hv.column_dimensions["A"].width = 50
for c in VC:
    hv.column_dimensions[c].width = 20
for c in SC:
    hv.column_dimensions[c].width = 42
hv.freeze_panes = "B5"

def H(key, j):
    return f"Raw_Data!${VC[j]}${ROW[key]}"

# ---------------- Parameters ----------------
pr = wb.create_sheet("Parameters")
pr["A1"] = "Parameters"; pr["A1"].font = title
pr["A3"] = "Days in a year"; pr["A3"].font = bold
pr["B3"] = 365; pr["B3"].font = blue; pr["B3"].fill = yellow
pr["C3"] = "If the course notes use 360 days, change this cell. All day-based ratios are linked to it."
pr["C3"].font = ital
for i, h in enumerate(["Parameter", *COLS]):
    c = pr.cell(row=5, column=i + 1, value=h)
    c.font, c.fill, c.border = hdrfont, hdrfill, box
pr["A6"] = "Income statement period length (months)"
pr["A7"] = "Annualization factor = 12 / period length"
PC = ["B", "C", "D", "E"]
for j in range(4):
    pr[f"{PC[j]}6"] = f"={H('ay', j)}"; pr[f"{PC[j]}6"].font = green
    pr[f"{PC[j]}7"] = f"=12/{PC[j]}6"; pr[f"{PC[j]}7"].number_format = "0.00"
    for rr in (6, 7):
        pr[f"{PC[j]}{rr}"].border = box
pr.column_dimensions["A"].width = 46
for c in PC:
    pr.column_dimensions[c].width = 16
DAYS = "Parameters!$B$3"
def FACTOR(j):
    return f"Parameters!${PC[j]}$7"

# ---------------- shared block writer ----------------
OC = ["D", "E", "F", "G"]   # value columns on ratio sheets
TC = ["I", "J", "K", "L"]   # "Calculation" text columns on ratio sheets
RESULT = {}                 # ratio name -> (sheet, row, format)
DEC = {TL: 0, RAT: 2, DAY: 1, "0": 0, "0.00": 2}

def calc_text(formula, row_fmt, res_ref, res_fmt):
    """'=D6/D7' -> =FIXED(D6,0)&" / "&FIXED(D7,0)&" = "&FIXED(D8,2)  (FIXED uses the local separators)"""
    parts = re.split(r"([+\-*/])", formula.lstrip("="))
    out = []
    for t in parts:
        t = t.strip()
        if t and t in "+-*/":
            out.append('" ' + {"-": "−", "*": "×"}.get(t, t) + ' "')
        elif t:
            m = re.match(r"\$?([A-Z]+)\$?(\d+)$", t)
            d = DEC.get(row_fmt.get(int(m.group(2)), TL), 0) if m else 0
            out.append(f"FIXED({t},{d})")
    out.append('" = "')
    out.append(f"FIXED({res_ref},{DEC.get(res_fmt, 2)})")
    return "=" + "&".join(out)

def setup_ratio_sheet(ws, ttl, note):
    ws["A1"] = ttl; ws["A1"].font = title
    ws["A2"] = note; ws["A2"].font = ital
    heads = ["Ratio", "Formula", "Step", *COLS, "", *[f"Calculation – {c}" for c in COLS]]
    for i, h in enumerate(heads):
        if not h:
            continue
        c = ws.cell(row=4, column=i + 1, value=h)
        c.font, c.fill, c.border = hdrfont, hdrfill, box
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.row_dimensions[4].height = 30
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 50
    ws.column_dimensions["C"].width = 42
    for c in OC:
        ws.column_dimensions[c].width = 18
    ws.column_dimensions["H"].width = 2
    for c in TC:
        ws.column_dimensions[c].width = 40
    ws.freeze_panes = "D5"

def block(ws, start, name, formula, steps, sheetname):
    """steps: list of (label, fn(j, rowmap) -> formula, format, kind); kind = 'link' | 'calc' | 'result'"""
    ws.cell(row=start, column=1, value=name).font = bold
    ws.cell(row=start, column=2, value=formula).font = norm
    for col in [1, 2, 3, 4, 5, 6, 7, 9, 10, 11, 12]:
        ws.cell(row=start, column=col).fill = secfill
    rowmap, row_fmt = {}, {}
    r = start + 1
    for label, fn, fmt, kind in steps:
        ws.cell(row=r, column=3, value=label).font = bold if kind == "result" else norm
        rowmap[label] = r
        for j in range(4):
            c = ws[f"{OC[j]}{r}"]
            c.value = fn(j, rowmap)
            c.number_format = fmt
            c.border = box
            c.font = green if kind == "link" else (bold if kind == "result" else norm)
            if kind == "result":
                c.fill = resfill
                t = ws[f"{TC[j]}{r}"]
                t.value = calc_text(c.value, row_fmt, f"{OC[j]}{r}", fmt)
                t.font = Font(name=F, size=9); t.border = box
        row_fmt[r] = fmt
        if kind == "result":
            RESULT[name] = (sheetname, r, fmt)
        r += 1
    return r + 1, rowmap

def ref(rowmap, label, j):
    return f"{OC[j]}{rowmap[label]}"

# ---------------- Liquidity ----------------
lk = wb.create_sheet("Liquidity")
setup_ratio_sheet(lk, "Liquidity Ratios",
                  "All values are live formulas linked to the Raw_Data sheet. Period-end balances are used.")
r = 5
r, _ = block(lk, r, "Current Ratio", "Current assets / Current liabilities", [
    ("Numerator: Total current assets", lambda j, m: f"={H('donen', j)}", TL, "link"),
    ("Denominator: Total current liabilities", lambda j, m: f"={H('kvyk', j)}", TL, "link"),
    ("Result: Current ratio (times)", lambda j, m: f"={ref(m,'Numerator: Total current assets',j)}/{ref(m,'Denominator: Total current liabilities',j)}", RAT, "result"),
], "Liquidity")
r, _ = block(lk, r, "Acid-Test (Quick) Ratio", "(Current assets − Inventories) / Current liabilities", [
    ("Total current assets", lambda j, m: f"={H('donen', j)}", TL, "link"),
    ("Inventories", lambda j, m: f"={H('stok', j)}", TL, "link"),
    ("Numerator: Current assets − Inventories", lambda j, m: f"={ref(m,'Total current assets',j)}-{ref(m,'Inventories',j)}", TL, "calc"),
    ("Denominator: Total current liabilities", lambda j, m: f"={H('kvyk', j)}", TL, "link"),
    ("Result: Acid-test ratio (times)", lambda j, m: f"={ref(m,'Numerator: Current assets − Inventories',j)}/{ref(m,'Denominator: Total current liabilities',j)}", RAT, "result"),
], "Liquidity")
r, _ = block(lk, r, "Cash Ratio", "(Cash and cash equivalents + Short-term financial investments) / Current liabilities", [
    ("Cash and cash equivalents", lambda j, m: f"={H('nakit', j)}", TL, "link"),
    ("Short-term financial investments", lambda j, m: f"={H('kvfy', j)}", TL, "link"),
    ("Numerator: Cash + Short-term financial investments", lambda j, m: f"={ref(m,'Cash and cash equivalents',j)}+{ref(m,'Short-term financial investments',j)}", TL, "calc"),
    ("Denominator: Total current liabilities", lambda j, m: f"={H('kvyk', j)}", TL, "link"),
    ("Result: Cash ratio (times)", lambda j, m: f"={ref(m,'Numerator: Cash + Short-term financial investments',j)}/{ref(m,'Denominator: Total current liabilities',j)}", RAT, "result"),
], "Liquidity")
r, _ = block(lk, r, "Net Working Capital (TRY thousand)", "Current assets − Current liabilities", [
    ("Total current assets", lambda j, m: f"={H('donen', j)}", TL, "link"),
    ("Total current liabilities", lambda j, m: f"={H('kvyk', j)}", TL, "link"),
    ("Result: Net working capital (TRY thousand)", lambda j, m: f"={ref(m,'Total current assets',j)}-{ref(m,'Total current liabilities',j)}", TL, "result"),
], "Liquidity")

# ---------------- Activity ----------------
fa = wb.create_sheet("Activity")
setup_ratio_sheet(fa, "Activity Ratios",
                  "Income statement figures are multiplied by the annualization factor (see Parameters): 2025Q4 ×1, "
                  "2026Q2 ×2. Cost of sales is used as an absolute value. Balance sheet items are period-end balances.")
r = 5
r, Y = block(fa, r, "Annualization (intermediate step)", "Income statement amount × (12 / period length)", [
    ("Revenue (for the period, as reported)", lambda j, m: f"={H('hasilat', j)}", TL, "link"),
    ("Cost of sales (for the period, absolute value)", lambda j, m: f"=ABS({H('smm', j)})", TL, "link"),
    ("Annualization factor", lambda j, m: f"={FACTOR(j)}", "0.00", "link"),
    ("Days in a year", lambda j, m: f"={DAYS}", "0", "link"),
    ("Annualized revenue", lambda j, m: f"={ref(m,'Revenue (for the period, as reported)',j)}*{ref(m,'Annualization factor',j)}", TL, "calc"),
    ("Annualized cost of sales", lambda j, m: f"={ref(m,'Cost of sales (for the period, absolute value)',j)}*{ref(m,'Annualization factor',j)}", TL, "calc"),
], "Activity")
def AREV(j): return f"{OC[j]}${Y['Annualized revenue']}"
def ACOS(j): return f"{OC[j]}${Y['Annualized cost of sales']}"
def ND(j): return f"{OC[j]}${Y['Days in a year']}"

r, _ = block(fa, r, "Inventory Turnover", "Annualized cost of sales / Inventories", [
    ("Numerator: Annualized cost of sales", lambda j, m: f"={ACOS(j)}", TL, "calc"),
    ("Denominator: Inventories", lambda j, m: f"={H('stok', j)}", TL, "link"),
    ("Result: Inventory turnover (times)", lambda j, m: f"={ref(m,'Numerator: Annualized cost of sales',j)}/{ref(m,'Denominator: Inventories',j)}", RAT, "result"),
], "Activity")
inv_t = RESULT["Inventory Turnover"][1]
r, _ = block(fa, r, "Days Sales in Inventory (days)", "Days in a year / Inventory turnover", [
    ("Days in a year", lambda j, m: f"={ND(j)}", "0", "calc"),
    ("Inventory turnover", lambda j, m: f"={OC[j]}{inv_t}", RAT, "calc"),
    ("Result: Days sales in inventory (days)", lambda j, m: f"={ref(m,'Days in a year',j)}/{ref(m,'Inventory turnover',j)}", DAY, "result"),
], "Activity")
r, _ = block(fa, r, "Receivables Turnover", "Annualized revenue / Trade receivables", [
    ("Numerator: Annualized revenue", lambda j, m: f"={AREV(j)}", TL, "calc"),
    ("Denominator: Trade receivables (current)", lambda j, m: f"={H('ticalacak', j)}", TL, "link"),
    ("Result: Receivables turnover (times)", lambda j, m: f"={ref(m,'Numerator: Annualized revenue',j)}/{ref(m,'Denominator: Trade receivables (current)',j)}", RAT, "result"),
], "Activity")
rec_t = RESULT["Receivables Turnover"][1]
r, _ = block(fa, r, "Average Collection Period (days)", "Days in a year / Receivables turnover", [
    ("Days in a year", lambda j, m: f"={ND(j)}", "0", "calc"),
    ("Receivables turnover", lambda j, m: f"={OC[j]}{rec_t}", RAT, "calc"),
    ("Result: Average collection period (days)", lambda j, m: f"={ref(m,'Days in a year',j)}/{ref(m,'Receivables turnover',j)}", DAY, "result"),
], "Activity")
r, _ = block(fa, r, "Payables Turnover", "Annualized cost of sales / Trade payables", [
    ("Numerator: Annualized cost of sales", lambda j, m: f"={ACOS(j)}", TL, "calc"),
    ("Denominator: Trade payables (current)", lambda j, m: f"={H('ticborc', j)}", TL, "link"),
    ("Result: Payables turnover (times)", lambda j, m: f"={ref(m,'Numerator: Annualized cost of sales',j)}/{ref(m,'Denominator: Trade payables (current)',j)}", RAT, "result"),
], "Activity")
pay_t = RESULT["Payables Turnover"][1]
r, _ = block(fa, r, "Average Payment Period (days)", "Days in a year / Payables turnover", [
    ("Days in a year", lambda j, m: f"={ND(j)}", "0", "calc"),
    ("Payables turnover", lambda j, m: f"={OC[j]}{pay_t}", RAT, "calc"),
    ("Result: Average payment period (days)", lambda j, m: f"={ref(m,'Days in a year',j)}/{ref(m,'Payables turnover',j)}", DAY, "result"),
], "Activity")
dsi = RESULT["Days Sales in Inventory (days)"][1]
acp = RESULT["Average Collection Period (days)"][1]
app = RESULT["Average Payment Period (days)"][1]
r, _ = block(fa, r, "Cash Conversion Cycle (days)", "Days sales in inventory + Average collection period − Average payment period", [
    ("Days sales in inventory", lambda j, m: f"={OC[j]}{dsi}", DAY, "calc"),
    ("Average collection period", lambda j, m: f"={OC[j]}{acp}", DAY, "calc"),
    ("Average payment period", lambda j, m: f"={OC[j]}{app}", DAY, "calc"),
    ("Result: Cash conversion cycle (days)", lambda j, m: f"={ref(m,'Days sales in inventory',j)}+{ref(m,'Average collection period',j)}-{ref(m,'Average payment period',j)}", DAY, "result"),
], "Activity")
r, _ = block(fa, r, "Fixed Asset Turnover", "Annualized revenue / Property, plant and equipment", [
    ("Numerator: Annualized revenue", lambda j, m: f"={AREV(j)}", TL, "calc"),
    ("Denominator: Property, plant and equipment", lambda j, m: f"={H('mdv', j)}", TL, "link"),
    ("Result: Fixed asset turnover (times)", lambda j, m: f"={ref(m,'Numerator: Annualized revenue',j)}/{ref(m,'Denominator: Property, plant and equipment',j)}", RAT, "result"),
], "Activity")
r, _ = block(fa, r, "Total Asset Turnover", "Annualized revenue / Total assets", [
    ("Numerator: Annualized revenue", lambda j, m: f"={AREV(j)}", TL, "calc"),
    ("Denominator: Total assets", lambda j, m: f"={H('toplamvar', j)}", TL, "link"),
    ("Result: Total asset turnover (times)", lambda j, m: f"={ref(m,'Numerator: Annualized revenue',j)}/{ref(m,'Denominator: Total assets',j)}", RAT, "result"),
], "Activity")

# ---------------- Comparison ----------------
ks = wb.create_sheet("Comparison")
ks["A1"] = "Comparison – 6-month change and MGROS vs. BIMAS difference"; ks["A1"].font = title
ks["A2"] = ("Change = 2026Q2 − 2025Q4; % change = change / |2025Q4|. "
            "Difference = MGROS − BIMAS.")
ks["A2"].font = ital
heads = ["Ratio", "MGROS 2025Q4", "MGROS 2026Q2", "MGROS change", "MGROS % change",
         "BIMAS 2025Q4", "BIMAS 2026Q2", "BIMAS change", "BIMAS % change",
         "Difference 2025Q4 (MGROS − BIMAS)", "Difference 2026Q2 (MGROS − BIMAS)"]
for i, h in enumerate(heads):
    c = ks.cell(row=4, column=i + 1, value=h)
    c.font, c.fill, c.border = hdrfont, hdrfill, box
    c.alignment = Alignment(horizontal="center", wrap_text=True)
rr = 5
for name, (sh, row, fmt) in RESULT.items():
    ks.cell(row=rr, column=1, value=name).font = norm
    m0, m1, b0, b1 = [f"{sh}!{OC[j]}{row}" for j in range(4)]
    vals = {"B": f"={m0}", "C": f"={m1}", "D": f"=C{rr}-B{rr}", "E": f"=IF(B{rr}=0,\"n/a\",D{rr}/ABS(B{rr}))",
            "F": f"={b0}", "G": f"={b1}", "H": f"=G{rr}-F{rr}", "I": f"=IF(F{rr}=0,\"n/a\",H{rr}/ABS(F{rr}))",
            "J": f"=B{rr}-F{rr}", "K": f"=C{rr}-G{rr}"}
    for col, v in vals.items():
        c = ks[f"{col}{rr}"]
        c.value = v
        c.number_format = PCT if col in ("E", "I") else fmt
        c.font = green if col in ("B", "C", "F", "G") else norm
        c.border = box
    rr += 1
ks.column_dimensions["A"].width = 38
for i in range(2, 12):
    ks.column_dimensions[L(i)].width = 17
ks.row_dimensions[4].height = 45

# ---------------- Checks ----------------
kt = wb.create_sheet("Checks")
kt["A1"] = "Data checks (all differences should be 0 and the status should be OK)"; kt["A1"].font = title
for i, h in enumerate(["Check", *COLS]):
    c = kt.cell(row=3, column=i + 1, value=h)
    c.font, c.fill, c.border = hdrfont, hdrfill, box
checks = [
    ("Total assets − (Total liabilities + Total equity)", lambda j: f"={H('toplamvar',j)}-({H('toplamyuk',j)}+{H('ozkaynak',j)})"),
    ("(Current assets + Non-current assets) − Total assets", lambda j: f"=({H('donen',j)}+{H('duran',j)})-{H('toplamvar',j)}"),
    ("(Current liabilities + Non-current liabilities) − Total liabilities", lambda j: f"=({H('kvyk',j)}+{H('uvyk',j)})-{H('toplamyuk',j)}"),
    ("Gross profit − (Revenue − |Cost of sales|)", lambda j: f"={H('brut',j)}-({H('hasilat',j)}-ABS({H('smm',j)}))"),
]
for i, (lab, fn) in enumerate(checks):
    rr = 4 + i
    kt.cell(row=rr, column=1, value=lab)
    for j in range(4):
        c = kt[f"{PC[j]}{rr}"]; c.value = fn(j); c.number_format = TL; c.border = box
st = 4 + len(checks)
kt.cell(row=st, column=1, value="Status").font = bold
for j in range(4):
    c = kt[f"{PC[j]}{st}"]
    c.value = f'=IF(SUMPRODUCT(ABS({PC[j]}4:{PC[j]}{st-1}))<1,"OK","CHECK")'
    c.font = bold; c.border = box
kt["A" + str(st + 2)] = ("Note: a tolerance of TRY 1 thousand is allowed for rounding. If a difference appears, "
                         "re-check the Raw_Data entry against the report.")
kt["A" + str(st + 2)].font = ital
kt.column_dimensions["A"].width = 62
for c in PC:
    kt.column_dimensions[c].width = 18

for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if c.font is None or c.font.name != F:
                c.font = Font(name=F, bold=c.font.bold, italic=c.font.italic, color=c.font.color, size=c.font.size)

wb.save(OUT)
print("Saved:", OUT)
