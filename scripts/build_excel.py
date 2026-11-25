"""FM_Proje_MGROS_BIMAS.xlsx oluşturur.
Ham rakamlar kap/veri.json dosyasından okunur (KAP raporlarından elle/PDF'ten alınan değerler).
Dosya yoksa ya da bir değer boşsa hücre boş (sarı) bırakılır.
"""
import json, os, sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_PATH = os.path.join(ROOT, "kap", "veri.json")
OUT = os.path.join(ROOT, "FM_Proje_MGROS_BIMAS.xlsx")

COLS = ["MGROS 2025Q4", "MGROS 2026Q2", "BIMAS 2025Q4", "BIMAS 2026Q2"]
VC = ["B", "C", "D", "E"]          # Ham_Veri değer sütunları
SC = ["F", "G", "H", "I"]          # Ham_Veri kaynak sütunları

# (anahtar, etiket, bölüm)
ITEMS = [
    ("H", "BİLANÇO (TL)", None),
    ("nakit", "Nakit ve nakit benzerleri", "b"),
    ("kvfy", "Kısa vadeli finansal yatırımlar", "b"),
    ("ticalacak", "Ticari alacaklar (kısa vadeli)", "b"),
    ("stok", "Stoklar", "b"),
    ("donen", "Toplam dönen varlıklar", "b"),
    ("mdv", "Maddi duran varlıklar", "b"),
    ("duran", "Toplam duran varlıklar", "b"),
    ("toplamvar", "Toplam varlıklar", "b"),
    ("ticborc", "Ticari borçlar (kısa vadeli)", "b"),
    ("kvyk", "Toplam kısa vadeli yükümlülükler", "b"),
    ("uvyk", "Toplam uzun vadeli yükümlülükler", "b"),
    ("toplamyuk", "Toplam yükümlülükler", "b"),
    ("ozkaynak", "Toplam özkaynaklar", "b"),
    ("anaozk", "Ana ortaklığa ait özkaynaklar", "b"),
    ("sermaye", "Ödenmiş sermaye (TL)", "b"),
    ("paysayisi", "Pay sayısı (adet)", "b"),
    ("H", "GELİR TABLOSU (TL, KAP işaretleriyle; giderler negatif)", None),
    ("hasilat", "Hasılat", "g"),
    ("smm", "Satışların maliyeti (-)", "g"),
    ("brut", "Brüt kâr", "g"),
    ("efk", "Esas faaliyet kârı", "g"),
    ("fingid", "Finansman giderleri (-)", "g"),
    ("vok", "Vergi öncesi kâr", "g"),
    ("netkar", "Dönem net kârı", "g"),
    ("ananetkar", "Ana ortaklığa ait net kâr", "g"),
    ("H", "DÖNEM BİLGİSİ", None),
    ("ay", "Gelir tablosu dönem uzunluğu (ay)", "p"),
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

# ---------------- Ham_Veri ----------------
hv = wb.active
hv.title = "Ham_Veri"
hv["A1"] = "Ham Veri — Migros (MGROS) ve BİM (BIMAS), KAP konsolide finansal raporları"
hv["A1"].font = title
hv["A2"] = ("Tutarlar TL (tam TL). TMS 29 enflasyon muhasebesi uygulanmış raporlar; her dönem için o dönemin kendi "
            "raporundaki cari dönem sütunu kullanılmıştır. Mavi = elle girilen KAP verisi, sarı = henüz girilmedi.")
hv["A2"].font = ital
hdr = ["Kalem"] + COLS + [f"Kaynak – {c}" for c in COLS]
for i, h in enumerate(hdr):
    c = hv.cell(row=4, column=i + 1, value=h)
    c.font, c.fill, c.alignment, c.border = hdrfont, hdrfill, Alignment(horizontal="center", wrap_text=True), box
ROW = {}
r = 5
for key, label, sec in ITEMS:
    if key == "H":
        hv.cell(row=r, column=1, value=label).font = bold
        for col in range(1, 10):
            hv.cell(row=r, column=col).fill = secfill
        r += 1
        continue
    ROW[key] = r
    hv.cell(row=r, column=1, value=label).font = norm
    for j, col in enumerate(COLS):
        cell = hv[f"{VC[j]}{r}"]
        src = hv[f"{SC[j]}{r}"]
        if key == "ay":
            cell.value = 12 if col.endswith("Q4") else 6
            src.value = "2025Q4: 01.01–31.12.2025 (12 ay); 2026Q2: 01.01–30.06.2026 (6 ay)"
            cell.font = blue
        else:
            d = data.get(col, {}).get(key)
            if d is not None and d.get("deger") is not None:
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
hv.column_dimensions["A"].width = 48
for c in VC:
    hv.column_dimensions[c].width = 20
for c in SC:
    hv.column_dimensions[c].width = 34
hv.freeze_panes = "B5"

def H(key, j):
    return f"Ham_Veri!${VC[j]}${ROW[key]}"

# ---------------- Parametreler ----------------
pr = wb.create_sheet("Parametreler")
pr["A1"] = "Parametreler"; pr["A1"].font = title
pr["A3"] = "Gün sayısı"; pr["A3"].font = bold
pr["B3"] = 365; pr["B3"].font = blue; pr["B3"].fill = yellow
pr["C3"] = "Ders notları 360 kullanıyorsa bu hücreyi değiştirin; tüm gün hesapları buna bağlı."; pr["C3"].font = ital
for i, h in enumerate(["", *COLS]):
    c = pr.cell(row=5, column=i + 1, value=h or "Parametre")
    c.font, c.fill, c.border = hdrfont, hdrfill, box
pr["A6"] = "Gelir tablosu dönem uzunluğu (ay)"
pr["A7"] = "Yıllıklandırma çarpanı = 12 / dönem uzunluğu"
PC = ["B", "C", "D", "E"]
for j in range(4):
    pr[f"{PC[j]}6"] = f"={H('ay', j)}"; pr[f"{PC[j]}6"].font = green
    pr[f"{PC[j]}7"] = f"=12/{PC[j]}6"; pr[f"{PC[j]}7"].number_format = "0.00"
    for rr in (6, 7):
        pr[f"{PC[j]}{rr}"].border = box
pr.column_dimensions["A"].width = 46
for c in PC:
    pr.column_dimensions[c].width = 16
GUN = "Parametreler!$B$3"
def CARP(j):
    return f"Parametreler!${PC[j]}$7"

# ---------------- ortak blok yazıcı ----------------
OC = ["D", "E", "F", "G"]   # oran sayfalarında değer sütunları
RESULT = {}  # oran adı -> (sayfa, satır, format)

def setup_ratio_sheet(ws, ttl, note):
    ws["A1"] = ttl; ws["A1"].font = title
    ws["A2"] = note; ws["A2"].font = ital
    for i, h in enumerate(["Oran", "Formül", "Adım", *COLS]):
        c = ws.cell(row=4, column=i + 1, value=h)
        c.font, c.fill, c.border = hdrfont, hdrfill, box
        c.alignment = Alignment(horizontal="center")
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 46
    ws.column_dimensions["C"].width = 44
    for c in OC:
        ws.column_dimensions[c].width = 18
    ws.freeze_panes = "D5"

def block(ws, start, name, formula, steps, sheetname):
    """steps: list of (etiket, fn(j, rowmap)->formül, format, tür) tür: 'link'|'calc'|'result'"""
    ws.cell(row=start, column=1, value=name).font = bold
    ws.cell(row=start, column=2, value=formula).font = norm
    for col in range(1, 8):
        ws.cell(row=start, column=col).fill = secfill
    rowmap = {}
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
        if kind == "result":
            RESULT[name] = (sheetname, r, fmt)
        r += 1
    return r + 1, rowmap

def ref(rowmap, label, j):
    return f"{OC[j]}{rowmap[label]}"

# ---------------- Likidite ----------------
lk = wb.create_sheet("Likidite")
setup_ratio_sheet(lk, "Likidite Oranları",
                  "Tüm değerler Ham_Veri sayfasına bağlı canlı formüllerdir. Dönem sonu bakiyeleri kullanılmıştır.")
r = 5
r, _ = block(lk, r, "Cari oran", "Dönen varlıklar / Kısa vadeli yabancı kaynaklar (KVYK)", [
    ("Pay: Toplam dönen varlıklar", lambda j, m: f"={H('donen', j)}", TL, "link"),
    ("Payda: KVYK", lambda j, m: f"={H('kvyk', j)}", TL, "link"),
    ("Sonuç: Cari oran", lambda j, m: f"={ref(m,'Pay: Toplam dönen varlıklar',j)}/{ref(m,'Payda: KVYK',j)}", RAT, "result"),
], "Likidite")
r, _ = block(lk, r, "Asit-test oranı", "(Dönen varlıklar − Stoklar) / KVYK", [
    ("Toplam dönen varlıklar", lambda j, m: f"={H('donen', j)}", TL, "link"),
    ("Stoklar", lambda j, m: f"={H('stok', j)}", TL, "link"),
    ("Pay: Dönen varlıklar − Stoklar", lambda j, m: f"={ref(m,'Toplam dönen varlıklar',j)}-{ref(m,'Stoklar',j)}", TL, "calc"),
    ("Payda: KVYK", lambda j, m: f"={H('kvyk', j)}", TL, "link"),
    ("Sonuç: Asit-test oranı", lambda j, m: f"={ref(m,'Pay: Dönen varlıklar − Stoklar',j)}/{ref(m,'Payda: KVYK',j)}", RAT, "result"),
], "Likidite")
r, _ = block(lk, r, "Nakit oranı", "(Nakit ve nakit benzerleri + KV finansal yatırımlar) / KVYK", [
    ("Nakit ve nakit benzerleri", lambda j, m: f"={H('nakit', j)}", TL, "link"),
    ("Kısa vadeli finansal yatırımlar", lambda j, m: f"={H('kvfy', j)}", TL, "link"),
    ("Pay: Nakit + KV finansal yatırımlar", lambda j, m: f"={ref(m,'Nakit ve nakit benzerleri',j)}+{ref(m,'Kısa vadeli finansal yatırımlar',j)}", TL, "calc"),
    ("Payda: KVYK", lambda j, m: f"={H('kvyk', j)}", TL, "link"),
    ("Sonuç: Nakit oranı", lambda j, m: f"={ref(m,'Pay: Nakit + KV finansal yatırımlar',j)}/{ref(m,'Payda: KVYK',j)}", RAT, "result"),
], "Likidite")
r, _ = block(lk, r, "Net işletme sermayesi (TL)", "Dönen varlıklar − KVYK", [
    ("Toplam dönen varlıklar", lambda j, m: f"={H('donen', j)}", TL, "link"),
    ("KVYK", lambda j, m: f"={H('kvyk', j)}", TL, "link"),
    ("Sonuç: Net işletme sermayesi (TL)", lambda j, m: f"={ref(m,'Toplam dönen varlıklar',j)}-{ref(m,'KVYK',j)}", TL, "result"),
], "Likidite")

# ---------------- Faaliyet ----------------
fa = wb.create_sheet("Faaliyet")
setup_ratio_sheet(fa, "Faaliyet Oranları",
                  "Gelir tablosu tutarları yıllıklandırma çarpanıyla (Parametreler) çarpılır: 2025Q4 ×1, 2026Q2 ×2. "
                  "Satışların maliyeti mutlak değerle kullanılır. Bilanço kalemleri dönem sonu bakiyesidir.")
r = 5
# Ortak yıllıklandırma bloğu
r, Y = block(fa, r, "Yıllıklandırma (ara adım)", "Gelir tablosu tutarı × (12 / dönem uzunluğu)", [
    ("Hasılat (dönem, KAP)", lambda j, m: f"={H('hasilat', j)}", TL, "link"),
    ("Satışların maliyeti (dönem, mutlak değer)", lambda j, m: f"=ABS({H('smm', j)})", TL, "link"),
    ("Yıllıklandırma çarpanı", lambda j, m: f"={CARP(j)}", "0.00", "link"),
    ("Gün sayısı", lambda j, m: f"={GUN}", "0", "link"),
    ("Yıllıklandırılmış hasılat", lambda j, m: f"={ref(m,'Hasılat (dönem, KAP)',j)}*{ref(m,'Yıllıklandırma çarpanı',j)}", TL, "calc"),
    ("Yıllıklandırılmış satışların maliyeti", lambda j, m: f"={ref(m,'Satışların maliyeti (dönem, mutlak değer)',j)}*{ref(m,'Yıllıklandırma çarpanı',j)}", TL, "calc"),
], "Faaliyet")
def YH(j): return f"{OC[j]}${Y['Yıllıklandırılmış hasılat']}"
def YS(j): return f"{OC[j]}${Y['Yıllıklandırılmış satışların maliyeti']}"
def GN(j): return f"{OC[j]}${Y['Gün sayısı']}"

r, _ = block(fa, r, "Stok devir hızı", "Yıllıklandırılmış satışların maliyeti / Stoklar", [
    ("Pay: Yıllıklandırılmış SMM", lambda j, m: f"={YS(j)}", TL, "calc"),
    ("Payda: Stoklar", lambda j, m: f"={H('stok', j)}", TL, "link"),
    ("Sonuç: Stok devir hızı (kez)", lambda j, m: f"={ref(m,'Pay: Yıllıklandırılmış SMM',j)}/{ref(m,'Payda: Stoklar',j)}", RAT, "result"),
], "Faaliyet")
stok_dh = RESULT["Stok devir hızı"][1]
r, _ = block(fa, r, "Stokta kalma süresi (gün)", "Gün sayısı / Stok devir hızı", [
    ("Gün sayısı", lambda j, m: f"={GN(j)}", "0", "calc"),
    ("Stok devir hızı", lambda j, m: f"={OC[j]}{stok_dh}", RAT, "calc"),
    ("Sonuç: Stokta kalma süresi (gün)", lambda j, m: f"={ref(m,'Gün sayısı',j)}/{ref(m,'Stok devir hızı',j)}", DAY, "result"),
], "Faaliyet")
r, _ = block(fa, r, "Alacak devir hızı", "Yıllıklandırılmış hasılat / Ticari alacaklar", [
    ("Pay: Yıllıklandırılmış hasılat", lambda j, m: f"={YH(j)}", TL, "calc"),
    ("Payda: Ticari alacaklar (KV)", lambda j, m: f"={H('ticalacak', j)}", TL, "link"),
    ("Sonuç: Alacak devir hızı (kez)", lambda j, m: f"={ref(m,'Pay: Yıllıklandırılmış hasılat',j)}/{ref(m,'Payda: Ticari alacaklar (KV)',j)}", RAT, "result"),
], "Faaliyet")
al_dh = RESULT["Alacak devir hızı"][1]
r, _ = block(fa, r, "Ortalama tahsil süresi (gün)", "Gün sayısı / Alacak devir hızı", [
    ("Gün sayısı", lambda j, m: f"={GN(j)}", "0", "calc"),
    ("Alacak devir hızı", lambda j, m: f"={OC[j]}{al_dh}", RAT, "calc"),
    ("Sonuç: Ortalama tahsil süresi (gün)", lambda j, m: f"={ref(m,'Gün sayısı',j)}/{ref(m,'Alacak devir hızı',j)}", DAY, "result"),
], "Faaliyet")
r, _ = block(fa, r, "Borç devir hızı", "Yıllıklandırılmış satışların maliyeti / Ticari borçlar", [
    ("Pay: Yıllıklandırılmış SMM", lambda j, m: f"={YS(j)}", TL, "calc"),
    ("Payda: Ticari borçlar (KV)", lambda j, m: f"={H('ticborc', j)}", TL, "link"),
    ("Sonuç: Borç devir hızı (kez)", lambda j, m: f"={ref(m,'Pay: Yıllıklandırılmış SMM',j)}/{ref(m,'Payda: Ticari borçlar (KV)',j)}", RAT, "result"),
], "Faaliyet")
bo_dh = RESULT["Borç devir hızı"][1]
r, _ = block(fa, r, "Ortalama ödeme süresi (gün)", "Gün sayısı / Borç devir hızı", [
    ("Gün sayısı", lambda j, m: f"={GN(j)}", "0", "calc"),
    ("Borç devir hızı", lambda j, m: f"={OC[j]}{bo_dh}", RAT, "calc"),
    ("Sonuç: Ortalama ödeme süresi (gün)", lambda j, m: f"={ref(m,'Gün sayısı',j)}/{ref(m,'Borç devir hızı',j)}", DAY, "result"),
], "Faaliyet")
sk = RESULT["Stokta kalma süresi (gün)"][1]; ts = RESULT["Ortalama tahsil süresi (gün)"][1]; od = RESULT["Ortalama ödeme süresi (gün)"][1]
r, _ = block(fa, r, "Nakit dönüşüm süresi (gün)", "Stokta kalma süresi + Ortalama tahsil süresi − Ortalama ödeme süresi", [
    ("Stokta kalma süresi", lambda j, m: f"={OC[j]}{sk}", DAY, "calc"),
    ("Ortalama tahsil süresi", lambda j, m: f"={OC[j]}{ts}", DAY, "calc"),
    ("Ortalama ödeme süresi", lambda j, m: f"={OC[j]}{od}", DAY, "calc"),
    ("Sonuç: Nakit dönüşüm süresi (gün)", lambda j, m: f"={ref(m,'Stokta kalma süresi',j)}+{ref(m,'Ortalama tahsil süresi',j)}-{ref(m,'Ortalama ödeme süresi',j)}", DAY, "result"),
], "Faaliyet")
r, _ = block(fa, r, "Maddi duran varlık devir hızı", "Yıllıklandırılmış hasılat / Maddi duran varlıklar", [
    ("Pay: Yıllıklandırılmış hasılat", lambda j, m: f"={YH(j)}", TL, "calc"),
    ("Payda: Maddi duran varlıklar", lambda j, m: f"={H('mdv', j)}", TL, "link"),
    ("Sonuç: MDV devir hızı (kez)", lambda j, m: f"={ref(m,'Pay: Yıllıklandırılmış hasılat',j)}/{ref(m,'Payda: Maddi duran varlıklar',j)}", RAT, "result"),
], "Faaliyet")
r, _ = block(fa, r, "Aktif devir hızı", "Yıllıklandırılmış hasılat / Toplam varlıklar", [
    ("Pay: Yıllıklandırılmış hasılat", lambda j, m: f"={YH(j)}", TL, "calc"),
    ("Payda: Toplam varlıklar", lambda j, m: f"={H('toplamvar', j)}", TL, "link"),
    ("Sonuç: Aktif devir hızı (kez)", lambda j, m: f"={ref(m,'Pay: Yıllıklandırılmış hasılat',j)}/{ref(m,'Payda: Toplam varlıklar',j)}", RAT, "result"),
], "Faaliyet")

# ---------------- Karsilastirma ----------------
ks = wb.create_sheet("Karsilastirma")
ks["A1"] = "Karşılaştırma — 6 aylık değişim ve MGROS − BIMAS farkı"; ks["A1"].font = title
ks["A2"] = "Değişim = 2026Q2 − 2025Q4; % değişim = değişim / |2025Q4|. Fark = MGROS − BIMAS."; ks["A2"].font = ital
heads = ["Oran", "MGROS 2025Q4", "MGROS 2026Q2", "MGROS değişim", "MGROS % değişim",
         "BIMAS 2025Q4", "BIMAS 2026Q2", "BIMAS değişim", "BIMAS % değişim",
         "Fark 2025Q4 (MGROS−BIMAS)", "Fark 2026Q2 (MGROS−BIMAS)"]
for i, h in enumerate(heads):
    c = ks.cell(row=4, column=i + 1, value=h)
    c.font, c.fill, c.border = hdrfont, hdrfill, box
    c.alignment = Alignment(horizontal="center", wrap_text=True)
rr = 5
for name, (sh, row, fmt) in RESULT.items():
    if name == "Yıllıklandırma (ara adım)":
        continue
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
ks.column_dimensions["A"].width = 32
for i in range(2, 12):
    ks.column_dimensions[L(i)].width = 16
ks.row_dimensions[4].height = 32

# ---------------- Kontroller ----------------
kt = wb.create_sheet("Kontroller")
kt["A1"] = "Veri kontrolleri (tümü 0 / TAMAM olmalı)"; kt["A1"].font = title
for i, h in enumerate(["Kontrol", *COLS]):
    c = kt.cell(row=3, column=i + 1, value=h)
    c.font, c.fill, c.border = hdrfont, hdrfill, box
checks = [
    ("Toplam varlıklar − (Toplam yükümlülükler + Toplam özkaynaklar)", lambda j: f"={H('toplamvar',j)}-({H('toplamyuk',j)}+{H('ozkaynak',j)})"),
    ("(Dönen + Duran varlıklar) − Toplam varlıklar", lambda j: f"=({H('donen',j)}+{H('duran',j)})-{H('toplamvar',j)}"),
    ("(KVYK + UVYK) − Toplam yükümlülükler", lambda j: f"=({H('kvyk',j)}+{H('uvyk',j)})-{H('toplamyuk',j)}"),
    ("Brüt kâr − (Hasılat − |Satışların maliyeti|)", lambda j: f"={H('brut',j)}-({H('hasilat',j)}-ABS({H('smm',j)}))"),
]
for i, (lab, fn) in enumerate(checks):
    rr = 4 + i
    kt.cell(row=rr, column=1, value=lab)
    for j in range(4):
        c = kt[f"{PC[j]}{rr}"]; c.value = fn(j); c.number_format = TL; c.border = box
st = 4 + len(checks)
kt.cell(row=st, column=1, value="Durum").font = bold
for j in range(4):
    c = kt[f"{PC[j]}{st}"]
    c.value = f'=IF(SUMPRODUCT(ABS({PC[j]}4:{PC[j]}{st-1}))<1,"TAMAM","KONTROL ET")'
    c.font = bold; c.border = box
kt["A" + str(st + 2)] = "Not: <1 TL tolerans (yuvarlama). Bir fark çıkarsa Ham_Veri girişini rapordan yeniden kontrol edin."
kt["A" + str(st + 2)].font = ital
kt.column_dimensions["A"].width = 58
for c in PC:
    kt.column_dimensions[c].width = 18

for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if c.font is None or c.font.name != F:
                c.font = Font(name=F, bold=c.font.bold, italic=c.font.italic, color=c.font.color, size=c.font.size)

wb.save(OUT)
print("Kaydedildi:", OUT)
