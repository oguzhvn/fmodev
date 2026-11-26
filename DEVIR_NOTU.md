# Devir notu — 1. aşamadan 2. ve 3. aşamaya

**Proje:** Financial Management grup projesi — Migros (MGROS) ve rakip BİM (BIMAS)
**Dönemler:** 2025Q4 (31.12.2025, 12 aylık) ve 2026Q2 (30.06.2026, 6 aylık)
**Grup:** Muhammed Oğuzhan Sağmak, Seher Naz Yomluer, Buse Çalış, Elif Aldemir

**GitHub reposu:** https://github.com/oguzhvn/fmodev — Dosyaların güncel hali burada. Sayfadaki yeşil **Code → Download ZIP** düğmesiyle hepsini indirebilirsiniz; GitHub hesabı veya program kurmak gerekmez (repo gizliyse önce davet edilmeniz gerekir).

## Dosyalar

| Dosya | Ne işe yarar |
|---|---|
| `FM_Project_MGROS_BIMAS.xlsx` | Ortak Excel. Bütün oranlar bu dosyadaki rakamlardan hesaplanmalı. |
| `KAP_Raporlari/*_EN.pdf` | KAP'tan indirilen İngilizce konsolide finansal raporlar (Excel'de kaynak olarak gösterilen sürüm) |
| `KAP_Raporlari/*.pdf` (ekisiz) | Aynı raporların Türkçe sürümü (rakamlar aynı) |
| `KAP_Raporlari/Rapor_Linkleri.md` | Raporların KAP bildirim linkleri ve yayın tarihleri |
| `teams_bildirim.txt` | Hocaya gönderilen grup ve şirket bildirimi |

## Excel'le nasıl çalışılır

- Excel'i doğrudan açıp düzenleyin; ek bir program gerekmez.
- **`Raw_Data`, `Liquidity`, `Activity`, `Comparison`, `Checks` sayfalarındaki hücreleri değiştirmeyin.** Bu sayfalar 1. aşamanın kontrol edilmiş sonucudur.
- Kendi bölümünüz için yeni bir sayfa açın (ör. `Leverage`, `Profitability`, `Market`, `DuPont`).
- Aynı dosyayı iki kişi aynı anda düzenlerse değişiklikler kaybolabilir. Ya sırayla çalışın, ya da herkes kendi kopyasında kendi sayfasını hazırlasın; en sonda bir kişi sayfaları tek dosyada birleştirsin (sayfa sekmesine sağ tık → "Taşı veya Kopyala").
- Eksik bir kalem gerekirse (ör. hisse fiyatı) `Raw_Data`'nın en altına yeni bir satır ekleyin ve yanına kaynağını yazın.

## Excel'in yapısı

- **Raw_Data:** Bilanço ve gelir tablosu kalemleri, 4 sütun (MGROS 2025Q4, MGROS 2026Q2, BIMAS 2025Q4, BIMAS 2026Q2). Her rakamın yanında kaynak (rapor + PDF sayfası) var.
- **Parameters:** Gün sayısı (365) ve yıllıklandırma çarpanı (2025Q4 = 1, 2026Q2 = 2).
- **Liquidity / Activity:** 1. aşamanın oranları; pay, payda, sonuç ve açık hesaplama metniyle.
- **Comparison:** 6 aylık değişim ve MGROS − BIMAS farkı.
- **Checks:** Bilanço eşitliği vb. kontroller (hepsi OK).

Kendi oranlarınızı `Raw_Data` hücrelerine bağlı formüllerle kurun (ör. MGROS 2025Q4 toplam varlıklar için `=Raw_Data!B13`); rakamı elle kopyalamayın. Aynı düzeni kullanırsanız (pay / payda / sonuç) rapor tutarlı görünür.

## Mutlaka dikkat edilmesi gerekenler

1. **Birim:** Tutarlar **bin TL** (raporlardaki gibi). Pay sayısı ise **adet**. Piyasa değeri hesaplarken bin TL'yi ×1.000 ile tam TL'ye çevirin.
2. **Yıllıklandırma:** 2026Q2 gelir tablosu 6 aylıktır. Gelir tablosu kalemini bilanço kalemine bölen oranlarda (ROA, ROE, DuPont'taki aktif devir hızı ve özkaynak çarpanı zinciri) 2026Q2 kâr ve hasılatını `Parameters` sayfasındaki çarpanla (×2) yıllıklandırın. Gelir tablosu kalemini yine gelir tablosu kalemine bölen oranlarda (kâr marjları, faiz karşılama oranı) yıllıklandırma gerekmez.
3. **Hisse fiyatları Excel'de yok:** Piyasa oranları (F/K, PD/DD, piyasa değeri) için 31.12.2025 ve 30.06.2026 kapanış fiyatlarını BIST/KAP'tan alıp `Raw_Data`'ya veya yeni bir sayfaya kaynağıyla birlikte ekleyin.
4. **BİM'in pay sayısı değişti:** Mayıs 2026'daki %100 bedelsiz sermaye artırımı nedeniyle 600 milyon → 1,2 milyar pay. 2025Q4 ile 2026Q2 hisse başına rakamları (EPS, defter değeri/hisse) doğrudan karşılaştırılamaz; raporda bunu belirtin ya da 2025Q4'ü 1,2 milyar paya göre düzeltin.
5. **Raporlanan EPS ile hesaplanan EPS farklı çıkabilir:** Şirketler ağırlıklı ortalama pay sayısını ve geri alınmış payları kullanır. Kendi hesabınızda hangi pay sayısını kullandığınızı yazın.
6. **Migros'un esas faaliyet sonucu zarar:** Her iki dönemde de "Operating loss" var (2025: −8.252.604 bin TL; 2026 ilk 6 ay: −10.858.537 bin TL). Bu yüzden faaliyet kârına dayalı faiz karşılama oranı negatif çıkar. Vergi öncesi kârın pozitif olması TMS 29'dan gelen **net parasal pozisyon kazancı** sayesindedir (2025: 21.102.605; 2026 6 ay: 16.733.047 bin TL). Yorumda bundan bahsetmek iyi olur.
7. **Finansman giderleri ve satışların maliyeti negatif işaretli** girildi (raporlardaki gibi). Oranlarda `ABS()` ile mutlak değer kullanın.
8. **TMS 29 (enflasyon muhasebesi):** Her dönem kendi raporundaki rakamla alındı. 2025Q4 rakamları Aralık 2025, 2026Q2 rakamları Haziran 2026 alım gücüyle ifade edilmiştir; TL tutarlarının 6 aylık karşılaştırmasında bu farkı dikkate alın (oranları büyük ölçüde etkilemez).
9. **Raporlardaki iki etiket hatası** (Excel'de doğru ele alındı, hoca PDF'e bakarsa diye): MGROS 2025 İngilizce bilançosunda (PDF s.9) kısa vadeli yükümlülükler toplamı yanlışlıkla "Total non-current liabilities" diye etiketlenmiş; BİM bilançosundaki "Total liabilities" satırı aslında "toplam kaynaklar"dır.

## Raw_Data'da 2. ve 3. aşama için hazır olan kalemler

Toplam varlıklar, toplam yükümlülükler, kısa/uzun vadeli yükümlülükler, toplam özkaynaklar, ana ortaklığa ait özkaynaklar, hasılat, brüt kâr, esas faaliyet kârı, finansman giderleri, vergi öncesi kâr, net kâr, ana ortaklığa ait net kâr, ödenmiş sermaye, pay sayısı.

Gerekirse raporlardan eklenebilecekler: banka kredileri ve kiralama yükümlülükleri (faizli borç), finansman gelirleri, vergi gideri, ödenen temettü.
