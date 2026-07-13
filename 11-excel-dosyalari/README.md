# Excel Dosyaları

Araç setinin üzerinde çalıştığı .xlsm çalışma kitapları. Gerçek veri içerenler
depoya dahil edilmemiştir; **anonimleştirilmiş bir kopya** ise incelenebilir
halde buradadır:

## ✅ Bu depoda: `TCH_STOK_TAKIP_CALISMASI GÜNCEL.xlsm` (anonim kopya)

Güncel stok takip sistemi — STOK HAREKETLERİ, DEPO, dinamik dashboard,
KRİTİK ANALİZ (B4 stok kaynağı seçici, ada dayalı SUMIF), personel/malzeme
listeleri. Yazım yardımcısı (02) ve personel sorgu (01) bu dosyayla çalışır.
Makrolar ve tüm biçimlendirme yerindedir.

**Anonimleştirme (ZIP cerrahisiyle, yapı bire bir korunarak):**
- ~1.550 personel adı **tutarlı kurgu isimlerle** değiştirildi (aynı kişi her
  sayfada aynı kurgu isme gider — sorgu/dashboard mantığı bozulmaz)
- Serbest metin sütunları (AÇIKLAMA/PLAKA), TEDARİKÇİ ve YEDEK (geri alma
  günlüğü) sayfası tamamen boşaltıldı
- Doğrulama: tüm gerçek isimler, dosyanın **bütün iç parçalarında** (hücreler,
  formül önbellekleri, paylaşılan metinler) tarandı — **0 kalıntı**

## ❌ Bu depoda olmayanlar (gerçek verili yedekler)

- **2026 DEPO STOK.xlsm** — eski ana stok kütüğü (~31.000 hareket, kritik
  analiz, KKD iade vb.; makro kaynakları `10-excel-vba/` klasöründe)
- **2025_SECIMLI_CALISAN_STOK_DASHBOARD.xlsm** — 2025 seçimli dashboard çalışması
- **2026 STOK ÇALIŞMASI MALZEME SINIFLANDIRMASI 1.xlsm** — malzeme sınıflandırma

Araçları denemek için `12`/`13` klasörlerindeki **sentetik örnek Excel'ler** de
kullanılabilir (bkz. `ORNEK_VERI_BEYANI.md`).
