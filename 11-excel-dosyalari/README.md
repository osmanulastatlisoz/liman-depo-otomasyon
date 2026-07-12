# Excel Dosyaları

Bu klasör, araç setinin üzerinde çalıştığı **gerçek .xlsm çalışma kitaplarının**
yerini tutar. Dosyalar gerçek veri (personel adları, stok kayıtları) içerdiği
için **bu depoya dahil edilmemiştir** (`.gitignore` ile dışlanır) — burada
yalnızca ne oldukları anlatılır:

- **2026 DEPO STOK.xlsm** — eski ana stok kütüğü (makrolu; STOK GİRİŞ ÇIKIŞ
  ~31.000 hareket, kritik analiz, KKD iade vb. sayfalar)
- **TCH_STOK_TAKIP_CALISMASI GÜNCEL.xlsm** — güncel stok takip sistemi
  (STOK HAREKETLERİ, DEPO, dinamik dashboard, resmi personel/malzeme
  listeleri; yazım yardımcısı bu dosyayla çalışır). KRİTİK ANALİZ'de
  B4 stok kaynağı seçici (GÜNCEL/LİMAN), KALAN/GİREN ada dayalı SUMIF,
  tür/sınıf/kullanım yeri kategorileri.
- **2025_SECIMLI_CALISAN_STOK_DASHBOARD.xlsm** — 2025 seçimli dashboard çalışması
- **2026 STOK ÇALIŞMASI MALZEME SINIFLANDIRMASI 1.xlsm** — malzeme
  sınıflandırma çalışması

Bu dosyaların **makro kaynak kodları** `10-excel-vba/` klasöründe dışa
aktarılmış olarak mevcuttur. Araçları denemek için `12`/`13` klasörlerindeki
**sentetik örnek Excel'ler** kullanılabilir (bkz. oradaki `ORNEK_VERI_BEYANI.md`).
