# Excel VBA Modülleri (yedek / referans)

![Ekran görüntüsü — demo veriler](../gorseller/10_excel_vba.png)

Stok Excel'lerinin İÇİNDE zaten kayıtlı olan makroların dışa aktarılmış
kopyaları. Dosyalar başka PC'ye taşınırken kod kaybolmaz (xlsm içinde gider);
buradakiler okumak, karşılaştırmak ve gerekirse yeniden kurmak içindir.

## İçerik
- **DEPO STOK - KOD ve KURULUM REHBERI.txt** — 2026 DEPO STOK.xlsm makrolarının
  tam kurulum rehberi (GuvenliKutu, hız optimizasyonları, manuel hesap notları)
- **modDinamikDashboard.bas** — TCH dosyasındaki dinamik dashboard modülü
- **STOK GIRIS CIKIS - SAYFA MODULU.txt / stok_giris_cikis_son_hali.txt** —
  2026 dosyası sayfa modülü (hızlı giriş, doğrulama)
- **stok hareketleri modülü - GUNCEL.txt / stok takip stok hareketleri sayfası
  modülü.txt** — TCH STOK HAREKETLERİ sayfa modülü sürümleri
- **MODUL3 / STOK_HAREKETLERI_SAYFASI** — kuruluma aktarılan modül sürümleri
- **MALZEME KISAYOL MAKRO.txt / YAZLIKLAR_yeni.txt / excel makro
  kısaltmalar.txt** — Worksheet_Change kısayol makroları
  (hücreye "KS" yaz → "POLO YAKA SWEAT S BEDEN" açılır)
- **personel ismi düzeltme.txt** — toplu isim düzeltme yardımcı makrosu

## modDinamikDashboard — kodlama tuzağı (yaşandı!)
`modDinamikDashboard.bas` UTF-8'dir; Excel'in VBA içe aktarıcısı ise ANSI
(cp1254) bekler. UTF-8'i doğrudan Import edersen Türkçe sabitler bozulur
(`"DİNAMİK DASHBOARD"` → `"DÄ°NAMÄ°K DASHBOARD"`) ve tüm makrolar "sayfa yok"
hatası verir. **Import için `modDinamikDashboard_cp1254.bas` kullan** (aynı
kodun cp1254 kodlamalı hâli, bu klasörde).

## Kurulum ipucu
VBA import ederken Türkçe karakter için cp1254 kodlamasına dikkat et;
VBProject erişimi yalnızca interaktif Excel'de ve Güven Merkezi'nden
"VBA proje nesne modeline erişime güven" açıkken çalışır.
