# Dolap Karşılaştırma (dolap_eslestir.py + dolap_verilecek_liste.py)

![Ekran görüntüsü — demo veriler](../gorseller/12_dolap.png)

"Dolap almamış personel listesi" görevi için eşleştirme ve rapor araçları.

## Girdiler (script'lerle AYNI klasörde; arşivde SENTETİK örnekleriyle gelir)
Bu klasördeki Excel'ler RASTGELE üretilmiş örnek veridir — bkz.
`ORNEK_VERI_BEYANI.md` (kişisel veri içermez; araçlar bunlarla hemen
denenebilir). Gerçek kullanımda kendi dosyalarınla değiştir:
- `ZIMMET LİSTESİ.xlsm` → DOLAP sayfası (B sütunu isimler; adı geçen dolap aldı)
- `DOĞRU PERSONEL İSİMLERİ.xlsx` → güncel personel (AD-SOYAD, TC)
- `AYRILMIŞ PERSONEL LİSTESİ.xlsx` → işten ayrılanlar (AD-SOYAD, TC)
- `SAHA PERSONELLERİ LİSTE.xlsx` → AD-SOYAD, TC, İŞE GİRİŞ TARİHİ

## dolap_eslestir.py — eşleştirme sihirbazı
1. Güncel listeyle birebir eşleşen → dolap ALMIŞ (otomatik).
   Ayrılmış listeyle birebir eşleşen → AYRILMIŞ (otomatik). İki listede
   birden olan güncel sayılır.
2. Kalanlar SIRAYLA sorulur: benzer isimler yüzdeyle listelenir (ayrılmışlar
   [AYRILMIŞ] etiketli), ilk öneri sadece ≥%80 benzerse hazır seçili gelir;
   kutuya yazınca canlı arama. ↓ seç + Enter, "Hiçbir listede yok" düğmesi,
   ↶ Geri. Kararlar `dolap_kararlar.json`a anında yazılır (yarıda bırak-devam et).
3. Bitince rapor: `DOLAP_KARSILASTIRMA_SONUC.xlsx` (DOLAP ALMAMIŞ + DÜZELTİLENLER
   + AYRILMIŞ + LİSTEDE YOK + ÖZET) ve düzeltmeler AÇIK Excel'deki ZIMMET
   LİSTESİ'ne COM ile OTOMATİK uygulanır (xlsm'e openpyxl KULLANILMAZ — makro
   ölür; dosya kaydedilmez, kontrol sana kalır).

## dolap_atama.py — ayrılmış dolapları devretme
AÇIKLAMA'sında "AYRILMIŞ" yazan (kırmızı) dolap satırlarını sırayla gösterir;
Ara kutusuna yeni kullanıcının adı yazılır (doğru listeden canlı süzülür, ilk
sonuç seçili), Enter = AÇIK Excel'de isim değişir + AYRILMIŞ notu ve kırmızı
biçim temizlenir. Sağ Shift/Atla = geç (kırmızı kalır; tekrar açılışta yine
gelir — kaldığı yerden devam otomatiktir, karar dosyası yoktur). Zaten dolabı
olan seçilirse "← DOLABI VAR!" uyarısı, ikinci Enter ile yazılır; "Son yazımı
geri al" tek adım geri alır (COM yazmaları Excel Ctrl+Z'sine girmez).
Kaydetmez; bitiş düğmesi kaydedip DOLAP VERİLECEK listesini yeniden üretir.
Başlarken otomatik yedek alır.

## dolap_verilecek_liste.py — teslim listesi
Saha listesinden dolap almışları düşer, `DOLAP VERİLECEK PERSONELLER.xlsx`
üretir (sıra, ad, TC, işe giriş tarihi; alfabetik + filtreli). Eşleştirme
kararlarını aynen kullanır; kararsız isim kaldıysa uyarır.

Gereksinim: `openpyxl`, `pywin32` (uygulama adımı için)
