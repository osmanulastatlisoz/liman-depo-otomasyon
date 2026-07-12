# Beden Listesi Eşleştir (beden_eslestir.py)

![Ekran görüntüsü — demo veriler](../gorseller/13_beden.png)

Beden listesindeki personel adlarını düzeltir, işten ayrılmışların satırlarını
listeden çıkarır. `12-dolap-karsilastirma` aracının satır-silmeli uyarlaması.

## Girdiler (script'le AYNI klasörde; arşivde SENTETİK örnekleriyle gelir)
Bu klasördeki Excel'ler RASTGELE üretilmiş örnek veridir — bkz.
`ORNEK_VERI_BEYANI.md` (kişisel veri içermez; araç bunlarla hemen
denenebilir). Gerçek kullanımda kendi dosyalarınla değiştir:
- `KIŞLIK BEDEN LİSTESİ.xlsx` → "KIŞLIK BEDEN LİSTESİ" sayfası
  (A=isim, B-E beden sütunları, F=renk — renk hücreleri boyalıdır)
- `DOĞRU PERSONEL İSİMLERİ.xlsx` → güncel personel (AD-SOYAD, TC)
- `AYRILMIŞ PERSONEL LİSTESİ.xlsx` → işten ayrılanlar (AD-SOYAD, TC)

## İşleyiş
1. Güncel listeyle birebir eşleşen KALIR (yazım farkı düzeltilir);
   ayrılmışla birebir eşleşen otomatik ÇIKARILIR; iki listede birden
   olan güncel sayılır.
2. Kalanlar SIRAYLA sorulur: öneriler yüzdeli, ayrılmışlar [AYRILMIŞ]
   etiketli, ≥%80 benzer ilk öneri hazır seçili; canlı arama; "Hiçbir
   listede yok" = çıkarılır. Kararlar `beden_kararlar.json`a anında yazılır.
3. Bitince otomatik: dosya YEDEKLENİR, isimler düzeltilir, çıkarılacak
   satırlar SİLİNİR (COM — boyalar/biçimler korunur, tersten silme +
   isim doğrulama). Dosya KAYDEDİLMEZ; kontrol edip Ctrl+S sana kalır.
4. Rapor `BEDEN_DUZELTME_SONUC.xlsx`: DÜZELTİLENLER + ÇIKARILANLAR
   (silinenlerin beden/renk verileri ve TC'leriyle — veri kaybolmaz) + ÖZET.

## Başka listeye uyarlama
Dosya başındaki AYARLAR bloğunda `BEDEN` (dosya adı), `SAYFA` ve
`VERI_SUTUN` değerlerini değiştir; rapordaki ÇIKARILANLAR sütun
başlıkları (PANTOLON/KAZAK/...) farklı düzende `rapor_yaz` içinde güncellenmeli.

## Notlar
- Uygulamayı BİR kez yap (silme sonrası satır numaraları kayar; isim
  doğrulaması ikinci çalıştırmayı zararsız atlatır ama gerek yoktur).
- Zombi Excel'e karşı bağlantı sağlık kontrolü vardır (Workbooks.Count).

Gereksinim: `openpyxl`, `pywin32`
