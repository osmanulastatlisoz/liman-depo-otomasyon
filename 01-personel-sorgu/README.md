# Personel Sorgu (birlesik_stok_sorgu.py)

![Ekran görüntüsü — demo veriler](../gorseller/01_personel_sorgu.png)

Formül yüklü stok Excel'lerini açmadan "bu personel ne almış?" sorusuna
saniyeler içinde cevap veren masaüstü arama penceresi.

## Ne yapar
- `2026 DEPO STOK.xlsm → STOK GİRİŞ ÇIKIŞ` ile
  `TCH_STOK_TAKIP_CALISMASI GÜNCEL.xlsm → STOK HAREKETLERİ` sayfalarını birleştirir.
- Geçiş döneminde iki dosyaya da girilmiş kayıtları teke indirir
  (anahtar: tarih + işlem + personel + malzeme + miktar, Türkçe harf katlamalı).
- Formülsüz `BİRLEŞİK STOK HAREKETLERİ.xlsx` dosyasını atomik olarak günceller.
- tkinter penceresinde personel ve/veya malzeme adıyla CANLI arama.

## Kullanım
- Çift tık: `PERSONEL SORGU.bat` — pencere önbellekten ANINDA açılır,
  arka planda kaynaklardan tazelenir (durum çubuğunda "✓ Güncellendi").
- Komut: `python birlesik_stok_sorgu.py --rebuild` (GUI'siz, sadece dosya üretimi)

## Özellikler
- Kelime sırası önemsiz arama ("yılmaz ahmet" = "ahmet yılmaz"), 2 harf yeter
- Bulamayınca difflib ile "Şunu mu demek istediniz?" önerisi
- İşlem filtresi (varsayılan ÇIKIŞ) ve "Malzeme toplamları" özet görünümü
- Kaynak dosyalar Excel'de açıkken bile çalışır (geçici kopyadan okur)

## Yollar (esnek)
Kaynak Excel'ler önce kayıtlı ayardan, sonra masaüstünden aranır; bulunamazsa
ilk çalıştırmada seçtirir ve `%LOCALAPPDATA%\DepoAraclari\yollar.json`'a
kaydeder. Birleşik xlsx çıktısı TCH dosyasının yanına yazılır; önbellek:
`%LOCALAPPDATA%\BirlesikStok\`

Gereksinim: `openpyxl`
