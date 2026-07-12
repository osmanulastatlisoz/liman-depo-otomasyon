# Mükerrer Temizlik (mukerrer_tespit.py + mukerrer_sil.py)

![Ekran görüntüsü — demo veriler](../gorseller/05_mukerrer_temizlik.png)

2026 DEPO STOK.xlsm'de yanlışlıkla çift girilmiş ÇIKIŞ kayıtlarını bulur ve
onayladıklarını Excel'i HİÇ AÇMADAN, makrolara dokunmadan siler.

## İki aşamalı güvenli akış
1. **`1_TARAMA YAP.bat`** (mukerrer_tespit.py): aynı tarih + personel + stok adıyla
   birden fazla ÇIKIŞ olan grupları bulur, `MUKERRER_KAYIT_RAPORU.xlsx` üretir.
   Kaynak dosyayı DEĞİŞTİRMEZ. Raporda her grubun yanında SİL? sütunu vardır.
2. Raporu Excel'de aç, silinecek grupların SİL? hücresine **E** yaz, kaydet.
3. **`2_MUKERRERLERI SIL.bat`** (mukerrer_sil.py): E işaretli grupların fazla
   kayıtlarını siler (her grubun İLK kaydı kalır). Silmeden önce tarihli yedek
   alır ve her hedef satırın içeriğini doğrular.

## Teknik: XML ameliyatı (_satir_sil_xml.py)
Silme, .xlsm ZIP'inin içindeki sayfa XML'i doğrudan düzenlenerek yapılır:
- Makrolar, butonlar/çizimler, diğer sayfalar bire bir korunur
- Sayfa XML yolu workbook.xml + rels'ten İSİMLE çözülür (sabit varsayılmaz)
- calcChain.xml kaldırılır (Excel ilk açılışta kendisi yeniden kurar)
- Excel lisansı/kurulumu gerektirmez

Kaynak dosya yolu esnektir: önce kayıtlı ayar, sonra masaüstü; bulunamazsa
ilk çalıştırmada seçtirir (`%LOCALAPPDATA%` altındaki `DepoAraclari/yollar.json`).
Rapor, kaynak dosyanın yanına yazılır.

Gereksinim: `openpyxl`
