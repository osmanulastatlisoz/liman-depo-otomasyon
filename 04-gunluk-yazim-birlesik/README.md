# Günlük Satır + Yazım Yardımcısı (gunluk_yazim_birlesik.py)

![Ekran görüntüsü — demo veriler](../gorseller/04_birlesik_mod.png)

İki aracı TEK süreçte, çakışmadan birlikte çalıştıran birleşik mod.
Akış: PDF seç → satır şeridi üstte kurulur → yazım yardımcısı altında açılır →
her şeyi yardımcı kutusundan yazarsın, Excel'e hiç geçmezsin.

## Kilit tasarım: Sağ Shift'te sentetik tuş YOK
Tek başına gunluk_satir, Sağ Shift'te Excel'e klavye tuşları gönderir — odak
yardımcıdayken bu tuşlar yanlış yere giderdi. Birleşik modda kayıtlı tuş dizisi
NET kaydırmaya çevrilir ve hücre COM ile taşınır (down,left,left = 1 aşağı 2 sol).
Odak hangi penceredeyse fark etmez.

## Sağ Shift = dört iş tek tuşta
1. Kutuda yazılan (veya ↓ ile seçilen) aktif hücreye YAZILIR
2. Excel imleci makro kadar taşınır (yeni satır başı)
3. PDF şeridi sıradaki satıra geçer
4. Yardımcının kutusu temizlenir

Sol Shift + Sağ Shift: şerit ilerler, Excel'e dokunulmaz.
Şeritteki "✎ Yardımcı" düğmesi yardımcı penceresini gizler/gösterir.

## Modül bağımlılığı (esnek)
Bu dosya orijinal iki aracı MODÜL olarak içe aktarır ve onları sırayla arar:
1. Kendi klasörü
2. Arşiv kardeş klasörleri (`../03-gunluk-satir`, `../02-yazim-yardimcisi`)
3. Masaüstü düzeni (`Desktop\TARAMA`, `Desktop\PYTON PROJELER`)

Yani bu arşiv klasörü tek başına yeterlidir; ayrıca py'yi iki modülle aynı
klasöre koyarsan da çalışır.

Gereksinim: her iki aracın gereksinimleri (openpyxl, pywin32, numpy, pypdfium2, Pillow, pynput)
