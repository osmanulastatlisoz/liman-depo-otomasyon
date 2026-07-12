# KKD Tarama Araçları (5 araç)

![Ekran görüntüsü — demo veriler](../gorseller/07_kkd_tarama.png)

Taranan KKD (kişisel koruyucu donanım) zimmet evraklarını bölme, personel
adıyla adlandırma ve Excel'e işleme zinciri. Tipik sıra: bol → adlandir → isle.

## kkd_bol.py
Tek büyük taranmış PDF'i her personel için 2'şer sayfalık dosyalara böler.
İlk sayfanın isim alanı pencerede gösterilir, ismi yazarsın, o adla kaydeder.

## kkd_adlandir.py
"ÇALIŞANIN ADI SOYADI:" satırından adı OKUR (pypdf, OCR'siz) ve dosyayı
`Ad Soyad.pdf` yapar. Uydurmaya karşı ÇAPRAZ KONTROL: aynı sayfadaki TC
kimlik no temiz 11 hane okunuyorsa metin katmanı güvenilir sayılır; değilse
"elle bak" penceresine düşer (tahmin kutuda hazır gelir). `--kuru` destekler.

## kkd_cikis.py / kkd_elektronik.py
İsimsiz (IMG_xxxx.pdf) taramaları tek tek göstererek elle hızlı adlandırma —
sırasıyla ayrılan personel iade listeleri ve elektronik cihaz zimmetleri için.

## kkd_isle.py
Adlandırılmış PDF'leri sırayla gezer; formun tablosu ekran üstü şeritte
gösterilir. Global tuşlar: Sağ Shift = personel ADINI Excel'e basar
(AD⏎AD⏎AD⏎→↑↑↑ dizisi), Sağ Ctrl = PDF'i İŞLENDİ klasörüne atıp sonrakine
geçer. Geri Al düğmesi son işleneni geri getirir.

pip: `pypdf`, `pypdfium2`, `Pillow`
