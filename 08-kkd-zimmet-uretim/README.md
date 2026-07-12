# KKD Zimmet Belge Üretimi

![Ekran görüntüsü — demo veriler](../gorseller/08_zimmet_uretim.png)

~700-800 personel için kişiselleştirilmiş Word zimmet belgesi üretme,
işaretleme ve teyitli toplu yazdırma hattı. Personel renk kategorilerine
ayrılıdır (gri, kırmızı, lacivert, mavi, turuncu, yeşil, ofis); her kategorinin
kendi Excel listesi ve Word şablonu vardır.

Detaylı mimari ve akış: `PROJE_OZETI.md` (projenin kendi dokümantasyonu),
`PROJE_README.md` (repo README'si). `ZIMMET_FORM_METNI.txt` form metni örneğidir.

## Araçlar
- **zimmet_olustur.py** — klasördeki tek .xlsx (AD, TC, görev) + tek .docx
  şablonundan `OLUSTURULAN/AD SOYAD.docx` belgeleri üretir
- **zimmet_yazdir.py** — belgeleri varsayılan yazıcıya sırayla gönderir;
  yazıcı KUYRUĞUNDAN tamamlandığını teyit eder, başarılıyı YAZDIRILDI'ya
  taşır, hata olursa DURUR, log tutar (kesintiden kaldığı yerden devam)
- **bot_alanlari_isaretle.py** — bot almış personellerin belgelerinde BOT
  satırının arka planını gri yapar (idempotent, tekrar çalıştırılabilir)
- **eslesmeyen_benzerlik.py** — listeyle eşleşmeyen isimleri fuzzy match ile
  bulur, numaralı onay akışıyla işaretler

## Not
Personel verileri (xlsx listeler, üretilen belgeler) KVKK gereği bu arşivde
YOKTUR; asıl proje klasörü `Desktop\kkd yazlık projesi\` (git repo'dur).

pip: `python-docx`, `openpyxl`, `pywin32`
