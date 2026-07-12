# Zimmet Belgesi Otomasyon Sistemi

Personel için kişiselleştirilmiş **kişisel koruyucu donanım (KKD) zimmet belgesi** oluşturma, işaretleme ve yazdırma otomasyonu. ~700-800 personel, 5 renk kategorisi (gri, kırmızı, lacivert, mavi, turuncu) + ofis/saha alt grupları.

## Dosya yapısı

```
.
├── PROJE_OZETI.md              ← Detaylı dokümantasyon (yeni Claude bunu okur)
├── README.md
├── .gitignore                  ← KVKK uyumlu (tüm personel verisi hariç)
│
├── bot_alanlari_isaretle.py    ← (kök) Bot almış personellerin docx'inde BOT satırını gri yapar
├── eslesmeyen_benzerlik.py     ← (kök) Eşleşmeyen isimleri fuzzy match ile bulup onay alır
├── zimmet_olustur.py           ← (kök, fallback) Tek klasör için belge üretir
├── zimmet_yazdir.py            ← (kök, fallback) Tek klasör için yazdırır
│
├── <renk>/                     ← gri, kırmızı, lacivert, mavi, turuncu, yeşil, ofis
│   ├── KIŞLIK ZİMMET FORMU <RENK>.docx   ← template (ignore'sız, repo'da)
│   ├── <renk>ler.xlsx                    ← personel listesi (GITIGNORE — yerel)
│   ├── zimmet_olustur.py                 ← belge üretici (Excel'i okur, template'i doldurur)
│   ├── zimmet_yazdir.py                  ← yazdırıcı (Word + Windows spooler)
│   ├── bot_alanlari_isaretle.py
│   ├── OLUSTURULAN/                      ← üretilenler (GITIGNORE)
│   └── YAZDIRILDI/                       ← yazdırılanlar (GITIGNORE)
│
├── bakım-inspekte/             ← Ofis personeli alt-grupları
│   ├── BAKIM PERSONELLERİ/
│   ├── FMC PERSONELLERİ/
│   ├── SUBSEA PERSONELLERİ/
│   └── İNSPEKTE PERSONELLERİ/
│
├── geçici görevlendirmeler/    ← Geçici personel için ayrı yapı
├── tanker operatörü/           ← Saha tanker operatörü (mavi giysi + özel görev)
└── YAZLIK BAŞLIKLAR/           ← Yazlık kıyafet başlık/etiket üretimi (ayrı iş)
```

## Hassas veri uyarısı

Bu repo **kod ve template'ları** içerir. Gerçek personel verisi (TC, isim, adres) `.gitignore` ile dışlanmıştır:

- `*.xlsx`, `*.xlsm` (tüm personel listeleri)
- `OLUSTURULAN/`, `YAZDIRILDI/` (üretilen belgeler)
- `yazdirma_log.txt`, `eslesmeyenler.txt`

Çalıştırmak için kendi makinanda bu dosyaları manuel oluşturmalısın (USB/cloud ile aktar).

## Bağımlılıklar

```
pip install python-docx openpyxl pywin32
```

`pywin32` sadece yazdırma için gerekli (Windows-only).

## Detay

Tüm script'lerin ne yaptığı, sırası, gotchas, edge case'ler için: **[PROJE_OZETI.md](PROJE_OZETI.md)**
