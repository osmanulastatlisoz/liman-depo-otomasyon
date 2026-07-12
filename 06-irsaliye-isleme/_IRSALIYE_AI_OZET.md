# İRSALİYE / FATURA İŞLEME SİSTEMİ — AI İÇİN KOMPAKT ÖZET

> Bu dosya, gelen e-Fatura/irsaliyeleri işleyen kurulumu ve iş akışını **bağlamsız bir
> yapay zekânın bile** anlayıp devam edebilmesi için yazıldı.

## 1. Amaç
Kullanıcı (liman deposu — depo/satınalma) gelen faturaları/irsaliyeleri tarıyor; iki çıktı isteniyor:
1. **Malzeme listesi Excel'i** (hangi faturada hangi malzeme, kaç adet).
2. PDF'lerin **fatura/irsaliye numarasıyla yeniden adlandırılması** (`IMG_0001.pdf` → `KTC2026000000179.pdf`).

Faturalar genelde **satıcı = KTC Güç Sistemleri**, **alıcı = şirket**. Bazen başka satıcı olur (ör. **Sönmez Filtre**). Hepsi **taranmış görüntü**.

## 2. Klasör düzeni
- `TARAMA\<GG.AA.YYYY> <başlık>` alt klasörleri. İrsaliyeler: adı **"… GELEN İRSALİYELERİ"** ya da **"… İRSALİYELER"** olan klasörlerde.
- Tarama dosyaları ya **`IMG_xxxx.pdf`** (telefon/çok-sayfa tarayıcı) ya da **NAPS2** ile **`ad 1.pdf`, `ad 2.pdf`** biçiminde gelir.

## 3. Kurulu araçlar (Windows)
- **Python 3.13** + paketler: `pypdf`, `pypdfium2`, `pillow`, `openpyxl`. (DİKKAT: `pdfplumber`, `pytesseract` **kurulu değil**.)
- **Tesseract 5.5** — diller: `tur`, `eng`. Yol: `C:\Program Files\Tesseract-OCR\tesseract.exe` (**PATH'te değil**, tam yolu kullan; bash'te `export PATH="$PATH:/c/Program Files/Tesseract-OCR"`).
- **NAPS2** (tarayıcı yazılımı): OCR=Türkçe, "aranabilir PDF", "Tekli sayfa dosyaları" **açık**. **LibreOffice** kurulu.
- Tarayıcı: **Canon GX7100** (A4, **300 dpi**, 24-bit Renk, ADF/Besleyici).

## 4. KRİTİK teknik gerçekler
- **Eski `IMG_xxxx.pdf`'lerde metin katmanı YOK** → `pypdf.extract_text()` = 0. Çözüm: `pypdfium2` ile sayfayı **scale 3–6** render edip görüntüden oku (AI vision) veya **Tesseract OCR**. Online OCR sitesine firma faturası (vergi no) **yükleme** — gizlilik.
- **NAPS2 OCR açıkken** yeni taramalar **metin katmanlı** gelir → `pypdf` ile metni bedavaya alırsın, AI resim okumaz → **token çok düşer**. Tercih edilen yol budur.
- **KTC e-Fatura yerleşimi** (kırpma oranları; `h`=sayfa yüksekliği, `w`=genişlik):
  - Sağ-üst bilgi kutusu (**FATURA NO + FATURA TARİHİ**) ≈ `y 0.25–0.41`, `x 0.55–1.0`.
  - Malzeme tablosu ≈ `y 0.39–0.64`, tüm genişlik.
  - **Sönmez Filtre** farklı tablo (Sıra No / İskonto kolonlu), numara yerleşimi de farklı.
- **Fatura no formatı:** 3 karakter + 13 hane → `KTC2026000000179`, `FT02026000000103`. Faturada **"İRSALİYE YERİNE GEÇER"** yazar → ayrı irsaliye no yok, **fatura no kullan**.
- **TUZAK:** Sayfa altındaki **"Fatura Açıklaması"** içinde `TEF…/MCR…/YDI…/OFF…/SFR…/TPR…` gibi BAŞKA numaralar var; bunlar gerçek fatura no **DEĞİL**. Numara ararken **yalnız sağ-üst kutuya** bak.
- **Doğrulama:** her satırda `birim fiyat × miktar = satır tutarı` ve KDV ile çapraz kontrol et (okuma hatasını yakalar).
- **`recalc.py`** (xlsx skill) Windows'ta **çalışmaz** (AF_UNIX yok). Excel'de formül yerine **hesaplanmış değer** yaz; ya da kullanıcı Excel'de açınca hesaplar.

## 5. Yardımcı script: `irsaliye_isle.py` (+ `irsaliye_isle.bat`)
**Ne yapar:** Bir klasördeki her PDF için → metin katmanı varsa onu, yoksa **sağ-üst kutuyu** Tesseract OCR'lar → **Fatura/İrsaliye No + Tarih** bulur → `<No>.pdf` olarak adlandırır (çakışma korumalı) → şu iki dosyayı döker:
- `irsaliye_ozet.xlsx` — sütunlar **Dosya / Fatura-İrsaliye No / Tarih / Durum**. Numarası bulunamayan satır **"⚠ ELLE BAK"** yazılır ve **sarıya** boyanır.
- `irsaliye_metinleri.txt` — her PDF'in çıkarılan tam metni (AI'nın malzeme Excel'ini **ucuza** üretmesi için).

**Çalıştırma:**
- Kolay: **`irsaliye_isle.bat`** çift tıkla → klasör seçme penceresi → "Uygula?" **Evet/Hayır**.
- Komut: `python irsaliye_isle.py "KLASOR"` (kuru deneme) · `… "KLASOR" --uygula` (gerçek adlandırma).

**Güvenlik:** Numarayı güvenle bulamazsa **UYDURMAZ**, "ELLE BAK" işaretler. Varsayılan kuru denemedir (dosya değiştirmez).

## 6. İstenen ÇIKTI: "Malzeme Listesi" Excel formatı
Sütunlar (önceki teslimlerle aynı tutulmalı):
`Orijinal Dosya | Fatura/İrsaliye No | Tarih | Firma | Malzeme Adı/Kod | Miktar | Birim | Birim Fiyat (TL) | Tutar KDV Hariç (TL) | Not`
- Font **Arial**, başlık **koyu mavi (1F4E78) + beyaz**, **freeze panes** + **autofilter**, `Tutar = Miktar × Birim Fiyat`.
- El yazısı notları "Not" sütununa al. Birim "Adet"/"kg" vb. faturadan.

## 7. AYLIK İŞ AKIŞI
1. **NAPS2 ile tara** (OCR Türkçe açık) → ayrı **metinli** PDF'ler bir tarih klasörüne.
2. **`irsaliye_isle.bat`** → klasörü seç → listeye bak.
3. **Sarı / "ELLE BAK"** satır varsa: o klasördeki **`irsaliye_metinleri.txt`**'i AI'ya ver, doğru numarayı buldur.
4. **Evet** de → PDF'ler fatura no'suyla adlanır.
5. **Malzeme Excel'i** için: **`irsaliye_metinleri.txt`**'i AI'ya ver → AI 6. bölümdeki formatta Excel'i üretir (resim değil metin okuduğu için ucuz).

## 8. BAŞKA BİR AI'YA TALİMAT (yeni klasör gelince)
- Önce PDF'lerde **metin katmanı var mı** bak (`pypdf`). Varsa metinden çalış (ucuz). Yoksa `pypdfium2` ile render et; kritik/küçük alanlarda **scale 5–6 kırp** ve oku (vision) veya Tesseract (`-l tur+eng`).
- **Fatura no**: yalnız sağ-üst kutudan al (alttaki açıklama numaralarını ALMA). **Malzeme**: tablodan ad+miktar+birim fiyat; `miktar×fiyat=tutar` ile doğrula.
- Excel'i 6. bölüm formatında üret. **Numarayı/sonucu UYDURMA** — emin değilsen kullanıcıya sor (dosya adına yazılan no birebir doğru olmalı).
- Çıktıları ilgili tarih klasörüne yaz; geçici PNG'leri temizle.

---
### Ek not — diğer scriptler (BU ÖZETİN KAPSAMI DIŞINDA, sadece bilgi)
TARAMA kökünde başka işlere ait scriptler de var, irsaliye sistemiyle ilgisiz:
- `gunluk_satir.py` (+ `.bat`, `gunluk_satir_durum.json`): "Günlük Verilen Malzeme" PDF'inde, Excel'e elle girerken ekranın üstünde **satır satır önizleme şeridi** gösterir.
- `kkd_bol.py`, `kkd_cikis.py`, `kkd_elektronik.py`: KKD (kişisel koruyucu donanım) / elektronik ile ilgili ayrı işler.
- `_test_*.py`, `_*.png`: gunluk_satir geliştirme artıkları.
