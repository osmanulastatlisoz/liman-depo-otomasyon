# Zimmet Belgesi Otomasyon — Proje Özeti

> Yeni Claude'a not: Bu dosya tüm projenin contextini içerir. Önce okuyup sonra çalış.

## Amaç

~700-800 personel için kişiselleştirilmiş **Kişisel Koruyucu Donanım (KKD) zimmet belgesi** (Word) oluştur ve yazdır.

- Personel renk kategorisine ayrılmış: **gri, kırmızı, lacivert, mavi, turuncu, yeşil** (saha) + **ofis** + alt gruplar
- Her kategori için ayrı Excel (personel listesi) + ayrı Word template var, yapıları aynı
- Excel: AD-SOYAD, TC, MESLEK kolonlarından oluşur
- Template: AD, TC, GÖREV placeholder'larını içerir; ayrıca KKD tablosu (pantolon, sweat/tişört, **bot** satırları)

## Klasör yapısı

```
kök/
├── PROJE_OZETI.md, README.md, .gitignore
├── 2026 BOT ALMIŞ PERSONEL LİSTESİ.xlsx     [GITIGNORE]
├── bot_alanlari_isaretle.py                 ← tüm renk klasörlerini gezer
├── eslesmeyen_benzerlik.py                  ← fuzzy match + manuel onay
├── zimmet_olustur.py, zimmet_yazdir.py      ← fallback (tek klasör için)
│
├── <renk>/                                  ← gri, kırmızı, lacivert, mavi, turuncu, yeşil, ofis
│   ├── KIŞLIK ZİMMET FORMU <RENK>.docx     ← template
│   ├── <renk>ler.xlsx                      ← personel listesi [GITIGNORE]
│   ├── zimmet_olustur.py                   ← belge üretici
│   ├── zimmet_yazdir.py                    ← yazdırıcı
│   ├── bot_alanlari_isaretle.py            ← (renk klasöründe de var, kök kopyası)
│   ├── OLUSTURULAN/  [GITIGNORE]
│   ├── YAZDIRILDI/   [GITIGNORE]
│   └── yazdirma_log.txt  [GITIGNORE]
│
├── bakım-inspekte/                          ← Ofis alt-grupları
│   ├── BAKIM PERSONELLERİ/, FMC PERSONELLERİ/, SUBSEA PERSONELLERİ/, İNSPEKTE PERSONELLERİ/
│   └── (her biri kendi xlsx + ofis template + scriptler)
│
├── geçici görevlendirmeler(operatör kırmızı)/   ← kırmızı/mavi/ofis/şoför geçici
├── tanker operatörü(saha personeli)/             ← Tehlikeli Atık Temizleme Operatörü (mavi giysi)
└── YAZLIK BAŞLIKLAR/                            ← Kıyafet stok başlık/etiket üretimi (ayrı iş)
    ├── BAŞLIKLAR.xlsx                          ← malzeme + beden listesi
    ├── ÖRNEK BAŞLIK.docx                       ← görünüm örneği (yatay A4, kalın Calibri 100pt)
    ├── baslik_olustur.py                       ← her satır için 1 yatay başlık üretir
    ├── baslik_yazdir.py                        ← TERS sırada yazdırır (Excel sırasının tersi)
    ├── OLUSTURULAN/  [GITIGNORE]
    └── YAZDIRILDI/   [GITIGNORE]
```

## Excel formatı (personel listesi)

1. satır başlık, atlanır. 3 kolon:
- **A:** AD-SOYAD (ör: `AHMET ÖZBAY`)
- **B:** TC (11 hane)
- **C:** MESLEK KODU - MESLEĞİ (ör: `Operatör`, `BAKIM PERSONELİ`, `Tehlikeli Atık Temizleme Operatörü (Tank/Tanker)`)

## Word template formatı

3 placeholder:
```
ÇALIŞANIN ADI SOYADI :  [çok boşluk]  TARİH : ……/……./2026   ← tek paragraf
TC KİMLİK NUMARASI :                                          ← tek paragraf,
GÖREVİ :                                                        içinde soft line break (\n)
```

Bunlara ek olarak tabloda KKD listesi (pantolon, sweat, BOT satırı vs.) var.

---

## Script'ler

### 1. `zimmet_olustur.py` (her renk klasöründe)

Klasördeki tek `.xlsx` ve tek `.docx` (template) bulur, her personel için belge üretir.

- Excel'in C kolonu boşsa GÖREVİ alanı boş bırakılır
- **TC boşsa personel atlanır** (zimmet için TC zorunlu)
- Çıktı: `OLUSTURULAN/AD SOYAD.docx`

Önemli teknik detaylar:
1. **Orantılı font kompansasyonu:** Ad eklendikten sonra TARİH satırın altına kaymasın diye, eklenen karakterin **2-3 katı** kadar boşluk silinir (`compensate=True`, sadece ad alanı için)
2. **Aynı paragrafta birden fazla placeholder:** TC KİMLİK ve GÖREVİ aynı paragrafta, `\n` ile ayrılmış. "İlk `:`" aramak çalışmaz — anahtar kelimeden sonra gelen `:` aranmalı
3. **Run bölünmesi:** Word template'inde "ÇALIŞANIN ADI SOYADI" tek run değil, bölünmüş olabilir. Tüm run'ları birleştirip pozisyon hesapla, sonra doğru run'a yaz
4. **Türkçe karakterler:** Dosya adlarında yasak karakterler (`<>:"/\|?*`) temizleniyor

### 2. `zimmet_yazdir.py` (her renk klasöründe)

OLUSTURULAN'daki belgeleri varsayılan Windows yazıcısına gönderir.

**Mevcut özellikler:**
- **Türkçe alfabetik sıra:** Ş, Ö, Ç gibi harflerle başlayanlar Z'den sonra değil, doğru yerlerinde sıralanır (`tr_sort_key`, modül başında)
- **Batch prompt:** Başlangıçta "Kaç tane basılsın?" sorar, Enter = hepsi, sayı = ilk N (sonra script kapanır, kullanıcı tekrar çalıştırınca kaldığı yerden devam)
- **`PAUSE_EVERY` sabiti** (default 0/kapalı): N belgede bir duraklayıp Enter bekler. Batch yaklaşımı tercih edildiği için default kapalı
- **Spooler poll:** Word'ün `PrintOut` dönmesi yetmez — spooler kuyruğundan düştüğünde tamamlandı say
- **Ctrl+C ile her an temiz çıkış:** Kalan belgeler OLUSTURULAN'da kalır, tekrar çalıştırınca kaldığı yerden devam (deduplicate: YAZDIRILDI'da olan dosyaları OLUSTURULAN'dan siler)
- **Yazdırılan belge YAZDIRILDI'ya taşınır**
- **Hata olursa durur** (kağıt bitti, offline, timeout 120s); log'a yazar

**Sabitler:**
- `TIMEOUT_PER_DOC = 120` s (Wi-Fi yazıcıda yavaşsa 180-240 yap)
- `PAUSE_EVERY = 0` (kapalı, batch prompt tercih edildi)

### 3. `bot_alanlari_isaretle.py` (kökte + her renk klasöründe)

Kök klasördeki `2026 BOT ALMIŞ PERSONEL LİSTESİ.xlsx` dosyasında adı geçen personellerin tüm renk klasörlerindeki `OLUSTURULAN/*.docx` belgelerinde, KKD tablosundaki **STL-9040-S3-BOT-STARLİNE** satırının arka planını gri (#BFBFBF) yapar.

- **Tüm satır 5 hücresinin tamamı gri** (sadece bot metni değil)
- Yazı normal kalır (eskiden gri+çizgi yapıyordu, kullanıcı isteğiyle değiştirildi)
- **Idempotent:** Tekrar çalıştırınca zaten gri olanları atlar
- **Eşleşmeyenler dosyaya yazılır:** Bot listesinde olup hiçbir renk klasöründe docx'i olmayanlar `kök/eslesmeyenler.txt`'ye kaydedilir
- **Türkçe-doğru isim normalize:** `tr_upper` (i→İ, ı→I) + boşluk topla

**Tarama mantığı:** Sadece `BASE_DIR.iterdir()` yapıyor — yani sadece kökteki klasörler içindeki OLUSTURULAN'a bakıyor. `bakım-inspekte/<altklasör>/OLUSTURULAN` gibi 2 seviye derin yapıları otomatik bulmaz.

### 4. `eslesmeyen_benzerlik.py` (kökte)

`eslesmeyenler.txt` deki her ismi tüm OLUSTURULAN dosya adlarıyla **fuzzy karşılaştırır** (`difflib.SequenceMatcher`, ASCII-fold dahil).

- Kullanıcı eşik girer (örn: 80), önerileri numaralı listeler
- Kullanıcı onayladığı numaraların docx'lerine BOT işaretini uygular
- Onaylananları `eslesmeyenler.txt`'den çıkarır
- Sonda Excel düzeltme önerisi listeler (kalıcı düzeltme için)
- `bot_alanlari_isaretle.py` mantığını import eder, davranış tutarlı

---

## Tipik Akış (her renk için)

1. Renk klasörüne git: `cd gri` (veya kırmızı, mavi vs.)
2. `python zimmet_olustur.py` → `OLUSTURULAN/` doldurulur
3. Bir-iki belgeyi görsel kontrol et (hizalama doğru mu)
4. (Opsiyonel) Kök klasöre dön: `python bot_alanlari_isaretle.py` → bot almış olanların satırı gri olur
5. (Opsiyonel) `python eslesmeyen_benzerlik.py` → yazım farklı isimleri eşleştir
6. Renk klasörüne dön: `python zimmet_yazdir.py` → batch sayısı sor + yazdır

Wi-Fi yazıcı kullanılıyorsa: varsayılan yazıcıyı değiştir, "Windows varsayılan yazıcımı yönetsin" ayarını **kapat**.

---

## Çözülmüş Problemler (bunlara tekrar girme)

- ❌ Ad uzayınca TARİH alt satıra kayıyordu → ✅ orantılı boşluk silme (2-3x katı)
- ❌ GÖREVİ alanı bulunamıyordu (aynı paragrafta TC ile birlikte) → ✅ multi-placeholder desteği
- ❌ Yazdırma anında klasör taşıma yapıyordu (Word return ettiği için) → ✅ spooler kuyruğunu poll et
- ❌ Türkçe Ş, Ö, Ç ile başlayanlar Z'den sonra geliyordu → ✅ `tr_sort_key` ile özel sıralama
- ❌ Bot işaretleme önce gri+strike yapıyordu, çok belirgindi → ✅ satır arka planı gri (#BFBFBF), yazı normal
- ❌ Bot listesinde isim var ama docx yok (yazım farkı) → ✅ `eslesmeyen_benzerlik.py` fuzzy + manuel onay
- ❌ Her N belgede duraklasın istiyordu ama edit edince çalıştırma anında değişmiyor → ✅ batch prompt (her seferinde "kaç tane?" sor, kaldığı yerden devam)

---

## TC Lookup (Eksik TC'leri master'dan doldurma)

Bazı alt klasör Excel'lerinde (özellikle `bakım-inspekte/*/`) TC kolonu boş gelebilir. Çözüm: master personel listesinden lookup.

**Master:** yerel `…\ZİMMET BİLGİLERİ\GÜNCEL PERSONEL LİSTESİ.xlsx` (gerçek yol kaldırıldı; ~1245 benzersiz isim, 6+ sheet)

**Sheet konfigleri** (header satırı, name kolonu, TC kolonu):
- `GÜNCEL PERSONELLER`: (0, 1, 3)
- `ENGELLİ PERSONELLER`: (0, 1, 2)
- `AYRILAN PERSONELLER`: (0, 1, 2)
- `ÇEVRE BİRİMİ`: (0, 1, 2)
- `İDARİ KADRO`: (1, 2, 3)
- `SIKINTILI AYRILANLAR`: (1, 0, 2)
- `GEÇİCİ GÖREVLENDİRMELER`: standart yapıda değil, manuel parse gerekir

**Akış:** Master'dan `name → TC` dict kur (normalize: `tr_upper` + boşluk topla). Eksik xlsx'leri gez, eşleşenleri doldur. Eşleşmeyenler için fuzzy match dene (≥85% güvenli, ≥75% manuel kontrol). Hiçbir şey bulunmazsa kullanıcı manuel ekler.

---

## Tanker Operatörü İş Akışı (özel grup)

Bazı saha personeli `Tehlikeli Atık Temizleme Operatörü (Tank/Tanker)` görevinde, **mavi** kıyafet alır ama kırmızı listede bulunmuş olabilir. Süreç:

1. `TANKER OPERATÖRÜ.xlsx` listesini al (Desktop'ta tutuluyor)
2. Bu isimleri kırmızı/OLUSTURULAN veya YAZDIRILDI'dan bul
3. `tanker operatörü(saha personeli)/` klasörü oluştur, **mavi template**'i kopyala
4. Mavi template'inden yeni docx üret (görev = `Tehlikeli Atık Temizleme Operatörü (Tank/Tanker)`)
5. kırmızılar.xlsx'ten bu isimleri çıkar
6. maviler.xlsx'e ekle (yeni görev ile)
7. `bot_alanlari_isaretle.py` koştur (yeni docx'lere bot işareti)

---

## Yazlık Başlık/Etiket Üretimi (ayrı iş — `YAZLIK BAŞLIKLAR/`)

Zimmet işinden bağımsız. Amaç: kıyafet stoğu için **her malzeme+beden kombinasyonuna ayrı bir yatay A4 etiket (başlık)** üretip yazdırmak (depo rafına/kutusuna asılır).

**Excel formatı** (`BAŞLIKLAR.xlsx`, ilk sayfa): 1. satır başlık, atlanır. 2 kolon:
- **A:** MALZEME ADI (ör: `MAVİ TİŞÖRT`, `LACİVERT PANTOLON`)
- **B:** BEDEN (`S/M/L/XL/2XL…` veya `48/50/52…` sayı)

**Görünüm** (`ÖRNEK BAŞLIK.docx`'ten alınır): yatay/landscape A4, 2.5 cm kenar, **ortalanmış kalın Calibri**, 1. satır malzeme / 2. satır beden.

### `baslik_olustur.py`

Excel'deki her satır için `OLUSTURULAN/`'a 1 Word belgesi üretir.

- **Tek sayfaya sığma GARANTİSİ:** yazı boyutu otomatik ayarlanır — üst sınır örnekteki **100pt**, malzeme adı tek satıra sığmıyorsa sığana kadar küçültülür (ör. `TURUNCU PANTOLON` → 72pt, `LACİVERT PANTOLON` → 74pt; kısa isimler 100pt kalır). Beden de aynı boyutta. İki satır sayfaya **dikey ortalanır** (`sectPr/w:vAlign=center`).
- **Font genişliği ölçümü:** PIL (`ImageFont.getlength`) + `calibrib.ttf` ile gerçek metrik; PIL/font yoksa kaba heuristik fallback (büyük harf ~0.62em, boşluk ~0.28em).
- **Dosya adı:** `001 - MALZEME - BEDEN.docx` … `097 - …`. Baştaki **sıfır dolgulu numara Excel sırasını korur ve YALNIZCA dosya adındadır — basılan sayfada GÖRÜNMEZ.** Yazdırma sırasını belirlemek için kullanılır.
- Geometri/font/üst sınır örnek docx'ten okunur; örnek yoksa A4-landscape varsayılanlarına düşer.

### `baslik_yazdir.py`

`zimmet_yazdir.py`'nin **birebir aynısı** (spooler poll, resume/dedup, batch prompt, Ctrl+C, log) — **TEK farkı sıralama:**
- Türkçe alfabetik (`tr_sort_key`) yerine **dosya adına göre AZALAN** sıra → `097…` ilk, `001…` son basılır = **Excel sırasının tersi**.
- **Neden ters:** çıktı yazıcıdan yüz üstü çıkıp üst üste birikince, en son basılan (001) en üstte kalır → deste Excel sırasında okunur.

### Çözülmüş problem
- ❌ Örnek `ÖRNEK BAŞLIK.docx` 100pt'de **2 sayfaya taşıyordu** (uzun isim tek satıra sığmıyor + sonda 2 adet 120pt boş paragraf) → ✅ otomatik font fit + dikey ortalama + boş paragraf yok ⇒ 97/97 belge tek sayfa (Word COM ile `ComputeStatistics` doğrulandı).

### Akış
```
cd "YAZLIK BAŞLIKLAR"
python baslik_olustur.py     # OLUSTURULAN/ -> 97 başlık
python baslik_yazdir.py      # ters sırada yazdır (pywin32 gerekli)
```

---

## Bağımlılıklar

```
pip install python-docx openpyxl pywin32
pip install Pillow            # opsiyonel — baslik_olustur.py'de font genişliği ölçümü için
```

`pywin32` sadece yazdırma için (Windows-only). Diğer scriptler Linux/Mac'te de çalışır.
`Pillow` kuruluysa başlık yazı boyutu birebir doğru ölçülür; yoksa kaba heuristiğe düşer (yine de güvenli çalışır).

---

## Kullanışlı snippet'ler

**Türkçe-doğru büyük harf + isim normalize:**
```python
def tr_upper(s):
    return s.translate(str.maketrans('iı', 'İI')).upper()

def normalize(s):
    return tr_upper(' '.join(str(s).split())).strip()
```

**Türkçe sıralama anahtarı:**
```python
TR_ALPHABET = "ABCÇDEFGĞHIİJKLMNOÖPQRSŞTUÜVWXYZ"
TR_ORDER = {c: i for i, c in enumerate(TR_ALPHABET, start=100)}

def tr_sort_key(path):
    s = path.name if hasattr(path, 'name') else str(path)
    return tuple(TR_ORDER.get(c, ord(c)) for c in tr_upper(s))
```

**Tablo hücresinin arka planını gri yap:**
```python
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def set_cell_shading(cell, color_hex):
    tc_pr = cell._tc.get_or_add_tcPr()
    for old in tc_pr.findall(qn('w:shd')):
        tc_pr.remove(old)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tc_pr.append(shd)
```

**Path Türkçe karakter sorunu (Windows):** `openpyxl.load_workbook` direkt string path ile bazen `FileNotFoundError` veriyor (NFC/NFD Unicode normalize farkı). Çözüm: `os.listdir(folder)` ile dosyayı bulup `os.path.join` ile yolu kur.
