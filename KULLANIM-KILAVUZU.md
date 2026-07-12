# DEPO ARAÇLARI KILAVUZU

Bu klasördeki (ve masaüstündeki) Python araçlarının kullanım kılavuzu.
Son güncelleme: 08.07.2026

## Hızlı Bakış

| Araç | Ne işe yarar | Nasıl açılır |
|---|---|---|
| **Personel Sorgu** | Personelin aldığı malzemeleri anında gösterir | Masaüstü → `PERSONEL SORGU.bat` |
| **Yazım Yardımcısı** | Excel'e veri girerken personel/malzeme otomatik tamamlama | Masaüstü → `YAZIM YARDIMCISI.bat` |
| **Günlük Satır** | Taranan PDF'i satır satır şerit halinde gösterir | `TARAMA\gunluk_satir.py` |
| **Günlük Satır + Yardımcı** | Üstteki ikisini birlikte çalıştırır | Masaüstü → `GÜNLÜK SATIR + YARDIMCI.bat` |

---

## 1. PERSONEL SORGU (`Desktop\birlesik_stok_sorgu.py`)

**Amaç:** Excel açmadan, personel adı yazınca aldığı malzemeleri anında görmek.
Excel bilmeyen biri de rahat kullansın diye yapıldı.

**Veri kaynağı:** İki dosyanın birleşimi (mükerrerler teke indirilmiş):
- `2026 DEPO STOK.xlsm` → STOK GİRİŞ ÇIKIŞ sayfası
- `TCH_STOK_TAKIP_CALISMASI GÜNCEL.xlsm` → STOK HAREKETLERİ sayfası

**Kullanım:**
- Pencere anında açılır (önbellekten), arka planda yeni kayıtları çekip
  hem ekranı hem masaüstündeki formülsüz **BİRLEŞİK STOK HAREKETLERİ.xlsx**
  dosyasını sessizce günceller.
- **Personel Adı** ve/veya **Malzeme Adı** kutusuna yaz (en az 2 harf).
  Kelime sırası önemsiz: "yılmaz ahmet" de "ahmet yılmaz" da bulur.
  Yanlış yazarsan "Şunu mu demek istediniz?" diye önerir.
- **İşlem** filtresi varsayılan ÇIKIŞ (personelin aldıkları). TÜMÜ/GİRİŞ seçilebilir.
- **Görünüm → Malzeme toplamları**: kişinin her malzemeden toplam kaç adet aldığı.
- **Excel'i Aç** düğmesi birleşik dosyayı açar. Esc kutuları temizler.

**Bilinmesi gerekenler:**
- Kaynak dosyaların **kayıtlı** (Ctrl+S yapılmış) hâli okunur.
- Birleşik Excel dosyası açıkken güncelleme o tura atlanır (alt çubukta uyarı çıkar);
  bir sonraki açılışta dosya da güncellenir.
- Mükerrer ayıklama anahtarı: tarih + işlem + personel + malzeme + miktar.

---

## 2. YAZIM YARDIMCISI (`PYTON PROJELER\excel_yazim_yardimcisi.py`)

**Amaç:** Excel'de stok hareketi girerken hep üstte duran küçük pencereden
yazmak — personel/malzeme adları yazdıkça süzülür, tek tuşla hücreye geçer.
Excel tüm özellikleriyle açık kalır.

**Sütunları kendisi tanır** (aktif hücre neredeyse ona göre):

| Sayfa | Sütun → Rol |
|---|---|
| STOK HAREKETLERİ (TCH) | A tarih, B depo, C işlem, D personel, E malzeme, F miktar, G açıklama |
| STOK GİRİŞ ÇIKIŞ (2026) | A işlem, B personel, C malzeme, D miktar, E açıklama, G tarih |

**Tuşlar:**

| Tuş | Ne yapar |
|---|---|
| ↑ ↓ | Listeden seç (ilk ↓ seçimi başlatır; en üstteyken ↑ seçimi bırakır) |
| → | Yazar ve SAĞDAKİ hücreye geçer |
| ← | Yazar ve SOLDAKİ hücreye geçer |
| Enter | Yazar ve sağa geçer |
| Ctrl+D | Üstteki satırdakini bulunduğun hücreye kopyalar (Excel'deki gibi) |
| Ctrl+N / ➕ | YENİ personel: adı önce PERSONEL LİSTESİ'ne ekler, sonra hücreye yazar (ilk basış onay ister) |
| Esc | Kutuyu temizler |

**Altın kural:** Listeden ↓ ile SEÇMEDİĞİN sürece yazdığın metin AYNEN yazılır
("e" yazıp → basarsan hücreye "e" girer — Excel kısayol makroların bozulmaz).
Kutu boş VE seçim yoksa → ← sadece hücre değiştirir (yazmadan atlama).
Kutu boş ama ↓ ile seçim yaptıysan seçilen yazılır (İŞLEM/DEPO gibi kısa listelerde pratik).

**Akıllı listeler:**
- **Malzeme** önerileri yanında liman deposunda kalan miktar parantezle görünür:
  `ELDİVEN    (2056)`. Kaynak: DEPO sayfası, LİMAN DEPO sütunu.
  Sayılar ~10 saniyede bir açık Excel'den CANLI tazelenir.
  Sıralama: stoğu çok olan üstte. Parantez hücreye YAZILMAZ, sadece görüntüdedir.
- **Personel** önerileri hareket sayısına göre sıralı: en çok malzeme alan en üstte.
- **İşlem / Depo** sütunlarında liste yazmadan hazır gelir (GİRİŞ/ÇIKIŞ/TRANSFER/İADE,
  LİMAN/GARAJ) — ↓ ile seç, → ile yaz.
- Tarih sütununda "5.7" yazsan bile gerçek tarih (05.07.2026) olarak yazılır.
- **⟳** düğmesi listeleri ve sayıları dosyadan yeniden okur
  (PERSONEL LİSTESİ / MALZEME LİSTESİ sayfalarına yeni kayıt eklediysen bas).

---

## 3. GÜNLÜK SATIR (`TARAMA\gunluk_satir.py`)

**Amaç:** Taranan "Günlük Verilen Malzeme" PDF'ini ekranın üstünde ince bir
şerit olarak satır satır göstermek. Fokus çalmaz; Excel'de çalışmaya devam edersin.

**Tuşlar (globaldir, Excel'deyken de çalışır):**

| Tuş | Ne yapar |
|---|---|
| Sağ Shift | Sıradaki PDF satırı (+ Excel'de imleç makrosu) |
| Sol Shift + Sağ Shift | Sıradaki satır ama Excel'e DOKUNMAZ |
| Sağ Ctrl | Önceki satır |

**Şerit düğmeleri:** ⌨ Ayarla (Excel'e gidecek tuş dizisini kaydet) • Excel ✓/✕
(imleç hareketini aç/kapat) • ⤓ Git (sayfa+satıra atla) • ✔ Bitti (dosya adına
"(işlendi)" ekler) • ✕ çıkış. Bilgi çubuğunu sürükleyerek şeridi taşıyabilirsin.

Kaldığın satır PDF başına hatırlanır; aynı PDF'i tekrar açınca sorar.
Eğik taramalar otomatik düzeltilir. Ayarlar dosyanın başındaki AYARLAR bloğunda.

---

## 4. GÜNLÜK SATIR + YARDIMCI (`PYTON PROJELER\gunluk_yazim_birlesik.py`)

**Amaç:** Yukarıdaki iki aracı tek pencere düzeninde, çakışmadan birlikte çalıştırmak.
Orijinal iki dosyaya dokunmaz; onları içe aktarır (ayar değişikliklerin buraya da yansır).

**Akış:**
1. Çift tıkla → PDF'i seç → şerit üstte kurulur, yardımcı penceresi altında açılır.
2. Yardımcı kutusundan satırın hücrelerini doldur (→ ile ilerle).
3. Satır bitince **Sağ Shift**: kutuda ne varsa (veya ↓ ile seçtiğin) hücreye
   YAZILIR + Excel imleci makro kadar taşınır (örn. 1 aşağı 2 sol) + PDF şeridi
   sıradaki satıra geçer + kutu temizlenir. **Dört iş tek tuşta.**

**Birleşik sürüme özel farklar:**
- Sağ Shift Excel'e tuş GÖNDERMEZ; hücreyi COM ile doğrudan taşır.
  Bu yüzden odak hangi penceredeyse fark etmez, ok tuşları asla karışmaz.
- Sol Shift + Sağ Shift yine Excel'e dokunmadan şeridi ilerletir.
- Şeritteki **✎ Yardımcı** düğmesi yardımcı penceresini gizler/gösterir.
- Yardımcının tüm özellikleri (stok parantezi, sıralamalar, Ctrl+D…) aynen geçerli.

---

## Sorun Giderme

| Belirti | Çözüm |
|---|---|
| "Excel'e bağlı değil" | Excel'i ve TCH dosyasını aç; pencere kendiliğinden bağlanır. |
| "Yazılamadı — Excel meşgul" | Excel'de bir hücre düzenleme modunda kalmış olabilir; Esc'e bas. |
| Stok parantezleri eski görünüyor | TCH dosyası Excel'de açıksa 10 sn bekle ya da ⟳'ye bas. |
| Sorguda "Kayıt bulunamadı" | Yazımı kontrol et; öneri verir. İşlem filtresini TÜMÜ yap. |
| Bat çift tıklayınca bir şey olmuyor | py dosyası taşınmışsa bat içindeki yol güncellenmedir. |
| Birleşik Excel güncellenmedi | Dosya Excel'de açıktı; kapat, sorguyu tekrar aç. |

## Dosya Haritası

```
Desktop\
├─ PERSONEL SORGU.bat            → birlesik_stok_sorgu.py
├─ YAZIM YARDIMCISI.bat          → PYTON PROJELER\excel_yazim_yardimcisi.py
├─ GÜNLÜK SATIR + YARDIMCI.bat   → PYTON PROJELER\gunluk_yazim_birlesik.py
├─ birlesik_stok_sorgu.py
├─ BİRLEŞİK STOK HAREKETLERİ.xlsx  (formülsüz birleşik kütük — otomatik güncellenir)
├─ TARAMA\gunluk_satir.py          (+ gunluk_satir_durum.json: kaldığın yer kaydı)
└─ PYTON PROJELER\
   ├─ excel_yazim_yardimcisi.py
   ├─ gunluk_yazim_birlesik.py
   └─ ARAÇLAR KILAVUZU.md          (bu dosya)

Önbellek: %LOCALAPPDATA%\BirlesikStok\birlesik_cache.pkl (sorgu aracının hızlı açılışı)
```
