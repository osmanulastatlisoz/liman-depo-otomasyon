# Yazım Yardımcısı (excel_yazim_yardimcisi.py)

![Ekran görüntüsü — demo veriler](../gorseller/02_yazim_yardimcisi.png)

Excel'e stok hareketi girerken hep üstte duran otomatik tamamlama penceresi.
Çalışan Excel'e COM ile bağlanır, aktif hücreyi 0,5 sn'de bir izler ve hangi
sütundaysan ona uygun öneri listesi gösterir. Excel tüm özellikleriyle açık kalır.

## Sütun tanıma
| Sayfa | Roller |
|---|---|
| STOK HAREKETLERİ (TCH) | A tarih, B depo, C işlem, D personel, E malzeme, F miktar, G açıklama |
| STOK GİRİŞ ÇIKIŞ (2026) | A işlem, B personel, C malzeme, D miktar, E açıklama, G tarih |

## Tuşlar
- ↑ ↓ listeden seç (ilk ↓ seçimi başlatır; en üstte ↑ seçimi bırakır)
- → / ← / Enter: yaz ve sağa/sola geç — **seçim yapılmadıysa yazılan AYNEN yazılır**
  (Excel'in kısayol makrolarıyla uyum için; "e" yazıp → basınca hücreye "e" girer)
- Ctrl+D: üstteki satırdakini bulunduğun hücreye kopyalar (Excel Ctrl+D'si, FillDown)
- Ctrl+N / ➕: YENİ personel — adı önce TCH PERSONEL LİSTESİ'nin sonuna ekler
  (koruma makrosu itiraz etmesin diye), sonra hücreye yazıp sağa geçer; ilk basış
  onay ister, ad zaten listedeyse eklemeden doğru yazımıyla yazar; öneriler yeni
  adı anında tanır
- Esc: kutuyu temizle · kutu boş + seçim yok: → ← sadece hücre değiştirir

## Akıllı listeler
- Malzeme yanında liman stoğu: `ELDİVEN (2056)` — DEPO sayfası LİMAN DEPO sütunu,
  ~10 sn'de bir açık Excel'den CANLI tazelenir; parantez hücreye yazılmaz
- Malzeme sıralaması stok çoktan aza; personel sıralaması hareket sayısı çoktan aza
- Tarih sütununda "5.7" → 05.07.2026 (gerçek tarih olarak yazılır)
- Kaynak listeler: TCH dosyasının PERSONEL LİSTESİ / MALZEME LİSTESİ / LISTELER
  sayfaları (⟳ düğmesi yeniler)

## Not
Office etkin değilse Excel'i COM ile BAŞLATMAK sihirbaz dialoguna takılır;
bu araç zaten AÇIK Excel'e bağlandığı için sorun yaşamaz.

TCH dosyasının yolu esnektir: önce kayıtlı ayar, sonra masaüstü; bulunamazsa
ilk çalıştırmada seçtirir (`%LOCALAPPDATA%\DepoAraclari\yollar.json`).

Gereksinim: `openpyxl`, `pywin32`
