# YIKAMA TUTANAK ÜRETİM TALİMATI

## GÖREV
Kullanıcı kısaltma kodları ve miktarlar verir. Bunları çözümle, birleştir, sırala ve YIKAMA_TUTANAK_FORMATI.xlsx şablonuna yaz.

## KISALTMA SÖZLÜĞÜ (kod → tam ad)

### KAZAKLAR
KS/KM/KL/KXL/K2XL/K3XL/K4XL/K5XL → POLO YAKA SWEAT [S/M/L/XL/2XL/3XL/4XL/5XL] BEDEN

### POLARLAR
SPS/SPM/SPL/SPXL/SP2XL-SP5XL → SİYAH POLAR [beden] BEDEN
GPS/GPM/GPL/GPXL/GP2XL-GP5XL → GRİ POLAR [beden] BEDEN
HPS/HPM/HPL/HPXL/HP2XL-HP4XL → HAKİ POLAR [beden] BEDEN
KPS/KPM/KPL/KPXL/KP2XL → KUM POLAR [beden] BEDEN

### GÖMLEKLER
KGM/KGL/KGXL/KG2XL → KUM GÖMLEK [beden] BEDEN
GGM/GGL/GGXL/GG2XL → GRİ GÖMLEK [beden] BEDEN
HGS/HGXS/HGM/HGL/HGXL/HG2XL/HG3XL → HAKİ GÖMLEK [beden] BEDEN

### T-SHİRTLER
STS/STXS/STM/STL/STXL/ST2XL-ST4XL → SİYAH T-SHİRT [beden] BEDEN

### TAKTİKAL PANTOLONLAR
TPXS/TPS/TPM/TPL/TPXL/TP2XL-TP5XL → TAKTİKAL PANTOLON [beden] BEDEN

### KABANLAR (M ön eki)
MKS-MK6XL → KABAN KIRMIZI [beden] BEDEN
MLS-ML6XL → KABAN LACİVERT [beden] BEDEN
MGS-MG4XL → KABAN GRİ [beden] BEDEN
MMS-MM5XL → KABAN MAVİ [beden] BEDEN
MTS-MT5XL → KABAN TURUNCU [beden] BEDEN
MYS-MY6XL → KABAN YEŞİL [beden] BEDEN

### MONTLAR
SMS/SMM/SML/SMXL/SM2XL-SM5XL → SİYAH MONT [beden] BEDEN
SAMS-SAM5XL → SARI MONT [beden]
MOS/MOM/MOL/MOXL/MO5XL → OXFORD KABAN [beden] BEDEN

### PANTOLONLAR (P + renk kodu + numara)
PM48-PM64 → PANTOLON MAVİ [numara] BEDEN
PK48-PK64 → PANTOLON KIRMIZI [numara] BEDEN
PL48-PL70 → PANTOLON LACİVERT [numara] BEDEN
PY48-PY66 → PANTOLON YEŞİL [numara] BEDEN
PT48-PT66 → PANTOLON TURUNCU [numara] BEDEN
PG48-PG66 → PANTOLON GRİ [numara] BEDEN

### TULUMLAR
KTS-KT5XL → KIRMIZI TULUM [beden] BEDEN

### YELEKLER
YS-Y5XL → YELEK SARI [beden]

### BARETLER
BM/BB/BK/BT → BARET MAVİ/BEYAZ/KIRMIZI/TURUNCU KARAM ENDÜSTRİYEL KB505

### AYAKKABILAR
BO37-BO47 (veya BOT37, B37) → AYAKKABI [numara] STARLINE 9040B-83
Ç37-Ç47 → ÇİZME [numara] STARLINE 9910-S5-YEŞİL

## İŞLEM KURALLARI

1. **Aynı ürünleri birleştir**: "pm54 1" + "pm54 2" = PANTOLON MAVİ 54 BEDEN → 3
2. **Kategori sırası**: Kazak → Polar → Gömlek → Kaban → Mont → Pantolon → Tulum → Yelek
3. **Alt sıralama**: Renk alfabetik, sonra beden küçükten büyüğe
4. **Beden sırası**: XS<S<M<L<XL<2XL<3XL<4XL<5XL<6XL (harf), 48<50<52...70 (numara)

## EXCEL ŞABLON YAPISI (YIKAMA_TUTANAK_FORMATI.xlsx)

- **Satır 1-3**: Logo + TARİH (I1)
- **Satır 4-5**: Başlık "YIKAMAYA GİDECEK TUTANAĞI" (A4:J5 merged)
- **Satır 6**: Sütun başlıkları — MALZEME ADI (A6:H6) | MİKTAR (I6:J6), font 16pt
- **Satır 7+**: Veri satırları — Her satır A:H merged (malzeme adı) + I:J merged (miktar)
  - Satır yüksekliği: 27, Font: 16pt bold
  - 26'dan fazla kalem varsa: yükseklik 18, font 11pt
- **İmza bloğu**: Veriden 1 boş satır sonra:
  - DEPO SORUMLUSU (A, F merged) | GENEL MÜDÜR (G:J merged) — font 11-12pt bold
  - AD SOYAD: SEZER GÜNAY (A:F, 2 satır) | AD SOYAD: OKAN DURAK (G:J, 2 satır) — font 10pt bold
  - GÖREV: DEPO SORUMLUSU (A:F, 2 satır) | GÖREV: GENEL MÜDÜR (G:J, 2 satır) — font 10pt bold
  - İMZA: (A:F, 2 satır) | İMZA: (G:J, 2 satır) — font 10pt bold

## ÖNEMLİ
- Kullanıcı bazen harf hatası yapar (y yerine h → haki polar gibi), sorarak düzelt
- "kutu" veya "koli" başlıkları varsa ayrı dosyalar üret
- "birleştir" denirse tüm dosyaları oku, aynı ürünleri topla, tek tutanak yap
- İmza bloğu HER ZAMAN verinin hemen altında olmalı, arada fazla boşluk bırakma
