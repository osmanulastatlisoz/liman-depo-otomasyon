"""
Eslesmeyen Isim Benzerlik + Manuel Onay
========================================
eslesmeyenler.txt'deki her ismi tum renk klasorlerindeki OLUSTURULAN dosya
adlariyla karsilastirir. Verilen esik orani uzerindeki eslesmeleri numarali
listeler. Onaylanan numaralarin docx'lerinde BOT satiri gri yapilir.

Kullanim:
    python eslesmeyen_benzerlik.py

Akis:
    1. Esik orani gir (orn: 80)
    2. Listelenen eslesmeleri kontrol et
    3. Onaylanacak numaralari yaz (orn: 1,3,5-8  |  a=hepsi  |  Enter=hicbiri)
    4. Yeni esik denemek icin tekrar oran gir, cikmak icin bos Enter.

Onaylanan isimler eslesmeyenler.txt'den otomatik cikarilir. Excel kalici
duzeltmesi icin son ozetteki ESKI -> YENI eslemesini boot Excel'ine uygulayin.
"""

from pathlib import Path
from difflib import SequenceMatcher
import sys

BASE_DIR = Path(__file__).parent.resolve()
ESLESMEYEN_FILE = BASE_DIR / "eslesmeyenler.txt"

sys.path.insert(0, str(BASE_DIR))
from bot_alanlari_isaretle import find_bot_row, mark_row, row_needs_update


def tr_upper(s):
    return s.translate(str.maketrans('iı', 'İI')).upper()


def ascii_fold(s):
    return s.translate(str.maketrans('ÇĞİÖŞÜ', 'CGIOSU'))


def normalize(name):
    return tr_upper(" ".join(str(name).split())).strip()


def similarity(a, b):
    r1 = SequenceMatcher(None, a, b).ratio()
    r2 = SequenceMatcher(None, ascii_fold(a), ascii_fold(b)).ratio()
    return max(r1, r2) * 100


def collect_olusturulan_files():
    files = []
    for klasor in sorted(BASE_DIR.iterdir()):
        if not klasor.is_dir():
            continue
        olu = klasor / "OLUSTURULAN"
        if not olu.is_dir():
            continue
        for docx_path in sorted(olu.glob("*.docx")):
            if docx_path.name.startswith("~$"):
                continue
            files.append((klasor.name, normalize(docx_path.stem), docx_path.stem))
    return files


def parse_selection(s, max_n):
    """'1,3,5-8' veya 'a' (hepsi) seklindeki secimi int set'e cevir."""
    s = s.strip().lower()
    if not s:
        return set()
    if s in ('a', 'all', 'h', 'hepsi'):
        return set(range(1, max_n + 1))
    result = set()
    for part in s.replace(' ', '').split(','):
        if not part:
            continue
        if '-' in part:
            try:
                a, b = map(int, part.split('-', 1))
                result.update(range(min(a, b), max(a, b) + 1))
            except ValueError:
                print(f"  Anlasilamadi: {part}")
        else:
            try:
                result.add(int(part))
            except ValueError:
                print(f"  Anlasilamadi: {part}")
    return {n for n in result if 1 <= n <= max_n}


def mark_docx(klasor_adi, docx_stem):
    """Belirtilen docx'in BOT satirini gri yap. (ok, mesaj) dondur."""
    from docx import Document
    docx_path = BASE_DIR / klasor_adi / "OLUSTURULAN" / f"{docx_stem}.docx"
    if not docx_path.is_file():
        return False, f"dosya yok: {docx_path}"
    try:
        doc = Document(str(docx_path))
        row = find_bot_row(doc)
        if row is None:
            return False, "BOT satiri bulunamadi"
        if not row_needs_update(row):
            return True, "zaten gri"
        mark_row(row)
        doc.save(str(docx_path))
        return True, "gri yapildi"
    except Exception as e:
        return False, f"hata: {e}"


def write_eslesmeyenler(names):
    if names:
        ESLESMEYEN_FILE.write_text(
            "\n".join(sorted(names)) + "\n", encoding="utf-8"
        )
    elif ESLESMEYEN_FILE.exists():
        ESLESMEYEN_FILE.unlink()


def main():
    if not ESLESMEYEN_FILE.is_file():
        sys.exit(
            f"HATA: {ESLESMEYEN_FILE.name} bulunamadi.\n"
            "Once bot_alanlari_isaretle.py'yi calistir."
        )

    eslesmeyenler = set(
        line.strip()
        for line in ESLESMEYEN_FILE.read_text(encoding="utf-8").splitlines()
        if line.strip()
    )
    docx_files = collect_olusturulan_files()

    print(f"Eslesmeyen isim sayisi  : {len(eslesmeyenler)}")
    print(f"OLUSTURULAN dosya sayisi: {len(docx_files)}")
    print()

    # Oturum boyunca yapilan tum esleme tarihcesi (sonda gosterilir)
    onaylanan_eslemeler = []  # [(eksik_isim, docx_orig, klasor)]

    while True:
        if not eslesmeyenler:
            print("Eslesmeyen kalmadi — listede herkes eslesti.")
            break

        try:
            esik_str = input(
                f"\nBenzerlik orani esigi (kalan {len(eslesmeyenler)} isim) "
                "[0-100, cikis icin Enter]: "
            ).strip()
        except (EOFError, KeyboardInterrupt):
            print()
            break
        if not esik_str or esik_str.lower() in ('q', 'quit', 'exit'):
            break
        try:
            esik = float(esik_str.replace('%', '').replace(',', '.'))
        except ValueError:
            print("Sayisal bir deger gir (orn: 80).")
            continue
        if not 0 <= esik <= 100:
            print("0-100 arasinda bir deger gir.")
            continue

        # Eslesmeleri topla
        matches = []  # (eksik, docx_orig, klasor, ratio)
        for eksik in eslesmeyenler:
            eksik_norm = normalize(eksik)
            for klasor, docx_norm, docx_orig in docx_files:
                ratio = similarity(eksik_norm, docx_norm)
                if ratio >= esik:
                    matches.append((eksik, docx_orig, klasor, ratio))
        # Once eksik isme gore, sonra ratio yukseyden asagiya
        matches.sort(key=lambda m: (m[0], -m[3]))

        if not matches:
            print(f"\n%{esik:.0f} ve uzeri oneri bulunamadi. Esigi dusur.\n")
            continue

        print(f"\n=== %{esik:.0f} ve uzeri ({len(matches)} oneri) ===\n")
        header = f"{'NO':>4s}  {'EKSIK ISIM (Excel)':32s} | {'OLUSTURULAN DOSYA':32s} | {'KLASOR':9s} | %"
        print(header)
        print("-" * len(header))
        for i, (eksik, docx_orig, klasor, ratio) in enumerate(matches, 1):
            print(f"[{i:3d}] {eksik:32s} | {docx_orig:32s} | {klasor:9s} | {ratio:5.1f}")
        print("-" * len(header))

        try:
            sec_str = input(
                "\nOnaylanacak numaralar (orn: 1,3,5-8  |  a=hepsi  |  Enter=hicbiri): "
            )
        except (EOFError, KeyboardInterrupt):
            print()
            break
        secilenler = parse_selection(sec_str, len(matches))
        if not secilenler:
            print("Hicbiri secilmedi, devam.")
            continue

        print()
        confirmed_eksik = set()
        for i in sorted(secilenler):
            eksik, docx_orig, klasor, ratio = matches[i - 1]
            ok, msg = mark_docx(klasor, docx_orig)
            simge = "OK   " if ok else "HATA "
            print(f"  {simge}[{i:3d}] {docx_orig} ({klasor}) — {msg}")
            if ok:
                confirmed_eksik.add(eksik)
                onaylanan_eslemeler.append((eksik, docx_orig, klasor))

        # Onaylananlari eslesmeyenler'den cikar ve dosyayi guncelle
        eslesmeyenler -= confirmed_eksik
        write_eslesmeyenler(eslesmeyenler)
        print(f"\n{len(confirmed_eksik)} isim eslesmeyenler.txt'den cikarildi "
              f"(kalan: {len(eslesmeyenler)}).")

    # Oturum sonu ozeti
    if onaylanan_eslemeler:
        print()
        print("=" * 70)
        print(f"Bu oturumda {len(onaylanan_eslemeler)} docx gri yapildi.")
        print()
        print("Excel'i de kalici duzeltmek istersen, boot Excel'inde su degisiklikleri yap:")
        print(f"{'EXCEL''DEKI ESKI':32s} -> YENI (OLUSTURULAN dosya adi)")
        print("-" * 70)
        for eksik, docx_orig, klasor in onaylanan_eslemeler:
            print(f"{eksik:32s} -> {docx_orig}   [{klasor}]")


if __name__ == "__main__":
    main()
