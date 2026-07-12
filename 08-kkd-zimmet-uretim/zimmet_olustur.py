"""
TCH Zimmet Belgesi Oluşturma Scripti
=====================================
Kullanım:
    Bu scripti bir renk klasörüne kopyala (ör. "griler" klasörü), terminalde:

        python zimmet_olustur.py

Klasörde 1 adet .xlsx (personel listesi) ve 1 adet .docx (zimmet template) olmalı.
Belgeler OLUSTURULAN alt klasörüne "AD SOYAD.docx" formatında kaydedilir.

Excel formatı:
    1. satır: başlık (atlanır)
    A kolonu: AD SOYAD
    B kolonu: TC KİMLİK NUMARASI
"""

from pathlib import Path

BASE_DIR = Path(__file__).parent.resolve()
OLUSTURULAN_DIR = BASE_DIR / "OLUSTURULAN"

PLACEHOLDER_AD = "ÇALIŞANIN ADI SOYADI"
PLACEHOLDER_TC = "TC KİMLİK NUMARASI"
PLACEHOLDER_GOREV = "GÖREVİ"


def find_files():
    """Klasördeki tek xlsx ve tek docx'i bul."""
    xlsx_files = [
        f for f in BASE_DIR.glob("*.xlsx")
        if not f.name.startswith("~$")
    ]
    docx_files = [
        f for f in BASE_DIR.glob("*.docx")
        if not f.name.startswith("~$")
    ]
    if len(xlsx_files) != 1:
        raise SystemExit(
            f"HATA: Klasörde tam olarak 1 .xlsx olmalı. Bulunan: {len(xlsx_files)}\n"
            f"  {[f.name for f in xlsx_files]}"
        )
    if len(docx_files) != 1:
        raise SystemExit(
            f"HATA: Klasörde tam olarak 1 .docx template olmalı. Bulunan: {len(docx_files)}\n"
            f"  {[f.name for f in docx_files]}"
        )
    return xlsx_files[0], docx_files[0]


def read_personnel(xlsx_path):
    """Excel'i oku, (ad, tc, görev) listesi döndür."""
    import openpyxl
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    personnel = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:  # başlık satırı
            continue
        if not row or row[0] is None or row[1] is None:
            continue
        name = str(row[0]).strip()
        tc = str(row[1]).strip()
        # C kolonu opsiyonel — yoksa boş string
        gorev = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""
        if not name or not tc:
            continue
        personnel.append((name, tc, gorev))
    return personnel


def safe_filename(name):
    """Windows'ta dosya adı olarak kullanılabilir hale getir."""
    for ch in '<>:"/\\|?*':
        name = name.replace(ch, '')
    return name.strip()


def insert_value(para, keyword, value, compensate=False):
    """Paragrafta 'keyword' metnini bul, ondan sonra gelen ilk ':' karakterinin
    arkasına değeri ekle. Birden fazla 'Etiket: değer' yapısı olan paragraflarda
    (ör. TC KİMLİK NUMARASI ve GÖREVİ aynı paragrafta ise) doğru yere yazar.

    Template'deki yazılar birden fazla 'run'a bölünmüş olabilir, bu yüzden
    önce tüm run'ları birleştirip pozisyon hesaplıyoruz.

    compensate=True ise, ':' sonrası karakterler boşluklardan oluşuyorsa,
    eklenen karakterin ~2 katı kadar boşluğu siler (TARİH gibi sağdaki
    alanların satır sonunda kaymasını önler)."""
    runs = para.runs
    if not runs:
        return False

    # Run'ların birleştirilmiş metindeki başlangıç pozisyonları
    combined = ""
    run_starts = []
    for r in runs:
        run_starts.append(len(combined))
        combined += r.text

    kw_idx = combined.find(keyword)
    if kw_idx == -1:
        return False

    # keyword'den sonraki ilk ':' pozisyonu
    colon_pos = combined.find(':', kw_idx + len(keyword))
    if colon_pos == -1:
        return False

    # Bu ':' hangi run'ın içinde?
    target_idx = None
    for i in range(len(runs)):
        end = run_starts[i] + len(runs[i].text)
        if run_starts[i] <= colon_pos < end:
            target_idx = i
            break
    if target_idx is None:
        return False

    target_run = runs[target_idx]
    offset = colon_pos - run_starts[target_idx]
    target_run.text = target_run.text[:offset + 1] + f" {value}" + target_run.text[offset + 1:]

    # Kompanse: önce hedef run'un kendi içindeki boşlukları (: sonrası), ardından
    # sonraki run'un baş boşluklarını kısalt. Toplam bütçe = eklenen karakter * 3.
    if compensate:
        inserted_len = len(value) + 1  # +1 eklenen boşluk için
        budget = inserted_len * 3

        # 1) target_run içinde eklenen değerin hemen ardındaki boşluklar
        prefix_len = offset + 1 + inserted_len
        after_value = target_run.text[prefix_len:]
        trailing_spaces = len(after_value) - len(after_value.lstrip(' '))
        remove_here = min(trailing_spaces, budget)
        if remove_here > 0:
            target_run.text = target_run.text[:prefix_len] + after_value[remove_here:]
        budget -= remove_here

        # 2) Kalan bütçeyle sonraki run'un baş boşluklarını kısalt
        if budget > 0 and target_idx + 1 < len(runs):
            next_run = runs[target_idx + 1]
            stripped = next_run.text.lstrip(' ')
            leading_spaces = len(next_run.text) - len(stripped)
            if leading_spaces > 0:
                remove_next = min(leading_spaces, budget)
                keep = max(1, leading_spaces - remove_next)
                next_run.text = ' ' * keep + stripped

    return True


def fill_template(template_path, name, tc, gorev, output_path):
    """Template'i aç, ad, tc ve görevi yerleştir, kaydet."""
    from docx import Document
    doc = Document(str(template_path))

    ad_done = False
    tc_done = False
    gorev_done = False if gorev else True  # görev boşsa zaten atla

    for para in doc.paragraphs:
        if not para.runs:
            continue

        # Aynı paragrafta birden fazla placeholder olabilir (TC ve GÖREVİ genelde
        # aynı paragrafta soft line break ile ayrılmış olur), o yüzden elif değil if
        if not ad_done and PLACEHOLDER_AD in para.text:
            if insert_value(para, PLACEHOLDER_AD, name, compensate=True):
                ad_done = True

        if not tc_done and PLACEHOLDER_TC in para.text:
            if insert_value(para, PLACEHOLDER_TC, tc, compensate=False):
                tc_done = True

        if not gorev_done and PLACEHOLDER_GOREV in para.text:
            if insert_value(para, PLACEHOLDER_GOREV, gorev, compensate=False):
                gorev_done = True

        if ad_done and tc_done and gorev_done:
            break

    if not ad_done:
        print(f"  UYARI: Ad yeri bulunamadi ({name})")
    if not tc_done:
        print(f"  UYARI: TC yeri bulunamadi ({name})")
    if not gorev_done and gorev:
        print(f"  UYARI: Gorev yeri bulunamadi ({name})")

    doc.save(str(output_path))


def main():
    xlsx, docx = find_files()
    print(f"Excel    : {xlsx.name}")
    print(f"Template : {docx.name}")
    print()

    personnel = read_personnel(xlsx)
    print(f"Personel sayisi: {len(personnel)}")
    print()

    OLUSTURULAN_DIR.mkdir(exist_ok=True)

    ok = 0
    for name, tc, gorev in personnel:
        filename = safe_filename(name) + ".docx"
        output = OLUSTURULAN_DIR / filename
        try:
            fill_template(docx, name, tc, gorev, output)
            print(f"  OK    {filename}")
            ok += 1
        except Exception as e:
            print(f"  HATA  {filename}: {e}")

    print()
    print(f"{ok}/{len(personnel)} belge olusturuldu")
    print(f"Klasor: {OLUSTURULAN_DIR}")


if __name__ == "__main__":
    main()
