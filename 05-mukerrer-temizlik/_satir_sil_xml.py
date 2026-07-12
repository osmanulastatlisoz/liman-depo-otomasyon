# -*- coding: utf-8 -*-
"""Excel'i hiç çalıştırmadan, .xlsm ZIP içindeki STOK GİRİŞ ÇIKIŞ sayfasından
belirtilen satırları siler. Diğer tüm parçalar (makrolar, butonlar/çizimler,
diğer tablolar) bire bir korunur. Sayfanın ve tablosunun XML dosya adları
sabit DEĞİLDİR; her çalıştırmada workbook.xml ve ilişki dosyalarından
isimle çözülür. calcChain.xml kaldırılır (Excel ilk açılışta yeniden kurar).

Bu bir motor modülüdür; mukerrer_sil.py tarafından çağrılır.
sil: silinecek Excel satır numaraları kümesi (1=başlık, 2..N veri)."""
import bisect
import re
import zipfile

SAYFA_ADI = "STOK GİRİŞ ÇIKIŞ"


def _sayfa_yolu_bul(icerik: dict) -> str:
    """workbook.xml + rels üzerinden sayfa adının XML yolunu çöz."""
    wbx = icerik["xl/workbook.xml"].decode("utf-8")
    rid = None
    for m in re.finditer(r'<sheet[^>]*name="([^"]*)"[^>]*r:id="(rId\d+)"', wbx):
        if m.group(1) == SAYFA_ADI:
            rid = m.group(2)
            break
    if rid is None:
        # öznitelik sırası farklı olabilir: r:id önce, name sonra
        for m in re.finditer(r'<sheet[^>]*r:id="(rId\d+)"[^>]*name="([^"]*)"', wbx):
            if m.group(2) == SAYFA_ADI:
                rid = m.group(1)
                break
    if rid is None:
        raise RuntimeError(f"'{SAYFA_ADI}' sayfası workbook.xml içinde bulunamadı")
    rels = icerik["xl/_rels/workbook.xml.rels"].decode("utf-8")
    m = re.search(r'Id="' + rid + r'"[^>]*Target="([^"]*)"', rels)
    if not m:
        m = re.search(r'Target="([^"]*)"[^>]*Id="' + rid + r'"', rels)
    if not m:
        raise RuntimeError(f"{rid} ilişkisi workbook.xml.rels içinde bulunamadı")
    hedef = m.group(1).lstrip("/")
    if not hedef.startswith("xl/"):
        hedef = "xl/" + hedef
    return hedef


def _tablo_yolu_bul(icerik: dict, sayfa_yolu: str) -> str | None:
    """Sayfanın .rels dosyasından bağlı tablo XML yolunu çöz (yoksa None)."""
    ad = sayfa_yolu.rsplit("/", 1)[-1]
    rels_yolu = f"xl/worksheets/_rels/{ad}.rels"
    if rels_yolu not in icerik:
        return None
    rels = icerik[rels_yolu].decode("utf-8")
    m = re.search(r'Type="[^"]*relationships/table"[^>]*Target="([^"]*)"', rels)
    if not m:
        m = re.search(r'Target="([^"]*tables/[^"]*)"', rels)
    if not m:
        return None
    hedef = m.group(1).replace("../", "")
    if not hedef.startswith("xl/"):
        hedef = "xl/" + hedef
    return hedef


def _sheetdata_isle(xml: str, sil: set) -> tuple[str, int, int]:
    """Sayfa XML'inden satırları sil, kalanları yeniden numarala."""
    sd = re.search(r"(<sheetData>)(.*)(</sheetData>)", xml, re.S)
    if not sd:
        raise RuntimeError("sheetData bulunamadı")
    bas, govde, son = sd.group(1), sd.group(2), sd.group(3)

    sil_sirali = sorted(sil)
    parcalar = []
    eski_max = 0
    bulunan = set()
    for m in re.finditer(r"<row\b[^>]*\sr=\"(\d+)\"[^>]*>.*?</row>", govde, re.S):
        old = int(m.group(1))
        eski_max = max(eski_max, old)
        if old in sil:
            bulunan.add(old)
            continue
        kayma = bisect.bisect_left(sil_sirali, old)
        chunk = m.group(0)
        if kayma:
            new = old - kayma
            chunk = re.sub(r"(\sr=\")([A-Z]*)" + str(old) + r"(\")",
                           lambda mm, n=new: mm.group(1) + mm.group(2) + str(n) + mm.group(3),
                           chunk)
        parcalar.append(chunk)

    eksik = sil - bulunan
    if eksik:
        raise RuntimeError(f"Silinecek satırlar sayfada bulunamadı: {sorted(eksik)[:5]}")
    if eski_max <= max(sil):
        raise RuntimeError(f"Sayfa beklenenden küçük (max {eski_max}) — yanlış sayfa olabilir")

    yeni_govde = "".join(parcalar)
    yeni_xml = xml[:sd.start()] + bas + yeni_govde + son + xml[sd.end():]
    yeni_max = eski_max - len(sil)
    yeni_xml = re.sub(r'(<dimension ref="A1:[A-Z]+)\d+(")',
                      lambda mm: mm.group(1) + str(yeni_max) + mm.group(2), yeni_xml, count=1)
    return yeni_xml, eski_max, yeni_max


def _tablo_isle(xml: str, yeni_max: int) -> str:
    """Tablo ref ve autoFilter ref'i küçült, eski sortState'i kaldır."""
    xml = re.sub(r'(<table[^>]*\sref="[A-Z]+\d+:[A-Z]+)\d+(")',
                 lambda m: m.group(1) + str(yeni_max) + m.group(2), xml, count=1)
    xml = re.sub(r'(<autoFilter ref="[A-Z]+\d+:[A-Z]+)\d+(")',
                 lambda m: m.group(1) + str(yeni_max) + m.group(2), xml, count=1)
    xml = re.sub(r"<sortState\b.*?</sortState>", "", xml, flags=re.S)
    xml = re.sub(r"<sortState\b[^>]*/>", "", xml)
    return xml


def satirlari_sil(kaynak: str, hedef: str, sil: set) -> dict:
    """kaynak .xlsm -> hedef .xlsm, STOK GİRİŞ ÇIKIŞ sayfasından sil."""
    with zipfile.ZipFile(kaynak, "r") as z:
        adlar = z.namelist()
        infolar = {i.filename: i for i in z.infolist()}
        icerik = {ad: z.read(ad) for ad in adlar}

    sayfa_yolu = _sayfa_yolu_bul(icerik)
    tablo_yolu = _tablo_yolu_bul(icerik, sayfa_yolu)

    yeni_sheet, eski_max, yeni_max = _sheetdata_isle(icerik[sayfa_yolu].decode("utf-8"), sil)
    icerik[sayfa_yolu] = yeni_sheet.encode("utf-8")

    if tablo_yolu and tablo_yolu in icerik:
        icerik[tablo_yolu] = _tablo_isle(icerik[tablo_yolu].decode("utf-8"), yeni_max).encode("utf-8")

    atilacak = set()
    if "xl/calcChain.xml" in icerik:
        atilacak.add("xl/calcChain.xml")
        ct = icerik["[Content_Types].xml"].decode("utf-8")
        ct = re.sub(r'<Override[^>]*calcChain[^>]*/>', "", ct)
        icerik["[Content_Types].xml"] = ct.encode("utf-8")
        wr = icerik["xl/_rels/workbook.xml.rels"].decode("utf-8")
        wr = re.sub(r'<Relationship[^>]*calcChain[^>]*/>', "", wr)
        icerik["xl/_rels/workbook.xml.rels"] = wr.encode("utf-8")

    with zipfile.ZipFile(hedef, "w", zipfile.ZIP_DEFLATED) as z:
        for ad in adlar:
            if ad in atilacak:
                continue
            bilgi = infolar[ad]
            yeni = zipfile.ZipInfo(ad, date_time=bilgi.date_time)
            yeni.compress_type = zipfile.ZIP_DEFLATED
            yeni.external_attr = bilgi.external_attr
            z.writestr(yeni, icerik[ad])

    return {"eski_max": eski_max, "yeni_max": yeni_max, "silinen": len(sil),
            "sayfa": sayfa_yolu, "tablo": tablo_yolu}


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    kaynak, hedef = sys.argv[1], sys.argv[2]
    sil = set(int(x) for x in sys.argv[3:])
    print(satirlari_sil(kaynak, hedef, sil))
