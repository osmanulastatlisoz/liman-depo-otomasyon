# -*- coding: utf-8 -*-
"""
SU TESLİM FİŞİ İŞLEME YARDIMCISI
================================
Bir klasördeki taranmış SU teslim fişlerini (PDF) işler:
  1) Her fişte "TESLİM FİŞİ" başlığının yanındaki 6 haneli SERİ NUMARASINI
     bulur (örn. 001841). Numaranın sayfadaki YERİ taramaya göre değişiyor
     (kimi fiş üstte, kimi ortada) — bu yüzden Tesseract'ın kelime-kutusu
     (TSV) çıktısıyla numaranın yerini TESPİT eder, sabit bir bölgeye güvenmez.
  2) Dosyayı  <seri no>.pdf  yapar  (örn. 001841.pdf).
  3) Adlandırılan dosyayı  "tarandı"  alt klasörüne TAŞIR.

ONAY PENCERESİ:
  OCR küçük rakamlarda yanılabildiği (000164'ü 000154 okuyabiliyor) ve bazı
  taramalar soluk olduğu için, her fiş için numaranın bulunduğu bölge BÜYÜK
  ve NET gösterilir; OCR tahmini kutuya hazır gelir. Doğruysa Enter, yanlışsa
  düzelt. Numara hiç okunamazsa fişin üst kısmı gösterilir, sen yazarsın.

KULLANIM:
  - Kolay: çift tıkla → klasör seç → her fiş için Enter / düzelt.
  - Komut:
      python su_irsaliye_isle.py                       # klasör seçme penceresi
      python su_irsaliye_isle.py "KLASOR"              # onaylı GUI
      python su_irsaliye_isle.py "KLASOR" --otomatik   # OCR'a güven, sormadan
      python su_irsaliye_isle.py "KLASOR" --kuru       # KURU DENEME (taşımaz)
"""

import os
import re
import io
import csv
import sys
import glob
import shutil
import subprocess
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox

import pypdfium2 as pdfium
from PIL import Image, ImageOps, ImageTk

# Konsol cp1254 olabilir → ok/Türkçe karakterler print'te çökmesin.
for _akis in (sys.stdout, sys.stderr):
    try:
        _akis.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# ---------- AYARLAR ----------
BASLANGIC_KLASORU = os.path.join(os.path.expanduser("~"), "Desktop", "TARAMA")
TESS = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
if not os.path.exists(TESS):          # farklı PC: PATH'teki tesseract'ı kullan
    TESS = shutil.which("tesseract") or TESS
TARANDI_KLASOR = "tarandı"     # adlandırılanlar buraya taşınır (seçilen klasörün altı)
RENDER_OLCEK = 3.0             # OCR/önizleme render çözünürlüğü
UST_ORAN = 0.95               # seri aranan bölge (fiş sayfada aşağıda da olabilir)
MIN_CONF = 70                 # bir 6 haneyi "seri" saymak için en az OCR güveni
GOSTER_W, GOSTER_H = 900, 620 # önizleme görüntüsü bu kutuya sığdırılır
# -----------------------------

ALTI_HANE = re.compile(r"^\d{6}$")
# Fişin sabit başlık kelimeleri — seri okunamazsa bunların yerinden bölge buluruz.
CAPA_RE = re.compile(r"(TESL|F[İIı][ŞS]|ÇUBUK|CUBUK|T[İIı]CARET|GÜL[ŞS]EN|GULSEN|ABANT|SIRMA|S[İIı]RMA)", re.I)


def temizle(no: str) -> str:
    """Dosya adı için sadece rakam bırak."""
    return re.sub(r"\D", "", no or "")


def benzersiz_yol(klasor: str, no: str) -> str:
    """Aynı no varsa _2, _3 ekler."""
    yol = os.path.join(klasor, f"{no}.pdf")
    s = 2
    while os.path.exists(yol):
        yol = os.path.join(klasor, f"{no}_{s}.pdf")
        s += 1
    return yol


# ----------------- OCR + numara tespiti -----------------
def _render(pdf_yolu: str) -> Image.Image:
    """PDF'in ilk sayfasını PIL görüntüsü olarak render et."""
    pdf = pdfium.PdfDocument(pdf_yolu)
    img = pdf[0].render(scale=RENDER_OLCEK).to_pil()
    pdf.close()
    return img


def _ocr_tsv(pil: Image.Image):
    """Görüntüyü Tesseract ile (kelime-kutusu/TSV) oku. Satır listesi döndürür."""
    tmp = os.path.join(tempfile.gettempdir(), "_su_tsv.png")
    pil.save(tmp)
    out = subprocess.run(
        [TESS, tmp, "stdout", "-l", "eng", "--psm", "11", "tsv"],
        capture_output=True, text=True, encoding="utf-8", errors="ignore"
    ).stdout or ""
    return list(csv.DictReader(io.StringIO(out), delimiter="\t"))


def _en_iyi_seri(rows):
    """TSV satırlarından en olası 6 haneli seriyi seç → (metin, (l,t,w,h)) veya None.
    Yüksek güven + sıfırla başlama tercih edilir; telefon/VD parçaları düşük
    güvenle elenir."""
    best = None
    for r in rows:
        t = (r.get("text") or "").strip()
        if not ALTI_HANE.match(t):
            continue
        try:
            conf = float(r.get("conf") or -1)
        except ValueError:
            conf = -1
        if conf < MIN_CONF:
            continue
        # Bu form kütüğünde seriler 00xxxx aralığında → "00" başlayanı en çok kayır.
        skor = conf + (60 if t.startswith("00") else 20 if t.startswith("0") else 0)
        if best is None or skor > best[0]:
            box = (int(r["left"]), int(r["top"]), int(r["width"]), int(r["height"]))
            best = (skor, t, box)
    return (best[1], best[2]) if best else None


def _capa_kutusu(rows):
    """Başlık kelimelerinin (TESLİM/ÇUBUKCU/GÜLŞEN...) birleşik kutusu → seri
    no buranın hemen altında/sağındadır. Yoksa None."""
    kutular = []
    for r in rows:
        t = (r.get("text") or "").strip()
        if not CAPA_RE.search(t):
            continue
        try:
            conf = float(r.get("conf") or -1)
        except ValueError:
            conf = -1
        if conf < 55:
            continue
        kutular.append((int(r["left"]), int(r["top"]), int(r["width"]), int(r["height"])))
    if not kutular:
        return None
    uL = min(b[0] for b in kutular)
    uT = min(b[1] for b in kutular)
    uR = max(b[0] + b[2] for b in kutular)
    uB = max(b[1] + b[3] for b in kutular)
    return (uL, uT, uR, uB)


def _odakli_no(crop):
    """Çapa kırpıntısında seri tahmini: EN ÜSTTEKİ 4-6 haneli rakam (seri hep
    üstte; telefon/VD no'su daha aşağıda). 6'ya sıfırla tamamlanır."""
    en_ust = None
    for r in _ocr_tsv(crop):
        t = (r.get("text") or "").strip()
        d = re.sub(r"\D", "", t)
        if not (4 <= len(d) <= 6):
            continue
        top = int(r["top"])
        if en_ust is None or top < en_ust[0]:
            en_ust = (top, d.zfill(6))
    return en_ust[1] if en_ust else ""


def seri_bul(pdf_yolu: str):
    """(tahmin, onizleme_PIL, guven) döndürür. guven: 'yuksek' | 'dusuk' | 'yok'.
    Fişin sayfadaki yeri değiştiği için üç katman:
      1) Numara doğrudan okunursa → çevresine zoom + 'yuksek'.
      2) Okunamaz ama başlık bulunursa → o bölgeyi kırp, en üstteki rakamı
         tahmin et (telefon değil) → 'dusuk' (kullanıcı kontrol etmeli).
      3) Hiçbiri → kontrastı artırılmış tüm sayfa, tahmin yok → 'yok'."""
    page = _render(pdf_yolu)
    rows = _ocr_tsv(page)

    # 1) doğrudan seri
    b = _en_iyi_seri(rows)
    if b is not None:
        metin, (l, t, ww, hh) = b
        pad = hh or 30
        box = (max(0, l - 9 * pad), max(0, t - 5 * pad),
               min(page.width, l + ww + 3 * pad), min(page.height, t + hh + 7 * pad))
        return metin, page.crop(box), "yuksek"

    # 2) başlık çapasıyla bölge
    cap = _capa_kutusu(rows)
    if cap is not None:
        uL, uT, uR, uB = cap
        uH = max(uB - uT, 40)
        box = (max(0, uL - 60), max(0, uT - 30),
               min(page.width, uR + 500), min(page.height, uB + int(2.4 * uH)))
        crop = page.crop(box)
        return _odakli_no(crop), crop, "dusuk"

    # 3) hiçbir şey okunamadı → kontrastlı tüm sayfa
    return "", ImageOps.autocontrast(page.convert("L"), cutoff=1), "yok"


def _olcekle(pil: Image.Image, maxw: int, maxh: int) -> Image.Image:
    """En-boy oranını koruyarak maxw×maxh kutusuna sığdır."""
    k = min(maxw / pil.width, maxh / pil.height)
    if k < 1:
        pil = pil.resize((max(1, int(pil.width * k)), max(1, int(pil.height * k))), Image.LANCZOS)
    return pil


# ----------------- onay / elle giriş penceresi -----------------
class OnayPenceresi:
    """Her fiş için numara bölgesini büyük gösterir; OCR tahmini kutuda hazır.
    Enter = kaydet & taşı, Esc = atla, Çık = bitir."""

    def __init__(self, klasor, yollar, hedef_klasor):
        self.klasor = klasor
        self.hedef = hedef_klasor
        self.yollar = yollar
        self.idx = 0
        self.tasinan = 0

        self.root = tk.Tk()
        self.root.title("SU Teslim Fişi — No Onayla / Taşı")
        eg = self.root.winfo_screenwidth()
        x = max(0, (eg - (GOSTER_W + 60)) // 2)
        self.root.geometry(f"{GOSTER_W + 60}x{GOSTER_H + 170}+{x}+10")

        self.bilgi = tk.Label(self.root, text="", font=("Segoe UI", 11, "bold"), pady=6)
        self.bilgi.pack(fill="x")
        self.gorsel = tk.Label(self.root, bg="#222", width=GOSTER_W, height=GOSTER_H)
        self.gorsel.pack(pady=4)
        self.gorsel.pack_propagate(False)

        alt = tk.Frame(self.root, pady=8)
        alt.pack(fill="x", padx=20)
        tk.Label(alt, text="Teslim fişi no:", font=("Segoe UI", 11)).pack(side="left")
        self.entry = tk.Entry(alt, font=("Segoe UI", 16), width=16)
        self.entry.pack(side="left", padx=8)
        tk.Button(alt, text="Kaydet & Taşı (Enter)", command=self.kaydet,
                  bg="#4CAF50", fg="white", font=("Segoe UI", 10), padx=10).pack(side="left", padx=4)
        tk.Button(alt, text="Atla (Esc)", command=self.atla,
                  font=("Segoe UI", 10), padx=10).pack(side="left", padx=4)
        tk.Button(alt, text="Çık", command=self.cik,
                  font=("Segoe UI", 10), padx=10).pack(side="left", padx=4)

        self.root.bind("<Return>", lambda e: self.kaydet())
        self.root.bind("<Escape>", lambda e: self.atla())
        self._yukle()
        self.root.mainloop()

    def _yukle(self):
        if self.idx >= len(self.yollar):
            messagebox.showinfo("Bitti", f"Tamamlandı. {self.tasinan} fiş taşındı. ✅")
            self.root.destroy()
            return
        yol = self.yollar[self.idx]
        ad = os.path.basename(yol)
        self.bilgi.config(text=f"[{self.idx+1}/{len(self.yollar)}]  {ad}   —   okunuyor…")
        self.root.update_idletasks()
        try:
            tahmin, prev, guven = seri_bul(yol)
            prev = _olcekle(prev, GOSTER_W, GOSTER_H)
            self.tk_img = ImageTk.PhotoImage(prev)
            self.gorsel.config(image=self.tk_img)
        except Exception as e:
            messagebox.showerror("Hata", f"{ad} okunamadı:\n{e}")
            self.idx += 1
            self._yukle()
            return
        durum = {
            "yuksek": f"OCR tahmini: {tahmin}  (doğruysa Enter)",
            "dusuk":  f"⚠ tahmin: {tahmin}  — fişten KONTROL ET / düzelt",
            "yok":    "OCR okuyamadı — fişe bakıp numarayı yaz",
        }[guven]
        self.bilgi.config(text=f"[{self.idx+1}/{len(self.yollar)}]  {ad}   —   {durum}")
        self.entry.delete(0, tk.END)
        if tahmin:
            self.entry.insert(0, tahmin)
            self.entry.select_range(0, tk.END)
        self.entry.focus()

    def kaydet(self):
        no = temizle(self.entry.get())
        if not no:
            self.atla()
            return
        yol = self.yollar[self.idx]
        os.makedirs(self.hedef, exist_ok=True)
        hedef = benzersiz_yol(self.hedef, no)
        try:
            shutil.move(yol, hedef)
            print(f"✓ {os.path.basename(yol)} → {os.path.relpath(hedef, self.klasor)}")
            self.tasinan += 1
        except PermissionError:
            messagebox.showwarning("Dosya açık",
                f"{os.path.basename(yol)} başka bir programda AÇIK (PDF görüntüleyici?).\n"
                "Lütfen kapatıp tekrar Kaydet'e bas.")
            return
        except Exception as e:
            messagebox.showerror("Hata", f"Taşınamadı:\n{e}")
            return
        self.idx += 1
        self._yukle()

    def atla(self):
        print(f"⊘ Atlandı: {os.path.basename(self.yollar[self.idx])}")
        self.idx += 1
        self._yukle()

    def cik(self):
        self.root.destroy()


# ----------------- akış -----------------
def pdf_listele(klasor):
    """Seçilen klasördeki PDF'ler (tarandı alt klasörü hariç — glob zaten
    recursive değil; ayrıca klasörün KENDİSİ tarandı ise boş döner)."""
    if os.path.basename(os.path.normpath(klasor)) == TARANDI_KLASOR:
        return []
    return sorted(glob.glob(os.path.join(klasor, "*.pdf")))


def otomatik_tasi(klasor, yollar, hedef_klasor, kuru=False):
    """Onaysız: OCR tahminiyle adlandırıp taşır. (--otomatik / --kuru)"""
    n = 0
    for yol in yollar:
        ad = os.path.basename(yol)
        no = temizle(seri_bul(yol)[0])
        if not no:
            print(f"⚠ ELLE BAK (no okunamadı): {ad}")
            continue
        if kuru:
            print(f"  {ad:12s} →  {TARANDI_KLASOR}\\{no}.pdf")
            continue
        os.makedirs(hedef_klasor, exist_ok=True)
        hedef = benzersiz_yol(hedef_klasor, no)
        try:
            shutil.move(yol, hedef)
            print(f"✓ {ad} → {os.path.relpath(hedef, klasor)}")
            n += 1
        except PermissionError:
            print(f"✗ KİLİTLİ (açık): {ad} — kapatıp tekrar çalıştır.")
        except Exception as e:
            print(f"✗ {ad}: {e}")
    return n


def calistir(klasor, otomatik=False, kuru=False, etkilesimli=True, root=None):
    yollar = pdf_listele(klasor)
    if not yollar:
        msg = f"Bu klasörde işlenecek PDF yok:\n{klasor}"
        print(msg)
        if etkilesimli and root:
            messagebox.showwarning("SU Teslim Fişi", msg, parent=root)
        if root is not None:
            root.destroy()
        return

    hedef_klasor = os.path.join(klasor, TARANDI_KLASOR)
    print(f"\n{len(yollar)} PDF — klasör: {klasor}")

    if kuru:
        print(f"\n(KURU DENEME — hiçbir dosya taşınmadı. Hedef: {TARANDI_KLASOR}\\)\n")
        otomatik_tasi(klasor, yollar, hedef_klasor, kuru=True)
        if root is not None:
            root.destroy()
        return

    if otomatik or not etkilesimli:
        n = otomatik_tasi(klasor, yollar, hedef_klasor, kuru=False)
        print(f"\n{n} fiş '{TARANDI_KLASOR}' klasörüne taşındı.")
        if root is not None:
            root.destroy()
        return

    if root is not None:
        root.destroy()
    OnayPenceresi(klasor, yollar, hedef_klasor)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    otomatik = "--otomatik" in sys.argv
    kuru = "--kuru" in sys.argv

    if args:
        calistir(args[0], otomatik=otomatik, kuru=kuru,
                 etkilesimli=not (otomatik or kuru), root=None)
        return

    root = tk.Tk()
    root.withdraw()
    klasor = filedialog.askdirectory(
        title="SU teslim fişlerinin olduğu klasörü seç",
        initialdir=BASLANGIC_KLASORU if os.path.isdir(BASLANGIC_KLASORU) else os.getcwd())
    if not klasor:
        print("Klasör seçilmedi, çıkılıyor.")
        root.destroy()
        return
    calistir(klasor, otomatik=False, kuru=False, etkilesimli=True, root=root)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        print(tb)
        try:
            r = tk.Tk(); r.withdraw()
            messagebox.showerror("SU Teslim Fişi — Hata", f"{e}\n\n{tb}")
            r.destroy()
        except Exception:
            pass
