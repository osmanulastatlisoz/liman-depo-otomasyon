# -*- coding: utf-8 -*-
"""
KKD ZİMMET TUTANAĞI - PDF'leri PERSONEL ADIYLA otomatik adlandırma.
====================================================================
Bir klasördeki taranmış KKD zimmet PDF'lerini işler. Her PDF'in
"ÇALIŞANIN ADI SOYADI : ......" satırından personelin adını okur ve
dosyayı  <Ad Soyad>.pdf  yapar. (OCR'ye gerek yok — PDF'lerde metin
katmanı var, pypdf ile bedavaya okunur.)

İKİ AŞAMA:
  1) GÜVENLİ olanlar  → otomatik adlandırılır.
  2) Okunamayan / emin olunamayan / el yazısı olanlar → küçük bir
     pencerede sayfanın ÜST kısmı gösterilir, ismi SEN yazarsın
     (kkd_cikis.py / kkd_bol.py'deki gibi). Tahmin varsa kutuya
     hazır yazılır; Enter'a basıp onaylar ya da düzeltirsin.

GÜVEN KONTROLÜ (uydurmaz!):
  İsim sadece harflerden oluşuyor diye "doğru" sayılmaz — el yazısı bir
  form da harf üretebilir. Bu yüzden ÇAPRAZ KONTROL: aynı sayfadaki
  "TC KİMLİK NUMARASI" temiz 11 hane okunuyorsa metin katmanı güvenilir
  demektir → isim otomatik uygulanır. Okunmuyorsa (el yazısı/bozuk)
  → ELLE BAK'a düşer.

KULLANIM:
  - Kolay: kkd_adlandir.bat'a çift tıkla → klasör seç.
  - Komut:
      python kkd_adlandir.py                       # klasör seçme penceresi
      python kkd_adlandir.py "KLASOR"              # o klasörü işle
      python kkd_adlandir.py "KLASOR" --kuru       # KURU DENEME (hiçbir şey değişmez)
"""

import os
import re
import sys
import glob
import tkinter as tk
from tkinter import filedialog, messagebox
from pypdf import PdfReader
import pypdfium2 as pdfium
from PIL import Image, ImageTk

# ---------- AYARLAR ----------
BASLANGIC_KLASORU = os.path.join(os.path.expanduser("~"), "Desktop", "TARAMA")
ON_IZLEME_ORANI = 0.22          # ELLE BAK penceresinde sayfanın üstten ne kadarı gösterilsin
RENDER_OLCEK = 2.0              # Önizleme görüntü çözünürlüğü
PENCERE_GENISLIK = 1000
PENCERE_UST_OFFSET = 10
BUYUK_HARF = True              # True: "METİN SEFERCİK.pdf" (formdaki gibi). False: "Metin Sefercik.pdf"
# -----------------------------

# "ÇALIŞANIN ADI SOYADI : <isim> TARİH ..."  — etiketten sonrasını al
LABEL_RE = re.compile(r"ADI\s*SOYAD[Iıİi]\s*[:：]?\s*(.+)", re.IGNORECASE)
TARIH_KES_RE = re.compile(r"\bTAR[İIıi]H\b", re.IGNORECASE)
# Geçerli isim: yalnız harf (Türkçe dahil) + boşluk + . ' -
GECERLI_ISIM_RE = re.compile(r"^[A-Za-zÇĞİIÖŞÜçğıiöşü .'\-]+$")
# Çapraz kontrol: TC KİMLİK NUMARASI satırında TEMİZ 11 hane var mı?
TC11_RE = re.compile(r"K[İIıi]ML[İIıi]K\s*NUMARAS[Iı]\s*[:：]?\s*(\d{11})\b", re.IGNORECASE)


def temizle_dosya_adi(ad: str) -> str:
    """Windows için yasak karakterleri sil, boşlukları normalize et."""
    ad = re.sub(r'[<>:"/\\|?*\n\r\t]', '', ad)
    ad = re.sub(r'\s+', ' ', ad).strip()
    return ad if ad else "isimsiz"


def benzersiz_yol(klasor: str, ad: str, simdiki: str = "") -> str:
    """Aynı isimde dosya varsa _2, _3 ekler. simdiki = bu dosyanın kendi yolu (çakışma sayılmaz)."""
    yol = os.path.join(klasor, f"{ad}.pdf")
    sayac = 2
    while os.path.exists(yol) and os.path.abspath(yol) != os.path.abspath(simdiki):
        yol = os.path.join(klasor, f"{ad}_{sayac}.pdf")
        sayac += 1
    return yol


def bicim_isim(ad: str) -> str:
    """İstenirse Türkçe-doğru başlık biçimi (İstanbul kuralları)."""
    if BUYUK_HARF:
        return ad
    out = []
    for kelime in ad.split(" "):
        if not kelime:
            continue
        ilk = kelime[0]
        ilk = "İ" if ilk in ("i", "I", "İ", "ı") and ilk in ("i",) else ilk
        # basit ve güvenli: ilk harf büyük (Türkçe i->İ), kalanı küçük (I->ı)
        bas = kelime[0].upper().replace("I", "İ")
        kalan = kelime[1:].lower().replace("i̇", "i")
        out.append(bas + kalan)
    return " ".join(out)


def isim_oku(text: str) -> str:
    """Metinden 'ÇALIŞANIN ADI SOYADI' ismini çıkar (boşsa '')."""
    for satir in text.splitlines():
        m = LABEL_RE.search(satir)
        if m:
            kalan = m.group(1)
            kt = TARIH_KES_RE.search(kalan)          # TARİH'te kes
            if kt:
                kalan = kalan[:kt.start()]
            kalan = re.sub(r"\s+", " ", kalan).strip(" :.-\t")
            return kalan
    return ""


def guven_durumu(isim: str, text: str) -> str:
    """'guvenli' | 'supheli' | 'yok' döndürür."""
    if not isim:
        return "yok"
    kelimeler = isim.split()
    isim_uygun = (
        bool(GECERLI_ISIM_RE.match(isim))
        and 4 <= len(isim) <= 40
        and 2 <= len(kelimeler) <= 4
    )
    tc_temiz = bool(TC11_RE.search(text))   # çapraz kontrol: temiz 11 haneli TC
    if isim_uygun and tc_temiz:
        return "guvenli"
    return "supheli"


def metni_al(pdf_yolu: str) -> str:
    try:
        r = PdfReader(pdf_yolu)
        return "\n".join((pg.extract_text() or "") for pg in r.pages)
    except Exception:
        return ""


def analiz(klasor: str):
    """Her PDF için karar üret — HİÇBİR dosyayı değiştirmez."""
    pdfs = sorted(glob.glob(os.path.join(klasor, "*.pdf")))
    sonuc = []
    for p in pdfs:
        ad_dosya = os.path.basename(p)
        text = metni_al(p)
        isim = isim_oku(text)
        durum = guven_durumu(isim, text)             # guvenli | supheli | yok
        hedef = ""
        if durum == "guvenli":
            temiz = temizle_dosya_adi(bicim_isim(isim))
            hedef = os.path.basename(benzersiz_yol(klasor, temiz, simdiki=p))
        sonuc.append({
            "yol": p, "ad": ad_dosya, "isim": isim,
            "durum": durum, "hedef": hedef,
        })
    return sonuc


def otomatik_adlandir(klasor, sonuclar):
    """durum=='guvenli' olanları adlandırır. Adlandırılan sayısını döndürür."""
    n = 0
    for r in sonuclar:
        if r["durum"] != "guvenli":
            continue
        if r["ad"] == r["hedef"]:                    # zaten doğru
            continue
        yeni = benzersiz_yol(klasor, os.path.splitext(r["hedef"])[0], simdiki=r["yol"])
        try:
            os.rename(r["yol"], yeni)
            r["yol"] = yeni
            r["ad"] = os.path.basename(yeni)
            print(f"✓ {yeni}")
            n += 1
        except Exception as e:
            print(f"✗ {r['ad']}: {e}")
            r["durum"] = "supheli"                    # adlandırılamadı → elle bak'a düşür
    return n


# ----------------- ELLE BAK penceresi -----------------
class ElleBakPenceresi:
    """Şüpheli/okunamayan PDF'ler için: üst kısmı göster, isim yazdır."""

    def __init__(self, klasor, isler):
        self.klasor = klasor
        self.isler = isler          # [{yol, ad, isim, durum, ...}]
        self.idx = 0

        self.root = tk.Tk()
        self.root.title("KKD - ELLE BAK / İsim Yaz")
        eg = self.root.winfo_screenwidth()
        x = (eg - PENCERE_GENISLIK) // 2
        self.root.geometry(f"{PENCERE_GENISLIK}x520+{x}+{PENCERE_UST_OFFSET}")
        self.root.resizable(True, True)

        self.bilgi = tk.Label(self.root, text="", font=("Segoe UI", 11, "bold"), pady=6)
        self.bilgi.pack(fill="x")

        self.gorsel = tk.Label(self.root, bg="#222")
        self.gorsel.pack(pady=4)

        alt = tk.Frame(self.root, pady=8)
        alt.pack(fill="x", padx=20)
        tk.Label(alt, text="Personel adı:", font=("Segoe UI", 11)).pack(side="left")
        self.entry = tk.Entry(alt, font=("Segoe UI", 13), width=40)
        self.entry.pack(side="left", padx=8, fill="x", expand=True)
        self.entry.focus()

        tk.Button(alt, text="Kaydet (Enter)", command=self.kaydet,
                  font=("Segoe UI", 10), bg="#4CAF50", fg="white", padx=10).pack(side="left", padx=4)
        tk.Button(alt, text="Atla (Esc)", command=self.atla,
                  font=("Segoe UI", 10), padx=10).pack(side="left", padx=4)
        tk.Button(alt, text="Çık", command=self.cik,
                  font=("Segoe UI", 10), padx=10).pack(side="left", padx=4)

        self.root.bind("<Return>", lambda e: self.kaydet())
        self.root.bind("<Escape>", lambda e: self.atla())

        self._yukle()
        self.root.mainloop()

    def _yukle(self):
        if self.idx >= len(self.isler):
            messagebox.showinfo("Bitti", "Elle bakılacak dosya kalmadı. ✅")
            self.root.destroy()
            return

        is_ = self.isler[self.idx]
        etiket = {"supheli": "⚠ EMİN DEĞİL", "yok": "⚠ İSİM OKUNAMADI"}.get(is_["durum"], "")
        self.bilgi.config(text=f"[{self.idx+1}/{len(self.isler)}]  {is_['ad']}   {etiket}")

        try:
            pdf = pdfium.PdfDocument(is_["yol"])
            img = pdf[0].render(scale=RENDER_OLCEK).to_pil()
            w, h = img.size
            kirp = img.crop((0, 0, w, int(h * ON_IZLEME_ORANI)))
            pdf.close()
            oran = (PENCERE_GENISLIK - 40) / kirp.width
            kirp = kirp.resize((int(kirp.width * oran), int(kirp.height * oran)), Image.LANCZOS)
            self.tk_img = ImageTk.PhotoImage(kirp)
            self.gorsel.config(image=self.tk_img)
        except Exception as e:
            messagebox.showerror("Hata", f"{is_['ad']} okunamadı:\n{e}")
            self.idx += 1
            self._yukle()
            return

        # Tahmin varsa kutuya hazır yaz (onayla/düzelt)
        self.entry.delete(0, tk.END)
        if is_["isim"]:
            self.entry.insert(0, is_["isim"])
            self.entry.select_range(0, tk.END)
        self.entry.focus()

    def kaydet(self):
        ad = self.entry.get().strip()
        if not ad:
            self.atla()
            return
        ad = temizle_dosya_adi(bicim_isim(ad))
        is_ = self.isler[self.idx]
        yeni = benzersiz_yol(self.klasor, ad, simdiki=is_["yol"])
        try:
            os.rename(is_["yol"], yeni)
            print(f"✓ (elle) {is_['ad']} → {os.path.basename(yeni)}")
        except Exception as e:
            messagebox.showerror("Hata", f"Adlandırma başarısız:\n{e}")
            return
        self.idx += 1
        self._yukle()

    def atla(self):
        print(f"⊘ Atlandı: {self.isler[self.idx]['ad']}")
        self.idx += 1
        self._yukle()

    def cik(self):
        self.root.destroy()


# ----------------- AKIŞ -----------------
def ozet_satir(r):
    if r["durum"] == "guvenli":
        return f"  {r['ad']:22s} →  {r['hedef']}"
    if r["durum"] == "supheli":
        return f"  {r['ad']:22s} ⚠ EMİN DEĞİL  (tahmin: {r['isim'] or '-'})"
    return f"  {r['ad']:22s} ⚠ İSİM OKUNAMADI"


def onay_penceresi(root, sonuclar, guvenli_n, elle_n):
    """Kaydırmalı liste + altta SABİT Evet/Hayır butonları olan onay penceresi.
    (messagebox çok dosyada ekrana sığmıyordu, butonlar aşağı kaçıyordu.)"""
    dlg = tk.Toplevel(root)
    dlg.title("KKD Adlandır — Onay")
    # DİKKAT: root.withdraw() ile gizliyken transient() onay penceresini de
    # gizler (görünmez takılır). Bu yüzden transient KULLANMIYORUZ.
    sw, sh = dlg.winfo_screenwidth(), dlg.winfo_screenheight()
    W, H = 780, min(740, sh - 80)
    dlg.geometry(f"{W}x{H}+{(sw - W)//2}+{max(10, (sh - H)//3)}")
    dlg.minsize(560, 400)

    sonuc = {"ok": False}

    def evet():
        sonuc["ok"] = True
        dlg.destroy()

    def hayir():
        sonuc["ok"] = False
        dlg.destroy()

    # Üst: özet
    tk.Label(dlg, font=("Segoe UI", 11, "bold"), pady=8, justify="left",
             text=(f"{len(sonuclar)} PDF bulundu.\n"
                   f"Otomatik adlandırılacak: {guvenli_n}      ⚠ Elle bakılacak: {elle_n}")
             ).pack(side="top", fill="x", padx=12)

    # Alt buton çubuğu — ÖNCE pack'lenir, böylece liste ne kadar uzun olursa olsun HEP görünür
    btn = tk.Frame(dlg, pady=10)
    btn.pack(side="bottom", fill="x")
    tk.Button(btn, text="Evet, adlandır", command=evet, bg="#4CAF50", fg="white",
              font=("Segoe UI", 11, "bold"), padx=18, pady=6).pack(side="right", padx=12)
    tk.Button(btn, text="Hayır / İptal", command=hayir,
              font=("Segoe UI", 11), padx=14, pady=6).pack(side="right")
    tk.Label(dlg, fg="#666", font=("Segoe UI", 9), justify="left",
             text="Sarı = emin değil · Turuncu = isim okunamadı  (bunlar adlandırılmaz, sonra elle bakılır)."
             ).pack(side="bottom", fill="x", padx=12)

    # Orta: kaydırmalı liste (Listbox + Scrollbar)
    cer = tk.Frame(dlg)
    cer.pack(side="top", fill="both", expand=True, padx=12, pady=6)
    sb = tk.Scrollbar(cer, orient="vertical")
    sb.pack(side="right", fill="y")
    lb = tk.Listbox(cer, yscrollcommand=sb.set, font=("Consolas", 10),
                    activestyle="none", borderwidth=1, relief="solid")
    lb.pack(side="left", fill="both", expand=True)
    sb.config(command=lb.yview)

    for i, r in enumerate(sonuclar):
        if r["durum"] == "guvenli":
            lb.insert("end", f"  {r['ad']}   →   {r['hedef']}")
        elif r["durum"] == "supheli":
            lb.insert("end", f"  {r['ad']}   ⚠ EMİN DEĞİL  (tahmin: {r['isim'] or '-'})")
            lb.itemconfig(i, bg="#FFF3B0")
        else:
            lb.insert("end", f"  {r['ad']}   ⚠ İSİM OKUNAMADI")
            lb.itemconfig(i, bg="#FFD2A6")

    def _wheel(e):
        lb.yview_scroll(int(-1 * (e.delta / 120)), "units")
    lb.bind("<MouseWheel>", _wheel)
    dlg.bind("<MouseWheel>", _wheel)
    dlg.bind("<Prior>", lambda e: lb.yview_scroll(-1, "pages"))   # PageUp
    dlg.bind("<Next>", lambda e: lb.yview_scroll(1, "pages"))     # PageDown
    dlg.bind("<Return>", lambda e: evet())
    dlg.bind("<Escape>", lambda e: hayir())
    dlg.protocol("WM_DELETE_WINDOW", hayir)

    dlg.update_idletasks()
    dlg.deiconify()                       # görünür olduğundan emin ol
    dlg.lift()
    dlg.attributes("-topmost", True)
    dlg.after(400, lambda: dlg.attributes("-topmost", False))
    dlg.focus_force()
    lb.focus_set()
    dlg.grab_set()
    dlg.wait_window(dlg)                  # bu pencere kapanana kadar bekle
    return sonuc["ok"]


def calistir(klasor, kuru=False, etkilesimli=True, root=None):
    sonuclar = analiz(klasor)
    if not sonuclar:
        msg = f"Bu klasörde PDF yok:\n{klasor}"
        print(msg)
        if etkilesimli and root:
            messagebox.showwarning("KKD Adlandır", msg, parent=root)
        return

    guvenli = [r for r in sonuclar if r["durum"] == "guvenli"]
    elle = [r for r in sonuclar if r["durum"] != "guvenli"]

    print(f"\n{len(sonuclar)} PDF — klasör: {klasor}\n")
    for r in sonuclar:
        print(ozet_satir(r))
    print(f"\nOtomatik adlandırılacak: {len(guvenli)}    ⚠ Elle bakılacak: {len(elle)}")

    if kuru:
        print("\n(KURU DENEME — hiçbir dosya değiştirilmedi.)")
        return

    # Onay — kaydırmalı, butonları hep görünen özel pencere
    devam = True
    if etkilesimli and root is not None:
        devam = onay_penceresi(root, sonuclar, len(guvenli), len(elle))
    if not devam:
        print("İptal edildi.")
        return

    n = otomatik_adlandir(klasor, sonuclar)
    print(f"\n{n} dosya otomatik adlandırıldı.")

    # Adlandırma sonrası elle bak listesini tazele (durumu değişmiş olabilir)
    elle = [r for r in sonuclar if r["durum"] != "guvenli"]

    if etkilesimli:
        if elle:
            if root is not None:
                messagebox.showinfo(
                    "KKD Adlandır",
                    f"{n} dosya otomatik adlandırıldı.\n\n"
                    f"Şimdi ⚠ {len(elle)} dosya için önizleme açılacak — isimleri sen yaz.",
                    parent=root)
                root.destroy()
                root = None
            ElleBakPenceresi(klasor, elle)
        else:
            if root is not None:
                messagebox.showinfo("KKD Adlandır — Bitti",
                                    f"{n} dosya adlandırıldı. Elle bakılacak yok. ✅", parent=root)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    kuru = "--kuru" in sys.argv

    if args:                                   # komut satırından klasör verildi
        calistir(args[0], kuru=kuru, etkilesimli=not kuru, root=tk.Tk() if not kuru else None)
        return

    # Argüman yok → klasör seçme penceresi
    root = tk.Tk()
    root.withdraw()
    klasor = filedialog.askdirectory(
        title="KKD zimmet PDF'lerinin olduğu klasörü seç",
        initialdir=BASLANGIC_KLASORU if os.path.isdir(BASLANGIC_KLASORU) else os.getcwd())
    if not klasor:
        print("Klasör seçilmedi, çıkılıyor.")
        root.destroy()
        return
    root.deiconify()
    root.withdraw()
    calistir(klasor, kuru=False, etkilesimli=True, root=root)
    try:
        root.destroy()
    except Exception:
        pass


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        print(tb)
        try:
            r = tk.Tk(); r.withdraw()
            messagebox.showerror("KKD Adlandır — Hata", f"{e}\n\n{tb}")
            r.destroy()
        except Exception:
            pass
