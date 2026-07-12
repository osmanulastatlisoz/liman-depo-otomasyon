"""
KKD Zimmet Tutanağı - PDF'i her personel için 2 sayfalık ayrı dosyalara böler.

Tek bir büyük taranmış PDF'i alır (her personel 2 ardışık sayfa).
Her çift için ilk sayfanın üst kısmı (isim alanı) küçük bir pencerede
ekranın üst kısmında gösterilir. Sen ismi yazarsın, 2 sayfalık PDF
o isimle çıktı klasörüne kaydedilir.

Kullanım:
  1. Bu script'i, bölmek istediğin PDF ile aynı klasöre koy.
  2. Aşağıdaki GIRIS_PDF değişkenini kendi dosya adına göre güncelle.
  3. python kkd_bol.py

Pencerede:
  - İsmi yaz + Enter      → o isimle kaydet, sonrakine geç
  - Enter (boş)           → atla (personel_XX.pdf olarak kaydet)
  - Atla butonu / Esc     → atla
  - Çık butonu            → kapat (o ana kadarki çıktılar korunur)
"""

import os
import re
import sys
import tkinter as tk
from tkinter import messagebox
from pypdf import PdfReader, PdfWriter
import pypdfium2 as pdfium
from PIL import Image, ImageTk

# ---------- AYARLAR ----------
GIRIS_PDF = "12.05.2026 TARANAN KKD.pdf"   # Bölünecek PDF dosyasının adı
CIKTI_KLASORU = "personeller"               # Çıktı klasörü
ON_IZLEME_ORANI = 0.20                      # Sayfanın üstten ne kadarı gösterilsin
RENDER_OLCEK = 2.0                          # Görüntü çözünürlüğü
PENCERE_GENISLIK = 1000                     # Pencere genişliği (px)
PENCERE_UST_OFFSET = 10                     # Ekran üstünden uzaklık (px)
# -----------------------------


def temizle_dosya_adi(ad: str) -> str:
    """Windows/Linux için yasak karakterleri sil, boşlukları normalize et."""
    ad = re.sub(r'[<>:"/\\|?*\n\r\t]', '', ad)
    ad = re.sub(r'\s+', ' ', ad).strip()
    return ad if ad else "isimsiz"


def benzersiz_yol(klasor: str, ad: str) -> str:
    """Aynı isimde dosya varsa _2, _3, ... ekler."""
    yol = os.path.join(klasor, f"{ad}.pdf")
    sayac = 2
    while os.path.exists(yol):
        yol = os.path.join(klasor, f"{ad}_{sayac}.pdf")
        sayac += 1
    return yol


class Uygulama:
    def __init__(self, reader: PdfReader, pdf_render: pdfium.PdfDocument,
                 cift_sayisi: int):
        self.reader = reader
        self.pdf_render = pdf_render
        self.cift_sayisi = cift_sayisi
        self.idx = 0

        self.root = tk.Tk()
        self.root.title("KKD Zimmet - PDF Bölme")

        # Pencereyi ekranın üst kısmına yerleştir
        ekran_genislik = self.root.winfo_screenwidth()
        x = (ekran_genislik - PENCERE_GENISLIK) // 2
        self.root.geometry(f"{PENCERE_GENISLIK}x500+{x}+{PENCERE_UST_OFFSET}")
        self.root.resizable(True, True)

        # En üstte: ilerleme bilgisi
        self.bilgi_etiketi = tk.Label(
            self.root, text="", font=("Segoe UI", 11, "bold"), pady=6
        )
        self.bilgi_etiketi.pack(fill="x")

        # Görüntü
        self.gorsel_etiketi = tk.Label(self.root, bg="#222")
        self.gorsel_etiketi.pack(pady=4)

        # Alt çerçeve: input + butonlar
        alt = tk.Frame(self.root, pady=8)
        alt.pack(fill="x", padx=20)

        tk.Label(alt, text="Personel adı:", font=("Segoe UI", 11)).pack(side="left")

        self.entry = tk.Entry(alt, font=("Segoe UI", 13), width=40)
        self.entry.pack(side="left", padx=8, fill="x", expand=True)
        self.entry.focus()

        tk.Button(
            alt, text="Kaydet (Enter)", command=self.kaydet,
            font=("Segoe UI", 10), bg="#4CAF50", fg="white", padx=10
        ).pack(side="left", padx=4)

        tk.Button(
            alt, text="Atla (Esc)", command=self.atla,
            font=("Segoe UI", 10), padx=10
        ).pack(side="left", padx=4)

        tk.Button(
            alt, text="Çık", command=self.cik,
            font=("Segoe UI", 10), padx=10
        ).pack(side="left", padx=4)

        # Klavye kısayolları
        self.root.bind("<Return>", lambda e: self.kaydet())
        self.root.bind("<Escape>", lambda e: self.atla())

        # İlk çifti yükle
        self._yukle()
        self.root.mainloop()

    def _yukle(self):
        if self.idx >= self.cift_sayisi:
            messagebox.showinfo(
                "Bitti",
                f"Tüm personeller işlendi. ✅\n\n"
                f"PDF'ler '{CIKTI_KLASORU}/' klasöründe."
            )
            self.root.destroy()
            return

        sayfa_idx = self.idx * 2

        self.bilgi_etiketi.config(
            text=f"[{self.idx+1}/{self.cift_sayisi}]  "
                 f"Sayfalar {sayfa_idx+1}-{sayfa_idx+2}"
        )

        try:
            page = self.pdf_render[sayfa_idx]
            img = page.render(scale=RENDER_OLCEK).to_pil()
            w, h = img.size
            kirpilan = img.crop((0, 0, w, int(h * ON_IZLEME_ORANI)))

            # Pencere genişliğine ölçekle
            oran = (PENCERE_GENISLIK - 40) / kirpilan.width
            yeni_boyut = (
                int(kirpilan.width * oran),
                int(kirpilan.height * oran),
            )
            kirpilan = kirpilan.resize(yeni_boyut, Image.LANCZOS)
            self.tk_img = ImageTk.PhotoImage(kirpilan)
            self.gorsel_etiketi.config(image=self.tk_img)
        except Exception as e:
            messagebox.showerror(
                "Hata", f"Sayfa {sayfa_idx+1} okunamadı:\n{e}"
            )
            self.idx += 1
            self._yukle()
            return

        self.entry.delete(0, tk.END)
        self.entry.focus()

    def _pdf_kaydet(self, ad: str):
        """2 sayfalık yeni PDF'i kaydet."""
        sayfa_idx = self.idx * 2
        writer = PdfWriter()
        writer.add_page(self.reader.pages[sayfa_idx])
        writer.add_page(self.reader.pages[sayfa_idx + 1])

        cikti = benzersiz_yol(CIKTI_KLASORU, ad)
        with open(cikti, "wb") as f:
            writer.write(f)
        print(f"✓ [{self.idx+1}/{self.cift_sayisi}] {cikti}")

    def kaydet(self):
        cevap = self.entry.get().strip()
        if not cevap:
            self.atla()
            return

        ad = temizle_dosya_adi(cevap)
        try:
            self._pdf_kaydet(ad)
        except Exception as e:
            messagebox.showerror("Hata", f"Kayıt başarısız:\n{e}")
            return

        self.idx += 1
        self._yukle()

    def atla(self):
        ad = f"personel_{self.idx+1:02d}"
        try:
            self._pdf_kaydet(ad)
            print(f"  → atlandı, '{ad}' olarak kaydedildi")
        except Exception as e:
            messagebox.showerror("Hata", f"Kayıt başarısız:\n{e}")
            return

        self.idx += 1
        self._yukle()

    def cik(self):
        self.root.destroy()


def main() -> None:
    if not os.path.isfile(GIRIS_PDF):
        print(f"HATA: '{GIRIS_PDF}' bulunamadı. GIRIS_PDF değişkenini düzelt.")
        sys.exit(1)

    os.makedirs(CIKTI_KLASORU, exist_ok=True)

    reader = PdfReader(GIRIS_PDF)
    pdf_render = pdfium.PdfDocument(GIRIS_PDF)

    toplam = len(reader.pages)
    if toplam % 2 != 0:
        print(f"UYARI: Toplam sayfa tek sayı ({toplam}). Son sayfa atlanacak.")

    cift_sayisi = toplam // 2
    print(f"📄 Toplam {toplam} sayfa → {cift_sayisi} personel\n")

    Uygulama(reader, pdf_render, cift_sayisi)

    print(f"\n✅ Bitti. PDF'ler '{CIKTI_KLASORU}/' klasöründe.")


if __name__ == "__main__":
    main()
