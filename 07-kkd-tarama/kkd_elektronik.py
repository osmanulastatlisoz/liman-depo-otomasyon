"""
ELEKTRONİK CİHAZ ZİMMET TUTANAĞI - PDF isimlendirme.

Bir klasördeki isimlendirilmemiş PDF'leri (IMG_0001.pdf, IMG_0002.pdf ...)
tek tek açar, üst kısmı (cihaz adı + AD SOYADI + GÖREV alanı) küçük bir
pencerede ekranın üst kısmında gösterir. Sen ismi yazarsın, dosya yeniden
adlandırılır.

İpucu: Dosya adına cihazı da ekleyebilirsin, örn: "Gülcan Yazıcı - Monitör"

Kullanım:
  1. Bu script'i PDF'lerin olduğu klasöre koy (veya KLASOR değişkenini ayarla).
  2. python kkd_elektronik.py

Pencerede:
  - İsmi yaz + Enter      → dosyayı yeniden adlandır, sonrakine geç
  - Enter (boş)           → atla (dosya adı değişmez)
  - Atla butonu / Esc     → atla
  - Çık butonu            → kapat

Not: Dosyalar YERİNDE yeniden adlandırılır. Yedek istersen klasörü
önceden kopyala.
"""

import os
import re
import sys
import tkinter as tk
from tkinter import messagebox
from pypdf import PdfReader
import pypdfium2 as pdfium
from PIL import Image, ImageTk

# ---------- AYARLAR ----------
KLASOR = "."                    # PDF'lerin olduğu klasör. "." = script'in olduğu klasör
ON_IZLEME_ORANI = 0.20          # Sayfanın üstten ne kadarı gösterilsin (elektronik formlar için biraz daha geniş)
RENDER_OLCEK = 2.0              # Görüntü çözünürlüğü
PENCERE_GENISLIK = 1000         # Pencere genişliği (px)
PENCERE_UST_OFFSET = 10         # Ekran üstünden uzaklık (px)
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


def kirpilmis_onizleme(pdf_yolu: str) -> Image.Image:
    """PDF'in ilk sayfasının üst kısmını kırpılmış görüntü olarak döndürür."""
    pdf = pdfium.PdfDocument(pdf_yolu)
    page = pdf[0]
    img = page.render(scale=RENDER_OLCEK).to_pil()
    w, h = img.size
    kirpilan = img.crop((0, 0, w, int(h * ON_IZLEME_ORANI)))
    pdf.close()
    return kirpilan


class Uygulama:
    def __init__(self, dosyalar: list[str], klasor: str):
        self.dosyalar = dosyalar
        self.klasor = klasor
        self.idx = 0

        self.root = tk.Tk()
        self.root.title("Elektronik Cihaz Zimmet - PDF İsimlendirme")

        # Pencereyi ekranın üst kısmına yerleştir
        ekran_genislik = self.root.winfo_screenwidth()
        x = (ekran_genislik - PENCERE_GENISLIK) // 2
        self.root.geometry(f"{PENCERE_GENISLIK}x440+{x}+{PENCERE_UST_OFFSET}")
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

        # İlk dosyayı yükle
        self._yukle()
        self.root.mainloop()

    def _yukle(self):
        if self.idx >= len(self.dosyalar):
            messagebox.showinfo("Bitti", "Tüm dosyalar işlendi. ✅")
            self.root.destroy()
            return

        dosya_adi = self.dosyalar[self.idx]
        tam_yol = os.path.join(self.klasor, dosya_adi)

        self.bilgi_etiketi.config(
            text=f"[{self.idx+1}/{len(self.dosyalar)}]  {dosya_adi}"
        )

        try:
            on_izleme = kirpilmis_onizleme(tam_yol)
            # Pencere genişliğine ölçekle
            oran = (PENCERE_GENISLIK - 40) / on_izleme.width
            yeni_boyut = (
                int(on_izleme.width * oran),
                int(on_izleme.height * oran),
            )
            on_izleme = on_izleme.resize(yeni_boyut, Image.LANCZOS)
            self.tk_img = ImageTk.PhotoImage(on_izleme)
            self.gorsel_etiketi.config(image=self.tk_img)
        except Exception as e:
            messagebox.showerror("Hata", f"{dosya_adi} okunamadı:\n{e}")
            self.idx += 1
            self._yukle()
            return

        self.entry.delete(0, tk.END)
        self.entry.focus()

    def kaydet(self):
        ad = self.entry.get().strip()
        if not ad:
            self.atla()
            return

        ad = temizle_dosya_adi(ad)
        eski_yol = os.path.join(self.klasor, self.dosyalar[self.idx])
        yeni_yol = benzersiz_yol(self.klasor, ad)

        try:
            os.rename(eski_yol, yeni_yol)
            print(f"✓ {self.dosyalar[self.idx]} → {os.path.basename(yeni_yol)}")
        except Exception as e:
            messagebox.showerror("Hata", f"Yeniden adlandırma başarısız:\n{e}")
            return

        self.idx += 1
        self._yukle()

    def atla(self):
        print(f"⊘ Atlandı: {self.dosyalar[self.idx]}")
        self.idx += 1
        self._yukle()

    def cik(self):
        self.root.destroy()


def main():
    klasor = os.path.abspath(KLASOR)
    if not os.path.isdir(klasor):
        print(f"HATA: Klasör bulunamadı: {klasor}")
        sys.exit(1)

    # PDF dosyalarını topla (alfabetik sıraya göre)
    dosyalar = sorted(
        f for f in os.listdir(klasor)
        if f.lower().endswith(".pdf") and os.path.isfile(os.path.join(klasor, f))
    )

    if not dosyalar:
        print(f"HATA: '{klasor}' içinde PDF dosyası yok.")
        sys.exit(1)

    print(f"📄 {len(dosyalar)} PDF bulundu: {klasor}\n")
    Uygulama(dosyalar, klasor)
    print("\n✅ Tamamlandı.")


if __name__ == "__main__":
    main()
