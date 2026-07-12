# -*- coding: utf-8 -*-
"""
KKD ZİMMET — Excel'e PERSONEL ADI yazma + işlenenleri ayırma yardımcısı.
========================================================================
Adı zaten "AD SOYAD.pdf" olarak konmuş KKD zimmet PDF'lerini SIRAYLA gezer.
Ekranın EN ÜSTÜNDE (Excel'den fokus çalmadan, hep üstte) o PDF'in
"KİŞİSEL KORUYUCU EKİPMAN LİSTESİ" tablosunu gösterir — sen ADET/BEDEN'i
oradan okuyup Excel'e yazarsın.

SAĞ SHIFT  → sadece personelin ADINI yazar — tuş dizisi (aşağıda ISLEM_DIZISI):
        AD ⏎  AD ⏎  AD ⏎  →  ↑ ↑ ↑
   (Excel'de: adı 3 satıra yazar, altta 1 BOŞ satır bırakır, sonra sağ
    sütunun en üst satırına döner.)
   Dosyayı TAŞIMAZ, sonrakine GEÇMEZ — istersen tekrar basabilirsin.

SAĞ CTRL   → o PDF'i "İŞLENDİ" alt klasörüne atar ve SIRADAKİ personele geçer.
   (Yani önce ADET/BEDEN'i yaz + Sağ Shift ile adı bas; bittiğinde Sağ Ctrl.)

TUŞLAR (Excel'deyken bile çalışır — global):
  Sağ Shift              → adı yaz
  Sağ Ctrl               → İŞLENDİ'ye at + sonraki
  Şeritteki  Atla        → işlemeden (taşımadan) sonrakine geç
  Şeritteki  ⟲ Geri Al   → son işleneni geri getir (yanlış yaptıysan)
  Şeritteki  TEST        → boş bir Excel hücresine örnek Türkçe ad yazıp dener
  Şeritteki  ✕           → çıkış
  Şeridi taşı: üstteki koyu bilgi çubuğunu fareyle tut, yukarı/aşağı sürükle

NOT: OCR YOK — ad doğrudan dosya adından alınır (zaten kkd_adlandir ile konmuştu).
GÜVENLİK: Yazmadan önce DOĞRU Excel hücresini seçili tut. Yanlış olursa
"Geri Al" ile PDF geri gelir; Excel'deki yanlışı Ctrl+Z ile alırsın.
"""

import os
import re
import sys
import time
import queue
import ctypes
from ctypes import wintypes
import tkinter as tk
from tkinter import filedialog, messagebox

import pypdfium2 as pdfium
from PIL import Image, ImageTk
from pynput import keyboard

# ============================ AYARLAR ============================
BASLANGIC_KLASORU = os.path.join(os.path.expanduser("~"), "Desktop", "TARAMA")   # Klasör seçme penceresi burada açılır (yoksa çalışma klasörü)
ISLENDI_KLASOR    = "İŞLENDİ"     # İşlenen PDF'lerin taşınacağı alt klasör adı
RENDER_OLCEK      = 2.0           # Tablo görüntüsü çözünürlüğü

# --- Tablo kırpma (sayfa yüksekliğine oran) — bu form şablonuna göre ayarlı ---
TABLO_UST_ORAN = 0.575           # Tablonun üstü (KİŞİSEL KORUYUCU ... başlığı)
TABLO_ALT_ORAN = 0.790           # Tablonun altı (satırlar bitince biraz pay)

# --- Sağ Shift'e basınca gönderilecek TUŞ DİZİSİ ---
#  ("ad",  None)      -> personelin adını yaz
#  ("tus", "enter" | "right" | "left" | "up" | "down" | "tab") -> o tuşa bas
# İstediğin gibi düzenle. Varsayılan: AD ⏎ AD ⏎ AD ⏎ → ↑ ↑ ↑
# (3. addan sonra fazladan bir ⏎ → ad sütununda altta bir BOŞ satır kalır;
#  sonra sağ + 3 kez yukarı ile imleç yine yan sütunun en üst satırına döner.)
ISLEM_DIZISI = [
    ("ad", None),  ("tus", "enter"),
    ("ad", None),  ("tus", "enter"),
    ("ad", None),  ("tus", "enter"),
    ("tus", "right"), ("tus", "up"), ("tus", "up"), ("tus", "up"),
]

# --- Ad biçimi ---  "aynen" | "buyuk" | "baslik"
AD_BICIM = "aynen"               # aynen = dosya adındaki gibi (BÜYÜK HARF)

# --- Zamanlama (Excel yetişemezse ARTIR) ---
BASLAMA_GECIKME = 0.08           # Shift bırakıldıktan sonra yazmaya başlamadan önce (sn)
HARF_GECIKME    = 0.004          # Harfler arası (sn)
TUS_GECIKME     = 0.03           # Adımlar (Enter/ok) arası (sn)

# Ad yazmadan ÖNCE bir kez ESC gönder: Excel'de açık kalmış "Yapıştırma
# Seçenekleri" menüsü / popup varsa kapanır, ad temiz hücreye yazılır.
ONCE_ESC = True

# --- Görünüm ---
MAKS_YUK_ORANI = 0.45            # Şerit en fazla ekran yüksekliğinin bu kadarı
BAR_YUK        = 34
DENEME_ISIM    = "ÇĞİÖŞÜ ışİ TEST"   # TEST butonunun yazacağı örnek
# ================================================================


# ---------------- Windows: hep üstte + fokus çalmama ----------------
_user32 = ctypes.windll.user32
_user32.GetWindowLongW.restype = ctypes.c_long
_user32.GetWindowLongW.argtypes = [ctypes.c_void_p, ctypes.c_int]
_user32.SetWindowLongW.restype = ctypes.c_long
_user32.SetWindowLongW.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_long]
_user32.GetParent.restype = ctypes.c_void_p
_user32.GetParent.argtypes = [ctypes.c_void_p]
_user32.SetWindowPos.argtypes = [ctypes.c_void_p, ctypes.c_void_p,
                                 ctypes.c_int, ctypes.c_int, ctypes.c_int,
                                 ctypes.c_int, ctypes.c_uint]
_GWL_EXSTYLE = -20
_WS_EX_NOACTIVATE = 0x08000000
_WS_EX_TOOLWINDOW = 0x00000080
_HWND_TOPMOST = -1
_SWP_NOSIZE = 0x0001
_SWP_NOMOVE = 0x0002
_SWP_NOACTIVATE = 0x0010


def _hwnd(root):
    root.update_idletasks()
    return _user32.GetParent(root.winfo_id()) or root.winfo_id()


def _fokus_alma(root):
    h = _hwnd(root)
    ex = _user32.GetWindowLongW(h, _GWL_EXSTYLE)
    _user32.SetWindowLongW(h, _GWL_EXSTYLE, ex | _WS_EX_NOACTIVATE | _WS_EX_TOOLWINDOW)


def _hep_ustte(root):
    h = _hwnd(root)
    _user32.SetWindowPos(h, ctypes.c_void_p(_HWND_TOPMOST), 0, 0, 0, 0,
                         _SWP_NOMOVE | _SWP_NOSIZE | _SWP_NOACTIVATE)


# ---------------- SendInput: Türkçe dahil her karakteri yaz ----------------
ULONG_PTR = wintypes.WPARAM


class _KEYBDINPUT(ctypes.Structure):
    _fields_ = [("wVk", wintypes.WORD), ("wScan", wintypes.WORD),
                ("dwFlags", wintypes.DWORD), ("time", wintypes.DWORD),
                ("dwExtraInfo", ULONG_PTR)]


class _INUNION(ctypes.Union):
    # MOUSEINPUT en büyük üye (x64'te 32 byte). cbSize'ın doğru çıkması için
    # union'ı 32 byte'a tamamla; aksi halde SendInput sessizce başarısız olur.
    _fields_ = [("ki", _KEYBDINPUT), ("_pad", ctypes.c_ubyte * 32)]


class _INPUT(ctypes.Structure):
    _fields_ = [("type", wintypes.DWORD), ("u", _INUNION)]


_user32.SendInput.argtypes = [wintypes.UINT, ctypes.c_void_p, ctypes.c_int]
_user32.SendInput.restype = wintypes.UINT

_INPUT_KEYBOARD = 1
_KEYEVENTF_KEYUP = 0x0002
_KEYEVENTF_UNICODE = 0x0004
_KEYEVENTF_EXTENDED = 0x0001
_VK = {"enter": 0x0D, "tab": 0x09, "escape": 0x1B, "left": 0x25,
       "up": 0x26, "right": 0x27, "down": 0x28}
_EXTENDED = {"left", "up", "right", "down"}


def _ki(wVk=0, wScan=0, dwFlags=0):
    return _INPUT(type=_INPUT_KEYBOARD, u=_INUNION(ki=_KEYBDINPUT(wVk, wScan, dwFlags, 0, 0)))


def _gonder(inputs):
    arr = (_INPUT * len(inputs))(*inputs)
    n = _user32.SendInput(len(inputs), ctypes.byref(arr), ctypes.sizeof(_INPUT))
    if n != len(inputs):
        err = ctypes.GetLastError()
        print(f"UYARI: SendInput {n}/{len(inputs)} gönderdi (hata {err}).")


def yaz_karakter(ch):
    """Tek bir karakteri (unicode) aktif pencereye yaz. Klavye düzeninden bağımsız."""
    c = ord(ch)
    _gonder([_ki(wScan=c, dwFlags=_KEYEVENTF_UNICODE),
             _ki(wScan=c, dwFlags=_KEYEVENTF_UNICODE | _KEYEVENTF_KEYUP)])


def bas_tus(isim):
    """enter/ok tuşlarına bas."""
    vk = _VK[isim]
    f = _KEYEVENTF_EXTENDED if isim in _EXTENDED else 0
    _gonder([_ki(wVk=vk, dwFlags=f),
             _ki(wVk=vk, dwFlags=f | _KEYEVENTF_KEYUP)])


# ---------------- Yardımcılar ----------------
def ad_bicimle(ad):
    if AD_BICIM == "buyuk":
        return ad.upper()
    if AD_BICIM == "baslik":
        return " ".join(k[:1].upper().replace("I", "İ") + k[1:].lower().replace("i̇", "i")
                        for k in ad.split() if k)
    return ad


def benzersiz_yol(klasor, dosya_adi):
    yol = os.path.join(klasor, dosya_adi)
    kok, uz = os.path.splitext(dosya_adi)
    s = 2
    while os.path.exists(yol):
        yol = os.path.join(klasor, f"{kok}_{s}{uz}")
        s += 1
    return yol


# ---------------- Uygulama ----------------
class KkdIsle:
    def __init__(self, klasor):
        self.klasor = klasor
        self.islendi_dir = os.path.join(klasor, ISLENDI_KLASOR)
        self.dosyalar = sorted(
            f for f in os.listdir(klasor)
            if f.lower().endswith(".pdf") and os.path.isfile(os.path.join(klasor, f))
        )
        if not self.dosyalar:
            raise RuntimeError(f"Bu klasörde PDF yok:\n{klasor}")

        self.idx = 0
        self.undo_stack = []          # [(idx, eski_yol, yeni_yol)]
        self.isleniyor = False
        self.komut_q = queue.Queue()
        self._td = {"shift": False, "ctrl": False, "shift_combo": False, "ctrl_combo": False}

        self._arayuz()
        self._dinleyici()
        self._yukle()
        self.root.after(40, self._pump)
        self.root.after(1500, self._ustte_dongu)
        self.root.mainloop()

    # ---- Arayüz ----
    def _arayuz(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self.ekran_w = self.root.winfo_screenwidth()
        self.ekran_h = self.root.winfo_screenheight()
        self.disp_w = self.ekran_w
        self.disp_h = int(self.ekran_h * MAKS_YUK_ORANI)

        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.pencere_y = 0
        self.root.geometry(f"{self.ekran_w}x{self.disp_h + BAR_YUK}+0+{self.pencere_y}")

        bar = tk.Frame(self.root, bg="#1f1f1f", height=BAR_YUK)
        bar.pack(fill="x")
        bar.pack_propagate(False)

        self.bilgi = tk.Label(bar, text="", bg="#1f1f1f", fg="white",
                              font=("Segoe UI", 11, "bold"), anchor="w", padx=12)
        self.bilgi.pack(side="left", fill="x", expand=True)

        ipucu = tk.Label(bar, text="Sağ Shift = Adı Yaz    Sağ Ctrl = İşlendi ▶    (çubuğu sürükle ↕)",
                         bg="#1f1f1f", fg="#9ad", font=("Segoe UI", 9), padx=8)
        ipucu.pack(side="left")

        cik = tk.Label(bar, text="  ✕  ", bg="#b33", fg="white",
                       font=("Segoe UI", 11, "bold"), cursor="hand2")
        cik.pack(side="right", fill="y")
        cik.bind("<Button-1>", lambda e: self._kapat())

        test = tk.Label(bar, text="  TEST  ", bg="#555", fg="white",
                        font=("Segoe UI", 10, "bold"), cursor="hand2")
        test.pack(side="right", fill="y")
        test.bind("<Button-1>", lambda e: self.komut_q.put("test"))

        geri = tk.Label(bar, text="  ⟲ Geri Al  ", bg="#2a7abf", fg="white",
                        font=("Segoe UI", 10, "bold"), cursor="hand2")
        geri.pack(side="right", fill="y")
        geri.bind("<Button-1>", lambda e: self.komut_q.put("undo"))

        atla = tk.Label(bar, text="  Atla  ", bg="#777", fg="white",
                        font=("Segoe UI", 10, "bold"), cursor="hand2")
        atla.pack(side="right", fill="y")
        atla.bind("<Button-1>", lambda e: self.komut_q.put("atla"))

        # Bilgi çubuğunu fareyle tutup yukarı/aşağı sürükleyerek şeridi taşı
        for w in (bar, self.bilgi, ipucu):
            w.bind("<Button-1>", self._suru_basla)
            w.bind("<B1-Motion>", self._suru)

        self.gorsel = tk.Label(self.root, bg="#f5f5f5")
        self.gorsel.pack(fill="both", expand=True)

        self.root.deiconify()
        self.root.update_idletasks()
        _fokus_alma(self.root)
        _hep_ustte(self.root)

    # ---- İsim (dosya adından) ----
    def _isim(self, idx):
        kok = os.path.splitext(self.dosyalar[idx])[0]
        kok = re.sub(r"_\d+$", "", kok)          # _2, _3 (çakışma eki) at
        return ad_bicimle(kok).strip()

    # ---- Tablo görüntüsü ----
    def _tablo_img(self, path):
        with open(path, "rb") as f:               # bytes ile aç (Türkçe yol sorununu aşar)
            data = f.read()
        pdf = pdfium.PdfDocument(data)
        img = pdf[0].render(scale=RENDER_OLCEK).to_pil().convert("RGB")
        pdf.close()
        w, h = img.size
        return img.crop((0, int(h * TABLO_UST_ORAN), w, int(h * TABLO_ALT_ORAN)))

    def _frame(self, img):
        cw, ch = img.size
        olcek = min(self.disp_w / cw, self.disp_h / ch)
        nw, nh = max(1, int(cw * olcek)), max(1, int(ch * olcek))
        img = img.resize((nw, nh), Image.LANCZOS)
        tuval = Image.new("RGB", (self.disp_w, self.disp_h), (245, 245, 245))
        tuval.paste(img, ((self.disp_w - nw) // 2, (self.disp_h - nh) // 2))
        return tuval

    # ---- Şeridi sürükleyerek taşı ----
    def _suru_basla(self, e):
        self._suru_y0 = e.y_root
        self._pencere_y0 = self.pencere_y

    def _suru(self, e):
        self.pencere_y = max(0, self._pencere_y0 + (e.y_root - self._suru_y0))
        self.root.geometry(f"+0+{self.pencere_y}")

    def _yukle(self):
        if self.idx >= len(self.dosyalar):
            messagebox.showinfo("Bitti",
                                f"Tüm personeller işlendi. ✅\n\n"
                                f"İşlenenler: '{ISLENDI_KLASOR}' alt klasöründe.")
            self._kapat()
            return

        ad = self._isim(self.idx)
        path = os.path.join(self.klasor, self.dosyalar[self.idx])
        self.bilgi.config(text=f"[{self.idx+1}/{len(self.dosyalar)}]   👤 {ad}")

        try:
            img = self._tablo_img(path)
            self.tk_img = ImageTk.PhotoImage(self._frame(img))
            self.gorsel.config(image=self.tk_img)
        except Exception as e:
            self.gorsel.config(image="", text=f"({self.dosyalar[self.idx]} okunamadı: {e})")

    # ---- Komutlar ----
    def _dizi_gonder(self, ad):
        time.sleep(BASLAMA_GECIKME)
        if ONCE_ESC:
            bas_tus("escape")          # açık kalmış yapıştırma menüsü/popup'ı kapat
            time.sleep(TUS_GECIKME)
        for tip, val in ISLEM_DIZISI:
            if tip == "ad":
                for ch in ad:
                    yaz_karakter(ch)
                    if HARF_GECIKME:
                        time.sleep(HARF_GECIKME)
            else:
                bas_tus(val)
            time.sleep(TUS_GECIKME)

    def _yaz(self):
        """SAĞ SHIFT: yalnızca adı yazar. Dosyayı taşımaz, sonrakine geçmez."""
        if self.isleniyor or self.idx >= len(self.dosyalar):
            return
        self.isleniyor = True
        try:
            ad = self._isim(self.idx)
            self._dizi_gonder(ad)                      # Excel'e yaz
            print(f"✎ yazıldı: {ad}")
        except Exception as e:
            messagebox.showerror("Hata", f"Yazma başarısız:\n{e}")
        finally:
            self.isleniyor = False

    def _islendi(self):
        """SAĞ CTRL: o PDF'i İŞLENDİ'ye taşır ve sıradakine geçer."""
        if self.isleniyor or self.idx >= len(self.dosyalar):
            return
        self.isleniyor = True
        try:
            os.makedirs(self.islendi_dir, exist_ok=True)
            eski = os.path.join(self.klasor, self.dosyalar[self.idx])
            yeni = benzersiz_yol(self.islendi_dir, self.dosyalar[self.idx])
            os.rename(eski, yeni)
            self.undo_stack.append((self.idx, eski, yeni))
            print(f"✓ {self._isim(self.idx)}  →  {ISLENDI_KLASOR}/")
            self.idx += 1
            self._yukle()
        except Exception as e:
            messagebox.showerror("Hata", f"İşlendi'ye taşıma başarısız:\n{e}")
        finally:
            self.isleniyor = False

    def _atla(self):
        if self.idx < len(self.dosyalar):
            print(f"⊘ Atlandı: {self.dosyalar[self.idx]}")
            self.idx += 1
            self._yukle()

    def _undo(self):
        if not self.undo_stack:
            return
        idx, eski, yeni = self.undo_stack.pop()
        try:
            os.rename(yeni, eski)         # İŞLENDİ'den geri getir
            self.idx = idx
            print(f"⟲ Geri alındı: {self.dosyalar[idx]}")
            self._yukle()
        except Exception as e:
            messagebox.showerror("Hata", f"Geri alma başarısız:\n{e}")

    def _test(self):
        """Boş bir hücreye örnek Türkçe ad + Enter yazar (Türkçe karakter denemesi)."""
        if self.isleniyor:
            return
        self.isleniyor = True
        try:
            time.sleep(BASLAMA_GECIKME)
            for ch in DENEME_ISIM:
                yaz_karakter(ch)
                time.sleep(HARF_GECIKME)
            bas_tus("enter")
        finally:
            self.isleniyor = False

    # ---- Global tuş dinleyici ----
    def _dinleyici(self):
        self.listener = keyboard.Listener(on_press=self._bas, on_release=self._birak)
        self.listener.daemon = True
        self.listener.start()

    def _bas(self, key):
        d = self._td
        if key == keyboard.Key.shift_r:
            d["shift"] = True; d["shift_combo"] = False
        elif key == keyboard.Key.ctrl_r:
            d["ctrl"] = True; d["ctrl_combo"] = False
        else:
            if d["shift"]: d["shift_combo"] = True
            if d["ctrl"]: d["ctrl_combo"] = True
        if d["shift"] and d["ctrl"]:
            # İki tuş yanlışlıkla birlikte basılırsa: hiçbir işlem yapma.
            # (Kapatma kısayolu kapalı — kapatmak için sağ üstteki ✕ butonu.)
            d["shift_combo"] = d["ctrl_combo"] = True

    def _birak(self, key):
        d = self._td
        if key == keyboard.Key.shift_r:
            if d["shift"] and not d["shift_combo"]:
                self.komut_q.put("yaz")            # Sağ Shift = adı yaz
            d["shift"] = False
        elif key == keyboard.Key.ctrl_r:
            if d["ctrl"] and not d["ctrl_combo"]:
                self.komut_q.put("islendi")        # Sağ Ctrl = işlendi + sonraki
            d["ctrl"] = False

    # ---- Ana thread: komut kuyruğu ----
    def _pump(self):
        try:
            while True:
                c = self.komut_q.get_nowait()
                if c == "yaz":       self._yaz()
                elif c == "islendi": self._islendi()
                elif c == "atla":    self._atla()
                elif c == "undo":    self._undo()
                elif c == "test":    self._test()
                elif c == "quit":    self._kapat(); return
        except queue.Empty:
            pass
        self.root.after(40, self._pump)

    def _ustte_dongu(self):
        _hep_ustte(self.root)
        self.root.after(1500, self._ustte_dongu)

    def _kapat(self):
        try:
            self.listener.stop()
        except Exception:
            pass
        try:
            self.root.destroy()
        except Exception:
            pass


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    kok = tk.Tk(); kok.withdraw()
    if args:
        klasor = args[0]
    else:
        klasor = filedialog.askdirectory(
            title="KKD zimmet PDF'lerinin (AD SOYAD.pdf) olduğu klasörü seç",
            initialdir=BASLANGIC_KLASORU if os.path.isdir(BASLANGIC_KLASORU) else os.getcwd())
    kok.destroy()
    if not klasor:
        print("Klasör seçilmedi, çıkılıyor.")
        return
    try:
        KkdIsle(klasor)
    except RuntimeError as e:
        r = tk.Tk(); r.withdraw()
        messagebox.showwarning("KKD İşle", str(e)); r.destroy()
    print("\n✅ Kapatıldı.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        print(tb)
        try:
            r = tk.Tk(); r.withdraw()
            messagebox.showerror("KKD İşle — Hata", f"{e}\n\n{tb}"); r.destroy()
        except Exception:
            pass
