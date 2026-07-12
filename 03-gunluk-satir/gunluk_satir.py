"""
GÜNLÜK VERİLEN MALZEME - Excel'e girerken SATIR SATIR önizleme şeridi.

Taranan "Günlük Verilen Malzeme Listesi" PDF'ini açar, sayfalardaki tablo
satırlarını otomatik bulur ve ekranın EN ÜSTÜNDE, hep üstte (always-on-top),
Excel'den FOKUS ÇALMAYAN ince bir şerit olarak gösterir.

Sen Excel'de yazmaya devam edersin; şerit yukarıda durur. Aktif satır net,
üst/alttaki komşu satırlar soluk görünür. Taşan el yazısı kesilmesin diye
üstten/alttan biraz pay bırakılır.

KULLANIM:
  1. python gunluk_satir.py
  2. Açılan pencereden PDF'i seç.
  3. Excel'e geç, yazmaya başla. Tuşlar Excel'deyken bile çalışır:

  TUŞLAR (global):
    Sağ Shift           → sıradaki satır (+ açıksa Excel'de kayıtlı tuş dizisi)
    Sol Shift + Sağ Shift → sıradaki satır AMA Excel'e tuş gitmez (tek seferlik)
    Sağ Ctrl            → önceki satır
    Şeritteki  ⌨ Ayarla  → Excel'e gönderilecek tuş dizisini kaydet (mini makro)
    Şeritteki  ✔ Bitti   → dosya adına "(işlendi)" ekle (bitenler belli olsun)
    Son satırda tekrar Sağ Shift → "belge bitti mi?" diye sorar, evet dersen
    dosya adına "(işlendi)" eklenir.
    Şeritteki  ✕  → çıkış
    Şeritteki bilgi çubuğunu fareyle tutup yukarı/aşağı sürükle → şeridi taşı

  Not: "Sağ Shift"i temiz basışta algılar. Excel'de Shift+Ok ile seçim
  yaparken yanlışlıkla ilerlemez (Sol Shift'i seçim için kullanmaya devam
  edebilirsin).

KALİBRASYON:
  Satır algılama PDF'e göre biraz şaşarsa aşağıdaki AYARLAR bloğundaki
  sayılarla oynayarak hizalanır. Çalıştırınca konsola sayfa başına bulunan
  satır sayısı yazılır; gerçeğiyle (≈26-30) karşılaştır.
"""

import os
import sys
import json
import queue
import statistics
import ctypes
import tkinter as tk
from tkinter import filedialog, messagebox

import numpy as np
import pypdfium2 as pdfium
from PIL import Image, ImageTk, ImageDraw
from pynput import keyboard

# ============================ AYARLAR ============================
BASLANGIC_KLASORU = os.path.join(os.path.expanduser("~"), "Desktop", "TARAMA")  # Dosya seçme penceresi burada açılır (yoksa çalışma klasörü)
RENDER_OLCEK      = 5.0      # PDF render çözünürlüğü (yükseltirsen daha net ama yavaş)

# --- Satır (çizgi) algılama ---
KOYU_ESIK    = 160          # Bu değerden koyu pikseller "dolu" sayılır (0-255). Silik çizgi için YÜKSELT
CIZGI_ORANI  = 0.45         # Bir y satırı, genişliğin bu kadarı koyuysa "yatay çizgi"dir
MIN_SATIR_PX = 26           # Bundan kısa bantlar (gürültü) atlanır  (render ölçeğinde px)
MAX_SATIR_PX = 130          # Bundan uzun bantlar (başlık bloğu vb.) atlanır
ILK_SATIR_ATLA = 1          # Her sayfanın başındaki başlık satırı sayısı (sütun başlıkları)
KENAR_UST    = 0.06         # Üst kenar payı (oran): bu sınıra kadar eksik satırları tamamla
KENAR_ALT    = 0.975        # Alt kenar payı (oran)

# --- Otomatik eğiklik düzeltme (deskew) ---
OTO_DUZELT     = True        # Eğik taranan sayfaları otomatik düzelt (sayfa atlanmasını önler)
DUZELT_ARALIK  = 8.0         # +/- kaç dereceye kadar denesin (makul tampon; çok büyütme!)
DUZELT_ADIM    = 0.10        # arama adımı (derece) — küçük = daha hassas düzeltme
DUZELT_ESIK    = 0.20        # bu açıdan küçük eğiklikte sayfayı döndürme (gereksiz)

# --- Yedek bölme (otomatik algılama tutmazsa) ---
YEDEK_BOL    = False        # True yaparsan algılamayı yok sayar, eşit böler
YEDEK_SATIR  = 28           # Eşit bölmede sayfa başına satır
YEDEK_UST    = 0.10         # Tablo alanı üst sınırı (sayfa yüksekliğine oran)
YEDEK_ALT    = 0.97         # Tablo alanı alt sınırı

# --- Görünüm ---
PAY_PX        = 8           # Şeridin en üst/altına bırakılan kırpma payı (px)
NET_PAY_PX    = 20          # Aktif satırın ÜST/ALTINA eklenen NET (soluk olmayan) pay. ~43px ≈ 1 satır. Net alanı büyütmek için ARTIR
SOLUK_ALFA    = 0.9        # Komşu satırların solukluğu (0=net ... 1=bembeyaz)
MAKS_YUK_ORANI = 0.42       # Şerit en fazla ekran yüksekliğinin bu kadarı olur
CERCEVE_RENK  = (0, 120, 215)  # Aktif satır çerçeve rengi
CERCEVE_KALIN = 5
BAR_YUK       = 30          # Üstteki bilgi çubuğu yüksekliği (px)

# --- Tuşlar ---
ILERI_TUS = keyboard.Key.shift_r   # sıradaki satır
GERI_TUS  = keyboard.Key.ctrl_r    # önceki satır

# --- İleri'de Excel'e otomatik imleç hareketi ---
# Sağ Shift ile sıradaki satıra geçerken Excel'de de imleci oynatır.
# Program açıkken üstteki "Excel" düğmesiyle açıp kapatabilirsin.
EXCEL_ILERI_AKTIF  = True                      # Başlangıçta açık mı?
EXCEL_ILERI_TUSLAR = ["down", "left", "left"]  # BAŞLANGIÇ dizisi (üstteki "⌨ Ayarla" ile değiştirilebilir)
# Kullanılabilen adlar: up, down, left, right, enter, tab, home, end,
#                       page_up, page_down, esc, space, backspace, delete

# Bu tuş BASILI tutulurken Sağ Shift'e basarsan, şeritte ilerler ama Excel'e
# tuş GÖNDERMEZ (tek seferlik atlama). Sol Shift seçildi çünkü Excel'de tek
# başına basılınca hiçbir şeyi tetiklemez. İstersen keyboard.Key.ctrl_l yapabilirsin.
EXCEL_BASTIR_TUS   = keyboard.Key.shift_l
# ================================================================

# İsim -> pynput tuşu (Excel'e gönderilen tuşlar için)
_TUS_HARITA = {
    "up": keyboard.Key.up, "down": keyboard.Key.down,
    "left": keyboard.Key.left, "right": keyboard.Key.right,
    "enter": keyboard.Key.enter, "tab": keyboard.Key.tab,
    "home": keyboard.Key.home, "end": keyboard.Key.end,
    "page_up": keyboard.Key.page_up, "page_down": keyboard.Key.page_down,
    "esc": keyboard.Key.esc, "space": keyboard.Key.space,
    "backspace": keyboard.Key.backspace, "delete": keyboard.Key.delete,
}


def _bastir_mi(key) -> bool:
    """Basılı tutulunca Excel'e tuş göndermeyi durduran tuş mu? (pynput bazı
    sürümlerde sol shift/ctrl'ü genel Key.shift/Key.ctrl olarak bildirir.)"""
    if key == EXCEL_BASTIR_TUS:
        return True
    if EXCEL_BASTIR_TUS in (keyboard.Key.shift_l, keyboard.Key.shift) \
            and key in (keyboard.Key.shift_l, keyboard.Key.shift):
        return True
    if EXCEL_BASTIR_TUS in (keyboard.Key.ctrl_l, keyboard.Key.ctrl) \
            and key in (keyboard.Key.ctrl_l, keyboard.Key.ctrl):
        return True
    return False


# Tkinter keysym -> ad (⌨ Ayarla kayıt penceresinde yakalanan tuşlar)
_TK_AD = {
    "Up": "up", "Down": "down", "Left": "left", "Right": "right",
    "KP_Up": "up", "KP_Down": "down", "KP_Left": "left", "KP_Right": "right",
    "Return": "enter", "KP_Enter": "enter", "Tab": "tab",
    "Home": "home", "End": "end", "KP_Home": "home", "KP_End": "end",
    "Prior": "page_up", "Next": "page_down",
    "KP_Prior": "page_up", "KP_Next": "page_down",
    "space": "space", "BackSpace": "backspace", "Delete": "delete",
}

# ad -> ekranda gösterilecek simge
_TUS_SIMGE = {
    "up": "↑", "down": "↓", "left": "←", "right": "→",
    "enter": "⏎", "tab": "⇥", "home": "Home", "end": "End",
    "page_up": "PgUp", "page_down": "PgDn",
    "space": "␣", "backspace": "⌫", "delete": "Del",
}


# ---- Windows: pencereyi fokus çalmadan hep üstte tutma yardımcıları ----
_user32 = ctypes.windll.user32
_user32.GetWindowLongW.restype = ctypes.c_long
_user32.GetWindowLongW.argtypes = [ctypes.c_void_p, ctypes.c_int]
_user32.SetWindowLongW.restype = ctypes.c_long
_user32.SetWindowLongW.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_long]
_user32.GetParent.restype = ctypes.c_void_p
_user32.GetParent.argtypes = [ctypes.c_void_p]
_user32.SetWindowPos.argtypes = [
    ctypes.c_void_p, ctypes.c_void_p,
    ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_uint,
]

_GWL_EXSTYLE      = -20
_WS_EX_NOACTIVATE = 0x08000000
_WS_EX_TOOLWINDOW = 0x00000080
_HWND_TOPMOST     = -1
_SWP_NOSIZE       = 0x0001
_SWP_NOMOVE       = 0x0002
_SWP_NOACTIVATE   = 0x0010


def _hwnd(root: tk.Tk) -> int:
    root.update_idletasks()
    h = _user32.GetParent(root.winfo_id())
    return h or root.winfo_id()


def _fokus_alma(root: tk.Tk) -> None:
    """Pencere tıklansa bile fokusu Excel'den almasın (NOACTIVATE)."""
    h = _hwnd(root)
    ex = _user32.GetWindowLongW(h, _GWL_EXSTYLE)
    _user32.SetWindowLongW(h, _GWL_EXSTYLE, ex | _WS_EX_NOACTIVATE | _WS_EX_TOOLWINDOW)


def _hep_ustte(root: tk.Tk) -> None:
    h = _hwnd(root)
    _user32.SetWindowPos(h, ctypes.c_void_p(_HWND_TOPMOST), 0, 0, 0, 0,
                         _SWP_NOMOVE | _SWP_NOSIZE | _SWP_NOACTIVATE)


# ---- Kaldığın yeri hatırlama (devam) ----
DURUM_DOSYA = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                           "gunluk_satir_durum.json")


def _durum_oku() -> dict:
    try:
        with open(DURUM_DOSYA, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _durum_yaz(d: dict) -> None:
    try:
        with open(DURUM_DOSYA, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False)
    except Exception:
        pass


# ---------------------- Satır algılama ----------------------
def _cizgi_merkezleri(gri: np.ndarray) -> list[int]:
    """Gri görüntüdeki yatay çizgilerin y-merkezlerini döndürür."""
    H, W = gri.shape
    cizgi = (gri < KOYU_ESIK).sum(axis=1) > (CIZGI_ORANI * W)
    merkez, y = [], 0
    while y < H:
        if cizgi[y]:
            y0 = y
            while y < H and cizgi[y]:
                y += 1
            merkez.append((y0 + y - 1) // 2)
        else:
            y += 1
    return merkez


def sayfa_bantlari(gri: np.ndarray) -> list[tuple[int, int]]:
    """Yatay çizgileri bulur; ızgara düzenli olduğu için eksik çizgileri
    tespit edilen satır adımıyla (pitch) tamamlayıp satır bantları döndürür."""
    H, _ = gri.shape
    merkez = _cizgi_merkezleri(gri)
    if len(merkez) < 3:
        return []

    gaps = np.diff(merkez)
    gecerli = gaps[(gaps >= MIN_SATIR_PX) & (gaps <= MAX_SATIR_PX)]
    if len(gecerli) == 0:
        return []
    adim = float(np.median(gecerli))            # satır adımı (~43 px)

    # 1) Çapa çizgiler arasındaki eksik çizgileri doldur (87≈2x, 130≈3x adım)
    tam = [merkez[0]]
    for a, b in zip(merkez, merkez[1:]):
        g = b - a
        k = max(1, round(g / adim))
        if k > 1 and abs(g - k * adim) <= 0.45 * adim:
            for i in range(1, k):
                tam.append(int(round(a + g * i / k)))
        tam.append(b)

    # 2) Üst/alt kenardaki eksik satırları kenar payına kadar uzat (en çok 3'er)
    for _ in range(3):
        if tam[0] - adim > H * KENAR_UST:
            tam.insert(0, int(round(tam[0] - adim)))
    for _ in range(3):
        if tam[-1] + adim < H * KENAR_ALT:
            tam.append(int(round(tam[-1] + adim)))

    bantlar = [(a, b) for a, b in zip(tam, tam[1:])
               if MIN_SATIR_PX <= (b - a) <= MAX_SATIR_PX]
    return bantlar


def yedek_bantlar(H: int) -> list[tuple[int, int]]:
    """Algılama kapalıyken tablo alanını eşit parçalara böler."""
    ust = int(H * YEDEK_UST)
    alt = int(H * YEDEK_ALT)
    adim = (alt - ust) / YEDEK_SATIR
    return [(int(ust + i * adim), int(ust + (i + 1) * adim)) for i in range(YEDEK_SATIR)]


def egim_bul(gri_img: Image.Image) -> float:
    """Eğik taranan sayfanın düzeltme açısını bulur (projeksiyon profili).
    Satır çizgileri hizalanınca yatay koyuluk profili keskinleşir; bu profili
    en yükselten küçük açıyı seçer. Küçük bir kopya üstünde, önce kaba sonra
    ince taramayla (iki kademe) çalışır — hassas ama hızlı."""
    k = max(1, gri_img.width // 800)
    kucuk = gri_img.resize((max(1, gri_img.width // k), max(1, gri_img.height // k)),
                           Image.BILINEAR)

    def _skor(a: float) -> float:
        d = kucuk.rotate(a, resample=Image.BILINEAR, fillcolor=255, expand=False)
        oran = (np.asarray(d) < KOYU_ESIK).mean(axis=1)
        return float((oran ** 2).sum())   # çizgiler hizalanınca tepe yükselir

    # 1) Kaba tarama (0.5° adım) — tüm aralık
    kaba = max(DUZELT_ADIM, 0.5)
    kaba_acilar = np.arange(-DUZELT_ARALIK, DUZELT_ARALIK + 1e-9, kaba)
    en_aci = max(kaba_acilar, key=_skor)
    # 2) İnce tarama — sadece en iyi kaba açının çevresinde, DUZELT_ADIM ile
    ince_acilar = np.arange(en_aci - kaba, en_aci + kaba + 1e-9, DUZELT_ADIM)
    en_aci = max(ince_acilar, key=_skor)
    return float(en_aci)


# ---------------------- Uygulama ----------------------
class SatirSerit:
    def __init__(self, pdf_yolu: str, baslangic: int = 0):
        self.pdf_yolu = pdf_yolu
        self.sayfa_img: list[Image.Image] = []   # RGB sayfa görüntüleri
        self.satirlar: list[dict] = []           # {page, y0, y1, local, page_total}

        self._pdf_yukle()
        if not self.satirlar:
            print("HATA: Hiç satır bulunamadı. AYARLAR'ı (KOYU_ESIK/CIZGI_ORANI) "
                  "ayarla ya da YEDEK_BOL=True yap.")
            sys.exit(1)

        # (sayfa, sıra no) -> indeks  (elle atlama için)
        self.indeks = {(s["page"], s["local"]): i for i, s in enumerate(self.satirlar)}
        self.cur = max(0, min(baslangic, len(self.satirlar) - 1))
        self.komut_q: queue.Queue = queue.Queue()
        self._tus_durum = {"shift": False, "ctrl": False,
                           "shift_combo": False, "ctrl_combo": False,
                           "bastir": False}
        self.klavye = keyboard.Controller()   # Excel'e tuş göndermek için
        self.excel_ileri_aktif = EXCEL_ILERI_AKTIF
        _kayit = _durum_oku().get("_excel_ileri_aktif")
        if isinstance(_kayit, bool):
            self.excel_ileri_aktif = _kayit
        self.excel_tuslar = list(EXCEL_ILERI_TUSLAR)   # gönderilecek tuş dizisi
        _kt = _durum_oku().get("_excel_tuslar")
        if isinstance(_kt, list) and _kt:
            self.excel_tuslar = [str(x).lower() for x in _kt]
        self.kayit_modu = False               # tuş kaydı sırasında navigasyon dursun
        self._bitti_soruldu = False           # son satır sorusu bir kez sorulsun

        self._arayuz_kur()
        self._dinleyici_baslat()
        self._goster()
        self.root.after(40, self._pump)
        self.root.after(1500, self._ustte_tut_dongu)
        self.root.mainloop()

    # ---- PDF yükle + satırları çıkar ----
    def _pdf_yukle(self):
        pdf = pdfium.PdfDocument(self.pdf_yolu)
        print(f"\n📄 {os.path.basename(self.pdf_yolu)} — {len(pdf)} sayfa\n")

        for p in range(len(pdf)):
            img = pdf[p].render(scale=RENDER_OLCEK).to_pil().convert("RGB")

            aci = 0.0
            if OTO_DUZELT:                       # eğik sayfayı düzelt (atlanmayı önler)
                aci = egim_bul(img.convert("L"))
                if abs(aci) >= DUZELT_ESIK:
                    img = img.rotate(aci, resample=Image.BICUBIC,
                                     fillcolor=(255, 255, 255), expand=False)

            self.sayfa_img.append(img)
            H = img.height

            if YEDEK_BOL:
                bantlar = yedek_bantlar(H)
            else:
                gri = np.asarray(img.convert("L"))
                bantlar = sayfa_bantlari(gri)
                if ILK_SATIR_ATLA > 0:
                    bantlar = bantlar[ILK_SATIR_ATLA:]
                if not bantlar:                  # yine boşsa: sayfayı ATLAMA, eşit böl
                    print(f"  ⚠ Sayfa {p+1}: çizgi bulunamadı → eşit bölündü")
                    bantlar = yedek_bantlar(H)

            for k, (y0, y1) in enumerate(bantlar, start=1):
                self.satirlar.append({
                    "page": p, "y0": y0, "y1": y1,
                    "local": k, "page_total": len(bantlar),
                })
            ek = f"  (eğiklik {aci:+.2f}° düzeltildi)" if abs(aci) >= DUZELT_ESIK else ""
            print(f"  Sayfa {p+1}: {len(bantlar)} satır{ek}")

        pdf.close()
        print(f"\n✅ Toplam {len(self.satirlar)} satır.\n"
              f"   Sağ Shift = ileri | Sağ Ctrl = geri | ✕ = çıkış\n")

        # Şerit yüksekliğini tipik satır boyuna göre belirle
        self.ekran_w = 0  # arayüzde set edilecek
        yuk = [s["y1"] - s["y0"] for s in self.satirlar]
        self.tipik_satir = int(statistics.median(yuk)) if yuk else 60

    # ---- Tkinter arayüz ----
    def _arayuz_kur(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self.ekran_w = self.root.winfo_screenwidth()
        ekran_h = self.root.winfo_screenheight()

        sayfa_w = self.sayfa_img[0].width
        ham_h = 3 * self.tipik_satir + 2 * PAY_PX
        self.disp_w = self.ekran_w
        self.disp_h = int(min(ekran_h * MAKS_YUK_ORANI,
                              ham_h * (self.ekran_w / sayfa_w)))
        pencere_h = self.disp_h + BAR_YUK

        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.pencere_y = 0
        self.root.geometry(f"{self.ekran_w}x{pencere_h}+0+{self.pencere_y}")

        # Üst bilgi çubuğu
        bar = tk.Frame(self.root, bg="#1f1f1f", height=BAR_YUK)
        bar.pack(fill="x")
        bar.pack_propagate(False)

        self.bilgi = tk.Label(bar, text="", bg="#1f1f1f", fg="white",
                              font=("Segoe UI", 10, "bold"), anchor="w", padx=10)
        self.bilgi.pack(side="left", fill="x", expand=True)

        ipucu = tk.Label(bar, text="Sağ Shift ▶   Sağ Ctrl ◀", bg="#1f1f1f",
                         fg="#9ad", font=("Segoe UI", 9), padx=8)
        ipucu.pack(side="left")

        cik = tk.Label(bar, text="  ✕  ", bg="#b33", fg="white",
                       font=("Segoe UI", 11, "bold"), cursor="hand2")
        cik.pack(side="right", fill="y")
        cik.bind("<Button-1>", lambda e: self._kapat())

        git = tk.Label(bar, text="  ⤓ Git  ", bg="#2a7abf", fg="white",
                       font=("Segoe UI", 10, "bold"), cursor="hand2")
        git.pack(side="right", fill="y")
        git.bind("<Button-1>", self._git_ac)

        # Excel'e otomatik imleç hareketini aç/kapat
        self.excel_btn = tk.Label(bar, fg="white", font=("Segoe UI", 10, "bold"),
                                  cursor="hand2")
        self.excel_btn.pack(side="right", fill="y")
        self.excel_btn.bind("<Button-1>", self._excel_toggle)
        self._excel_btn_guncelle()

        # Excel'e gönderilen tuş dizisini elle kaydet (mini makro)
        ayarla = tk.Label(bar, text="  ⌨ Ayarla  ", bg="#6a4ca8", fg="white",
                          font=("Segoe UI", 10, "bold"), cursor="hand2")
        ayarla.pack(side="right", fill="y")
        ayarla.bind("<Button-1>", self._excel_kaydet_ac)

        # Belgeyi "(işlendi)" olarak işaretle
        bitti = tk.Label(bar, text="  ✔ Bitti  ", bg="#00796b", fg="white",
                         font=("Segoe UI", 10, "bold"), cursor="hand2")
        bitti.pack(side="right", fill="y")
        bitti.bind("<Button-1>", lambda e: self._bitti_sor())

        # Bilgi çubuğunu sürükleyerek şeridi taşı
        for w in (bar, self.bilgi, ipucu):
            w.bind("<Button-1>", self._suru_basla)
            w.bind("<B1-Motion>", self._suru)

        # Görüntü
        self.gorsel = tk.Label(self.root, bg="#f5f5f5")
        self.gorsel.pack(fill="both", expand=True)

        self.root.deiconify()
        self.root.update_idletasks()
        _fokus_alma(self.root)
        _hep_ustte(self.root)

    # ---- Sürükleme ----
    def _suru_basla(self, e):
        self._suru_y0 = e.y_root
        self._pencere_y0 = self.pencere_y

    def _suru(self, e):
        self.pencere_y = max(0, self._pencere_y0 + (e.y_root - self._suru_y0))
        self.root.geometry(f"+0+{self.pencere_y}")

    # ---- Görüntü oluşturma ----
    def _komsular(self, cur):
        p = self.satirlar[cur]["page"]
        onc = self.satirlar[cur - 1] if cur - 1 >= 0 and self.satirlar[cur - 1]["page"] == p else None
        son = self.satirlar[cur + 1] if cur + 1 < len(self.satirlar) and self.satirlar[cur + 1]["page"] == p else None
        return onc, son

    def _kompozit(self, cur) -> Image.Image:
        s = self.satirlar[cur]
        img = self.sayfa_img[s["page"]]
        W, H = img.size
        onc, son = self._komsular(cur)

        ust = (onc["y0"] if onc else s["y0"]) - PAY_PX
        alt = (son["y1"] if son else s["y1"]) + PAY_PX
        ust = max(0, ust)
        alt = min(H, alt)

        kirp = img.crop((0, ust, W, alt))
        beyaz = Image.new("RGB", kirp.size, (255, 255, 255))
        soluk = Image.blend(kirp, beyaz, SOLUK_ALFA)

        # Aktif satırı net yapıştır: NET_PAY_PX kadar üst/alta taşarak
        # net (soluk olmayan) alanı genişlet
        a0 = max(0, s["y0"] - ust - NET_PAY_PX)
        a1 = min(kirp.height, s["y1"] - ust + NET_PAY_PX)
        soluk.paste(kirp.crop((0, a0, W, a1)), (0, a0))

        # Net alanın dış kenarına çerçeve (pay sayesinde yazıyı örtmez)
        if CERCEVE_KALIN > 0:
            ImageDraw.Draw(soluk).rectangle(
                (1, a0, W - 2, a1 - 1), outline=CERCEVE_RENK, width=CERCEVE_KALIN
            )
        return soluk

    def _frame(self, cur) -> Image.Image:
        img = self._kompozit(cur)
        cw, ch = img.size
        olcek = min(self.disp_w / cw, self.disp_h / ch)
        nw, nh = max(1, int(cw * olcek)), max(1, int(ch * olcek))
        img = img.resize((nw, nh), Image.LANCZOS)
        tuval = Image.new("RGB", (self.disp_w, self.disp_h), (245, 245, 245))
        tuval.paste(img, ((self.disp_w - nw) // 2, (self.disp_h - nh) // 2))
        return tuval

    def _goster(self):
        s = self.satirlar[self.cur]
        self.tk_img = ImageTk.PhotoImage(self._frame(self.cur))
        self.gorsel.config(image=self.tk_img)
        self.bilgi.config(
            text=f"📄 {os.path.basename(self.pdf_yolu)}    "
                 f"Sayfa {s['page']+1}    •    "
                 f"Satır {s['local']}/{s['page_total']}    •    "
                 f"Toplam {self.cur+1}/{len(self.satirlar)}"
        )
        self._durum_kaydet()

    def _durum_kaydet(self):
        """Kaldığın satırı (bu PDF için) diske yaz — sonraki açılışta sorulur."""
        s = self.satirlar[self.cur]
        d = _durum_oku()
        d[self.pdf_yolu] = {
            "i": self.cur,
            "etiket": f"Sayfa {s['page']+1} • Satır {s['local']}/{s['page_total']}",
        }
        _durum_yaz(d)

    def _git_ac(self, *_):
        """Sayfa + sıra no yazıp o satıra atlama kutusu; sonra fokus Excel'e döner."""
        onceki = _user32.GetForegroundWindow()   # büyük olasılıkla Excel
        s = self.satirlar[self.cur]

        dlg = tk.Toplevel(self.root)
        dlg.title("Git")
        dlg.attributes("-topmost", True)
        dlg.resizable(False, False)
        dw, dh = 260, 160
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        gx = (sw - dw) // 2
        gy = (sh - dh) // 2          # ekranın ortası (üstteki şeridin altında kalır)
        dlg.geometry(f"{dw}x{dh}+{gx}+{gy}")

        tk.Label(dlg, text="Sayfa:", font=("Segoe UI", 11)).grid(
            row=0, column=0, padx=8, pady=(14, 4), sticky="e")
        e_sayfa = tk.Entry(dlg, width=6, font=("Segoe UI", 13))
        e_sayfa.grid(row=0, column=1, padx=8, pady=(14, 4), sticky="w")
        e_sayfa.insert(0, str(s["page"] + 1))

        tk.Label(dlg, text="Satır (sıra no):", font=("Segoe UI", 11)).grid(
            row=1, column=0, padx=8, pady=4, sticky="e")
        e_sira = tk.Entry(dlg, width=6, font=("Segoe UI", 13))
        e_sira.grid(row=1, column=1, padx=8, pady=4, sticky="w")
        e_sira.insert(0, str(s["local"]))

        uyari = tk.Label(dlg, text="", fg="#b33", font=("Segoe UI", 8))
        uyari.grid(row=2, column=0, columnspan=2)

        def kapat(*_a):
            dlg.destroy()
            if onceki:
                try:
                    _user32.SetForegroundWindow(onceki)   # fokus Excel'e geri
                except Exception:
                    pass

        def git(*_a):
            try:
                p = int(e_sayfa.get()) - 1
                sr = int(e_sira.get())
            except ValueError:
                uyari.config(text="Lütfen sayı gir.")
                return
            idx = self.indeks.get((p, sr))
            if idx is None:
                uyari.config(text="Böyle bir satır yok.")
                return
            self.cur = idx
            self._goster()
            kapat()

        tk.Button(dlg, text="Git", command=git, bg="#4CAF50", fg="white",
                  font=("Segoe UI", 10, "bold"), padx=14).grid(
            row=3, column=0, columnspan=2, pady=8)
        dlg.bind("<Return>", git)
        dlg.bind("<Escape>", kapat)
        dlg.protocol("WM_DELETE_WINDOW", kapat)

        dlg.update_idletasks()
        try:
            _user32.SetForegroundWindow(_hwnd(dlg))
        except Exception:
            pass
        e_sira.focus_force()
        e_sira.select_range(0, tk.END)

    # ---- Navigasyon (ana thread'de çağrılır) ----
    def _ileri(self) -> bool:
        if self.cur < len(self.satirlar) - 1:
            self.cur += 1
            self._goster()
            return True
        return False          # zaten son satırdaydık

    def _geri(self):
        if self.cur > 0:
            self.cur -= 1
            self._goster()

    # ---- Excel'e otomatik imleç hareketi ----
    def _excel_tus_gonder(self):
        """İleri'de Excel'e ayarlı tuşları (ör. aşağı/sol/sol) gönderir."""
        if not self.excel_ileri_aktif:
            return
        for ad in self.excel_tuslar:
            k = _TUS_HARITA.get(str(ad).lower())
            if k is None:
                continue
            self.klavye.press(k)
            self.klavye.release(k)

    def _excel_btn_guncelle(self):
        if self.excel_ileri_aktif:
            self.excel_btn.config(text="  Excel ✓  ", bg="#2e7d32")
        else:
            self.excel_btn.config(text="  Excel ✕  ", bg="#555")

    def _excel_toggle(self, *_):
        self.excel_ileri_aktif = not self.excel_ileri_aktif
        self._excel_btn_guncelle()
        d = _durum_oku()
        d["_excel_ileri_aktif"] = self.excel_ileri_aktif
        _durum_yaz(d)

    def _excel_kaydet_ac(self, *_):
        """Mini makro kaydı: yön/Enter/Tab tuşlarına sırayla bas, Kaydet'e
        tıkla; artık Sağ Shift bu diziyi Excel'e gönderir. Kayıt sırasında
        şerit gezinmesi durur (kayit_modu)."""
        onceki = _user32.GetForegroundWindow()
        self.kayit_modu = True
        kayit = list(self.excel_tuslar)          # üstünde çalışılan kopya

        dlg = tk.Toplevel(self.root)
        dlg.title("Excel kısayolunu kaydet")
        dlg.attributes("-topmost", True)
        dlg.resizable(False, False)
        dlg.configure(bg="#1f1f1f")
        dw, dh = 460, 280
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        dlg.geometry(f"{dw}x{dh}+{(sw - dw) // 2}+{(sh - dh) // 2}")

        tk.Label(dlg, text="Yön tuşlarına / Enter / Tab'a SIRAYLA bas",
                 bg="#1f1f1f", fg="white", font=("Segoe UI", 12, "bold")
                 ).pack(pady=(16, 2))
        tk.Label(dlg, text="Sağ Shift'e basınca Excel'de bu tuşlar gönderilecek.",
                 bg="#1f1f1f", fg="#9ad", font=("Segoe UI", 9)).pack()

        dizi_lbl = tk.Label(dlg, text="", bg="#111", fg="#7CFC00",
                            font=("Consolas", 20, "bold"), height=2,
                            relief="sunken", bd=2)
        dizi_lbl.pack(pady=14, padx=16, fill="x")

        def yenile():
            if kayit:
                dizi_lbl.config(text="  ".join(_TUS_SIMGE.get(a, a) for a in kayit))
            else:
                dizi_lbl.config(text="(boş — tuşlara bas)")

        def tus_bas(e):
            ad = _TK_AD.get(e.keysym)
            if ad:
                kayit.append(ad)
                yenile()
            return "break"                       # varsayılan davranışı (Tab vb.) engelle

        def geri_al(*_a):
            if kayit:
                kayit.pop()
            yenile()
            dlg.focus_set()

        def temizle(*_a):
            kayit.clear()
            yenile()
            dlg.focus_set()

        def kapat(*_a):
            self.kayit_modu = False
            dlg.destroy()
            if onceki:
                try:
                    _user32.SetForegroundWindow(onceki)
                except Exception:
                    pass

        def kaydet(*_a):
            if not kayit:                        # boş kaydetme
                return
            self.excel_tuslar = list(kayit)
            d = _durum_oku()
            d["_excel_tuslar"] = self.excel_tuslar
            if not self.excel_ileri_aktif:       # kaydeden kullanmak ister → aç
                self.excel_ileri_aktif = True
                d["_excel_ileri_aktif"] = True
                self._excel_btn_guncelle()
            _durum_yaz(d)
            kapat()

        buton = tk.Frame(dlg, bg="#1f1f1f")
        buton.pack(pady=8)
        for txt, cmd, renk in (("↶ Geri al", geri_al, "#555"),
                               ("Temizle", temizle, "#555"),
                               ("Vazgeç", kapat, "#555"),
                               ("Kaydet", kaydet, "#4CAF50")):
            tk.Button(buton, text=txt, command=cmd, bg=renk, fg="white",
                      font=("Segoe UI", 10, "bold"), width=9, takefocus=0,
                      relief="flat").pack(side="left", padx=4)

        yenile()
        dlg.bind("<KeyPress>", tus_bas)
        dlg.bind("<Escape>", kapat)              # Escape: kaydetmeden çık
        dlg.protocol("WM_DELETE_WINDOW", kapat)

        dlg.update_idletasks()
        try:
            _user32.SetForegroundWindow(_hwnd(dlg))
        except Exception:
            pass
        dlg.focus_force()

    # ---- "(işlendi)" işaretleme ----
    def _islendi_mi(self) -> bool:
        return "(işlendi)" in os.path.basename(self.pdf_yolu)

    def _bitti_sor(self, otomatik: bool = False):
        """Belgeyi bitirdin mi diye sorar; evet dersen dosya adına
        "(işlendi)" ekler. otomatik=True: son satırda Sağ Shift'e tekrar
        basılınca çağrılır ve oturum başına bir kez sorar."""
        if self._islendi_mi():
            return
        if otomatik:
            if self._bitti_soruldu:
                return
            self._bitti_soruldu = True

        self.kayit_modu = True                  # soru açıkken navigasyon dursun
        onceki = _user32.GetForegroundWindow()  # büyük olasılıkla Excel
        try:
            baslik = "Belge bitti mi?" if otomatik else "İşlendi işaretle"
            metin = (("Son satırdasın.\n\n" if otomatik else "") +
                     f"{os.path.basename(self.pdf_yolu)}\n\n"
                     f"dosya adına \"(işlendi)\" eklensin mi?")
            if messagebox.askyesno(baslik, metin, parent=self.root):
                self._bitti_isaretle()
        finally:
            self.kayit_modu = False
            if onceki:
                try:
                    _user32.SetForegroundWindow(onceki)   # fokus Excel'e geri
                except Exception:
                    pass

    def _bitti_isaretle(self):
        """Dosyayı '... (işlendi).pdf' olarak yeniden adlandırır; kaldığın-yer
        kaydını da yeni ada taşır."""
        eski = self.pdf_yolu
        kok, uz = os.path.splitext(eski)
        yeni = f"{kok} (işlendi){uz}"
        try:
            os.rename(eski, yeni)
        except OSError as e:
            messagebox.showerror(
                "Yeniden adlandırılamadı",
                f"{e}\n\nPDF başka bir programda açık olabilir; kapatıp tekrar dene.",
                parent=self.root)
            return
        d = _durum_oku()
        if eski in d:
            d[yeni] = d.pop(eski)
            _durum_yaz(d)
        self.pdf_yolu = yeni
        self._goster()                          # üst çubuktaki ad da güncellensin
        print(f"✔ İşlendi: {os.path.basename(yeni)}")

    # ---- Global tuş dinleyici (ayrı thread) ----
    def _dinleyici_baslat(self):
        self.listener = keyboard.Listener(on_press=self._bas, on_release=self._birak)
        self.listener.daemon = True
        self.listener.start()

    def _bas(self, key):
        if self.kayit_modu:        # tuş kaydı sırasında şeritte gezinme
            return
        d = self._tus_durum
        if key == ILERI_TUS:
            d["shift"] = True
            d["shift_combo"] = False
        elif key == GERI_TUS:
            d["ctrl"] = True
            d["ctrl_combo"] = False
        elif _bastir_mi(key):
            d["bastir"] = True          # basılıyken Excel'e tuş gitmez
        else:
            if d["shift"]:
                d["shift_combo"] = True
            if d["ctrl"]:
                d["ctrl_combo"] = True

    def _birak(self, key):
        if self.kayit_modu:        # tuş kaydı sırasında şeritte gezinme
            return
        d = self._tus_durum
        if key == ILERI_TUS:
            if d["shift"] and not d["shift_combo"]:
                self.komut_q.put("next_sessiz" if d["bastir"] else "next")
            d["shift"] = False
        elif key == GERI_TUS:
            if d["ctrl"] and not d["ctrl_combo"]:
                self.komut_q.put("prev")
            d["ctrl"] = False
        elif _bastir_mi(key):
            d["bastir"] = False

    # ---- Ana thread: komut kuyruğunu işle ----
    def _pump(self):
        try:
            while True:
                c = self.komut_q.get_nowait()
                if c == "next":
                    ilerledi = self._ileri()
                    self._excel_tus_gonder()
                    if not ilerledi:       # son satırda bir daha basıldı → bitti mi?
                        self._bitti_sor(otomatik=True)
                elif c == "next_sessiz":
                    if not self._ileri():  # Sol Shift basılıydı: Excel'e tuş yok
                        self._bitti_sor(otomatik=True)
                elif c == "prev":
                    self._geri()
                elif c == "quit":
                    self._kapat()
                    return
        except queue.Empty:
            pass
        self.root.after(40, self._pump)

    def _ustte_tut_dongu(self):
        _hep_ustte(self.root)
        self.root.after(1500, self._ustte_tut_dongu)

    def _kapat(self):
        try:
            self.listener.stop()
        except Exception:
            pass
        self.root.destroy()


def main():
    kok = tk.Tk()
    kok.withdraw()
    pdf_yolu = filedialog.askopenfilename(
        title="Günlük Verilen Malzeme PDF'ini seç",
        initialdir=BASLANGIC_KLASORU if os.path.isdir(BASLANGIC_KLASORU) else os.getcwd(),
        filetypes=[("PDF dosyaları", "*.pdf"), ("Tümü", "*.*")],
    )

    if not pdf_yolu:
        kok.destroy()
        print("Dosya seçilmedi, çıkılıyor.")
        return

    # Kaldığın yer kayıtlıysa sor
    baslangic = 0
    kayit = _durum_oku().get(pdf_yolu)
    if kayit:
        etiket = kayit.get("etiket") or f"{kayit.get('i', 0) + 1}. satır"
        if messagebox.askyesno(
            "Kaldığın yerden devam",
            f"Bu PDF'de en son şurada kalmıştın:\n\n{etiket}\n\nOradan devam edilsin mi?",
            parent=kok,
        ):
            baslangic = int(kayit.get("i", 0))
    kok.destroy()

    SatirSerit(pdf_yolu, baslangic)
    print("\n✅ Kapatıldı.")


if __name__ == "__main__":
    main()
