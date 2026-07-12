# -*- coding: utf-8 -*-
"""
GÜNLÜK SATIR + YAZIM YARDIMCISI (BİRLEŞİK)

Tek dosya, iki araç birlikte:
  1. Önce GÜNLÜK SATIR açılır: PDF'i seçersin, satır şeridi ekranın üstüne kurulur.
  2. Şerit kurulunca YAZIM YARDIMCISI penceresi de kendiliğinden açılır.

Birlikte çalışma kuralları:
  - Sağ Shift artık Excel'e TUŞ GÖNDERMEZ; hücreyi COM ile doğrudan taşır
    (⌨ Ayarla'daki yön dizisi toplam kaydırmaya çevrilir: down,left,left = 1 aşağı 2 sol).
    Böylece odak yazım yardımcısındayken bile ok tuşları karışmaz.
  - Sağ Shift'te kutuda yazı varsa ÖNCE o aktif hücreye yazılır (→ gerekmez),
    sonra hücre taşınır ve kutu temizlenir → yeni satıra temiz başlarsın.
  - Sol Shift + Sağ Shift: şerit ilerler ama Excel'de hücre OYNAMAZ (eskisi gibi).
  - Şeritteki "✎ Yardımcı" düğmesi yardımcı penceresini gizler/gösterir.

Orijinal dosyalar değişmedi:
  TARAMA\\gunluk_satir.py  ve  PYTON PROJELER\\excel_yazim_yardimcisi.py
Oradaki AYARLAR bloklarını değiştirirsen burası da otomatik kullanır.

Kullanım: GÜNLÜK SATIR + YARDIMCI.bat (çift tık)
"""
import os
import sys

# pythonw ile açılınca print patlamasın (gunluk_satir konsola yazar)
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w", encoding="utf-8")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w", encoding="utf-8")

MASAUSTU = os.path.join(os.path.expanduser("~"), "Desktop")
BURA = os.path.dirname(os.path.abspath(__file__))

# Gerekli iki modülü esnek ara: yanımdaki klasör -> arşiv kardeş klasörleri ->
# masaüstündeki bilinen yerler. Böylece hem iş PC düzeninde hem arşiv
# kopyasında (TCH-DEPO-ARACLARI) hem farklı bir bilgisayarda çalışır.
_MODUL_ADAYLARI = {
    "gunluk_satir.py": [
        BURA,
        os.path.normpath(os.path.join(BURA, "..", "03-gunluk-satir")),
        os.path.join(MASAUSTU, "TARAMA"),
    ],
    "excel_yazim_yardimcisi.py": [
        BURA,
        os.path.normpath(os.path.join(BURA, "..", "02-yazim-yardimcisi")),
        os.path.join(MASAUSTU, "PYTON PROJELER"),
    ],
}
_eksik = []
for _dosya, _adaylar in _MODUL_ADAYLARI.items():
    for _d in _adaylar:
        if os.path.isfile(os.path.join(_d, _dosya)):
            if _d not in sys.path:
                sys.path.insert(0, _d)
            break
    else:
        _eksik.append(f"{_dosya}  (arandı: {', '.join(_adaylar)})")

import tkinter as tk
from tkinter import ttk, messagebox

try:
    if _eksik:
        raise ImportError("\n".join(_eksik))
    import gunluk_satir as gs
    import excel_yazim_yardimcisi as eyy
except ImportError as e:
    r = tk.Tk(); r.withdraw()
    messagebox.showerror("Dosya eksik",
                         "Gerekli modül bulunamadı:\n\n" + str(e))
    sys.exit(1)


# ---------------- Yazım yardımcısı: şeride bağlı Toplevel sürümü ----------------
class YazimYardimcisi:
    """excel_yazim_yardimcisi.gui() ile aynı davranış; kendi Tk kökü yerine
    şeridin köküne bağlı Toplevel olarak yaşar, mainloop'u şeritle paylaşır."""

    def __init__(self, parent, y=140):
        self.durum = {"excel": None, "rol": None, "sayfa": None, "adres": ""}
        self.gercek = []
        liste_hata = ""
        try:
            self.listeler = eyy.listeleri_oku()
        except Exception as e:
            self.listeler = {}
            liste_hata = f"Listeler okunamadı: {e}"

        win = self.win = tk.Toplevel(parent)
        win.title("YAZIM YARDIMCISI")
        win.geometry(f"430x470+40+{y}")
        win.attributes("-topmost", True)
        win.minsize(360, 380)
        win.protocol("WM_DELETE_WINDOW", win.withdraw)  # ✕ = gizle (şerit yaşasın)

        self.baslik_var = tk.StringVar(value="Excel aranıyor…")
        ttk.Label(win, textvariable=self.baslik_var, font=("Segoe UI", 14, "bold"),
                  padding=(12, 8, 12, 2)).pack(fill="x")

        cerceve = ttk.Frame(win, padding=(12, 2))
        cerceve.pack(fill="x")
        self.metin_var = tk.StringVar()
        self.giris = ttk.Entry(cerceve, textvariable=self.metin_var, font=("Segoe UI", 15))
        self.giris.pack(side="left", fill="x", expand=True)
        self.giris.focus_set()
        ttk.Button(cerceve, text="⟳", width=3, command=self.listeleri_yenile).pack(
            side="left", padx=(6, 0))
        ttk.Button(cerceve, text="➕", width=3, command=self.yeni_personel).pack(
            side="left", padx=(4, 0))

        self.kutu = tk.Listbox(win, font=("Segoe UI", 13), activestyle="dotbox",
                               selectbackground="#1F4E78", selectforeground="white",
                               height=9)
        self.kutu.pack(fill="both", expand=True, padx=12, pady=(4, 2))

        ipucu = ("Seçmezsen yazdığın AYNEN yazılır — ↓ seç | Ctrl+D üsttekini kopyala | "
                 "Ctrl+N/➕ YENİ personel\n"
                 "→ yaz + sağa   ← yaz + sola   Enter yaz + sağa   Esc temizle\n"
                 "Sağ Shift: kutudakini/seçileni YAZAR + hücreyi taşır + PDF satırı ilerler")
        ttk.Label(win, text=ipucu, font=("Segoe UI", 10), foreground="#555",
                  padding=(12, 2)).pack(fill="x")

        self.durum_var = tk.StringVar(value=liste_hata)
        ttk.Label(win, textvariable=self.durum_var, relief="sunken", anchor="w",
                  padding=(10, 4), font=("Segoe UI", 10)).pack(fill="x", side="bottom")

        g = self.giris
        g.bind("<Right>", lambda e: self.yaz(1, 0))
        g.bind("<Left>", lambda e: self.yaz(-1, 0))
        g.bind("<Return>", lambda e: self.yaz(1, 0))
        g.bind("<Down>", lambda e: self.kaydir(1))
        g.bind("<Up>", lambda e: self.kaydir(-1))
        g.bind("<Escape>", lambda e: (self.metin_var.set(""), self.oner(), "break")[-1])
        g.bind("<Control-d>", self.ctrl_d)
        g.bind("<Control-D>", self.ctrl_d)
        win.bind("<Control-d>", self.ctrl_d)
        win.bind("<Control-D>", self.ctrl_d)
        g.bind("<Control-n>", self.yeni_personel)
        g.bind("<Control-N>", self.yeni_personel)
        win.bind("<Control-n>", self.yeni_personel)
        win.bind("<Control-N>", self.yeni_personel)
        g.bind("<KeyRelease>", lambda e: None if e.keysym in
               ("Right", "Left", "Up", "Down", "Return", "Escape") else self.oner())
        self.kutu.bind("<Double-Button-1>", lambda e: self.yaz(1, 0))

        if eyy.excel_bul() is None:
            try:
                os.startfile(eyy.DOSYA_TCH)
                self.durum_var.set("Excel açılıyor, dosya yüklenince bağlanacak…")
            except OSError:
                self.durum_var.set("TCH dosyası bulunamadı — Excel'i elle aç.")
        self.dongu()

    # --- dışarıya (şerit entegrasyonu için) ---
    def excel_ver(self):
        return self.durum["excel"]

    def bekleyeni_yaz(self):
        """Sağ Shift akışı: kutuda yazı ya da listeden seçim varsa aktif
        hücreye YERİNDE yazar (hücre taşımaz — taşımayı makro yapar)."""
        secim = self.kutu.curselection() if self.durum["rol"] in eyy.LISTE_ROLLER else ()
        if not self.metin_var.get().strip() and not secim:
            return
        excel = self.durum["excel"]
        if excel is None:
            return
        metin = self.secili_metin()
        try:
            eyy.hucreye_yaz(excel, metin, 0, 0)
            self.durum_var.set(f"{self.durum['adres']} ← {metin[:40]}")
            self.metin_var.set("")
        except Exception:
            self.durum_var.set("Yazılamadı — Excel meşgul (hücre düzenleme modunda olabilir).")

    def satir_gecisi(self):
        """Sağ Shift ile PDF satırı ilerleyince: kutu temizlenir, mod tazelenir."""
        self.metin_var.set("")
        self.izle()

    # --- iç işleyiş (excel_yazim_yardimcisi.gui ile birebir mantık) ---
    def listeleri_yenile(self):
        try:
            self.listeler.update(eyy.listeleri_oku())
            self.oner()
            self.durum_var.set("Listeler yenilendi.")
        except Exception as e:
            self.durum_var.set(f"Liste yenilenemedi: {e}")

    def oner(self, *_):
        kutu = self.kutu
        kutu.delete(0, "end")
        self.gercek = []   # listbox sırası -> hücreye yazılacak GERÇEK değer
        rol = self.durum["rol"]
        if rol not in eyy.LISTE_ROLLER:
            return
        adaylar = self.listeler.get(rol, [])
        aranan = eyy.fold(self.metin_var.get())
        kelimeler = aranan.split()
        stoklar = self.listeler.get("MALZEME_STOK", {}) if rol == "MALZEME" else {}

        def stok_sirasi(a):
            try:
                return -float(stoklar.get(eyy.fold(a)))
            except (TypeError, ValueError):
                return float("inf")   # stoğu bilinmeyen en alta

        sayilar = self.listeler.get("PERSONEL_SAYI", {}) if rol == "PERSONEL" else {}
        if kelimeler:
            eslesen = [a for a in adaylar if all(k in eyy.fold(a) for k in kelimeler)]
            if rol == "MALZEME":   # stoğu çok olan üstte
                eslesen.sort(key=lambda a: (stok_sirasi(a), eyy.fold(a)))
            elif rol == "PERSONEL":   # hareketi çok olan üstte
                eslesen.sort(key=lambda a: (-sayilar.get(eyy.fold(a), 0), eyy.fold(a)))
            else:
                eslesen.sort(key=lambda a: (not eyy.fold(a).startswith(kelimeler[0]), eyy.fold(a)))
        else:
            eslesen = adaylar if len(adaylar) <= 12 else []
        for a in eslesen[:60]:
            self.gercek.append(a)
            k = eyy.fold(a)
            kutu.insert("end", f"{a}    ({eyy.stok_metni(stoklar[k])})" if k in stoklar else a)
        # bilerek ↑↓ ile seçilmedikçe öneri seçili gelmez: ham metin aynen yazılır
        if kutu.size():
            kutu.see(0)

    def secili_metin(self):
        sec = self.kutu.curselection()
        if self.durum["rol"] in eyy.LISTE_ROLLER and sec:
            return self.gercek[sec[0]]
        return self.metin_var.get()

    def yaz(self, dx, dy):
        excel = self.durum["excel"]
        if excel is None:
            self.durum_var.set("Excel'e bağlı değil.")
            return "break"
        # kutu boş olsa bile listeden ↓ ile seçim yapıldıysa onu yaz
        secim = self.kutu.curselection() if self.durum["rol"] in eyy.LISTE_ROLLER else ()
        metin = self.secili_metin() if (self.metin_var.get().strip() or secim) else ""
        try:
            eyy.hucreye_yaz(excel, metin, dx, dy)
            if metin:
                self.durum_var.set(f"{self.durum['adres']} ← {metin[:40]}")
            self.metin_var.set("")
            self.izle()
        except Exception:
            self.durum_var.set("Yazılamadı — Excel meşgul (hücre düzenleme modunda olabilir).")
        return "break"

    def yeni_personel(self, _e=None):
        """Ctrl+N: kutudaki adı ÖNCE PERSONEL LİSTESİ'ne ekler, SONRA hücreye
        yazıp sağa geçer. İlk basış onay ister; ad listedeyse sadece yazar."""
        excel = self.durum["excel"]
        if excel is None:
            self.durum_var.set("Excel'e bağlı değil.")
            return "break"
        if self.durum["rol"] != "PERSONEL":
            self.durum_var.set("Yeni personel ekleme PERSONEL sütunundayken kullanılır.")
            return "break"
        ad = eyy.tr_upper(self.metin_var.get())
        if len(ad.split()) < 2:
            self.durum_var.set("Önce AD SOYAD yaz (en az iki kelime), sonra Ctrl+N.")
            return "break"
        mevcut = next((a for a in self.listeler.get("PERSONEL", [])
                       if eyy.fold(a) == eyy.fold(ad)), None)
        if mevcut:
            try:
                eyy.hucreye_yaz(excel, mevcut, 1, 0)
                self.metin_var.set("")
                self.izle()
                self.durum_var.set(f"{mevcut} zaten listedeydi — hücreye yazıldı.")
            except Exception:
                self.durum_var.set("Yazılamadı — Excel meşgul.")
            return "break"
        if self.durum.get("yeni_onay") != eyy.fold(ad):
            self.durum["yeni_onay"] = eyy.fold(ad)
            self.durum_var.set(f"YENİ PERSONEL 『{ad}』 listeye eklenecek — onay için tekrar Ctrl+N.")
            return "break"
        self.durum["yeni_onay"] = None
        try:
            satir, hata = eyy.personel_listeye_ekle(excel, ad)
        except Exception as e:
            satir, hata = None, str(e)
        if hata:
            self.durum_var.set("Listeye eklenemedi: " + hata)
            return "break"
        self.listeler["PERSONEL"].append(ad)
        self.listeler["PERSONEL"].sort(key=eyy.fold)
        try:
            eyy.hucreye_yaz(excel, ad, 1, 0)
            self.metin_var.set("")
            self.izle()
            self.durum_var.set(f"➕ {ad} → PERSONEL LİSTESİ satır {satir} + hücreye yazıldı.")
        except Exception:
            self.durum_var.set(f"Listeye eklendi (satır {satir}) ama hücreye yazılamadı — tekrar dene.")
        return "break"

    def ctrl_d(self, _e=None):
        """Excel'in Ctrl+D'si: üstteki satırdakini bulunduğun hücreye kopyalar."""
        excel = self.durum["excel"]
        if excel is None:
            self.durum_var.set("Excel'e bağlı değil.")
            return "break"
        try:
            if eyy.ustu_kopyala(excel):
                self.durum_var.set(f"{self.durum['adres']} ← üstteki satırdan kopyalandı (Ctrl+D)")
            else:
                self.durum_var.set("En üst satırdasın, üstünde kopyalanacak hücre yok.")
            self.izle()
        except Exception:
            self.durum_var.set("Kopyalanamadı — Excel meşgul (hücre düzenleme modunda olabilir).")
        return "break"

    def kaydir(self, yon):
        kutu = self.kutu
        if self.durum["rol"] in eyy.LISTE_ROLLER and kutu.size():
            sec = kutu.curselection()
            if not sec:
                if yon > 0:  # ilk ↓ ile seçim başlar
                    kutu.selection_set(0)
                    kutu.see(0)
            else:
                i = sec[0] + yon
                kutu.selection_clear(0, "end")
                if i >= 0:  # en üstteyken ↑ = seçimi bırak, ham metne dön
                    i = min(kutu.size() - 1, i)
                    kutu.selection_set(i)
                    kutu.see(i)
        else:
            self.yaz(0, yon)
        return "break"

    def izle(self):
        excel = self.durum["excel"]
        if excel is None:
            excel = eyy.excel_bul()
            if excel is None:
                self.baslik_var.set("Excel açık değil")
                return
            self.durum["excel"] = excel
        try:
            bilgi = eyy.hucre_bilgi(excel)
            if bilgi is None:
                self.baslik_var.set("Excel'de dosya açık değil")
                return
            sayfa, adres, satir, kolon, rol = bilgi
            self.durum["adres"] = adres
            degisti = (rol != self.durum["rol"]) or (sayfa != self.durum["sayfa"])
            self.durum["rol"], self.durum["sayfa"] = rol, sayfa
            self.baslik_var.set(f"{adres}  —  {rol}" if rol != "SERBEST"
                                else f"{adres}  —  serbest yazım")
            if degisti:
                self.oner()
        except Exception:
            self.durum["excel"] = eyy.excel_bul()
            self.baslik_var.set("Excel meşgul…")

    def dongu(self):
        self.izle()
        self._sayac = getattr(self, "_sayac", 0) + 1
        if self._sayac % 20 == 0 and self.durum["excel"] is not None:
            # ~10 sn'de bir liman stoğunu açık Excel'den CANLI tazele
            stok = eyy.depo_stok_com(self.durum["excel"])
            if stok:
                self.listeler["MALZEME_STOK"] = stok
                if self.durum["rol"] == "MALZEME" and not self.kutu.curselection():
                    self.oner()
        self.win.after(500, self.dongu)


# ---------------- Şerit: Sağ Shift'te tuş yerine COM ile hücre taşıyan sürüm ----------------
class BirlesikSerit(gs.SatirSerit):
    yardimci = None

    def _arayuz_kur(self):
        super()._arayuz_kur()
        # yardımcıyı şeridin hemen altına yerleştir
        self.yardimci = YazimYardimcisi(self.root, y=self.disp_h + gs.BAR_YUK + 16)
        # şeridin üst çubuğuna gizle/göster düğmesi
        bar = self.root.winfo_children()[0]
        btn = tk.Label(bar, text="  ✎ Yardımcı  ", bg="#8a6d3b", fg="white",
                       font=("Segoe UI", 10, "bold"), cursor="hand2")
        btn.pack(side="right", fill="y")
        btn.bind("<Button-1>", lambda e: self._yardimci_toggle())

    def _yardimci_toggle(self):
        w = self.yardimci.win
        if w.state() == "withdrawn":
            w.deiconify()
        else:
            w.withdraw()

    def _excel_tus_gonder(self):
        """Sağ Shift'te Excel'e SENTETİK TUŞ GÖNDERİLMEZ (odak yardımcıdayken
        ok tuşları karışırdı). Kayıtlı yön dizisi toplam kaydırmaya çevrilir ve
        aktif hücre COM ile taşınır — odak nerede olursa olsun doğru çalışır."""
        if not self.excel_ileri_aktif:
            return
        # önce kutuda bekleyen yazıyı aktif hücreye yaz (→ tuşuna gerek kalmasın)
        if self.yardimci:
            self.yardimci.bekleyeni_yaz()
        dr = dc = 0
        for ad in self.excel_tuslar:
            a = str(ad).lower()
            if a in ("down", "enter"):
                dr += 1
            elif a == "up":
                dr -= 1
            elif a in ("right", "tab"):
                dc += 1
            elif a == "left":
                dc -= 1
        excel = self.yardimci.excel_ver() if self.yardimci else None
        if excel is None:
            excel = eyy.excel_bul()
        if excel is not None and (dr or dc):
            try:
                h = excel.ActiveCell
                h.Worksheet.Cells(max(1, h.Row + dr), max(1, h.Column + dc)).Select()
            except Exception:
                pass
        if self.yardimci:
            self.yardimci.satir_gecisi()


def main():
    gs.SatirSerit = BirlesikSerit  # gs.main() bizim sürümü kullansın
    gs.main()


if __name__ == "__main__":
    main()
