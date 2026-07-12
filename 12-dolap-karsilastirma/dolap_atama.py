# -*- coding: utf-8 -*-
"""
DOLAP ATAMA — ayrılmış personelin dolabını güncel kullanıcısına devretme.

ZIMMET LİSTESİ.xlsm DOLAP sayfasında AÇIKLAMA'sında "AYRILMIŞ" yazan (kırmızı)
satırları SIRAYLA gösterir. Ara kutusuna yeni kullanıcının adını yaz (doğru
personel listesinden canlı süzülür, ilk sonuç seçili gelir), Enter = AÇIK
Excel'de o satırın adı değişir, kırmızı boya ve AYRILMIŞ notu temizlenir.

  Enter / ✔ YAZ : seçili adı yaz (zaten dolabı olana yazarken 2. Enter ister)
  Sağ Shift / Atla : bu dolabı geç (kullananı bilinmiyor — kırmızı kalır)
  ↶ Son yazımı geri al : yanlış yazdıysan eski adı ve kırmızıyı geri getirir

Dosya KAYDEDİLMEZ; bitiş ekranındaki düğme kaydedip DOLAP VERİLECEK
PERSONELLER listesini de günceller. Başlamadan yedek alınır.
Araç kapatılıp açılırsa kaldığı yerden sürer (değişen satırlar artık
"AYRILMIŞ" olmadığı için kuyruğa girmez).
"""
import os
import sys
import time
import shutil
import datetime

BURA = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BURA)
import dolap_eslestir as de

fold = de.fold


def isaretli_satirlar():
    """DOLAP'ta AÇIKLAMA'sında AYRILMIŞ geçen satırlar + mevcut sahipler."""
    wb, k = de._kopyadan_ac(de.ZIMMET)
    ws = wb["DOLAP"]
    kuyruk, sahipler = [], {}
    for i, r in enumerate(ws.iter_rows(min_row=2, max_col=5, values_only=True), start=2):
        ad = str(r[1]).strip() if r[1] else ""
        if r[3] and "AYRILMIS" in fold(r[3]):
            kuyruk.append({"satir": i, "no": r[0], "ad": ad, "aciklama": str(r[3])})
        elif ad:
            sahipler[fold(ad)] = (i, r[0])
    wb.close(); os.remove(k)
    return kuyruk, sahipler


def excel_baglan():
    import win32com.client
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
        excel.Workbooks.Count
    except Exception:
        excel = win32com.client.DispatchEx("Excel.Application")
        time.sleep(2)
        try:
            excel.Visible = True
        except Exception:
            pass
    hedef = fold(os.path.basename(de.ZIMMET))
    for w in excel.Workbooks:
        if fold(w.Name) == hedef:
            return excel, w
    return excel, excel.Workbooks.Open(de.ZIMMET, 0)


XL_YOK = -4142       # xlPatternNone
XL_OTOMATIK = -4105  # xlColorIndexAutomatic


def gui():
    import tkinter as tk
    from tkinter import ttk, messagebox

    dogru, tc, saha, ayrilmis, ayr_tc, dolap = de.veri_yukle()
    adlar = sorted(dogru.values(), key=fold)
    kuyruk, sahipler = isaretli_satirlar()

    if not kuyruk:
        r = tk.Tk(); r.withdraw()
        messagebox.showinfo("Bitti", "AYRILMIŞ işaretli satır kalmadı.")
        return

    try:
        excel, wb = excel_baglan()
        ws = wb.Worksheets("DOLAP")
    except Exception as e:
        r = tk.Tk(); r.withdraw()
        messagebox.showerror("Excel", f"ZIMMET LİSTESİ'ne bağlanılamadı:\n{e}")
        return

    ts = datetime.datetime.now().strftime("%Y-%m-%d_%H%M")
    yedek = os.path.join(BURA, f"ZIMMET LİSTESİ_YEDEK_ATAMA_{ts}.xlsm")
    shutil.copy2(de.ZIMMET, yedek)

    root = tk.Tk()
    root.title("DOLAP ATAMA — AYRILMIŞ DOLAPLARI DEVRET")
    root.geometry("780x640")
    stil = ttk.Style(root)
    try:
        stil.theme_use("vista")
    except Exception:
        pass

    durum = {"i": 0, "atlanan": [], "yazilan": 0, "son": None, "onay": None}
    atananlar = set()   # bu oturumda atananlar (mükerrer uyarısı için)

    ust_var = tk.StringVar()
    ttk.Label(root, textvariable=ust_var, font=("Segoe UI", 11),
              padding=(14, 8, 14, 0)).pack(fill="x")
    ad_var = tk.StringVar()
    ttk.Label(root, textvariable=ad_var, font=("Segoe UI", 20, "bold"),
              foreground="#B33", padding=(14, 2)).pack(fill="x")
    bilgi_var = tk.StringVar()
    ttk.Label(root, textvariable=bilgi_var, font=("Segoe UI", 10),
              foreground="#555", padding=(14, 0)).pack(fill="x")

    ara_cerceve = ttk.Frame(root, padding=(14, 8, 14, 2))
    ara_cerceve.pack(fill="x")
    ttk.Label(ara_cerceve, text="Yeni kullanıcı:", font=("Segoe UI", 12)).pack(side="left")
    ara_var = tk.StringVar()
    ara = ttk.Entry(ara_cerceve, textvariable=ara_var, font=("Segoe UI", 14))
    ara.pack(side="left", fill="x", expand=True, padx=(8, 0))

    kutu = tk.Listbox(root, font=("Segoe UI", 13), height=10,
                      selectbackground="#1F4E78", selectforeground="white")
    kutu.pack(fill="both", expand=True, padx=14, pady=(4, 2))
    gercek = []

    ipucu = ("Yaz → ilk sonuç seçili gelir → Enter = Excel'de değiştir (anında)\n"
             "Sağ Shift veya Atla = bu dolabı geç   |   zaten dolabı olana 2. Enter gerekir")
    ttk.Label(root, text=ipucu, font=("Segoe UI", 10), foreground="#555",
              padding=(14, 2)).pack(fill="x")

    dugme = ttk.Frame(root, padding=(14, 4, 14, 10))
    dugme.pack(fill="x")
    yaz_btn = tk.Button(dugme, text="✔ YAZ (Enter)", bg="#2e7d32", fg="white",
                        font=("Segoe UI", 12, "bold"), padx=14)
    yaz_btn.pack(side="left")
    atla_btn = tk.Button(dugme, text="⏭ Atla (Sağ Shift)", bg="#666", fg="white",
                         font=("Segoe UI", 12, "bold"), padx=14)
    atla_btn.pack(side="left", padx=10)
    geri_btn = tk.Button(dugme, text="↶ Son yazımı geri al", font=("Segoe UI", 11), padx=10)
    geri_btn.pack(side="left")

    alt_var = tk.StringVar()
    ttk.Label(root, textvariable=alt_var, relief="sunken", anchor="w",
              padding=(10, 4), font=("Segoe UI", 10)).pack(fill="x", side="bottom")

    def oneri_doldur():
        kutu.delete(0, "end")
        gercek.clear()
        aranan = fold(ara_var.get())
        if len(aranan) < 2:
            return
        kelimeler = aranan.split()
        for a in adlar:
            f = fold(a)
            if all(k in f for k in kelimeler):
                gercek.append(f)
                ek = "   ← DOLABI VAR!" if (f in sahipler or f in atananlar) else ""
                kutu.insert("end", a + ek)
                if kutu.size() >= 40:
                    break
        if kutu.size():
            kutu.selection_set(0)

    def goster():
        durum["onay"] = None
        if durum["i"] >= len(kuyruk):
            bitir()
            return
        k = kuyruk[durum["i"]]
        ust_var.set(f"Dolap {durum['i'] + 1} / {len(kuyruk)}   "
                    f"(yazılan: {durum['yazilan']}, atlanan: {len(durum['atlanan'])})")
        ad_var.set(k["ad"])
        bilgi_var.set(f"Dolap No: {k['no']}  •  satır {k['satir']}  •  {k['aciklama']}")
        ara_var.set("")
        kutu.delete(0, "end")
        gercek.clear()
        ara.focus_set()

    def hucre_yaz(k, yeni_ad):
        hucre = ws.Cells(k["satir"], 2)
        if hucre.Value is None or fold(hucre.Value) != fold(k["ad"]):
            return "Satır beklenen ismi içermiyor — Excel'de kontrol et."
        hucre.Value = yeni_ad
        ws.Cells(k["satir"], 4).Value = None
        alan = ws.Range(f"A{k['satir']}:E{k['satir']}")
        alan.Interior.Pattern = XL_YOK
        alan.Font.ColorIndex = XL_OTOMATIK
        alan.Font.Bold = False
        return None

    def yaz(*_):
        sec = kutu.curselection()
        if not sec:
            alt_var.set("Önce ad yaz ve listeden seç.")
            return "break"
        hedef = gercek[sec[0]]
        yeni_ad = dogru[hedef]
        k = kuyruk[durum["i"]]
        if (hedef in sahipler or hedef in atananlar) and durum["onay"] != hedef:
            nerede = f" (satır {sahipler[hedef][0]}, dolap {sahipler[hedef][1]})" if hedef in sahipler else " (bu oturumda atandı)"
            alt_var.set(f"DİKKAT: {yeni_ad} zaten dolap sahibi{nerede} — eminsen tekrar Enter'a bas.")
            durum["onay"] = hedef
            return "break"
        try:
            hata = hucre_yaz(k, yeni_ad)
        except Exception as e:
            hata = str(e)
        if hata:
            alt_var.set("Yazılamadı: " + hata)
            return "break"
        atananlar.add(hedef)
        durum["son"] = (durum["i"], k, yeni_ad, hedef)
        durum["yazilan"] += 1
        alt_var.set(f"Dolap {k['no']}: {k['ad']}  →  {yeni_ad}")
        durum["i"] += 1
        goster()
        return "break"

    def atla(*_):
        k = kuyruk[durum["i"]]
        durum["atlanan"].append(f"Dolap {k['no']} ({k['ad']})")
        durum["i"] += 1
        goster()
        return "break"

    def geri_al(*_):
        if not durum["son"]:
            return
        i, k, yeni_ad, hedef = durum["son"]
        try:
            hucre = ws.Cells(k["satir"], 2)
            if hucre.Value is None or fold(hucre.Value) != fold(yeni_ad):
                alt_var.set("Geri alınamadı: satır beklenen ismi içermiyor.")
                return
            hucre.Value = k["ad"]
            ws.Cells(k["satir"], 4).Value = k["aciklama"]
            alan = ws.Range(f"A{k['satir']}:E{k['satir']}")
            alan.Interior.Color = 255
            alan.Font.Color = 16777215
            alan.Font.Bold = True
        except Exception as e:
            alt_var.set("Geri alınamadı: " + str(e))
            return
        atananlar.discard(hedef)
        durum["yazilan"] -= 1
        durum["son"] = None
        durum["i"] = i
        goster()
        alt_var.set(f"Geri alındı: dolap {k['no']} yeniden {k['ad']} (kırmızı).")

    def bitir():
        for w in root.winfo_children():
            w.destroy()
        ttk.Label(root, text="Atama tamamlandı ✔", font=("Segoe UI", 18, "bold"),
                  padding=(16, 16)).pack()
        ozet = (f"Yazılan: {durum['yazilan']}   |   Atlanan: {len(durum['atlanan'])}\n"
                "Excel'de değişiklikler duruyor ama dosya HENÜZ KAYDEDİLMEDİ.")
        if durum["atlanan"]:
            ozet += "\n\nAtlananlar (kırmızı kaldı):\n" + "\n".join(durum["atlanan"][:15])
        ttk.Label(root, text=ozet, font=("Segoe UI", 11), padding=(16, 4),
                  justify="left").pack()
        sonuc_var = tk.StringVar()
        ttk.Label(root, textvariable=sonuc_var, font=("Segoe UI", 11),
                  foreground="#2e7d32", padding=(16, 4)).pack()

        def kaydet_guncelle():
            try:
                wb.Save()
            except Exception as e:
                sonuc_var.set("Kaydedilemedi: " + str(e))
                return
            try:
                import dolap_verilecek_liste as dvl
                toplam, almis, verilecek, kararsiz = dvl.uret(ac=True)
                sonuc_var.set(f"Kaydedildi ✓  DOLAP VERİLECEK güncellendi: {verilecek} kişi "
                              f"(dolap almış: {almis}/{toplam})")
            except Exception as e:
                sonuc_var.set("Kaydedildi ✓ ama liste güncellenemedi: " + str(e))

        tk.Button(root, text="💾 Kaydet + DOLAP VERİLECEK listesini güncelle",
                  command=kaydet_guncelle, bg="#1F4E78", fg="white",
                  font=("Segoe UI", 13, "bold"), padx=16, pady=6).pack(pady=10)
        ttk.Label(root, text=f"Yedek: {os.path.basename(yedek)}",
                  font=("Segoe UI", 9), foreground="#777", padding=(10, 6)).pack()

    # Sağ Shift: temiz basışta atla (başka tuşla kombinasyonda değil)
    shift = {"basili": False, "combo": False}

    def tus_bas(e):
        if e.keysym == "Shift_R":
            shift["basili"] = True
            shift["combo"] = False
        elif shift["basili"]:
            shift["combo"] = True

    def tus_birak(e):
        if e.keysym == "Shift_R":
            if shift["basili"] and not shift["combo"]:
                atla()
            shift["basili"] = False

    root.bind_all("<KeyPress>", tus_bas, add="+")
    root.bind_all("<KeyRelease>", tus_birak, add="+")

    yaz_btn.config(command=yaz)
    atla_btn.config(command=atla)
    geri_btn.config(command=geri_al)
    ara.bind("<Return>", yaz)
    ara.bind("<Down>", lambda e: (kutu.size() and (kutu.selection_clear(0, "end"),
             kutu.selection_set(min((kutu.curselection() or (0,))[0] + 1, kutu.size() - 1)),
             kutu.see(kutu.curselection()[0])), "break")[-1])
    ara.bind("<Up>", lambda e: (kutu.size() and (lambda i=max((kutu.curselection() or (0,))[0] - 1, 0):
             (kutu.selection_clear(0, "end"), kutu.selection_set(i), kutu.see(i)))(), "break")[-1])
    ara.bind("<KeyRelease>", lambda e: None if e.keysym in
             ("Return", "Up", "Down", "Escape", "Shift_R") else (oneri_doldur(), durum.update(onay=None)))
    ara.bind("<Escape>", lambda e: ara_var.set(""))
    kutu.bind("<Double-Button-1>", yaz)

    goster()
    root.mainloop()


if __name__ == "__main__":
    gui()
