"""
TCH Zimmet Belgesi Yazdırma Scripti
====================================
Kullanım:
    Bu scripti OLUSTURULAN klasörünün bulunduğu renk klasörüne kopyala, terminalde:

        python zimmet_yazdir.py

    Gereksinim: pip install pywin32

İşleyiş:
    1. YAZDIRILDI klasöründe zaten olan belgeleri OLUSTURULAN'dan siler
       (tekrardan yazdırmamak için — script interrupt olursa kaldığı yerden devam).
    2. Her belgeyi sırayla varsayılan yazıcıya gönderir.
    3. Yazıcı kuyruğundan belgenin tamamlandığını teyit eder (sadece Word bitirince değil,
       spooler'dan da düştüğünde).
    4. Başarılıysa belgeyi YAZDIRILDI klasörüne taşır.
    5. Herhangi bir hata olursa DURUR — kalan belgelere dokunmaz.
    6. yazdirma_log.txt dosyasına zaman damgalı kayıt tutar.
"""

import sys
import time
import shutil
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).parent.resolve()
OLUSTURULAN_DIR = BASE_DIR / "OLUSTURULAN"
YAZDIRILDI_DIR = BASE_DIR / "YAZDIRILDI"
LOG_FILE = BASE_DIR / "yazdirma_log.txt"

TIMEOUT_PER_DOC = 120    # saniye — yazdırma bu sürede bitmezse DURUR
POLL_INTERVAL = 0.5      # saniye — yazıcı kuyruğu kontrol sıklığı
EMPTY_QUEUE_THRESHOLD = 4  # kuyruk bu kadar ardışık boş görünürse başarılı say
PAUSE_EVERY = 0          # bu kadar belgede bir duraklayıp Enter bekle (0 = kapalı)
                         # Önerilen: 0 (kapalı). Bunun yerine başlangıçta
                         # "kaç tane?" sorulur, batch bitince script çıkar.

# Türkçe alfabetik sıralama (Ç,Ğ,İ,Ö,Ş,Ü Z'den sonra değil, doğru yerlerinde)
TR_ALPHABET = "ABCÇDEFGĞHIİJKLMNOÖPQRSŞTUÜVWXYZ"
TR_ORDER = {c: i for i, c in enumerate(TR_ALPHABET, start=100)}


def tr_upper(s):
    """Türkçe-doğru büyük harf (i → İ, ı → I)."""
    return s.translate(str.maketrans('iı', 'İI')).upper()


def tr_sort_key(path):
    """Path veya string için Türkçe alfabetik sıralama anahtarı.
    Türkçe harfler 100+ pozisyondan başlar; boşluk/rakam (ASCII < 100)
    bu yüzden tüm harflerden önce sıralanır."""
    s = path.name if hasattr(path, 'name') else str(path)
    return tuple(TR_ORDER.get(c, ord(c)) for c in tr_upper(s))

# Windows spooler job status bit flags
# (https://docs.microsoft.com/en-us/windows/win32/printdocs/job-info-1)
JOB_STATUS_ERROR = 0x00000002
JOB_STATUS_OFFLINE = 0x00000020
JOB_STATUS_PAPEROUT = 0x00000040
JOB_STATUS_BLOCKED_DEVQ = 0x00000200
JOB_STATUS_USER_INTERVENTION = 0x00000400


def log(msg):
    """Log dosyasına ve konsola yaz."""
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def deduplicate():
    """YAZDIRILDI'da zaten var olan dosyaları OLUSTURULAN'dan sil."""
    if not YAZDIRILDI_DIR.exists() or not OLUSTURULAN_DIR.exists():
        return 0
    printed = {f.name for f in YAZDIRILDI_DIR.glob("*.docx")}
    removed = 0
    for f in OLUSTURULAN_DIR.glob("*.docx"):
        if f.name in printed:
            f.unlink()
            removed += 1
    return removed


def check_job_errors(job):
    """Job'ta bilinen bir hata durumu varsa Türkçe sebep döndür, yoksa None."""
    status = job.get('Status', 0)
    if status & JOB_STATUS_ERROR:
        return "yazıcı hatası"
    if status & JOB_STATUS_OFFLINE:
        return "yazıcı çevrimdışı"
    if status & JOB_STATUS_PAPEROUT:
        return "kağıt bitti"
    if status & JOB_STATUS_BLOCKED_DEVQ:
        return "yazıcı kuyruğu bloklu"
    if status & JOB_STATUS_USER_INTERVENTION:
        return "kullanıcı müdahalesi gerekiyor"
    return None


def wait_for_print_completion(printer_name, doc_filename, timeout):
    """Spooler kuyruğunda doc_filename içeren job gözükene ve sonra
    kuyruktan düşene kadar bekle. (success, reason_if_failed) döner."""
    import win32print

    deadline = time.time() + timeout
    seen_in_queue = False
    empty_checks = 0

    while time.time() < deadline:
        h = win32print.OpenPrinter(printer_name)
        try:
            jobs = win32print.EnumJobs(h, 0, 999, 1)
        finally:
            win32print.ClosePrinter(h)

        our_jobs = [
            j for j in jobs
            if doc_filename in (j.get('pDocument') or '')
        ]

        if our_jobs:
            seen_in_queue = True
            empty_checks = 0
            for job in our_jobs:
                err = check_job_errors(job)
                if err:
                    return False, err
        else:
            if seen_in_queue:
                # Bizim job'u gördük, şimdi kuyrukta yok — bitti
                return True, None
            empty_checks += 1
            if empty_checks >= EMPTY_QUEUE_THRESHOLD:
                # Hiç görmedik ama kuyruk uzun süredir boş
                # Muhtemelen çok hızlı tamamlandı, başarı say
                return True, None

        time.sleep(POLL_INTERVAL)

    return False, f"zaman aşımı ({timeout}s)"


def print_document(word_app, doc_path, printer_name, timeout):
    """Tek belgeyi Word üzerinden yazdır ve spooler'dan düşmesini bekle.
    (success: bool, reason_if_failed: str|None) döner."""
    doc = None
    try:
        doc = word_app.Documents.Open(
            str(doc_path),
            ReadOnly=True,
            ConfirmConversions=False,
            AddToRecentFiles=False,
        )
        # Background=False: Word spooler'a yolluyana kadar bu çağrı döner
        doc.PrintOut(Background=False)
    except Exception as e:
        return False, f"Word hatası: {e}"
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass

    return wait_for_print_completion(printer_name, doc_path.name, timeout)


def main():
    try:
        import win32com.client
        import win32print
        import pythoncom
    except ImportError:
        print("HATA: pywin32 kurulu değil. CMD'de şunu çalıştır:")
        print("    pip install pywin32")
        sys.exit(1)

    if not OLUSTURULAN_DIR.exists():
        raise SystemExit("HATA: OLUSTURULAN klasörü yok. Önce zimmet_olustur.py çalıştır.")

    YAZDIRILDI_DIR.mkdir(exist_ok=True)

    removed = deduplicate()
    if removed:
        print(f"Bilgi: {removed} belge daha önce yazdırılmıştı, OLUSTURULAN'dan silindi.")

    files = sorted(OLUSTURULAN_DIR.glob("*.docx"), key=tr_sort_key)
    if not files:
        print("Yazdırılacak belge yok.")
        return

    try:
        printer = win32print.GetDefaultPrinter()
    except Exception as e:
        raise SystemExit(f"HATA: Varsayılan yazıcı alınamadı: {e}")

    print(f"Klasör           : {BASE_DIR.name}")
    print(f"Varsayılan yazıcı: {printer}")
    print(f"Kalan belge      : {len(files)}")
    print(f"Zaman aşımı      : {TIMEOUT_PER_DOC} saniye/belge")
    print()
    print("Yazıcının hazır olduğundan, kağıt ve tonerin yeterli olduğundan emin ol.")
    print("Ctrl+C ile her an durdurabilirsin — yazdırılanlar YAZDIRILDI'ya alınır,")
    print("kalanlar OLUSTURULAN'da kalır, tekrar çalıştırınca kaldığı yerden devam eder.")
    print()
    try:
        secim = input(
            f"Kaç tane basılsın? [Enter = hepsi ({len(files)}), "
            f"sayı = ilk N, iptal için Ctrl+C]: "
        ).strip()
    except KeyboardInterrupt:
        print("\nİptal edildi.")
        return

    if secim:
        try:
            n = int(secim)
        except ValueError:
            print(f"Sayı anlaşılamadı: {secim!r}, iptal edildi.")
            return
        if n < 1:
            print("En az 1 belge gerekli, iptal edildi.")
            return
        if n < len(files):
            files = files[:n]
            print(f"-> Bu seferki batch: {len(files)} belge "
                  f"(kalan {len(list(OLUSTURULAN_DIR.glob('*.docx'))) - n} sonraki çalıştırmada).")
        else:
            print(f"-> Hepsi basılacak ({len(files)} belge).")
    else:
        print(f"-> Hepsi basılacak ({len(files)} belge).")

    log(f"=== YAZDIRMA BASLADI === Klasor: {BASE_DIR.name}, Belge: {len(files)}, Yazici: {printer}")

    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # wdAlertsNone — popup'ları bastır
    except Exception as e:
        log(f"HATA: Word baslatilamadi: {e}")
        pythoncom.CoUninitialize()
        raise SystemExit(f"HATA: Word başlatılamadı: {e}\nWord yüklü mü?")

    printed = 0
    failed_file = None
    failure_reason = None

    try:
        for i, f in enumerate(files, 1):
            print(f"[{i}/{len(files)}] {f.name} ... ", end="", flush=True)
            try:
                success, error = print_document(word, f, printer, TIMEOUT_PER_DOC)
            except KeyboardInterrupt:
                print("IPTAL (Ctrl+C)")
                failure_reason = "kullanıcı iptal etti"
                failed_file = f.name
                break

            if success:
                dest = YAZDIRILDI_DIR / f.name
                try:
                    shutil.move(str(f), str(dest))
                    print("OK")
                    log(f"YAZDIRILDI: {f.name}")
                    printed += 1
                except Exception as e:
                    print(f"TASIMA HATASI: {e}")
                    log(f"HATA (tasima): {f.name}: {e}")
                    failed_file = f.name
                    failure_reason = f"taşıma hatası: {e}"
                    break

                # Her PAUSE_EVERY belgede bir duraklayıp yazıcının dinlenmesi için bekle
                if (PAUSE_EVERY > 0
                        and printed % PAUSE_EVERY == 0
                        and i < len(files)):
                    kalan = len(files) - i
                    print()
                    print(f"--- {printed} belge yazdırıldı, yazıcı dinlensin ({kalan} kaldı) ---")
                    try:
                        input("Devam etmek için Enter (durmak için Ctrl+C): ")
                    except KeyboardInterrupt:
                        print()
                        print("Duraklamada durduruldu. Kalan belgeler OLUSTURULAN'da.")
                        log(f"=== DURAKLAMADA DURDURULDU === Yazdirilan: {printed}, Kalan: {kalan}")
                        failed_file = "(kullanıcı duraklamada durdu)"
                        failure_reason = "kullanıcı duraklamada Ctrl+C"
                        break
                    print()
            else:
                print(f"HATA: {error}")
                log(f"HATA (yazdirma): {f.name}: {error}")
                failed_file = f.name
                failure_reason = error
                break
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    remaining = len(list(OLUSTURULAN_DIR.glob("*.docx")))
    print()
    print(f"Yazdırılan : {printed}")
    print(f"Kalan      : {remaining}")
    if failed_file:
        print(f"DURDU      : {failed_file} -> {failure_reason}")
        print("Yazıcıyı kontrol et, sonra scripti tekrar çalıştır. Kaldığı yerden devam eder.")
        log(f"=== DURDU === Yazdirilan: {printed}, Kalan: {remaining}, Sebep: {failure_reason}")
    else:
        log(f"=== TAMAMLANDI === Yazdirilan: {printed}")


if __name__ == "__main__":
    main()
