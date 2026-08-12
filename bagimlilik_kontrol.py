import importlib
import importlib.metadata
import re
import subprocess
import sys


ZORUNLU_PAKETLER = {
    "ttkbootstrap": ("ttkbootstrap", "1.10.1"),
    "pandas": ("pandas", "1.5.0"),
    "tkintermapview": ("tkintermapview", "1.29"),
    "Pillow": ("PIL", "9.0.0"),
    "numpy": ("numpy", "1.23.0"),
    "python-docx": ("docx", "0.8.11"),
    "openpyxl": ("openpyxl", "3.0.10"),
}

OPSIYONEL_PAKETLER = {
    "opencv-python-headless": ("cv2", "4.8.0"),
    "pypdf": ("pypdf", "3.0.0"),
    "xlrd": ("xlrd", "2.0.1"),
}


def surum_parcalari(surum):
    """Standart paket sürümlerinin sayısal bölümünü karşılaştırılabilir yapar."""
    eslesme = re.match(r"^(\d+(?:\.\d+)*)", surum)
    if not eslesme:
        return ()
    return tuple(int(parca) for parca in eslesme.group(1).split("."))


def surum_yeterli_mi(mevcut, minimum):
    mevcut_parcalar = surum_parcalari(mevcut)
    minimum_parcalar = surum_parcalari(minimum)
    uzunluk = max(len(mevcut_parcalar), len(minimum_parcalar))
    mevcut_parcalar += (0,) * (uzunluk - len(mevcut_parcalar))
    minimum_parcalar += (0,) * (uzunluk - len(minimum_parcalar))
    return bool(mevcut_parcalar) and mevcut_parcalar >= minimum_parcalar


def paketleri_kontrol_et(paketler):
    sorunlar = []
    for paket_adi, (modul_adi, minimum_surum) in paketler.items():
        try:
            importlib.import_module(modul_adi)
        except Exception as hata:
            sorunlar.append((paket_adi, f"ice aktarilamadi: {hata}"))
            continue

        try:
            mevcut_surum = importlib.metadata.version(paket_adi)
        except importlib.metadata.PackageNotFoundError:
            sorunlar.append((paket_adi, "kurulu surum bilgisi bulunamadi"))
            continue

        if not surum_yeterli_mi(mevcut_surum, minimum_surum):
            sorunlar.append(
                (paket_adi, f"{mevcut_surum} kurulu; en az {minimum_surum} gerekli")
            )
    return sorunlar


def tkinter_kontrol_et():
    kok = None
    try:
        import tkinter as tk

        # K-1 doğrudan bir Tk ana penceresi oluşturur. Windows'ta yalnızca
        # ``Tcl()`` açıp sonradan ``package require Tk`` çalıştırmak, Python'ın
        # DLLs klasöründeki geçerli tk86t.dll yerine Tcl'in ``bin`` yolunu
        # arayabildiği için yanlış negatif sonuç üretir.
        kok = tk.Tk()
        kok.withdraw()
        kok.update_idletasks()
        tk_surumu = str(kok.tk.call("info", "patchlevel"))
        if not surum_yeterli_mi(tk_surumu, "8.6"):
            return f"Tk {tk_surumu} kurulu; en az Tk 8.6 gerekli"
    except Exception as hata:
        return f"Tk kullanilamiyor: {hata}"
    finally:
        if kok is not None:
            try:
                kok.destroy()
            except Exception:
                pass
    return None


def main():
    sorunlu_zorunlu = paketleri_kontrol_et(ZORUNLU_PAKETLER)
    sorunlu_opsiyonel = paketleri_kontrol_et(OPSIYONEL_PAKETLER)
    tkinter_sorunu = tkinter_kontrol_et()
    if tkinter_sorunu:
        sorunlu_zorunlu.insert(0, ("tkinter", tkinter_sorunu))

    if sorunlu_zorunlu:
        print("K-1 icin zorunlu Python bilesenlerinde sorun var:")
        for paket, ayrinti in sorunlu_zorunlu:
            print(f"  - {paket}: {ayrinti}")
        if sorunlu_zorunlu != [("tkinter", tkinter_sorunu)]:
            print()
            print("Eksik Python paketleri icin bu klasorde su komutu calistirin:")
            komut = subprocess.list2cmdline(
                [sys.executable, "-m", "pip", "install", "-r", "requirements.txt"]
            )
            print(f"  {komut}")
        if tkinter_sorunu:
            print()
            print(
                "Tkinter sorunu pip ile giderilemez; Python kurulumunda "
                "Tcl/Tk destegini Modify/Repair ile onarin."
            )
        return 1

    if sorunlu_opsiyonel:
        print("Uyari: Bazi istege bagli paketlerde sorun var:")
        for paket, ayrinti in sorunlu_opsiyonel:
            print(f"  - {paket}: {ayrinti}")
        print("Program acilir; ancak ilgili ozellikler sinirli calisabilir.")
        print()

    print("Bagimlilik kontrolu tamam.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
