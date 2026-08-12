import copy
import datetime as _datetime
import hashlib
import math
import os
import re
import uuid


ZEMIN_SINIFLARI = ("ZA", "ZB", "ZC", "ZD", "ZE", "ZF")

IS_DURUMLARI = {
    "belirlenmedi": "Durum Belirlenmedi",
    "yeni": "Yeni İş",
    "on_deger_verildi": "Ön Değer Verildi",
    "yazim_asamasinda": "Yazım Aşamasında",
    "duzeltme_asamasinda": "Düzeltme Aşamasında",
    "bitti": "Bitti",
}

NORMAL_DURUM_GECISLERI = {
    "belirlenmedi": {"yeni", "on_deger_verildi", "yazim_asamasinda", "bitti"},
    "yeni": {"on_deger_verildi", "yazim_asamasinda"},
    "on_deger_verildi": {"yazim_asamasinda"},
    "yazim_asamasinda": {"bitti"},
    "duzeltme_asamasinda": {"bitti"},
    "bitti": {"duzeltme_asamasinda"},
}


def simdi_iso():
    return _datetime.datetime.now().astimezone().isoformat(timespec="seconds")


def bos_on_deger_verisi():
    return {
        "guncel": {
            "qt": "",
            "ks": "",
            "zemin_sinifi": "",
            "aciklama": "",
        },
        "ilk": None,
        "revizyonlar": [],
    }


def bos_tdth_verisi():
    return {
        "aktif": None,
        "gecmis": [],
        "durum": "eksik",
        "uyarilar": [],
    }


def bos_is_akisi_verisi(durum="yeni"):
    if durum not in IS_DURUMLARI:
        durum = "yeni"
    return {
        "proje_id": str(uuid.uuid4()),
        "durum": durum,
        "revizyon_no": 1,
        "duzeltme_nedeni": "",
        "tamamlanma_tarihi": "",
        "son_rapor_yolu": "",
        "son_nihai_pdf_yolu": "",
        "tamamlanan_revizyonlar": [],
        "gecmis": [],
    }


def _metin(deger):
    if deger is None:
        return ""
    if isinstance(deger, (dict, list, tuple)):
        return ""
    return str(deger).strip()


def _sonlu_pozitif_sayi_metni(deger, etiket):
    metin = _metin(deger)
    if not metin:
        raise ValueError(f"{etiket} boş bırakılamaz.")
    try:
        sayi = float(metin.replace(",", "."))
    except ValueError:
        raise ValueError(f"{etiket} sayısal olmalıdır.") from None
    if not math.isfinite(sayi) or sayi <= 0:
        raise ValueError(f"{etiket} sonlu ve sıfırdan büyük olmalıdır.")
    return metin


def normalize_on_deger(veri):
    sonuc = bos_on_deger_verisi()
    if not isinstance(veri, dict):
        return sonuc

    guncel = veri.get("guncel", {})
    if isinstance(guncel, dict):
        sonuc["guncel"] = {
            "qt": _metin(guncel.get("qt")),
            "ks": _metin(guncel.get("ks")),
            "zemin_sinifi": _metin(guncel.get("zemin_sinifi")).upper(),
            "aciklama": _metin(guncel.get("aciklama")),
        }

    revizyonlar = veri.get("revizyonlar", [])
    if isinstance(revizyonlar, list):
        for item in revizyonlar:
            if not isinstance(item, dict):
                continue
            zemin = _metin(item.get("zemin_sinifi")).upper()
            sonuc["revizyonlar"].append({
                "id": _metin(item.get("id")) or str(uuid.uuid4()),
                "qt": _metin(item.get("qt")),
                "ks": _metin(item.get("ks")),
                "zemin_sinifi": zemin if zemin in ZEMIN_SINIFLARI else "",
                "aciklama": _metin(item.get("aciklama")),
                "tdth_hash": _metin(item.get("tdth_hash")),
                "tarih": _metin(item.get("tarih")),
            })

    ilk = veri.get("ilk")
    if isinstance(ilk, dict):
        ilk_id = _metin(ilk.get("id"))
        sonuc["ilk"] = next(
            (copy.deepcopy(r) for r in sonuc["revizyonlar"] if r["id"] == ilk_id),
            {
                "id": ilk_id or str(uuid.uuid4()),
                "qt": _metin(ilk.get("qt")),
                "ks": _metin(ilk.get("ks")),
                "zemin_sinifi": _metin(ilk.get("zemin_sinifi")).upper(),
                "aciklama": _metin(ilk.get("aciklama")),
                "tdth_hash": _metin(ilk.get("tdth_hash")),
                "tarih": _metin(ilk.get("tarih")),
            },
        )
    elif sonuc["revizyonlar"]:
        sonuc["ilk"] = copy.deepcopy(sonuc["revizyonlar"][0])
    return sonuc


def on_deger_durumu(veri):
    """Ön değer varlığını iş aşamasından bağımsız olarak revizyon geçmişinden türet."""
    return "verildi" if normalize_on_deger(veri).get("revizyonlar") else "verilmedi"


def on_deger_revizyonu_ekle(veri, qt, ks, zemin_sinifi, aciklama="", tdth_hash=""):
    sonuc = normalize_on_deger(veri)
    qt = _sonlu_pozitif_sayi_metni(qt, "Ön qₜ")
    ks = _sonlu_pozitif_sayi_metni(ks, "Ön kₛ")
    zemin_sinifi = _metin(zemin_sinifi).upper()
    if zemin_sinifi not in ZEMIN_SINIFLARI:
        raise ValueError("Yerel zemin sınıfı ZA-ZF aralığından seçilmelidir.")
    if not _metin(tdth_hash):
        raise ValueError("Ön değer kaydı için geçerli bir TDTH PDF seçilmelidir.")

    revizyon = {
        "id": str(uuid.uuid4()),
        "qt": qt,
        "ks": ks,
        "zemin_sinifi": zemin_sinifi,
        "aciklama": _metin(aciklama),
        "tdth_hash": _metin(tdth_hash),
        "tarih": simdi_iso(),
    }
    sonuc["guncel"] = {
        "qt": qt,
        "ks": ks,
        "zemin_sinifi": zemin_sinifi,
        "aciklama": _metin(aciklama),
    }
    sonuc["revizyonlar"].append(revizyon)
    if sonuc["ilk"] is None:
        sonuc["ilk"] = copy.deepcopy(revizyon)
    return sonuc, copy.deepcopy(revizyon)


def normalize_tdth(veri):
    sonuc = bos_tdth_verisi()
    if not isinstance(veri, dict):
        return sonuc

    def temiz_kayit(item):
        if not isinstance(item, dict):
            return None
        degerler = item.get("degerler", {})
        if not isinstance(degerler, dict):
            degerler = {}
        return {
            "id": _metin(item.get("id")) or str(uuid.uuid4()),
            "pdf_yolu": _metin(item.get("pdf_yolu")),
            "orijinal_dosya_adi": _metin(item.get("orijinal_dosya_adi")),
            "sha256": _metin(item.get("sha256")),
            "sayfa_sayisi": max(0, int(item.get("sayfa_sayisi") or 0)),
            "rapor_basligi": _metin(item.get("rapor_basligi")),
            "dd_duzeyi": _metin(item.get("dd_duzeyi")).upper(),
            "zemin_sinifi": _metin(item.get("zemin_sinifi")).upper(),
            "enlem": _metin(item.get("enlem")),
            "boylam": _metin(item.get("boylam")),
            "degerler": {str(k).upper(): _metin(v) for k, v in degerler.items()},
            "ice_aktarim_tarihi": _metin(item.get("ice_aktarim_tarihi")),
        }

    sonuc["aktif"] = temiz_kayit(veri.get("aktif"))
    gecmis = veri.get("gecmis", [])
    if isinstance(gecmis, list):
        sonuc["gecmis"] = [kayit for kayit in (temiz_kayit(x) for x in gecmis) if kayit]
    durum = _metin(veri.get("durum"))
    sonuc["durum"] = durum if durum in {"eksik", "gecerli", "uyari", "yenilenmeli"} else "eksik"
    uyarilar = veri.get("uyarilar", [])
    sonuc["uyarilar"] = [_metin(x) for x in uyarilar if _metin(x)] if isinstance(uyarilar, list) else []
    if sonuc["aktif"] is None:
        sonuc["durum"] = "eksik"
    return sonuc


def normalize_is_akisi(veri, eski_proje=False):
    if not isinstance(veri, dict):
        return bos_is_akisi_verisi("belirlenmedi" if eski_proje else "yeni")
    sonuc = bos_is_akisi_verisi()
    sonuc["proje_id"] = _metin(veri.get("proje_id")) or sonuc["proje_id"]
    durum = _metin(veri.get("durum"))
    sonuc["durum"] = durum if durum in IS_DURUMLARI else ("belirlenmedi" if eski_proje else "yeni")
    try:
        sonuc["revizyon_no"] = max(1, int(veri.get("revizyon_no") or 1))
    except (TypeError, ValueError):
        sonuc["revizyon_no"] = 1
    sonuc["duzeltme_nedeni"] = _metin(veri.get("duzeltme_nedeni"))
    sonuc["tamamlanma_tarihi"] = _metin(veri.get("tamamlanma_tarihi"))
    sonuc["son_rapor_yolu"] = _metin(veri.get("son_rapor_yolu"))
    sonuc["son_nihai_pdf_yolu"] = _metin(veri.get("son_nihai_pdf_yolu"))
    tamamlanan = veri.get("tamamlanan_revizyonlar", [])
    if isinstance(tamamlanan, list):
        for item in tamamlanan:
            if not isinstance(item, dict):
                continue
            sonuc["tamamlanan_revizyonlar"].append({
                "revizyon_no": max(1, int(item.get("revizyon_no") or 1)),
                "tamamlanma_tarihi": _metin(item.get("tamamlanma_tarihi")),
                "arsiv_tarihi": _metin(item.get("arsiv_tarihi")),
                "qt_nihai": _metin(item.get("qt_nihai")),
                "ks_nihai": _metin(item.get("ks_nihai")),
                "zemin_sinifi": _metin(item.get("zemin_sinifi")),
                "tdth_hash": _metin(item.get("tdth_hash")),
                "tdth_pdf_yolu": _metin(item.get("tdth_pdf_yolu")),
                "rapor_yolu": _metin(item.get("rapor_yolu")),
                "nihai_pdf_yolu": _metin(item.get("nihai_pdf_yolu")),
            })
    gecmis = veri.get("gecmis", [])
    if isinstance(gecmis, list):
        for item in gecmis:
            if not isinstance(item, dict):
                continue
            sonuc["gecmis"].append({
                "eski": _metin(item.get("eski")),
                "yeni": _metin(item.get("yeni")),
                "tarih": _metin(item.get("tarih")),
                "aciklama": _metin(item.get("aciklama")),
                "revizyon_no": max(1, int(item.get("revizyon_no") or 1)),
            })
    return sonuc


def is_durumu_degistir(veri, yeni_durum, aciklama="", zorla=False):
    sonuc = normalize_is_akisi(veri)
    eski = sonuc["durum"]
    if yeni_durum not in IS_DURUMLARI:
        raise ValueError("Geçersiz iş aşaması.")
    if eski == yeni_durum:
        return sonuc
    if not zorla and yeni_durum not in NORMAL_DURUM_GECISLERI.get(eski, set()):
        raise ValueError(
            f"{IS_DURUMLARI.get(eski, eski)} durumundan "
            f"{IS_DURUMLARI.get(yeni_durum, yeni_durum)} durumuna geçilemez."
        )
    if eski == "bitti" and yeni_durum == "duzeltme_asamasinda":
        sonuc["revizyon_no"] += 1
        sonuc["duzeltme_nedeni"] = _metin(aciklama)
        sonuc["tamamlanma_tarihi"] = ""
        sonuc["son_rapor_yolu"] = ""
        sonuc["son_nihai_pdf_yolu"] = ""
    if yeni_durum == "bitti":
        sonuc["tamamlanma_tarihi"] = simdi_iso()
    sonuc["durum"] = yeni_durum
    sonuc["gecmis"].append({
        "eski": eski,
        "yeni": yeni_durum,
        "tarih": simdi_iso(),
        "aciklama": _metin(aciklama),
        "revizyon_no": sonuc["revizyon_no"],
    })
    return sonuc


def bitmis_revizyonu_arsivle(is_akisi, proje_verisi):
    sonuc = normalize_is_akisi(is_akisi)
    if not isinstance(proje_verisi, dict):
        return sonuc
    tasima = proje_verisi.get("_TASIMA_", {})
    if not isinstance(tasima, dict):
        tasima = {}
    tdth = normalize_tdth(proje_verisi.get("_TDTH_"))
    aktif_tdth = tdth.get("aktif") or {}
    revizyon_no = sonuc.get("revizyon_no", 1)
    mevcut = next(
        (x for x in sonuc["tamamlanan_revizyonlar"] if x.get("revizyon_no") == revizyon_no),
        None,
    )
    arsiv = {
        "revizyon_no": revizyon_no,
        "tamamlanma_tarihi": sonuc.get("tamamlanma_tarihi", ""),
        "arsiv_tarihi": simdi_iso(),
        "qt_nihai": _metin(tasima.get("qt_nihai")),
        "ks_nihai": _metin(tasima.get("ks_nihai")),
        "zemin_sinifi": _metin(proje_verisi.get("YEREL_ZEMIN_SINIFI")),
        "tdth_hash": _metin(aktif_tdth.get("sha256")),
        "tdth_pdf_yolu": _metin(aktif_tdth.get("pdf_yolu")),
        "rapor_yolu": sonuc.get("son_rapor_yolu", ""),
        "nihai_pdf_yolu": sonuc.get("son_nihai_pdf_yolu", ""),
    }
    if mevcut is None:
        sonuc["tamamlanan_revizyonlar"].append(arsiv)
    else:
        mevcut.update(arsiv)
    return sonuc


def dosya_sha256(yol):
    ozet = hashlib.sha256()
    with open(yol, "rb") as dosya:
        for blok in iter(lambda: dosya.read(1024 * 1024), b""):
            ozet.update(blok)
    return ozet.hexdigest()


def _ilk_eslesme(metin, desenler, flags=re.IGNORECASE):
    for desen in desenler:
        eslesme = re.search(desen, metin, flags)
        if eslesme:
            return _metin(eslesme.group(1))
    return ""


def tdth_pdf_bilgilerini_oku(yol):
    if not os.path.isfile(yol):
        raise ValueError("TDTH PDF dosyası bulunamadı.")
    if os.path.splitext(yol)[1].lower() != ".pdf":
        raise ValueError("TDTH raporu PDF biçiminde olmalıdır.")
    try:
        from pypdf import PdfReader
    except ImportError:
        try:
            from PyPDF2 import PdfReader
        except ImportError as exc:
            raise RuntimeError("TDTH PDF okumak için pypdf veya PyPDF2 gereklidir.") from exc

    reader = PdfReader(yol)
    if getattr(reader, "is_encrypted", False) and not reader.decrypt(""):
        raise ValueError("TDTH PDF parola korumalıdır.")
    if not reader.pages:
        raise ValueError("TDTH PDF sayfa içermiyor.")
    metin = "\n".join((sayfa.extract_text() or "") for sayfa in reader.pages)
    if not metin.strip():
        raise ValueError("TDTH PDF içindeki metin okunamadı.")

    sayi = r"([-+]?\d+(?:[.,]\d+)?)"
    degerler = {}
    etiketler = {
        "PGA": [rf"\bPGA\s*[:=]\s*{sayi}"],
        "PGV": [rf"\bPGV\s*[:=]\s*{sayi}"],
        "SS": [rf"\bS[Ss]\s*[:=]\s*{sayi}"],
        "S1": [rf"\bS1\s*[:=]\s*{sayi}"],
        "FS": [rf"\bF[Ss]\s*[:=]\s*{sayi}"],
        "F1": [rf"\bF1\s*[:=]\s*{sayi}"],
        "SDS": [rf"(?m)^\s*SDS[^\r\n]*=[ \t]*{sayi}[ \t]*$", rf"\bSDS[ \t]*:[ \t]*{sayi}"],
        "SD1": [rf"(?m)^\s*SD1[^\r\n]*=[ \t]*{sayi}[ \t]*$", rf"\bSD1[ \t]*:[ \t]*{sayi}"],
        "TA": [rf"\bTA\s*[:=]\s*{sayi}"],
        "TB": [rf"\bTB\s*[:=]\s*{sayi}"],
        "TL": [rf"\bTL\s*[:=]\s*{sayi}"],
    }
    for kod, desenler in etiketler.items():
        deger = _ilk_eslesme(metin, desenler)
        if deger:
            degerler[kod] = deger.replace(",", ".")

    rapor_basligi = _ilk_eslesme(
        metin,
        [
            r"Rapor\s*Başlığı\s*[:\-]?\s*([^\n\r]+)",
            r"Rapor\s*Basligi\s*[:\-]?\s*([^\n\r]+)",
        ],
    )
    zemin_sinifi = _ilk_eslesme(
        metin,
        [r"Yerel\s+Zemin\s+Sınıfı\s*[:\-]?\s*(Z[ABCDEF])", r"Yerel\s+Zemin\s+Sinifi\s*[:\-]?\s*(Z[ABCDEF])"],
    ).upper()
    dd_duzeyi = _ilk_eslesme(metin, [r"(DD\s*-\s*[1-4])\b"]).replace(" ", "").upper()
    enlem = _ilk_eslesme(metin, [rf"(?:Enlem|Latitude)[ \t]*[:=]?[ \t]*{sayi}"]).replace(",", ".")
    boylam = _ilk_eslesme(metin, [rf"(?:Boylam|Longitude)[ \t]*[:=]?[ \t]*{sayi}"]).replace(",", ".")

    return {
        "id": str(uuid.uuid4()),
        "pdf_yolu": os.path.abspath(yol),
        "orijinal_dosya_adi": os.path.basename(yol),
        "sha256": dosya_sha256(yol),
        "sayfa_sayisi": len(reader.pages),
        "rapor_basligi": rapor_basligi,
        "dd_duzeyi": dd_duzeyi,
        "zemin_sinifi": zemin_sinifi,
        "enlem": enlem,
        "boylam": boylam,
        "degerler": degerler,
        "ice_aktarim_tarihi": simdi_iso(),
    }


def tdth_kaydi_etkinlestir(veri, kayit, uyarilar=None):
    sonuc = normalize_tdth(veri)
    if sonuc["aktif"] and sonuc["aktif"].get("sha256") != kayit.get("sha256"):
        sonuc["gecmis"].append(copy.deepcopy(sonuc["aktif"]))
    sonuc["aktif"] = copy.deepcopy(kayit)
    sonuc["uyarilar"] = [_metin(x) for x in (uyarilar or []) if _metin(x)]
    sonuc["durum"] = "uyari" if sonuc["uyarilar"] else "gecerli"
    return sonuc


def tdth_zemin_sinifi_guncelle(veri, zemin_sinifi):
    sonuc = normalize_tdth(veri)
    aktif = sonuc.get("aktif")
    secim = _metin(zemin_sinifi).upper()
    if aktif and secim in ZEMIN_SINIFLARI and aktif.get("zemin_sinifi") and aktif.get("zemin_sinifi") != secim:
        sonuc["durum"] = "yenilenmeli"
        uyari = f"TDTH PDF zemin sınıfı {aktif.get('zemin_sinifi')}; seçilen sınıf {secim}."
        if uyari not in sonuc["uyarilar"]:
            sonuc["uyarilar"].append(uyari)
    elif aktif and secim == aktif.get("zemin_sinifi"):
        sonuc["uyarilar"] = [
            uyari for uyari in sonuc["uyarilar"]
            if not uyari.startswith("TDTH PDF zemin sınıfı ")
        ]
        sonuc["durum"] = "uyari" if sonuc["uyarilar"] else "gecerli"
    return sonuc
