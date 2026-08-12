"""İlçe klasöründeki proje paketlerini bulup toplu jeoloji aktarımına hazırlar."""

from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path

from harita_islemleri import kml_dosyalari_bul, kml_poligonlarini_oku
from jeoloji_kutuphanesi import jeoloji_anahtari
from jeoloji_kunye_uzlasma import kunye_uzlasmasi_olustur
from jeoloji_word_aktarimi import word_dosyalari_bul, word_rapor_adaylarini_oku


ILCELER = (
    "Merkez", "Ayvacık", "Bayramiç", "Biga", "Bozcaada", "Çan", "Eceabat",
    "Ezine", "Gelibolu", "Gökçeada", "Lapseki", "Yenice",
)
ILCE_ADLARI = {jeoloji_anahtari(value): value for value in ILCELER}
YARDIMCI_KLASORLER = {
    "jeofizik", "lab", "laboratuvar", "evrak", "evraklar", "fotograf",
    "fotograflar", "resim", "resimler", "video", "videolar", "ek", "ekler",
}


class TopluTaramaIptalEdildi(RuntimeError):
    pass


def _iptal_kontrol(iptal_event):
    if iptal_event is not None and iptal_event.is_set():
        raise TopluTaramaIptalEdildi("Toplu klasör taraması kullanıcı tarafından iptal edildi.")


def _ilerleme(callback, mevcut, toplam, mesaj):
    if callback is not None:
        callback(int(mevcut), int(toplam), str(mesaj))


def _altinda(path, root):
    try:
        hedef = os.path.normcase(os.path.abspath(os.fspath(path)))
        kok = os.path.normcase(os.path.abspath(os.fspath(root)))
        return os.path.commonpath((hedef, kok)) == kok
    except (OSError, ValueError, TypeError):
        return False


def _yardimci_klasorde(path, root):
    try:
        parts = Path(os.path.abspath(path)).relative_to(Path(os.path.abspath(root))).parts[:-1]
    except (OSError, ValueError, TypeError):
        parts = Path(path).parts[:-1]
    return any(jeoloji_anahtari(part) in YARDIMCI_KLASORLER for part in parts)


def _dosya_hashi(path):
    hasher = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def _formasyon_secimi_gecerli(value):
    return jeoloji_anahtari(value) not in ("", "seciniz")


def _formasyon_uyarisi_mi(value):
    key = jeoloji_anahtari(value)
    return key.startswith("formasyon adi algilanamadi") or key.startswith(
        "birden fazla formasyon adayi bulundu"
    )


def toplu_proje_formasyonunu_belirle(proje, formasyon, kutuphane=None):
    """Toplu önizlemedeki çözülememiş formasyonu kullanıcı seçimiyle tamamlar."""
    value = str(formasyon or "").strip()
    if not _formasyon_secimi_gecerli(value):
        raise ValueError("Geçerli bir formasyon seçin veya adını yazın.")
    record = proje.get("record") or {}
    if not record or proje.get("kml_adayi") is None:
        raise ValueError("Formasyon seçilebilmesi için Word ve KML eşleşmesi tamamlanmalıdır.")

    record["formasyon"] = value
    word_adayi = proje.get("word_adayi") or {}
    word = word_adayi.get("sonuc")
    if word is not None:
        word.formasyon = value
        word.formasyon_adaylari = (value,)
        word.uyarilar = tuple(
            warning for warning in word.uyarilar if not _formasyon_uyarisi_mi(warning)
        )
        record["notlar"] = word.kutuphane_kaydi()["notlar"]
    proje["uyarilar"] = [
        warning for warning in proje.get("uyarilar", ()) if not _formasyon_uyarisi_mi(warning)
    ]

    tekrar = (
        kutuphane.toplu_kayit_durumu(
            record,
            proje.get("kaynak_rapor_hash", ""),
            proje.get("kml_hash", ""),
        )
        if kutuphane is not None
        else {"durum": "yeni", "kayit": None}
    )
    if tekrar["durum"] == "ayni":
        proje["durum_kodu"] = "mevcut"
        proje["durum"] = "Zaten kütüphanede"
        proje["guncellenecek_kayit_id"] = None
        proje["secili"] = False
    elif tekrar["durum"] == "revizyon":
        proje["durum_kodu"] = "revizyon"
        proje["durum"] = "Olası yeni revizyon"
        proje["guncellenecek_kayit_id"] = None
        proje["secili"] = False
    else:
        proje["durum_kodu"] = "hazir"
        if proje.get("guncellenecek_kayit_id"):
            proje["durum"] = "Hazır · Eksik formasyon tamamlandı"
        else:
            proje["durum"] = "Hazır · Formasyon kullanıcı tarafından seçildi"
        proje["secili"] = True
    return proje


def _ilce_ve_yerlesim_yedegi(secili_root, proje_root):
    root = Path(secili_root)
    project = Path(proje_root)
    ilce = ILCE_ADLARI.get(jeoloji_anahtari(root.name), "")
    if not ilce:
        for part in reversed(root.parts):
            ilce = ILCE_ADLARI.get(jeoloji_anahtari(part), "")
            if ilce:
                break
    yerlesim = ""
    try:
        relative = project.relative_to(root)
        if len(relative.parts) >= 2:
            yerlesim = relative.parts[-2]
    except ValueError:
        pass
    if jeoloji_anahtari(yerlesim) in ILCE_ADLARI:
        yerlesim = ""
    return ilce, yerlesim


def _proje_koku_birlestir(root_bilgileri):
    """Aynı metadata anahtarına sahip iç içe Word köklerinden dıştakini korur."""
    secilen = []
    for info in sorted(root_bilgileri, key=lambda item: len(Path(item["klasor"]).parts)):
        top = info["word_adaylari"][0] if info["word_adaylari"] else None
        sonuc = top["sonuc"] if top else None
        key = (
            jeoloji_anahtari(getattr(sonuc, "ilce", "")),
            jeoloji_anahtari(getattr(sonuc, "yerlesim", "")),
            jeoloji_anahtari(getattr(sonuc, "ada", "")),
            jeoloji_anahtari(getattr(sonuc, "parsel", "")),
        )
        cakisan = next(
            (
                existing
                for existing in secilen
                if key != ("", "", "", "")
                and existing["metadata_key"] == key
                and _altinda(info["klasor"], existing["klasor"])
            ),
            None,
        )
        if cakisan is None:
            item = dict(info)
            item["metadata_key"] = key
            secilen.append(item)
    return secilen


def ilce_klasorunu_tara(root_path, kutuphane=None, ilerleme=None, iptal_event=None):
    """İlçe kökündeki proje klasörlerini ve kesin Word-KML çiftlerini hazırlar."""
    root = Path(root_path)
    if not root.is_dir():
        raise ValueError("Seçilen ilçe klasörü bulunamadı.")
    _ilerleme(ilerleme, 0, 1, "Dosya adları taranıyor")
    ham_word = []
    ham_kml = []
    for current, directories, filenames in os.walk(
        root, topdown=True, onerror=lambda _error: None, followlinks=False
    ):
        _iptal_kontrol(iptal_event)
        directories[:] = sorted(
            (name for name in directories if not name.startswith(".")), key=str.casefold
        )
        current_path = Path(current)
        for filename in filenames:
            if filename.startswith(".") or filename.startswith("~$"):
                continue
            path = current_path / filename
            suffix = path.suffix.lower()
            if suffix == ".docx":
                ham_word.append(str(path))
            elif suffix == ".kml":
                ham_kml.append(str(path))

    word_paths = [
        path for path in word_dosyalari_bul(ham_word) if not _yardimci_klasorde(path, root)
    ]
    word_groups = {}
    for path in word_paths:
        word_groups.setdefault(str(Path(path).parent), []).append(path)

    root_bilgileri = []
    total_groups = max(len(word_groups), 1)
    for index, (folder, paths) in enumerate(sorted(word_groups.items()), start=1):
        _iptal_kontrol(iptal_event)
        _ilerleme(ilerleme, index - 1, total_groups, f"Word adayları okunuyor: {Path(folder).name}")
        adaylar = word_rapor_adaylarini_oku(paths)
        if adaylar and adaylar[0]["puan"] >= 60:
            root_bilgileri.append({"klasor": folder, "word_adaylari": adaylar})
    root_bilgileri = _proje_koku_birlestir(root_bilgileri)

    _ilerleme(ilerleme, 0, max(len(ham_kml), 1), "KML dosyaları doğrulanıyor")
    gecerli_kml = []
    gecersiz_kml = []
    for index, path in enumerate(kml_dosyalari_bul(ham_kml), start=1):
        _iptal_kontrol(iptal_event)
        _ilerleme(ilerleme, index - 1, max(len(ham_kml), 1), f"KML okunuyor: {Path(path).name}")
        try:
            poligonlar = kml_poligonlarini_oku(path)
        except (OSError, ValueError) as exc:
            gecersiz_kml.append({"dosya_yolu": path, "hata": str(exc)})
            continue
        gecerli_kml.append({"dosya_yolu": path, "poligonlar": poligonlar})

    # KML'leri yalnızca içinde bulundukları en yakın proje köküne bağla.
    for info in root_bilgileri:
        info["kml_adaylari"] = []
        info["kml_hatalari"] = []
    sahipsiz_kml = []
    for aday in gecerli_kml:
        sahipler = [info for info in root_bilgileri if _altinda(aday["dosya_yolu"], info["klasor"])]
        if sahipler:
            sahip = max(sahipler, key=lambda item: len(Path(item["klasor"]).parts))
            sahip["kml_adaylari"].append(aday)
        else:
            sahipsiz_kml.append(aday)
    for hata in gecersiz_kml:
        sahipler = [info for info in root_bilgileri if _altinda(hata["dosya_yolu"], info["klasor"])]
        if sahipler:
            sahip = max(sahipler, key=lambda item: len(Path(item["klasor"]).parts))
            sahip["kml_hatalari"].append(hata)

    # Ana raporu bulunmayan KML klasörlerini de toplu sonuçta görünür kıl.
    for aday in sahipsiz_kml:
        folder = str(Path(aday["dosya_yolu"]).parent)
        mevcut = next((info for info in root_bilgileri if info["klasor"] == folder), None)
        if mevcut is None:
            root_bilgileri.append(
                {"klasor": folder, "word_adaylari": [], "kml_adaylari": [aday], "kml_hatalari": []}
            )

    projeler = []
    total_projects = max(len(root_bilgileri), 1)
    for index, info in enumerate(sorted(root_bilgileri, key=lambda item: item["klasor"]), start=1):
        _iptal_kontrol(iptal_event)
        _ilerleme(ilerleme, index - 1, total_projects, f"Eşleştiriliyor: {Path(info['klasor']).name}")
        word_adaylari = info.get("word_adaylari", [])
        top = word_adaylari[0] if word_adaylari else None
        second = word_adaylari[1] if len(word_adaylari) > 1 else None
        word = top["sonuc"] if top else None
        ilce_yedek, yerlesim_yedek = _ilce_ve_yerlesim_yedegi(root, info["klasor"])
        durum_kodu = "hazir"
        durum = "Hazır"
        uyarilar = []
        record = None
        kml_adayi = None
        eslesme = None

        if word is None or top["puan"] < 100 or not (
            word.genel_jeoloji_metni or word.inceleme_alani_jeolojisi
        ):
            durum_kodu, durum = "ana_rapor_yok", "Ana rapor bulunamadı"
        elif second is not None and top["puan"] - second["puan"] < 40:
            durum_kodu, durum = "word_belirsiz", "Birden fazla Word adayı"
        else:
            record = word.kutuphane_kaydi()
            record["il"] = record.get("il") or "Çanakkale"
            record["ilce"] = record.get("ilce") or ilce_yedek
            record["yerlesim"] = record.get("yerlesim") or yerlesim_yedek
            record["kaynak_klasor_path"] = info["klasor"]
            if not record["ilce"]:
                durum_kodu, durum = "kunye_eksik", "İlçe belirlenemedi"

        if durum_kodu == "hazir":
            degerlendirilen_kml = []
            for aday in info.get("kml_adaylari", []):
                aday_eslesmesi = kunye_uzlasmasi_olustur(
                    secili_root=root,
                    proje_klasoru=info["klasor"],
                    word_sonucu=word,
                    kml_adayi=aday,
                )
                item = dict(aday)
                item["eslesme"] = aday_eslesmesi
                degerlendirilen_kml.append(item)
            durum_sirasi = {"tam": 4, "ada0": 3, "duzeltildi": 3, "belirsiz": 1, "celiski": 0}
            degerlendirilen_kml.sort(
                key=lambda item: (
                    -durum_sirasi.get(item["eslesme"]["durum"], 0),
                    -int(item["eslesme"].get("guven_puani", 0)),
                    os.path.normcase(item["dosya_yolu"]).casefold(),
                )
            )
            if not degerlendirilen_kml:
                if info.get("kml_hatalari"):
                    durum_kodu, durum = "kml_gecersiz", "KML geçersiz"
                else:
                    durum_kodu, durum = "kml_yok", "KML bulunamadı"
            else:
                kml_adayi = degerlendirilen_kml[0]
                eslesme = kml_adayi["eslesme"]
                ikinci_kml = degerlendirilen_kml[1] if len(degerlendirilen_kml) > 1 else None
                if (
                    ikinci_kml is not None
                    and eslesme["hazir"]
                    and ikinci_kml["eslesme"]["hazir"]
                    and eslesme["guven_puani"] - ikinci_kml["eslesme"]["guven_puani"] < 2
                ):
                    durum_kodu, durum = "kml_belirsiz", "Birden fazla KML adayı"
                elif eslesme["durum"] == "celiski":
                    durum_kodu, durum = "kml_uyusmuyor", "Ada/parsel kaynakları açıkça çelişiyor"
                elif not eslesme["hazir"]:
                    durum_kodu, durum = "kml_belirsiz", "Ada/parsel eşleşmesi belirsiz"
                else:
                    canonical = eslesme["kanonik"]
                    for field in ("ilce", "yerlesim", "ada", "parsel"):
                        if canonical.get(field):
                            record[field] = canonical[field]
                    record["kunye_kaynaklari_json"] = json.dumps(
                        {
                            "kaynaklar": eslesme["kaynaklar"],
                            "kanonik": canonical,
                            "durum": eslesme["durum"],
                        },
                        ensure_ascii=False,
                        sort_keys=True,
                    )
                    record["kunye_duzeltme_notu"] = "\n".join(eslesme["uyarilar"])
                    uyarilar.extend(eslesme["uyarilar"])
                    if eslesme["durum"] == "duzeltildi":
                        durum = "Hazır · Word künyesi düzeltildi"
                    elif eslesme["durum"] == "ada0":
                        durum = "Hazır · Ada 0 / parsel eşleşti"

        kaynak_hash = ""
        kml_hash = ""
        guncellenecek_kayit_id = None
        if durum_kodu == "hazir" and record is not None and kml_adayi is not None:
            kaynak_hash = kutuphane.kaynak_dosya_hashi(word.dosya_yolu) if kutuphane else _dosya_hashi(word.dosya_yolu)
            kml_hash = _dosya_hashi(kml_adayi["dosya_yolu"])
            record["kaynak_rapor_hash"] = kaynak_hash
            if not _formasyon_secimi_gecerli(record.get("formasyon")):
                if kutuphane is not None:
                    bos_formasyon = kutuphane.toplu_kayit_durumu(record, kaynak_hash, kml_hash)
                    if bos_formasyon.get("kayit"):
                        guncellenecek_kayit_id = bos_formasyon["kayit"].get("id")
                aday_sayisi = len(getattr(word, "formasyon_adaylari", ()) or ())
                durum_kodu = "formasyon_gerekli"
                durum = (
                    f"Formasyon seçimi gerekli ({aday_sayisi} aday)"
                    if aday_sayisi
                    else "Formasyon seçimi gerekli"
                )
            elif kutuphane is not None:
                tekrar = kutuphane.toplu_kayit_durumu(record, kaynak_hash, kml_hash)
                if tekrar["durum"] == "ayni":
                    durum_kodu, durum = "mevcut", "Zaten kütüphanede"
                elif tekrar["durum"] == "revizyon":
                    durum_kodu, durum = "revizyon", "Olası yeni revizyon"

        try:
            relative = str(Path(info["klasor"]).relative_to(root))
        except ValueError:
            relative = Path(info["klasor"]).name
        projeler.append(
            {
                "klasor": info["klasor"],
                "goreli_klasor": relative,
                "word_adaylari": word_adaylari,
                "word_adayi": top,
                "kml_adaylari": info.get("kml_adaylari", []),
                "kml_adayi": kml_adayi,
                "eslesme": eslesme,
                "eslesme_durumu": eslesme.get("durum", "") if eslesme else "",
                "record": record,
                "kaynak_rapor_hash": kaynak_hash,
                "kml_hash": kml_hash,
                "guncellenecek_kayit_id": guncellenecek_kayit_id,
                "durum_kodu": durum_kodu,
                "durum": durum,
                "uyarilar": uyarilar,
                "secili": durum_kodu == "hazir",
            }
        )
    _ilerleme(ilerleme, total_projects, total_projects, f"{len(projeler)} proje klasörü değerlendirildi")
    return {"root": str(root.resolve()), "projeler": projeler}


def toplu_kayitlari_aktar(projeler, kutuphane, onayli=True, ilerleme=None, iptal_event=None):
    """Seçili ve hazır proje paketlerini birbirinden bağımsız olarak kaydeder."""
    secilen = [p for p in projeler if p.get("secili") and p.get("durum_kodu") == "hazir"]
    sonuc = {"basarili": [], "atlanan": [], "hatali": []}
    total = max(len(secilen), 1)
    for index, proje in enumerate(secilen, start=1):
        _iptal_kontrol(iptal_event)
        _ilerleme(ilerleme, index - 1, total, f"Aktarılıyor: {proje['goreli_klasor']}")
        if not _formasyon_secimi_gecerli((proje.get("record") or {}).get("formasyon")):
            proje["durum_kodu"] = "formasyon_gerekli"
            proje["durum"] = "Formasyon seçimi gerekli"
            proje["secili"] = False
            sonuc["atlanan"].append({"proje": proje, "neden": "formasyon_gerekli"})
            continue
        record = dict(proje["record"])
        record["onay_durumu"] = "onayli" if onayli else "taslak"
        guncellenecek_id = proje.get("guncellenecek_kayit_id")
        try:
            if guncellenecek_id:
                mevcut = kutuphane.getir(guncellenecek_id)
                if mevcut is None or _formasyon_secimi_gecerli(mevcut.get("formasyon")):
                    sonuc["atlanan"].append(
                        {"proje": proje, "neden": "guncellenecek_kayit_degisti"}
                    )
                    continue
                record_id = kutuphane.kaydet(record, kayit_id=guncellenecek_id)
                kutuphane.geometrileri_degistir(
                    record_id,
                    proje["kml_adayi"]["poligonlar"],
                    proje["kml_adayi"]["dosya_yolu"],
                )
            else:
                tekrar = kutuphane.toplu_kayit_durumu(
                    record, proje["kaynak_rapor_hash"], proje["kml_hash"]
                )
                if tekrar["durum"] != "yeni":
                    sonuc["atlanan"].append({"proje": proje, "neden": tekrar["durum"]})
                    continue
                record_id = kutuphane.yeni_kaydi_geometriyle_kaydet(
                    record,
                    proje["kml_adayi"]["poligonlar"],
                    proje["kml_adayi"]["dosya_yolu"],
                )
        except Exception as exc:
            sonuc["hatali"].append({"proje": proje, "hata": str(exc)})
            continue
        proje["durum_kodu"] = "aktarildi"
        proje["durum"] = f"Aktarıldı (#{record_id})"
        proje["secili"] = False
        sonuc["basarili"].append({"proje": proje, "kayit_id": record_id})
    _ilerleme(ilerleme, total, total, "Toplu aktarım tamamlandı")
    return sonuc


__all__ = [
    "TopluTaramaIptalEdildi",
    "ilce_klasorunu_tara",
    "toplu_proje_formasyonunu_belirle",
    "toplu_kayitlari_aktar",
]
