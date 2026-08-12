"""Klasör, Word ve KML kaynaklarından denetlenebilir kanonik proje künyesi üretir."""

from __future__ import annotations

from pathlib import Path
import re

from jeoloji_kutuphanesi import jeoloji_anahtari


ILCELER = (
    "Merkez", "Ayvacık", "Bayramiç", "Biga", "Bozcaada", "Çan", "Eceabat",
    "Ezine", "Gelibolu", "Gökçeada", "Lapseki", "Yenice",
)
ILCE_ADLARI = {jeoloji_anahtari(value): value for value in ILCELER}
KAYNAK_AGIRLIKLARI = {
    "kml": 4,
    "proje_klasoru": 3,
    "word_dosya_adi": 3,
    "word_icerigi": 1,
    "ilce_klasoru": 5,
}


def _temiz(value):
    return " ".join(str(value or "").replace("_", " ").split()).strip(" -_/.,;:")


def _ilce_bul(path):
    for part in reversed(Path(path).parts):
        result = ILCE_ADLARI.get(jeoloji_anahtari(part))
        if result:
            return result
    return ""


def _yerlesim_metinden(text, ilce=""):
    clean = _temiz(text)
    mahalle = re.search(
        r"([^/\\,;]+?)\s+(?:Mahallesi|Köyü)\b", clean, re.IGNORECASE
    )
    if mahalle:
        value = _temiz(mahalle.group(1).split("/")[-1].split("\\")[-1])
        if ilce and jeoloji_anahtari(value).startswith(jeoloji_anahtari(ilce) + " "):
            value = _temiz(value[len(ilce):])
        return value
    if ilce:
        match = re.search(
            rf"\b{re.escape(ilce)}\s+(.+?)\s+(?=\d{{1,6}}(?:\s*[-/]\s*\d{{1,6}})?\b)",
            clean,
            re.IGNORECASE,
        )
        if match:
            return _temiz(match.group(1))
    return ""


def metinden_kunye_adayi(text, *, ilce=""):
    """Dosya/klasör metninden açık ada-parsel ve tek parsel adayını çıkarır."""
    clean = _temiz(text)
    numeric_text = re.sub(
        r"\b(?:19|20)\d{2}[-/.]\d{1,2}[-/.]\d{1,2}\b|"
        r"\b\d{1,2}[-/.]\d{1,2}[-/.](?:19|20)\d{2}\b",
        " ",
        clean,
    )
    result = {"ilce": ilce, "yerlesim": _yerlesim_metinden(clean, ilce), "ada": "", "parsel": "", "tek_parsel": ""}
    word_pair = list(
        re.finditer(
            r"(?<!\d)(\d{1,6}[A-Za-z]?)\s*[-_]?\s*ada\s*[,;:/-]?\s*"
            r"(\d{1,7}[A-Za-z]?)\s*[-_]?\s*(?:numaralı\s+)?parsel\b",
            numeric_text,
            re.IGNORECASE,
        )
    )
    if word_pair:
        result["ada"], result["parsel"] = word_pair[-1].group(1), word_pair[-1].group(2)
        return result
    pairs = list(
        re.finditer(
            r"(?<!\d)(\d{1,6}[A-Za-z]?)\s*[-/]\s*(\d{1,7}[A-Za-z]?)"
            r"(?!\s*[-/]\s*\d)",
            numeric_text,
            re.IGNORECASE,
        )
    )
    if pairs:
        result["ada"], result["parsel"] = pairs[-1].group(1), pairs[-1].group(2)
        return result
    parcel_matches = list(
        re.finditer(
            r"(?<!\d)(\d{1,7}[A-Za-z]?)\s*[-_]?\s*(?:numaralı\s+)?parsel\b",
            numeric_text,
            re.IGNORECASE,
        )
    )
    if parcel_matches:
        result["parsel"] = parcel_matches[-1].group(1)
        return result
    singles = re.findall(r"(?<!\d)(\d{1,7})(?!\d)", numeric_text)
    if singles:
        result["tek_parsel"] = singles[-1]
    return result


def _kml_metni(kml_adayi):
    if not kml_adayi:
        return ""
    parts = [Path(kml_adayi.get("dosya_yolu", "")).stem]
    for polygon in kml_adayi.get("poligonlar", ()):
        parts.extend((polygon.get("ad", ""), polygon.get("aciklama", "")))
    return " ".join(part for part in parts if part)


def _kanonik_deger(kaynaklar, field):
    totals = {}
    raw_values = {}
    sources = {}
    for source, data in kaynaklar.items():
        value = _temiz(data.get(field, ""))
        if not value:
            continue
        key = jeoloji_anahtari(value)
        totals[key] = totals.get(key, 0) + KAYNAK_AGIRLIKLARI.get(source, 1)
        raw_values.setdefault(key, value)
        sources.setdefault(key, []).append(source)
    if not totals:
        return "", [], False
    ranked = sorted(totals.items(), key=lambda item: (-item[1], item[0]))
    tie = len(ranked) > 1 and ranked[0][1] == ranked[1][1]
    key = ranked[0][0]
    return raw_values[key], sources[key], tie


def kunye_uzlasmasi_olustur(*, secili_root, proje_klasoru, word_sonucu, kml_adayi):
    """Kaynakları karşılaştırıp KML ile semantik eşleşme ve kanonik künye üretir."""
    ilce = _ilce_bul(secili_root) or _temiz(getattr(word_sonucu, "ilce", ""))
    project_path = Path(proje_klasoru)
    project_data = metinden_kunye_adayi(project_path.name, ilce=ilce)
    try:
        relative = project_path.relative_to(Path(secili_root))
        if len(relative.parts) >= 2:
            parent_settlement = _temiz(relative.parts[-2])
            if jeoloji_anahtari(parent_settlement) not in ILCE_ADLARI:
                project_data["yerlesim"] = parent_settlement
    except ValueError:
        pass
    word_file_data = metinden_kunye_adayi(
        Path(getattr(word_sonucu, "dosya_yolu", "")).stem, ilce=ilce
    )
    word_content = {
        "ilce": _temiz(getattr(word_sonucu, "ilce", "")),
        "yerlesim": _temiz(getattr(word_sonucu, "yerlesim", "")),
        "ada": _temiz(getattr(word_sonucu, "ada", "")),
        "parsel": _temiz(getattr(word_sonucu, "parsel", "")),
        "tek_parsel": "",
    }
    kml_data = metinden_kunye_adayi(_kml_metni(kml_adayi), ilce=ilce)
    kml_parsel = kml_data.get("parsel") or kml_data.get("tek_parsel")
    kml_data["parsel"] = kml_parsel
    for data in (project_data, word_file_data):
        if not data.get("parsel") and data.get("tek_parsel") and kml_parsel:
            if jeoloji_anahtari(data["tek_parsel"]) == jeoloji_anahtari(kml_parsel):
                data["parsel"] = data["tek_parsel"]
                data["ada"] = "0"
    root_data = {"ilce": ilce, "yerlesim": "", "ada": "", "parsel": "", "tek_parsel": ""}
    kaynaklar = {
        "ilce_klasoru": root_data,
        "proje_klasoru": project_data,
        "word_dosya_adi": word_file_data,
        "word_icerigi": word_content,
        "kml": kml_data,
    }
    canonical = {}
    supports = {}
    ties = {}
    for field in ("ilce", "yerlesim", "ada", "parsel"):
        canonical[field], supports[field], ties[field] = _kanonik_deger(kaynaklar, field)
    if not canonical["ada"] and canonical["parsel"] and not kml_data.get("ada"):
        zero_sources = [
            source for source in ("proje_klasoru", "word_dosya_adi", "word_icerigi")
            if kaynaklar[source].get("ada") == "0"
        ]
        if zero_sources:
            canonical["ada"] = "0"
            supports["ada"] = zero_sources

    warnings = []
    conflicts = []
    if not kml_parsel:
        conflicts.append("KML dosya adı veya Placemark içinde parsel numarası belirlenemedi.")
    elif canonical["parsel"] and jeoloji_anahtari(canonical["parsel"]) != jeoloji_anahtari(kml_parsel):
        conflicts.append(
            f"KML parseli {kml_parsel}, diğer kaynakların desteklediği {canonical['parsel']} ile çelişiyor."
        )
    if kml_data.get("ada") and canonical["ada"]:
        if jeoloji_anahtari(kml_data["ada"]) != jeoloji_anahtari(canonical["ada"]):
            conflicts.append(
                f"KML adası {kml_data['ada']}, diğer kaynakların desteklediği {canonical['ada']} ile çelişiyor."
            )
    corrections = []
    labels = {"ilce": "ilçe", "yerlesim": "yerleşim", "ada": "ada", "parsel": "parsel"}
    for field in ("ilce", "yerlesim", "ada", "parsel"):
        old = word_content.get(field, "")
        new = canonical.get(field, "")
        if old and new and jeoloji_anahtari(old) != jeoloji_anahtari(new):
            message = f"Word içeriğindeki {labels[field]} '{old}' yerine '{new}' kullanıldı."
            corrections.append(message)
            warnings.append(message)
    if canonical.get("ada") == "0" and not kml_data.get("ada") and kml_parsel:
        warnings.append("KML ada belirtmiyor; ada 0 diğer proje kaynaklarından doğrulandı.")

    parcel_support = len(supports.get("parsel", []))
    ambiguous = bool(ties.get("parsel")) or parcel_support < 2 or not canonical.get("parsel")
    if conflicts:
        status = "celiski"
    elif ambiguous:
        status = "belirsiz"
    elif corrections:
        status = "duzeltildi"
    elif canonical.get("ada") == "0" and not kml_data.get("ada"):
        status = "ada0"
    else:
        status = "tam"
    return {
        "durum": status,
        "hazir": status in ("tam", "ada0", "duzeltildi"),
        "kanonik": canonical,
        "kaynaklar": kaynaklar,
        "destekler": supports,
        "uyarilar": warnings,
        "celiskiler": conflicts,
        "guven_puani": parcel_support + len(supports.get("yerlesim", [])),
    }


__all__ = ["kunye_uzlasmasi_olustur", "metinden_kunye_adayi"]
