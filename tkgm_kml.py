from __future__ import annotations

import json
import math
import os
import re
import tempfile
import unicodedata
import urllib.error
import urllib.request
from xml.sax.saxutils import escape


TKGM_API_BASE = "https://cbsapi.tkgm.gov.tr/megsiswebapi.v3.1/api/"
TKGM_IL_LISTE_URL = (
    "https://parselsorgu.tkgm.gov.tr/app/modules/administrativeQuery/data/ilListe.json"
)
TKGM_AZAMI_YANIT_BOYUTU = 5 * 1024 * 1024
TKGM_AZAMI_NOKTA_SAYISI = 20_000


class TKGMSorguHatasi(RuntimeError):
    """TKGM parsel sorgusu kullanıcıya gösterilebilir bir hata ile başarısız oldu."""


def konum_adi_normalize_et(value):
    text = str(value or "").strip().casefold().replace("ı", "i")
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(
        r"\b(mahallesi|mahalle|mah|mh|koyu|koy|beldesi|belde)\b",
        " ",
        text,
    )
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def ada_parsel_no_temizle(value, default="0"):
    text = str(value or "").strip()
    if not text:
        return default
    match = re.search(r"\d+", text)
    if not match:
        raise TKGMSorguHatasi(f"Ada/parsel değeri sayısal değil: {text}")
    return match.group(0)


def tkgm_kml_dosya_adi(kunye):
    ada = ada_parsel_no_temizle(kunye.get("ada"), default="0")
    parsel = ada_parsel_no_temizle(kunye.get("parsel"), default="")
    if not parsel:
        raise TKGMSorguHatasi("Parsel bilgisi boş.")
    return f"TKGM_{ada}_{parsel}.kml"


def _json_getir(url, timeout=25):
    request = urllib.request.Request(
        url,
        headers={
            "Accept": "application/json,text/plain,*/*",
            "Accept-Language": "tr-TR,tr;q=0.9,en;q=0.7",
            "User-Agent": "K-1/1.0 (TKGM KML)",
            "Referer": "https://parselsorgu.tkgm.gov.tr/",
        },
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            content_type = (response.headers.get_content_type() or "").casefold()
            raw = response.read(TKGM_AZAMI_YANIT_BOYUTU + 1)
            charset = response.headers.get_content_charset() or "utf-8"
    except urllib.error.HTTPError as exc:
        try:
            detail = exc.read(300).decode("utf-8", errors="replace").strip()
        except Exception:
            detail = ""
        suffix = f" - {detail}" if detail else ""
        raise TKGMSorguHatasi(f"TKGM servisi {exc.code} hatası verdi{suffix}.") from exc
    except urllib.error.URLError as exc:
        raise TKGMSorguHatasi(f"TKGM servisine ulaşılamadı: {exc.reason}") from exc
    except TimeoutError as exc:
        raise TKGMSorguHatasi("TKGM servisi zaman aşımına uğradı.") from exc

    if len(raw) > TKGM_AZAMI_YANIT_BOYUTU:
        raise TKGMSorguHatasi("TKGM servis yanıtı güvenli boyut sınırını aşıyor.")
    if "html" in content_type:
        raise TKGMSorguHatasi("TKGM servisi JSON yerine bir web sayfası döndürdü.")
    try:
        return json.loads(raw.decode(charset, errors="strict"))
    except (LookupError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise TKGMSorguHatasi("TKGM servisi beklenen JSON yanıtını döndürmedi.") from exc


def _liste_elemanlari(payload):
    if isinstance(payload, list):
        source = payload
    elif isinstance(payload, dict):
        source = (
            payload.get("features")
            or payload.get("data")
            or payload.get("result")
            or payload.get("items")
            or []
        )
    else:
        source = []

    items = []
    for item in source:
        if not isinstance(item, dict):
            continue
        properties = item.get("properties")
        items.append(
            {
                "raw": item,
                "properties": properties if isinstance(properties, dict) else item,
            }
        )
    return items


def _etiket_al(item):
    props = item.get("properties", {}) if isinstance(item, dict) else {}
    for key in (
        "text",
        "ad",
        "adi",
        "name",
        "label",
        "ilAdi",
        "ilceAdi",
        "mahalleAdi",
        "value",
    ):
        value = props.get(key)
        if value not in (None, ""):
            return str(value)
    return ""


def _id_al(item):
    props = item.get("properties", {}) if isinstance(item, dict) else {}
    for key in ("id", "kod", "value", "mahalleId", "ilceId", "ilId"):
        value = props.get(key)
        if value not in (None, ""):
            return value
    raw = item.get("raw", {}) if isinstance(item, dict) else {}
    for key in ("id", "kod"):
        value = raw.get(key)
        if value not in (None, ""):
            return value
    return None


def _konum_bul(payload, aranan, alan_adi):
    hedef = konum_adi_normalize_et(aranan)
    if not hedef:
        raise TKGMSorguHatasi(f"{alan_adi} bilgisi boş.")

    items = _liste_elemanlari(payload)
    if not items:
        raise TKGMSorguHatasi(f"TKGM {alan_adi} listesi boş geldi.")

    exact = []
    contains = []
    for item in items:
        normalized = konum_adi_normalize_et(_etiket_al(item))
        if normalized == hedef:
            exact.append(item)
        elif normalized and (hedef in normalized or normalized in hedef):
            contains.append(item)

    candidates = exact or contains
    if len(candidates) > 1:
        names = ", ".join(_etiket_al(item) for item in candidates[:8])
        raise TKGMSorguHatasi(
            f"{alan_adi} eşleşmesi belirsiz: {aranan}. Eşleşen kayıtlar: {names}"
        )
    if not candidates:
        examples = ", ".join(_etiket_al(item) for item in items[:8] if _etiket_al(item))
        suffix = f" TKGM listesinden örnekler: {examples}" if examples else ""
        raise TKGMSorguHatasi(f"{alan_adi} bulunamadı: {aranan}.{suffix}")

    match = candidates[0]
    match_id = _id_al(match)
    if match_id in (None, ""):
        raise TKGMSorguHatasi(
            f"{alan_adi} için TKGM kimliği bulunamadı: {_etiket_al(match)}"
        )
    return match_id, _etiket_al(match)


def _geometry_ve_properties(payload):
    if isinstance(payload, dict):
        if payload.get("type") == "Feature":
            return payload.get("geometry"), payload.get("properties", {})
        if isinstance(payload.get("geometry"), dict):
            return payload.get("geometry"), payload.get("properties", {})
        features = payload.get("features")
        if isinstance(features, list):
            for feature in features:
                geometry, properties = _geometry_ve_properties(feature)
                if geometry:
                    return geometry, properties
        for key in ("data", "result", "entity"):
            geometry, properties = _geometry_ve_properties(payload.get(key))
            if geometry:
                return geometry, properties
    elif isinstance(payload, list):
        for item in payload:
            geometry, properties = _geometry_ve_properties(item)
            if geometry:
                return geometry, properties
    return None, {}


def _polygonlar(geometry):
    if not isinstance(geometry, dict):
        return []
    geometry_type = str(geometry.get("type") or "").casefold()
    coordinates = geometry.get("coordinates")
    if geometry_type == "polygon" and isinstance(coordinates, list):
        return [coordinates]
    if geometry_type == "multipolygon" and isinstance(coordinates, list):
        return coordinates
    return []


def _ring_dogrula(ring, kalan_nokta):
    if not isinstance(ring, list):
        raise TKGMSorguHatasi("TKGM parsel koordinat halkası geçersiz.")
    points = []
    for point in ring:
        if not isinstance(point, (list, tuple)) or len(point) < 2:
            raise TKGMSorguHatasi("TKGM parsel geometrisinde eksik koordinat var.")
        try:
            lon = float(point[0])
            lat = float(point[1])
            alt = float(point[2]) if len(point) > 2 else 0.0
        except (TypeError, ValueError) as exc:
            raise TKGMSorguHatasi("TKGM parsel geometrisinde sayısal olmayan koordinat var.") from exc
        if not all(math.isfinite(value) for value in (lon, lat, alt)):
            raise TKGMSorguHatasi("TKGM parsel geometrisinde sonlu olmayan koordinat var.")
        if not -180.0 <= lon <= 180.0 or not -90.0 <= lat <= 90.0:
            raise TKGMSorguHatasi("TKGM parsel koordinatı WGS84 sınırlarının dışında.")
        points.append((lon, lat, alt))
        if len(points) > kalan_nokta:
            raise TKGMSorguHatasi("TKGM parsel geometrisi güvenli nokta sınırını aşıyor.")

    unique = {(round(lon, 12), round(lat, 12)) for lon, lat, _ in points}
    if len(unique) < 3:
        raise TKGMSorguHatasi("TKGM parsel poligonu en az üç farklı nokta içermelidir.")
    if points[0][:2] != points[-1][:2]:
        if len(points) >= kalan_nokta:
            raise TKGMSorguHatasi("TKGM parsel geometrisi güvenli nokta sınırını aşıyor.")
        points.append(points[0])
    return points


def _polygonlari_dogrula(geometry):
    raw_polygons = _polygonlar(geometry)
    if not raw_polygons:
        raise TKGMSorguHatasi("TKGM parsel geometrisi Polygon/MultiPolygon olarak gelmedi.")

    polygons = []
    point_count = 0
    for polygon in raw_polygons:
        if not isinstance(polygon, list) or not polygon:
            raise TKGMSorguHatasi("TKGM parsel poligonu boş veya geçersiz.")
        rings = []
        for ring in polygon:
            points = _ring_dogrula(ring, TKGM_AZAMI_NOKTA_SAYISI - point_count)
            point_count += len(points)
            rings.append(points)
        polygons.append(rings)
    return polygons


def _property_value(properties, *keys):
    if not isinstance(properties, dict):
        return ""
    for key in keys:
        value = properties.get(key)
        if value not in (None, ""):
            return str(value).strip()
    return ""


def _parsel_kunyesini_dogrula(properties, mahalle_id, ada, parsel):
    returned_ada = _property_value(properties, "adaNo", "ada")
    returned_parsel = _property_value(properties, "parselNo", "parsel")
    returned_mahalle_id = _property_value(properties, "mahalleId", "mahalle_id")

    if returned_ada and ada_parsel_no_temizle(returned_ada, default="") != ada:
        raise TKGMSorguHatasi(
            f"TKGM farklı ada döndürdü: istenen {ada}, dönen {returned_ada}."
        )
    if returned_parsel and ada_parsel_no_temizle(returned_parsel, default="") != parsel:
        raise TKGMSorguHatasi(
            f"TKGM farklı parsel döndürdü: istenen {parsel}, dönen {returned_parsel}."
        )
    if returned_mahalle_id and returned_mahalle_id != str(mahalle_id):
        raise TKGMSorguHatasi("TKGM geometrisi farklı bir mahalle/köy kaydına ait.")


def geojson_kml_olustur(geometry, name="TKGM Parsel", description=""):
    polygons = _polygonlari_dogrula(geometry)
    placemarks = []
    for index, polygon in enumerate(polygons, start=1):
        outer = " ".join(
            f"{lon:.8f},{lat:.8f},{alt:.2f}" for lon, lat, alt in polygon[0]
        )
        inner_parts = []
        for ring in polygon[1:]:
            inner = " ".join(
                f"{lon:.8f},{lat:.8f},{alt:.2f}" for lon, lat, alt in ring
            )
            inner_parts.append(
                "<innerBoundaryIs><LinearRing><coordinates>"
                + inner
                + "</coordinates></LinearRing></innerBoundaryIs>"
            )
        label = name if len(polygons) == 1 else f"{name} - {index}"
        placemarks.append(
            "<Placemark>"
            f"<name>{escape(label)}</name>"
            f"<description>{escape(description)}</description>"
            "<styleUrl>#parselStyle</styleUrl>"
            "<Polygon><tessellate>1</tessellate>"
            f"<outerBoundaryIs><LinearRing><coordinates>{outer}</coordinates>"
            "</LinearRing></outerBoundaryIs>"
            + "".join(inner_parts)
            + "</Polygon></Placemark>"
        )

    return (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<kml xmlns="http://www.opengis.net/kml/2.2">\n'
        "<Document>\n"
        f"<name>{escape(name)}</name>\n"
        '<Style id="parselStyle">'
        '<LineStyle><color>ff0000ff</color><width>2</width></LineStyle>'
        '<PolyStyle><color>2600ffff</color></PolyStyle>'
        "</Style>\n"
        + "\n".join(placemarks)
        + "\n</Document>\n</kml>\n"
    )


def _merkez_hesapla(polygons):
    points = [point for polygon in polygons for point in polygon[0]]
    if not points:
        return None
    lons = [point[0] for point in points]
    lats = [point[1] for point in points]
    return (min(lats) + max(lats)) / 2.0, (min(lons) + max(lons)) / 2.0


def _atomik_metin_yaz(path, text):
    directory = os.path.dirname(os.path.abspath(path))
    os.makedirs(directory, exist_ok=True)
    temp_path = ""
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            newline="\n",
            prefix=".tkgm_",
            suffix=".tmp",
            dir=directory,
            delete=False,
        ) as handle:
            handle.write(text)
            temp_path = handle.name
        os.replace(temp_path, path)
    finally:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass


def tkgm_parsel_kml_olustur(kunye, output_dir, timeout=25, fetcher=_json_getir):
    if not isinstance(kunye, dict):
        raise TKGMSorguHatasi("Proje bilgileri okunamadı.")

    il = str(kunye.get("il") or "").strip()
    ilce = str(kunye.get("ilce") or "").strip()
    mahalle = str(kunye.get("mahalle") or kunye.get("koy") or "").strip()
    ada = ada_parsel_no_temizle(kunye.get("ada"), default="0")
    parsel = ada_parsel_no_temizle(kunye.get("parsel"), default="")

    missing = []
    if not il:
        missing.append("İl")
    if not ilce:
        missing.append("İlçe")
    if not mahalle:
        missing.append("Mahalle/Köy")
    if not parsel:
        missing.append("Parsel")
    if missing:
        raise TKGMSorguHatasi("Eksik proje bilgileri: " + ", ".join(missing))

    il_id, il_label = _konum_bul(fetcher(TKGM_IL_LISTE_URL, timeout=timeout), il, "İl")
    ilce_id, ilce_label = _konum_bul(
        fetcher(f"{TKGM_API_BASE}idariYapi/ilceListe/{il_id}", timeout=timeout),
        ilce,
        "İlçe",
    )
    mahalle_id, mahalle_label = _konum_bul(
        fetcher(f"{TKGM_API_BASE}idariYapi/mahalleListe/{ilce_id}", timeout=timeout),
        mahalle,
        "Mahalle/Köy",
    )

    parsel_sorgu_url = f"{TKGM_API_BASE}parsel/{mahalle_id}/{ada}/{parsel}"
    payload = fetcher(parsel_sorgu_url, timeout=timeout)
    geometry, properties = _geometry_ve_properties(payload)
    if not geometry:
        raise TKGMSorguHatasi(f"TKGM parsel geometrisi bulunamadı: {ada}/{parsel}")

    _parsel_kunyesini_dogrula(properties, mahalle_id, ada, parsel)
    polygons = _polygonlari_dogrula(geometry)
    label = f"{il_label} {ilce_label} {mahalle_label} {ada}/{parsel}"
    description = "TKGM Parsel Sorgu servisinden K-1 ile oluşturuldu."
    kml_text = geojson_kml_olustur(geometry, name=label, description=description)

    output_path = os.path.join(output_dir, tkgm_kml_dosya_adi({"ada": ada, "parsel": parsel}))
    _atomik_metin_yaz(output_path, kml_text)
    return {
        "path": output_path,
        "center": _merkez_hesapla(polygons),
        "label": label,
        "properties": properties,
        "il": il_label,
        "ilce": ilce_label,
        "mahalle": mahalle_label,
        "ada": ada,
        "parsel": parsel,
        "mahalle_id": mahalle_id,
        "polygon_count": len(polygons),
        "source_url": parsel_sorgu_url,
    }


__all__ = [
    "TKGM_API_BASE",
    "TKGMSorguHatasi",
    "ada_parsel_no_temizle",
    "geojson_kml_olustur",
    "konum_adi_normalize_et",
    "tkgm_kml_dosya_adi",
    "tkgm_parsel_kml_olustur",
]
