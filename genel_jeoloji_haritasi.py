"""Parsel merkezli 1/100.000 genel jeoloji haritası ve çevre birimi analizi."""

from __future__ import annotations

import hashlib
import math
import os
from pathlib import Path
import textwrap

import cv2
import numpy as np
from PIL import Image, ImageDraw, ImageFont

from formasyon_metin_kutuphanesi import (
    birimleri_yasli_gence_sirala,
    jeolojik_yas_tahmin_et,
)
from jeoloji_kutuphanesi import jeoloji_anahtari
from jeoloji_pafta_kutuphanesi import (
    JeolojiPaftaHatasi,
    kmz_goruntusunu_ac,
)
from jeoloji_pafta_tanima import (
    _gorsel_ozelligi,
    _lejant_siniflari,
    _ozellik_benzerligi,
)


HARITA_GENISLIK_KM = 15.0
HARITA_YUKSEKLIK_KM = 9.0
HARITA_PIKSEL = (1772, 1063)  # 15 cm genişlikte yaklaşık 300 dpi


def parsel_geometri_hashi(points):
    normalized = ";".join(f"{float(lat):.9f},{float(lon):.9f}" for lat, lon in points or ())
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()


def _font(size, bold=False):
    candidates = (
        r"C:\Windows\Fonts\arialbd.ttf" if bold else r"C:\Windows\Fonts\arial.ttf",
        r"C:\Windows\Fonts\calibrib.ttf" if bold else r"C:\Windows\Fonts\calibri.ttf",
        "DejaVuSans-Bold.ttf" if bold else "DejaVuSans.ttf",
    )
    for path in candidates:
        try:
            return ImageFont.truetype(path, size=size)
        except OSError:
            continue
    return ImageFont.load_default()


def _merkez_ve_sinir(points, width_km, height_km):
    cleaned = [(float(lat), float(lon)) for lat, lon in points or ()]
    if len(cleaned) > 1 and cleaned[0] == cleaned[-1]:
        cleaned.pop()
    if len(cleaned) < 3:
        raise JeolojiPaftaHatasi("Genel jeoloji haritası için en az üç parsel noktası gerekir.")
    center_lat = sum(point[0] for point in cleaned) / len(cleaned)
    center_lon = sum(point[1] for point in cleaned) / len(cleaned)
    half_lat = (height_km / 2.0) / 111.32
    longitude_km = 111.32 * max(math.cos(math.radians(center_lat)), 0.05)
    half_lon = (width_km / 2.0) / longitude_km
    extent = {
        "north": center_lat + half_lat,
        "south": center_lat - half_lat,
        "east": center_lon + half_lon,
        "west": center_lon - half_lon,
    }
    return cleaned, (center_lat, center_lon), extent


def _record_haritalama_dizileri(record, extent, width, height):
    bounds = record.get("bounds", {})
    north = float(bounds["north"])
    south = float(bounds["south"])
    east = float(bounds["east"])
    west = float(bounds["west"])
    center_lat = (north + south) / 2.0
    center_lon = (east + west) / 2.0
    scale = math.cos(math.radians(center_lat))
    east_span = (east - west) * scale
    north_span = north - south
    lon_values = np.linspace(extent["west"], extent["east"], width, dtype=np.float32)
    lat_values = np.linspace(extent["north"], extent["south"], height, dtype=np.float32)
    lon_grid, lat_grid = np.meshgrid(lon_values, lat_values)
    east_offset = (lon_grid - center_lon) * scale
    north_offset = lat_grid - center_lat
    angle = math.radians(float(record.get("rotation", 0) or 0))
    cosine = math.cos(angle)
    sine = math.sin(angle)
    source_east = cosine * east_offset + sine * north_offset
    source_north = -sine * east_offset + cosine * north_offset
    x_normal = 0.5 + source_east / east_span
    y_normal = 0.5 - source_north / north_span
    valid = (x_normal >= 0) & (x_normal <= 1) & (y_normal >= 0) & (y_normal <= 1)
    return x_normal, y_normal, valid


def _pafta_mozaiği(library, records, extent, size):
    width, height = size
    canvas = np.full((height, width, 3), 255, dtype=np.uint8)
    source_index = np.full((height, width), -1, dtype=np.int16)
    loaded = []
    try:
        for index, record in enumerate(records):
            image = kmz_goruntusunu_ac(record)
            rgb = np.asarray(image.convert("RGB"), dtype=np.uint8)
            image.close()
            profile = library.lejant_getir(record.get("lejant_id"))
            classes = _lejant_siniflari(profile) if profile and profile.get("ogeler") else []
            x_normal, y_normal, valid = _record_haritalama_dizileri(record, extent, width, height)
            map_x = (x_normal * (rgb.shape[1] - 1)).astype(np.float32)
            map_y = (y_normal * (rgb.shape[0] - 1)).astype(np.float32)
            remapped = cv2.remap(
                rgb,
                map_x,
                map_y,
                interpolation=cv2.INTER_LINEAR,
                borderMode=cv2.BORDER_CONSTANT,
                borderValue=(255, 255, 255),
            )
            writable = valid & (source_index < 0)
            canvas[writable] = remapped[writable]
            source_index[writable] = index
            loaded.append(
                {
                    "record": dict(record),
                    "rgb": rgb,
                    "classes": classes,
                    "x_normal": x_normal,
                    "y_normal": y_normal,
                    "valid": valid,
                }
            )
    except Exception:
        for source in loaded:
            for entry in source.get("classes", []):
                preview = entry.get("preview")
                if preview is not None:
                    preview.close()
        raise
    if not loaded or not bool((source_index >= 0).any()):
        raise JeolojiPaftaHatasi("Harita alanını örten kullanılabilir 1/100.000 KMZ görüntüsü bulunamadı.")
    return Image.fromarray(canvas, mode="RGB"), source_index, loaded


def _birimleri_tani(source_index, loaded, ana_birim=None, satir=11, sutun=17):
    height, width = source_index.shape
    merged = {}
    total_used = 0
    for row in range(satir):
        y = max(0, min(height - 1, round((row + 0.5) * height / satir)))
        for column in range(sutun):
            x = max(0, min(width - 1, round((column + 0.5) * width / sutun)))
            source_no = int(source_index[y, x])
            if source_no < 0 or source_no >= len(loaded):
                continue
            source = loaded[source_no]
            classes = source["classes"]
            if not classes:
                continue
            source_x = float(source["x_normal"][y, x]) * (source["rgb"].shape[1] - 1)
            source_y = float(source["y_normal"][y, x]) * (source["rgb"].shape[0] - 1)
            radius = max(5, min(22, round(min(source["rgb"].shape[:2]) / 480)))
            left = max(0, int(source_x) - radius)
            top = max(0, int(source_y) - radius)
            right = min(source["rgb"].shape[1], int(source_x) + radius + 1)
            bottom = min(source["rgb"].shape[0], int(source_y) + radius + 1)
            if right - left < 3 or bottom - top < 3:
                continue
            patch = Image.fromarray(source["rgb"][top:bottom, left:right], mode="RGB")
            feature = _gorsel_ozelligi(patch)
            patch.close()
            scores = [
                max(_ozellik_benzerligi(feature, reference) for reference in entry["features"])
                for entry in classes
            ]
            order = sorted(range(len(scores)), key=lambda i: scores[i], reverse=True)
            best_index = order[0]
            best = scores[best_index]
            second = scores[order[1]] if len(order) > 1 else 0.0
            if best < 0.34 or best - second < 0.004:
                continue
            entry = classes[best_index]
            key = (jeoloji_anahtari(entry.get("kod")), jeoloji_anahtari(entry.get("ad")))
            unit = merged.setdefault(
                key,
                {
                    "kod": entry.get("kod", ""),
                    "ad": entry.get("ad", ""),
                    "adet": 0,
                    "puan_toplami": 0.0,
                    "pafta_idleri": set(),
                    "pafta_adlari": set(),
                    "kaynak_sirasi": max(entry.get("kaynak_siralari") or [-1]),
                    "onizleme": entry.get("preview").copy() if entry.get("preview") is not None else None,
                },
            )
            unit["adet"] += 1
            unit["puan_toplami"] += best
            unit["pafta_idleri"].add(source["record"].get("id", ""))
            unit["pafta_adlari"].add(source["record"].get("ad", ""))
            total_used += 1

    ana_kod = jeoloji_anahtari((ana_birim or {}).get("birim_kodu"))
    ana_ad = jeoloji_anahtari((ana_birim or {}).get("birim_adi"))
    for source in loaded:
        for entry in source.get("classes", []):
            entry_kod = jeoloji_anahtari(entry.get("kod"))
            entry_ad = jeoloji_anahtari(entry.get("ad"))
            if not ((ana_kod and entry_kod == ana_kod) or (ana_ad and entry_ad == ana_ad)):
                continue
            key = (entry_kod, entry_ad)
            if key not in merged:
                merged[key] = {
                    "kod": entry.get("kod", ""),
                    "ad": entry.get("ad", ""),
                    "adet": 1,
                    "puan_toplami": 1.0,
                    "pafta_idleri": {source["record"].get("id", "")},
                    "pafta_adlari": {source["record"].get("ad", "")},
                    "kaynak_sirasi": max(entry.get("kaynak_siralari") or [-1]),
                    "onizleme": entry.get("preview").copy() if entry.get("preview") is not None else None,
                }
            merged[key]["ana_birim"] = True

    units = []
    for unit in merged.values():
        share = 100.0 * unit["adet"] / max(total_used, 1)
        if unit["adet"] < 2 and share < 1.5 and not unit.get("ana_birim"):
            preview = unit.get("onizleme")
            if preview is not None:
                preview.close()
            continue
        yas, yas_sirasi = jeolojik_yas_tahmin_et(unit["kod"], unit["ad"])
        units.append(
            {
                "kod": unit["kod"],
                "ad": unit["ad"],
                "jeolojik_yas": yas,
                "yas_sirasi": yas_sirasi,
                "kaynak_sirasi": unit["kaynak_sirasi"],
                "oran": round(share, 1),
                "guven": round(100.0 * unit["puan_toplami"] / max(unit["adet"], 1), 1),
                "pafta_idleri": sorted(value for value in unit["pafta_idleri"] if value),
                "pafta_adlari": sorted(value for value in unit["pafta_adlari"] if value),
                "ana_birim": bool(unit.get("ana_birim")),
                "onizleme": unit.get("onizleme"),
            }
        )
    units.sort(key=lambda unit: (-int(unit.get("ana_birim", False)), -unit["oran"], -unit["guven"]))
    kept = units[:12]
    for unit in units[12:]:
        preview = unit.get("onizleme")
        if preview is not None:
            preview.close()
    return kept


def _lejant_katalogu(loaded, ana_birim=None):
    """Kapsayan paftalardaki bütün tanımlı lejant birimlerini birleştirir.

    Yerel örnekleme yalnız en güçlü adayları döndürür. Kod okuyan bir yapay zekâ
    denetiminin bu ön elemeye bağımlı kalmaması için tüm lejant kataloğu ayrıca
    analiz sonucunda tutulur.
    """
    merged = {}
    main_code = jeoloji_anahtari((ana_birim or {}).get("birim_kodu"))
    main_name = jeoloji_anahtari((ana_birim or {}).get("birim_adi"))
    for source in loaded:
        record = source.get("record", {})
        for entry in source.get("classes", []):
            code_key = jeoloji_anahtari(entry.get("kod"))
            name_key = jeoloji_anahtari(entry.get("ad"))
            key = code_key or name_key
            if not key:
                continue
            unit = merged.get(key)
            if unit is None:
                age, age_order = jeolojik_yas_tahmin_et(entry.get("kod"), entry.get("ad"))
                unit = {
                    "kod": entry.get("kod", ""),
                    "ad": entry.get("ad", ""),
                    "jeolojik_yas": age,
                    "yas_sirasi": age_order,
                    "kaynak_sirasi": max(entry.get("kaynak_siralari") or [-1]),
                    "pafta_idleri": set(),
                    "pafta_adlari": set(),
                    "ana_birim": False,
                    "onizleme": (
                        entry.get("preview").copy()
                        if entry.get("preview") is not None else None
                    ),
                }
                merged[key] = unit
            unit["pafta_idleri"].add(record.get("id", ""))
            unit["pafta_adlari"].add(record.get("ad", ""))
            unit["kaynak_sirasi"] = max(
                int(unit.get("kaynak_sirasi", -1)),
                max(entry.get("kaynak_siralari") or [-1]),
            )
            if (main_code and code_key == main_code) or (main_name and name_key == main_name):
                unit["ana_birim"] = True
    result = []
    for unit in merged.values():
        unit["pafta_idleri"] = sorted(value for value in unit["pafta_idleri"] if value)
        unit["pafta_adlari"] = sorted(value for value in unit["pafta_adlari"] if value)
        result.append(unit)
    result.sort(
        key=lambda unit: (
            int(unit.get("yas_sirasi", 9999)),
            int(unit.get("kaynak_sirasi", -1)),
            jeoloji_anahtari(unit.get("kod")),
        )
    )
    return result


def _geo_to_pixel(lat, lon, extent, width, height):
    x = (float(lon) - extent["west"]) / (extent["east"] - extent["west"]) * (width - 1)
    y = (extent["north"] - float(lat)) / (extent["north"] - extent["south"]) * (height - 1)
    return round(x), round(y)


def _haritayi_isaretle(image, polygon, extent, width_km):
    result = image.copy().convert("RGB")
    draw = ImageDraw.Draw(result)
    points = [_geo_to_pixel(lat, lon, extent, result.width, result.height) for lat, lon in polygon]
    if points and points[0] != points[-1]:
        points.append(points[0])
    draw.line(points, fill="white", width=12, joint="curve")
    draw.line(points, fill="#c00020", width=6, joint="curve")
    if points:
        cx = round(sum(point[0] for point in points[:-1]) / max(len(points) - 1, 1))
        cy = round(sum(point[1] for point in points[:-1]) / max(len(points) - 1, 1))
        label = "ÇALIŞMA ALANI"
        font = _font(28, bold=True)
        bbox = draw.textbbox((0, 0), label, font=font)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        tx = max(10, min(result.width - tw - 20, cx - tw // 2))
        ty = max(10, min(result.height - th - 20, cy - th - 18))
        draw.rectangle((tx - 7, ty - 5, tx + tw + 7, ty + th + 5), fill="white", outline="#c00020", width=2)
        draw.text((tx, ty), label, fill="#a00018", font=font)

    # Kuzey oku ve 2 km ölçek çubuğu.
    arrow_x = result.width - 68
    draw.line((arrow_x, 95, arrow_x, 28), fill="black", width=5)
    draw.polygon(((arrow_x, 18), (arrow_x - 12, 42), (arrow_x + 12, 42)), fill="black")
    draw.text((arrow_x - 11, 98), "K", fill="black", font=_font(26, bold=True))
    bar_km = 2.0
    bar_width = round(result.width * bar_km / width_km)
    bar_x = 42
    bar_y = result.height - 42
    draw.line((bar_x, bar_y, bar_x + bar_width, bar_y), fill="white", width=10)
    draw.line((bar_x, bar_y, bar_x + bar_width, bar_y), fill="black", width=4)
    for offset in (0, bar_width // 2, bar_width):
        draw.line((bar_x + offset, bar_y - 10, bar_x + offset, bar_y + 10), fill="black", width=3)
    draw.text((bar_x, bar_y - 39), "0", fill="black", font=_font(23, bold=True))
    draw.text((bar_x + bar_width // 2 - 8, bar_y - 39), "1", fill="black", font=_font(23, bold=True))
    draw.text((bar_x + bar_width - 8, bar_y - 39), "2 km", fill="black", font=_font(23, bold=True))
    draw.rectangle((2, 2, result.width - 3, result.height - 3), outline="black", width=4)
    return result


def genel_jeoloji_verisini_hazirla(
    library,
    points,
    ana_birim=None,
    width_km=HARITA_GENISLIK_KM,
    height_km=HARITA_YUKSEKLIK_KM,
):
    polygon, center, extent = _merkez_ve_sinir(points, width_km, height_km)
    corners = [
        (extent["north"], extent["west"]),
        (extent["north"], extent["east"]),
        (extent["south"], extent["east"]),
        (extent["south"], extent["west"]),
        center,
    ]
    records = library.kapsayan_paftalar(corners)
    ready = []
    for record in records:
        profile = library.lejant_getir(record.get("lejant_id"))
        if record.get("kmz_path") and profile and profile.get("ogeler"):
            ready.append(record)
    if not ready:
        raise JeolojiPaftaHatasi(
            "Parsel çevresini kapsayan, KMZ görüntüsü ve hazırlanmış lejantı bulunan 1/100.000 pafta yok."
        )
    base, source_index, loaded = _pafta_mozaiği(library, ready, extent, HARITA_PIKSEL)
    units = []
    catalog = []
    marked = None
    try:
        catalog = _lejant_katalogu(loaded, ana_birim=ana_birim)
        units = _birimleri_tani(source_index, loaded, ana_birim=ana_birim)
        marked = _haritayi_isaretle(base, polygon, extent, width_km)
    except Exception:
        for collection in (units, catalog):
            for unit in collection:
                preview = unit.get("onizleme")
                if preview is not None:
                    preview.close()
        if marked is not None:
            marked.close()
        raise
    finally:
        base.close()
        for source in loaded:
            for entry in source.get("classes", []):
                preview = entry.get("preview")
                if preview is not None:
                    preview.close()
    if not units:
        marked.close()
        for unit in catalog:
            preview = unit.get("onizleme")
            if preview is not None:
                preview.close()
        raise JeolojiPaftaHatasi("Harita çevresinde güvenilir formasyon adayı belirlenemedi.")
    return {
        "harita": marked,
        "birimler": units,
        "lejant_birimleri": catalog,
        "merkez": [round(center[0], 8), round(center[1], 8)],
        "sinir": {key: round(value, 8) for key, value in extent.items()},
        "genislik_km": width_km,
        "yukseklik_km": height_km,
        "pafta_idleri": [record.get("id", "") for record in ready],
        "pafta_adlari": [record.get("ad", "") for record in ready],
        "geometri_hash": parsel_geometri_hashi(points),
    }


def _wrapped_lines(draw, text, font, max_width):
    words = str(text or "").split()
    if not words:
        return [""]
    lines = []
    current = words[0]
    for word in words[1:]:
        candidate = f"{current} {word}"
        if draw.textlength(candidate, font=font) <= max_width:
            current = candidate
        else:
            lines.append(current)
            current = word
    lines.append(current)
    return lines


def _lejant_satirlari(draw, text, font, max_width, maksimum=3):
    lines = _wrapped_lines(draw, text, font, max_width)
    if len(lines) <= maksimum:
        return lines
    lines = lines[:maksimum]
    last = lines[-1].rstrip(" .")
    while last and draw.textlength(last + "…", font=font) > max_width:
        last = last[:-1].rstrip()
    lines[-1] = (last or "") + "…"
    return lines


def genel_jeoloji_gorselini_olustur(harita, birimler, hedef_yolu):
    """Harita panelini, yaşlıdan gence gerçek lejant örnekleriyle tek JPG yapar."""
    ordered = birimleri_yasli_gence_sirala(birimler)
    width = HARITA_PIKSEL[0]
    margin = 44
    title_height = 108
    row_min_height = 104
    description_font = _font(29)
    title_font = _font(43, bold=True)
    unit_font = _font(31, bold=True)
    age_font = _font(26, bold=True)

    dummy = Image.new("RGB", (width, 100), "white")
    dummy_draw = ImageDraw.Draw(dummy)
    row_heights = []
    for unit in ordered:
        description = unit.get("lejant_aciklamasi") or unit.get("ad") or ""
        lines = _lejant_satirlari(dummy_draw, description, description_font, width - 390)
        row_heights.append(max(row_min_height, 27 + len(lines) * 36))
    dummy.close()
    footer_height = 92
    total_height = harita.height + title_height + sum(row_heights) + footer_height + margin
    canvas = Image.new("RGB", (width, total_height), "white")
    canvas.paste(harita.resize((width, HARITA_PIKSEL[1]), Image.Resampling.LANCZOS), (0, 0))
    draw = ImageDraw.Draw(canvas)
    y = harita.height
    draw.line((margin, y + 4, width - margin, y + 4), fill="#202020", width=4)
    draw.text((margin, y + 29), "AÇIKLAMALAR", fill="black", font=title_font)
    draw.text((width - 520, y + 39), "Yaşlıdan gence doğru", fill="#424242", font=age_font)
    y += title_height
    for unit, row_height in zip(ordered, row_heights):
        preview = unit.get("onizleme")
        sample_box = (margin, y + 15, margin + 190, y + row_height - 15)
        draw.rectangle(sample_box, fill="white", outline="#333333", width=2)
        if preview is not None:
            sample = preview.convert("RGB")
            sample.thumbnail((sample_box[2] - sample_box[0] - 8, sample_box[3] - sample_box[1] - 8), Image.Resampling.LANCZOS)
            canvas.paste(
                sample,
                (
                    sample_box[0] + (sample_box[2] - sample_box[0] - sample.width) // 2,
                    sample_box[1] + (sample_box[3] - sample_box[1] - sample.height) // 2,
                ),
            )
            sample.close()
        code = str(unit.get("kod") or "").strip()
        name = str(unit.get("ad") or "").strip()
        age = str(unit.get("jeolojik_yas") or "Yaşı belirtilmedi").strip()
        heading = f"{code}  {name}" if code else name
        text_x = margin + 220
        draw.text((text_x, y + 11), heading, fill="black", font=unit_font)
        age_width = draw.textlength(age, font=age_font)
        draw.text((width - margin - age_width, y + 15), age, fill="#7a1f2b", font=age_font)
        description = unit.get("lejant_aciklamasi") or name
        lines = _lejant_satirlari(draw, description, description_font, width - text_x - margin)
        for line_no, line in enumerate(lines):
            draw.text((text_x, y + 52 + line_no * 36), line, fill="#202020", font=description_font)
        y += row_height
        draw.line((margin, y, width - margin, y), fill="#d2d2d2", width=2)

    footer = "Çalışma alanı merkez alınarak hazırlanmıştır. Harita paneli baskıda 15 cm genişlikte 1/100.000 ölçektedir."
    footer_lines = textwrap.wrap(footer, width=115)
    for index, line in enumerate(footer_lines):
        draw.text((margin, y + 23 + index * 30), line, fill="#424242", font=_font(23))
    draw.rectangle((3, 3, width - 4, total_height - 4), outline="black", width=5)
    target = Path(hedef_yolu)
    target.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(target, format="JPEG", quality=95, subsampling=0, optimize=True)
    canvas.close()
    return str(target.resolve())


def birim_serilestir(birim):
    """PIL nesnelerini proje JSON'una taşımadan birim anlık görüntüsü üret."""
    allowed = (
        "kod", "ad", "jeolojik_yas", "yas_sirasi", "kaynak_sirasi", "oran", "guven",
        "pafta_idleri", "pafta_adlari", "ana_birim", "lejant_aciklamasi",
        "bolgesel_jeoloji_metni", "kutuphane_id", "kutuphane_revizyon_no", "metin_kaynagi",
        "yerel_aday",
        "ai_gemini_guven", "ai_openai_guven", "ai_birlesik_guven", "ai_kanit",
        "ai_gemini_kanit", "ai_openai_kanit",
        "ai_gemini_aciklama", "ai_openai_aciklama",
        "ai_durum", "ai_oneri", "ai_aciklama", "ai_saglayicilar",
    )
    return {key: birim.get(key) for key in allowed if key in birim}


def analiz_gorsellerini_kapat(analiz):
    try:
        harita = (analiz or {}).get("harita")
        if harita is not None:
            harita.close()
    except Exception:
        pass
    for collection_name in ("birimler", "lejant_birimleri"):
        for unit in (analiz or {}).get(collection_name, []):
            try:
                preview = unit.get("onizleme")
                if preview is not None:
                    preview.close()
            except Exception:
                pass


__all__ = [
    "HARITA_GENISLIK_KM",
    "HARITA_YUKSEKLIK_KM",
    "analiz_gorsellerini_kapat",
    "birim_serilestir",
    "genel_jeoloji_gorselini_olustur",
    "genel_jeoloji_verisini_hazirla",
    "parsel_geometri_hashi",
]
