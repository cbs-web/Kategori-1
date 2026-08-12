"""Parsel geometrisini 1/100.000 jeoloji paftasındaki lejant örnekleriyle karşılaştırır."""

from __future__ import annotations

import math
import os
from pathlib import Path

from PIL import Image, ImageDraw

from jeoloji_pafta_kutuphanesi import (
    GORSEL_AZAMI_PIKSEL,
    JeolojiPaftaHatasi,
    kmz_goruntusunu_ac,
    pafta_anahtari,
    pafta_gorsel_koordinati,
)


def _numpy_al():
    try:
        import numpy as np
    except Exception as exc:
        raise JeolojiPaftaHatasi(
            "Jeoloji paftası tanıması için numpy paketi kurulmalıdır."
        ) from exc
    return np


def _histogram(np, values, bins, value_range):
    histogram = np.histogram(values, bins=bins, range=value_range)[0].astype("float32")
    total = float(histogram.sum())
    if total:
        histogram /= total
    return histogram


def _gorsel_ozelligi(image):
    """Renk ve tarama dokusunu, farklı çözünürlüklere dayanıklı bir imzaya çevirir."""
    np = _numpy_al()
    rgb = image.convert("RGB")
    if max(rgb.size) > 180:
        oran = 180.0 / max(rgb.size)
        rgb = rgb.resize(
            (max(1, round(rgb.width * oran)), max(1, round(rgb.height * oran))),
            Image.Resampling.LANCZOS,
        )
    pixels = np.asarray(rgb, dtype="float32").reshape(-1, 3) / 255.0
    if not len(pixels):
        raise JeolojiPaftaHatasi("Boş bir lejant örneği kullanılamaz.")

    # Lejant kutusunun beyaz kenarı ve siyah çerçevesi sınıflandırmayı bozmasın.
    kullan = ~((pixels.min(axis=1) > 0.965) | (pixels.max(axis=1) < 0.035))
    if int(kullan.sum()) >= max(24, len(pixels) // 12):
        pixels = pixels[kullan]

    quantized = np.clip((pixels * 4).astype("int16"), 0, 3)
    rgb_bins = quantized[:, 0] * 16 + quantized[:, 1] * 4 + quantized[:, 2]
    rgb_hist = np.bincount(rgb_bins, minlength=64).astype("float32")
    rgb_hist /= max(float(rgb_hist.sum()), 1.0)

    gray_values = pixels @ np.asarray((0.299, 0.587, 0.114), dtype="float32")
    gray_hist = _histogram(np, gray_values, 16, (0.0, 1.0))

    gray = np.asarray(
        rgb.convert("L").resize((96, 96), Image.Resampling.BILINEAR),
        dtype="float32",
    ) / 255.0
    gx = np.zeros_like(gray)
    gy = np.zeros_like(gray)
    gx[:, 1:-1] = (gray[:, 2:] - gray[:, :-2]) / 2.0
    gy[1:-1, :] = (gray[2:, :] - gray[:-2, :]) / 2.0
    magnitude = np.sqrt(gx * gx + gy * gy)
    grad_hist = _histogram(np, np.clip(magnitude, 0.0, 0.55), 10, (0.0, 0.55))
    aci = np.mod(np.arctan2(gy, gx), math.pi)
    orient_index = np.clip((aci / math.pi * 8).astype("int16"), 0, 7)
    orient_hist = np.bincount(
        orient_index.reshape(-1),
        weights=magnitude.reshape(-1),
        minlength=8,
    ).astype("float32")
    orient_hist /= max(float(orient_hist.sum()), 1.0e-8)

    return {
        "rgb_hist": rgb_hist,
        "gray_hist": gray_hist,
        "grad_hist": grad_hist,
        "orient_hist": orient_hist,
        "mean": pixels.mean(axis=0),
        "std": pixels.std(axis=0),
        "edge": float((magnitude > 0.075).mean()),
    }


def _hist_benzerligi(np, first, second):
    return float(np.sqrt(np.clip(first, 0, None) * np.clip(second, 0, None)).sum())


def _ozellik_benzerligi(first, second):
    np = _numpy_al()
    rgb = _hist_benzerligi(np, first["rgb_hist"], second["rgb_hist"])
    gray = _hist_benzerligi(np, first["gray_hist"], second["gray_hist"])
    grad = _hist_benzerligi(np, first["grad_hist"], second["grad_hist"])
    orient = _hist_benzerligi(np, first["orient_hist"], second["orient_hist"])
    mean = math.exp(-float(np.linalg.norm(first["mean"] - second["mean"])) / 0.30)
    std = math.exp(-float(np.linalg.norm(first["std"] - second["std"])) / 0.24)
    edge = math.exp(-abs(first["edge"] - second["edge"]) / 0.18)
    score = (
        0.43 * rgb
        + 0.11 * gray
        + 0.12 * mean
        + 0.07 * std
        + 0.11 * grad
        + 0.10 * orient
        + 0.06 * edge
    )
    return max(0.0, min(1.0, float(score)))


def _nokta_poligonda_mi(lat, lon, points):
    inside = False
    previous_lat, previous_lon = points[-1]
    for current_lat, current_lon in points:
        crosses = (current_lat > lat) != (previous_lat > lat)
        if crosses:
            boundary_lon = (
                (previous_lon - current_lon)
                * (lat - current_lat)
                / ((previous_lat - current_lat) or 1.0e-15)
                + current_lon
            )
            if lon < boundary_lon:
                inside = not inside
        previous_lat, previous_lon = current_lat, current_lon
    return inside


def _ornek_noktalar(points, maksimum=36):
    cleaned = [(float(lat), float(lon)) for lat, lon in points or ()]
    if len(cleaned) > 1 and cleaned[0] == cleaned[-1]:
        cleaned.pop()
    if len(cleaned) < 3:
        raise JeolojiPaftaHatasi("Formasyon tanıması için kapalı bir parsel poligonu gerekir.")
    min_lat = min(point[0] for point in cleaned)
    max_lat = max(point[0] for point in cleaned)
    min_lon = min(point[1] for point in cleaned)
    max_lon = max(point[1] for point in cleaned)
    if min_lat == max_lat or min_lon == max_lon:
        raise JeolojiPaftaHatasi("Parsel poligonunun alanı sıfır olamaz.")

    result = []
    grid = max(3, int(math.sqrt(maksimum)) + 1)
    for row in range(grid):
        lat = min_lat + (row + 0.5) * (max_lat - min_lat) / grid
        for column in range(grid):
            lon = min_lon + (column + 0.5) * (max_lon - min_lon) / grid
            if _nokta_poligonda_mi(lat, lon, cleaned):
                result.append((lat, lon))
    center = (sum(p[0] for p in cleaned) / len(cleaned), sum(p[1] for p in cleaned) / len(cleaned))
    if _nokta_poligonda_mi(center[0], center[1], cleaned):
        result.insert(0, center)
    if not result:
        result = [center]
    if len(result) > maksimum:
        step = len(result) / maksimum
        result = [result[min(len(result) - 1, int(index * step))] for index in range(maksimum)]
    return cleaned, result


def _piksel(record, image, lat, lon):
    x, y = pafta_gorsel_koordinati(record, lat, lon)
    return x * (image.width - 1), y * (image.height - 1)


def _lejant_siniflari(profile):
    jpeg_path = Path((profile or {}).get("jpeg_path", ""))
    if not jpeg_path.is_file():
        raise JeolojiPaftaHatasi("Paftaya bağlı açıklamalı JPEG bulunamadı.")
    items = (profile or {}).get("ogeler", [])
    if not items:
        raise JeolojiPaftaHatasi("Bu paftanın lejant örnekleri henüz tanımlanmamış.")
    try:
        full = Image.open(jpeg_path)
        if full.width * full.height > GORSEL_AZAMI_PIKSEL:
            full.close()
            raise JeolojiPaftaHatasi("Açıklamalı JPEG güvenli piksel sınırını aşıyor.")
        full.load()
        full = full.convert("RGB")
    except (OSError, Image.DecompressionBombError) as exc:
        raise JeolojiPaftaHatasi(f"Açıklamalı JPEG açılamadı: {exc}") from exc

    classes = {}
    try:
        for item in items:
            rect = item.get("rect", [])
            if not isinstance(rect, list) or len(rect) != 4:
                continue
            left = max(0, min(full.width - 1, round(float(rect[0]) * full.width)))
            top = max(0, min(full.height - 1, round(float(rect[1]) * full.height)))
            right = max(left + 1, min(full.width, round(float(rect[2]) * full.width)))
            bottom = max(top + 1, min(full.height, round(float(rect[3]) * full.height)))
            crop = full.crop((left, top, right, bottom))
            key = (pafta_anahtari(item.get("kod")), pafta_anahtari(item.get("ad")))
            entry = classes.setdefault(
                key,
                {
                    "kod": str(item.get("kod", "")).strip(),
                    "ad": str(item.get("ad", "")).strip(),
                    "features": [],
                    "preview": None,
                    "kaynak_siralari": [],
                },
            )
            entry["features"].append(_gorsel_ozelligi(crop))
            try:
                kaynak_sirasi = int(item.get("sira"))
            except (TypeError, ValueError):
                # Eski hazırlanmış profillerde sıra alanı yoktur. Lejantın düşey
                # konumu gençten yaşlıya kaynak sırası olarak yeterli bir geri dönüş sağlar.
                kaynak_sirasi = round(float(rect[1]) * 10_000)
            entry["kaynak_siralari"].append(kaynak_sirasi)
            if entry["preview"] is None:
                preview = crop.convert("RGB")
                preview.thumbnail((420, 180), Image.Resampling.LANCZOS)
                entry["preview"] = preview
            crop.close()
    finally:
        full.close()
    classes = [entry for entry in classes.values() if entry["features"]]
    if not classes:
        raise JeolojiPaftaHatasi("Geçerli bir lejant örneği bulunamadı.")
    return classes


def _kanit_gorseli(record, image, polygon, label):
    pixel_points = [_piksel(record, image, lat, lon) for lat, lon in polygon]
    xs = [point[0] for point in pixel_points]
    ys = [point[1] for point in pixel_points]
    span = max(max(xs) - min(xs), max(ys) - min(ys), 12.0)
    padding = max(45.0, span * 1.25)
    left = max(0, int(min(xs) - padding))
    top = max(0, int(min(ys) - padding))
    right = min(image.width, int(max(xs) + padding) + 1)
    bottom = min(image.height, int(max(ys) + padding) + 1)
    crop = image.crop((left, top, right, bottom)).convert("RGB")
    local_points = [(round(x - left), round(y - top)) for x, y in pixel_points]
    if local_points and local_points[0] != local_points[-1]:
        local_points.append(local_points[0])
    draw = ImageDraw.Draw(crop)
    draw.line(local_points, fill="white", width=8, joint="curve")
    draw.line(local_points, fill="#d10018", width=4, joint="curve")

    if crop.width < 720:
        scale = min(4.0, 720.0 / max(crop.width, 1))
        crop = crop.resize(
            (round(crop.width * scale), round(crop.height * scale)),
            Image.Resampling.LANCZOS,
        )
    if crop.width > 1400 or crop.height > 1000:
        scale = min(1400.0 / crop.width, 1000.0 / crop.height)
        crop = crop.resize(
            (max(1, round(crop.width * scale)), max(1, round(crop.height * scale))),
            Image.Resampling.LANCZOS,
        )
    draw = ImageDraw.Draw(crop)
    text = str(label or "Parsel")
    box = draw.textbbox((0, 0), text)
    width = box[2] - box[0]
    height = box[3] - box[1]
    draw.rectangle((9, 9, 21 + width, 19 + height), fill="#202020")
    draw.text((15, 13), text, fill="white")
    return crop


def paftada_birim_tahmin_et(record, profile, points):
    """Tek bir KMZ paftasında parsel için aday birimleri ve kanıt görselini döndürür."""
    polygon, samples = _ornek_noktalar(points)
    classes = _lejant_siniflari(profile)
    image = kmz_goruntusunu_ac(record)
    previews_transferred = False
    try:
        polygon_pixels = [_piksel(record, image, lat, lon) for lat, lon in polygon]
        parcel_span = max(
            max(point[0] for point in polygon_pixels) - min(point[0] for point in polygon_pixels),
            max(point[1] for point in polygon_pixels) - min(point[1] for point in polygon_pixels),
            4.0,
        )
        radius = max(3, min(14, round(parcel_span / 9.0)))
        score_totals = [0.0] * len(classes)
        probability_totals = [0.0] * len(classes)
        used = 0
        for lat, lon in samples:
            x, y = _piksel(record, image, lat, lon)
            if not (0 <= x < image.width and 0 <= y < image.height):
                continue
            left = max(0, int(x) - radius)
            top = max(0, int(y) - radius)
            right = min(image.width, int(x) + radius + 1)
            bottom = min(image.height, int(y) + radius + 1)
            patch = image.crop((left, top, right, bottom))
            feature = _gorsel_ozelligi(patch)
            patch.close()
            scores = [
                max(_ozellik_benzerligi(feature, reference) for reference in entry["features"])
                for entry in classes
            ]
            maximum = max(scores)
            weights = [math.exp((score - maximum) / 0.055) for score in scores]
            weight_total = sum(weights) or 1.0
            for index, score in enumerate(scores):
                score_totals[index] += score
                probability_totals[index] += weights[index] / weight_total
            used += 1
        if not used:
            raise JeolojiPaftaHatasi("Parsel pafta görüntüsünün dışında kalıyor.")

        ranked_candidates = []
        for index, entry in enumerate(classes):
            mean_score = score_totals[index] / used
            share = probability_totals[index] / used
            confidence = max(0.0, min(1.0, 0.62 * mean_score + 0.38 * share))
            ranked_candidates.append(
                (
                    {
                        "kod": entry["kod"],
                        "ad": entry["ad"],
                        "oran": round(share * 100, 1),
                        "guven": round(confidence * 100, 1),
                        "puan": round(mean_score * 100, 1),
                        "pafta_id": record.get("id", ""),
                        "pafta_adi": record.get("ad", ""),
                        "lejant_id": record.get("lejant_id", ""),
                    },
                    entry.get("preview"),
                )
            )
        ranked_candidates.sort(
            key=lambda item: (-item[0]["oran"], -item[0]["puan"], item[0]["ad"])
        )
        selected = ranked_candidates[: min(8, len(ranked_candidates))]
        selected_previews = {id(preview) for _candidate, preview in selected if preview is not None}
        for _candidate, preview in ranked_candidates:
            if preview is not None and id(preview) not in selected_previews:
                preview.close()
        candidates = [candidate for candidate, _preview in selected]
        legend_previews = [preview for _candidate, preview in selected]
        top = candidates[0]
        evidence = _kanit_gorseli(
            record,
            image,
            polygon,
            f"{top['kod']} - {top['ad']}" if top["kod"] else top["ad"],
        )
        result = {
            "pafta": dict(record),
            "adaylar": candidates,
            "lejant_gorselleri": legend_previews,
            "ornek_sayisi": used,
            "kanit_gorseli": evidence,
            "dusuk_guven": bool(top["guven"] < 58 or top["oran"] < 45),
        }
        previews_transferred = True
        return result
    finally:
        image.close()
        if not previews_transferred:
            for entry in classes:
                try:
                    preview = entry.get("preview")
                    if preview is not None:
                        preview.close()
                except Exception:
                    pass


__all__ = [
    "paftada_birim_tahmin_et",
    "_gorsel_ozelligi",
    "_lejant_siniflari",
    "_ozellik_benzerligi",
]
