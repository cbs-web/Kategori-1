"""1/100.000 ölçekli raster jeoloji paftalarını indeksleme ve eşleştirme araçları."""

from __future__ import annotations

import datetime as dt
import hashlib
import io
import json
import math
import os
from pathlib import Path
import re
import shutil
import tempfile
import unicodedata
import xml.etree.ElementTree as ET
from zipfile import BadZipFile, ZipFile

from PIL import Image


PAFTA_KUTUPHANE_SURUMU = 1
KMZ_AZAMI_BOYUT = 250 * 1024 * 1024
GORSEL_AZAMI_BOYUT = 220 * 1024 * 1024
GORSEL_AZAMI_PIKSEL = 100_000_000


class JeolojiPaftaHatasi(ValueError):
    pass


def _simdi():
    return dt.datetime.now().isoformat(timespec="seconds")


def pafta_anahtari(value):
    text = unicodedata.normalize("NFKD", str(value or "").strip().casefold())
    text = "".join(char for char in text if not unicodedata.combining(char))
    text = text.replace("ı", "i")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def _yerel_ad(tag):
    return str(tag or "").rsplit("}", 1)[-1]


def _ilk_cocuk(node, name):
    if node is None:
        return None
    return next((child for child in node if _yerel_ad(child.tag) == name), None)


def _alt_metin(node, *names):
    current = node
    for name in names:
        current = _ilk_cocuk(current, name)
        if current is None:
            return ""
    return str(current.text or "").strip()


def _sonlu_sayi(value, label):
    try:
        number = float(value)
    except (TypeError, ValueError):
        raise JeolojiPaftaHatasi(f"KMZ içindeki {label} değeri sayısal değil.") from None
    if not (-1.0e12 < number < 1.0e12):
        raise JeolojiPaftaHatasi(f"KMZ içindeki {label} değeri geçersiz.")
    return number


def pafta_gorsel_koordinati(record, lat, lon):
    """Bir coğrafi noktayı GroundOverlay üzerinde 0-1 arası görüntü koordinatına çevirir."""
    bounds = (record or {}).get("bounds", {})
    north = _sonlu_sayi(bounds.get("north"), "north")
    south = _sonlu_sayi(bounds.get("south"), "south")
    east = _sonlu_sayi(bounds.get("east"), "east")
    west = _sonlu_sayi(bounds.get("west"), "west")
    if north <= south or east <= west:
        raise JeolojiPaftaHatasi("Paftanın koordinat sınırları geçersiz.")

    center_lat = (north + south) / 2.0
    center_lon = (east + west) / 2.0
    # Boylam derecesinin fiziksel karşılığı enleme göre değişir. Dönüşü pafta
    # merkezindeki yerel doğu-kuzey düzleminde yapmak görüntü oranını korur.
    longitude_scale = math.cos(math.radians(center_lat))
    if abs(longitude_scale) < 1.0e-12:
        raise JeolojiPaftaHatasi("Kutup noktasındaki GroundOverlay dönüştürülemiyor.")

    east_span = (east - west) * longitude_scale
    north_span = north - south
    east_offset = (_sonlu_sayi(lon, "boylam") - center_lon) * longitude_scale
    north_offset = _sonlu_sayi(lat, "enlem") - center_lat

    # KML'de pozitif rotation saat yönünün tersidir. Dünya noktasını kaynak
    # görüntüye taşımak için bu dönüşün tersini uygularız.
    angle = math.radians(_sonlu_sayi((record or {}).get("rotation", 0) or 0, "rotation"))
    cosine = math.cos(angle)
    sine = math.sin(angle)
    source_east = cosine * east_offset + sine * north_offset
    source_north = -sine * east_offset + cosine * north_offset

    return (
        0.5 + source_east / east_span,
        0.5 - source_north / north_span,
    )


def _zip_gorsel_girdisi(archive, href):
    wanted = str(href or "").replace("\\", "/").lstrip("./")
    for entry in archive.infolist():
        if entry.filename.replace("\\", "/").lstrip("./").casefold() == wanted.casefold():
            return entry
    wanted_stem = pafta_anahtari(Path(wanted).stem)
    images = [
        entry
        for entry in archive.infolist()
        if Path(entry.filename).suffix.lower() in (".jpg", ".jpeg", ".png", ".tif", ".tiff")
    ]
    same_stem = [entry for entry in images if pafta_anahtari(Path(entry.filename).stem) == wanted_stem]
    if len(same_stem) == 1:
        return same_stem[0]
    if len(images) == 1:
        return images[0]
    return None


def kmz_bilgilerini_oku(path):
    source = Path(path)
    if source.suffix.lower() != ".kmz" or not source.is_file():
        raise JeolojiPaftaHatasi("Geçerli bir KMZ dosyası seçilmelidir.")
    if source.stat().st_size > KMZ_AZAMI_BOYUT:
        raise JeolojiPaftaHatasi("KMZ dosyası izin verilen boyuttan büyük.")
    results = []
    try:
        with ZipFile(source) as archive:
            kml_entries = [entry for entry in archive.infolist() if entry.filename.lower().endswith(".kml")]
            if not kml_entries:
                raise JeolojiPaftaHatasi("KMZ içinde KML belgesi bulunamadı.")
            for kml_entry in kml_entries:
                if kml_entry.file_size > 20 * 1024 * 1024:
                    raise JeolojiPaftaHatasi("KMZ içindeki KML belgesi çok büyük.")
                root = ET.fromstring(archive.read(kml_entry))
                for overlay in root.iter():
                    if _yerel_ad(overlay.tag) != "GroundOverlay":
                        continue
                    box = next(
                        (node for node in overlay.iter() if _yerel_ad(node.tag) == "LatLonBox"),
                        None,
                    )
                    if box is None:
                        continue
                    href = _alt_metin(overlay, "Icon", "href")
                    asset = _zip_gorsel_girdisi(archive, href)
                    if asset is None:
                        raise JeolojiPaftaHatasi(
                            f"GroundOverlay görüntüsü KMZ içinde bulunamadı: {href or '—'}"
                        )
                    if asset.file_size > GORSEL_AZAMI_BOYUT:
                        raise JeolojiPaftaHatasi("KMZ içindeki pafta görüntüsü çok büyük.")
                    bounds = {
                        "north": _sonlu_sayi(_alt_metin(box, "north"), "north"),
                        "south": _sonlu_sayi(_alt_metin(box, "south"), "south"),
                        "east": _sonlu_sayi(_alt_metin(box, "east"), "east"),
                        "west": _sonlu_sayi(_alt_metin(box, "west"), "west"),
                    }
                    if not (
                        -90 <= bounds["south"] < bounds["north"] <= 90
                        and -180 <= bounds["west"] < bounds["east"] <= 180
                    ):
                        raise JeolojiPaftaHatasi("GroundOverlay koordinat sınırları geçersiz.")
                    rotation_text = _alt_metin(box, "rotation")
                    rotation = _sonlu_sayi(rotation_text or 0.0, "rotation")
                    name = _alt_metin(overlay, "name") or source.stem
                    identity = "|".join(
                        (
                            os.path.normcase(str(source.resolve())).casefold(),
                            asset.filename,
                            *(f"{bounds[key]:.10f}" for key in ("north", "south", "east", "west")),
                        )
                    )
                    results.append(
                        {
                            "id": hashlib.sha256(identity.encode("utf-8")).hexdigest()[:20],
                            "ad": name,
                            "kmz_path": str(source.resolve()),
                            "overlay_href": asset.filename,
                            "bounds": bounds,
                            "rotation": rotation,
                        }
                    )
    except (BadZipFile, ET.ParseError, OSError) as exc:
        raise JeolojiPaftaHatasi(f"KMZ okunamadı: {exc}") from exc
    if not results:
        raise JeolojiPaftaHatasi("KMZ içinde LatLonBox kullanan GroundOverlay bulunamadı.")
    return results


def kmz_goruntusunu_ac(record):
    path = Path((record or {}).get("kmz_path", ""))
    href = (record or {}).get("overlay_href", "")
    if not path.is_file():
        raise JeolojiPaftaHatasi(f"KMZ dosyası bulunamadı: {path}")
    try:
        with ZipFile(path) as archive:
            entry = _zip_gorsel_girdisi(archive, href)
            if entry is None:
                raise JeolojiPaftaHatasi("KMZ içindeki pafta görüntüsü bulunamadı.")
            if entry.file_size > GORSEL_AZAMI_BOYUT:
                raise JeolojiPaftaHatasi("KMZ içindeki pafta görüntüsü çok büyük.")
            data = archive.read(entry)
        image = Image.open(io.BytesIO(data))
        if image.width * image.height > GORSEL_AZAMI_PIKSEL:
            image.close()
            raise JeolojiPaftaHatasi("Pafta görüntüsünün piksel sayısı güvenli sınırı aşıyor.")
        image.load()
        converted = image.convert("RGB")
        image.close()
        return converted
    except (BadZipFile, OSError) as exc:
        raise JeolojiPaftaHatasi(f"KMZ görüntüsü açılamadı: {exc}") from exc


def _dosyalari_bul(root, suffixes):
    root = Path(root)
    if root.is_file():
        candidates = [root]
    elif root.is_dir():
        candidates = [path for path in root.rglob("*") if path.is_file()]
    else:
        return []
    return sorted(
        (str(path.resolve()) for path in candidates if path.suffix.lower() in suffixes),
        key=lambda value: os.path.normcase(value).casefold(),
    )


def kmz_dosyalari_bul(root):
    return _dosyalari_bul(root, {".kmz"})


def jpeg_dosyalari_bul(root):
    return _dosyalari_bul(root, {".jpg", ".jpeg"})


def _ad_eslesme_puani(record, jpeg_path):
    jpeg_key = pafta_anahtari(Path(jpeg_path).stem)
    href_key = pafta_anahtari(Path(record.get("overlay_href", "")).stem)
    kmz_key = pafta_anahtari(Path(record.get("kmz_path", "")).stem)
    overlay_key = pafta_anahtari(record.get("ad", ""))
    if href_key and href_key == jpeg_key:
        return 1000
    if kmz_key and kmz_key == jpeg_key:
        return 950
    jpeg_tokens = set(jpeg_key.split())
    best = 0
    for key in (href_key, kmz_key, overlay_key):
        tokens = set(key.split())
        if not tokens or not jpeg_tokens:
            continue
        common = len(tokens & jpeg_tokens)
        union = len(tokens | jpeg_tokens)
        best = max(best, int(500 * common / max(union, 1)))
    return best


def _cv_gorsel_oku(path):
    try:
        import cv2
        import numpy as np
    except Exception:
        return None
    try:
        data = np.fromfile(os.fspath(path), dtype=np.uint8)
        image = cv2.imdecode(data, cv2.IMREAD_GRAYSCALE)
    except (OSError, ValueError):
        return None
    return image


def _cv_boyutlandir(image, maximum=2200):
    import cv2

    scale = min(1.0, float(maximum) / max(image.shape[:2]))
    if scale >= 1.0:
        return image
    return cv2.resize(image, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA)


def _gorsel_icerik_eslesme_puani(record, jpeg_path):
    """KMZ kırpımını tam paftada ORB/RANSAC ile arar; OpenCV yoksa sıfır döner."""
    try:
        import cv2
        import numpy as np
    except Exception:
        return 0
    try:
        overlay_image = kmz_goruntusunu_ac(record)
        overlay_array = np.asarray(overlay_image.convert("L"))
        overlay_image.close()
        full_array = _cv_gorsel_oku(jpeg_path)
        if full_array is None:
            return 0
        overlay_array = _cv_boyutlandir(overlay_array)
        full_array = _cv_boyutlandir(full_array)
        orb = cv2.ORB_create(nfeatures=7000, scaleFactor=1.2, nlevels=10)
        overlay_points, overlay_desc = orb.detectAndCompute(overlay_array, None)
        full_points, full_desc = orb.detectAndCompute(full_array, None)
        if overlay_desc is None or full_desc is None:
            return 0
        pairs = cv2.BFMatcher(cv2.NORM_HAMMING).knnMatch(overlay_desc, full_desc, k=2)
        good = [first for first, second in pairs if first.distance < 0.72 * second.distance]
        if len(good) < 8:
            return 0
        source_points = np.float32(
            [overlay_points[match.queryIdx].pt for match in good]
        ).reshape(-1, 1, 2)
        target_points = np.float32(
            [full_points[match.trainIdx].pt for match in good]
        ).reshape(-1, 1, 2)
        _matrix, mask = cv2.findHomography(source_points, target_points, cv2.RANSAC, 5.0)
        return int(mask.sum()) if mask is not None else 0
    except Exception:
        return 0


def paftalari_eslestir(kmz_paths, jpeg_paths, ilerleme=None):
    records = []
    errors = []
    for index, path in enumerate(kmz_paths, start=1):
        if ilerleme:
            ilerleme(index - 1, max(len(kmz_paths), 1), f"KMZ okunuyor: {Path(path).name}")
        try:
            records.extend(kmz_bilgilerini_oku(path))
        except JeolojiPaftaHatasi as exc:
            errors.append({"path": path, "hata": str(exc)})
    jpeg_paths = [str(Path(path).resolve()) for path in jpeg_paths if Path(path).is_file()]
    direct_used = set()
    unmatched = []
    for record in records:
        scored = sorted(
            ((_ad_eslesme_puani(record, path), path) for path in jpeg_paths),
            reverse=True,
        )
        best_score, best_path = scored[0] if scored else (0, "")
        if best_score >= 900:
            record["jpeg_path"] = best_path
            record["eslesme_yontemi"] = "dosya_adi"
            direct_used.add(os.path.normcase(best_path).casefold())
        else:
            record["jpeg_path"] = ""
            record["eslesme_yontemi"] = ""
            unmatched.append(record)

    for record_index, record in enumerate(unmatched, start=1):
        if ilerleme:
            ilerleme(
                len(kmz_paths) + record_index - 1,
                max(len(kmz_paths) + len(unmatched), 1),
                f"Görüntü eşleştiriliyor: {record['ad']}",
            )
        scores = sorted(
            ((_gorsel_icerik_eslesme_puani(record, path), path) for path in jpeg_paths),
            reverse=True,
        )
        best_score, best_path = scores[0] if scores else (0, "")
        second_score = scores[1][0] if len(scores) > 1 else 0
        if best_score >= 10 and best_score >= second_score + 4:
            record["jpeg_path"] = best_path
            record["eslesme_yontemi"] = "gorsel_icerik"
            direct_used.add(os.path.normcase(best_path).casefold())

    unmatched_jpegs = [
        path for path in jpeg_paths if os.path.normcase(path).casefold() not in direct_used
    ]
    return {"paftalar": records, "hatalar": errors, "eslesmeyen_jpeg": unmatched_jpegs}


class JeolojiPaftaKutuphanesi:
    def __init__(self, root_dir):
        self.root_dir = Path(root_dir)
        self.root_dir.mkdir(parents=True, exist_ok=True)
        self.json_path = self.root_dir / "paftalar.json"
        self.kanit_dir = self.root_dir / "kanitlar"
        self.kmz_dir = self.root_dir / "kmz"
        self.jpeg_dir = self.root_dir / "jpeg"
        self._data = self._oku()

    @staticmethod
    def _bos_veri():
        return {"version": PAFTA_KUTUPHANE_SURUMU, "paftalar": [], "lejantlar": {}}

    def _oku(self):
        try:
            with self.json_path.open("r", encoding="utf-8-sig") as stream:
                data = json.load(stream)
        except (OSError, json.JSONDecodeError):
            return self._bos_veri()
        if not isinstance(data, dict):
            return self._bos_veri()
        data.setdefault("paftalar", [])
        data.setdefault("lejantlar", {})
        data["version"] = PAFTA_KUTUPHANE_SURUMU
        return data

    def _yaz(self):
        self.root_dir.mkdir(parents=True, exist_ok=True)
        fd, temporary = tempfile.mkstemp(prefix=".paftalar.", suffix=".tmp", dir=self.root_dir)
        try:
            with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as stream:
                json.dump(self._data, stream, ensure_ascii=False, indent=2, allow_nan=False)
                stream.flush()
                os.fsync(stream.fileno())
            os.replace(temporary, self.json_path)
        except Exception:
            try:
                os.close(fd)
            except OSError:
                pass
            try:
                os.remove(temporary)
            except OSError:
                pass
            raise

    @staticmethod
    def _profil_id(jpeg_path):
        key = os.path.normcase(str(Path(jpeg_path).resolve())).casefold()
        return hashlib.sha256(key.encode("utf-8")).hexdigest()[:20]

    def _varligi_yerellestir(self, source_path, directory, suffixes, maximum_size):
        source = Path(source_path)
        if source.suffix.lower() not in suffixes or not source.is_file():
            raise JeolojiPaftaHatasi(f"Kütüphaneye alınacak dosya bulunamadı: {source}")
        if source.stat().st_size > maximum_size:
            raise JeolojiPaftaHatasi(f"Kütüphaneye alınacak dosya çok büyük: {source.name}")

        directory.mkdir(parents=True, exist_ok=True)
        source_resolved = source.resolve()
        directory_resolved = directory.resolve()
        if source_resolved.parent == directory_resolved:
            return str(source_resolved)

        source_key = os.path.normcase(str(source_resolved)).casefold()
        asset_id = hashlib.sha256(source_key.encode("utf-8")).hexdigest()[:20]
        target = directory_resolved / f"{asset_id}{source.suffix.lower()}"
        fd, temporary = tempfile.mkstemp(
            prefix=f".{asset_id}.",
            suffix=f"{source.suffix.lower()}.tmp",
            dir=directory_resolved,
        )
        os.close(fd)
        try:
            shutil.copy2(source_resolved, temporary)
            os.replace(temporary, target)
        except Exception:
            try:
                os.remove(temporary)
            except OSError:
                pass
            raise
        return str(target.resolve())

    def varliklari_yerellestir(self):
        """Mevcut dış KMZ/JPEG bağlantılarını kalıcı K-1 kopyalarına dönüştür."""
        changed_records = 0
        copied_paths = set()
        errors = []
        for record in self._data.get("paftalar", []):
            record_changed = False
            for field, directory, suffixes, maximum in (
                ("kmz_path", self.kmz_dir, {".kmz"}, KMZ_AZAMI_BOYUT),
                ("jpeg_path", self.jpeg_dir, {".jpg", ".jpeg"}, GORSEL_AZAMI_BOYUT),
            ):
                source_path = str(record.get(field) or "").strip()
                if not source_path:
                    continue
                try:
                    local_path = self._varligi_yerellestir(
                        source_path, directory, suffixes, maximum
                    )
                except (OSError, JeolojiPaftaHatasi) as exc:
                    errors.append(
                        {"pafta": record.get("ad", "Adsız pafta"), "alan": field, "hata": str(exc)}
                    )
                    continue
                copied_paths.add(local_path)
                if os.path.normcase(source_path) != os.path.normcase(local_path):
                    record[field] = local_path
                    record_changed = True
                if field == "jpeg_path" and record.get("lejant_id"):
                    profile = self._data.get("lejantlar", {}).get(record["lejant_id"])
                    if profile is not None:
                        profile["jpeg_path"] = local_path
                        profile["guncelleme_tarihi"] = _simdi()
            if record_changed:
                record["guncelleme_tarihi"] = _simdi()
                changed_records += 1
        if changed_records:
            self._yaz()
        return {
            "degisen_pafta": changed_records,
            "yerel_dosya": len(copied_paths),
            "hatalar": errors,
        }

    def listele(self):
        return [dict(record) for record in self._data.get("paftalar", [])]

    def getir(self, record_id):
        return next(
            (dict(record) for record in self._data.get("paftalar", []) if record.get("id") == record_id),
            None,
        )

    def lejant_getir(self, profile_id):
        profile = self._data.get("lejantlar", {}).get(str(profile_id or ""))
        return json.loads(json.dumps(profile, ensure_ascii=False)) if profile else None

    def _profili_hazirla(self, jpeg_path):
        profile_id = self._profil_id(jpeg_path)
        profiles = self._data.setdefault("lejantlar", {})
        profile = profiles.setdefault(
            profile_id,
            {
                "id": profile_id,
                "ad": Path(jpeg_path).stem,
                "jpeg_path": str(Path(jpeg_path).resolve()),
                "ogeler": [],
                "guncelleme_tarihi": _simdi(),
            },
        )
        profile["jpeg_path"] = str(Path(jpeg_path).resolve())
        if not profile.get("ogeler"):
            try:
                from jeoloji_pafta_hazir_profilleri import hazir_profil_ogeleri

                prepared = hazir_profil_ogeleri(profile["jpeg_path"])
            except Exception:
                prepared = []
            if prepared:
                items = []
                for item in prepared:
                    identity = (
                        f"{profile_id}|hazir|{item.get('sira', '')}|"
                        f"{item.get('kod', '')}|{item.get('ad', '')}"
                    )
                    items.append(
                        {
                            "id": hashlib.sha256(identity.encode("utf-8")).hexdigest()[:16],
                            "kod": item["kod"],
                            "ad": item["ad"],
                            "rect": item["rect"],
                            "sira": item.get("sira"),
                            "kaynak": "hazir_profil",
                        }
                    )
                profile["ogeler"] = items
                profile["kaynak"] = "K-1 hazır 1/100.000 pafta profili"
                profile["guncelleme_tarihi"] = _simdi()
        return profile_id

    def toplu_ice_aktar(self, kmz_root, jpeg_root, ilerleme=None):
        kmz_paths = kmz_dosyalari_bul(kmz_root)
        jpeg_paths = jpeg_dosyalari_bul(jpeg_root)
        result = paftalari_eslestir(kmz_paths, jpeg_paths, ilerleme=ilerleme)
        existing = {record.get("id"): record for record in self._data.get("paftalar", [])}
        merged = []
        for record in result["paftalar"]:
            previous = existing.get(record["id"], {})
            record["kmz_path"] = self._varligi_yerellestir(
                record["kmz_path"], self.kmz_dir, {".kmz"}, KMZ_AZAMI_BOYUT
            )
            if record.get("jpeg_path"):
                record["jpeg_path"] = self._varligi_yerellestir(
                    record["jpeg_path"],
                    self.jpeg_dir,
                    {".jpg", ".jpeg"},
                    GORSEL_AZAMI_BOYUT,
                )
                previous_profile_id = previous.get("lejant_id", "")
                previous_profile = self._data.get("lejantlar", {}).get(previous_profile_id)
                if previous_profile is not None:
                    previous_profile["jpeg_path"] = record["jpeg_path"]
                    previous_profile["guncelleme_tarihi"] = _simdi()
                    record["lejant_id"] = previous_profile_id
                else:
                    record["lejant_id"] = self._profili_hazirla(record["jpeg_path"])
            else:
                record["lejant_id"] = previous.get("lejant_id", "")
                if record["lejant_id"]:
                    profile = self._data.get("lejantlar", {}).get(record["lejant_id"], {})
                    record["jpeg_path"] = profile.get("jpeg_path", "")
            record["ekleme_tarihi"] = previous.get("ekleme_tarihi") or _simdi()
            record["guncelleme_tarihi"] = _simdi()
            merged.append(record)
        scanned_ids = {record["id"] for record in merged}
        merged.extend(record for record in existing.values() if record.get("id") not in scanned_ids)
        self._data["paftalar"] = sorted(
            merged, key=lambda item: pafta_anahtari(item.get("ad", ""))
        )
        self._yaz()
        result["kmz_sayisi"] = len(kmz_paths)
        result["jpeg_sayisi"] = len(jpeg_paths)
        return result

    def jpeg_bagla(self, record_id, jpeg_path, yontem="kullanici"):
        path = Path(jpeg_path)
        if path.suffix.lower() not in (".jpg", ".jpeg") or not path.is_file():
            raise JeolojiPaftaHatasi("Geçerli bir açıklamalı JPEG seçilmelidir.")
        record = next(
            (item for item in self._data.get("paftalar", []) if item.get("id") == record_id),
            None,
        )
        if record is None:
            raise JeolojiPaftaHatasi("Güncellenecek pafta bulunamadı.")
        local_path = self._varligi_yerellestir(
            path, self.jpeg_dir, {".jpg", ".jpeg"}, GORSEL_AZAMI_BOYUT
        )
        record["jpeg_path"] = local_path
        record["lejant_id"] = self._profili_hazirla(local_path)
        record["eslesme_yontemi"] = yontem
        record["guncelleme_tarihi"] = _simdi()
        self._yaz()
        return dict(record)

    def lejant_ogesi_kaydet(self, profile_id, code, name, rect, item_id=None):
        profile = self._data.get("lejantlar", {}).get(profile_id)
        if not profile:
            raise JeolojiPaftaHatasi("Lejant profili bulunamadı.")
        code = str(code or "").strip()
        name = str(name or "").strip()
        if not code or not name:
            raise JeolojiPaftaHatasi("Birim kodu ve adı zorunludur.")
        if not isinstance(rect, (list, tuple)) or len(rect) != 4:
            raise JeolojiPaftaHatasi("Lejant örneği için geçerli bir dikdörtgen seçilmelidir.")
        try:
            clean_rect = [max(0.0, min(1.0, float(value))) for value in rect]
        except (TypeError, ValueError):
            raise JeolojiPaftaHatasi("Lejant dikdörtgeni sayısal olmalıdır.") from None
        if clean_rect[2] - clean_rect[0] < 0.0005 or clean_rect[3] - clean_rect[1] < 0.0005:
            raise JeolojiPaftaHatasi("Seçilen lejant örneği çok küçük.")
        if item_id is None:
            identity = f"{profile_id}|{code}|{name}|{_simdi()}"
            item_id = hashlib.sha256(identity.encode("utf-8")).hexdigest()[:16]
        item = {"id": item_id, "kod": code, "ad": name, "rect": clean_rect}
        items = profile.setdefault("ogeler", [])
        index = next((i for i, current in enumerate(items) if current.get("id") == item_id), None)
        if index is None:
            items.append(item)
        else:
            items[index] = item
        profile["guncelleme_tarihi"] = _simdi()
        self._yaz()
        return dict(item)

    def lejant_ogesi_sil(self, profile_id, item_id):
        profile = self._data.get("lejantlar", {}).get(profile_id)
        if not profile:
            return False
        before = len(profile.get("ogeler", []))
        profile["ogeler"] = [item for item in profile.get("ogeler", []) if item.get("id") != item_id]
        if len(profile["ogeler"]) == before:
            return False
        profile["guncelleme_tarihi"] = _simdi()
        self._yaz()
        return True

    def kapsayan_paftalar(self, points):
        points = [(float(lat), float(lon)) for lat, lon in points or ()]
        if not points:
            return []
        min_lat = min(point[0] for point in points)
        max_lat = max(point[0] for point in points)
        min_lon = min(point[1] for point in points)
        max_lon = max(point[1] for point in points)
        center_lat = (min_lat + max_lat) / 2
        center_lon = (min_lon + max_lon) / 2
        results = []
        for record in self._data.get("paftalar", []):
            try:
                image_points = [pafta_gorsel_koordinati(record, lat, lon) for lat, lon in points]
                center_x, center_y = pafta_gorsel_koordinati(record, center_lat, center_lon)
            except JeolojiPaftaHatasi:
                continue
            min_x = min(point[0] for point in image_points)
            max_x = max(point[0] for point in image_points)
            min_y = min(point[1] for point in image_points)
            max_y = max(point[1] for point in image_points)
            overlap_x = max(0.0, min(max_x, 1.0) - max(min_x, 0.0))
            overlap_y = max(0.0, min(max_y, 1.0) - max(min_y, 0.0))
            center_inside = bool(0.0 <= center_x <= 1.0 and 0.0 <= center_y <= 1.0)
            if overlap_x <= 0 or overlap_y <= 0:
                if not center_inside:
                    continue
                coverage = 1.0
            else:
                parcel_area = max((max_x - min_x) * (max_y - min_y), 1.0e-15)
                coverage = min(1.0, overlap_x * overlap_y / parcel_area)
            item = dict(record)
            item["kapsama_orani"] = coverage
            item["merkezi_kapsiyor"] = center_inside
            results.append(item)
        return sorted(
            results,
            key=lambda item: (-int(item["merkezi_kapsiyor"]), -item["kapsama_orani"], pafta_anahtari(item["ad"])),
        )

    def kanit_kaydet(self, image, target_dir=None):
        directory = Path(target_dir) if target_dir else self.kanit_dir
        directory.mkdir(parents=True, exist_ok=True)
        target = directory / "Jeoloji_Birim_Tespiti.jpg"
        image.convert("RGB").save(target, format="JPEG", quality=94, optimize=True)
        return str(target.resolve())


__all__ = [
    "JeolojiPaftaHatasi",
    "JeolojiPaftaKutuphanesi",
    "jpeg_dosyalari_bul",
    "kmz_bilgilerini_oku",
    "kmz_dosyalari_bul",
    "kmz_goruntusunu_ac",
    "pafta_anahtari",
    "pafta_gorsel_koordinati",
    "paftalari_eslestir",
]
