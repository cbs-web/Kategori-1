"""1/100.000 jeoloji paftaları için isteğe bağlı yapay zekâ denetimi.

Yerel görüntü sınıflandırması ana akış olarak kalır. Gemini ve GPT-5.6 Sol,
kullanıcının bastığı iki ayrı düğmeyle birbirinden bağımsız çalışır. API
anahtarları proje JSON'una veya bu modülün ayar dosyasına yazılmaz; Windows
Kimlik Bilgileri Yöneticisi kullanılır.
"""

from __future__ import annotations

import base64
import hashlib
import io
import json
import os
from pathlib import Path
import socket
import time
import urllib.error
import urllib.request

from PIL import Image, ImageDraw, ImageFont

from jeoloji_kutuphanesi import jeoloji_anahtari


PROMPT_SURUMU = "2026-08-10.1"
VARSAYILAN_AYARLAR = {
    "birincil_model": "gemini-3.6-flash",
    "ikinci_model": "gpt-5.6-sol",
    "onbellek_etkin": True,
    "zaman_asimi_saniye": 90,
}
SAGLAYICILAR = {
    "gemini": {
        "hedef": "K-1/JeolojiAI/Gemini",
        "env": "GEMINI_API_KEY",
        "etiket": "Gemini 3.6 Flash",
    },
    "openai": {
        "hedef": "K-1/JeolojiAI/OpenAI",
        "env": "OPENAI_API_KEY",
        "etiket": "GPT-5.6 Sol",
    },
}
KANITLAR = {"kod_okundu", "kod_ve_desen", "yalniz_desen"}


class JeolojiYapayZekaHatasi(RuntimeError):
    """Kullanıcıya gösterilebilecek yapay zekâ denetimi hatası."""


class YapayZekaAnahtariEksik(JeolojiYapayZekaHatasi):
    """Hiçbir kullanılabilir API anahtarı bulunamadığında yükseltilir."""


def _font(size, bold=False):
    candidates = (
        r"C:\Windows\Fonts\arialbd.ttf" if bold else r"C:\Windows\Fonts\arial.ttf",
        "DejaVuSans-Bold.ttf" if bold else "DejaVuSans.ttf",
    )
    for candidate in candidates:
        try:
            return ImageFont.truetype(candidate, size=size)
        except OSError:
            continue
    return ImageFont.load_default()


def _metni_kisalt(value, length=48):
    text = " ".join(str(value or "").split())
    return text if len(text) <= length else text[: max(1, length - 1)].rstrip() + "…"


class JeolojiYapayZekaAyarlari:
    """Gizli olmayan seçenekleri ve Windows'taki güvenli anahtarları yönetir."""

    def __init__(self, kullanici_veri_klasoru):
        self.root = Path(kullanici_veri_klasoru) / "jeoloji_yapay_zeka"
        self.ayar_yolu = self.root / "ayarlar.json"
        self.onbellek_klasoru = self.root / "onbellek"

    def oku(self):
        settings = dict(VARSAYILAN_AYARLAR)
        try:
            with self.ayar_yolu.open("r", encoding="utf-8-sig") as stream:
                loaded = json.load(stream)
        except (OSError, json.JSONDecodeError):
            loaded = {}
        if isinstance(loaded, dict):
            for key in settings:
                if key in loaded:
                    settings[key] = loaded[key]
        settings["birincil_model"] = str(settings.get("birincil_model") or "gemini-3.6-flash")
        settings["ikinci_model"] = str(settings.get("ikinci_model") or "gpt-5.6-sol")
        settings["onbellek_etkin"] = bool(settings.get("onbellek_etkin", True))
        try:
            settings["zaman_asimi_saniye"] = max(
                30, min(180, int(settings.get("zaman_asimi_saniye", 90)))
            )
        except (TypeError, ValueError):
            settings["zaman_asimi_saniye"] = 90
        return settings

    def kaydet(self, settings):
        clean = self.oku()
        for key in ("onbellek_etkin",):
            if key in settings:
                clean[key] = settings[key]
        # Model adları kod tarafından yönetilir; anahtarlar bu dosyaya asla girmez.
        clean["birincil_model"] = VARSAYILAN_AYARLAR["birincil_model"]
        clean["ikinci_model"] = VARSAYILAN_AYARLAR["ikinci_model"]
        clean = JeolojiYapayZekaAyarlari._temiz_ayarlar(clean)
        self.root.mkdir(parents=True, exist_ok=True)
        temporary = self.ayar_yolu.with_suffix(".tmp")
        temporary.write_text(
            json.dumps(clean, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        os.replace(temporary, self.ayar_yolu)
        return clean

    @staticmethod
    def _temiz_ayarlar(settings):
        result = dict(VARSAYILAN_AYARLAR)
        result.update({key: settings.get(key, value) for key, value in result.items()})
        result["onbellek_etkin"] = bool(result["onbellek_etkin"])
        return result

    @staticmethod
    def _win32cred():
        try:
            import win32cred  # type: ignore
        except Exception as exc:
            raise JeolojiYapayZekaHatasi(
                "API anahtarı güvenli biçimde kaydedilemedi. pywin32/Windows Kimlik "
                "Bilgileri Yöneticisi kullanılamıyor; GEMINI_API_KEY veya OPENAI_API_KEY "
                "ortam değişkenini kullanın."
            ) from exc
        return win32cred

    @staticmethod
    def _blob_coz(blob):
        if isinstance(blob, str):
            return blob.strip()
        if not isinstance(blob, (bytes, bytearray)):
            return ""
        raw = bytes(blob)
        prefix = b"K1AI-UTF8:"
        if raw.startswith(prefix):
            return raw[len(prefix):].decode("utf-8", errors="strict").strip()
        for encoding in ("utf-16-le", "utf-8"):
            try:
                return raw.decode(encoding).rstrip("\x00").strip()
            except UnicodeDecodeError:
                continue
        return ""

    def anahtar_al(self, provider):
        info = SAGLAYICILAR.get(provider)
        if not info:
            return ""
        environmental = str(os.environ.get(info["env"], "")).strip()
        if environmental:
            return environmental
        try:
            win32cred = self._win32cred()
            credential = win32cred.CredRead(
                info["hedef"], win32cred.CRED_TYPE_GENERIC, 0
            )
        except Exception:
            return ""
        return self._blob_coz(credential.get("CredentialBlob"))

    def anahtar_kaynagi(self, provider):
        info = SAGLAYICILAR.get(provider)
        if not info:
            return "yok"
        if str(os.environ.get(info["env"], "")).strip():
            return f"{info['env']} ortam değişkeni"
        try:
            win32cred = self._win32cred()
            credential = win32cred.CredRead(
                info["hedef"], win32cred.CRED_TYPE_GENERIC, 0
            )
            if self._blob_coz(credential.get("CredentialBlob")):
                return "Windows Kimlik Bilgileri Yöneticisi"
        except Exception:
            pass
        return "yok"

    def anahtar_kaydet(self, provider, key):
        info = SAGLAYICILAR.get(provider)
        key = str(key or "").strip()
        if not info or not key:
            raise JeolojiYapayZekaHatasi("Kaydedilecek API anahtarı boş olamaz.")
        win32cred = self._win32cred()
        win32cred.CredWrite(
            {
                "Type": win32cred.CRED_TYPE_GENERIC,
                "TargetName": info["hedef"],
                # Bu pywin32 sürümünde CredWrite, CredentialBlob için Unicode
                # metin bekliyor; bytes verilirse "cannot be converted to
                # Unicode" hatası oluşuyor. Değer yine Windows Kimlik Bilgileri
                # Yöneticisi'nin şifreli kullanıcı kasasında tutulur.
                "CredentialBlob": key,
                "Persist": win32cred.CRED_PERSIST_LOCAL_MACHINE,
                "UserName": "K-1",
                "Comment": "K-1 jeoloji paftası yapay zekâ denetimi",
            },
            0,
        )

    def anahtar_sil(self, provider):
        info = SAGLAYICILAR.get(provider)
        if not info:
            return False
        win32cred = self._win32cred()
        try:
            win32cred.CredDelete(info["hedef"], win32cred.CRED_TYPE_GENERIC, 0)
        except Exception:
            return False
        return True


def _jpeg_bytes(image, maximum=1900, quality=92):
    working = image.convert("RGB")
    if max(working.size) > maximum:
        ratio = maximum / max(working.size)
        resized = working.resize(
            (max(1, round(working.width * ratio)), max(1, round(working.height * ratio))),
            Image.Resampling.LANCZOS,
        )
        working.close()
        working = resized
    stream = io.BytesIO()
    working.save(stream, format="JPEG", quality=quality, subsampling=0, optimize=True)
    working.close()
    return stream.getvalue()


def _harita_gorselleri(harita):
    result = [{"ad": "Genel harita", "mime_type": "image/jpeg", "data": _jpeg_bytes(harita)}]
    width, height = harita.size
    boxes = (
        (0, 0, round(width * 0.56), round(height * 0.56)),
        (round(width * 0.44), 0, width, round(height * 0.56)),
        (0, round(height * 0.44), round(width * 0.56), height),
        (round(width * 0.44), round(height * 0.44), width, height),
    )
    names = ("Kuzeybatı ayrıntısı", "Kuzeydoğu ayrıntısı", "Güneybatı ayrıntısı", "Güneydoğu ayrıntısı")
    for name, box in zip(names, boxes):
        crop = harita.crop(box)
        result.append(
            {"ad": name, "mime_type": "image/jpeg", "data": _jpeg_bytes(crop, maximum=1650, quality=94)}
        )
        crop.close()
    return result


def _lejant_sayfalari(catalog):
    with_preview = [unit for unit in catalog if unit.get("onizleme") is not None]
    if not with_preview:
        return []
    columns, rows = 4, 5
    cell_width, cell_height = 350, 170
    header_height = 54
    page_size = columns * rows
    images = []
    for page_no, start in enumerate(range(0, len(with_preview), page_size), 1):
        page_units = with_preview[start:start + page_size]
        canvas = Image.new("RGB", (columns * cell_width, header_height + rows * cell_height), "white")
        draw = ImageDraw.Draw(canvas)
        draw.text(
            (18, 13),
            f"Tanımlı pafta lejantları — sayfa {page_no}",
            fill="black",
            font=_font(25, bold=True),
        )
        for index, unit in enumerate(page_units):
            row, column = divmod(index, columns)
            left = column * cell_width
            top = header_height + row * cell_height
            draw.rectangle(
                (left + 5, top + 5, left + cell_width - 5, top + cell_height - 5),
                fill="white",
                outline="#707070",
                width=2,
            )
            preview = unit.get("onizleme").convert("RGB")
            preview.thumbnail((cell_width - 24, 112), Image.Resampling.LANCZOS)
            canvas.paste(
                preview,
                (left + 12 + (cell_width - 24 - preview.width) // 2, top + 10),
            )
            preview.close()
            code = str(unit.get("kod") or "").strip()
            name = _metni_kisalt(unit.get("ad"), 39)
            draw.text(
                (left + 12, top + 128),
                f"{code} — {name}" if code else name,
                fill="black",
                font=_font(19, bold=True),
            )
        images.append(
            {
                "ad": f"Lejant örnekleri {page_no}",
                "mime_type": "image/jpeg",
                "data": _jpeg_bytes(canvas, maximum=1900, quality=94),
            }
        )
        canvas.close()
    return images


def yapay_zeka_paketi_hazirla(analysis):
    """PIL nesnelerini ağ iş parçacığından önce bağımsız JPEG baytlarına çevirir."""
    harita = (analysis or {}).get("harita")
    if harita is None:
        raise JeolojiYapayZekaHatasi("Yapay zekâ denetimi için genel jeoloji haritası bulunamadı.")
    catalog_source = (analysis or {}).get("lejant_birimleri") or (analysis or {}).get("birimler") or []
    catalog = []
    seen = set()
    for unit in catalog_source:
        code = str(unit.get("kod") or "").strip()
        name = str(unit.get("ad") or "").strip()
        key = jeoloji_anahtari(code)
        if not code or not key or key in seen:
            continue
        seen.add(key)
        catalog.append(
            {
                "kod": code,
                "ad": name,
                "pafta_adlari": list(unit.get("pafta_adlari") or []),
                "onizleme": unit.get("onizleme"),
            }
        )
    if not catalog:
        raise JeolojiYapayZekaHatasi("Paftaya ait tanımlı lejant birimleri bulunamadı.")
    local = [
        {
            "kod": str(unit.get("kod") or "").strip(),
            "ad": str(unit.get("ad") or "").strip(),
            "oran": float(unit.get("oran") or 0),
            "guven": float(unit.get("guven") or 0),
            "ana_birim": bool(unit.get("ana_birim")),
        }
        for unit in (analysis or {}).get("birimler", [])
        if str(unit.get("kod") or "").strip()
        and not bool(unit.get("ai_eklenen", False))
        and bool(unit.get("yerel_aday", True))
    ]
    images = _harita_gorselleri(harita)
    images.extend(_lejant_sayfalari(catalog))
    digest = hashlib.sha256()
    digest.update(PROMPT_SURUMU.encode("utf-8"))
    digest.update(str((analysis or {}).get("geometri_hash", "")).encode("utf-8"))
    for image in images:
        digest.update(image["ad"].encode("utf-8"))
        digest.update(image["data"])
    digest.update(
        json.dumps(
            [{"kod": unit["kod"], "ad": unit["ad"]} for unit in catalog],
            ensure_ascii=False,
            sort_keys=True,
        ).encode("utf-8")
    )
    return {
        "hash": digest.hexdigest(),
        "gorseller": images,
        "katalog": [{key: value for key, value in unit.items() if key != "onizleme"} for unit in catalog],
        "yerel_birimler": local,
        "pafta_adlari": list((analysis or {}).get("pafta_adlari") or []),
    }


def _response_schema(codes):
    code_values = list(dict.fromkeys([""] + [str(code) for code in codes if str(code).strip()]))
    return {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "birimler": {
                "type": "array",
                "maxItems": min(40, max(1, len(code_values) - 1)),
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "kod": {"type": "string", "enum": code_values[1:]},
                        "haritada_goruldu": {"type": "boolean"},
                        "kanit": {
                            "type": "string",
                            "enum": ["kod_okundu", "kod_ve_desen", "yalniz_desen"],
                        },
                        "guven": {"type": "integer", "minimum": 0, "maximum": 100},
                        "konum_aciklamasi": {"type": "string"},
                    },
                    "required": ["kod", "haritada_goruldu", "kanit", "guven", "konum_aciklamasi"],
                },
            },
            "ana_parsel_kodu": {"type": "string", "enum": code_values},
            "genel_guven": {"type": "integer", "minimum": 0, "maximum": 100},
            "notlar": {"type": "string"},
        },
        "required": ["birimler", "ana_parsel_kodu", "genel_guven", "notlar"],
    }


def _prompt(package):
    catalog_lines = "\n".join(
        f"- {unit['kod']} = {unit.get('ad') or 'Ad belirtilmedi'}"
        for unit in package["katalog"]
    )
    local_lines = "\n".join(
        f"- {unit['kod']} ({unit.get('ad')}): örnek payı %{unit.get('oran', 0):g}, "
        f"yerel görsel güveni %{unit.get('guven', 0):g}, ana parsel={unit.get('ana_birim')}"
        for unit in package["yerel_birimler"]
    ) or "- Yerel aday yok"
    sheets = ", ".join(package.get("pafta_adlari") or []) or "adı belirtilmeyen pafta"
    return f"""Görev: MTA 1/100.000 jeoloji paftasının parsel merkezli 15 x 9 km kesitinde gerçekten görülen jeolojik birimleri denetle.

Görsellerde önce tüm harita, sonra dört büyütülmüş harita parçası, ardından tanımlı lejant örnek sayfaları vardır. Kırmızı-beyaz sınır ve ÇALIŞMA ALANI etiketi ana parseldir. Harita çevresindeki bütün birimleri bildir; yalnız parsel içini bildirmekle yetinme.

Zorunlu kurallar:
1. Yalnız aşağıdaki izinli kodlardan seçim yap. Yeni kod veya formasyon uydurma.
2. Haritadaki basılı birim kodunu okuyabiliyorsan kodu esas al ve kanıtı kod_okundu veya kod_ve_desen seç.
3. Kod okunamıyor, yalnız renk/tarama lejantla eşleşiyorsa kanıtı yalniz_desen seç ve güveni 65'i geçirme.
4. Yol, yerleşim, deniz, boş beyaz alan, koordinat yazısı ve sembolleri formasyon sayma.
5. Aynı kodu yalnız bir kez döndür. Haritada görülmeyen kodları listeye ekleme.
6. Yerel aday listesi yalnız zayıf bir ön tahmindir; yanlış olabilir ve cevabı ona göre zorlamamalısın.
7. Ana parsel kodu net değilse boş metin döndür.

Pafta(lar): {sheets}

İzinli kod ve birimler:
{catalog_lines}

Yerel görüntü örnekleme adayları:
{local_lines}
"""


def _hata_mesaji(body, fallback):
    try:
        parsed = json.loads(body)
        error = parsed.get("error", {}) if isinstance(parsed, dict) else {}
        message = error.get("message") or parsed.get("message")
        violations = []
        for detail in error.get("details", []) if isinstance(error, dict) else []:
            if not isinstance(detail, dict):
                continue
            for violation in detail.get("fieldViolations", []):
                if not isinstance(violation, dict):
                    continue
                field = str(violation.get("field") or "").strip()
                description = str(violation.get("description") or "").strip()
                if field and description:
                    violations.append(f"{field}: {description}")
                elif field or description:
                    violations.append(field or description)
        if message and violations:
            message = f"{message} ({'; '.join(violations[:3])})"
        if message:
            return _metni_kisalt(message, 500)
    except (TypeError, ValueError, json.JSONDecodeError):
        pass
    return fallback


GEMINI_YENIDEN_DENEME_DURUMLARI = {408, 429, 500, 502, 503, 504}
GEMINI_AZAMI_DENEME = 3


def _gemini_http_hatasi(code):
    if code == 400:
        return "Gönderilen Gemini isteği geçersiz (HTTP 400); model ve istek ayarlarını kontrol edin."
    if code in {401, 403}:
        return "Gemini API anahtarı kabul edilmedi; anahtar ve yetki ayarlarını kontrol edin."
    if code == 503:
        return (
            "Gemini hizmeti yanıtı zamanında tamamlayamadı (HTTP 503). "
            "Bağlantınızı kontrol edip biraz sonra yeniden deneyin."
        )
    return f"Gemini hizmeti isteği tamamlayamadı (HTTP {code}). Lütfen biraz sonra yeniden deneyin."


def _post_json(
    url,
    headers,
    payload,
    timeout,
    *,
    retry_attempts=1,
    retry_http_statuses=None,
    retry_label="API",
):
    request = urllib.request.Request(
        url,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={"Content-Type": "application/json", **headers},
        method="POST",
    )
    retry_http_statuses = set(retry_http_statuses or ())
    attempts = max(1, int(retry_attempts or 1))
    for attempt in range(1, attempts + 1):
        try:
            with urllib.request.urlopen(request, timeout=timeout) as response:
                body = response.read().decode("utf-8")
            break
        except urllib.error.HTTPError as exc:
            try:
                error_body = exc.read(6000).decode("utf-8", errors="replace")
            except Exception:
                error_body = ""
            if exc.code in retry_http_statuses and attempt < attempts:
                time.sleep(min(4.0, 2 ** (attempt - 1)))
                continue
            if retry_label == "Gemini":
                raise JeolojiYapayZekaHatasi(_gemini_http_hatasi(exc.code)) from None
            raise JeolojiYapayZekaHatasi(
                f"API HTTP {exc.code}: {_hata_mesaji(error_body, exc.reason or 'istek reddedildi')}"
            ) from None
        except (urllib.error.URLError, TimeoutError, socket.timeout) as exc:
            if attempt < attempts:
                time.sleep(min(4.0, 2 ** (attempt - 1)))
                continue
            if retry_label == "Gemini":
                raise JeolojiYapayZekaHatasi(
                    "Gemini hizmetine bağlanılamadı veya yanıt zaman aşımına uğradı. "
                    "Bağlantınızı kontrol edip biraz sonra yeniden deneyin."
                ) from None
            if isinstance(exc, urllib.error.URLError):
                reason = getattr(exc, "reason", exc)
                raise JeolojiYapayZekaHatasi(f"API bağlantısı kurulamadı: {reason}") from None
            raise JeolojiYapayZekaHatasi("Yapay zekâ isteği zaman aşımına uğradı.") from None
    try:
        return json.loads(body)
    except json.JSONDecodeError:
        raise JeolojiYapayZekaHatasi("API geçerli bir JSON yanıtı döndürmedi.") from None


def _json_metnini_oku(text):
    raw = str(text or "").strip()
    if raw.startswith("```"):
        raw = raw.split("\n", 1)[1] if "\n" in raw else raw
        raw = raw.rsplit("```", 1)[0].strip()
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        start, end = raw.find("{"), raw.rfind("}")
        if start >= 0 and end > start:
            try:
                return json.loads(raw[start:end + 1])
            except json.JSONDecodeError:
                pass
    raise JeolojiYapayZekaHatasi("Modelin yapılandırılmış cevabı okunamadı.")


def _gemini_cagir(package, key, model, timeout):
    schema = _response_schema([unit["kod"] for unit in package["katalog"]])
    parts = [{"text": _prompt(package)}]
    for image in package["gorseller"]:
        parts.append({"text": f"Görsel: {image['ad']}"})
        parts.append(
            {
                "inline_data": {
                    "mime_type": image["mime_type"],
                    "data": base64.b64encode(image["data"]).decode("ascii"),
                }
            }
        )
    payload = {
        "contents": [{"role": "user", "parts": parts}],
        "generationConfig": {
            # Gemini 3.6 ve sonraki modeller, kullanımdan kaldırılan sampling
            # alanları gönderildiğinde HTTP 400 döndürebilir. Denetim yalnız
            # yapılandırılmış çıktı alanlarını kullanır.
            "responseMimeType": "application/json",
            "responseJsonSchema": schema,
        },
    }
    response = _post_json(
        f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent",
        {"x-goog-api-key": key},
        payload,
        timeout,
        retry_attempts=GEMINI_AZAMI_DENEME,
        retry_http_statuses=GEMINI_YENIDEN_DENEME_DURUMLARI,
        retry_label="Gemini",
    )
    try:
        parts = response["candidates"][0]["content"]["parts"]
        text = "".join(part.get("text", "") for part in parts)
    except (KeyError, IndexError, TypeError):
        feedback = response.get("promptFeedback", {}) if isinstance(response, dict) else {}
        raise JeolojiYapayZekaHatasi(
            f"Gemini yanıtında sonuç bulunamadı. {feedback.get('blockReason', '')}".strip()
        ) from None
    return _json_metnini_oku(text)


def _openai_cagir(package, key, model, timeout):
    schema = _response_schema([unit["kod"] for unit in package["katalog"]])
    content = [{"type": "input_text", "text": _prompt(package)}]
    for image in package["gorseller"]:
        content.append({"type": "input_text", "text": f"Görsel: {image['ad']}"})
        data_url = (
            f"data:{image['mime_type']};base64,"
            + base64.b64encode(image["data"]).decode("ascii")
        )
        content.append(
            {"type": "input_image", "image_url": data_url, "detail": "original"}
        )
    payload = {
        "model": model,
        "input": [{"role": "user", "content": content}],
        "reasoning": {"effort": "medium"},
        "text": {
            "verbosity": "low",
            "format": {
                "type": "json_schema",
                "name": "jeoloji_formasyon_denetimi",
                "strict": True,
                "schema": schema,
            },
        },
        "max_output_tokens": 5000,
        "store": False,
    }
    response = _post_json(
        "https://api.openai.com/v1/responses",
        {"Authorization": f"Bearer {key}"},
        payload,
        timeout,
    )
    text = str(response.get("output_text") or "") if isinstance(response, dict) else ""
    if not text:
        for item in response.get("output", []) if isinstance(response, dict) else []:
            if item.get("type") != "message":
                continue
            for part in item.get("content", []):
                if part.get("type") in {"output_text", "text"} and part.get("text"):
                    text += str(part["text"])
    if not text:
        raise JeolojiYapayZekaHatasi("OpenAI yanıtında metin sonucu bulunamadı.")
    return _json_metnini_oku(text)


def _sonucu_dogrula(raw, catalog):
    code_map = {jeoloji_anahtari(unit["kod"]): unit for unit in catalog}
    units = {}
    for item in raw.get("birimler", []) if isinstance(raw, dict) else []:
        if not isinstance(item, dict) or not bool(item.get("haritada_goruldu")):
            continue
        key = jeoloji_anahtari(item.get("kod"))
        catalog_unit = code_map.get(key)
        if not catalog_unit:
            continue
        try:
            confidence = max(0, min(100, int(round(float(item.get("guven", 0))))))
        except (TypeError, ValueError):
            confidence = 0
        evidence = str(item.get("kanit") or "yalniz_desen")
        if evidence not in KANITLAR:
            evidence = "yalniz_desen"
        if evidence == "yalniz_desen":
            confidence = min(confidence, 65)
        normalized = {
            "kod": catalog_unit["kod"],
            "ad": catalog_unit.get("ad", ""),
            "kanit": evidence,
            "guven": confidence,
            "konum_aciklamasi": _metni_kisalt(item.get("konum_aciklamasi"), 300),
        }
        previous = units.get(key)
        if previous is None or normalized["guven"] > previous["guven"]:
            units[key] = normalized
    main_key = jeoloji_anahtari(raw.get("ana_parsel_kodu")) if isinstance(raw, dict) else ""
    main = code_map.get(main_key, {}).get("kod", "")
    try:
        general_confidence = max(0, min(100, int(round(float(raw.get("genel_guven", 0))))))
    except (AttributeError, TypeError, ValueError):
        general_confidence = 0
    return {
        "birimler": list(units.values()),
        "ana_parsel_kodu": main,
        "genel_guven": general_confidence,
        "notlar": _metni_kisalt(raw.get("notlar", "") if isinstance(raw, dict) else "", 700),
    }


class JeolojiYapayZekaServisi:
    """Gemini ve GPT-5.6 Sol'u yalnız seçilen düğme için çalıştırır."""

    def __init__(self, kullanici_veri_klasoru):
        self.ayarlar = JeolojiYapayZekaAyarlari(kullanici_veri_klasoru)

    def paket_hazirla(self, analysis):
        return yapay_zeka_paketi_hazirla(analysis)

    def saglayici_ile_analiz_et(self, package, saglayici, zorla_yenile=False):
        if saglayici not in SAGLAYICILAR:
            raise JeolojiYapayZekaHatasi("Bilinmeyen yapay zekâ sağlayıcısı seçildi.")
        settings = self.ayarlar.oku()
        key = self.ayarlar.anahtar_al(saglayici)
        if not key:
            raise YapayZekaAnahtariEksik(
                f"{SAGLAYICILAR[saglayici]['etiket']} API anahtarı bulunamadı. "
                "AI Ayarları bölümünden bu sağlayıcının anahtarını kaydedin."
            )
        model = (
            settings["birincil_model"]
            if saglayici == "gemini" else settings["ikinci_model"]
        )
        cache_identity = {
            "paket": package["hash"],
            "prompt": PROMPT_SURUMU,
            "saglayici": saglayici,
            "model": model,
        }
        cache_key = hashlib.sha256(
            json.dumps(cache_identity, sort_keys=True).encode("utf-8")
        ).hexdigest()
        cache_path = self.ayarlar.onbellek_klasoru / f"{cache_key}.json"
        if settings["onbellek_etkin"] and not zorla_yenile:
            try:
                cached = json.loads(cache_path.read_text(encoding="utf-8-sig"))
                if (
                    isinstance(cached, dict)
                    and cached.get("saglayici") == saglayici
                    and isinstance(cached.get("birimler"), list)
                ):
                    cached["onbellekten"] = True
                    return cached
            except (OSError, json.JSONDecodeError):
                pass

        timeout = settings["zaman_asimi_saniye"]
        if saglayici == "gemini":
            raw = _gemini_cagir(package, key, model, timeout)
        else:
            raw = _openai_cagir(package, key, model, timeout)
        normalized = _sonucu_dogrula(raw, package["katalog"])
        result = {
            "surum": 2,
            "prompt_surumu": PROMPT_SURUMU,
            "saglayici": saglayici,
            "saglayici_etiketi": SAGLAYICILAR[saglayici]["etiket"],
            "model": model,
            "birimler": normalized.get("birimler", []),
            "ana_parsel_kodu": normalized.get("ana_parsel_kodu", ""),
            "genel_guven": normalized.get("genel_guven", 0),
            "notlar": normalized.get("notlar", ""),
            "onbellekten": False,
        }
        if settings["onbellek_etkin"]:
            try:
                self.ayarlar.onbellek_klasoru.mkdir(parents=True, exist_ok=True)
                temporary = cache_path.with_suffix(".tmp")
                temporary.write_text(
                    json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
                )
                os.replace(temporary, cache_path)
            except OSError:
                pass
        return result

    def analiz_et(self, package, zorla_yenile=False, saglayici="gemini"):
        """Eski çağıranlar için uyumluluk; otomatik ikinci sağlayıcı çalıştırmaz."""
        return self.saglayici_ile_analiz_et(
            package, saglayici=saglayici, zorla_yenile=zorla_yenile
        )


__all__ = [
    "JeolojiYapayZekaAyarlari",
    "JeolojiYapayZekaHatasi",
    "JeolojiYapayZekaServisi",
    "YapayZekaAnahtariEksik",
    "yapay_zeka_paketi_hazirla",
]
