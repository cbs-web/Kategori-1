"""1/100.000 pafta birimleri için kalıcı yaş ve metin kütüphanesi."""

from __future__ import annotations

import datetime as dt
from contextlib import contextmanager
import json
from pathlib import Path
import re
import sqlite3

from jeoloji_kutuphanesi import jeoloji_anahtari, kullanici_yolu


YAS_SECENEKLERI = (
    ("Prekambriyen", 100),
    ("Paleozoyik", 200),
    ("Permiyen", 260),
    ("Triyas", 300),
    ("Jura", 400),
    ("Jura-Kretase", 450),
    ("Kretase", 500),
    ("Paleosen", 600),
    ("Eosen", 700),
    ("Oligosen", 800),
    ("Oligosen-Miyosen", 850),
    ("Alt Miyosen", 900),
    ("Orta Miyosen", 920),
    ("Üst Miyosen", 940),
    ("Miyosen", 950),
    ("Pliyosen", 1000),
    ("Pliyosen-Kuvaterner", 1050),
    ("Kuvaterner", 1100),
    ("Güncel", 1200),
    ("Yaşı kullanıcı tarafından belirlenecek", 9999),
)

YAS_SIRA_SOZLUGU = {ad: sira for ad, sira in YAS_SECENEKLERI}


def _simdi():
    return dt.datetime.now().isoformat(timespec="seconds")


def jeolojik_yas_tahmin_et(kod, ad):
    """Pafta kodu/birim adından yalnız güvenli ve genel bir yaş önerisi üret."""
    kod_ham = str(kod or "").strip()
    kod_key = jeoloji_anahtari(kod_ham).replace(" ", "")
    ad_key = jeoloji_anahtari(ad)
    combined = f"{kod_key} {ad_key}"

    anahtarlar = (
        (("guncel",), "Güncel"),
        (("kuvaterner", "aluvyon", "yamaç molozu", "teras"), "Kuvaterner"),
        (("pliyosen kuvaterner",), "Pliyosen-Kuvaterner"),
        (("pliyosen",), "Pliyosen"),
        (("ust miyosen",), "Üst Miyosen"),
        (("orta miyosen",), "Orta Miyosen"),
        (("alt miyosen",), "Alt Miyosen"),
        (("oligosen miyosen", "oligo miyosen"), "Oligosen-Miyosen"),
        (("miyosen",), "Miyosen"),
        (("oligosen",), "Oligosen"),
        (("eosen",), "Eosen"),
        (("paleosen",), "Paleosen"),
        (("jura kretase",), "Jura-Kretase"),
        (("kretase",), "Kretase"),
        (("jura",), "Jura"),
        (("triyas",), "Triyas"),
        (("permiyen",), "Permiyen"),
        (("paleozoyik",), "Paleozoyik"),
        (("prekambriyen",), "Prekambriyen"),
    )
    for kelimeler, yas in anahtarlar:
        if any(kelime in combined for kelime in kelimeler):
            return yas, YAS_SIRA_SOZLUGU[yas]

    # Türkiye 1/100.000 jeoloji paftalarında yaygın dönem önekleri. Bunlar yalnız
    # öneridir; arayüzde kullanıcı tarafından değiştirilebilir.
    lower = kod_ham.casefold()
    if lower.startswith("q"):
        yas = "Kuvaterner"
    elif lower.startswith("pl") or lower.startswith("tpl"):
        yas = "Pliyosen"
    elif lower.startswith("tr"):
        yas = "Triyas"
    elif lower.startswith("te") or lower.startswith("e"):
        yas = "Eosen"
    elif lower.startswith("t"):
        yas = "Miyosen"
    elif lower.startswith("k"):
        yas = "Kretase"
    elif lower.startswith("j"):
        yas = "Jura"
    elif lower.startswith("p"):
        yas = "Paleozoyik"
    else:
        yas = "Yaşı kullanıcı tarafından belirlenecek"
    return yas, YAS_SIRA_SOZLUGU[yas]


def birimleri_yasli_gence_sirala(birimler):
    """Aynı listeyi yaşlıdan gence; eşit yaşta pafta stratigrafisine göre sırala."""
    def anahtar(birim):
        try:
            yas_sirasi = int(birim.get("yas_sirasi", 9999))
        except (TypeError, ValueError):
            yas_sirasi = 9999
        try:
            kaynak_sirasi = int(birim.get("kaynak_sirasi", -1))
        except (TypeError, ValueError):
            kaynak_sirasi = -1
        # Pafta lejantları genel olarak gençten yaşlıya dizildiği için eşit/belirsiz
        # yaşta aşağıdaki (büyük sıra numaralı) birim önce gelir.
        return (yas_sirasi, -kaynak_sirasi, jeoloji_anahtari(birim.get("ad")))

    return sorted((dict(birim) for birim in birimler or ()), key=anahtar)


_BIRIM_BASLIGI_TERIMLERI = {
    "aluvyon",
    "bazalti",
    "formasyonu",
    "gnaysi",
    "graniti",
    "granitoidleri",
    "ignimbiriti",
    "kirectasi",
    "kompleksi",
    "kumtasi",
    "melanji",
    "mermer",
    "metamorfitleri",
    "molozu",
    "ofiyoliti",
    "riyoliti",
    "sisti",
    "traverten",
    "uyesi",
    "volkaniti",
}


def _kod_anahtari(value):
    return jeoloji_anahtari(value).replace(" ", "")


def _birim_adindan_kodu_ayir(value):
    text = str(value or "").strip()
    match = re.search(r"\(([^()]*)\)\s*$", text)
    code = match.group(1).strip() if match else ""
    name = text[: match.start()].strip(" :-") if match else text
    return name, code


def _2_1_sinir_satiri_mi(value):
    key = jeoloji_anahtari(value)
    if not key:
        return False
    if key.startswith(("sekil ", "tablo ", "cizelge ", "fotograf ", "figure ")):
        return True
    return bool(
        "2 1 1" in key
        or "yapisal jeoloji" in key
        or "aktif tektonik" in key
        or re.match(r"^2\s+[2-9](?:\s|$)", key)
    )


def _birim_basligini_ayir(value):
    """Bir paragraf birim başlığıysa ``(başlık, devam metni)`` döndür."""
    text = " ".join(str(value or "").strip().split())
    if not text or _2_1_sinir_satiri_mi(text):
        return None

    heading = text
    continuation = ""
    if ":" in text:
        candidate, tail = text.split(":", 1)
        candidate_key = jeoloji_anahtari(candidate)
        if any(term in candidate_key.split() for term in _BIRIM_BASLIGI_TERIMLERI):
            heading = candidate.strip()
            continuation = tail.strip()

    heading_key = jeoloji_anahtari(heading)
    words = heading_key.split()
    if not words or len(heading) > 180 or len(words) > 16:
        return None
    has_unit_term = any(term in words for term in _BIRIM_BASLIGI_TERIMLERI)
    _name, code = _birim_adindan_kodu_ayir(heading)
    if not has_unit_term or (not code and len(words) > 7):
        return None
    return heading, continuation


def _baslik_birim_anahtari(heading, targets):
    heading_name, heading_code = _birim_adindan_kodu_ayir(heading)
    heading_name_key = jeoloji_anahtari(heading_name)
    heading_code_key = _kod_anahtari(heading_code)
    if heading_code_key:
        for target in targets:
            if target["kod_key"] and target["kod_key"] == heading_code_key:
                return target["key"]
    for target in targets:
        if target["ad_key"] and target["ad_key"] == heading_name_key:
            return target["key"]
    return None


def eski_metin_birimlere_dagit(metin, birimler):
    """Eski 2.1 metnini birim başlıkları arasındaki gerçek gövdelere ayır.

    Önceki yöntem, birimin adı başka bir formasyonun açıklamasında geçtiğinde o
    paragrafı da yanlış birime ekliyordu. Burada yalnız başlıkla başlayan bölüm,
    bir sonraki birim/şekil/2.1.1 başlığına kadar alınır.
    """
    targets = []
    result = {}
    for unit in birimler or ():
        code = str(unit.get("kod") or "").strip()
        name, name_code = _birim_adindan_kodu_ayir(unit.get("ad"))
        code = code or name_code
        key = (jeoloji_anahtari(code), jeoloji_anahtari(str(unit.get("ad") or "")))
        targets.append(
            {
                "key": key,
                "kod_key": _kod_anahtari(code),
                "ad_key": jeoloji_anahtari(name),
            }
        )
        result[key] = []

    active_key = None
    paragraphs = [
        paragraph.strip()
        for paragraph in re.split(r"(?:\r?\n)+", str(metin or ""))
        if paragraph.strip()
    ]
    for paragraph in paragraphs:
        if _2_1_sinir_satiri_mi(paragraph):
            active_key = None
            continue
        heading = _birim_basligini_ayir(paragraph)
        if heading is not None:
            heading_text, continuation = heading
            active_key = _baslik_birim_anahtari(heading_text, targets)
            if active_key is not None and continuation:
                result[active_key].append(continuation)
            continue
        if active_key is not None:
            cleaned = re.sub(
                r"Hata!\s*Yer işareti başvurusu geçersiz\.?",
                "",
                paragraph,
                flags=re.IGNORECASE,
            ).strip()
            if cleaned:
                result[active_key].append(cleaned)

    return {key: "\n\n".join(parts).strip() for key, parts in result.items()}


class FormasyonMetinKutuphanesi:
    """Jeoloji birimi metinlerini ana rapor kütüphanesiyle aynı SQLite'ta tutar."""

    def __init__(self, db_path=None):
        self.db_path = db_path or kullanici_yolu("jeoloji", "canakkale_jeoloji.db")
        self._sema_olustur()

    @contextmanager
    def _baglan(self):
        connection = sqlite3.connect(self.db_path)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        connection.execute("PRAGMA busy_timeout = 5000")
        try:
            yield connection
            connection.commit()
        except Exception:
            connection.rollback()
            raise
        finally:
            connection.close()

    def _sema_olustur(self):
        with self._baglan() as connection:
            connection.executescript(
                """
                CREATE TABLE IF NOT EXISTS formasyon_metinleri (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    birim_kodu TEXT NOT NULL DEFAULT '',
                    birim_adi TEXT NOT NULL,
                    kod_key TEXT NOT NULL DEFAULT '',
                    ad_key TEXT NOT NULL,
                    jeolojik_yas TEXT NOT NULL DEFAULT '',
                    yas_sirasi INTEGER NOT NULL DEFAULT 9999,
                    lejant_aciklamasi TEXT NOT NULL DEFAULT '',
                    bolgesel_jeoloji_metni TEXT NOT NULL DEFAULT '',
                    kaynak_notu TEXT NOT NULL DEFAULT '',
                    revizyon_no INTEGER NOT NULL DEFAULT 1,
                    aktif INTEGER NOT NULL DEFAULT 1,
                    olusturma_tarihi TEXT NOT NULL,
                    guncelleme_tarihi TEXT NOT NULL,
                    UNIQUE(kod_key, ad_key)
                );

                CREATE TABLE IF NOT EXISTS formasyon_metni_revizyonlari (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    formasyon_id INTEGER NOT NULL,
                    revizyon_no INTEGER NOT NULL,
                    anlik_goruntu TEXT NOT NULL,
                    kayit_tarihi TEXT NOT NULL,
                    FOREIGN KEY(formasyon_id) REFERENCES formasyon_metinleri(id) ON DELETE CASCADE,
                    UNIQUE(formasyon_id, revizyon_no)
                );

                CREATE UNIQUE INDEX IF NOT EXISTS uq_formasyon_metin_kodu
                ON formasyon_metinleri(kod_key)
                WHERE kod_key <> '' AND aktif = 1;
                """
            )

    @staticmethod
    def _satir(row):
        return dict(row) if row is not None else None

    def getir(self, kod="", ad=""):
        kod_key = jeoloji_anahtari(kod)
        ad_key = jeoloji_anahtari(ad)
        with self._baglan() as connection:
            row = connection.execute(
                """
                SELECT * FROM formasyon_metinleri
                WHERE aktif = 1 AND (
                    (kod_key <> '' AND kod_key = ?) OR ad_key = ?
                )
                ORDER BY CASE WHEN kod_key = ? AND kod_key <> '' THEN 0 ELSE 1 END, id
                LIMIT 1
                """,
                (kod_key, ad_key, kod_key),
            ).fetchone()
        return self._satir(row)

    def listele(self):
        with self._baglan() as connection:
            rows = connection.execute(
                """
                SELECT * FROM formasyon_metinleri WHERE aktif = 1
                ORDER BY yas_sirasi, ad_key
                """
            ).fetchall()
        return [self._satir(row) for row in rows]

    @staticmethod
    def _birim_anahtari(unit):
        code = str(unit.get("kod") or "").strip()
        if not code:
            _name, code = _birim_adindan_kodu_ayir(unit.get("ad"))
        return (
            jeoloji_anahtari(code),
            jeoloji_anahtari(unit.get("ad")),
        )

    def eski_jeoloji_wordlerinden_tamamla(self, birimler):
        """Boş 2.1 birim metinlerini onaylı eski Word kayıtlarından doldur.

        Kaynak Word'lerin 2.1 metni daha önce ``jeoloji_kayitlari`` tablosuna
        çıkarıldığı için DOCX dosyaları her harita açılışında yeniden okunmaz.
        Kullanıcının elle kaydettiği dolu metinlere kesinlikle dokunulmaz.
        """
        missing = []
        for unit in birimler or ():
            code = str(unit.get("kod") or "").strip()
            name = str(unit.get("ad") or "").strip()
            if not name:
                continue
            current = self.getir(code, name) or {}
            if str(current.get("bolgesel_jeoloji_metni") or "").strip():
                continue
            candidate = dict(unit)
            candidate.update({"kod": code, "ad": name, "current": current})
            missing.append(candidate)
        if not missing:
            return {"aktarilan": [], "bulunamayan": []}

        with self._baglan() as connection:
            table = connection.execute(
                "SELECT 1 FROM sqlite_master WHERE type='table' AND name='jeoloji_kayitlari'"
            ).fetchone()
            if table is None:
                return {
                    "aktarilan": [],
                    "bulunamayan": [unit["ad"] for unit in missing],
                }
            sources = connection.execute(
                """
                SELECT id, formasyon, genel_jeoloji_metni, kaynak_rapor_path,
                       guncelleme_tarihi
                FROM jeoloji_kayitlari
                WHERE aktif = 1 AND onay_durumu = 'onayli'
                  AND TRIM(genel_jeoloji_metni) <> ''
                ORDER BY guncelleme_tarihi DESC, id DESC
                """
            ).fetchall()

        candidates = {self._birim_anahtari(unit): [] for unit in missing}
        for source in sources:
            distributed = eski_metin_birimlere_dagit(
                source["genel_jeoloji_metni"], missing
            )
            source_formation_key = jeoloji_anahtari(source["formasyon"])
            for unit in missing:
                unit_key = self._birim_anahtari(unit)
                text = str(distributed.get(unit_key) or "").strip()
                if len(text) < 80:
                    continue
                unit_name, _unit_code = _birim_adindan_kodu_ayir(unit["ad"])
                name_key = jeoloji_anahtari(unit_name)
                main_formation_match = int(
                    bool(name_key and name_key in source_formation_key)
                )
                candidates[unit_key].append(
                    (
                        main_formation_match,
                        len(text),
                        str(source["guncelleme_tarihi"] or ""),
                        int(source["id"]),
                        text,
                        str(source["kaynak_rapor_path"] or ""),
                    )
                )

        transferred = []
        not_found = []
        for unit in missing:
            unit_key = self._birim_anahtari(unit)
            options = candidates.get(unit_key, [])
            if not options:
                not_found.append(unit["ad"])
                continue
            _main, _length, _date, record_id, text, source_path = max(options)
            current = unit.get("current") or {}
            age = (
                current.get("jeolojik_yas")
                or unit.get("jeolojik_yas")
                or jeolojik_yas_tahmin_et(unit.get("kod"), unit.get("ad"))[0]
            )
            age_order = current.get("yas_sirasi", unit.get("yas_sirasi"))
            source_name = Path(source_path).name if source_path else f"kayıt #{record_id}"
            automatic_note = (
                f"Jeoloji kütüphanesi kayıt #{record_id} içindeki 2.1 bölümünden "
                f"otomatik ayıklandı: {source_name}"
            )
            previous_note = str(current.get("kaynak_notu") or "").strip()
            self.kaydet(
                kod=unit.get("kod", ""),
                ad=unit.get("ad", ""),
                jeolojik_yas=age,
                yas_sirasi=age_order,
                lejant_aciklamasi=(
                    current.get("lejant_aciklamasi")
                    or unit.get("lejant_aciklamasi", "")
                ),
                bolgesel_jeoloji_metni=text,
                kaynak_notu=(
                    f"{previous_note}\n{automatic_note}".strip()
                    if previous_note else automatic_note
                ),
            )
            transferred.append(
                {"kod": unit.get("kod", ""), "ad": unit["ad"], "kayit_id": record_id}
            )
        return {"aktarilan": transferred, "bulunamayan": not_found}

    def kaydet(
        self,
        *,
        kod,
        ad,
        jeolojik_yas,
        yas_sirasi=None,
        lejant_aciklamasi="",
        bolgesel_jeoloji_metni="",
        kaynak_notu="",
    ):
        kod = str(kod or "").strip()
        ad = str(ad or "").strip()
        if not ad:
            raise ValueError("Formasyon/birim adı zorunludur.")
        jeolojik_yas = str(jeolojik_yas or "").strip()
        if not jeolojik_yas:
            jeolojik_yas, tahmin_sirasi = jeolojik_yas_tahmin_et(kod, ad)
        else:
            tahmin_sirasi = YAS_SIRA_SOZLUGU.get(jeolojik_yas, 9999)
        try:
            yas_sirasi = int(yas_sirasi if yas_sirasi is not None else tahmin_sirasi)
        except (TypeError, ValueError):
            yas_sirasi = tahmin_sirasi
        now = _simdi()
        kod_key = jeoloji_anahtari(kod)
        ad_key = jeoloji_anahtari(ad)
        with self._baglan() as connection:
            current = connection.execute(
                """
                SELECT * FROM formasyon_metinleri
                WHERE aktif = 1 AND ((kod_key <> '' AND kod_key = ?) OR ad_key = ?)
                ORDER BY CASE WHEN kod_key = ? AND kod_key <> '' THEN 0 ELSE 1 END, id
                LIMIT 1
                """,
                (kod_key, ad_key, kod_key),
            ).fetchone()
            if current is None:
                cursor = connection.execute(
                    """
                    INSERT INTO formasyon_metinleri(
                        birim_kodu, birim_adi, kod_key, ad_key, jeolojik_yas,
                        yas_sirasi, lejant_aciklamasi, bolgesel_jeoloji_metni,
                        kaynak_notu, revizyon_no, aktif, olusturma_tarihi, guncelleme_tarihi
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 1, 1, ?, ?)
                    """,
                    (
                        kod, ad, kod_key, ad_key, jeolojik_yas, yas_sirasi,
                        str(lejant_aciklamasi or "").strip(),
                        str(bolgesel_jeoloji_metni or "").strip(),
                        str(kaynak_notu or "").strip(), now, now,
                    ),
                )
                record_id = int(cursor.lastrowid)
                revision = 1
            else:
                record_id = int(current["id"])
                revision = int(current["revizyon_no"]) + 1
                connection.execute(
                    """
                    UPDATE formasyon_metinleri SET
                        birim_kodu = ?, birim_adi = ?, kod_key = ?, ad_key = ?,
                        jeolojik_yas = ?, yas_sirasi = ?,
                        lejant_aciklamasi = ?, bolgesel_jeoloji_metni = ?, kaynak_notu = ?,
                        revizyon_no = ?, aktif = 1, guncelleme_tarihi = ?
                    WHERE id = ?
                    """,
                    (
                        kod, ad, kod_key, ad_key, jeolojik_yas, yas_sirasi,
                        str(lejant_aciklamasi or "").strip(),
                        str(bolgesel_jeoloji_metni or "").strip(),
                        str(kaynak_notu or "").strip(), revision, now, record_id,
                    ),
                )
            row = connection.execute(
                "SELECT * FROM formasyon_metinleri WHERE id = ?", (record_id,)
            ).fetchone()
            snapshot = dict(row)
            connection.execute(
                """
                INSERT OR REPLACE INTO formasyon_metni_revizyonlari(
                    formasyon_id, revizyon_no, anlik_goruntu, kayit_tarihi
                ) VALUES (?, ?, ?, ?)
                """,
                (record_id, revision, json.dumps(snapshot, ensure_ascii=False), now),
            )
        return snapshot


__all__ = [
    "FormasyonMetinKutuphanesi",
    "YAS_SECENEKLERI",
    "YAS_SIRA_SOZLUGU",
    "birimleri_yasli_gence_sirala",
    "eski_metin_birimlere_dagit",
    "jeolojik_yas_tahmin_et",
]
