"""Çanakkale jeoloji metinleri ve haritaları için yerel SQLite kütüphanesi."""

from __future__ import annotations

import datetime as dt
from contextlib import contextmanager
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import sqlite3
import tempfile
import unicodedata
from zipfile import BadZipFile, ZIP_DEFLATED, ZipFile

from jeoloji_bolum_paketi import JeolojiBolumPaketiHatasi, jeoloji_bolumunu_ayir


def kullanici_yolu(*parcalar):
    """KATEGORI_1 için kullanıcı verisi yolunu üretir."""
    kok = (
        os.environ.get("LOCALAPPDATA")
        or os.environ.get("APPDATA")
        or tempfile.gettempdir()
    )
    klasor = Path(kok) / "K-1"
    klasor.mkdir(parents=True, exist_ok=True)
    return klasor.joinpath(*parcalar)


ONAY_DURUMLARI = ("taslak", "onayli")


class JeolojiKutuphanesiHatasi(ValueError):
    """Kütüphane girdisi eksik veya geçersiz olduğunda kullanılır."""


class AyniJeolojiKaydiHatasi(JeolojiKutuphanesiHatasi):
    """Aynı konum ve formasyon için ikinci etkin kayıt açılmasını engeller."""

    def __init__(self, kayit_id):
        self.kayit_id = int(kayit_id)
        super().__init__(
            f"Aynı ilçe, yerleşim, ada, parsel ve formasyon için {kayit_id} numaralı kayıt zaten var."
        )


def jeoloji_anahtari(value):
    """Türkçe karakter ve yazım farklarına dayanıklı karşılaştırma anahtarı üret."""
    text = unicodedata.normalize("NFKD", str(value or "").strip().casefold())
    text = "".join(char for char in text if not unicodedata.combining(char))
    text = text.replace("ı", "i")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def _temiz(value):
    return str(value or "").strip()


def _simdi():
    return dt.datetime.now().isoformat(timespec="seconds")


class JeolojiKutuphanesi:
    ALANLAR = (
        "il",
        "ilce",
        "yerlesim",
        "ada",
        "parsel",
        "formasyon",
        "genel_jeoloji_metni",
        "inceleme_alani_jeolojisi",
        "bolum_docx_path",
        "bolum_hash",
        "harita_path",
        "harita_aciklamasi",
        "harita_kaynagi",
        "harita_olcegi",
        "kaynak_rapor_path",
        "kaynak_klasor_path",
        "kaynak_rapor_hash",
        "kunye_kaynaklari_json",
        "kunye_duzeltme_notu",
        "notlar",
        "onay_durumu",
    )

    def __init__(self, db_path=None):
        if db_path is None:
            db_path = kullanici_yolu("jeoloji", "canakkale_jeoloji.db")
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self.harita_dizini = self.db_path.parent / "haritalar"
        self.bolum_dizini = self.db_path.parent / "jeoloji_bolumleri"
        self.kml_dizini = self.db_path.parent / "parsel_kml"
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
                CREATE TABLE IF NOT EXISTS jeoloji_kayitlari (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    il TEXT NOT NULL,
                    ilce TEXT NOT NULL,
                    yerlesim TEXT NOT NULL DEFAULT '',
                    ada TEXT NOT NULL DEFAULT '',
                    parsel TEXT NOT NULL DEFAULT '',
                    formasyon TEXT NOT NULL DEFAULT '',
                    il_key TEXT NOT NULL,
                    ilce_key TEXT NOT NULL,
                    yerlesim_key TEXT NOT NULL,
                    ada_key TEXT NOT NULL DEFAULT '',
                    parsel_key TEXT NOT NULL DEFAULT '',
                    formasyon_key TEXT NOT NULL,
                    genel_jeoloji_metni TEXT NOT NULL DEFAULT '',
                    inceleme_alani_jeolojisi TEXT NOT NULL DEFAULT '',
                    bolum_docx_path TEXT NOT NULL DEFAULT '',
                    bolum_hash TEXT NOT NULL DEFAULT '',
                    harita_path TEXT NOT NULL DEFAULT '',
                    harita_aciklamasi TEXT NOT NULL DEFAULT '',
                    harita_kaynagi TEXT NOT NULL DEFAULT '',
                    harita_olcegi TEXT NOT NULL DEFAULT '',
                    kaynak_rapor_path TEXT NOT NULL DEFAULT '',
                    kaynak_klasor_path TEXT NOT NULL DEFAULT '',
                    kaynak_rapor_hash TEXT NOT NULL DEFAULT '',
                    kunye_kaynaklari_json TEXT NOT NULL DEFAULT '',
                    kunye_duzeltme_notu TEXT NOT NULL DEFAULT '',
                    notlar TEXT NOT NULL DEFAULT '',
                    onay_durumu TEXT NOT NULL DEFAULT 'taslak',
                    revizyon_no INTEGER NOT NULL DEFAULT 1,
                    aktif INTEGER NOT NULL DEFAULT 1,
                    olusturma_tarihi TEXT NOT NULL,
                    guncelleme_tarihi TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS jeoloji_revizyonlari (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    kayit_id INTEGER NOT NULL,
                    revizyon_no INTEGER NOT NULL,
                    anlik_goruntu TEXT NOT NULL,
                    kayit_tarihi TEXT NOT NULL,
                    FOREIGN KEY(kayit_id) REFERENCES jeoloji_kayitlari(id) ON DELETE CASCADE,
                    UNIQUE(kayit_id, revizyon_no)
                );

                CREATE TABLE IF NOT EXISTS jeoloji_geometrileri (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    kayit_id INTEGER NOT NULL,
                    sira INTEGER NOT NULL,
                    placemark_adi TEXT NOT NULL DEFAULT '',
                    aciklama TEXT NOT NULL DEFAULT '',
                    noktalar_json TEXT NOT NULL,
                    min_enlem REAL NOT NULL,
                    max_enlem REAL NOT NULL,
                    min_boylam REAL NOT NULL,
                    max_boylam REAL NOT NULL,
                    merkez_enlem REAL NOT NULL,
                    merkez_boylam REAL NOT NULL,
                    kml_path TEXT NOT NULL DEFAULT '',
                    kml_hash TEXT NOT NULL DEFAULT '',
                    olusturma_tarihi TEXT NOT NULL,
                    FOREIGN KEY(kayit_id) REFERENCES jeoloji_kayitlari(id) ON DELETE CASCADE,
                    UNIQUE(kayit_id, sira)
                );
                """
            )
            existing_columns = {
                row["name"]
                for row in connection.execute("PRAGMA table_info(jeoloji_kayitlari)").fetchall()
            }
            for column in (
                "ada",
                "parsel",
                "ada_key",
                "parsel_key",
                "bolum_docx_path",
                "bolum_hash",
                "kaynak_klasor_path",
                "kaynak_rapor_hash",
                "kunye_kaynaklari_json",
                "kunye_duzeltme_notu",
            ):
                if column not in existing_columns:
                    connection.execute(
                        f"ALTER TABLE jeoloji_kayitlari ADD COLUMN {column} TEXT NOT NULL DEFAULT ''"
                    )
            connection.executescript(
                """
                CREATE UNIQUE INDEX IF NOT EXISTS uq_jeoloji_etkin_konum_v2
                ON jeoloji_kayitlari(
                    il_key, ilce_key, yerlesim_key, ada_key, parsel_key, formasyon_key
                ) WHERE aktif = 1;

                CREATE INDEX IF NOT EXISTS ix_jeoloji_arama_v2
                ON jeoloji_kayitlari(
                    il_key, ilce_key, yerlesim_key, ada_key, parsel_key,
                    formasyon_key, onay_durumu, aktif
                );

                CREATE INDEX IF NOT EXISTS ix_jeoloji_geometri_bbox
                ON jeoloji_geometrileri(min_boylam, max_boylam, min_enlem, max_enlem);

                CREATE INDEX IF NOT EXISTS ix_jeoloji_geometri_kayit
                ON jeoloji_geometrileri(kayit_id);

                DROP INDEX IF EXISTS uq_jeoloji_etkin_konum;
                DROP INDEX IF EXISTS ix_jeoloji_arama;
                """
            )

    def _normalize_kayit(self, kayit):
        result = {field: _temiz((kayit or {}).get(field)) for field in self.ALANLAR}
        result["il"] = result["il"] or "Çanakkale"
        if not result["ilce"]:
            raise JeolojiKutuphanesiHatasi("İlçe zorunludur.")
        if result["onay_durumu"] not in ONAY_DURUMLARI:
            raise JeolojiKutuphanesiHatasi("Onay durumu 'taslak' veya 'onayli' olmalıdır.")
        if not any(
            result[key]
            for key in (
                "genel_jeoloji_metni",
                "inceleme_alani_jeolojisi",
                "bolum_docx_path",
                "harita_path",
            )
        ):
            raise JeolojiKutuphanesiHatasi(
                "JEOLOJİ bölümü, inceleme alanı jeolojisi veya harita bilgilerinden en az biri girilmelidir."
            )
        for field in ("il", "ilce", "yerlesim", "ada", "parsel", "formasyon"):
            result[f"{field}_key"] = jeoloji_anahtari(result[field])
        if not result["kaynak_rapor_hash"] and result["kaynak_rapor_path"]:
            source = Path(result["kaynak_rapor_path"])
            try:
                if source.is_file():
                    result["kaynak_rapor_hash"] = self.kaynak_dosya_hashi(source)
            except (OSError, ValueError, BadZipFile):
                pass
        return result

    @staticmethod
    def _dosya_hashi(path):
        hasher = hashlib.sha256()
        with Path(path).open("rb") as stream:
            for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                hasher.update(chunk)
        return hasher.hexdigest()

    @staticmethod
    def _docx_icerik_hashi(path):
        """ZIP zaman damgalarından etkilenmeyen, DOCX parça içerik hash'i üret."""
        hasher = hashlib.sha256()
        with ZipFile(path, "r") as archive:
            for name in sorted(archive.namelist()):
                if name.endswith("/"):
                    continue
                hasher.update(name.encode("utf-8"))
                hasher.update(b"\0")
                hasher.update(archive.read(name))
                hasher.update(b"\0")
        return hasher.hexdigest()

    def kaynak_dosya_hashi(self, path):
        """Kaynak rapor için DOCX içeriğine dayanıklı veya normal SHA-256 üretir."""
        source = Path(path)
        if source.suffix.lower() == ".docx":
            return self._docx_icerik_hashi(source)
        return self._dosya_hashi(source)

    def bolum_dosyasi_ekle(self, source_path):
        """Bir raporun JEOLOJİ bölümünü kütüphaneye içerik-adresli DOCX olarak al."""
        source = Path(source_path)
        if source.suffix.lower() != ".docx" or not source.is_file():
            raise JeolojiKutuphanesiHatasi("Seçilen JEOLOJİ kaynak Word dosyası bulunamadı.")
        self.bolum_dizini.mkdir(parents=True, exist_ok=True)
        try:
            with tempfile.TemporaryDirectory(
                prefix="jeoloji_bolumu_", dir=self.bolum_dizini
            ) as temp_dir:
                temporary_package = Path(temp_dir) / "bolum.docx"
                result = jeoloji_bolumunu_ayir(source, temporary_package)
                digest = self._docx_icerik_hashi(temporary_package)
                target = self.bolum_dizini / f"{digest[:24]}.docx"
                if not target.exists():
                    os.replace(temporary_package, target)
        except (OSError, JeolojiBolumPaketiHatasi) as exc:
            raise JeolojiKutuphanesiHatasi(
                f"JEOLOJİ bölümü Word paketi oluşturulamadı: {exc}"
            ) from exc
        return str(target), digest, result

    def _bolum_paketini_hazirla(self, kayit):
        prepared = dict(kayit or {})
        package_text = _temiz(prepared.get("bolum_docx_path"))
        source_text = _temiz(prepared.get("kaynak_rapor_path"))
        package_path = Path(package_text) if package_text else None
        library_root = self.bolum_dizini.resolve()
        package_is_stored = False
        if package_path is not None and package_path.is_file():
            try:
                package_is_stored = package_path.resolve().is_relative_to(library_root)
            except (OSError, ValueError):
                package_is_stored = False
        if package_is_stored:
            prepared["bolum_docx_path"] = str(package_path.resolve())
            prepared["bolum_hash"] = _temiz(prepared.get("bolum_hash")) or self._docx_icerik_hashi(
                package_path
            )
            return prepared

        source = package_path if package_path is not None and package_path.is_file() else Path(source_text)
        should_extract = bool(package_text or _temiz(prepared.get("genel_jeoloji_metni")))
        if should_extract and source_text and source.suffix.lower() == ".docx" and source.is_file():
            stored_path, digest, _result = self.bolum_dosyasi_ekle(source)
            prepared["bolum_docx_path"] = stored_path
            prepared["bolum_hash"] = digest
        elif should_extract and package_path is not None and package_path.is_file():
            stored_path, digest, _result = self.bolum_dosyasi_ekle(package_path)
            prepared["bolum_docx_path"] = stored_path
            prepared["bolum_hash"] = digest
        return prepared

    @staticmethod
    def _satir_sozluk(row):
        if row is None:
            return None
        result = dict(row)
        result["aktif"] = bool(result.get("aktif"))
        return result

    def _ayni_kayit_id(self, connection, kayit, haric_id=None):
        params = [
            kayit["il_key"],
            kayit["ilce_key"],
            kayit["yerlesim_key"],
            kayit["ada_key"],
            kayit["parsel_key"],
            kayit["formasyon_key"],
        ]
        sql = (
            "SELECT id FROM jeoloji_kayitlari "
            "WHERE aktif = 1 AND il_key = ? AND ilce_key = ? AND yerlesim_key = ? "
            "AND ada_key = ? AND parsel_key = ? AND formasyon_key = ?"
        )
        if haric_id is not None:
            sql += " AND id <> ?"
            params.append(int(haric_id))
        row = connection.execute(sql, params).fetchone()
        return int(row["id"]) if row else None

    def _revizyon_yaz(self, connection, kayit_id, revizyon_no):
        row = connection.execute(
            "SELECT * FROM jeoloji_kayitlari WHERE id = ?", (int(kayit_id),)
        ).fetchone()
        snapshot = self._satir_sozluk(row)
        connection.execute(
            """
            INSERT INTO jeoloji_revizyonlari(kayit_id, revizyon_no, anlik_goruntu, kayit_tarihi)
            VALUES (?, ?, ?, ?)
            """,
            (
                int(kayit_id),
                int(revizyon_no),
                json.dumps(snapshot, ensure_ascii=False, sort_keys=True),
                _simdi(),
            ),
        )

    def kaydet(self, kayit, kayit_id=None):
        normalized = self._normalize_kayit(self._bolum_paketini_hazirla(kayit))
        now = _simdi()
        with self._baglan() as connection:
            duplicate_id = self._ayni_kayit_id(connection, normalized, haric_id=kayit_id)
            if duplicate_id is not None:
                raise AyniJeolojiKaydiHatasi(duplicate_id)

            if kayit_id is None:
                columns = (
                    *self.ALANLAR,
                    "il_key",
                    "ilce_key",
                    "yerlesim_key",
                    "ada_key",
                    "parsel_key",
                    "formasyon_key",
                )
                values = [normalized[column] for column in columns]
                placeholders = ", ".join("?" for _ in columns)
                cursor = connection.execute(
                    f"""
                    INSERT INTO jeoloji_kayitlari(
                        {', '.join(columns)}, revizyon_no, aktif, olusturma_tarihi, guncelleme_tarihi
                    ) VALUES ({placeholders}, 1, 1, ?, ?)
                    """,
                    (*values, now, now),
                )
                saved_id = int(cursor.lastrowid)
                self._revizyon_yaz(connection, saved_id, 1)
                return saved_id

            existing = connection.execute(
                "SELECT * FROM jeoloji_kayitlari WHERE id = ? AND aktif = 1", (int(kayit_id),)
            ).fetchone()
            if existing is None:
                raise JeolojiKutuphanesiHatasi("Güncellenecek etkin kayıt bulunamadı.")
            revision = int(existing["revizyon_no"]) + 1
            columns = (
                *self.ALANLAR,
                "il_key",
                "ilce_key",
                "yerlesim_key",
                "ada_key",
                "parsel_key",
                "formasyon_key",
            )
            assignments = ", ".join(f"{column} = ?" for column in columns)
            values = [normalized[column] for column in columns]
            connection.execute(
                f"""
                UPDATE jeoloji_kayitlari
                SET {assignments}, revizyon_no = ?, guncelleme_tarihi = ?
                WHERE id = ?
                """,
                (*values, revision, now, int(kayit_id)),
            )
            self._revizyon_yaz(connection, int(kayit_id), revision)
            return int(kayit_id)

    def getir(self, kayit_id, aktif_olmayan=False):
        sql = (
            "SELECT jeoloji_kayitlari.*, "
            "(SELECT COUNT(*) FROM jeoloji_geometrileri "
            " WHERE kayit_id = jeoloji_kayitlari.id) AS geometri_sayisi "
            "FROM jeoloji_kayitlari WHERE id = ?"
        )
        if not aktif_olmayan:
            sql += " AND aktif = 1"
        with self._baglan() as connection:
            row = connection.execute(sql, (int(kayit_id),)).fetchone()
        return self._satir_sozluk(row)

    def listele(
        self,
        *,
        il="",
        ilce="",
        yerlesim="",
        ada="",
        parsel="",
        formasyon="",
        onay_durumu="",
        arama="",
        aktif=True,
    ):
        clauses = ["aktif = ?"]
        params = [1 if aktif else 0]
        for column, value in (
            ("il_key", il),
            ("ilce_key", ilce),
            ("yerlesim_key", yerlesim),
            ("ada_key", ada),
            ("parsel_key", parsel),
            ("formasyon_key", formasyon),
        ):
            if _temiz(value):
                clauses.append(f"{column} = ?")
                params.append(jeoloji_anahtari(value))
        if onay_durumu in ONAY_DURUMLARI:
            clauses.append("onay_durumu = ?")
            params.append(onay_durumu)
        if _temiz(arama):
            clauses.append(
                "(il_key LIKE ? OR ilce_key LIKE ? OR yerlesim_key LIKE ? OR ada_key LIKE ? "
                "OR parsel_key LIKE ? OR formasyon_key LIKE ? "
                "OR genel_jeoloji_metni LIKE ? OR inceleme_alani_jeolojisi LIKE ?)"
            )
            key = jeoloji_anahtari(arama)
            raw = f"%{_temiz(arama)}%"
            params.extend(
                (
                    f"%{key}%",
                    f"%{key}%",
                    f"%{key}%",
                    f"%{key}%",
                    f"%{key}%",
                    f"%{key}%",
                    raw,
                    raw,
                )
            )
        where = " AND ".join(clauses)
        with self._baglan() as connection:
            rows = connection.execute(
                f"""
                SELECT jeoloji_kayitlari.*,
                       (SELECT COUNT(*) FROM jeoloji_geometrileri
                        WHERE kayit_id = jeoloji_kayitlari.id) AS geometri_sayisi
                FROM jeoloji_kayitlari
                WHERE {where}
                ORDER BY ilce_key, yerlesim_key, ada_key, parsel_key, formasyon_key, guncelleme_tarihi DESC
                """,
                params,
            ).fetchall()
        return [self._satir_sozluk(row) for row in rows]

    def revizyonlar(self, kayit_id):
        with self._baglan() as connection:
            rows = connection.execute(
                """
                SELECT revizyon_no, anlik_goruntu, kayit_tarihi
                FROM jeoloji_revizyonlari
                WHERE kayit_id = ? ORDER BY revizyon_no DESC
                """,
                (int(kayit_id),),
            ).fetchall()
        return [
            {
                "revizyon_no": int(row["revizyon_no"]),
                "kayit_tarihi": row["kayit_tarihi"],
                "kayit": json.loads(row["anlik_goruntu"]),
            }
            for row in rows
        ]

    def arsivle(self, kayit_id):
        with self._baglan() as connection:
            row = connection.execute(
                "SELECT revizyon_no FROM jeoloji_kayitlari WHERE id = ? AND aktif = 1",
                (int(kayit_id),),
            ).fetchone()
            if row is None:
                return False
            revision = int(row["revizyon_no"]) + 1
            connection.execute(
                """
                UPDATE jeoloji_kayitlari
                SET aktif = 0, revizyon_no = ?, guncelleme_tarihi = ?
                WHERE id = ?
                """,
                (revision, _simdi(), int(kayit_id)),
            )
            self._revizyon_yaz(connection, int(kayit_id), revision)
        return True

    def arsivden_cikar(self, kayit_id):
        with self._baglan() as connection:
            row = connection.execute(
                "SELECT * FROM jeoloji_kayitlari WHERE id = ? AND aktif = 0", (int(kayit_id),)
            ).fetchone()
            if row is None:
                return False
            record = dict(row)
            duplicate_id = self._ayni_kayit_id(connection, record, haric_id=kayit_id)
            if duplicate_id is not None:
                raise AyniJeolojiKaydiHatasi(duplicate_id)
            revision = int(row["revizyon_no"]) + 1
            connection.execute(
                """
                UPDATE jeoloji_kayitlari
                SET aktif = 1, revizyon_no = ?, guncelleme_tarihi = ?
                WHERE id = ?
                """,
                (revision, _simdi(), int(kayit_id)),
            )
            self._revizyon_yaz(connection, int(kayit_id), revision)
        return True

    def harita_dosyasi_ekle(self, source_path):
        source = Path(source_path)
        if not source.is_file():
            raise JeolojiKutuphanesiHatasi("Seçilen harita dosyası bulunamadı.")
        hasher = hashlib.sha256()
        with source.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                hasher.update(chunk)
        digest = hasher.hexdigest()[:20]
        suffix = source.suffix.lower() or ".bin"
        self.harita_dizini.mkdir(parents=True, exist_ok=True)
        target = self.harita_dizini / f"{digest}{suffix}"
        if not target.exists():
            shutil.copy2(source, target)
        return str(target)

    def _kml_dosyasi_ekle(self, source_path):
        source = Path(source_path)
        if source.suffix.lower() != ".kml" or not source.is_file():
            raise JeolojiKutuphanesiHatasi("Seçilen parsel KML dosyası bulunamadı.")
        digest = self._dosya_hashi(source)
        self.kml_dizini.mkdir(parents=True, exist_ok=True)
        target = self.kml_dizini / f"{digest[:24]}.kml"
        if not target.exists():
            shutil.copy2(source, target)
        return str(target), digest

    @staticmethod
    def _geometriyi_dogrula(poligon):
        ham_noktalar = poligon.get("noktalar", ()) if isinstance(poligon, dict) else poligon
        noktalar = []
        for nokta in ham_noktalar or ():
            if not isinstance(nokta, (list, tuple)) or len(nokta) < 2:
                raise JeolojiKutuphanesiHatasi("KML poligonunda geçersiz koordinat bulundu.")
            try:
                enlem = float(nokta[0])
                boylam = float(nokta[1])
            except (TypeError, ValueError) as exc:
                raise JeolojiKutuphanesiHatasi(
                    "KML poligon koordinatları sayısal olmalıdır."
                ) from exc
            if not (-90 <= enlem <= 90 and -180 <= boylam <= 180):
                raise JeolojiKutuphanesiHatasi("KML koordinatı WGS84 aralığının dışında.")
            noktalar.append([enlem, boylam])
        if len({(round(p[0], 12), round(p[1], 12)) for p in noktalar}) < 3:
            raise JeolojiKutuphanesiHatasi("Parsel poligonu en az üç farklı nokta içermelidir.")
        return noktalar

    def geometrileri_degistir(self, kayit_id, poligonlar, kml_path=""):
        """Bir rapor kaydının bütün parsel geometrilerini doğrulayıp yeniler."""
        hazir = []
        for sira, poligon in enumerate(poligonlar or (), start=1):
            noktalar = self._geometriyi_dogrula(poligon)
            enlemler = [p[0] for p in noktalar]
            boylamlar = [p[1] for p in noktalar]
            hazir.append(
                {
                    "sira": sira,
                    "ad": _temiz(poligon.get("ad")) if isinstance(poligon, dict) else "",
                    "aciklama": _temiz(poligon.get("aciklama")) if isinstance(poligon, dict) else "",
                    "noktalar": noktalar,
                    "min_enlem": min(enlemler),
                    "max_enlem": max(enlemler),
                    "min_boylam": min(boylamlar),
                    "max_boylam": max(boylamlar),
                }
            )
        if not hazir:
            raise JeolojiKutuphanesiHatasi("Kaydedilecek parsel poligonu bulunamadı.")

        stored_path = ""
        digest = ""
        if _temiz(kml_path):
            stored_path, digest = self._kml_dosyasi_ekle(kml_path)
        now = _simdi()
        with self._baglan() as connection:
            record = connection.execute(
                "SELECT id FROM jeoloji_kayitlari WHERE id = ? AND aktif = 1",
                (int(kayit_id),),
            ).fetchone()
            if record is None:
                raise JeolojiKutuphanesiHatasi("Geometri bağlanacak etkin kayıt bulunamadı.")
            connection.execute(
                "DELETE FROM jeoloji_geometrileri WHERE kayit_id = ?", (int(kayit_id),)
            )
            for item in hazir:
                connection.execute(
                    """
                    INSERT INTO jeoloji_geometrileri(
                        kayit_id, sira, placemark_adi, aciklama, noktalar_json,
                        min_enlem, max_enlem, min_boylam, max_boylam,
                        merkez_enlem, merkez_boylam, kml_path, kml_hash, olusturma_tarihi
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        int(kayit_id), item["sira"], item["ad"], item["aciklama"],
                        json.dumps(item["noktalar"], separators=(",", ":")),
                        item["min_enlem"], item["max_enlem"],
                        item["min_boylam"], item["max_boylam"],
                        (item["min_enlem"] + item["max_enlem"]) / 2,
                        (item["min_boylam"] + item["max_boylam"]) / 2,
                        stored_path, digest, now,
                    ),
                )
        return len(hazir)

    def yeni_kaydi_geometriyle_kaydet(self, kayit, poligonlar, kml_path):
        """Yeni kaydı ve zorunlu geometrisini tek proje işlemi olarak kaydeder."""
        record_id = self.kaydet(kayit)
        try:
            self.geometrileri_degistir(record_id, poligonlar, kml_path)
        except Exception:
            with self._baglan() as connection:
                connection.execute(
                    "DELETE FROM jeoloji_kayitlari WHERE id = ? AND revizyon_no = 1",
                    (int(record_id),),
                )
            raise
        return record_id

    def geometrileri_getir(self, kayit_id):
        with self._baglan() as connection:
            rows = connection.execute(
                """
                SELECT * FROM jeoloji_geometrileri
                WHERE kayit_id = ? ORDER BY sira
                """,
                (int(kayit_id),),
            ).fetchall()
        sonuc = []
        for row in rows:
            item = dict(row)
            item["noktalar"] = [tuple(point) for point in json.loads(item.pop("noktalar_json"))]
            sonuc.append(item)
        return sonuc

    def toplu_kayit_durumu(self, kayit, kaynak_rapor_hash="", kml_hash=""):
        """Toplu aktarım adayını yeni, aynı veya olası revizyon olarak sınıflandırır."""
        keys = {
            f"{field}_key": jeoloji_anahtari((kayit or {}).get(field, ""))
            for field in ("il", "ilce", "yerlesim", "ada", "parsel", "formasyon")
        }
        with self._baglan() as connection:
            row = connection.execute(
                """
                SELECT k.*,
                       COALESCE((SELECT g.kml_hash FROM jeoloji_geometrileri g
                                 WHERE g.kayit_id = k.id AND g.kml_hash <> ''
                                 ORDER BY g.sira LIMIT 1), '') AS mevcut_kml_hash
                FROM jeoloji_kayitlari k
                WHERE k.aktif = 1 AND k.il_key = ? AND k.ilce_key = ?
                  AND k.yerlesim_key = ? AND k.ada_key = ? AND k.parsel_key = ?
                  AND k.formasyon_key = ?
                LIMIT 1
                """,
                tuple(
                    keys[f"{field}_key"]
                    for field in ("il", "ilce", "yerlesim", "ada", "parsel", "formasyon")
                ),
            ).fetchone()
        if row is None:
            return {"durum": "yeni", "kayit": None}
        existing = self._satir_sozluk(row)
        rapor_ayni = bool(
            kaynak_rapor_hash
            and existing.get("kaynak_rapor_hash")
            and kaynak_rapor_hash == existing.get("kaynak_rapor_hash")
        )
        kml_ayni = bool(kml_hash and kml_hash == existing.get("mevcut_kml_hash"))
        if rapor_ayni and kml_ayni:
            return {"durum": "ayni", "kayit": existing}
        return {"durum": "revizyon", "kayit": existing}

    def harita_kayitlari(
        self,
        *,
        min_enlem,
        max_enlem,
        min_boylam,
        max_boylam,
        taslaklari_goster=False,
        limit=500,
    ):
        """Görünen koordinat kutusuyla kesişen etkin rapor geometrilerini döndürür."""
        min_enlem, max_enlem = sorted((float(min_enlem), float(max_enlem)))
        min_boylam, max_boylam = sorted((float(min_boylam), float(max_boylam)))
        clauses = [
            "k.aktif = 1",
            "g.min_enlem <= ?", "g.max_enlem >= ?",
            "g.min_boylam <= ?", "g.max_boylam >= ?",
        ]
        params = [max_enlem, min_enlem, max_boylam, min_boylam]
        if not taslaklari_goster:
            clauses.append("k.onay_durumu = 'onayli'")
        params.append(max(1, min(int(limit), 2000)))
        with self._baglan() as connection:
            rows = connection.execute(
                f"""
                SELECT k.*, g.id AS geometri_id, g.sira AS geometri_sirasi,
                       g.placemark_adi, g.aciklama AS geometri_aciklamasi,
                       g.noktalar_json, g.kml_path, g.merkez_enlem, g.merkez_boylam
                FROM jeoloji_geometrileri g
                JOIN jeoloji_kayitlari k ON k.id = g.kayit_id
                WHERE {' AND '.join(clauses)}
                ORDER BY k.onay_durumu DESC, k.guncelleme_tarihi DESC, k.id
                LIMIT ?
                """,
                params,
            ).fetchall()
        sonuc = []
        for row in rows:
            item = self._satir_sozluk(row)
            item["noktalar"] = [tuple(point) for point in json.loads(item.pop("noktalar_json"))]
            sonuc.append(item)
        return sonuc

    def yedek_paketi_olustur(self, target_path):
        target = Path(target_path)
        target.parent.mkdir(parents=True, exist_ok=True)
        manifest = {
            "tur": "RaporPro Çanakkale Jeoloji Kütüphanesi",
            "olusturma_tarihi": _simdi(),
            "veritabani": self.db_path.name,
        }
        with ZipFile(target, "w", compression=ZIP_DEFLATED) as archive:
            archive.write(self.db_path, arcname=self.db_path.name)
            archive.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2))
            if self.harita_dizini.is_dir():
                for path in sorted(self.harita_dizini.rglob("*")):
                    if path.is_file():
                        archive.write(path, arcname=str(Path("haritalar") / path.relative_to(self.harita_dizini)))
            if self.bolum_dizini.is_dir():
                for path in sorted(self.bolum_dizini.rglob("*")):
                    if path.is_file():
                        archive.write(
                            path,
                            arcname=str(
                                Path("jeoloji_bolumleri") / path.relative_to(self.bolum_dizini)
                            ),
                        )
            if self.kml_dizini.is_dir():
                for path in sorted(self.kml_dizini.rglob("*")):
                    if path.is_file():
                        archive.write(
                            path,
                            arcname=str(Path("parsel_kml") / path.relative_to(self.kml_dizini)),
                        )
        return str(target)

    @staticmethod
    def _adaylari_sirala(records, yerlesim_key, ada_key, parsel_key, formasyon_key, alan):
        records = [record for record in records if _temiz(record.get(alan))]
        if not records:
            return None

        def rank(record):
            yerlesim_exact = bool(yerlesim_key and record["yerlesim_key"] == yerlesim_key)
            ada_exact = bool(ada_key and record["ada_key"] == ada_key)
            parsel_exact = bool(parsel_key and record["parsel_key"] == parsel_key)
            formasyon_exact = bool(formasyon_key and record["formasyon_key"] == formasyon_key)
            yerlesim_generic = not record["yerlesim_key"]
            ada_generic = not record["ada_key"]
            parsel_generic = not record["parsel_key"]
            formasyon_generic = not record["formasyon_key"]
            if alan == "inceleme_alani_jeolojisi":
                specificity = (
                    parsel_exact,
                    ada_exact,
                    formasyon_exact,
                    yerlesim_exact,
                    parsel_generic,
                    ada_generic,
                    formasyon_generic,
                    yerlesim_generic,
                )
            else:
                specificity = (
                    yerlesim_exact,
                    parsel_exact,
                    ada_exact,
                    formasyon_exact,
                    yerlesim_generic,
                    parsel_generic,
                    ada_generic,
                    formasyon_generic,
                )
            return (*specificity, record.get("guncelleme_tarihi", ""))

        return max(records, key=rank)

    def uygun_icerigi_bul(
        self, *, il="Çanakkale", ilce, yerlesim="", ada="", parsel="", formasyon=""
    ):
        il_key = jeoloji_anahtari(il or "Çanakkale")
        ilce_key = jeoloji_anahtari(ilce)
        yerlesim_key = jeoloji_anahtari(yerlesim)
        ada_key = jeoloji_anahtari(ada)
        parsel_key = jeoloji_anahtari(parsel)
        formasyon_key = jeoloji_anahtari(formasyon)
        if not ilce_key:
            return None
        with self._baglan() as connection:
            rows = connection.execute(
                """
                SELECT * FROM jeoloji_kayitlari
                WHERE aktif = 1 AND onay_durumu = 'onayli'
                  AND il_key = ? AND ilce_key = ?
                  AND (yerlesim_key = '' OR yerlesim_key = ?)
                  AND (ada_key = '' OR ada_key = ?)
                  AND (parsel_key = '' OR parsel_key = ?)
                  AND (formasyon_key = '' OR formasyon_key = ?)
                """,
                (il_key, ilce_key, yerlesim_key, ada_key, parsel_key, formasyon_key),
            ).fetchall()
        records = [self._satir_sozluk(row) for row in rows]
        if not records:
            return None

        result = {}
        source_ids = set()
        for field in ("genel_jeoloji_metni", "inceleme_alani_jeolojisi"):
            source = self._adaylari_sirala(
                records, yerlesim_key, ada_key, parsel_key, formasyon_key, field
            )
            result[field] = source.get(field, "") if source else ""
            result[f"{field}_kayit_id"] = source.get("id") if source else None
            if field == "genel_jeoloji_metni":
                result["bolum_docx_path"] = source.get("bolum_docx_path", "") if source else ""
                result["bolum_hash"] = source.get("bolum_hash", "") if source else ""
            if source:
                source_ids.add(int(source["id"]))
        map_source = self._adaylari_sirala(
            records, yerlesim_key, ada_key, parsel_key, formasyon_key, "harita_path"
        )
        for field in ("harita_path", "harita_aciklamasi", "harita_kaynagi", "harita_olcegi"):
            result[field] = map_source.get(field, "") if map_source else ""
            result[f"{field}_kayit_id"] = map_source.get("id") if map_source else None
        if map_source:
            source_ids.add(int(map_source["id"]))
        result["kayit_idleri"] = sorted(source_ids)
        result["il"] = _temiz(il) or "Çanakkale"
        result["ilce"] = _temiz(ilce)
        result["yerlesim"] = _temiz(yerlesim)
        result["ada"] = _temiz(ada)
        result["parsel"] = _temiz(parsel)
        result["formasyon"] = _temiz(formasyon)
        return result
