import datetime
import copy
import json
import logging
import math
import os
import shutil
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox

from harita_dosyalari import RAPOR_HARITA_ALANLARI, harita_verisini_mevcut_dosyalarla_tamamla
from harita_renkleri import CALISAN_PARSEL_SINIR_KALINLIGI, CALISAN_PARSEL_SINIR_RENGI
from on_deger import (
    bitmis_revizyonu_arsivle,
    is_durumu_degistir,
    normalize_is_akisi,
    normalize_on_deger,
    normalize_tdth,
)


logger = logging.getLogger("ZeminRaporPro")

SON_PROJE_LIMITI = 8
SCHEMA_VERSION = 7
YEDEK_LIMITI = 20
TAAHHUT_BILGI_ALANLARI = (
    "JEOFIZIK_MUH_AD",
    "JEOFIZIK_MUH_SICIL",
    "JEOFIZIK_MUH_ADRES",
    "JEOFIZIK_MUH_TELEFON",
    "JEOLOJI_MUH_AD",
    "JEOLOJI_MUH_SICIL",
    "JEOLOJI_MUH_ADRES",
    "JEOLOJI_MUH_TELEFON",
)


class KayitYoneticisi:
    def __init__(self, app):
        object.__setattr__(self, "app", app)

    def __getattr__(self, name):
        return getattr(self.app, name)

    def __setattr__(self, name, value):
        if name == "app":
            object.__setattr__(self, name, value)
        else:
            setattr(self.app, name, value)

    def atomik_json_yaz(self, yol, veriler, indent=4):
        hedef = os.path.abspath(yol)
        klasor = os.path.dirname(hedef) or os.getcwd()
        os.makedirs(klasor, exist_ok=True)
        fd, gecici_yol = tempfile.mkstemp(
            prefix=f".{os.path.basename(hedef)}.",
            suffix=".tmp",
            dir=klasor,
        )
        try:
            with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as f:
                json.dump(veriler, f, ensure_ascii=False, indent=indent, allow_nan=False)
                f.flush()
                os.fsync(f.fileno())
            os.replace(gecici_yol, hedef)
        except Exception:
            try:
                os.close(fd)
            except OSError:
                pass
            try:
                os.remove(gecici_yol)
            except OSError:
                pass
            raise
        return hedef

    def atomik_dosya_kopyala(self, kaynak_yolu, hedef_yolu):
        hedef = os.path.abspath(hedef_yolu)
        klasor = os.path.dirname(hedef) or os.getcwd()
        os.makedirs(klasor, exist_ok=True)
        fd, gecici_yol = tempfile.mkstemp(
            prefix=f".{os.path.basename(hedef)}.",
            suffix=".tmp",
            dir=klasor,
        )
        try:
            with open(kaynak_yolu, "rb") as kaynak, os.fdopen(fd, "wb") as hedef_dosya:
                shutil.copyfileobj(kaynak, hedef_dosya)
                hedef_dosya.flush()
                os.fsync(hedef_dosya.fileno())
            os.replace(gecici_yol, hedef)
            try:
                shutil.copystat(kaynak_yolu, hedef)
            except OSError:
                pass
        except Exception:
            try:
                os.close(fd)
            except OSError:
                pass
            try:
                os.remove(gecici_yol)
            except OSError:
                pass
            raise
        return hedef

    def son_projeler_dosya_yolu(self):
        return os.path.join(self.kullanici_veri_klasoru_bul(), "son_projeler.json")

    def son_projeleri_yukle(self):
        yeni_yol = self.son_projeler_dosya_yolu()
        eski_yol = os.path.join(self.uygulama_klasoru_bul(), "son_projeler.json")
        veriler = None
        kullanilan_yol = None
        for yol in (yeni_yol, eski_yol):
            if kullanilan_yol is not None:
                break
            if os.path.normcase(os.path.abspath(yol)) == os.path.normcase(os.path.abspath(yeni_yol)) and yol != yeni_yol:
                continue
            try:
                with open(yol, "r", encoding="utf-8-sig") as f:
                    veriler = json.load(f)
                kullanilan_yol = yol
            except (OSError, json.JSONDecodeError):
                continue

        if not isinstance(veriler, list):
            return []
        temiz = []
        for proje_yolu in veriler:
            if not isinstance(proje_yolu, str) or not proje_yolu.strip():
                continue
            tam_yol = os.path.abspath(proje_yolu)
            if os.path.exists(tam_yol) and tam_yol not in temiz:
                temiz.append(tam_yol)
        temiz = temiz[:SON_PROJE_LIMITI]
        if kullanilan_yol and os.path.normcase(os.path.abspath(kullanilan_yol)) != os.path.normcase(os.path.abspath(yeni_yol)):
            try:
                self.atomik_json_yaz(yeni_yol, temiz, indent=2)
            except OSError as e:
                self.hata_kaydet("Eski son proje listesi kullanıcı klasörüne taşınamadı", e)
        return temiz

    def son_projeleri_kaydet(self):
        try:
            self.atomik_json_yaz(
                self.son_projeler_dosya_yolu(),
                getattr(self, "son_projeler", []),
                indent=2,
            )
        except OSError as e:
            self.hata_kaydet("Son açılan projeler kaydedilemedi", e)

    def son_proje_ekle(self, yol):
        if not yol:
            return
        tam_yol = os.path.abspath(yol)
        mevcut = [
            p for p in getattr(self, "son_projeler", [])
            if os.path.normcase(os.path.abspath(p)) != os.path.normcase(tam_yol)
        ]
        self.son_projeler = [tam_yol] + mevcut
        self.son_projeler = self.son_projeler[:SON_PROJE_LIMITI]
        self.son_projeleri_kaydet()
        self.son_projeler_menusunu_guncelle()

    def son_projeleri_temizle(self):
        self.son_projeler = []
        self.son_projeleri_kaydet()
        self.son_projeler_menusunu_guncelle()
        if hasattr(self, "durum_mesaji_yaz"):
            self.durum_mesaji_yaz("Son açılan projeler listesi temizlendi")

    def son_projeler_menusunu_guncelle(self):
        menu = getattr(self, "son_projeler_menusu", None)
        if not menu:
            return
        menu.delete(0, tk.END)
        projeler = getattr(self, "son_projeler", [])
        if not projeler:
            menu.add_command(label="Kayıt yok", state="disabled")
            return

        for index, yol in enumerate(projeler, start=1):
            etiket = f"{index}. {os.path.basename(yol)}"
            menu.add_command(label=etiket, command=lambda p=yol: self.son_proje_ac(p))
        menu.add_separator()
        menu.add_command(label="Listeyi Temizle", command=self.son_projeleri_temizle)

    def son_proje_ac(self, yol):
        if not os.path.exists(yol):
            messagebox.showwarning("Son Açılan Proje", "Seçilen proje dosyası artık bulunamıyor.")
            self.son_projeler = [
                p for p in getattr(self, "son_projeler", [])
                if os.path.normcase(os.path.abspath(p)) != os.path.normcase(os.path.abspath(yol))
            ]
            self.son_projeleri_kaydet()
            self.son_projeler_menusunu_guncelle()
            return False
        return self.proje_dosyasini_ac(yol)

    def yedek_klasoru(self):
        klasor = os.path.join(self.kullanici_veri_klasoru_bul(), "yedekler")
        os.makedirs(klasor, exist_ok=True)
        return klasor

    def guvenli_dosya_adi(self, deger):
        metin = str(deger or "K1").strip() or "K1"
        yasak = '<>:"/\\|?*'
        for karakter in yasak:
            metin = metin.replace(karakter, "_")
        return metin[:80].strip(" .") or "K1"

    def otomatik_yedek_yolu(self, kaynak_yolu):
        proje_adi = os.path.splitext(os.path.basename(kaynak_yolu or ""))[0] or self.proje_deger("PROJE_ADI", "K1")
        proje_adi = self.guvenli_dosya_adi(proje_adi)
        zaman = datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        temel = os.path.join(self.yedek_klasoru(), f"{proje_adi}_{zaman}")
        aday = f"{temel}.json"
        sayac = 1
        while os.path.exists(aday):
            aday = f"{temel}_{sayac}.json"
            sayac += 1
        return aday

    def yedekleri_sinirla(self, kaynak_yolu):
        proje_adi = os.path.splitext(os.path.basename(kaynak_yolu or ""))[0] or self.proje_deger("PROJE_ADI", "K1")
        onek = self.guvenli_dosya_adi(proje_adi) + "_"
        klasor = self.yedek_klasoru()
        adaylar = []
        try:
            for kayit in os.scandir(klasor):
                if kayit.is_file() and kayit.name.startswith(onek) and kayit.name.lower().endswith(".json"):
                    adaylar.append(kayit)
            adaylar.sort(key=lambda kayit: (kayit.stat().st_mtime_ns, kayit.name), reverse=True)
            for kayit in adaylar[YEDEK_LIMITI:]:
                try:
                    os.remove(kayit.path)
                except OSError as e:
                    self.hata_kaydet(f"Eski yedek silinemedi: {kayit.path}", e)
        except OSError as e:
            self.hata_kaydet("Yedek listesi sınırlandırılamadı", e)

    def otomatik_yedek_olustur(self, kaynak_yolu, veriler=None):
        try:
            yedek_yolu = self.otomatik_yedek_yolu(kaynak_yolu)
            if veriler is not None:
                self.atomik_json_yaz(yedek_yolu, veriler)
            elif kaynak_yolu and os.path.exists(kaynak_yolu):
                self.atomik_dosya_kopyala(kaynak_yolu, yedek_yolu)
            else:
                return ""
            self.yedekleri_sinirla(kaynak_yolu)
            logger.info("Otomatik yedek oluşturuldu: %s", yedek_yolu)
            return yedek_yolu
        except Exception as e:
            self.hata_kaydet("Otomatik yedek oluşturulamadı", e)
            return ""

    def varsayilan_proje_verisi_al(self):
        varsayilan = getattr(self, "varsayilan_proje_verisi", None)
        if not isinstance(varsayilan, dict):
            raise RuntimeError("Varsayılan proje durumu henüz hazırlanmadı.")
        return copy.deepcopy(varsayilan)

    def _metne_cevir(self, deger, alan):
        if deger is None:
            return ""
        if isinstance(deger, (dict, list)):
            raise ValueError(f"{alan} metin veya sayı olmalıdır.")
        return str(deger)

    def _bool_deger(self, deger, varsayilan=False):
        if isinstance(deger, bool):
            return deger
        if isinstance(deger, (int, float)):
            return bool(deger)
        if isinstance(deger, str):
            kucuk = deger.strip().lower()
            if kucuk in ("1", "true", "evet", "yes", "on"):
                return True
            if kucuk in ("0", "false", "hayır", "hayir", "no", "off", ""):
                return False
        return bool(varsayilan)

    def _sonlu_sayi(self, deger, varsayilan, alan, alt=None, ust=None):
        try:
            sonuc = float(deger)
        except (TypeError, ValueError):
            raise ValueError(f"{alan} sayısal olmalıdır.") from None
        if not math.isfinite(sonuc):
            raise ValueError(f"{alan} sonlu bir sayı olmalıdır.")
        if alt is not None and sonuc < alt:
            raise ValueError(f"{alan} en az {alt} olmalıdır.")
        if ust is not None and sonuc > ust:
            raise ValueError(f"{alan} en fazla {ust} olmalıdır.")
        return sonuc

    def _tablo_satirlarini_dogrula(self, satirlar, alan):
        if satirlar is None:
            return []
        if not isinstance(satirlar, list):
            raise ValueError(f"{alan} bir satır listesi olmalıdır.")
        sonuc = []
        for index, satir in enumerate(satirlar):
            if not isinstance(satir, list):
                raise ValueError(f"{alan}[{index}] bir hücre listesi olmalıdır.")
            sonuc.append([
                self._metne_cevir(hucre, f"{alan}[{index}]")
                for hucre in satir
            ])
        return sonuc

    def _koordinat_ciftini_dogrula(self, deger, alan):
        if not isinstance(deger, (list, tuple)) or len(deger) < 2:
            raise ValueError(f"{alan} [enlem, boylam] biçiminde olmalıdır.")
        enlem = self._sonlu_sayi(deger[0], 0.0, f"{alan}.enlem", -90.0, 90.0)
        boylam = self._sonlu_sayi(deger[1], 0.0, f"{alan}.boylam", -180.0, 180.0)
        return [enlem, boylam]

    def _jeoloji_pafta_sonucunu_dogrula(self, deger):
        if deger in (None, ""):
            return {}
        if not isinstance(deger, dict):
            raise ValueError("__HARITA__.jeoloji_pafta_sonucu bir JSON nesnesi olmalıdır.")
        if not deger.get("birim_adi"):
            return {}
        sonuc = {
            "pafta_id": self._metne_cevir(deger.get("pafta_id", ""), "jeoloji_pafta_sonucu.pafta_id"),
            "pafta_adi": self._metne_cevir(deger.get("pafta_adi", ""), "jeoloji_pafta_sonucu.pafta_adi"),
            "birim_kodu": self._metne_cevir(deger.get("birim_kodu", ""), "jeoloji_pafta_sonucu.birim_kodu"),
            "birim_adi": self._metne_cevir(deger.get("birim_adi", ""), "jeoloji_pafta_sonucu.birim_adi"),
            "oran": self._sonlu_sayi(deger.get("oran", 0), 0, "jeoloji_pafta_sonucu.oran", 0, 100),
            "puan": self._sonlu_sayi(deger.get("puan", 0), 0, "jeoloji_pafta_sonucu.puan", 0, 100),
            "guven": self._sonlu_sayi(deger.get("guven", 0), 0, "jeoloji_pafta_sonucu.guven", 0, 100),
            "kanit_yolu": self._metne_cevir(deger.get("kanit_yolu", ""), "jeoloji_pafta_sonucu.kanit_yolu"),
            "tarih": self._metne_cevir(deger.get("tarih", ""), "jeoloji_pafta_sonucu.tarih"),
            "adaylar": [],
        }
        adaylar = deger.get("adaylar", [])
        if not isinstance(adaylar, list):
            raise ValueError("jeoloji_pafta_sonucu.adaylar bir liste olmalıdır.")
        if len(adaylar) > 20:
            raise ValueError("jeoloji_pafta_sonucu.adaylar en fazla 20 kayıt içerebilir.")
        for index, aday in enumerate(adaylar):
            if not isinstance(aday, dict):
                raise ValueError(f"jeoloji_pafta_sonucu.adaylar[{index}] bir JSON nesnesi olmalıdır.")
            sonuc["adaylar"].append({
                "kod": self._metne_cevir(aday.get("kod", ""), f"adaylar[{index}].kod"),
                "ad": self._metne_cevir(aday.get("ad", ""), f"adaylar[{index}].ad"),
                "oran": self._sonlu_sayi(aday.get("oran", 0), 0, f"adaylar[{index}].oran", 0, 100),
                "puan": self._sonlu_sayi(aday.get("puan", 0), 0, f"adaylar[{index}].puan", 0, 100),
                "guven": self._sonlu_sayi(aday.get("guven", 0), 0, f"adaylar[{index}].guven", 0, 100),
                "pafta_id": self._metne_cevir(aday.get("pafta_id", ""), f"adaylar[{index}].pafta_id"),
                "pafta_adi": self._metne_cevir(aday.get("pafta_adi", ""), f"adaylar[{index}].pafta_adi"),
                "lejant_id": self._metne_cevir(aday.get("lejant_id", ""), f"adaylar[{index}].lejant_id"),
            })
        return sonuc

    def _genel_jeoloji_verisini_dogrula(self, deger):
        if deger in (None, ""):
            return {}
        if not isinstance(deger, dict):
            raise ValueError("__HARITA__.genel_jeoloji_verisi bir JSON nesnesi olmalıdır.")
        birimler = deger.get("birimler", [])
        if not isinstance(birimler, list):
            raise ValueError("genel_jeoloji_verisi.birimler bir liste olmalıdır.")
        if len(birimler) > 60:
            raise ValueError("genel_jeoloji_verisi.birimler en fazla 60 kayıt içerebilir.")
        temiz_birimler = []
        for index, birim in enumerate(birimler):
            if not isinstance(birim, dict):
                raise ValueError(f"genel_jeoloji_verisi.birimler[{index}] bir JSON nesnesi olmalıdır.")
            prefix = f"genel_jeoloji_verisi.birimler[{index}]"
            try:
                library_id = int(birim["kutuphane_id"]) if birim.get("kutuphane_id") not in (None, "") else None
                library_revision = (
                    int(birim["kutuphane_revizyon_no"])
                    if birim.get("kutuphane_revizyon_no") not in (None, "")
                    else None
                )
            except (TypeError, ValueError):
                raise ValueError(f"{prefix} kütüphane kimliği/revizyonu tam sayı olmalıdır.") from None
            temiz_birimler.append({
                "kod": self._metne_cevir(birim.get("kod", ""), f"{prefix}.kod"),
                "ad": self._metne_cevir(birim.get("ad", ""), f"{prefix}.ad"),
                "jeolojik_yas": self._metne_cevir(birim.get("jeolojik_yas", ""), f"{prefix}.jeolojik_yas"),
                "yas_sirasi": int(self._sonlu_sayi(birim.get("yas_sirasi", 9999), 9999, f"{prefix}.yas_sirasi", 0, 100000)),
                "kaynak_sirasi": int(self._sonlu_sayi(birim.get("kaynak_sirasi", -1), -1, f"{prefix}.kaynak_sirasi", -1, 100000)),
                "oran": self._sonlu_sayi(birim.get("oran", 0), 0, f"{prefix}.oran", 0, 100),
                "guven": self._sonlu_sayi(birim.get("guven", 0), 0, f"{prefix}.guven", 0, 100),
                "pafta_idleri": [
                    self._metne_cevir(value, f"{prefix}.pafta_idleri")
                    for value in (birim.get("pafta_idleri", []) if isinstance(birim.get("pafta_idleri", []), list) else [])[:20]
                ],
                "pafta_adlari": [
                    self._metne_cevir(value, f"{prefix}.pafta_adlari")
                    for value in (birim.get("pafta_adlari", []) if isinstance(birim.get("pafta_adlari", []), list) else [])[:20]
                ],
                "ana_birim": bool(birim.get("ana_birim", False)),
                "kullan": bool(birim.get("kullan", True)),
                "lejant_aciklamasi": self._metne_cevir(birim.get("lejant_aciklamasi", ""), f"{prefix}.lejant_aciklamasi"),
                "bolgesel_jeoloji_metni": self._metne_cevir(birim.get("bolgesel_jeoloji_metni", ""), f"{prefix}.bolgesel_jeoloji_metni"),
                "kutuphane_id": library_id,
                "kutuphane_revizyon_no": library_revision,
                "metin_kaynagi": self._metne_cevir(birim.get("metin_kaynagi", ""), f"{prefix}.metin_kaynagi"),
                "yerel_aday": bool(birim.get("yerel_aday", True)),
                "ai_gemini_guven": self._sonlu_sayi(
                    birim.get("ai_gemini_guven", 0), 0, f"{prefix}.ai_gemini_guven", 0, 100
                ),
                "ai_openai_guven": self._sonlu_sayi(
                    birim.get("ai_openai_guven", 0), 0, f"{prefix}.ai_openai_guven", 0, 100
                ),
                "ai_birlesik_guven": self._sonlu_sayi(
                    birim.get("ai_birlesik_guven", 0), 0, f"{prefix}.ai_birlesik_guven", 0, 100
                ),
                "ai_kanit": self._metne_cevir(birim.get("ai_kanit", ""), f"{prefix}.ai_kanit"),
                "ai_gemini_kanit": self._metne_cevir(
                    birim.get("ai_gemini_kanit", ""), f"{prefix}.ai_gemini_kanit"
                ),
                "ai_openai_kanit": self._metne_cevir(
                    birim.get("ai_openai_kanit", ""), f"{prefix}.ai_openai_kanit"
                ),
                "ai_gemini_aciklama": self._metne_cevir(
                    birim.get("ai_gemini_aciklama", ""), f"{prefix}.ai_gemini_aciklama"
                ),
                "ai_openai_aciklama": self._metne_cevir(
                    birim.get("ai_openai_aciklama", ""), f"{prefix}.ai_openai_aciklama"
                ),
                "ai_durum": self._metne_cevir(birim.get("ai_durum", ""), f"{prefix}.ai_durum"),
                "ai_oneri": self._metne_cevir(birim.get("ai_oneri", ""), f"{prefix}.ai_oneri"),
                "ai_aciklama": self._metne_cevir(birim.get("ai_aciklama", ""), f"{prefix}.ai_aciklama"),
                "ai_saglayicilar": [
                    self._metne_cevir(value, f"{prefix}.ai_saglayicilar")
                    for value in (
                        birim.get("ai_saglayicilar", [])
                        if isinstance(birim.get("ai_saglayicilar", []), list) else []
                    )[:4]
                ],
            })
        mode = self._metne_cevir(deger.get("kaynak_modu", "kutuphane"), "genel_jeoloji_verisi.kaynak_modu")
        if mode not in {"kutuphane", "eski_rapor", "karma"}:
            mode = "kutuphane"
        try:
            old_id = int(deger["eski_rapor_kayit_id"]) if deger.get("eski_rapor_kayit_id") not in (None, "") else None
            old_revision = int(deger["eski_rapor_revizyon_no"]) if deger.get("eski_rapor_revizyon_no") not in (None, "") else None
        except (TypeError, ValueError):
            raise ValueError("genel_jeoloji_verisi eski rapor kimliği/revizyonu tam sayı olmalıdır.") from None
        center = deger.get("merkez", [])
        clean_center = self._koordinat_ciftini_dogrula(center, "genel_jeoloji_verisi.merkez") if center else []
        boundary = deger.get("sinir", {}) if isinstance(deger.get("sinir", {}), dict) else {}
        clean_boundary = {
            key: self._sonlu_sayi(boundary.get(key, 0), 0, f"genel_jeoloji_verisi.sinir.{key}", -180, 180)
            for key in ("north", "south", "east", "west")
        } if boundary else {}
        raw_provider_summaries = deger.get("ai_saglayici_ozetleri", {})
        clean_provider_summaries = {}
        if isinstance(raw_provider_summaries, dict):
            for provider in ("gemini", "openai"):
                summary = raw_provider_summaries.get(provider)
                if not isinstance(summary, dict):
                    continue
                clean_provider_summaries[provider] = {
                    "model": self._metne_cevir(
                        summary.get("model", ""),
                        f"genel_jeoloji_verisi.ai_saglayici_ozetleri.{provider}.model",
                    ),
                    "genel_guven": self._sonlu_sayi(
                        summary.get("genel_guven", 0), 0,
                        f"genel_jeoloji_verisi.ai_saglayici_ozetleri.{provider}.genel_guven",
                        0, 100,
                    ),
                    "ana_parsel_kodu": self._metne_cevir(
                        summary.get("ana_parsel_kodu", ""),
                        f"genel_jeoloji_verisi.ai_saglayici_ozetleri.{provider}.ana_parsel_kodu",
                    ),
                    "notlar": self._metne_cevir(
                        summary.get("notlar", ""),
                        f"genel_jeoloji_verisi.ai_saglayici_ozetleri.{provider}.notlar",
                    ),
                    "onbellekten": bool(summary.get("onbellekten", False)),
                }
        return {
            "surum": int(self._sonlu_sayi(deger.get("surum", 1), 1, "genel_jeoloji_verisi.surum", 1, 20)),
            "kaynak_modu": mode,
            "eski_rapor_kayit_id": old_id,
            "eski_rapor_revizyon_no": old_revision,
            "merkez": clean_center,
            "sinir": clean_boundary,
            "genislik_km": self._sonlu_sayi(deger.get("genislik_km", 15), 15, "genel_jeoloji_verisi.genislik_km", 1, 100),
            "yukseklik_km": self._sonlu_sayi(deger.get("yukseklik_km", 9), 9, "genel_jeoloji_verisi.yukseklik_km", 1, 100),
            "geometri_hash": self._metne_cevir(deger.get("geometri_hash", ""), "genel_jeoloji_verisi.geometri_hash"),
            "pafta_idleri": [self._metne_cevir(value, "genel_jeoloji_verisi.pafta_idleri") for value in (deger.get("pafta_idleri", []) if isinstance(deger.get("pafta_idleri", []), list) else [])[:30]],
            "pafta_adlari": [self._metne_cevir(value, "genel_jeoloji_verisi.pafta_adlari") for value in (deger.get("pafta_adlari", []) if isinstance(deger.get("pafta_adlari", []), list) else [])[:30]],
            "birimler": temiz_birimler,
            "bolgesel_jeoloji_metni": self._metne_cevir(deger.get("bolgesel_jeoloji_metni", ""), "genel_jeoloji_verisi.bolgesel_jeoloji_metni"),
            "eksik_metinler": [self._metne_cevir(value, "genel_jeoloji_verisi.eksik_metinler") for value in (deger.get("eksik_metinler", []) if isinstance(deger.get("eksik_metinler", []), list) else [])[:30]],
            "ai_inceleme_tarihi": self._metne_cevir(
                deger.get("ai_inceleme_tarihi", ""), "genel_jeoloji_verisi.ai_inceleme_tarihi"
            ),
            "ai_saglayicilar": [
                self._metne_cevir(value, "genel_jeoloji_verisi.ai_saglayicilar")
                for value in (
                    deger.get("ai_saglayicilar", [])
                    if isinstance(deger.get("ai_saglayicilar", []), list) else []
                )[:4]
            ],
            "ai_saglayici_ozetleri": clean_provider_summaries,
            "ai_onbellekten": bool(deger.get("ai_onbellekten", False)),
            "ai_uyarilar": [
                self._metne_cevir(value, "genel_jeoloji_verisi.ai_uyarilar")
                for value in (
                    deger.get("ai_uyarilar", [])
                    if isinstance(deger.get("ai_uyarilar", []), list) else []
                )[:10]
            ],
            "olusturma_tarihi": self._metne_cevir(deger.get("olusturma_tarihi", ""), "genel_jeoloji_verisi.olusturma_tarihi"),
        }

    def _tasima_rapor_imzasini_dogrula(self, imza):
        if imza is None:
            return None
        if not isinstance(imza, (list, tuple)) or len(imza) != 4:
            raise ValueError("_TASIMA_.rapor_imzasi geçersizdir.")
        tur, varsayim, girdiler, qt_nihai = imza
        if not isinstance(girdiler, (list, tuple)):
            raise ValueError("_TASIMA_.rapor_imzasi girdileri bir liste olmalıdır.")
        temiz_girdiler = []
        for index, girdi in enumerate(girdiler):
            if not isinstance(girdi, (list, tuple)) or len(girdi) != 2:
                raise ValueError(f"_TASIMA_.rapor_imzasi.girdiler[{index}] geçersizdir.")
            temiz_girdiler.append((
                self._metne_cevir(girdi[0], f"rapor_imzasi.girdiler[{index}].kod"),
                self._metne_cevir(girdi[1], f"rapor_imzasi.girdiler[{index}].deger"),
            ))
        return (
            self._metne_cevir(tur, "rapor_imzasi.tur"),
            self._bool_deger(varsayim),
            tuple(temiz_girdiler),
            self._metne_cevir(qt_nihai, "rapor_imzasi.qt_nihai"),
        )

    def _yol_kok_icinde_mi(self, yol, kok):
        if not yol or not kok:
            return False
        try:
            tam_yol = os.path.normcase(os.path.abspath(yol))
            tam_kok = os.path.normcase(os.path.abspath(kok))
            return os.path.commonpath((tam_yol, tam_kok)) == tam_kok
        except (OSError, ValueError):
            return False

    def _uygulama_sablon_yolu_mu(self, yol):
        """Paket/kullanıcı veri alanındaki uygulama şablonlarını proje varlığı sayma."""
        kokler = []
        sablon_kok_adaylari = getattr(self.app, "sablon_kok_adaylari", None)
        if callable(sablon_kok_adaylari):
            try:
                kokler.extend(sablon_kok_adaylari())
            except Exception as e:
                self.hata_kaydet("Şablon kökleri taşınabilir kayıt için belirlenemedi", e)
        try:
            kokler.append(os.path.join(self.uygulama_klasoru_bul(), "ornek_sablonlar"))
        except Exception:
            pass
        return any(self._yol_kok_icinde_mi(yol, kok) for kok in kokler)

    def _proje_yolu_kayda_hazirla(self, yol, proje_yolu, sablon=False):
        if not isinstance(yol, str) or not yol.strip():
            return yol
        if not os.path.isabs(yol):
            return yol
        if sablon and self._uygulama_sablon_yolu_mu(yol):
            return yol
        proje_klasoru = os.path.dirname(os.path.abspath(proje_yolu))
        if not self._yol_kok_icinde_mi(yol, proje_klasoru):
            return yol
        return os.path.relpath(os.path.abspath(yol), proje_klasoru)

    def _proje_yolu_coz(self, yol, proje_yolu, sablon=False):
        del sablon  # Dönüştürücülerin ortak çağrı imzasını korur.
        if not isinstance(yol, str) or not yol.strip() or os.path.isabs(yol):
            return yol
        proje_klasoru = os.path.dirname(os.path.abspath(proje_yolu))
        return os.path.abspath(os.path.join(proje_klasoru, yol))

    def _proje_dosya_yollarini_donustur(self, veriler, donustur):
        """Proje içindeki tüm dosya referanslarını bir kopya üzerinde dönüştürür."""
        sonuc = copy.deepcopy(veriler)
        if not isinstance(sonuc, dict):
            return sonuc

        for anahtar in ("_RAPOR_SABLONU_", "_TAAHHUT_WORD_SABLONU_", "_JEOLOJI_SABLONU_"):
            if anahtar in sonuc:
                sonuc[anahtar] = donustur(sonuc[anahtar], sablon=True)

        if "_LAB_KAYNAK_" in sonuc:
            sonuc["_LAB_KAYNAK_"] = donustur(sonuc["_LAB_KAYNAK_"])

        jeofizik = sonuc.get("_JEOFIZIK_")
        if isinstance(jeofizik, dict) and "excel_yolu" in jeofizik:
            jeofizik["excel_yolu"] = donustur(jeofizik["excel_yolu"])

        harita = sonuc.get("__HARITA__")
        if isinstance(harita, dict):
            for alan in (
                "kml_yolu",
                "parsel_haritasi_yolu",
                *(json_alani for json_alani, _ in RAPOR_HARITA_ALANLARI.values()),
            ):
                if alan in harita:
                    harita[alan] = donustur(harita[alan])
            pafta_sonucu = harita.get("jeoloji_pafta_sonucu")
            if isinstance(pafta_sonucu, dict) and "kanit_yolu" in pafta_sonucu:
                pafta_sonucu["kanit_yolu"] = donustur(pafta_sonucu["kanit_yolu"])

        kutuphane_bolumu = sonuc.get("_JEOLOJI_KUTUPHANE_BOLUMU_")
        if isinstance(kutuphane_bolumu, dict) and "bolum_docx_path" in kutuphane_bolumu:
            kutuphane_bolumu["bolum_docx_path"] = donustur(
                kutuphane_bolumu["bolum_docx_path"], sablon=True
            )

        ekler = sonuc.get("_EKLER_")
        if isinstance(ekler, dict):
            for kategori_verisi in ekler.values():
                if not isinstance(kategori_verisi, list):
                    continue
                for index, ek in enumerate(kategori_verisi):
                    if isinstance(ek, str):
                        kategori_verisi[index] = donustur(ek)
                    elif isinstance(ek, dict) and "yol" in ek:
                        ek["yol"] = donustur(ek["yol"])

        tdth = sonuc.get("_TDTH_")
        if isinstance(tdth, dict):
            aktif = tdth.get("aktif")
            if isinstance(aktif, dict) and "pdf_yolu" in aktif:
                aktif["pdf_yolu"] = donustur(aktif["pdf_yolu"])
            gecmis = tdth.get("gecmis", [])
            if isinstance(gecmis, list):
                for kayit in gecmis:
                    if isinstance(kayit, dict) and "pdf_yolu" in kayit:
                        kayit["pdf_yolu"] = donustur(kayit["pdf_yolu"])

        is_akisi = sonuc.get("_IS_AKISI_")
        if isinstance(is_akisi, dict):
            for anahtar in ("son_rapor_yolu", "son_nihai_pdf_yolu"):
                if anahtar in is_akisi:
                    is_akisi[anahtar] = donustur(is_akisi[anahtar])
            tamamlanan = is_akisi.get("tamamlanan_revizyonlar", [])
            if isinstance(tamamlanan, list):
                for revizyon in tamamlanan:
                    if not isinstance(revizyon, dict):
                        continue
                    for anahtar in ("tdth_pdf_yolu", "rapor_yolu", "nihai_pdf_yolu"):
                        if anahtar in revizyon:
                            revizyon[anahtar] = donustur(revizyon[anahtar])

        return sonuc

    def proje_yollarini_kayda_hazirla(self, veriler, proje_yolu):
        return self._proje_dosya_yollarini_donustur(
            veriler,
            lambda yol, sablon=False: self._proje_yolu_kayda_hazirla(
                yol, proje_yolu, sablon=sablon
            ),
        )

    def proje_yollarini_coz(self, veriler, proje_yolu):
        return self._proje_dosya_yollarini_donustur(
            veriler,
            lambda yol, sablon=False: self._proje_yolu_coz(
                yol, proje_yolu, sablon=sablon
            ),
        )

    def _taahhut_bilgilerini_topla(self):
        kaynak = getattr(self, "taahhut_bilgileri", None)
        if not isinstance(kaynak, dict):
            kaynak = getattr(self, "taahhut_varsayilanlari", {})
        return {
            kod: self._metne_cevir(kaynak.get(kod, ""), f"_TAAHHUT_BILGILERI_.{kod}").strip()
            for kod in TAAHHUT_BILGI_ALANLARI
        }

    def _taahhut_bilgilerini_yerlestir(self, veriler):
        kaynak = veriler if isinstance(veriler, dict) else {}
        self.taahhut_bilgileri = {
            kod: self._metne_cevir(kaynak.get(kod, ""), f"_TAAHHUT_BILGILERI_.{kod}").strip()
            for kod in TAAHHUT_BILGI_ALANLARI
        }

    def proje_verisini_normalize_et(self, veriler):
        """Legacy kayıtları boş varsayılan proje üzerine güvenli biçimde taşır."""
        if not isinstance(veriler, dict):
            raise ValueError("Proje dosyasının kök değeri bir JSON nesnesi olmalıdır.")

        ham_surum = veriler.get("schema_version", 0)
        try:
            surum = int(ham_surum)
        except (TypeError, ValueError):
            raise ValueError("schema_version tam sayı olmalıdır.") from None
        if surum < 0:
            raise ValueError("schema_version negatif olamaz.")
        if surum > SCHEMA_VERSION:
            raise ValueError(
                f"Bu proje daha yeni bir kayıt biçimi kullanıyor (v{surum}). "
                f"Programın desteklediği en yeni biçim v{SCHEMA_VERSION}."
            )

        sonuc = self.varsayilan_proje_verisi_al()
        sonuc["schema_version"] = SCHEMA_VERSION

        sonuc["_ON_DEGER_"] = normalize_on_deger(veriler.get("_ON_DEGER_"))
        sonuc["_TDTH_"] = normalize_tdth(veriler.get("_TDTH_"))
        sonuc["_IS_AKISI_"] = normalize_is_akisi(
            veriler.get("_IS_AKISI_"),
            eski_proje=("_IS_AKISI_" not in veriler and surum < 3),
        )

        for kod in getattr(self, "veri_alanlari", {}):
            if kod in veriler:
                sonuc[kod] = self._metne_cevir(veriler[kod], kod)

        if "_BINA_" in veriler:
            bina = veriler["_BINA_"]
            if not isinstance(bina, dict):
                raise ValueError("_BINA_ bir JSON nesnesi olmalıdır.")
            hedef_bina = copy.deepcopy(sonuc.get("_BINA_", {}))
            for etiket in getattr(self, "bina_alanlari", {}):
                if etiket in bina:
                    hedef_bina[etiket] = self._metne_cevir(bina[etiket], f"_BINA_.{etiket}")
            sonuc["_BINA_"] = hedef_bina

        for anahtar in (
            "_FORMASYON_",
            "_FORMASYON_METNI_",
            "_MUHENDISLIK_JEOLOJISI_METNI_",
            "_RAPOR_SABLONU_",
            "_TAAHHUT_WORD_SABLONU_",
            "_JEOLOJI_SABLONU_",
            "_LAB_KAYNAK_",
        ):
            if anahtar in veriler:
                sonuc[anahtar] = self._metne_cevir(veriler[anahtar], anahtar)

        kutuphane_bolumu = veriler.get("_JEOLOJI_KUTUPHANE_BOLUMU_", {})
        if not isinstance(kutuphane_bolumu, dict):
            raise ValueError("_JEOLOJI_KUTUPHANE_BOLUMU_ bir JSON nesnesi olmalıdır.")
        kayit_id = kutuphane_bolumu.get("kayit_id")
        if kayit_id not in (None, ""):
            try:
                kayit_id = int(kayit_id)
            except (TypeError, ValueError):
                raise ValueError(
                    "_JEOLOJI_KUTUPHANE_BOLUMU_.kayit_id tam sayı olmalıdır."
                ) from None
            if kayit_id <= 0:
                raise ValueError(
                    "_JEOLOJI_KUTUPHANE_BOLUMU_.kayit_id sıfırdan büyük olmalıdır."
                )
        else:
            kayit_id = None
        sonuc["_JEOLOJI_KUTUPHANE_BOLUMU_"] = {
            "aktif": self._bool_deger(kutuphane_bolumu.get("aktif", False)),
            "kayit_id": kayit_id,
            "bolum_docx_path": self._metne_cevir(
                kutuphane_bolumu.get("bolum_docx_path", ""),
                "_JEOLOJI_KUTUPHANE_BOLUMU_.bolum_docx_path",
            ),
            "bolum_hash": self._metne_cevir(
                kutuphane_bolumu.get("bolum_hash", ""),
                "_JEOLOJI_KUTUPHANE_BOLUMU_.bolum_hash",
            ),
            "uygulanan_genel": self._metne_cevir(
                kutuphane_bolumu.get("uygulanan_genel", ""),
                "_JEOLOJI_KUTUPHANE_BOLUMU_.uygulanan_genel",
            ),
            "uygulanan_inceleme": self._metne_cevir(
                kutuphane_bolumu.get("uygulanan_inceleme", ""),
                "_JEOLOJI_KUTUPHANE_BOLUMU_.uygulanan_inceleme",
            ),
        }

        taahhut = veriler.get("_TAAHHUT_BILGILERI_", {})
        if not isinstance(taahhut, dict):
            raise ValueError("_TAAHHUT_BILGILERI_ bir JSON nesnesi olmalıdır.")
        hedef_taahhut = {kod: "" for kod in TAAHHUT_BILGI_ALANLARI}
        varsayilan_taahhut = sonuc.get("_TAAHHUT_BILGILERI_", {})
        if isinstance(varsayilan_taahhut, dict):
            for kod in TAAHHUT_BILGI_ALANLARI:
                hedef_taahhut[kod] = self._metne_cevir(
                    varsayilan_taahhut.get(kod, ""), f"_TAAHHUT_BILGILERI_.{kod}"
                ).strip()
        for kod in TAAHHUT_BILGI_ALANLARI:
            if kod in taahhut:
                hedef_taahhut[kod] = self._metne_cevir(
                    taahhut[kod], f"_TAAHHUT_BILGILERI_.{kod}"
                ).strip()
        sonuc["_TAAHHUT_BILGILERI_"] = hedef_taahhut

        if "_AC_SEKMELERI_" in veriler:
            ac_verileri = veriler["_AC_SEKMELERI_"]
            if not isinstance(ac_verileri, list):
                raise ValueError("_AC_SEKMELERI_ bir kayıt listesi olmalıdır.")
            temiz_ac = []
            for index, ac_data in enumerate(ac_verileri):
                if not isinstance(ac_data, dict):
                    raise ValueError(f"_AC_SEKMELERI_[{index}] bir JSON nesnesi olmalıdır.")
                temiz_ac.append({
                    "isim": self._metne_cevir(ac_data.get("isim", f"AÇ{index + 1}"), f"AÇ[{index}].isim"),
                    "derinlik": self._metne_cevir(ac_data.get("derinlik", ""), f"AÇ[{index}].derinlik"),
                    "enlem": self._metne_cevir(ac_data.get("enlem", ""), f"AÇ[{index}].enlem"),
                    "boylam": self._metne_cevir(ac_data.get("boylam", ""), f"AÇ[{index}].boylam"),
                    "tarih": self._metne_cevir(ac_data.get("tarih", ""), f"AÇ[{index}].tarih"),
                    "satirlar": self._tablo_satirlarini_dogrula(
                        ac_data.get("satirlar", []), f"_AC_SEKMELERI_[{index}].satirlar"
                    ),
                    "aciklama": self._metne_cevir(ac_data.get("aciklama", ""), f"AÇ[{index}].aciklama"),
                })
            sonuc["_AC_SEKMELERI_"] = temiz_ac

        if "_JEOFIZIK_" in veriler:
            jeo = veriler["_JEOFIZIK_"]
            if not isinstance(jeo, dict):
                raise ValueError("_JEOFIZIK_ bir JSON nesnesi olmalıdır.")
            hedef_jeo = copy.deepcopy(sonuc.get("_JEOFIZIK_", {}))
            hedef_jeo["excel_yolu"] = self._metne_cevir(jeo.get("excel_yolu", ""), "_JEOFIZIK_.excel_yolu")
            hedef_jeo["tree_sis"] = self._tablo_satirlarini_dogrula(
                jeo.get("tree_sis", []), "_JEOFIZIK_.tree_sis"
            )
            dizilim = jeo.get("jeofon_dizilim", {})
            if not isinstance(dizilim, dict):
                raise ValueError("_JEOFIZIK_.jeofon_dizilim bir JSON nesnesi olmalıdır.")
            hedef_dizilim = copy.deepcopy(hedef_jeo.get("jeofon_dizilim", {}))
            for kod, deger in dizilim.items():
                if kod in hedef_dizilim:
                    hedef_dizilim[kod] = self._metne_cevir(deger, f"jeofon_dizilim.{kod}")
            hedef_jeo["jeofon_dizilim"] = hedef_dizilim
            sonuc["_JEOFIZIK_"] = hedef_jeo

        if "_TASIMA_" in veriler:
            tasima = veriler["_TASIMA_"]
            hedef_tasima = copy.deepcopy(sonuc.get("_TASIMA_", {}))
            if isinstance(tasima, str):
                hedef_tasima["secim"] = tasima
            elif isinstance(tasima, dict):
                girdiler = tasima.get("girdiler", {})
                if not isinstance(girdiler, dict):
                    raise ValueError("_TASIMA_.girdiler bir JSON nesnesi olmalıdır.")
                hedef_girdiler = copy.deepcopy(hedef_tasima.get("girdiler", {}))
                for kod, deger in girdiler.items():
                    if kod in getattr(self, "tg_girdiler", {}):
                        hedef_girdiler[kod] = self._metne_cevir(deger, f"_TASIMA_.girdiler.{kod}")
                hedef_tasima["girdiler"] = hedef_girdiler
                for anahtar in ("secim", "qt_nihai", "ks_nihai", "son_qk", "son_qt", "rapor_metni"):
                    if anahtar in tasima:
                        hedef_tasima[anahtar] = self._metne_cevir(tasima[anahtar], f"_TASIMA_.{anahtar}")
                hedef_tasima["varsayim_onayi"] = self._bool_deger(
                    tasima.get("varsayim_onayi", hedef_tasima.get("varsayim_onayi", False))
                )
                hedef_tasima["dayanim_23_uygulandi"] = self._bool_deger(
                    tasima.get(
                        "dayanim_23_uygulandi",
                        hedef_tasima.get("dayanim_23_uygulandi", False),
                    )
                )
                for anahtar in ("dayanim_23_kaynak_c", "dayanim_23_kaynak_phi"):
                    if anahtar in tasima:
                        hedef_tasima[anahtar] = self._metne_cevir(
                            tasima[anahtar], f"_TASIMA_.{anahtar}"
                        )
                if hedef_tasima["dayanim_23_uygulandi"] and not all(
                    str(hedef_tasima.get(anahtar, "")).strip()
                    for anahtar in ("dayanim_23_kaynak_c", "dayanim_23_kaynak_phi")
                ):
                    hedef_tasima["dayanim_23_uygulandi"] = False
                if "yass_var" in tasima:
                    hedef_tasima["yass_var"] = self._bool_deger(tasima.get("yass_var"))
                else:
                    raw_yass = str(hedef_girdiler.get("yass", "")).strip().replace(",", ".")
                    try:
                        hedef_tasima["yass_var"] = bool(raw_yass) and float(raw_yass) < 999.0
                    except ValueError:
                        hedef_tasima["yass_var"] = bool(raw_yass)
                hedef_tasima["rapor_imzasi"] = self._tasima_rapor_imzasini_dogrula(
                    tasima.get("rapor_imzasi")
                )
            else:
                raise ValueError("_TASIMA_ bir JSON nesnesi olmalıdır.")
            sonuc["_TASIMA_"] = hedef_tasima

        for anahtar in ("_LAB_AC_", "_LAB_YN_"):
            if anahtar in veriler:
                sonuc[anahtar] = self._tablo_satirlarini_dogrula(veriler[anahtar], anahtar)

        if "_EKLER_" in veriler:
            ekler = veriler["_EKLER_"]
            if not isinstance(ekler, dict):
                raise ValueError("_EKLER_ bir JSON nesnesi olmalıdır.")
            temiz_ekler = {}
            for kategori in getattr(self, "ek_kategorileri", []):
                kategori_verisi = ekler.get(kategori, [])
                if not isinstance(kategori_verisi, list):
                    raise ValueError(f"_EKLER_.{kategori} bir liste olmalıdır.")
                temiz_ekler[kategori] = []
                for index, ek in enumerate(kategori_verisi):
                    if isinstance(ek, str):
                        temiz_ekler[kategori].append({"baslik": "", "yol": ek})
                    elif isinstance(ek, dict):
                        temiz_ekler[kategori].append({
                            "baslik": self._metne_cevir(ek.get("baslik", ""), f"_EKLER_.{kategori}[{index}].baslik"),
                            "yol": self._metne_cevir(ek.get("yol", ""), f"_EKLER_.{kategori}[{index}].yol"),
                        })
                    else:
                        raise ValueError(f"_EKLER_.{kategori}[{index}] geçersizdir.")
            sonuc["_EKLER_"] = temiz_ekler

        if "__HARITA__" in veriler:
            harita = veriler["__HARITA__"]
            if not isinstance(harita, dict):
                raise ValueError("__HARITA__ bir JSON nesnesi olmalıdır.")
            hedef_harita = copy.deepcopy(sonuc.get("__HARITA__", {}))
            hedef_harita["lat"] = self._sonlu_sayi(
                harita.get("lat", hedef_harita.get("lat", 39.524)),
                39.524,
                "__HARITA__.lat",
                -90.0,
                90.0,
            )
            hedef_harita["lon"] = self._sonlu_sayi(
                harita.get("lon", hedef_harita.get("lon", 26.120)),
                26.120,
                "__HARITA__.lon",
                -180.0,
                180.0,
            )
            zoom = self._sonlu_sayi(
                harita.get("zoom", hedef_harita.get("zoom", 15)),
                15,
                "__HARITA__.zoom",
                0,
                22,
            )
            hedef_harita["zoom"] = int(round(zoom))
            hedef_harita["kml_yolu"] = self._metne_cevir(harita.get("kml_yolu", ""), "__HARITA__.kml_yolu")
            hedef_harita["parsel_haritasi_yolu"] = self._metne_cevir(
                harita.get("parsel_haritasi_yolu", ""),
                "__HARITA__.parsel_haritasi_yolu",
            )
            hedef_harita["parsel_haritasi_geometri_hash"] = self._metne_cevir(
                harita.get("parsel_haritasi_geometri_hash", ""),
                "__HARITA__.parsel_haritasi_geometri_hash",
            )
            hedef_harita["parsel_haritasi_ada"] = self._metne_cevir(
                harita.get("parsel_haritasi_ada", ""),
                "__HARITA__.parsel_haritasi_ada",
            )
            hedef_harita["parsel_haritasi_parsel"] = self._metne_cevir(
                harita.get("parsel_haritasi_parsel", ""),
                "__HARITA__.parsel_haritasi_parsel",
            )
            hedef_harita["parsel_haritasi_kaynak_url"] = self._metne_cevir(
                harita.get("parsel_haritasi_kaynak_url", ""),
                "__HARITA__.parsel_haritasi_kaynak_url",
            )
            hedef_harita["jeoloji_pafta_sonucu"] = self._jeoloji_pafta_sonucunu_dogrula(
                harita.get("jeoloji_pafta_sonucu", {})
            )
            hedef_harita["genel_jeoloji_verisi"] = self._genel_jeoloji_verisini_dogrula(
                harita.get("genel_jeoloji_verisi", {})
            )
            for _, (json_alani, _) in RAPOR_HARITA_ALANLARI.items():
                hedef_harita[json_alani] = self._metne_cevir(
                    harita.get(json_alani, ""),
                    f"__HARITA__.{json_alani}",
                )

            kml_points = harita.get("kml_points", [])
            if not isinstance(kml_points, list):
                raise ValueError("__HARITA__.kml_points bir liste olmalıdır.")
            if len(kml_points) > 100_000:
                raise ValueError("__HARITA__.kml_points en fazla 100000 nokta içerebilir.")
            hedef_harita["kml_points"] = [
                self._koordinat_ciftini_dogrula(nokta, f"__HARITA__.kml_points[{index}]")
                for index, nokta in enumerate(kml_points)
            ]

            sayaclar = harita.get("sayaclar", {})
            if not isinstance(sayaclar, dict):
                raise ValueError("__HARITA__.sayaclar bir JSON nesnesi olmalıdır.")
            temiz_sayaclar = {"AÇ": 1, "YN": 1, "SS": 1}
            for tip in temiz_sayaclar:
                try:
                    temiz_sayaclar[tip] = max(1, int(sayaclar.get(tip, temiz_sayaclar[tip])))
                except (TypeError, ValueError):
                    raise ValueError(f"__HARITA__.sayaclar.{tip} tam sayı olmalıdır.") from None
            hedef_harita["sayaclar"] = temiz_sayaclar

            isaretler = harita.get("isaretler", {})
            if not isinstance(isaretler, dict):
                raise ValueError("__HARITA__.isaretler bir JSON nesnesi olmalıdır.")
            temiz_isaretler = {}
            for isim, veri in isaretler.items():
                if not isinstance(isim, str) or not isinstance(veri, dict):
                    raise ValueError("Harita işaretleri ad/nesne çiftlerinden oluşmalıdır.")
                tip = veri.get("tip")
                if tip not in ("M", "AÇ", "YN", "SS"):
                    raise ValueError(f"{isim} için harita işaret tipi geçersizdir.")
                if tip == "SS":
                    temiz_isaretler[isim] = {
                        "tip": tip,
                        "n1": self._koordinat_ciftini_dogrula(veri.get("n1"), f"{isim}.n1"),
                        "n2": self._koordinat_ciftini_dogrula(veri.get("n2"), f"{isim}.n2"),
                    }
                else:
                    temiz_isaretler[isim] = {
                        "tip": tip,
                        "lat": self._sonlu_sayi(veri.get("lat"), 0.0, f"{isim}.lat", -90.0, 90.0),
                        "lon": self._sonlu_sayi(veri.get("lon"), 0.0, f"{isim}.lon", -180.0, 180.0),
                    }
            hedef_harita["isaretler"] = temiz_isaretler
            sonuc["__HARITA__"] = hedef_harita

        return sonuc

    def kaydedilmemis_degisiklik_var_mi(self):
        if getattr(self, "proje_salt_okunur", False):
            return False
        son_kayit = getattr(self, "son_kayit_verisi", None)
        if son_kayit is None:
            return False
        return son_kayit != self.verileri_topla()

    def degisiklik_gecisine_izin_ver(self, eylem):
        if not self.kaydedilmemis_degisiklik_var_mi():
            return True
        cevap = messagebox.askyesnocancel(
            eylem,
            f"{eylem} işleminden önce kaydedilmemiş değişiklikler var. Şimdi kaydetmek ister misiniz?",
        )
        if cevap is None:
            return False
        if cevap is False:
            return True
        return bool(self.kaydet())

    def proje_dosyasini_ac(self, yol):
        if not self.degisiklik_gecisine_izin_ver("Proje Aç"):
            return False
        onceki_veri = self.verileri_topla()
        onceki_dosya_yolu = getattr(self, "guncel_dosya_yolu", None)
        onceki_son_kayit = copy.deepcopy(getattr(self, "son_kayit_verisi", None))
        onceki_son_projeler = copy.deepcopy(getattr(self, "son_projeler", None))
        onceki_salt_okunur = bool(getattr(self, "proje_salt_okunur", False))
        try:
            onceki_pencere_basligi = self.root.title()
        except Exception:
            onceki_pencere_basligi = None
        yerlestirme_basladi = False
        try:
            tam_yol = os.path.abspath(yol)
            with open(tam_yol, "r", encoding="utf-8-sig") as f:
                veriler = json.load(f)
            veriler = self.proje_yollarini_coz(veriler, tam_yol)
            temiz_veriler = self.proje_verisini_normalize_et(veriler)
            temiz_veriler["__HARITA__"] = harita_verisini_mevcut_dosyalarla_tamamla(
                temiz_veriler.get("__HARITA__", {}),
                tam_yol,
            )
            kaynak_temiz_veriler = copy.deepcopy(temiz_veriler)
            akis = temiz_veriler.get("_IS_AKISI_", {})
            asama_degisti = False
            if akis.get("durum") == "belirlenmedi":
                secilen_asama = self.eski_proje_asama_secimi()
                if not secilen_asama:
                    return False
                akis = is_durumu_degistir(
                    akis,
                    secilen_asama,
                    "Eski proje ilk kez iş takip sistemine alındı.",
                    zorla=True,
                )
                temiz_veriler["_IS_AKISI_"] = akis
                asama_degisti = True

            salt_okunur_ac = False
            if akis.get("durum") == "bitti":
                secim, neden = self.bitmis_proje_acilis_secimi()
                if not secim:
                    return False
                if secim == "duzeltme":
                    akis = bitmis_revizyonu_arsivle(akis, temiz_veriler)
                    temiz_veriler["_IS_AKISI_"] = is_durumu_degistir(
                        akis, "duzeltme_asamasinda", neden
                    )
                    asama_degisti = True
                else:
                    salt_okunur_ac = True

            if hasattr(self, "proje_salt_okunur_ayarla"):
                self.proje_salt_okunur_ayarla(False)
            yerlestirme_basladi = True
            self.verileri_yerlestir(temiz_veriler, dogrulandi=True)
            self.guncel_dosya_yolu = tam_yol
            self.root.title(f"K-1 - {tam_yol}")
            self.son_kayit_verisi = (
                kaynak_temiz_veriler if asama_degisti else copy.deepcopy(self.verileri_topla())
            )
            if hasattr(self, "proje_salt_okunur_ayarla"):
                self.proje_salt_okunur_ayarla(salt_okunur_ac)
            self.son_proje_ekle(tam_yol)
            if hasattr(self, "is_takibi_kaydi_guncelle"):
                self.is_takibi_kaydi_guncelle(tam_yol, kaynak_temiz_veriler)
            if hasattr(self, "durum_mesaji_yaz"):
                self.durum_mesaji_yaz("Proje dosyası açıldı", os.path.basename(tam_yol))
            if hasattr(self, "proje_durum_seridi_guncelle"):
                self.proje_durum_seridi_guncelle()
            return True
        except json.JSONDecodeError as e:
            self.hata_kaydet(f"JSON dosyası bozuk veya boş: {yol}", e)
            messagebox.showerror("Açma Hatası", "Seçilen dosya geçerli bir JSON formatında değil, boş veya bozuk.")
        except Exception as e:
            if yerlestirme_basladi:
                try:
                    self.verileri_yerlestir(onceki_veri, dogrulandi=True)
                except Exception as geri_alma_hatasi:
                    self.hata_kaydet("Başarısız proje açma işlemi geri alınamadı", geri_alma_hatasi)
                self.guncel_dosya_yolu = onceki_dosya_yolu
                self.son_kayit_verisi = onceki_son_kayit
                if onceki_son_projeler is not None:
                    son_projeler_degisti = getattr(self, "son_projeler", None) != onceki_son_projeler
                    self.son_projeler = onceki_son_projeler
                    if son_projeler_degisti:
                        try:
                            self.son_projeleri_kaydet()
                            self.son_projeler_menusunu_guncelle()
                        except Exception as geri_alma_hatasi:
                            self.hata_kaydet("Son proje listesi geri alınamadı", geri_alma_hatasi)
                if onceki_pencere_basligi is not None:
                    try:
                        self.root.title(onceki_pencere_basligi)
                    except Exception as geri_alma_hatasi:
                        self.hata_kaydet("Pencere başlığı geri alınamadı", geri_alma_hatasi)
                if hasattr(self, "proje_salt_okunur_ayarla"):
                    try:
                        self.proje_salt_okunur_ayarla(onceki_salt_okunur)
                    except Exception as geri_alma_hatasi:
                        self.hata_kaydet("Proje düzenleme modu geri alınamadı", geri_alma_hatasi)
            self.hata_kaydet(f"Proje dosyası açılırken hata oluştu: {yol}", e)
            messagebox.showerror("Açma Hatası", f"Dosya okunurken bir hata oluştu:\n{e}")
        return False

    def verileri_topla(self):
        veriler = {"schema_version": SCHEMA_VERSION}
        veriler.update({kod: entry.get() for kod, entry in self.veri_alanlari.items()})
        if hasattr(self, "on_deger_verisini_topla"):
            veriler["_ON_DEGER_"] = self.on_deger_verisini_topla()
        else:
            veriler["_ON_DEGER_"] = copy.deepcopy(getattr(self, "on_deger_verisi", {}))
        veriler["_TDTH_"] = copy.deepcopy(getattr(self, "tdth_verisi", {}))
        veriler["_IS_AKISI_"] = copy.deepcopy(getattr(self, "is_akisi_verisi", {}))
        veriler["_BINA_"] = {k: e.get() for k, e in self.bina_alanlari.items()}
        veriler["_FORMASYON_"] = self.combo_formasyon.get()
        if hasattr(self, "txt_formasyon_rapor"):
            veriler["_FORMASYON_METNI_"] = self.txt_formasyon_rapor.get("1.0", tk.END).strip()
        if hasattr(self, "txt_muhendislik_jeolojisi"):
            veriler["_MUHENDISLIK_JEOLOJISI_METNI_"] = self.txt_muhendislik_jeolojisi.get("1.0", tk.END).strip()

        ac_sekmeleri = []
        for kayit in self.ac_yn_sekme_kayitlari():
            ac_sekmeleri.append(self.ac_yn_kaydi_verisini_oku(kayit))
        veriler["_AC_SEKMELERI_"] = ac_sekmeleri

        veriler["_JEOFIZIK_"] = {
            "excel_yolu": self.jeofizik_excel_yolu_al(),
            "tree_sis": self.jeofizik_koordinatlari_al(),
            "jeofon_dizilim": self.jeofon_dizilim_bilgileri_al(),
        }

        qt_nihai = self.entry_qt_nihai.get() if hasattr(self, "entry_qt_nihai") else ""
        ks_nihai = self.entry_ks_nihai.get() if hasattr(self, "entry_ks_nihai") else ""
        rapor_metni = self.txt_tasima_rapor.get("1.0", tk.END).strip() if hasattr(self, "txt_tasima_rapor") else ""

        veriler["_TASIMA_"] = {
            "secim": self.zemin_kaya_var.get(),
            "girdiler": {k: e.get() for k, e in self.tg_girdiler.items()},
            "qt_nihai": qt_nihai,
            "ks_nihai": ks_nihai,
            "son_qk": getattr(self, "son_qk", "-"),
            "son_qt": getattr(self, "son_qt", "-"),
            "rapor_metni": rapor_metni,
            "varsayim_onayi": bool(self.tasima_varsayim_onayi.get())
            if hasattr(self, "tasima_varsayim_onayi")
            else False,
            "yass_var": bool(self.tasima_yass_var.get())
            if hasattr(self, "tasima_yass_var")
            else False,
            "dayanim_23_uygulandi": bool(self.tasima_dayanim_23_uygulandi.get())
            if hasattr(self, "tasima_dayanim_23_uygulandi")
            else False,
            "dayanim_23_kaynak_c": str(
                getattr(self, "tasima_dayanim_23_kaynak_c", "")
            ),
            "dayanim_23_kaynak_phi": str(
                getattr(self, "tasima_dayanim_23_kaynak_phi", "")
            ),
            "rapor_imzasi": copy.deepcopy(getattr(self, "tasima_rapor_imzasi", None)),
        }

        veriler["__HARITA__"] = {
            "zoom": self.map_widget.zoom,
            "lat": self.map_widget.get_position()[0],
            "lon": self.map_widget.get_position()[1],
            "kml_yolu": getattr(self, "yuklu_kml_yolu", ""),
            "kml_points": copy.deepcopy(getattr(self, "yuklu_kml_points", [])),
            "parsel_haritasi_yolu": getattr(self, "img_parsel_haritasi", "") or "",
            "parsel_haritasi_geometri_hash": getattr(self, "parsel_haritasi_geometri_hash", ""),
            "parsel_haritasi_ada": getattr(self, "parsel_haritasi_ada", ""),
            "parsel_haritasi_parsel": getattr(self, "parsel_haritasi_parsel", ""),
            "parsel_haritasi_kaynak_url": getattr(
                self, "parsel_haritasi_kaynak_url", ""
            ),
            "jeoloji_pafta_sonucu": copy.deepcopy(getattr(self, "jeoloji_pafta_sonucu", {})),
            "genel_jeoloji_verisi": copy.deepcopy(getattr(self, "genel_jeoloji_verisi", {})),
            **{
                json_alani: getattr(self, uygulama_alani, "") or ""
                for uygulama_alani, (json_alani, _) in RAPOR_HARITA_ALANLARI.items()
            },
            "sayaclar": copy.deepcopy(self.harita_nokta_sayaclari),
            "isaretler": {
                isim: copy.deepcopy({k: v for k, v in data.items() if k not in ["marker", "path"]})
                for isim, data in self.harita_isaretleri.items()
            }
        }

        if hasattr(self, "tree_lab_ac"):
            veriler["_LAB_AC_"] = self.lab_ac_satirlari_al()
        if hasattr(self, "tree_lab_yn"):
            veriler["_LAB_YN_"] = self.lab_yn_satirlari_al()
        veriler["_LAB_KAYNAK_"] = (
            self.lbl_lab_excel.cget("text")
            if hasattr(self, "lbl_lab_excel") and self.lbl_lab_excel.cget("text") != "Yok"
            else ""
        )
        veriler["_EKLER_"] = self.ekler_verisini_topla()
        veriler["_TAAHHUT_BILGILERI_"] = self._taahhut_bilgilerini_topla()
        veriler["_RAPOR_SABLONU_"] = getattr(self, "sablon_yolu", "")
        veriler["_TAAHHUT_WORD_SABLONU_"] = self.taahhut_word_sablon_yolu
        veriler["_JEOLOJI_SABLONU_"] = getattr(self, "jeoloji_sablon_yolu", "")
        veriler["_JEOLOJI_KUTUPHANE_BOLUMU_"] = {
            "aktif": bool(getattr(self, "jeoloji_kutuphanesi_bolumu_aktif", False)),
            "kayit_id": getattr(self, "jeoloji_kutuphanesi_kayit_id", None),
            "bolum_docx_path": getattr(self, "jeoloji_kutuphanesi_bolum_yolu", ""),
            "bolum_hash": getattr(self, "jeoloji_kutuphanesi_bolum_hash", ""),
            "uygulanan_genel": getattr(self, "jeoloji_kutuphanesi_uygulanan_genel", ""),
            "uygulanan_inceleme": getattr(self, "jeoloji_kutuphanesi_uygulanan_inceleme", ""),
        }

        return veriler

    def gecici_proje_durumunu_temizle(self):
        after_id = getattr(self, "_harita_yeniden_ciz_after_id", None)
        if after_id is not None:
            try:
                self.root.after_cancel(after_id)
            except Exception:
                pass
            self._harita_yeniden_ciz_after_id = None

        map_widget = getattr(self, "map_widget", None)
        if map_widget is not None:
            for metot_adi in ("delete_all_marker", "delete_all_path", "delete_all_polygon"):
                metot = getattr(map_widget, metot_adi, None)
                if callable(metot):
                    try:
                        metot()
                    except Exception as e:
                        self.hata_kaydet(f"Harita geçici nesnesi temizlenemedi: {metot_adi}", e)

        self.temp_ss_marker = None
        self.ss_ilk_nokta = None
        self.kml_polygon_obj = None
        self.harita_isaretleri = {}
        self.harita_nokta_sayaclari = {"AÇ": 1, "YN": 1, "SS": 1}
        self.yuklu_kml_yolu = ""
        self.yuklu_kml_points = []
        self.jeoloji_pafta_sonucu = {}
        self.genel_jeoloji_verisi = {}
        if hasattr(self, "aktif_harita_araci"):
            self.aktif_harita_araci.set("Yok")

        self.img_mjh = None
        self.img_jeofizik_lok = None
        self.img_jeoloji_lok = None
        self.img_yerbulduru = None
        self.img_parsel_haritasi = None
        self.img_genel_jeoloji = None
        self.img_pga_haritasi = None
        self.parsel_haritasi_geometri_hash = ""
        self.parsel_haritasi_ada = ""
        self.parsel_haritasi_parsel = ""
        self.parsel_haritasi_kaynak_url = ""
        if hasattr(self, "lbl_lab_excel"):
            self.lbl_lab_excel.config(text="Yok")

    def verileri_yerlestir(self, veriler, dogrulandi=False):
        if not dogrulandi:
            veriler = self.proje_verisini_normalize_et(veriler)

        self.gecici_proje_durumunu_temizle()

        self.on_deger_verisi = normalize_on_deger(veriler.get("_ON_DEGER_"))
        self.tdth_verisi = normalize_tdth(veriler.get("_TDTH_"))
        self.is_akisi_verisi = normalize_is_akisi(veriler.get("_IS_AKISI_"))

        for kod, entry in self.veri_alanlari.items():
            entry.delete(0, tk.END)
            entry.insert(0, veriler.get(kod, ""))

        formasyon_secimi = veriler.get("_FORMASYON_", "Seçiniz...")
        formasyon_metni = veriler.get("_FORMASYON_METNI_", "")
        if formasyon_secimi and formasyon_secimi not in self.formasyonlar:
            self.formasyonlar.append(formasyon_secimi)
            self.combo_formasyon.configure(values=self.formasyonlar)
        if formasyon_secimi and formasyon_metni:
            self.formasyon_metinleri[formasyon_secimi] = formasyon_metni
        self.combo_formasyon.set(formasyon_secimi)
        if hasattr(self, "txt_formasyon_rapor"):
            self.txt_formasyon_rapor.delete("1.0", tk.END)
            self.txt_formasyon_rapor.insert("1.0", veriler.get("_FORMASYON_METNI_", ""))
        if hasattr(self, "txt_muhendislik_jeolojisi"):
            self.txt_muhendislik_jeolojisi.delete("1.0", tk.END)
            self.txt_muhendislik_jeolojisi.insert("1.0", veriler.get("_MUHENDISLIK_JEOLOJISI_METNI_", ""))
        bina_verileri = veriler.get("_BINA_", {})
        for k, e in self.bina_alanlari.items():
            e.delete(0, tk.END)
            e.insert(0, bina_verileri.get(k, ""))

        self.ac_yn_sekmelerini_temizle()

        for index, ac_data in enumerate(veriler.get("_AC_SEKMELERI_", []), start=1):
            sekme_adi = ac_data.get("isim") or f"AÇ{index}"
            sekme = self.cukur_sekmesi_ekle(
                sekme_adi,
                ac_data.get("enlem", ""),
                ac_data.get("boylam", ""),
                ac_data.get("tarih", ""),
            )
            kayit = self.ac_yn_sekme_bilgisi(sekme)
            if not kayit:
                continue

            kayit["derinlik_entry"].delete(0, tk.END)
            kayit["derinlik_entry"].insert(0, ac_data.get("derinlik", ""))

            tree = kayit["tree"]
            for item in tree.get_children():
                tree.delete(item)
            for satir in ac_data.get("satirlar", []):
                tree.insert("", "end", values=satir)
            self.stripe_tree(tree)

            kayit["aciklama_text"].delete("1.0", tk.END)
            kayit["aciklama_text"].insert("1.0", ac_data.get("aciklama", ""))

        jeo = veriler.get("_JEOFIZIK_", {})
        self.jeofizik_excel_yolu_ayarla(jeo.get("excel_yolu", ""))
        self.jeofizik_koordinatlari_yerlestir(jeo.get("tree_sis", []))
        self.jeofon_dizilim_bilgileri_yerlestir(jeo.get("jeofon_dizilim", {}))

        self.lab_ac_satirlari_yerlestir(veriler.get("_LAB_AC_", []))
        self.lab_yn_satirlari_yerlestir(veriler.get("_LAB_YN_", []))
        if hasattr(self, "lbl_lab_excel"):
            self.lbl_lab_excel.config(text=veriler.get("_LAB_KAYNAK_", "") or "Yok")

        if hasattr(self, "on_deger_verisini_yerlestir"):
            self.on_deger_verisini_yerlestir(
                self.on_deger_verisi,
                self.tdth_verisi,
                self.is_akisi_verisi,
            )

        self.ekler_verisini_yerlestir(veriler.get("_EKLER_", {}))
        self._taahhut_bilgilerini_yerlestir(veriler.get("_TAAHHUT_BILGILERI_", {}))
        self.sablon_yolu = veriler.get("_RAPOR_SABLONU_", "")
        if hasattr(self, "lbl_sablon"):
            sablon_metni = f"Şablon: {self.sablon_yolu}" if self.sablon_yolu else "Seçilen Şablon: Yok"
            self.lbl_sablon.config(text=sablon_metni)

        self.taahhut_word_sablon_yolu = veriler.get("_TAAHHUT_WORD_SABLONU_", "")
        if hasattr(self, "lbl_taahhut_word"):
            text = self.taahhut_word_sablon_yolu if self.taahhut_word_sablon_yolu else "Seçilmedi"
            self.lbl_taahhut_word.config(text=text)
        self.jeoloji_sablon_yolu = veriler.get("_JEOLOJI_SABLONU_", "")
        kutuphane_bolumu = veriler.get("_JEOLOJI_KUTUPHANE_BOLUMU_", {})
        if not isinstance(kutuphane_bolumu, dict):
            kutuphane_bolumu = {}
        self.jeoloji_kutuphanesi_bolumu_aktif = bool(kutuphane_bolumu.get("aktif", False))
        self.jeoloji_kutuphanesi_kayit_id = kutuphane_bolumu.get("kayit_id")
        self.jeoloji_kutuphanesi_bolum_yolu = str(
            kutuphane_bolumu.get("bolum_docx_path") or ""
        )
        self.jeoloji_kutuphanesi_bolum_hash = str(kutuphane_bolumu.get("bolum_hash") or "")
        self.jeoloji_kutuphanesi_uygulanan_genel = str(
            kutuphane_bolumu.get("uygulanan_genel") or ""
        )
        self.jeoloji_kutuphanesi_uygulanan_inceleme = str(
            kutuphane_bolumu.get("uygulanan_inceleme") or ""
        )
        if hasattr(self, "lbl_jeoloji_sablon"):
            self.lbl_jeoloji_sablon.config(text=self.jeoloji_sablon_etiket_metni())

        tasima = veriler.get("_TASIMA_", {})
        self.zemin_kaya_var.set(tasima.get("secim", "zemin"))
        girdiler = tasima.get("girdiler", {})
        for k, entry in self.tg_girdiler.items():
            if k in {"ks_carpani", "yass"}:
                entry.configure(state="normal")
            entry.delete(0, tk.END)
            entry.insert(0, girdiler.get(k, ""))
        if hasattr(self, "tasima_varsayim_onayi"):
            self.tasima_varsayim_onayi.set(bool(tasima.get("varsayim_onayi", False)))
        if hasattr(self, "tasima_yass_var"):
            if "yass_var" in tasima:
                yass_var = bool(tasima.get("yass_var", False))
            else:
                # Eski projelerde 999/999.0, YASS bulunmadığı anlamında kullanılıyordu.
                raw_yass = str(girdiler.get("yass", "")).strip().replace(",", ".")
                try:
                    yass_var = bool(raw_yass) and float(raw_yass) < 999.0
                except ValueError:
                    yass_var = bool(raw_yass)
            self.tasima_yass_var.set(yass_var)
        if hasattr(self, "tasima_dayanim_23_uygulandi"):
            self.tasima_dayanim_23_uygulandi.set(
                bool(tasima.get("dayanim_23_uygulandi", False))
            )
        self.tasima_dayanim_23_kaynak_c = str(
            tasima.get("dayanim_23_kaynak_c", "")
        )
        self.tasima_dayanim_23_kaynak_phi = str(
            tasima.get("dayanim_23_kaynak_phi", "")
        )
        if hasattr(self, "tasima_ekran_guncelle"):
            self.tasima_ekran_guncelle()

        if hasattr(self, "entry_qt_nihai"):
            self.entry_qt_nihai.delete(0, tk.END)
            self.entry_qt_nihai.insert(0, tasima.get("qt_nihai", ""))
        if hasattr(self, "entry_ks_nihai"):
            self.entry_ks_nihai.delete(0, tk.END)
            self.entry_ks_nihai.insert(0, tasima.get("ks_nihai", ""))
        self.son_qk = tasima.get("son_qk", "-")
        self.son_qt = tasima.get("son_qt", "-")
        if hasattr(self, "lbl_sonuc"):
            self.lbl_sonuc.config(text=str(self.son_qk) if str(self.son_qk).strip() else "-")
        if hasattr(self, "txt_tasima_rapor"):
            self.txt_tasima_rapor.delete("1.0", tk.END)
            self.txt_tasima_rapor.insert("1.0", tasima.get("rapor_metni", ""))
        self.tasima_rapor_imzasi = self._tasima_rapor_imzasini_dogrula(
            tasima.get("rapor_imzasi")
        )

        hd = veriler.get("__HARITA__", {})
        self.map_widget.set_position(hd.get("lat", 39.524), hd.get("lon", 26.120))
        self.map_widget.set_zoom(hd.get("zoom", 15))
        self.yuklu_kml_yolu = hd.get("kml_yolu", "")
        self.yuklu_kml_points = copy.deepcopy(hd.get("kml_points", []))
        self.jeoloji_pafta_sonucu = copy.deepcopy(hd.get("jeoloji_pafta_sonucu", {}))
        self.genel_jeoloji_verisi = copy.deepcopy(hd.get("genel_jeoloji_verisi", {}))
        self.img_parsel_haritasi = hd.get("parsel_haritasi_yolu", "") or None
        for uygulama_alani, (json_alani, _) in RAPOR_HARITA_ALANLARI.items():
            setattr(self, uygulama_alani, hd.get(json_alani, "") or None)
        self.parsel_haritasi_geometri_hash = hd.get("parsel_haritasi_geometri_hash", "")
        self.parsel_haritasi_ada = hd.get("parsel_haritasi_ada", "")
        self.parsel_haritasi_parsel = hd.get("parsel_haritasi_parsel", "")
        self.parsel_haritasi_kaynak_url = hd.get("parsel_haritasi_kaynak_url", "")
        self.harita_nokta_sayaclari = copy.deepcopy(hd.get("sayaclar", {"AÇ": 1, "YN": 1, "SS": 1}))
        self.harita_isaretleri = copy.deepcopy(hd.get("isaretler", {}))

        self.root.update_idletasks()

        def gecikmeli_cizim():
            self._harita_yeniden_ciz_after_id = None
            try:
                if self.yuklu_kml_points:
                    tuple_points = [tuple(pt) for pt in self.yuklu_kml_points]
                    self.kml_polygon_obj = self.map_widget.set_polygon(
                        tuple_points,
                        fill_color=None,
                        outline_color=CALISAN_PARSEL_SINIR_RENGI,
                        border_width=CALISAN_PARSEL_SINIR_KALINLIGI,
                    )
                self.sil_ve_yeniden_ciz(True, True, True)
                if hasattr(self, "jeoloji_harita_katmanini_yenile"):
                    self.jeoloji_harita_katmanini_yenile(zorla=True)
                if hasattr(self, "jeoloji_pafta_durumunu_guncelle"):
                    self.jeoloji_pafta_durumunu_guncelle()
                if hasattr(self, "genel_jeoloji_durumunu_guncelle"):
                    self.genel_jeoloji_durumunu_guncelle()
            except Exception as e:
                self.hata_kaydet("Proje haritası yeniden çizilemedi", e)

        self._harita_yeniden_ciz_after_id = self.root.after(400, gecikmeli_cizim)

    def mojibake_puani(self, metin):
        isaretler = "ÃƒÃ„Ã…Ã‚Ã¢"
        puan = sum(metin.count(ch) for ch in isaretler)
        puan += sum(1 for ch in metin if 0x80 <= ord(ch) <= 0x9F)
        return puan

    def mojibake_metin_onar(self, metin):
        if not isinstance(metin, str):
            return metin, 0

        sonuc = metin
        if self.mojibake_puani(sonuc) > 0:
            for encoding in ("cp1252", "latin1"):
                try:
                    aday = sonuc.encode(encoding).decode("utf-8")
                except UnicodeError:
                    continue
                if aday != sonuc and self.mojibake_puani(aday) <= self.mojibake_puani(sonuc):
                    sonuc = aday
                    break

        yama_tablosu = {
            "\u00c3\u2021": "\u00c7", "\u00c3\u0087": "\u00c7",
            "\u00c3\u00a7": "\u00e7",
            "\u00c4\u017e": "\u011e", "\u00c4\u009e": "\u011e",
            "\u00c4\u0178": "\u011f", "\u00c4\u009f": "\u011f",
            "\u00c4\u00b0": "\u0130",
            "\u00c4\u00b1": "\u0131",
            "\u00c3\u2013": "\u00d6", "\u00c3\u0096": "\u00d6",
            "\u00c3\u00b6": "\u00f6",
            "\u00c5\u017e": "\u015e", "\u00c5\u009e": "\u015e",
            "\u00c5\u0178": "\u015f", "\u00c5\u009f": "\u015f",
            "\u00c3\u0153": "\u00dc", "\u00c3\u009c": "\u00dc",
            "\u00c3\u00bc": "\u00fc",
            "\u00c2\u00b0": "\u00b0",
            "\u00e2\u20ac\u0153": "\u201c", "\u00e2\u20ac\u009d": "\u201d",
            "\u00e2\u20ac\u02dc": "\u2018", "\u00e2\u20ac\u2122": "\u2019",
            "\u00e2\u20ac\u201c": "\u2013", "\u00e2\u20ac\u201d": "\u2014",
        }
        for bozuk, dogru in yama_tablosu.items():
            sonuc = sonuc.replace(bozuk, dogru)

        return sonuc, 1 if sonuc != metin else 0

    def json_verisini_karakter_onar(self, veri):
        if isinstance(veri, dict):
            yeni = {}
            degisim = 0
            for anahtar, deger in veri.items():
                if isinstance(anahtar, str):
                    yeni_anahtar, anahtar_degisim = self.mojibake_metin_onar(anahtar)
                else:
                    yeni_anahtar, anahtar_degisim = anahtar, 0
                yeni_deger, deger_degisim = self.json_verisini_karakter_onar(deger)
                if yeni_anahtar in yeni and yeni_anahtar != anahtar:
                    raise ValueError(
                        f"JSON onarımında anahtar çakışması oluştu: {anahtar!r} -> {yeni_anahtar!r}"
                    )
                yeni[yeni_anahtar] = yeni_deger
                degisim += anahtar_degisim + deger_degisim
            return yeni, degisim
        if isinstance(veri, list):
            yeni_liste = []
            degisim = 0
            for item in veri:
                yeni_item, item_degisim = self.json_verisini_karakter_onar(item)
                yeni_liste.append(yeni_item)
                degisim += item_degisim
            return yeni_liste, degisim
        if isinstance(veri, str):
            return self.mojibake_metin_onar(veri)
        return veri, 0

    def json_onarim_yedek_yolu(self, yol):
        base, ext = os.path.splitext(yol)
        ext = ext or ".json"
        yedek = f"{base}.onceki_yedek{ext}"
        if not os.path.exists(yedek):
            return yedek
        zaman = datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        yedek = f"{base}.onceki_yedek_{zaman}{ext}"
        sayac = 1
        while os.path.exists(yedek):
            yedek = f"{base}.onceki_yedek_{zaman}_{sayac}{ext}"
            sayac += 1
        return yedek

    def eski_json_karakterlerini_onar(self):
        yol = filedialog.askopenfilename(filetypes=[("JSON Dosyası", "*.json")])
        if not yol:
            return

        try:
            with open(yol, "r", encoding="utf-8-sig") as f:
                veriler = json.load(f)

            onarilmis, degisim_sayisi = self.json_verisini_karakter_onar(veriler)
            if degisim_sayisi == 0:
                messagebox.showinfo("Onarım Gerekmedi", "Seçilen kayıt dosyasında bozuk Türkçe karakter bulunamadı.")
                return

            onarilmis = self.proje_verisini_normalize_et(onarilmis)
            yedek_yolu = self.json_onarim_yedek_yolu(yol)
            self.atomik_dosya_kopyala(yol, yedek_yolu)
            self.atomik_json_yaz(yol, onarilmis)

            logger.info(
                "JSON karakter onarımı tamamlandı. Dosya: %s, Yedek: %s, Değişen alan: %s",
                yol,
                yedek_yolu,
                degisim_sayisi,
            )
            mesaj = (
                f"Kayıt dosyası onarıldı.\n\n"
                f"Değişen alan sayısı: {degisim_sayisi}\n"
                f"Yedek dosya:\n{yedek_yolu}\n\n"
                "Onarılan dosyayı şimdi açmak ister misiniz?"
            )
            if messagebox.askyesno("Onarım Tamamlandı", mesaj):
                self.proje_dosyasini_ac(yol)
        except json.JSONDecodeError as e:
            self.hata_kaydet(f"JSON karakter onarımı için seçilen dosya bozuk veya boş: {yol}", e)
            messagebox.showerror("Onarım Hatası", "Seçilen dosya geçerli bir JSON formatında değil, boş veya bozuk.")
        except Exception as e:
            self.hata_kaydet(f"JSON karakter onarımı başarısız: {yol}", e)
            messagebox.showerror("Onarım Hatası", f"Kayıt dosyası onarılamadı:\n{e}")

    def yeni_dosya(self):
        if not self.degisiklik_gecisine_izin_ver("Yeni Proje"):
            return False
        onceki_veri = self.verileri_topla()
        try:
            if hasattr(self, "proje_salt_okunur_ayarla"):
                self.proje_salt_okunur_ayarla(False)
            self.verileri_yerlestir(self.varsayilan_proje_verisi_al(), dogrulandi=True)
            self.guncel_dosya_yolu = None
            self.root.title("K-1 - Yeni Proje")
            self.son_kayit_verisi = copy.deepcopy(self.verileri_topla())
            if hasattr(self, "proje_durum_seridi_guncelle"):
                self.proje_durum_seridi_guncelle(kaydedilmedi=False)
            if hasattr(self, "durum_mesaji_yaz"):
                self.durum_mesaji_yaz("Yeni proje hazırlandı", "Yeni proje")
            return True
        except Exception as e:
            try:
                self.verileri_yerlestir(onceki_veri, dogrulandi=True)
            except Exception as geri_alma_hatasi:
                self.hata_kaydet("Başarısız yeni proje işlemi geri alınamadı", geri_alma_hatasi)
            self.hata_kaydet("Yeni proje hazırlanamadı", e)
            messagebox.showerror("Yeni Proje Hatası", f"Yeni proje güvenli biçimde hazırlanamadı:\n{e}")
            return False

    def dosya_ac(self):
        yol = filedialog.askopenfilename(filetypes=[("JSON Dosyası", "*.json")])
        if yol:
            self.proje_dosyasini_ac(yol)

    def kaydet(self):
        if getattr(self, "proje_salt_okunur", False):
            messagebox.showwarning(
                "İzleme Modu",
                "Bu proje izleme modunda açıldı. Değişiklik kaydedilemez.",
            )
            return False
        if self.guncel_dosya_yolu:
            try:
                veriler = self.verileri_topla()
                yedek_yolu = ""
                if os.path.exists(self.guncel_dosya_yolu):
                    yedek_yolu = self.otomatik_yedek_olustur(self.guncel_dosya_yolu)
                    if not yedek_yolu:
                        raise OSError("Mevcut proje yedeklenemedi; asıl dosya korunarak kayıt iptal edildi.")
                kayit_verileri = self.proje_yollarini_kayda_hazirla(veriler, self.guncel_dosya_yolu)
                self.atomik_json_yaz(self.guncel_dosya_yolu, kayit_verileri)
                self.son_kayit_verisi = copy.deepcopy(veriler)
                self.son_proje_ekle(self.guncel_dosya_yolu)
                if hasattr(self, "is_takibi_kaydi_guncelle"):
                    self.is_takibi_kaydi_guncelle(self.guncel_dosya_yolu, veriler)
                if hasattr(self, "proje_durum_seridi_guncelle"):
                    self.proje_durum_seridi_guncelle(kaydedilmedi=False)
                if hasattr(self, "durum_mesaji_yaz"):
                    mesaj = "Proje kaydedildi"
                    if yedek_yolu:
                        mesaj += " / yedek alındı"
                    self.durum_mesaji_yaz(mesaj, os.path.basename(self.guncel_dosya_yolu))
                messagebox.showinfo("Başarılı", "Proje Kaydedildi.")
                return True
            except Exception as e:
                self.hata_kaydet(f"Proje dosyası kaydedilirken hata oluştu: {self.guncel_dosya_yolu}", e)
                messagebox.showerror("Kaydetme Hatası", f"Dosya kaydedilirken bir hata oluştu:\n{e}")
                return False
        else:
            return self.farkli_kaydet()

    def farkli_kaydet(self):
        if getattr(self, "proje_salt_okunur", False):
            messagebox.showwarning(
                "İzleme Modu",
                "Bu proje izleme modunda açıldı. Farklı Kaydet kullanılamaz.",
            )
            return False
        yol = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON Dosyası", "*.json")])
        if not yol:
            return False
        try:
            tam_yol = os.path.abspath(yol)
            yedek_yolu = ""
            if os.path.exists(tam_yol):
                yedek_yolu = self.otomatik_yedek_olustur(tam_yol)
                if not yedek_yolu:
                    raise OSError("Hedef dosya yedeklenemedi; üzerine yazma iptal edildi.")
            veriler = self.verileri_topla()
            kayit_verileri = self.proje_yollarini_kayda_hazirla(veriler, tam_yol)
            self.atomik_json_yaz(tam_yol, kayit_verileri)
            self.guncel_dosya_yolu = tam_yol
            self.root.title(f"K-1 - {tam_yol}")
            self.son_kayit_verisi = copy.deepcopy(veriler)
            self.son_proje_ekle(tam_yol)
            if hasattr(self, "is_takibi_kaydi_guncelle"):
                self.is_takibi_kaydi_guncelle(tam_yol, veriler)
            if hasattr(self, "proje_durum_seridi_guncelle"):
                self.proje_durum_seridi_guncelle(kaydedilmedi=False)
            if hasattr(self, "durum_mesaji_yaz"):
                mesaj = "Proje farklı kaydedildi"
                if yedek_yolu:
                    mesaj += " / eski hedef yedeklendi"
                self.durum_mesaji_yaz(mesaj, os.path.basename(tam_yol))
            return True
        except Exception as e:
            self.hata_kaydet(f"Proje dosyası farklı kaydedilirken hata oluştu: {yol}", e)
            messagebox.showerror("Kaydetme Hatası", f"Dosya kaydedilirken bir hata oluştu:\n{e}")
            return False
