import math
import os
from pathlib import Path
import stat
import threading
import tkinter as tk
import xml.etree.ElementTree as ET
from tkinter import filedialog, messagebox, ttk

from harita_renkleri import (
    CALISAN_PARSEL_SINIR_KALINLIGI,
    CALISAN_PARSEL_SINIR_RENGI,
    SS_SERIM_KALINLIGI,
    SS_SERIM_KESIK_DESENI,
    SS_SERIM_RENGI,
)
from harita_hassas_zoom import KESIRLI_ZOOM_ADIMI, kesirli_zoom_degerini_duzelt
from tkgm_kml import tkgm_kml_dosya_adi, tkgm_parsel_kml_olustur


KML_AZAMI_DOSYA_BOYUTU = 10 * 1024 * 1024
KML_AZAMI_NOKTA_SAYISI = 20_000
PARSEL_ETKILESIMLI_KENAR_PAYI = 0.10
PARSEL_KADRAJ_MIN_BOYUT = 64
PARSEL_KADRAJ_BEKLEME_DENEMESI = 10


def _parsel_kadraj_mercator(enlem, boylam):
    enlem = max(-85.05112878, min(85.05112878, float(enlem)))
    x = (float(boylam) + 180.0) / 360.0
    y = (1.0 - math.asinh(math.tan(math.radians(enlem))) / math.pi) / 2.0
    return x, y


def parsel_etkilesimli_kadraj_hesapla(
    noktalar,
    genislik,
    yukseklik,
    kenar_payi=PARSEL_ETKILESIMLI_KENAR_PAYI,
    min_zoom=1,
    max_zoom=22,
):
    """Etkileşimli harita için merkez, tam karo zoom'u ve ideal zoom'u hesaplar.

    Adapter widget, ``round(self.zoom)`` seviyesindeki kaynak karoyu kesirli
    efektif karo boyutuna ölçekler. Bu nedenle burada güvenli kenar payını
    koruyan en yüksek 0,25 adımlı zoom döndürülür; ``ideal_zoom`` ise ölçülen
    ham değeri teşhis ve test için korur.
    """
    temiz = []
    for nokta in noktalar or []:
        if not isinstance(nokta, (list, tuple)) or len(nokta) < 2:
            raise ValueError("Parsel koordinatı [enlem, boylam] biçiminde olmalıdır.")
        try:
            enlem, boylam = float(nokta[0]), float(nokta[1])
        except (TypeError, ValueError):
            raise ValueError("Parsel koordinatları sayısal olmalıdır.") from None
        if not math.isfinite(enlem) or not math.isfinite(boylam):
            raise ValueError("Parsel koordinatları sonlu olmalıdır.")
        if not -90.0 <= enlem <= 90.0 or not -180.0 <= boylam <= 180.0:
            raise ValueError("Parsel koordinatı WGS84 aralığının dışında.")
        temiz.append((enlem, boylam))
    if len({(round(p[0], 10), round(p[1], 10)) for p in temiz}) < 3:
        raise ValueError("Parsel kadrajı için en az üç farklı koordinat gerekir.")

    try:
        genislik = max(1.0, float(genislik))
        yukseklik = max(1.0, float(yukseklik))
    except (TypeError, ValueError):
        raise ValueError("Parsel kadrajı için geçerli harita boyutu gerekir.") from None
    if not math.isfinite(genislik) or not math.isfinite(yukseklik):
        raise ValueError("Parsel kadrajı harita boyutları sonlu olmalıdır.")

    try:
        pay = float(kenar_payi)
    except (TypeError, ValueError):
        pay = PARSEL_ETKILESIMLI_KENAR_PAYI
    if not math.isfinite(pay):
        pay = PARSEL_ETKILESIMLI_KENAR_PAYI
    pay = min(0.30, max(0.05, pay))

    try:
        min_zoom = max(0, int(min_zoom))
        max_zoom = max(min_zoom, int(max_zoom))
    except (TypeError, ValueError):
        min_zoom, max_zoom = 1, 22

    mercator = [_parsel_kadraj_mercator(enlem, boylam) for enlem, boylam in temiz]
    xs = [nokta[0] for nokta in mercator]
    ys = [nokta[1] for nokta in mercator]
    span_x = max(xs) - min(xs)
    span_y = max(ys) - min(ys)
    kullanilabilir_w = genislik * (1.0 - 2.0 * pay)
    kullanilabilir_h = yukseklik * (1.0 - 2.0 * pay)
    zoom_x = math.log2(kullanilabilir_w / (256.0 * span_x)) if span_x > 0 else float(max_zoom)
    zoom_y = math.log2(kullanilabilir_h / (256.0 * span_y)) if span_y > 0 else float(max_zoom)
    ideal_zoom = min(zoom_x, zoom_y)
    ideal_zoom = max(float(min_zoom), min(float(max_zoom), ideal_zoom))
    zoom = math.floor((ideal_zoom + 1e-9) / KESIRLI_ZOOM_ADIMI) * KESIRLI_ZOOM_ADIMI
    zoom = kesirli_zoom_degerini_duzelt(
        zoom,
        min_zoom=min_zoom,
        max_zoom=max_zoom,
        adim=KESIRLI_ZOOM_ADIMI,
    )

    merkez = (
        math.degrees(math.atan(math.sinh(math.pi * (1.0 - 2.0 * ((min(ys) + max(ys)) / 2.0))))),
        ((min(xs) + max(xs)) / 2.0 * 360.0) - 180.0,
    )
    return merkez, zoom, ideal_zoom


def isim_bazli_birlestirme_plani(mevcut_isimler, hedef_isimler):
    """İsim bazlı, sırası kararlı bir ekle/güncelle/orphan planı döndürür."""
    mevcut = [str(isim) for isim in mevcut_isimler]
    hedef = [str(isim) for isim in hedef_isimler]
    mevcut_kume = set(mevcut)
    hedef_kume = set(hedef)
    return {
        "guncellenecek": [isim for isim in hedef if isim in mevcut_kume],
        "eklenecek": [isim for isim in hedef if isim not in mevcut_kume],
        "orphan": [isim for isim in mevcut if isim not in hedef_kume],
    }


def _xml_yerel_ad(tag):
    return str(tag).rsplit("}", 1)[-1]


def _kml_koordinat_metnini_ayristir(metin, azami_nokta_sayisi):
    koordinatlar = []
    for parca in (metin or "").split():
        alanlar = parca.split(",")
        if len(alanlar) < 2:
            raise ValueError("KML koordinat kaydı enlem ve boylam içermiyor.")
        try:
            boylam = float(alanlar[0])
            enlem = float(alanlar[1])
        except ValueError as exc:
            raise ValueError(f"KML içinde sayısal olmayan koordinat var: {parca}") from exc
        if not math.isfinite(enlem) or not math.isfinite(boylam):
            raise ValueError("KML koordinatları sonlu sayılar olmalıdır.")
        if not -90.0 <= enlem <= 90.0 or not -180.0 <= boylam <= 180.0:
            raise ValueError(f"KML koordinatı geçerli WGS84 aralığının dışında: {parca}")
        koordinatlar.append((enlem, boylam))
        if len(koordinatlar) > azami_nokta_sayisi:
            raise ValueError(f"KML poligonu en fazla {azami_nokta_sayisi} nokta içerebilir.")

    benzersiz = {(round(enlem, 12), round(boylam, 12)) for enlem, boylam in koordinatlar}
    if len(benzersiz) < 3:
        raise ValueError("KML poligonu en az üç farklı koordinat içermelidir.")
    return koordinatlar


def kml_poligon_koordinatlarini_oku(
    dosya_yolu,
    azami_dosya_boyutu=KML_AZAMI_DOSYA_BOYUTU,
    azami_nokta_sayisi=KML_AZAMI_NOKTA_SAYISI,
):
    """Sınırlı boyuttaki KML'den ilk gerçek Polygon dış halkasını güvenle okur."""
    poligonlar = kml_poligonlarini_oku(
        dosya_yolu,
        azami_dosya_boyutu=azami_dosya_boyutu,
        azami_nokta_sayisi=azami_nokta_sayisi,
    )
    return poligonlar[0]["noktalar"]


def kml_poligonlarini_oku(
    dosya_yolu,
    azami_dosya_boyutu=KML_AZAMI_DOSYA_BOYUTU,
    azami_nokta_sayisi=KML_AZAMI_NOKTA_SAYISI,
):
    """KML içindeki bütün Placemark Polygon dış halkalarını okur.

    Sonuç koordinatları tkintermapview ile uyumlu ``(enlem, boylam)``
    sırasındadır. MultiGeometry içindeki Polygon öğeleri de ayrı halkalar
    olarak döndürülür.
    """
    boyut = os.path.getsize(dosya_yolu)
    if boyut > azami_dosya_boyutu:
        raise ValueError(
            f"KML dosyası çok büyük ({boyut / (1024 * 1024):.1f} MB). "
            f"Üst sınır {azami_dosya_boyutu / (1024 * 1024):.0f} MB'dir."
        )
    with open(dosya_yolu, "rb") as dosya:
        icerik = dosya.read(azami_dosya_boyutu + 1)
    if b"<!doctype" in icerik.lower() or b"<!entity" in icerik.lower():
        raise ValueError("DTD veya harici varlık içeren KML dosyaları güvenlik nedeniyle desteklenmiyor.")
    try:
        kok = ET.fromstring(icerik)
    except ET.ParseError as exc:
        raise ValueError(f"KML XML yapısı geçersiz: {exc}") from exc

    sonuc = []
    toplam_nokta = 0
    placemarkler = [eleman for eleman in kok.iter() if _xml_yerel_ad(eleman.tag) == "Placemark"]
    kokler = placemarkler or [kok]
    for placemark in kokler:
        ad = ""
        aciklama = ""
        for child in placemark:
            yerel_ad = _xml_yerel_ad(child.tag)
            if yerel_ad == "name" and not ad:
                ad = (child.text or "").strip()
            elif yerel_ad == "description" and not aciklama:
                aciklama = (child.text or "").strip()
        for polygon in placemark.iter():
            if _xml_yerel_ad(polygon.tag) != "Polygon":
                continue
            noktalar = None
            for dis_sinir in polygon:
                if _xml_yerel_ad(dis_sinir.tag) != "outerBoundaryIs":
                    continue
                for halka in dis_sinir:
                    if _xml_yerel_ad(halka.tag) != "LinearRing":
                        continue
                    for eleman in halka:
                        if _xml_yerel_ad(eleman.tag) == "coordinates" and (eleman.text or "").strip():
                            kalan = azami_nokta_sayisi - toplam_nokta
                            if kalan < 3:
                                raise ValueError(
                                    f"KML dosyası toplam en fazla {azami_nokta_sayisi} nokta içerebilir."
                                )
                            noktalar = _kml_koordinat_metnini_ayristir(eleman.text, kalan)
                            break
                    if noktalar:
                        break
                if noktalar:
                    break
            if noktalar:
                toplam_nokta += len(noktalar)
                sonuc.append({"ad": ad, "aciklama": aciklama, "noktalar": noktalar})
    if not sonuc:
        raise ValueError("KML içinde geçerli bir Polygon koordinat halkası bulunamadı.")
    return sonuc


def kml_dosyalari_bul(paths):
    """Dosya ve klasörlerden görünür, makul boyutlu KML adaylarını bulur."""

    def gizli_mi(path):
        if any(part.startswith(".") for part in path.parts if part not in (".", "..")):
            return True
        try:
            attributes = getattr(path.stat(), "st_file_attributes", 0)
            return bool(attributes & getattr(stat, "FILE_ATTRIBUTE_HIDDEN", 2))
        except OSError:
            return True

    adaylar = []
    for raw_path in paths or ():
        path = Path(raw_path)
        if path.is_dir():
            for root, directories, filenames in os.walk(
                path, topdown=True, onerror=lambda _error: None, followlinks=False
            ):
                root_path = Path(root)
                directories[:] = sorted(
                    (name for name in directories if not gizli_mi(root_path / name)),
                    key=str.casefold,
                )
                for filename in sorted(filenames, key=str.casefold):
                    adaylar.append(root_path / filename)
        else:
            adaylar.append(path)

    sonuc = []
    gorulen = set()
    for aday in adaylar:
        if aday.suffix.lower() != ".kml" or gizli_mi(aday):
            continue
        try:
            if not aday.is_file() or aday.stat().st_size > KML_AZAMI_DOSYA_BOYUTU:
                continue
            cozulmus = aday.resolve(strict=True)
        except (OSError, ValueError):
            continue
        anahtar = os.path.normcase(str(cozulmus)).casefold()
        if anahtar not in gorulen:
            gorulen.add(anahtar)
            sonuc.append(str(cozulmus))
    return sorted(sonuc, key=lambda item: os.path.normcase(item).casefold())


class HaritaIslemleri:
    def __init__(self, app):
        object.__setattr__(self, "app", app)

    def __getattr__(self, name):
        return getattr(self.app, name)

    def __setattr__(self, name, value):
        if name == "app":
            object.__setattr__(self, name, value)
        else:
            setattr(self.app, name, value)

    def _parsel_harita_boyutunu_al(self):
        widget = getattr(self, "map_widget", None)
        if widget is None:
            return None
        try:
            widget.update_idletasks()
            canvas = getattr(widget, "canvas", widget)
            genislik = int(canvas.winfo_width())
            yukseklik = int(canvas.winfo_height())
        except (AttributeError, tk.TclError, TypeError, ValueError):
            return None
        if genislik < PARSEL_KADRAJ_MIN_BOYUT or yukseklik < PARSEL_KADRAJ_MIN_BOYUT:
            return None
        return genislik, yukseklik

    def _parsel_kadraj_bekleyenini_iptal_et(self):
        root = getattr(self, "root", None)
        after_id = getattr(self, "_parsel_kadraj_after_id", None)
        if after_id is not None and root is not None:
            try:
                root.after_cancel(after_id)
            except (AttributeError, tk.TclError):
                pass
        self._parsel_kadraj_after_id = None

        widget = getattr(self, "map_widget", None)
        binding_id = getattr(self, "_parsel_kadraj_configure_binding_id", None)
        if widget is not None and binding_id:
            try:
                widget.unbind("<Configure>", binding_id)
            except (AttributeError, tk.TclError):
                pass
        self._parsel_kadraj_configure_binding_id = None
        self._parsel_kadraj_bekliyor = False
        self._parsel_kadraj_deneme = 0

    def parsel_kadraj_beklemesini_iptal_et(self):
        """Proje geçişinde bekleyen tek seferlik parsel kadrajını iptal eder."""
        self._parsel_kadraj_bekleyenini_iptal_et()
        self._parsel_odak_aktif = False
        self._parsel_odak_geri_gorunumu = None
        self._parsel_odak_uygulama_bekleniyor = False
        self._parsel_odak_otomatik = False
        self._parseli_odakla_dugmesini_guncelle()

    def _parseli_odakla_configure(self, _event=None):
        if not getattr(self, "_parsel_kadraj_bekliyor", False):
            return
        if getattr(self, "_parsel_kadraj_after_id", None) is not None:
            return
        root = getattr(self, "root", None)
        if root is None:
            return
        try:
            self._parsel_kadraj_after_id = root.after_idle(
                self._parseli_odakla_bekleyenini_calistir
            )
        except (AttributeError, tk.TclError):
            self._parsel_kadraj_after_id = None

    def _parseli_odaklamayi_planla(self):
        self._parsel_kadraj_bekliyor = True
        self._parsel_kadraj_deneme = 0
        widget = getattr(self, "map_widget", None)
        if widget is not None and not getattr(self, "_parsel_kadraj_configure_binding_id", None):
            try:
                self._parsel_kadraj_configure_binding_id = widget.bind(
                    "<Configure>", self._parseli_odakla_configure, add="+"
                )
            except (AttributeError, tk.TclError):
                self._parsel_kadraj_configure_binding_id = None
        self._parseli_odakla_configure()

    def _parseli_odakla_bekleyenini_calistir(self):
        self._parsel_kadraj_after_id = None
        if not getattr(self, "_parsel_kadraj_bekliyor", False):
            return False
        if not getattr(self, "yuklu_kml_points", None):
            self._parsel_kadraj_bekleyenini_iptal_et()
            return False

        boyut = self._parsel_harita_boyutunu_al()
        if boyut is not None:
            uygulandi = self._parseli_odakla_uygula(boyut)
            if uygulandi:
                self._parseli_odaklama_tamamlandi()
                self._parsel_kadraj_bekleyenini_iptal_et()
            return uygulandi

        self._parsel_kadraj_deneme = getattr(self, "_parsel_kadraj_deneme", 0) + 1
        root = getattr(self, "root", None)
        if root is not None and self._parsel_kadraj_deneme < PARSEL_KADRAJ_BEKLEME_DENEMESI:
            try:
                self._parsel_kadraj_after_id = root.after(
                    80, self._parseli_odakla_bekleyenini_calistir
                )
            except (AttributeError, tk.TclError):
                self._parsel_kadraj_after_id = None
        return False

    def _parseli_odakla_uygula(self, boyut):
        try:
            min_zoom = int(getattr(self.map_widget, "min_zoom", 1))
            max_zoom = int(getattr(self.map_widget, "max_zoom", 22))
            merkez, zoom, ideal_zoom = parsel_etkilesimli_kadraj_hesapla(
                getattr(self, "yuklu_kml_points", []),
                boyut[0],
                boyut[1],
                min_zoom=min_zoom,
                max_zoom=max_zoom,
            )
            self.map_widget.set_zoom(zoom)
            self.map_widget.set_position(merkez[0], merkez[1])
            self._parsel_kadraj_sonucu = {
                "merkez": merkez,
                "zoom": zoom,
                "ideal_zoom": ideal_zoom,
                "genislik": boyut[0],
                "yukseklik": boyut[1],
            }
            return True
        except Exception as exc:
            self._parsel_odak_uygulama_bekleniyor = False
            self._parsel_odak_otomatik = False
            try:
                self.hata_kaydet("Parsel haritası kadrajı ayarlanamadı", exc)
            except Exception:
                pass
            return False

    def _parsel_harita_gorunumunu_al(self):
        try:
            konum = self.map_widget.get_position()
            zoom = float(getattr(self.map_widget, "zoom", 15))
            enlem, boylam = float(konum[0]), float(konum[1])
            if not all(math.isfinite(deger) for deger in (enlem, boylam, zoom)):
                return None
            return {"lat": enlem, "lon": boylam, "zoom": zoom}
        except (AttributeError, IndexError, TypeError, ValueError):
            return None

    def _parseli_odakla_dugmesini_guncelle(self):
        dugme = getattr(self, "btn_parseli_odakla", None)
        if dugme is None:
            return
        metin = (
            "Yakın Görünüme Dön"
            if getattr(self, "_parsel_odak_aktif", False)
            else "Parseli Odakla"
        )
        try:
            dugme.configure(text=metin)
        except tk.TclError:
            pass

    def parseli_odakla_dugmesini_guncelle(self):
        """Refresh the focus button after loading a saved map view."""
        self._parseli_odakla_dugmesini_guncelle()

    def _parseli_odaklama_tamamlandi(self):
        otomatik = bool(getattr(self, "_parsel_odak_otomatik", False))
        self._parsel_odak_uygulama_bekleniyor = False
        self._parsel_odak_otomatik = False
        if otomatik:
            sonuc = getattr(self, "_parsel_kadraj_sonucu", {}) or {}
            merkez = sonuc.get("merkez")
            zoom = sonuc.get("zoom")
            if merkez and zoom is not None:
                self._parsel_odak_geri_gorunumu = {
                    "lat": merkez[0],
                    "lon": merkez[1],
                    "zoom": kesirli_zoom_degerini_duzelt(
                        float(zoom) + KESIRLI_ZOOM_ADIMI,
                        min_zoom=getattr(self.map_widget, "min_zoom", 1),
                        max_zoom=getattr(self.map_widget, "max_zoom", 22),
                        adim=KESIRLI_ZOOM_ADIMI,
                    ),
                }
        self._parsel_odak_aktif = True
        self._parseli_odakla_dugmesini_guncelle()
        if otomatik and hasattr(self, "durum_mesaji_yaz"):
            self.durum_mesaji_yaz(
                "Parsel tam görünümde. Yakın çalışma görünümüne dönmek için düğmeye basın."
            )

    def _yakin_gorunume_don(self):
        gorunum = getattr(self, "_parsel_odak_geri_gorunumu", None)
        if not gorunum:
            self._parsel_odak_aktif = False
            self._parseli_odakla_dugmesini_guncelle()
            return False
        try:
            self._parsel_kadraj_bekleyenini_iptal_et()
            self.map_widget.set_zoom(gorunum["zoom"])
            self.map_widget.set_position(gorunum["lat"], gorunum["lon"])
        except Exception as exc:
            try:
                self.hata_kaydet("Önceki harita görünümüne dönülemedi", exc)
            except Exception:
                pass
            return False

        self._parsel_odak_geri_gorunumu = None
        self._parsel_odak_aktif = False
        self._parseli_odakla_dugmesini_guncelle()
        if hasattr(self, "durum_mesaji_yaz"):
            self.durum_mesaji_yaz("Yakın çalışma görünümüne dönüldü.")
        return True

    def parseli_odakla(self, otomatik=False):
        """Parseli sığdırır veya saklanan yakın/pan görünümünü geri yükler."""
        if not otomatik and getattr(self, "_parsel_odak_aktif", False):
            return self._yakin_gorunume_don()
        if not getattr(self, "yuklu_kml_points", None):
            return False

        if getattr(self, "_parsel_kadraj_bekliyor", False):
            if not otomatik:
                return False
            self._parsel_kadraj_bekleyenini_iptal_et()

        onceki_gorunum = self._parsel_harita_gorunumunu_al()
        if otomatik:
            # Yeni KML için eski, ilgisiz harita konumu yerine parsel merkezinde
            # bir üst yakınlık düzeyi saklanır. Böylece ikinci tık çalışma
            # alanına yakın, anlamlı bir görünüme döner.
            try:
                boyut = self._parsel_harita_boyutunu_al()
                if boyut is not None:
                    merkez, zoom, _ = parsel_etkilesimli_kadraj_hesapla(
                        self.yuklu_kml_points,
                        boyut[0],
                        boyut[1],
                        min_zoom=int(getattr(self.map_widget, "min_zoom", 1)),
                        max_zoom=int(getattr(self.map_widget, "max_zoom", 22)),
                    )
                    onceki_gorunum = {
                        "lat": merkez[0],
                        "lon": merkez[1],
                        "zoom": kesirli_zoom_degerini_duzelt(
                            float(zoom) + KESIRLI_ZOOM_ADIMI,
                            min_zoom=getattr(self.map_widget, "min_zoom", 1),
                            max_zoom=getattr(self.map_widget, "max_zoom", 22),
                            adim=KESIRLI_ZOOM_ADIMI,
                        ),
                    }
            except (TypeError, ValueError):
                pass
        self._parsel_odak_geri_gorunumu = onceki_gorunum
        self._parsel_odak_uygulama_bekleniyor = True
        self._parsel_odak_otomatik = bool(otomatik)
        self._parsel_odak_aktif = False
        self._parseli_odakla_dugmesini_guncelle()

        boyut = self._parsel_harita_boyutunu_al()
        if boyut is None:
            self._parseli_odaklamayi_planla()
            return False

        self._parsel_kadraj_bekleyenini_iptal_et()
        uygulandi = self._parseli_odakla_uygula(boyut)
        if uygulandi:
            self._parseli_odaklama_tamamlandi()
            if not otomatik and hasattr(self, "durum_mesaji_yaz"):
                self.durum_mesaji_yaz(
                    "Parsel tam görünümde. Önceki yakın/pan görünümüne dönmek için düğmeye basın."
                )
        return uygulandi

    def harita_araci_degisti(self):
        if self.ss_ilk_nokta is not None:
            if self.temp_ss_marker: self.temp_ss_marker.delete(); self.temp_ss_marker = None
            self.ss_ilk_nokta = None

        for isim, data in self.harita_isaretleri.items():
            tip = data["tip"]
            if tip == "M":
                 pass
                      
    def sil_ve_yeniden_ciz(self, ac_yn_goster=True, ss_goster=True, m_goster=True):
        for isim, data in self.harita_isaretleri.items():
            if "marker" in data and data["marker"]: data["marker"].delete(); data["marker"] = None
            if "path" in data and data["path"]: data["path"].delete(); data["path"] = None

        for isim, data in self.harita_isaretleri.items():
            tip = data["tip"]
            if (tip in ["AÇ", "YN"] and ac_yn_goster) or (tip == "SS" and ss_goster) or (tip == "M" and m_goster):
                if tip == "M": 
                     if m_goster: data["marker"] = self.map_widget.set_marker(data["lat"], data["lon"], text="M", icon=self.ikon_m, text_color="black", font=("Arial", 28, "bold"))
                elif tip == "AÇ": data["marker"] = self.map_widget.set_marker(data["lat"], data["lon"], text=isim, icon=self.ikon_ac, text_color="black", font=("Arial", 28, "bold"))
                elif tip == "YN": data["marker"] = self.map_widget.set_marker(data["lat"], data["lon"], text=isim, icon=self.ikon_yn, text_color="black", font=("Arial", 28, "bold"))
                elif tip == "SS":
                    p = self.map_widget.set_path(
                        [data["n1"], data["n2"]],
                        color=SS_SERIM_RENGI,
                        width=SS_SERIM_KALINLIGI,
                    )
                    if hasattr(p, 'canvas_line'): self.map_widget.canvas.itemconfig(p.canvas_line, dash=SS_SERIM_KESIK_DESENI)
                    elif hasattr(p, 'canvas_line_list'):
                        for line in p.canvas_line_list: self.map_widget.canvas.itemconfig(line, dash=SS_SERIM_KESIK_DESENI)
                    data["marker"] = self.map_widget.set_marker(
                        data["n1"][0],
                        data["n1"][1],
                        text=isim,
                        marker_color_outside=SS_SERIM_RENGI,
                        marker_color_circle=SS_SERIM_RENGI,
                        text_color="white",
                        font=("Arial", 30, "bold"),
                    )
                    data["path"] = p

    def harita_sol_tik(self, coords):
        arac = self.aktif_harita_araci.get()
        if arac == "Yok": return
        
        if arac == "M":
             isim = "Merkez"
             # Varsa eski merkezi sil
             if "Merkez" in self.harita_isaretleri:
                 old_m = self.harita_isaretleri.pop("Merkez")
                 if "marker" in old_m and old_m["marker"]: old_m["marker"].delete()
             self.harita_isaretleri[isim] = {"tip": "M", "lat": coords[0], "lon": coords[1]}
             self.sil_ve_yeniden_ciz(True, True, True)
             return

        sayac = self.harita_nokta_sayaclari[arac]
        isim = f"{arac}{sayac}"
        if arac in ["AÇ", "YN"]:
            self.harita_isaretleri[isim] = {"tip": arac, "lat": coords[0], "lon": coords[1]}
            self.harita_nokta_sayaclari[arac] += 1
            self.sil_ve_yeniden_ciz(True, True, True)
        elif arac == "SS":
            if self.ss_ilk_nokta is None:
                self.ss_ilk_nokta = coords
                self.temp_ss_marker = self.map_widget.set_marker(coords[0], coords[1], text=f"{isim} (Başlangıç)", marker_color_outside="red", marker_color_circle="white", font=("Arial", 28, "bold"))
            else:
                self.harita_isaretleri[isim] = {"tip": "SS", "n1": self.ss_ilk_nokta, "n2": coords}
                if self.temp_ss_marker: self.temp_ss_marker.delete(); self.temp_ss_marker = None
                self.harita_nokta_sayaclari[arac] += 1
                self.ss_ilk_nokta = None
                self.sil_ve_yeniden_ciz(True, True, True)

    def harita_sag_tik(self, coords):
        if not self.harita_isaretleri: return
        min_mesafe = float('inf'); en_yakin_isim = None
        for isim, data in self.harita_isaretleri.items():
            pos = data["n1"] if data["tip"] == "SS" else (data["lat"], data["lon"])
            mesafe = ((pos[0] - coords[0])**2 + (pos[1] - coords[1])**2)**0.5
            if mesafe < min_mesafe: min_mesafe = mesafe; en_yakin_isim = isim
        if min_mesafe < 0.002 and en_yakin_isim:
            if not messagebox.askyesno(
                "Harita Noktasını Sil",
                f"{en_yakin_isim} haritadan silinsin mi? İlişkili tablo verileri bir sonraki "
                "senkronizasyonda ayrıca korunacak veya silinmek üzere size sorulacaktır.",
            ):
                return
            data = self.harita_isaretleri.pop(en_yakin_isim)
            if "marker" in data and data["marker"]: data["marker"].delete()
            if "path" in data and data["path"]: data["path"].delete()
            if en_yakin_isim == "Merkez": return # Merkez noktasının sayacı yok
            prefix = ''.join([c for c in en_yakin_isim if not c.isdigit()]); num = int(''.join([c for c in en_yakin_isim if c.isdigit()]))
            if self.harita_nokta_sayaclari.get(prefix, 0) - 1 == num: self.harita_nokta_sayaclari[prefix] -= 1

    def harita_verilerini_senkronize_et(self):
        isaretler = getattr(self, "harita_isaretleri", {}) or {}
        hedef_ac_yn = [isim for isim, data in isaretler.items() if data.get("tip") in ("AÇ", "YN")]
        mevcut_kayitlar = {kayit["isim"]: kayit for kayit in self.ac_yn_sekme_kayitlari()}
        ac_yn_plan = isim_bazli_birlestirme_plani(mevcut_kayitlar, hedef_ac_yn)

        hedef_ss = [isim for isim, data in isaretler.items() if data.get("tip") == "SS"]
        mevcut_ss = {}
        if hasattr(self, "tree_sis"):
            for item in self.tree_sis.get_children():
                degerler = self.tree_sis.item(item).get("values", [])
                if degerler:
                    mevcut_ss[str(degerler[0])] = item
        ss_plan = isim_bazli_birlestirme_plani(mevcut_ss, hedef_ss)

        orphan_isimler = ac_yn_plan["orphan"]
        orphan_ss = ss_plan["orphan"]
        orphanlari_sil = False
        if orphan_isimler or orphan_ss:
            parcalar = []
            if orphan_isimler:
                parcalar.append("AÇ/YN: " + ", ".join(orphan_isimler))
            if orphan_ss:
                parcalar.append("Jeofizik: " + ", ".join(orphan_ss))
            cevap = messagebox.askyesnocancel(
                "Haritada Olmayan Kayıtlar",
                "Haritada artık bulunmayan ancak tablolarda verisi olan kayıtlar var:\n\n"
                + "\n".join(parcalar)
                + "\n\nEvet: Bu orphan kayıtları ve aynı numaralı laboratuvar satırlarını sil."
                  "\nHayır: Verileri koru ve yalnız mevcut harita noktalarını güncelle."
                  "\nİptal: Senkronizasyondan vazgeç.",
            )
            if cevap is None:
                return
            orphanlari_sil = bool(cevap)

        if orphanlari_sil:
            for isim in orphan_isimler:
                kayit = mevcut_kayitlar.get(isim)
                if not kayit:
                    continue
                sekme = kayit.get("sekme")
                if sekme is not None:
                    self.ac_yn_sekme_bilgileri.pop(str(sekme), None)
                    sekme.destroy()
            if hasattr(self, "tree_sis"):
                for isim in orphan_ss:
                    item = mevcut_ss.get(isim)
                    if item:
                        self.tree_sis.delete(item)
            self._harita_orphan_lab_satirlarini_sil(orphan_isimler)

        ss_count = 0
        ac_yn_count = 0
        merkez_guncellendi = False
        mevcut_ac = set(self.lab_ac_numaralari_al()) if hasattr(self, "tree_lab_ac") else set()
        mevcut_yn = set(self.lab_yn_numaralari_al()) if hasattr(self, "tree_lab_yn") else set()

        for isim, data in isaretler.items():
            tip = data.get("tip")
            if tip == "SS":
                degerler = (isim, f"{data['n1'][0]:.6f}", f"{data['n1'][1]:.6f}")
                if isim in mevcut_ss:
                    self.tree_sis.item(mevcut_ss[isim], values=degerler)
                else:
                    self.jeofizik_koordinat_ekle(*degerler)
                ss_count += 1
            elif tip == "M":
                if "ENLEM" in self.veri_alanlari:
                    self.veri_alanlari["ENLEM"].delete(0, tk.END)
                    self.veri_alanlari["ENLEM"].insert(0, f"{data['lat']:.6f}")
                if "BOYLAM" in self.veri_alanlari:
                    self.veri_alanlari["BOYLAM"].delete(0, tk.END)
                    self.veri_alanlari["BOYLAM"].insert(0, f"{data['lon']:.6f}")
                merkez_guncellendi = True
            elif tip in ("AÇ", "YN"):
                kayit = mevcut_kayitlar.get(isim)
                enlem = f"{data['lat']:.6f}"
                boylam = f"{data['lon']:.6f}"
                if kayit:
                    for entry, deger in ((kayit["enlem_entry"], enlem), (kayit["boylam_entry"], boylam)):
                        entry.delete(0, tk.END)
                        entry.insert(0, deger)
                else:
                    self.cukur_sekmesi_ekle(isim, enlem, boylam)
                ac_yn_count += 1
                if tip == "AÇ" and hasattr(self, "tree_lab_ac") and isim not in mevcut_ac:
                    self.lab_ac_bos_satir_ekle(isim)
                    mevcut_ac.add(isim)
                elif tip == "YN" and hasattr(self, "tree_lab_yn") and isim not in mevcut_yn:
                    self.lab_yn_bos_satir_ekle(isim)
                    mevcut_yn.add(isim)

        if hasattr(self, "tree_lab_ac"):
            self.stripe_tree(self.tree_lab_ac)
        if hasattr(self, "tree_lab_yn"):
            self.stripe_tree(self.tree_lab_yn)
        if hasattr(self, "senkronize_ac_tablo"):
            self.senkronize_ac_tablo(temizle_eslesmeyen=False)

        mesaj = f"Haritadaki {ac_yn_count} nokta ve {ss_count} serim tablolarda birleştirilip güncellendi."
        if (orphan_isimler or orphan_ss) and not orphanlari_sil:
            mesaj += "\n\nHaritada olmayan kayıtlar veri kaybını önlemek için korundu."
        if merkez_guncellendi:
            mesaj += "\n\nMerkez koordinatları 'Arazi Bilgileri' sekmesine aktarıldı."
        if hasattr(self, "durum_mesaji_yaz"):
            self.durum_mesaji_yaz(mesaj.replace("\n", " · "))

    def _harita_orphan_lab_satirlarini_sil(self, orphan_isimler):
        hedefler = {str(isim) for isim in orphan_isimler}
        for tree_attr in ("tree_lab_ac", "tree_lab_yn"):
            tree = getattr(self, tree_attr, None)
            if not tree:
                continue
            for item in tree.get_children():
                degerler = tree.item(item).get("values", [])
                if degerler and str(degerler[0]) in hedefler:
                    tree.delete(item)
        if hasattr(self, "lab_sayaclari_guncelle"):
            self.lab_sayaclari_guncelle()

    def kml_haritaya_yukle(self):
        dosya_yolu = filedialog.askopenfilename(
            initialdir=self.sablon_alt_klasoru("kml"),
            filetypes=[("KML Dosyaları", "*.kml")]
        )
        if not dosya_yolu:
            return False
        return self.kml_dosyasini_yukle(dosya_yolu)

    def kml_dosyasini_yukle(self, dosya_yolu, basari_bildir=True):
        """Verilen KML'yi doğrulayıp çalışma parseli olarak haritaya bağlar."""
        try:
            poligonlar = kml_poligonlarini_oku(dosya_yolu)
            ana_poligon = max(
                poligonlar,
                key=lambda item: self._kml_poligon_kapsama_alani(item["noktalar"]),
            )
            poly_coords = ana_poligon["noktalar"]
            yeni_polygon = self.map_widget.set_polygon(
                poly_coords,
                fill_color=None,
                outline_color=CALISAN_PARSEL_SINIR_RENGI,
                border_width=CALISAN_PARSEL_SINIR_KALINLIGI,
            )
            eski_polygon = self.kml_polygon_obj
            if eski_polygon:
                try:
                    eski_polygon.delete()
                except Exception as exc:
                    self.hata_kaydet("Eski KML poligonu haritadan silinemedi", exc)
            self.kml_polygon_obj = yeni_polygon
            self.yuklu_kml_points = poly_coords
            self.yuklu_kml_yolu = os.path.abspath(dosya_yolu)
            self.parsel_haritasi_kaynak_url = ""
            self.jeoloji_pafta_sonucu = {}
            self.genel_jeoloji_verisi = {}
            self.img_genel_jeoloji = None
            self.img_parsel_haritasi = None
            self.parsel_haritasi_geometri_hash = ""
            self.parsel_haritasi_ada = ""
            self.parsel_haritasi_parsel = ""
            if hasattr(self, "proje_durum_seridi_guncelle"):
                self.proje_durum_seridi_guncelle()
            # KML yüklendiğinde sabit zoom yerine gerçek widget boyutuna göre
            # güvenli kadraj uygula. Widget henüz yerleşmediyse metot bunu
            # Configure/after_idle sonrasında tek seferlik olarak tamamlar.
            self.parseli_odakla(otomatik=True)
            if hasattr(self, "jeoloji_harita_katmanini_yenile"):
                self.jeoloji_harita_katmanini_yenile(zorla=True)
            if hasattr(self, "jeoloji_pafta_durumunu_guncelle"):
                self.jeoloji_pafta_durumunu_guncelle()
            if hasattr(self, "genel_jeoloji_durumunu_guncelle"):
                self.genel_jeoloji_durumunu_guncelle()
            if hasattr(self, "durum_mesaji_yaz"):
                self.durum_mesaji_yaz("KML çalışma parseline bağlandı", os.path.basename(dosya_yolu))
            if basari_bildir:
                messagebox.showinfo(
                    "Başarılı",
                    "KML poligonu (çalışma alanı sınırı) doğrulanıp haritaya eklendi.",
                )
            return True
        except Exception as e:
            self.hata_kaydet("KML yüklenirken hata oluştu", e)
            messagebox.showerror("Hata", f"KML yüklenirken hata oluştu: {e}")
            return False

    @staticmethod
    def _kml_poligon_kapsama_alani(noktalar):
        if not noktalar:
            return 0.0
        enlemler = [nokta[0] for nokta in noktalar]
        boylamlar = [nokta[1] for nokta in noktalar]
        return (max(enlemler) - min(enlemler)) * (max(boylamlar) - min(boylamlar))

    def _tkgm_kunye_olustur(self):
        return {
            "il": self.proje_deger("IL", ""),
            "ilce": self.proje_deger("ILCE", ""),
            "koy": self.proje_deger("KOY", ""),
            "ada": self.proje_deger("ADA", "0"),
            "parsel": self.proje_deger("PARSEL", ""),
        }

    def _tkgm_kml_output_dir(self):
        proje_yolu = getattr(self, "guncel_dosya_yolu", None)
        if proje_yolu:
            return os.path.join(os.path.dirname(os.path.abspath(proje_yolu)), "Haritalar")
        return os.path.join(self.kullanici_veri_klasoru_bul(), "TKGM_KML")

    def tkgm_kml_al(self):
        """Formdaki künye ile TKGM geometrisini alıp mevcut KML akışına bağlar."""
        if getattr(self, "_tkgm_kml_calisisiyor", False):
            messagebox.showinfo("TKGM KML", "TKGM parsel sorgusu zaten çalışıyor.")
            return False

        kunye = self._tkgm_kunye_olustur()
        eksikler = []
        for key, label in (("il", "İl"), ("ilce", "İlçe"), ("koy", "Köy"), ("parsel", "Parsel")):
            if not str(kunye.get(key) or "").strip():
                eksikler.append(label)
        if eksikler:
            messagebox.showwarning(
                "TKGM KML",
                "TKGM'den KML almak için şu alanlar doldurulmalıdır:\n- "
                + "\n- ".join(eksikler),
            )
            return False

        try:
            output_dir = self._tkgm_kml_output_dir()
            output_path = os.path.join(output_dir, tkgm_kml_dosya_adi(kunye))
        except Exception as exc:
            messagebox.showerror("TKGM KML", str(exc))
            return False

        if os.path.exists(output_path) and not messagebox.askyesno(
            "TKGM KML",
            f"Bu parsel için daha önce oluşturulmuş KML var:\n{output_path}\n\nYenilensin mi?",
        ):
            return False

        progress = self.animasyonlu_pencere()
        progress.title("TKGM KML")
        progress.transient(self.root)
        progress.resizable(False, False)
        progress.protocol("WM_DELETE_WINDOW", lambda: None)
        ttk.Label(
            progress,
            text="TKGM Parsel Sorgu'dan parsel geometrisi alınıyor...",
        ).pack(anchor="w", padx=16, pady=(16, 6))
        ttk.Label(
            progress,
            text=(
                f"{kunye['il']} / {kunye['ilce']} / {kunye['koy']} / "
                f"{kunye.get('ada') or '0'}/{kunye['parsel']}"
            ),
            style="Muted.TLabel",
        ).pack(anchor="w", padx=16, pady=(0, 8))
        bar = ttk.Progressbar(progress, mode="indeterminate", length=390)
        bar.pack(fill="x", padx=16, pady=(0, 16))
        bar.start(12)
        progress.update_idletasks()

        button = getattr(self, "btn_tkgm_kml", None)
        if button is not None:
            button.configure(state="disabled")
        self._tkgm_kml_calisisiyor = True
        if hasattr(self, "durum_mesaji_yaz"):
            self.durum_mesaji_yaz("TKGM parsel geometrisi alınıyor")

        def islemi_bitir():
            self._tkgm_kml_calisisiyor = False
            if button is not None:
                button.configure(state="normal")
            if progress.winfo_exists():
                bar.stop()
                progress.destroy()

        def basarili(result):
            islemi_bitir()
            path = result.get("path", "")
            if not path or not self.kml_dosyasini_yukle(path, basari_bildir=False):
                return
            self.parsel_haritasi_kaynak_url = str(result.get("source_url") or "").strip()
            center = result.get("center")
            if center:
                for code, value in (("ENLEM", center[0]), ("BOYLAM", center[1])):
                    entry = self.veri_alanlari.get(code)
                    if entry is not None and not entry.get().strip():
                        entry.delete(0, tk.END)
                        entry.insert(0, f"{value:.8f}")
            parca_notu = ""
            if result.get("polygon_count", 1) > 1:
                parca_notu = "\n\nParsel birden fazla geometrik parçadan oluşuyor; ana parça çalışma sınırı olarak gösterildi."
            messagebox.showinfo(
                "TKGM KML",
                f"TKGM KML oluşturuldu ve haritaya bağlandı:\n{path}{parca_notu}",
            )

        def basarisiz(exc):
            islemi_bitir()
            self.hata_kaydet("TKGM KML alınamadı", exc)
            messagebox.showerror(
                "TKGM KML",
                "TKGM'den parsel KML'si alınamadı.\n\n"
                f"{exc}\n\nMevcut yüklü KML değiştirilmedi.",
            )
            if hasattr(self, "durum_mesaji_yaz"):
                self.durum_mesaji_yaz("TKGM KML alınamadı")

        def worker():
            try:
                result = tkgm_parsel_kml_olustur(kunye, output_dir)
            except Exception as exc:
                try:
                    self.root.after(0, lambda error=exc: basarisiz(error))
                except tk.TclError:
                    pass
                return
            try:
                self.root.after(0, lambda value=result: basarili(value))
            except tk.TclError:
                pass

        threading.Thread(target=worker, name="tkgm-kml", daemon=True).start()
        return True

    # --- HATCH / TARAMA ÇİZİM MOTORU ---
