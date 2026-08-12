from on_deger import bos_is_akisi_verisi, bos_on_deger_verisi, is_durumu_degistir, on_deger_revizyonu_ekle
from proje_durumu_islemleri import (
    ProjeDurumuIslemleri,
    parsel_haritasi_durumu_hazirla,
    proje_durum_ozeti_hazirla,
)


class _SahteRoot:
    def __init__(self):
        self.planlanan = None
        self.iptal = []

    def after(self, gecikme, callback):
        self.planlanan = (gecikme, callback)
        return "after-1"

    def after_cancel(self, kimlik):
        self.iptal.append(kimlik)


class _SahteWidget:
    def __init__(self, sinif):
        self.sinif = sinif

    def winfo_class(self):
        return self.sinif


class _SahteOlay:
    def __init__(self, olay_tipi, widget_sinifi=None):
        self.type = olay_tipi
        self.widget = _SahteWidget(widget_sinifi) if widget_sinifi else None


class _SahteApp:
    def __init__(self):
        self.root = _SahteRoot()
        self._proje_kirli = False
        self._proje_durumu_after_id = None


def test_dogrudan_yazim_seridi_on_deger_yok_gosterir():
    akis = is_durumu_degistir(bos_is_akisi_verisi(), "yazim_asamasinda")
    ozet = proje_durum_ozeti_hazirla({
        "PROJE_ADI": "Ercan Şahin", "ILCE": "Ayvacık", "KOY": "Kozlu",
        "MEVKII": "Köyiçi", "ADA": "151", "PARSEL": "10",
        "_IS_AKISI_": akis, "_ON_DEGER_": bos_on_deger_verisi(),
    }, kaydedilmedi=True)

    assert ozet["kimlik"] == "Ayvacık / Kozlu / 151-10"
    assert ozet["asama"] == "Yazım Aşamasında"
    assert ozet["on_deger"] == "Ön Değer: Verilmedi"
    assert ozet["kaydedilmedi"] is True
    assert ozet["haritali_ozet"] is True
    assert ozet["proje_adi"] == "Ercan Şahin"
    assert ozet["konum"] == "Ayvacık / Kozlu / Köyiçi"
    assert ozet["ada_parsel"] == "ADA 151 — PARSEL 10"
    assert ozet["pencere_basligi"].endswith("KAYDEDİLMEDİ")


def test_bitmis_izleme_ozeti_baslikta_yalniz_mevcut_asamayi_gosterir():
    akis = bos_is_akisi_verisi()
    akis = is_durumu_degistir(akis, "on_deger_verildi")
    akis = is_durumu_degistir(akis, "yazim_asamasinda")
    akis = is_durumu_degistir(akis, "bitti")
    on_deger, _ = on_deger_revizyonu_ekle(bos_on_deger_verisi(), "20", "1000", "ZD", "", "abc")
    ozet = proje_durum_ozeti_hazirla({
        "ILCE": "Bayramiç", "KOY": "Akçakıl", "ADA": "0", "PARSEL": "673",
        "_IS_AKISI_": akis, "_ON_DEGER_": on_deger,
    }, salt_okunur=True)

    assert ozet["on_deger"] == "Ön Değer: Verildi"
    assert ozet["mod"] == "İzleme"
    assert ozet["pencere_basligi"] == "K-1 — BİTTİ — BAYRAMİÇ / AKÇAKIL / 0-673"
    assert "İZLEME" not in ozet["pencere_basligi"]


def test_parsel_haritasi_yalniz_uygun_asamada_gosterilir(tmp_path):
    harita = tmp_path / "Parsel_Haritasi.png"
    harita.write_bytes(b"test")
    ortak = {
        "harita_yolu": str(harita),
        "kml_noktalari": [[39.0, 26.0], [39.1, 26.0], [39.0, 26.1]],
        "kayitli_hash": "abc",
        "guncel_hash": "abc",
        "kayitli_ada": "151",
        "kayitli_parsel": "10",
        "guncel_ada": "151",
        "guncel_parsel": "10",
    }

    assert parsel_haritasi_durumu_hazirla("yeni", **ortak)["kod"] == "asama_disinda"
    assert parsel_haritasi_durumu_hazirla("yazim_asamasinda", **ortak)["kod"] == "hazir"


def test_parsel_haritasi_eski_geometriyi_gostermez(tmp_path):
    harita = tmp_path / "Parsel_Haritasi.png"
    harita.write_bytes(b"test")
    sonuc = parsel_haritasi_durumu_hazirla(
        "bitti",
        harita_yolu=str(harita),
        kml_noktalari=[[39.0, 26.0], [39.1, 26.0], [39.0, 26.1]],
        kayitli_hash="eski",
        guncel_hash="yeni",
        kayitli_ada="151",
        kayitli_parsel="10",
        guncel_ada="151",
        guncel_parsel="10",
    )

    assert sonuc["kod"] == "geometri_degisti"
    assert sonuc["goster"] is False


def test_kullanici_etkilesimi_kirli_bayragini_isaretleyip_hizli_yenileme_planlar():
    app = _SahteApp()

    ProjeDurumuIslemleri(app).proje_durumu_yenilemeyi_planla(_SahteOlay("2"))

    assert app._proje_kirli is True
    assert app.root.planlanan[0] == 160


def test_notebook_sekme_degistirmek_projeyi_kirli_yapmaz():
    app = _SahteApp()

    ProjeDurumuIslemleri(app).proje_durumu_yenilemeyi_planla(_SahteOlay("35"))

    assert app._proje_kirli is False


def test_notebook_uzerindeki_fare_birakma_projeyi_kirli_yapmaz():
    app = _SahteApp()

    ProjeDurumuIslemleri(app).proje_durumu_yenilemeyi_planla(_SahteOlay("5", "TNotebook"))

    assert app._proje_kirli is False
