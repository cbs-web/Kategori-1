import copy
import json
import os
from pathlib import Path

import pytest

from kayit import KayitYoneticisi, SCHEMA_VERSION, TAAHHUT_BILGI_ALANLARI, YEDEK_LIMITI


class SahteUygulama:
    def __init__(self, veri_klasoru):
        self._veri_klasoru = str(veri_klasoru)
        self.veri_alanlari = {"PROJE_ADI": object(), "IL": object()}
        self.bina_alanlari = {"Bina Kullanım Amacı": object()}
        self.tg_girdiler = {"B": object(), "L": object(), "ks_carpani": object(), "gsat": object(), "yass": object()}
        self.ek_kategorileri = ["EVRAKLAR", "LOG"]
        self.hatalar = []
        self.taahhut_varsayilanlari = {kod: "" for kod in TAAHHUT_BILGI_ALANLARI}
        self.taahhut_bilgileri = dict(self.taahhut_varsayilanlari)
        self.varsayilan_proje_verisi = {
            "schema_version": SCHEMA_VERSION,
            "PROJE_ADI": "",
            "IL": "",
            "_BINA_": {"Bina Kullanım Amacı": ""},
            "_FORMASYON_": "Seçiniz...",
            "_FORMASYON_METNI_": "",
            "_MUHENDISLIK_JEOLOJISI_METNI_": "",
            "_AC_SEKMELERI_": [],
            "_JEOFIZIK_": {
                "excel_yolu": "",
                "tree_sis": [],
                "jeofon_dizilim": {"jeofon_sayisi": "12"},
            },
            "_TASIMA_": {
                "secim": "zemin",
                "girdiler": {
                    "B": "1.0",
                    "L": "2.0",
                    "ks_carpani": "1.0",
                    "gsat": "20.0",
                    "yass": "999.0",
                },
                "qt_nihai": "",
                "ks_nihai": "",
                "son_qk": "-",
                "son_qt": "-",
                "rapor_metni": "",
                "varsayim_onayi": False,
                "dayanim_23_uygulandi": False,
                "dayanim_23_kaynak_c": "",
                "dayanim_23_kaynak_phi": "",
                "rapor_imzasi": None,
            },
            "__HARITA__": {
                "zoom": 15,
                "lat": 39.524,
                "lon": 26.120,
                "kml_yolu": "",
                "kml_points": [],
                "sayaclar": {"AÇ": 1, "YN": 1, "SS": 1},
                "isaretler": {},
            },
            "_LAB_AC_": [],
            "_LAB_YN_": [],
            "_LAB_KAYNAK_": "",
            "_EKLER_": {"EVRAKLAR": [], "LOG": []},
            "_TAAHHUT_BILGILERI_": dict(self.taahhut_varsayilanlari),
            "_RAPOR_SABLONU_": "varsayilan.docx",
            "_TAAHHUT_WORD_SABLONU_": "taahhut.docx",
            "_JEOLOJI_SABLONU_": "",
        }

    def kullanici_veri_klasoru_bul(self):
        return self._veri_klasoru

    def uygulama_klasoru_bul(self):
        return self._veri_klasoru

    def proje_deger(self, kod, varsayilan=""):
        return varsayilan

    def hata_kaydet(self, baslik, hata=None):
        self.hatalar.append((baslik, hata))


class SahteDegisken:
    def __init__(self, deger):
        self.deger = deger

    def set(self, deger):
        self.deger = deger


class SahteEtiket:
    def __init__(self, metin):
        self.metin = metin

    def config(self, **ayarlar):
        self.metin = ayarlar.get("text", self.metin)


class SahteHarita:
    def __init__(self):
        self.silinenler = []

    def delete_all_marker(self):
        self.silinenler.append("marker")

    def delete_all_path(self):
        self.silinenler.append("path")

    def delete_all_polygon(self):
        self.silinenler.append("polygon")


class SahteRoot:
    def __init__(self):
        self.iptal_edilen = None

    def after_cancel(self, after_id):
        self.iptal_edilen = after_id


class SahteBaslikRoot:
    def __init__(self, baslik="K-1 - Yeni Proje"):
        self.baslik = baslik

    def title(self, *deger):
        if deger:
            self.baslik = deger[0]
        return self.baslik


class KayitAkisiYoneticisi(KayitYoneticisi):
    """Dosya aç/kaydet sözleşmesini arayüz widget'larından bağımsız sınar."""

    def degisiklik_gecisine_izin_ver(self, eylem):
        return True

    def verileri_topla(self):
        return copy.deepcopy(self.app.durum)

    def verileri_yerlestir(self, veriler, dogrulandi=False):
        self.app.durum = copy.deepcopy(veriler)

    def proje_verisini_normalize_et(self, veriler):
        return copy.deepcopy(veriler)

    def son_proje_ekle(self, yol):
        self.son_projeler = [os.path.abspath(yol)] + [
            eski for eski in self.son_projeler if os.path.abspath(eski) != os.path.abspath(yol)
        ]
        if getattr(self.app, "son_proje_hatasi", False):
            raise RuntimeError("son proje listesi güncellenemedi")


@pytest.fixture
def yonetici(tmp_path):
    return KayitYoneticisi(SahteUygulama(tmp_path))


def test_legacy_kayit_mevcut_durumla_degil_bos_varsayilanla_birlestirilir(yonetici):
    legacy = {
        "PROJE_ADI": "Yeni proje",
        "_BINA_": {"Bina Kullanım Amacı": "Konut"},
        "_TASIMA_": {"secim": "zemin", "girdiler": {"B": "3.0"}},
    }

    sonuc = yonetici.proje_verisini_normalize_et(legacy)

    assert sonuc["schema_version"] == SCHEMA_VERSION
    assert sonuc["PROJE_ADI"] == "Yeni proje"
    assert sonuc["IL"] == ""
    assert sonuc["_TASIMA_"]["girdiler"] == {
        "B": "3.0",
        "L": "2.0",
        "ks_carpani": "1.0",
        "gsat": "20.0",
        "yass": "999.0",
    }
    assert sonuc["_TASIMA_"]["varsayim_onayi"] is False
    assert sonuc["_TASIMA_"]["yass_var"] is False
    assert sonuc["_TASIMA_"]["dayanim_23_uygulandi"] is False
    assert sonuc["_RAPOR_SABLONU_"] == "varsayilan.docx"
    assert sonuc["_LAB_AC_"] == []
    assert sonuc["_TAAHHUT_BILGILERI_"] == {kod: "" for kod in TAAHHUT_BILGI_ALANLARI}


def test_kutuphane_jeoloji_word_baglantisi_proje_verisinde_korunur(yonetici):
    result = yonetici.proje_verisini_normalize_et({
        "_JEOLOJI_KUTUPHANE_BOLUMU_": {
            "aktif": True,
            "kayit_id": 27,
            "bolum_docx_path": r"C:\veri\jeoloji_bolumu.docx",
            "bolum_hash": "abc123",
            "uygulanan_genel": "Genel metin",
            "uygulanan_inceleme": "İnceleme alanı metni",
        }
    })

    assert result["_JEOLOJI_KUTUPHANE_BOLUMU_"] == {
        "aktif": True,
        "kayit_id": 27,
        "bolum_docx_path": r"C:\veri\jeoloji_bolumu.docx",
        "bolum_hash": "abc123",
        "uygulanan_genel": "Genel metin",
        "uygulanan_inceleme": "İnceleme alanı metni",
    }


@pytest.mark.parametrize(
    "veri,mesaj",
    [
        ([], "kök değeri"),
        ({"schema_version": SCHEMA_VERSION + 1}, "daha yeni"),
        ({"_AC_SEKMELERI_": "geçersiz"}, "kayıt listesi"),
        ({"_TAAHHUT_BILGILERI_": []}, "_TAAHHUT_BILGILERI_ bir JSON nesnesi"),
        ({"__HARITA__": {"lat": 100, "lon": 26, "zoom": 15}}, "en fazla"),
    ],
)
def test_gecersiz_proje_semasi_yuklemeden_once_reddedilir(yonetici, veri, mesaj):
    with pytest.raises(ValueError, match=mesaj):
        yonetici.proje_verisini_normalize_et(veri)


def test_legacy_harita_isaretleri_widget_referansi_olmadan_migrate_edilir(yonetici):
    sonuc = yonetici.proje_verisini_normalize_et({
        "__HARITA__": {
            "lat": 39.0,
            "lon": 26.0,
            "zoom": 12,
            "sayaclar": {"AÇ": 2, "YN": 1, "SS": 1},
            "isaretler": {
                "AÇ1": {"tip": "AÇ", "lat": 39.1, "lon": 26.1, "marker": "eski"},
                "SS1": {"tip": "SS", "n1": [39.1, 26.1], "n2": [39.2, 26.2]},
            },
        }
    })

    assert sonuc["__HARITA__"]["isaretler"] == {
        "AÇ1": {"tip": "AÇ", "lat": 39.1, "lon": 26.1},
        "SS1": {"tip": "SS", "n1": [39.1, 26.1], "n2": [39.2, 26.2]},
    }


def test_tasima_rapor_imzasi_json_listesinden_canonical_tuple_olarak_yuklenir(yonetici):
    sonuc = yonetici.proje_verisini_normalize_et({
        "_TASIMA_": {
            "secim": "zemin",
            "girdiler": {},
            "rapor_imzasi": [
                "zemin",
                True,
                [["B", "1.0"], ["L", "2.0"]],
                "10.50",
            ],
        }
    })

    assert sonuc["_TASIMA_"]["rapor_imzasi"] == (
        "zemin",
        True,
        (("B", "1.0"), ("L", "2.0")),
        "10.50",
    )


def test_taahhut_bilgileri_normalize_topla_ve_yerlestir_round_trip(yonetici):
    kayit = {
        "JEOFIZIK_MUH_AD": "  Jeofizik İmza  ",
        "JEOFIZIK_MUH_SICIL": 1234,
        "JEOLOJI_MUH_AD": "Jeoloji İmza",
        "BILINMEYEN": "kaydedilmemeli",
    }

    normalize = yonetici.proje_verisini_normalize_et({"_TAAHHUT_BILGILERI_": kayit})
    yonetici._taahhut_bilgilerini_yerlestir(normalize["_TAAHHUT_BILGILERI_"])
    toplanan = yonetici._taahhut_bilgilerini_topla()

    assert set(toplanan) == set(TAAHHUT_BILGI_ALANLARI)
    assert toplanan["JEOFIZIK_MUH_AD"] == "Jeofizik İmza"
    assert toplanan["JEOFIZIK_MUH_SICIL"] == "1234"
    assert toplanan["JEOLOJI_MUH_AD"] == "Jeoloji İmza"
    assert toplanan["JEOLOJI_MUH_TELEFON"] == ""
    assert "BILINMEYEN" not in toplanan


def test_atomik_json_hatasi_mevcut_dosyayi_korur_ve_gecici_dosya_birakmaz(yonetici, tmp_path):
    hedef = tmp_path / "proje.json"
    hedef.write_text('{"eski": true}', encoding="utf-8")

    with pytest.raises(TypeError):
        yonetici.atomik_json_yaz(str(hedef), {"gecersiz": {1, 2}})

    assert json.loads(hedef.read_text(encoding="utf-8")) == {"eski": True}
    assert not list(tmp_path.glob(".proje.json.*.tmp"))


def test_yedekler_benzersizdir_ve_proje_bazinda_sinirlanir(yonetici, tmp_path):
    kaynak = tmp_path / "ornek.json"
    kaynak.write_text('{"deger": 1}', encoding="utf-8")

    uretilenler = {
        yonetici.otomatik_yedek_olustur(str(kaynak))
        for _ in range(YEDEK_LIMITI + 5)
    }

    assert "" not in uretilenler
    assert len(uretilenler) == YEDEK_LIMITI + 5
    kalanlar = list((tmp_path / "yedekler").glob("ornek_*.json"))
    assert len(kalanlar) == YEDEK_LIMITI
    assert all(Path(yol).name.startswith("ornek_") for yol in uretilenler)


def test_proje_gecisinde_tum_gecici_harita_ve_gorsel_durumu_temizlenir(tmp_path):
    app = SahteUygulama(tmp_path)
    app.root = SahteRoot()
    app.map_widget = SahteHarita()
    app._harita_yeniden_ciz_after_id = "after-1"
    app.temp_ss_marker = object()
    app.ss_ilk_nokta = (39.0, 26.0)
    app.kml_polygon_obj = object()
    app.harita_isaretleri = {"Merkez": {"tip": "M"}}
    app.harita_nokta_sayaclari = {"AÇ": 9, "YN": 8, "SS": 7}
    app.yuklu_kml_yolu = "eski.kml"
    app.yuklu_kml_points = [[39.0, 26.0]]
    app.aktif_harita_araci = SahteDegisken("SS")
    app.img_mjh = "eski-mjh.jpg"
    app.img_jeofizik_lok = "eski-jeofizik.jpg"
    app.img_jeoloji_lok = "eski-jeoloji.jpg"
    app.img_yerbulduru = "eski-yer.jpg"
    app.lbl_lab_excel = SahteEtiket("eski.xlsx")

    KayitYoneticisi(app).gecici_proje_durumunu_temizle()

    assert app.root.iptal_edilen == "after-1"
    assert app.map_widget.silinenler == ["marker", "path", "polygon"]
    assert app.temp_ss_marker is None
    assert app.ss_ilk_nokta is None
    assert app.kml_polygon_obj is None
    assert app.harita_isaretleri == {}
    assert app.harita_nokta_sayaclari == {"AÇ": 1, "YN": 1, "SS": 1}
    assert app.yuklu_kml_yolu == ""
    assert app.yuklu_kml_points == []
    assert app.aktif_harita_araci.deger == "Yok"
    assert app.img_mjh is None
    assert app.img_jeofizik_lok is None
    assert app.img_jeoloji_lok is None
    assert app.img_yerbulduru is None
    assert app.lbl_lab_excel.metin == "Yok"


def test_proje_dosya_yollari_goreli_kaydedilip_mutlak_cozulur(yonetici, tmp_path):
    proje_klasoru = tmp_path / "proje"
    varlik_klasoru = proje_klasoru / "varliklar"
    sablon_koku = proje_klasoru / "ornek_sablonlar"
    varlik_klasoru.mkdir(parents=True)
    sablon_koku.mkdir()
    proje_yolu = proje_klasoru / "ornek.json"

    yerel_yollar = {
        "rapor": varlik_klasoru / "rapor.docx",
        "jeoloji": varlik_klasoru / "jeoloji.docx",
        "lab": varlik_klasoru / "lab.xlsx",
        "jeofizik": varlik_klasoru / "jeofizik.xlsx",
        "kml": varlik_klasoru / "parsel.kml",
        "mjh": varlik_klasoru / "Mühendislik_Jeolojisi_Haritasi.jpg",
        "jeofizik_lokasyon": varlik_klasoru / "Jeofizik_Lokasyon_Haritasi.jpg",
        "jeoloji_lokasyon": varlik_klasoru / "Jeoloji_Lokasyon_Haritasi.jpg",
        "yerbulduru": varlik_klasoru / "Yerbulduru_Haritasi.jpg",
        "ek": varlik_klasoru / "ek.pdf",
        "legacy_ek": varlik_klasoru / "legacy.pdf",
    }
    paket_sablonu = sablon_koku / "taahhut.docx"
    harici_sablon = tmp_path / "harici.docx"
    yonetici.app.sablon_kok_adaylari = lambda: [str(sablon_koku)]

    veriler = {
        "_RAPOR_SABLONU_": str(yerel_yollar["rapor"]),
        "_TAAHHUT_WORD_SABLONU_": str(paket_sablonu),
        "_JEOLOJI_SABLONU_": str(yerel_yollar["jeoloji"]),
        "_LAB_KAYNAK_": str(yerel_yollar["lab"]),
        "_JEOFIZIK_": {"excel_yolu": str(yerel_yollar["jeofizik"])},
        "__HARITA__": {
            "kml_yolu": str(yerel_yollar["kml"]),
            "mjh_yolu": str(yerel_yollar["mjh"]),
            "jeofizik_lokasyon_yolu": str(yerel_yollar["jeofizik_lokasyon"]),
            "jeoloji_lokasyon_yolu": str(yerel_yollar["jeoloji_lokasyon"]),
            "yerbulduru_yolu": str(yerel_yollar["yerbulduru"]),
        },
        "_EKLER_": {
            "EVRAKLAR": [{"baslik": "Ek", "yol": str(yerel_yollar["ek"])}],
            "LOG": [str(yerel_yollar["legacy_ek"])],
        },
        "HARICI": str(harici_sablon),
    }

    kayit_verileri = yonetici.proje_yollarini_kayda_hazirla(veriler, str(proje_yolu))

    assert not os.path.isabs(kayit_verileri["_RAPOR_SABLONU_"])
    assert not os.path.isabs(kayit_verileri["_LAB_KAYNAK_"])
    assert not os.path.isabs(kayit_verileri["_JEOFIZIK_"]["excel_yolu"])
    assert not os.path.isabs(kayit_verileri["__HARITA__"]["kml_yolu"])
    assert not os.path.isabs(kayit_verileri["__HARITA__"]["mjh_yolu"])
    assert not os.path.isabs(kayit_verileri["__HARITA__"]["jeofizik_lokasyon_yolu"])
    assert not os.path.isabs(kayit_verileri["__HARITA__"]["jeoloji_lokasyon_yolu"])
    assert not os.path.isabs(kayit_verileri["__HARITA__"]["yerbulduru_yolu"])
    assert not os.path.isabs(kayit_verileri["_EKLER_"]["EVRAKLAR"][0]["yol"])
    assert not os.path.isabs(kayit_verileri["_EKLER_"]["LOG"][0])
    assert kayit_verileri["_TAAHHUT_WORD_SABLONU_"] == str(paket_sablonu)
    assert veriler["_RAPOR_SABLONU_"] == str(yerel_yollar["rapor"])

    cozulen = yonetici.proje_yollarini_coz(kayit_verileri, str(proje_yolu))
    assert cozulen["_RAPOR_SABLONU_"] == str(yerel_yollar["rapor"])
    assert cozulen["_LAB_KAYNAK_"] == str(yerel_yollar["lab"])
    assert cozulen["_JEOFIZIK_"]["excel_yolu"] == str(yerel_yollar["jeofizik"])
    assert cozulen["__HARITA__"]["kml_yolu"] == str(yerel_yollar["kml"])
    assert cozulen["__HARITA__"]["mjh_yolu"] == str(yerel_yollar["mjh"])
    assert cozulen["__HARITA__"]["jeofizik_lokasyon_yolu"] == str(yerel_yollar["jeofizik_lokasyon"])
    assert cozulen["__HARITA__"]["jeoloji_lokasyon_yolu"] == str(yerel_yollar["jeoloji_lokasyon"])
    assert cozulen["__HARITA__"]["yerbulduru_yolu"] == str(yerel_yollar["yerbulduru"])
    assert cozulen["_EKLER_"]["EVRAKLAR"][0]["yol"] == str(yerel_yollar["ek"])
    assert cozulen["_EKLER_"]["LOG"][0] == str(yerel_yollar["legacy_ek"])


def test_farkli_kaydet_ve_ac_tasinabilir_yollari_donusturur(tmp_path, monkeypatch):
    proje_klasoru = tmp_path / "tasinabilir"
    proje_klasoru.mkdir()
    lab_yolu = proje_klasoru / "lab.xlsx"
    proje_yolu = proje_klasoru / "proje.json"

    app = SahteUygulama(tmp_path / "app1")
    app.root = SahteBaslikRoot()
    app.durum = {"_LAB_KAYNAK_": str(lab_yolu)}
    app.guncel_dosya_yolu = None
    app.son_kayit_verisi = copy.deepcopy(app.durum)
    app.son_projeler = []
    app.son_proje_hatasi = False
    yonetici = KayitAkisiYoneticisi(app)
    monkeypatch.setattr("kayit.filedialog.asksaveasfilename", lambda **kwargs: str(proje_yolu))

    assert yonetici.farkli_kaydet() is True
    ham_veri = json.loads(proje_yolu.read_text(encoding="utf-8"))
    assert ham_veri["_LAB_KAYNAK_"] == "lab.xlsx"
    assert app.durum["_LAB_KAYNAK_"] == str(lab_yolu)

    acilan_app = SahteUygulama(tmp_path / "app2")
    acilan_app.root = SahteBaslikRoot()
    acilan_app.durum = {"eski": True}
    acilan_app.guncel_dosya_yolu = None
    acilan_app.son_kayit_verisi = copy.deepcopy(acilan_app.durum)
    acilan_app.son_projeler = []
    acilan_app.son_proje_hatasi = False
    acan_yonetici = KayitAkisiYoneticisi(acilan_app)

    assert acan_yonetici.proje_dosyasini_ac(str(proje_yolu)) is True
    assert acilan_app.durum["_LAB_KAYNAK_"] == str(lab_yolu)


def test_gec_hatada_proje_acma_dosyayi_basligi_ve_veriyi_geri_alir(tmp_path, monkeypatch):
    eski_yol = tmp_path / "eski.json"
    yeni_yol = tmp_path / "yeni.json"
    yeni_yol.write_text(json.dumps({"PROJE_ADI": "Yeni"}), encoding="utf-8")

    app = SahteUygulama(tmp_path)
    app.root = SahteBaslikRoot(f"K-1 - {eski_yol}")
    app.durum = {"PROJE_ADI": "Eski"}
    app.guncel_dosya_yolu = str(eski_yol)
    app.son_kayit_verisi = {"PROJE_ADI": "Eski"}
    app.son_projeler = [str(eski_yol)]
    app.son_proje_hatasi = True
    yonetici = KayitAkisiYoneticisi(app)
    monkeypatch.setattr("kayit.messagebox.showerror", lambda *args, **kwargs: None)

    assert yonetici.proje_dosyasini_ac(str(yeni_yol)) is False
    assert app.durum == {"PROJE_ADI": "Eski"}
    assert app.guncel_dosya_yolu == str(eski_yol)
    assert app.root.baslik == f"K-1 - {eski_yol}"
    assert app.son_kayit_verisi == {"PROJE_ADI": "Eski"}
    assert app.son_projeler == [str(eski_yol)]
