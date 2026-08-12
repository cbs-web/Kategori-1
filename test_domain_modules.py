import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import pandas as pd
from PIL import Image, ImageDraw

from cizimler import CizimUretici, derinlik_araligi_oku, guvenli_dosya_adi
from harita_islemleri import (
    isim_bazli_birlestirme_plani,
    kml_poligon_koordinatlarini_oku,
    kml_poligonlarini_oku,
)
from jeofizik_islemleri import jeofon_dizilim_bilgilerini_dogrula
from laboratuvar import (
    laboratuvar_dataframe_satirlari,
    laboratuvar_dosyasi_oku,
    laboratuvar_pano_verisini_donustur,
    laboratuvar_satirlarini_birlestir,
)
from laboratuvar_islemleri import lab_eslesen_satiri_bul


class SahteEntry:
    def __init__(self, deger):
        self.deger = deger

    def get(self, *_args):
        return self.deger


class HaritaVeKmlTestleri(unittest.TestCase):
    def test_isim_bazli_plan_veriyi_orphan_olarak_korur(self):
        plan = isim_bazli_birlestirme_plani(["AÇ1", "AÇ2"], ["AÇ2", "YN1"])
        self.assertEqual(plan["guncellenecek"], ["AÇ2"])
        self.assertEqual(plan["eklenecek"], ["YN1"])
        self.assertEqual(plan["orphan"], ["AÇ1"])

    def test_kml_yalniz_polygon_ve_gecerli_koordinat_kabul_eder(self):
        icerik = b"""<?xml version='1.0' encoding='UTF-8'?>
        <kml xmlns='http://www.opengis.net/kml/2.2'><Document><Placemark><Polygon>
        <outerBoundaryIs><LinearRing><coordinates>
        26.0,40.0,0 26.1,40.0,0 26.1,40.1,0 26.0,40.0,0
        </coordinates></LinearRing></outerBoundaryIs></Polygon></Placemark></Document></kml>"""
        with tempfile.TemporaryDirectory() as klasor:
            yol = Path(klasor, "alan.kml")
            yol.write_bytes(icerik)
            noktalar = kml_poligon_koordinatlarini_oku(yol)
        self.assertEqual(noktalar[0], (40.0, 26.0))
        self.assertEqual(len(noktalar), 4)

    def test_kml_linestring_polygon_diye_kabul_edilmez(self):
        with tempfile.TemporaryDirectory() as klasor:
            yol = Path(klasor, "cizgi.kml")
            yol.write_text(
                "<kml><LineString><coordinates>26,40 27,40 27,41 26,40</coordinates></LineString></kml>",
                encoding="utf-8",
            )
            with self.assertRaises(ValueError):
                kml_poligon_koordinatlarini_oku(yol)

    def test_kml_ic_halka_dis_sinir_yerine_kullanilmaz(self):
        icerik = """<kml><Polygon>
        <innerBoundaryIs><LinearRing><coordinates>
        30,40 30.1,40 30.1,40.1 30,40
        </coordinates></LinearRing></innerBoundaryIs>
        <outerBoundaryIs><LinearRing><coordinates>
        26,40 26.1,40 26.1,40.1 26,40
        </coordinates></LinearRing></outerBoundaryIs>
        </Polygon></kml>"""
        with tempfile.TemporaryDirectory() as klasor:
            yol = Path(klasor, "halkalar.kml")
            yol.write_text(icerik, encoding="utf-8")
            noktalar = kml_poligon_koordinatlarini_oku(yol)
        self.assertEqual(noktalar[0], (40.0, 26.0))

    def test_kml_polygon_altindaki_serbest_koordinat_gecersizdir(self):
        with tempfile.TemporaryDirectory() as klasor:
            yol = Path(klasor, "eksik_halka.kml")
            yol.write_text(
                "<kml><Polygon><coordinates>26,40 27,40 27,41 26,40</coordinates></Polygon></kml>",
                encoding="utf-8",
            )
            with self.assertRaises(ValueError):
                kml_poligon_koordinatlarini_oku(yol)

    def test_kml_aralik_disi_koordinati_reddeder(self):
        with tempfile.TemporaryDirectory() as klasor:
            yol = Path(klasor, "bozuk.kml")
            yol.write_text(
                "<kml><Polygon><coordinates>26,95 27,40 27,41 26,95</coordinates></Polygon></kml>",
                encoding="utf-8",
            )
            with self.assertRaises(ValueError):
                kml_poligon_koordinatlarini_oku(yol)

    def test_kml_multigeometry_icindeki_butun_poligonlari_okur(self):
        icerik = """<kml><Placemark><name>12 ada 3 parsel</name><MultiGeometry>
        <Polygon><outerBoundaryIs><LinearRing><coordinates>
        26,40 26.01,40 26.01,40.01 26,40
        </coordinates></LinearRing></outerBoundaryIs></Polygon>
        <Polygon><outerBoundaryIs><LinearRing><coordinates>
        26.02,40 26.03,40 26.03,40.01 26.02,40
        </coordinates></LinearRing></outerBoundaryIs></Polygon>
        </MultiGeometry></Placemark></kml>"""
        with tempfile.TemporaryDirectory() as klasor:
            yol = Path(klasor, "coklu.kml")
            yol.write_text(icerik, encoding="utf-8")
            poligonlar = kml_poligonlarini_oku(yol)
        self.assertEqual(len(poligonlar), 2)
        self.assertEqual(poligonlar[0]["ad"], "12 ada 3 parsel")


class LaboratuvarTestleri(unittest.TestCase):
    @staticmethod
    def lab1_satiri(no="AÇ-1", derinlik="0,00-3,00"):
        satir = [""] * 29
        satir[0] = no
        satir[2] = derinlik
        satir[3] = "0"
        satir[4] = "7,68"
        satir[5] = "90,32"
        satir[6] = "2"
        satir[7] = "51,2"
        satir[8] = "24,7"
        satir[9] = "26,5"
        satir[10] = "31,9"
        satir[11] = "1,866"
        satir[12] = "1,412"
        satir[14] = "CIH"
        satir[20] = "94,11"
        satir[21] = "7,93"
        satir[28] = "6,4"
        return satir

    def test_lab1_pano_satiri_k1_ac_sutunlarina_eslesir(self):
        baslik = ["Sondaj No", "Numune No", "Derinlik"] + [""] * 26
        pano = "\n".join("\t".join(row) for row in (baslik, self.lab1_satiri()))

        sonuc = laboratuvar_pano_verisini_donustur(pano, "ac")

        self.assertEqual(sonuc["format"], "lab1")
        self.assertEqual(
            sonuc["ac_satirlari"],
            [("AÇ-1", "0,00-3,00", "0", "7.68", "92.32", "51.2", "24.7", "26.5", "31.9", "1.866", "1.412", "CIH", "94.11", "7.93")],
        )

    def test_lab1_pano_ac_yn_ayirir_ve_ayni_kaydi_gunceller(self):
        ac = self.lab1_satiri()
        yn = self.lab1_satiri(no="YN-1", derinlik="0-0,20")
        pano = "\n".join("\t".join(row) for row in (ac, yn))

        sonuc = laboratuvar_pano_verisini_donustur(pano, "ac")
        birlesim = laboratuvar_satirlarini_birlestir(
            [("AÇ1", "0-3", "eski")], sonuc["ac_satirlari"]
        )

        self.assertEqual(len(sonuc["ac_satirlari"]), 1)
        self.assertEqual(sonuc["yn_satirlari"], [("YN-1", "0-0,20", "1.866", "6.4")])
        self.assertEqual(birlesim["eklenen"], 0)
        self.assertEqual(birlesim["guncellenen"], 1)
        self.assertEqual(birlesim["satirlar"][0][2], "0")

    def test_eski_on_dort_sutunlu_ac_panosu_korunur(self):
        degerler = ["AÇ-2", "1-2"] + [str(index) for index in range(12)]
        sonuc = laboratuvar_pano_verisini_donustur("\t".join(degerler), "ac")
        self.assertEqual(sonuc["format"], "standart_ac")
        self.assertEqual(sonuc["standart_satirlar"], [tuple(degerler)])

    def test_liquid_limit_id_alt_dizesi_numune_kolonunu_calamaz(self):
        df = pd.DataFrame(
            [[50, "AÇ1", "0-1"]],
            columns=["Liquid Limit", "Sample No", "Depth"],
        )
        sonuc = laboratuvar_dataframe_satirlari(df)
        self.assertEqual(sonuc["ac_satirlari"][0][0], "AÇ1")
        self.assertEqual(sonuc["ac_satirlari"][0][5], "50")

    def test_yn_is50_degerini_aktarir(self):
        df = pd.DataFrame(
            [["YN1", "0-0.2", 2.65, 7.4]],
            columns=["No", "Depth", "BHA", "IS50"],
        )
        sonuc = laboratuvar_dataframe_satirlari(df)
        self.assertEqual(sonuc["yn_satirlari"], [("YN1", "0-0.2", "2.65", "7.4")])

    def test_buyuk_harfli_cp1254_csv_okunur(self):
        with tempfile.TemporaryDirectory() as klasor:
            yol = Path(klasor, "LAB.CSV")
            yol.write_bytes("No;Derinlik;LL\nAÇ1;0-1;45\n".encode("cp1254"))
            sonuc = laboratuvar_dosyasi_oku(yol)
        self.assertEqual(sonuc["ac_satirlari"][0][0], "AÇ1")
        self.assertEqual(sonuc["ac_satirlari"][0][5], "45")

    def test_laboratuvar_desteklenmeyen_uzantiyi_reddeder(self):
        with tempfile.TemporaryDirectory() as klasor:
            yol = Path(klasor, "laboratuvar.txt")
            yol.write_text("No,Derinlik\nAÇ1,0-1\n", encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "xlsx"):
                laboratuvar_dosyasi_oku(yol)

    def test_laboratuvar_asiri_satir_sayisini_reddeder(self):
        with tempfile.TemporaryDirectory() as klasor:
            yol = Path(klasor, "buyuk.csv")
            satirlar = ["No,Derinlik"] + [f"AÇ{i},0-1" for i in range(10001)]
            yol.write_text("\n".join(satirlar), encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "en fazla 10000"):
                laboratuvar_dosyasi_oku(yol)

    def test_tek_lab_sonucu_farkli_derinlige_yayilmaz(self):
        satir = {"orijinal_derinlik": "1-2", "LL": "40"}
        self.assertIsNone(lab_eslesen_satiri_bul([satir], "2-3"))
        self.assertIs(lab_eslesen_satiri_bul([satir], "1,0 - 2,0"), satir)


class CizimTestleri(unittest.TestCase):
    def test_dosya_adi_klasor_kacisini_engeller(self):
        ad = guvenli_dosya_adi(r"..\C:\temp/rapor:*?")
        self.assertNotIn("/", ad)
        self.assertNotIn("\\", ad)
        self.assertNotIn(":", ad)
        self.assertFalse(ad.startswith(".."))
        self.assertEqual(guvenli_dosya_adi("CON.txt"), "_CON.txt")

    def test_derinlik_araligi_gecersiz_ve_ters_araligi_reddeder(self):
        self.assertEqual(derinlik_araligi_oku("0,0 - 3,5"), (0.0, 3.5))
        self.assertIsNone(derinlik_araligi_oku("3-2"))
        self.assertIsNone(derinlik_araligi_oku("metin"))

    def test_mevcut_cikti_icin_uzerine_yazma_onayi_ister(self):
        with tempfile.TemporaryDirectory() as klasor:
            yol = Path(klasor, "mevcut.jpg")
            yol.write_bytes(b"x")
            with patch("cizimler.messagebox.askyesno", return_value=False) as sor:
                sonuc = CizimUretici(SimpleNamespace()).mevcut_dosyalar_icin_onay_al([yol])
        self.assertFalse(sonuc)
        sor.assert_called_once()

    def test_tarama_deseni_deterministiktir(self):
        uretici = CizimUretici(SimpleNamespace())
        bir = Image.new("RGB", (180, 120), "white")
        iki = Image.new("RGB", (180, 120), "white")
        uretici.ciz_tarama_deseni(ImageDraw.Draw(bir), "KUM", 5, 5, 175, 115)
        uretici.ciz_tarama_deseni(ImageDraw.Draw(iki), "KUM", 5, 5, 175, 115)
        self.assertEqual(bir.tobytes(), iki.tobytes())

    def test_kesit_tum_kayitlari_ve_azami_derinligi_kapsar(self):
        uretici = CizimUretici(SimpleNamespace())
        kayitlar = [
            {"isim": "AÇ1", "satirlar": [["0-2", "", "", "KİL"]]},
            {"isim": "AÇ2", "satirlar": [["0-4", "", "", "SİLT"]]},
            {"isim": "YN1", "satirlar": [["0-7.5", "", "", "KUM"]]},
        ]
        with tempfile.TemporaryDirectory() as klasor:
            yol = Path(klasor, "kesit.jpg")
            uretici.kesit_ciz_olustur(kayitlar, yol)
            with Image.open(yol) as img:
                self.assertGreaterEqual(img.width, 850)
                self.assertGreater(img.height, 900)

    def test_log_aciklamayi_dogru_widgettan_alir_ve_derinligi_kirpmaz(self):
        app = SimpleNamespace()
        app.veri_alanlari = {
            "PROJE_ADI": SahteEntry("Test Projesi"),
            "IL": SahteEntry("Çanakkale"),
            "ILCE": SahteEntry("Merkez"),
            "ADA": SahteEntry("1"),
            "PARSEL": SahteEntry("2"),
        }
        app.ac_yn_satirlari = lambda kayit: kayit["satirlar"]
        uretici = CizimUretici(app)
        kayit = {
            "derinlik_entry": SahteEntry("7"),
            "enlem_entry": SahteEntry("40"),
            "boylam_entry": SahteEntry("26"),
            "tarih_entry": SahteEntry("01/01/2026"),
            "aciklama_text": SahteEntry("Açıklama"),
            "satirlar": [["0-7", "DS1", "-", "KİLTAŞI", "CL", "", "", "", "", "", ""]],
        }
        with tempfile.TemporaryDirectory() as klasor:
            yol = Path(klasor, "log.jpg")
            uretici.tekil_log_ciz(kayit, "AÇ1", yol)
            with Image.open(yol) as img:
                self.assertGreater(img.height, 2262)


class JeofizikTestleri(unittest.TestCase):
    def test_jeofon_girdileri_sonlu_ve_aralikli_olmalidir(self):
        sonuc, hatalar = jeofon_dizilim_bilgilerini_dogrula(
            {"jeofon_sayisi": "12.5", "jeofon_araligi": "-2", "duz_offset": "nan"}
        )
        self.assertGreaterEqual(len(hatalar), 3)
        self.assertEqual(sonuc["jeofon_sayisi"], "12")
        self.assertEqual(sonuc["jeofon_araligi"], "2")
        self.assertEqual(sonuc["duz_offset"], "0")


if __name__ == "__main__":
    unittest.main()
