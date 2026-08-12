from on_deger import bos_is_akisi_verisi, bos_on_deger_verisi, is_durumu_degistir, on_deger_revizyonu_ekle
from proje_durumu_islemleri import parsel_haritasi_durumu_hazirla, proje_durum_ozeti_hazirla


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
