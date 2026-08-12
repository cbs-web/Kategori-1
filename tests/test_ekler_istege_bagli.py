from ekler import (
    ek_kategori_durumunu_hazirla,
    ek_taahhutname_mi,
    ekleri_denetle,
)


def test_bos_ek_kategorisi_zorunlu_hata_degildir():
    metin, durum = ek_kategori_durumunu_hazirla({"EVRAKLAR": []}, "EVRAKLAR")
    assert "İsteğe bağlı" in metin
    assert durum == "secondary"


def test_taahhutname_ekler_listesinden_cikarilir():
    ek = {"baslik": "Jeoloji Taahhütnamesi", "yol": "C:/rapor/taahhutname.pdf"}
    assert ek_taahhutname_mi(ek)
    denetim = ekleri_denetle({"EVRAKLAR": [ek]}, ["EVRAKLAR"])
    assert denetim["toplam"] == 0
    assert denetim["sirali"] == []
