from pathlib import Path
from types import SimpleNamespace

from jeoloji_kunye_uzlasma import kunye_uzlasmasi_olustur


def _word(path, *, ilce="Bayramiç", yerlesim="Akçakıl", ada="0", parsel="1"):
    return SimpleNamespace(
        dosya_yolu=str(path),
        ilce=ilce,
        yerlesim=yerlesim,
        ada=ada,
        parsel=parsel,
    )


def _kml(path, ad=""):
    return {"dosya_yolu": str(path), "poligonlar": [{"ad": ad, "aciklama": ""}]}


def test_ada_sifir_parsel_kml_ile_semantik_eslesir(tmp_path):
    root = tmp_path / "BAYRAMİÇ"
    project = root / "AKÇAKIL" / "Bayramiç Akçakıl 0-673"
    word = _word(project / "Bayramiç Akçakıl 0-673 Rapor.docx", parsel="673")
    kml = _kml(project / "tkgm-parsel-sorgu-sonuc-673-parsel.kml")

    result = kunye_uzlasmasi_olustur(
        secili_root=root, proje_klasoru=project, word_sonucu=word, kml_adayi=kml
    )

    assert result["hazir"] is True
    assert result["durum"] == "ada0"
    assert result["kanonik"]["ada"] == "0"
    assert result["kanonik"]["parsel"] == "673"


def test_klasor_dosya_ve_kml_word_icerigindeki_eski_kunyeyi_duzeltir(tmp_path):
    root = tmp_path / "BAYRAMİÇ"
    project = root / "AKÇAKIL" / "Bayramiç Akçakıl 679"
    word = _word(
        project / "Bayramiç AKÇAKIL 679.docx",
        yerlesim="Zeytinli",
        ada="0",
        parsel="1",
    )
    kml = _kml(project / "tkgm-parsel-sorgu-sonuc-679-parsel.kml")

    result = kunye_uzlasmasi_olustur(
        secili_root=root, proje_klasoru=project, word_sonucu=word, kml_adayi=kml
    )

    assert result["hazir"] is True
    assert result["durum"] == "duzeltildi"
    assert result["kanonik"]["yerlesim"] == "AKÇAKIL"
    assert result["kanonik"]["ada"] == "0"
    assert result["kanonik"]["parsel"] == "679"
    assert any("Zeytinli" in warning for warning in result["uyarilar"])
    assert any("'1' yerine '679'" in warning for warning in result["uyarilar"])


def test_acik_ada_celiskisi_otomatik_aktarilmaz(tmp_path):
    root = tmp_path / "AYVACIK"
    project = root / "KOZLU" / "Ayvacık Kozlu 151-10"
    word = _word(
        project / "Ayvacık Kozlu 151-10 Rapor.docx",
        ilce="Ayvacık",
        yerlesim="Kozlu",
        ada="151",
        parsel="10",
    )
    kml = _kml(project / "tkgm-152-ada-10-parsel.kml")

    result = kunye_uzlasmasi_olustur(
        secili_root=root, proje_klasoru=project, word_sonucu=word, kml_adayi=kml
    )

    assert result["hazir"] is False
    assert result["durum"] == "celiski"
    assert result["celiskiler"]
