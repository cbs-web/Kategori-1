from pathlib import Path

import pytest

from harita_islemleri import kml_poligonlarini_oku
from tkgm_kml import TKGMSorguHatasi, geojson_kml_olustur, tkgm_parsel_kml_olustur


def _feature(ada="0", parsel="673"):
    return {
        "type": "Feature",
        "geometry": {
            "type": "Polygon",
            "coordinates": [
                [
                    [26.4000, 39.8000],
                    [26.4010, 39.8000],
                    [26.4010, 39.8010],
                    [26.4000, 39.8010],
                    [26.4000, 39.8000],
                ]
            ],
        },
        "properties": {
            "adaNo": ada,
            "parselNo": parsel,
            "mahalleId": 3,
        },
    }


def _fetcher_olustur(feature):
    urls = []

    def fetcher(url, timeout=25):
        del timeout
        urls.append(url)
        if url.endswith("ilListe.json"):
            return [{"id": 1, "text": "Çanakkale"}]
        if "ilceListe/1" in url:
            return [{"id": 2, "text": "Bayramiç"}]
        if "mahalleListe/2" in url:
            return [{"id": 3, "text": "Akçakıl Köyü"}]
        return feature

    return fetcher, urls


def test_ada_sifir_tkgm_kml_uretilir_ve_k1_okur(tmp_path):
    fetcher, urls = _fetcher_olustur(_feature())
    result = tkgm_parsel_kml_olustur(
        {
            "il": "Çanakkale",
            "ilce": "Bayramiç",
            "koy": "Akçakıl",
            "ada": "0",
            "parsel": "673",
        },
        tmp_path,
        fetcher=fetcher,
    )

    assert urls[-1].endswith("/parsel/3/0/673")
    assert Path(result["path"]).name == "TKGM_0_673.kml"
    assert len(kml_poligonlarini_oku(result["path"])) == 1


def test_tkgm_farkli_parsel_dondururse_dosya_yazilmaz(tmp_path):
    fetcher, _urls = _fetcher_olustur(_feature(parsel="674"))

    with pytest.raises(TKGMSorguHatasi, match="farklı parsel"):
        tkgm_parsel_kml_olustur(
            {
                "il": "Çanakkale",
                "ilce": "Bayramiç",
                "koy": "Akçakıl",
                "ada": "0",
                "parsel": "673",
            },
            tmp_path,
            fetcher=fetcher,
        )

    assert list(tmp_path.iterdir()) == []


def test_gecersiz_tkgm_koordinati_reddedilir():
    geometry = _feature()["geometry"]
    geometry["coordinates"][0][1] = [26.4010, 95.0]

    with pytest.raises(TKGMSorguHatasi, match="WGS84"):
        geojson_kml_olustur(geometry)
