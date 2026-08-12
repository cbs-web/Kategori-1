from types import SimpleNamespace

from PIL import Image

from cizimler import CizimUretici


def test_tkgm_kaynak_metni_url_ve_altligi_yazar():
    app = SimpleNamespace(
        yuklu_kml_yolu="C:/proje/TKGM_0_656.kml",
        parsel_haritasi_kaynak_url=(
            "https://cbsapi.tkgm.gov.tr/megsiswebapi.v3.1/api/parsel/123/0/656"
        ),
    )
    metin = CizimUretici(app).parsel_haritasi_kaynak_metni()
    assert "https://cbsapi.tkgm.gov.tr" in metin
    assert "Altlık: Google Uydu" in metin
    assert "Taslak parsel gösterimidir" not in metin


def test_yerbulduru_ciktisi_a4_portre_oranina_yakindir(tmp_path):
    app = SimpleNamespace()
    uretici = CizimUretici(app)
    genis = Image.new("RGB", (1000, 600), "green")
    yakin = Image.new("RGB", (1000, 600), "blue")
    hedef = tmp_path / "yerbulduru.jpg"

    uretici.yerbulduru_iki_panel_resim_olustur(
        genis,
        yakin,
        hedef,
        genis_pin=(500, 300),
        yakin_pin=(500, 300),
    )

    with Image.open(hedef) as image:
        oran = image.height / image.width
    assert 1.35 <= oran <= 1.55
