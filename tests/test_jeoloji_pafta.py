from zipfile import ZIP_DEFLATED, ZipFile

from PIL import Image

from jeoloji_pafta_kutuphanesi import kmz_bilgilerini_oku, pafta_gorsel_koordinati
from jeoloji_pafta_tanima import paftada_birim_tahmin_et


def test_parsel_lejant_ornegindeki_birimi_secer(tmp_path):
    overlay = Image.new("RGB", (240, 120), "#b13a31")
    for x in range(120, 240):
        for y in range(120):
            overlay.putpixel((x, y), (45, 86, 170))
    overlay_path = tmp_path / "overlay.png"
    overlay.save(overlay_path)

    kml = """<?xml version="1.0" encoding="UTF-8"?>
    <kml xmlns="http://www.opengis.net/kml/2.2"><Document><GroundOverlay>
      <name>Pilot Pafta</name><Icon><href>overlay.png</href></Icon>
      <LatLonBox><north>40</north><south>39</south><east>27</east><west>26</west></LatLonBox>
    </GroundOverlay></Document></kml>"""
    kmz_path = tmp_path / "pilot.kmz"
    with ZipFile(kmz_path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("doc.kml", kml)
        archive.write(overlay_path, "overlay.png")

    legend = Image.new("RGB", (240, 80), "white")
    for x in range(0, 100):
        for y in range(80):
            legend.putpixel((x, y), (177, 58, 49))
    for x in range(140, 240):
        for y in range(80):
            legend.putpixel((x, y), (45, 86, 170))
    legend_path = tmp_path / "legend.jpg"
    legend.save(legend_path, quality=100, subsampling=0)

    record = kmz_bilgilerini_oku(kmz_path)[0]
    record["lejant_id"] = "pilot"
    profile = {
        "jpeg_path": str(legend_path),
        "ogeler": [
            {"id": "red", "kod": "Kr", "ad": "Kırmızı Birim", "rect": [0, 0, 0.40, 1]},
            {"id": "blue", "kod": "Mb", "ad": "Mavi Birim", "rect": [0.60, 0, 1, 1]},
        ],
    }
    parcel = [(39.25, 26.10), (39.25, 26.35), (39.55, 26.35), (39.55, 26.10)]
    result = paftada_birim_tahmin_et(record, profile, parcel)
    try:
        assert result["adaylar"][0]["kod"] == "Kr"
        assert result["ornek_sayisi"] > 0
    finally:
        result["kanit_gorseli"].close()
        for image in result.get("lejant_gorselleri", []):
            if image is not None:
                image.close()


def test_dondurulmus_groundoverlay_noktayi_dogru_piksele_tasir():
    record = {
        "bounds": {"north": 1, "south": -1, "east": 1, "west": -1},
        "rotation": 90,
    }
    # Kaynak görüntüde (0.75, 0.25) olan nokta +90 derece dönüşten sonra
    # dünya koordinatında lat=0.5, lon=-0.5 konumuna gelir.
    x, y = pafta_gorsel_koordinati(record, 0.5, -0.5)
    assert abs(x - 0.75) < 1.0e-9
    assert abs(y - 0.25) < 1.0e-9
