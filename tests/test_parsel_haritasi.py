from cizimler import parsel_gorunum_hesapla, parsel_noktalari_hashi


def test_parsel_gorunumu_poligon_merkezini_bulup_makul_zoom_secer():
    points = [
        (39.52390, 26.11980),
        (39.52390, 26.12030),
        (39.52425, 26.12030),
        (39.52425, 26.11980),
        (39.52390, 26.11980),
    ]

    center, zoom = parsel_gorunum_hesapla(points, 1000, 650)

    assert 39.52390 <= center[0] <= 39.52425
    assert 26.11980 <= center[1] <= 26.12030
    assert 17 <= zoom <= 21
    assert parsel_noktalari_hashi(points) == parsel_noktalari_hashi(list(points))
