from harita_renkleri import (
    CALISAN_PARSEL_SINIR_RENGI,
    JEOLOJI_HARITA_RENK_ACIKLAMASI,
    JEOLOJI_ONAYLI_DOLGU_RENGI,
    JEOLOJI_ONAYLI_SINIR_RENGI,
    JEOLOJI_TASLAK_SINIR_RENGI,
)


def test_jeoloji_harita_renk_anlamlari_birbirinden_ayri():
    assert JEOLOJI_ONAYLI_DOLGU_RENGI == "#ef9a9a"
    assert JEOLOJI_ONAYLI_SINIR_RENGI == "#b71c1c"
    assert len({CALISAN_PARSEL_SINIR_RENGI, JEOLOJI_ONAYLI_SINIR_RENGI, JEOLOJI_TASLAK_SINIR_RENGI}) == 3
    assert JEOLOJI_HARITA_RENK_ACIKLAMASI == (
        "Mavi: çalışan parsel · Kırmızı: onaylı rapor · Turuncu: taslak"
    )
