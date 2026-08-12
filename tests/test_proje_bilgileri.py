from temel_bilgiler_islemleri import PROJE_ALANLARI


def test_proje_bilgileri_sekmesinde_firma_idare_ve_rapor_alanlari_yoktur():
    alan_kodlari = {kod for _, kod in PROJE_ALANLARI}

    assert alan_kodlari.isdisjoint(
        {
            "FIRMA_ADI",
            "FIRMA_ADRESI",
            "FIRMA_TELEFON",
            "ILGILI_IDARE",
            "RAPOR_TARIHI",
        }
    )
