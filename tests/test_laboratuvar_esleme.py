from laboratuvar import (
    laboratuvar_numune_anahtari,
    laboratuvar_satirlarini_birlestir,
)


def test_ac_yazim_bicimleri_ayni_numune_anahtarini_verir():
    anahtarlar = {
        laboratuvar_numune_anahtari(value)
        for value in ("AÇ1", "AÇ-1", "AÇ 1", "AC1")
    }
    assert anahtarlar == {"AC1"}


def test_excel_ac_1_mevcut_ac1_satirini_gunceller_ve_etiketi_korur():
    mevcut = [("AÇ1", "0.00-1.00", "eski")]
    gelen = [("AÇ-1", "0,00 - 1,00", "yeni")]

    sonuc = laboratuvar_satirlarini_birlestir(mevcut, gelen)

    assert sonuc["eklenen"] == 0
    assert sonuc["guncellenen"] == 1
    assert sonuc["satirlar"] == [("AÇ1", "0,00 - 1,00", "yeni")]
