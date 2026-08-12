import math

import pytest

from tasima import TBDY2018TasimaGucu
from tasima_islemleri import TasimaIslemleri, tasima_qt_asagi_yuvarla


class _SahteDeger:
    def __init__(self, deger):
        self.deger = deger

    def get(self, *args):
        return self.deger

    def set(self, deger):
        self.deger = deger

    def delete(self, *args):
        self.deger = ""

    def insert(self, index, deger):
        self.deger = str(deger)

    def config(self, **ayarlar):
        if "text" in ayarlar:
            self.deger = ayarlar["text"]

    def cget(self, anahtar):
        return self.deger if anahtar == "text" else ""


def _sahte_zemin_uygulamasi():
    class SahteUygulama:
        def hata_kaydet(self, *args):
            raise AssertionError(f"Beklenmeyen hata kaydı: {args}")

    app = SahteUygulama()
    app.tg_girdiler = {
        kod: _SahteDeger(deger)
        for kod, deger in {
            "c": "10",
            "phi": "30",
            "gn": "1.8",
            "gsat": "2.0",
            "yass": "999",
            "B": "1.5",
            "L": "2.0",
            "Df": "1.0",
            "RvGk": "1.4",
            "ks_carpani": "56",
            "qt": "34",
            "Gk": "3",
        }.items()
    }
    app.entry_qt_nihai = _SahteDeger("")
    app.entry_ks_nihai = _SahteDeger("")
    app.lbl_sonuc = _SahteDeger("-")
    app.txt_tasima_rapor = _SahteDeger("")
    app.tasima_varsayim_onayi = _SahteDeger(True)
    app.zemin_kaya_var = _SahteDeger("zemin")
    app.tasima_rapor_imzasi = None
    return app


def test_tasima_gucu_su_seviyesini_hesaba_katar():
    kuru = TBDY2018TasimaGucu(10, 30, 18.0, 20.0, 999.0)
    sulu = TBDY2018TasimaGucu(10, 30, 18.0, 20.0, 0.0)

    kuru_qk, kuru_qt = kuru.analiz_yap(1.5, 2.0, 1.0)
    sulu_qk, sulu_qt = sulu.analiz_yap(1.5, 2.0, 1.0)

    assert sulu_qk < kuru_qk
    assert sulu_qt < kuru_qt


def test_yass_tiki_derinligin_zorunlulugunu_ve_hesap_degerini_belirler():
    yass_yok = _sahte_zemin_uygulamasi()
    yass_yok.tasima_yass_var = _SahteDeger(False)
    yass_yok.tg_girdiler["yass"].deger = ""
    yok_degerleri = TasimaIslemleri(yass_yok).zemin_tasima_girdilerini_oku()

    assert yok_degerleri["yass_var"] is False
    assert yok_degerleri["yass"] == pytest.approx(
        yok_degerleri["Df"] + yok_degerleri["B"]
    )

    yass_var = _sahte_zemin_uygulamasi()
    yass_var.tasima_yass_var = _SahteDeger(True)
    yass_var.tg_girdiler["yass"].deger = "0.75"
    var_degerleri = TasimaIslemleri(yass_var).zemin_tasima_girdilerini_oku()

    assert var_degerleri["yass_var"] is True
    assert var_degerleri["yass"] == pytest.approx(0.75)


def test_tasima_gucu_duzeltme_katsayilarini_uygular():
    analiz = TBDY2018TasimaGucu(10, 30, 18.0, 20.0, 999.0)
    varsayilan_qk, _ = analiz.analiz_yap(1.5, 2.0, 1.0)
    azaltilmis_qk, _ = analiz.analiz_yap(
        1.5,
        2.0,
        1.0,
        duzeltme_katsayilari={"ic": 0.8, "iq": 0.8, "igamma": 0.8},
    )

    assert azaltilmis_qk < varsayilan_qk


def test_tasima_tasarim_dayanimi_gamma_rv_ile_azaltilir():
    analiz = TBDY2018TasimaGucu(10, 30, 18.0, 20.0, 999.0)
    qk, qt = analiz.analiz_yap(1.5, 2.0, 1.0, gamma_Rv=1.4)

    assert qt == pytest.approx(qk / 1.4)


def test_phi_sifira_yaklasirken_sonuc_sonlu_ve_pozitiftir():
    sifir = TBDY2018TasimaGucu(10, 0, 18.0, 20.0, 999.0)
    cok_kucuk = TBDY2018TasimaGucu(10, 1e-16, 18.0, 20.0, 999.0)

    sifir_qk, _ = sifir.analiz_yap(1.5, 2.0, 1.0)
    kucuk_qk, _ = cok_kucuk.analiz_yap(1.5, 2.0, 1.0)

    assert math.isfinite(kucuk_qk)
    assert kucuk_qk > 0
    assert kucuk_qk == pytest.approx(sifir_qk)


def test_duzeltme_katsayisi_birden_buyuk_olamaz():
    analiz = TBDY2018TasimaGucu(10, 30, 18.0, 20.0, 999.0)
    with pytest.raises(ValueError, match="0 ile 1"):
        analiz.analiz_yap(1.5, 2.0, 1.0, duzeltme_katsayilari={"iq": 1.01})


def test_qt_rapor_degeri_yukariya_yuvarlanmaz():
    assert tasima_qt_asagi_yuvarla(94.519) == pytest.approx(94.51)
    assert tasima_qt_asagi_yuvarla(94.51000000000001) == pytest.approx(94.51)


def test_su_etkisi_taban_ustu_ici_ve_altinda_tutarlidir():
    su_yuzeyde = TBDY2018TasimaGucu(10, 30, 18.0, 20.0, 0.0)
    su_taban_ile_b_arasinda = TBDY2018TasimaGucu(10, 30, 18.0, 20.0, 2.0)
    su_etkisiz_derinde = TBDY2018TasimaGucu(10, 30, 18.0, 20.0, 3.0)

    assert su_yuzeyde._su_etkisi_duzeltmesi(1.0, 2.0) == pytest.approx((10.19, 10.19))
    assert su_taban_ile_b_arasinda._su_etkisi_duzeltmesi(1.0, 2.0) == pytest.approx((14.095, 18.0))
    assert su_etkisiz_derinde._su_etkisi_duzeltmesi(1.0, 2.0) == pytest.approx((18.0, 18.0))


@pytest.mark.parametrize(
    "b,l,df",
    [
        (2.0, 1.0, 1.0),
        (0.0, 1.0, 1.0),
        (1.0, 1.0, -0.1),
        (math.nan, 1.0, 1.0),
    ],
)
def test_tasima_gucu_gecersiz_geometriyi_reddeder(b, l, df):
    analiz = TBDY2018TasimaGucu(10, 30, 18.0, 20.0, 999.0)
    with pytest.raises(ValueError):
        analiz.analiz_yap(b, l, df)


def test_tasima_gucu_gecersiz_malzeme_girdilerini_reddeder():
    with pytest.raises(ValueError):
        TBDY2018TasimaGucu(0, 30, 18.0, 9.81, 1.0)
    with pytest.raises(ValueError):
        TBDY2018TasimaGucu(0, "gecersiz", 18.0, 20.0, 1.0)


def test_tasima_raporu_girdi_degistiginde_eskimis_sayilir():
    class SahteUygulama:
        pass

    app = SahteUygulama()
    app.tg_girdiler = {"ks_carpani": _SahteDeger("56"), "c": _SahteDeger("5")}
    app.entry_qt_nihai = _SahteDeger("20")
    app.tasima_varsayim_onayi = _SahteDeger(True)
    app.zemin_kaya_var = _SahteDeger("zemin")
    app.txt_tasima_rapor = _SahteDeger("geçerli rapor")
    app.tasima_rapor_imzasi = None
    islemler = TasimaIslemleri(app)

    app.tasima_rapor_imzasi = islemler.tasima_girdi_imzasi_olustur()
    assert islemler.tasima_raporu_guncel_mi()

    app.tg_girdiler["c"].deger = "6"
    assert not islemler.tasima_raporu_guncel_mi()


def test_zemin_hesabi_yeterlilik_hukmu_vermeden_metin_uretir():
    app = _sahte_zemin_uygulamasi()
    islemler = TasimaIslemleri(app)

    islemler.tasima_hesapla()
    islemler.tasima_metni_olustur()

    assert "q_t =" in app.txt_tasima_rapor.deger
    assert "k_s = 40 × q_t × G_k" in app.txt_tasima_rapor.deger
    assert "yeterliliği hakkında hüküm üretilmemiştir" in app.txt_tasima_rapor.deger
    assert "taşıma gücü sorunu yoktur" not in app.txt_tasima_rapor.deger
    assert "Hesapta c =" not in app.txt_tasima_rapor.deger
    assert "N_c =" not in app.txt_tasima_rapor.deger
    assert float(app.entry_ks_nihai.deger) == pytest.approx(
        float(app.entry_qt_nihai.deger) * 40.0 * 1.4,
        abs=0.005,
    )
    assert islemler.tasima_raporu_guncel_mi()

    app.entry_ks_nihai.deger = "9999"
    assert not islemler.tasima_raporu_guncel_mi()


def test_dayanim_parametreleri_23_oranina_indirilir_ve_tek_cumle_kullanilir():
    app = _sahte_zemin_uygulamasi()
    app.tasima_dayanim_23_uygulandi = _SahteDeger(False)
    app.tasima_dayanim_23_kaynak_c = ""
    app.tasima_dayanim_23_kaynak_phi = ""
    app.btn_tasima_dayanim_23 = _SahteDeger("")
    app.lbl_tasima_dayanim_23 = _SahteDeger("")
    islemler = TasimaIslemleri(app)

    islemler.tasima_dayanim_23_degistir()

    assert float(app.tg_girdiler["c"].deger) == pytest.approx(10 * 2 / 3)
    assert float(app.tg_girdiler["phi"].deger) == pytest.approx(20.0)
    assert app.tasima_dayanim_23_uygulandi.get() is True

    app.tg_girdiler["c"].deger = "7"
    app.tg_girdiler["phi"].deger = "21"
    islemler.tasima_hesapla()
    islemler.tasima_metni_olustur()

    beklenen = (
        "Taşıma gücü hesabında kullanılan efektif dayanım parametreleri, zemin "
        "koşullarındaki belirsizlikler dikkate alınarak güvenli tarafta kalınması "
        "amacıyla başlangıç değerlerinin 2/3’üne düşürülmüş; hesaplarda "
        "c′=0.71 t/m² ve φ′=21.00° değerleri esas alınmıştır."
    )
    assert beklenen in app.txt_tasima_rapor.deger

    islemler.tasima_dayanim_23_degistir()
    assert app.tg_girdiler["c"].deger == "10"
    assert app.tg_girdiler["phi"].deger == "30"
    assert app.tasima_dayanim_23_uygulandi.get() is False


def test_tasima_imzasi_yalniz_aktif_dalin_girdilerini_izler():
    app = _sahte_zemin_uygulamasi()
    islemler = TasimaIslemleri(app)
    islemler.tasima_hesapla()
    islemler.tasima_metni_olustur()
    assert islemler.tasima_raporu_guncel_mi()

    app.tg_girdiler["qt"].deger = "999"
    app.tg_girdiler["Gk"].deger = "9"
    assert islemler.tasima_raporu_guncel_mi()

    app.tg_girdiler["c"].deger = "11"
    assert not islemler.tasima_raporu_guncel_mi()


def test_kaya_hesabi_uygunluk_hukmu_vermeden_rapor_uretir():
    app = _sahte_zemin_uygulamasi()
    app.zemin_kaya_var.deger = "kaya"
    app.tasima_varsayim_onayi.deger = False
    islemler = TasimaIslemleri(app)

    islemler.tasima_hesapla()
    islemler.tasima_metni_olustur()

    assert "Kaya Birimi" in app.txt_tasima_rapor.deger
    assert "yeterliliği hakkında hüküm üretilmemiştir" in app.txt_tasima_rapor.deger
    assert islemler.tasima_raporu_guncel_mi()


def test_hesaplanandan_buyuk_nihai_qt_eski_metni_birakmaz(monkeypatch):
    app = _sahte_zemin_uygulamasi()
    islemler = TasimaIslemleri(app)
    uyarilar = []
    monkeypatch.setattr(
        "tasima_islemleri.messagebox.showwarning",
        lambda *args, **kwargs: uyarilar.append(args),
    )

    islemler.tasima_hesapla()
    app.txt_tasima_rapor.deger = "ESKİ METİN"
    app.entry_qt_nihai.deger = "10000"
    islemler.tasima_metni_olustur()

    assert app.txt_tasima_rapor.deger == ""
    assert uyarilar
    assert not islemler.tasima_raporu_guncel_mi()
