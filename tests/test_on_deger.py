from is_takibi import IsTakibiDeposu
from on_deger import (
    bos_is_akisi_verisi,
    bos_on_deger_verisi,
    bos_tdth_verisi,
    is_durumu_degistir,
    on_deger_durumu,
    on_deger_revizyonu_ekle,
    tdth_kaydi_etkinlestir,
    tdth_zemin_sinifi_guncelle,
)


def _tdth_kaydi(zemin="ZD", ozet="abc"):
    return {
        "id": "tdth-1",
        "pdf_yolu": "rapor.pdf",
        "orijinal_dosya_adi": "rapor.pdf",
        "sha256": ozet,
        "sayfa_sayisi": 6,
        "rapor_basligi": "463 Ada 73 Parsel",
        "dd_duzeyi": "DD-2",
        "zemin_sinifi": zemin,
        "enlem": "40.1",
        "boylam": "26.4",
        "degerler": {"PGA": "0.303"},
        "ice_aktarim_tarihi": "2026-08-06T10:00:00+03:00",
    }


def test_on_deger_ilk_kaydi_korur_ve_revizyon_ekler():
    veri, ilk = on_deger_revizyonu_ekle(bos_on_deger_verisi(), "12,5", "700", "ZD", "ilk", "abc")
    veri, ikinci = on_deger_revizyonu_ekle(veri, "15", "850", "ZC", "hesap sonrası", "def")

    assert veri["ilk"]["id"] == ilk["id"]
    assert veri["ilk"]["qt"] == "12,5"
    assert veri["guncel"]["qt"] == "15"
    assert len(veri["revizyonlar"]) == 2
    assert ikinci["tdth_hash"] == "def"
    assert "qo" not in veri["guncel"]


def test_zemin_sinifi_degisiminde_tdth_yenilenmeli_olur():
    tdth = tdth_kaydi_etkinlestir(bos_tdth_verisi(), _tdth_kaydi("ZD"))
    tdth = tdth_zemin_sinifi_guncelle(tdth, "ZC")

    assert tdth["durum"] == "yenilenmeli"
    assert tdth["aktif"]["zemin_sinifi"] == "ZD"


def test_yeni_is_on_deger_olmadan_dogrudan_yazima_gecebilir():
    akis = is_durumu_degistir(
        bos_is_akisi_verisi(),
        "yazim_asamasinda",
        "Ön değer verilmeden doğrudan yazım başlatıldı.",
    )

    assert akis["durum"] == "yazim_asamasinda"
    assert akis["gecmis"][-1]["eski"] == "yeni"


def test_yazim_sirasinda_on_deger_eklemek_is_asamasini_geriye_almaz():
    akis = is_durumu_degistir(bos_is_akisi_verisi(), "yazim_asamasinda")
    on_deger = bos_on_deger_verisi()
    assert on_deger_durumu(on_deger) == "verilmedi"

    on_deger, _ = on_deger_revizyonu_ekle(on_deger, "18", "900", "ZD", "sonradan verildi", "abc")

    assert on_deger_durumu(on_deger) == "verildi"
    assert akis["durum"] == "yazim_asamasinda"


def test_bitmis_projeden_duzeltme_yeni_revizyon_baslatir():
    akis = bos_is_akisi_verisi()
    akis = is_durumu_degistir(akis, "on_deger_verildi")
    akis = is_durumu_degistir(akis, "yazim_asamasinda")
    akis = is_durumu_degistir(akis, "bitti")
    akis = is_durumu_degistir(akis, "duzeltme_asamasinda", "Parsel bilgisi düzeltilecek")

    assert akis["durum"] == "duzeltme_asamasinda"
    assert akis["revizyon_no"] == 2
    assert akis["duzeltme_nedeni"] == "Parsel bilgisi düzeltilecek"
    assert akis["tamamlanma_tarihi"] == ""


def test_is_takibi_indeksi_proje_ozetini_saklar(tmp_path):
    depo = IsTakibiDeposu(tmp_path / "is_takibi.db")
    akis = bos_is_akisi_verisi()
    on_deger, _ = on_deger_revizyonu_ekle(bos_on_deger_verisi(), "20", "1000", "ZD", "", "abc")
    veri = {
        "PROJE_ADI": "Örnek Proje",
        "ILCE": "Ayvacık",
        "KOY": "Kozlu",
        "ADA": "151",
        "PARSEL": "10",
        "_IS_AKISI_": akis,
        "_ON_DEGER_": on_deger,
        "_TDTH_": tdth_kaydi_etkinlestir(bos_tdth_verisi(), _tdth_kaydi()),
    }
    depo.kaydet(tmp_path / "proje.json", veri)

    kayitlar = depo.listele("Kozlu")
    assert len(kayitlar) == 1
    assert kayitlar[0]["on_qt"] == "20"
    assert kayitlar[0]["zemin_sinifi"] == "ZD"
    assert kayitlar[0]["on_deger_durumu"] == "verildi"
