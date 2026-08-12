from pathlib import Path

from cizimler import CizimUretici
from harita_dosyalari import (
    RAPOR_HARITA_ALANLARI,
    harita_verisini_mevcut_dosyalarla_tamamla,
    proje_klasorundeki_rapor_haritalarini_bul,
)


class SahteUygulama:
    pass


def test_eski_proje_klasorundeki_haritalar_bulunur_eksik_olan_uydurulmaz(tmp_path):
    proje_yolu = tmp_path / "eski.json"
    proje_yolu.write_text("{}", encoding="utf-8")
    beklenen = {}
    for uygulama_alani, (_, dosya_adi) in RAPOR_HARITA_ALANLARI.items():
        if uygulama_alani == "img_yerbulduru":
            continue
        yol = tmp_path / dosya_adi
        yol.write_bytes(b"gorsel")
        beklenen[uygulama_alani] = str(yol)

    bulunan = proje_klasorundeki_rapor_haritalarini_bul(str(proje_yolu))
    tamamlanan = harita_verisini_mevcut_dosyalarla_tamamla(
        {"mjh_yolu": str(tmp_path / "artik-yok.jpg")},
        str(proje_yolu),
    )

    assert bulunan == beklenen
    assert tamamlanan["mjh_yolu"] == beklenen["img_mjh"]
    assert "yerbulduru_yolu" not in tamamlanan


def test_mevcut_haritalar_yeniden_uretilmeden_kullanilabilir(tmp_path, monkeypatch):
    yollar = []
    for _, dosya_adi in RAPOR_HARITA_ALANLARI.values():
        yol = tmp_path / dosya_adi
        yol.write_bytes(b"gorsel")
        yollar.append(str(yol))
    monkeypatch.setattr("cizimler.messagebox.askyesnocancel", lambda *args, **kwargs: True)

    secim = CizimUretici(SahteUygulama()).mevcut_harita_dosyalari_secimi(yollar)

    assert secim == "kullan"
