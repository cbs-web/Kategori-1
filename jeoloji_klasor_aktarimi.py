"""Proje klasöründeki jeoloji raporu ve parsel KML adaylarını eşleştirme araçları."""

from __future__ import annotations

from pathlib import Path
import os
import re

from harita_islemleri import kml_dosyalari_bul, kml_poligonlarini_oku
from jeoloji_kutuphanesi import jeoloji_anahtari
from jeoloji_kunye_uzlasma import metinden_kunye_adayi
from jeoloji_word_aktarimi import word_rapor_adaylarini_oku


def _sayisal_eslesme(metin, deger):
    deger = str(deger or "").strip()
    if not deger:
        return False
    return bool(re.search(rf"(?<!\d){re.escape(deger)}(?!\d)", metin))


def kml_adaylarini_sirala(kml_adaylari, word_sonucu):
    """KML adaylarını Word'deki ada/parsel ve dosya açıklamasına göre sıralar."""
    sonuc = []
    for aday in kml_adaylari:
        metin = " ".join(
            [
                Path(aday["dosya_yolu"]).stem,
                *(poligon.get("ad", "") for poligon in aday["poligonlar"]),
                *(poligon.get("aciklama", "") for poligon in aday["poligonlar"]),
            ]
        )
        anahtar = jeoloji_anahtari(metin)
        kml_kunye = metinden_kunye_adayi(metin)
        word_ada = str(getattr(word_sonucu, "ada", "") or "").strip()
        word_parsel = str(getattr(word_sonucu, "parsel", "") or "").strip()
        puan = 10
        if word_ada == "0" and not kml_kunye.get("ada"):
            puan += 25
        elif _sayisal_eslesme(metin, word_ada):
            puan += 25
        elif word_ada and kml_kunye.get("ada"):
            puan -= 40
        if _sayisal_eslesme(metin, word_parsel):
            puan += 30
        elif word_parsel and kml_kunye.get("parsel"):
            puan -= 45
        if "ada" in anahtar:
            puan += 5
        if "parsel" in anahtar or "tkgm" in anahtar:
            puan += 10
        if any(kelime in anahtar for kelime in ("sondaj", "jeofizik", "nokta", "serim")):
            puan -= 25
        puan += min(len(aday["poligonlar"]), 5)
        item = dict(aday)
        item["puan"] = puan
        sonuc.append(item)
    return sorted(
        sonuc,
        key=lambda item: (-item["puan"], os.path.normcase(item["dosya_yolu"]).casefold()),
    )


def proje_klasorunu_incele(klasor, task_context=None):
    """Bir proje klasöründeki geçerli Word ve KML adaylarını okur."""
    root = Path(klasor)
    if not root.is_dir():
        raise ValueError("Seçilen proje klasörü bulunamadı.")

    word_adaylari = word_rapor_adaylarini_oku([root], task_context=task_context)
    kml_adaylari = []
    kml_hatalari = []
    for path in kml_dosyalari_bul([root]):
        try:
            poligonlar = kml_poligonlarini_oku(path)
        except (OSError, ValueError) as exc:
            kml_hatalari.append({"dosya_yolu": path, "hata": str(exc)})
            continue
        kml_adaylari.append(
            {
                "dosya_yolu": path,
                "poligonlar": poligonlar,
                "poligon_sayisi": len(poligonlar),
            }
        )
    return {
        "klasor": str(root.resolve()),
        "word_adaylari": word_adaylari,
        "kml_adaylari": kml_adaylari,
        "kml_hatalari": kml_hatalari,
    }


__all__ = ["kml_adaylarini_sirala", "proje_klasorunu_incele"]
