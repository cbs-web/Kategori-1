import os


# Uygulama alanı -> (__HARITA__ JSON alanı, standart dosya adı)
RAPOR_HARITA_ALANLARI = {
    "img_mjh": ("mjh_yolu", "Mühendislik_Jeolojisi_Haritasi.jpg"),
    "img_jeofizik_lok": ("jeofizik_lokasyon_yolu", "Jeofizik_Lokasyon_Haritasi.jpg"),
    "img_jeoloji_lok": ("jeoloji_lokasyon_yolu", "Jeoloji_Lokasyon_Haritasi.jpg"),
    "img_yerbulduru": ("yerbulduru_yolu", "Yerbulduru_Haritasi.jpg"),
    "img_genel_jeoloji": ("genel_jeoloji_yolu", "Genel_Jeoloji_Haritasi.jpg"),
    "img_pga_haritasi": ("pga_haritasi_yolu", "PGA_Haritasi.jpg"),
}


def proje_klasorundeki_rapor_haritalarini_bul(proje_yolu):
    """Eski projelerde standart adlarla bulunan rapor haritalarını döndür."""
    if not proje_yolu:
        return {}
    proje_yolu = os.path.abspath(str(proje_yolu))
    klasor = proje_yolu if os.path.isdir(proje_yolu) else os.path.dirname(proje_yolu)
    if not os.path.isdir(klasor):
        return {}

    try:
        arama_klasorleri = [klasor]
        haritalar_alt_klasoru = os.path.join(klasor, "Haritalar")
        if os.path.isdir(haritalar_alt_klasoru):
            arama_klasorleri.append(haritalar_alt_klasoru)
        ad_eslemesi = {}
        for arama_klasoru in arama_klasorleri:
            for ad in os.listdir(arama_klasoru):
                yol = os.path.join(arama_klasoru, ad)
                if os.path.isfile(yol):
                    ad_eslemesi.setdefault(ad.casefold(), yol)
    except OSError:
        return {}

    bulunan = {}
    for uygulama_alani, (_, dosya_adi) in RAPOR_HARITA_ALANLARI.items():
        yol = ad_eslemesi.get(dosya_adi.casefold())
        if yol:
            bulunan[uygulama_alani] = os.path.abspath(yol)
    return bulunan


def harita_verisini_mevcut_dosyalarla_tamamla(harita_verisi, proje_yolu):
    """Kayıtlı yolu korur, eksik eski alanları proje klasöründen tamamlar."""
    sonuc = dict(harita_verisi) if isinstance(harita_verisi, dict) else {}
    bulunan = proje_klasorundeki_rapor_haritalarini_bul(proje_yolu)
    for uygulama_alani, (json_alani, _) in RAPOR_HARITA_ALANLARI.items():
        kayitli = str(sonuc.get(json_alani) or "").strip()
        if kayitli and os.path.isfile(kayitli):
            continue
        bulunan_yol = bulunan.get(uygulama_alani)
        if bulunan_yol:
            sonuc[json_alani] = bulunan_yol
    return sonuc
