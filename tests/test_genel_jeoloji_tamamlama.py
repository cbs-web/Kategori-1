from genel_jeoloji_islemleri import GenelJeolojiIslemleri


class SahteMetinKutuphanesi:
    def eski_jeoloji_wordlerinden_tamamla(self, birimler):
        return {"aktarilan": [], "bulunamayan": []}

    def getir(self, kod="", ad=""):
        if kod == "Tg":
            return {
                "id": 17,
                "revizyon_no": 3,
                "bolgesel_jeoloji_metni": "Kütüphaneden gelen güvenilir bölgesel jeoloji açıklaması.",
            }
        return None


class SahteUygulama:
    def __init__(self):
        self.genel_jeoloji_verisi = {
            "kaynak_modu": "kutuphane",
            "bolgesel_jeoloji_metni": "Kullanıcının elle düzenlediği birleşik 2.1 metni.",
            "birimler": [
                {
                    "kod": "Tg",
                    "ad": "Oligo-Miyosen Granitoyidleri",
                    "kullan": True,
                    "bolgesel_jeoloji_metni": "",
                }
            ],
        }
        self._proje_kirli = False
        self.hatalar = []

    def hata_kaydet(self, baslik, hata=None):
        self.hatalar.append((baslik, hata))


def test_eksik_birim_tamamlanirken_kullanicinin_birlesik_metni_korunur(monkeypatch):
    app = SahteUygulama()
    islemler = GenelJeolojiIslemleri(app)
    monkeypatch.setattr(
        GenelJeolojiIslemleri,
        "_metin_kutuphanesi",
        lambda self: SahteMetinKutuphanesi(),
    )

    sonuc = islemler.genel_jeoloji_eksik_metinlerini_tamamla()

    assert sonuc == {
        "tamamlanan": ["Oligo-Miyosen Granitoyidleri"],
        "bulunamayan": [],
    }
    assert app.genel_jeoloji_verisi["bolgesel_jeoloji_metni"] == (
        "Kullanıcının elle düzenlediği birleşik 2.1 metni."
    )
    birim = app.genel_jeoloji_verisi["birimler"][0]
    assert birim["bolgesel_jeoloji_metni"].startswith("Kütüphaneden gelen")
    assert birim["metin_kaynagi"] == "kalici_kutuphane"
    assert app._proje_kirli is True
