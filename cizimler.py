import os
import hashlib
import math
import re
import time
import tkinter as tk
from tkinter import filedialog, messagebox

from PIL import Image, ImageDraw, ImageGrab, ImageFont, ImageStat, ImageTk

from harita_renkleri import (
    CALISAN_PARSEL_SINIR_KALINLIGI,
    CALISAN_PARSEL_SINIR_RENGI,
    SS_SERIM_KALINLIGI,
    SS_SERIM_RENGI,
)
from tkgm_kml import TKGM_API_BASE

WINDOWS_AYGIT_ADLARI = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}


def guvenli_dosya_adi(ad, varsayilan="Kayit", azami_uzunluk=80):
    """Kullanıcı/proje kaynaklı bir etiketi tek, güvenli dosya adı bileşenine çevirir."""
    metin = re.sub(r"[<>:\"/\\|?*\x00-\x1f]", "_", str(ad or ""))
    metin = re.sub(r"\s+", " ", metin).strip(" .")
    if not metin:
        metin = varsayilan
    if metin.split(".", 1)[0].upper() in WINDOWS_AYGIT_ADLARI:
        metin = f"_{metin}"
    metin = metin[:azami_uzunluk].rstrip(" .")
    return metin or varsayilan


def derinlik_araligi_oku(deger):
    metin = str(deger or "").strip().replace(",", ".")
    eslesme = re.fullmatch(
        r"\s*(-?\d+(?:\.\d+)?)\s*[-–—]\s*(-?\d+(?:\.\d+)?)\s*",
        metin,
    )
    if not eslesme:
        return None
    ust, alt = float(eslesme.group(1)), float(eslesme.group(2))
    if not math.isfinite(ust) or not math.isfinite(alt) or ust < 0 or alt <= ust:
        return None
    return ust, alt


def _deterministik_rng(zemin_tipi, x1, y1, x2, y2):
    anahtar = f"{zemin_tipi}|{int(x1)}|{int(y1)}|{int(x2)}|{int(y2)}".encode("utf-8")
    tohum = int.from_bytes(hashlib.sha256(anahtar).digest()[:8], "big")
    import random
    return random.Random(tohum)


def parsel_noktalarini_dogrula(noktalar):
    sonuc = []
    for nokta in noktalar or []:
        if not isinstance(nokta, (list, tuple)) or len(nokta) < 2:
            raise ValueError("Parsel koordinatı [enlem, boylam] biçiminde olmalıdır.")
        try:
            enlem, boylam = float(nokta[0]), float(nokta[1])
        except (TypeError, ValueError):
            raise ValueError("Parsel koordinatları sayısal olmalıdır.") from None
        if not math.isfinite(enlem) or not math.isfinite(boylam):
            raise ValueError("Parsel koordinatları sonlu olmalıdır.")
        if not (-85.05112878 <= enlem <= 85.05112878 and -180 <= boylam <= 180):
            raise ValueError("Parsel koordinatı Web Mercator aralığının dışında.")
        sonuc.append((enlem, boylam))
    if len({(round(p[0], 10), round(p[1], 10)) for p in sonuc}) < 3:
        raise ValueError("Parsel haritası için en az üç farklı koordinat gerekir.")
    return sonuc


def _mercator_birim(enlem, boylam):
    enlem = max(-85.05112878, min(85.05112878, float(enlem)))
    x = (float(boylam) + 180.0) / 360.0
    y = (1.0 - math.asinh(math.tan(math.radians(enlem))) / math.pi) / 2.0
    return x, y


def _mercator_birimden_koordinat(x, y):
    boylam = (float(x) * 360.0) - 180.0
    enlem = math.degrees(math.atan(math.sinh(math.pi * (1.0 - 2.0 * float(y)))))
    return enlem, boylam


def parsel_gorunum_hesapla(noktalar, genislik, yukseklik, kenar_payi=0.18):
    """Parseli hedef piksel alanına sığdıran Web Mercator merkezini ve zoom'u üretir."""
    temiz = parsel_noktalarini_dogrula(noktalar)
    genislik = max(200, int(genislik))
    yukseklik = max(160, int(yukseklik))
    pay = min(0.35, max(0.08, float(kenar_payi)))
    mercator = [_mercator_birim(enlem, boylam) for enlem, boylam in temiz]
    xs = [p[0] for p in mercator]
    ys = [p[1] for p in mercator]
    span_x = max(xs) - min(xs)
    span_y = max(ys) - min(ys)
    kullanilabilir_w = genislik * (1.0 - 2.0 * pay)
    kullanilabilir_h = yukseklik * (1.0 - 2.0 * pay)
    zoom_x = math.log2(kullanilabilir_w / (256.0 * span_x)) if span_x > 0 else 22
    zoom_y = math.log2(kullanilabilir_h / (256.0 * span_y)) if span_y > 0 else 22
    zoom = max(1, min(21, int(math.floor(min(zoom_x, zoom_y)))))
    merkez = _mercator_birimden_koordinat(
        (min(xs) + max(xs)) / 2.0,
        (min(ys) + max(ys)) / 2.0,
    )
    return merkez, zoom


def parsel_noktalari_hashi(noktalar):
    temiz = parsel_noktalarini_dogrula(noktalar)
    metin = ";".join(f"{enlem:.8f},{boylam:.8f}" for enlem, boylam in temiz)
    return hashlib.sha256(metin.encode("ascii")).hexdigest()


class CizimUretici:
    def __init__(self, app):
        object.__setattr__(self, "app", app)

    def __getattr__(self, name):
        return getattr(self.app, name)

    def __setattr__(self, name, value):
        if name == "app":
            object.__setattr__(self, name, value)
        else:
            setattr(self.app, name, value)

    def proje_degeri(self, kod, varsayilan=""):
        alan = getattr(self, "veri_alanlari", {}).get(kod)
        if alan is None:
            return varsayilan
        try:
            return str(alan.get()).strip()
        except Exception:
            return varsayilan

    def proje_yass_degeri(self):
        try:
            for kayit in self.ac_yn_sekme_kayitlari():
                for satir in self.ac_yn_satirlari(kayit):
                    if len(satir) > 2:
                        deger = str(satir[2]).strip()
                        if deger and deger != "-":
                            return deger
        except Exception:
            pass
        return "-"

    def mevcut_dosyalar_icin_onay_al(self, yollar, baslik="Dosyaların Üzerine Yazılsın mı?"):
        mevcut = [yol for yol in yollar if os.path.exists(yol)]
        if not mevcut:
            return True
        adlar = "\n".join(f"- {os.path.basename(yol)}" for yol in mevcut[:12])
        if len(mevcut) > 12:
            adlar += f"\n- ... ve {len(mevcut) - 12} dosya daha"
        return messagebox.askyesno(
            baslik,
            "Aşağıdaki dosyalar zaten var ve yeniden üretimde üzerlerine yazılacak:\n\n"
            f"{adlar}\n\nDevam edilsin mi?",
        )

    def mevcut_harita_dosyalari_secimi(self, yollar, baslik="Mevcut Haritalar Bulundu"):
        """Haritaları kullanma, yeniden üretme veya işlemi iptal etme seçimini al."""
        mevcut = [yol for yol in yollar if os.path.isfile(yol)]
        if not mevcut:
            return "yeniden"

        adlar = "\n".join(f"- {os.path.basename(yol)}" for yol in mevcut)
        if len(mevcut) == len(yollar):
            cevap = messagebox.askyesnocancel(
                baslik,
                "Rapor için gerekli harita dosyaları zaten var:\n\n"
                f"{adlar}\n\n"
                "Evet: Mevcut dosyaları kullan\n"
                "Hayır: Haritaları yeniden oluştur\n"
                "İptal: İşlemden vazgeç",
            )
            if cevap is True:
                return "kullan"
            if cevap is False:
                return "yeniden"
            return "iptal"

        cevap = messagebox.askyesno(
            baslik,
            "Harita takımının yalnız bir bölümü mevcut:\n\n"
            f"{adlar}\n\nEksik dosyalarla birlikte tüm takım yeniden oluşturulsun mu?",
        )
        return "yeniden" if cevap else "iptal"

    def harita_kirpma_miktarlari(self, genislik, yukseklik):
        if genislik < 500 or yukseklik < 250:
            return 0, 0, 0, 0
        sol = min(160, max(0, int(genislik * 0.14)))
        ust = min(65, max(0, int(yukseklik * 0.10)))
        sag = min(10, max(0, int(genislik * 0.02)))
        alt = min(12, max(0, int(yukseklik * 0.02)))
        if genislik - sol - sag < 300:
            sol = sag = 0
        if yukseklik - ust - alt < 200:
            ust = alt = 0
        return sol, ust, sag, alt

    def harita_yakalama_widgeti(self):
        canvas = getattr(self.map_widget, "canvas", None)
        if canvas and canvas.winfo_exists():
            return canvas
        return self.map_widget

    def harita_goruntusu_yakala(self, ortali_kirp=False):
        widget = self.harita_yakalama_widgeti()
        widget.update_idletasks()
        self.root.update()
        time.sleep(0.1)

        genislik = max(1, widget.winfo_width())
        yukseklik = max(1, widget.winfo_height())
        sol, ust, sag, alt = self.harita_kirpma_miktarlari(genislik, yukseklik)
        if ortali_kirp:
            sag = sol
            alt = ust

        def kirp(img):
            if sol or ust or sag or alt:
                sag_kenar = max(sol + 1, img.width - sag)
                alt_kenar = max(ust + 1, img.height - alt)
                return img.crop((sol, ust, sag_kenar, alt_kenar))
            return img

        try:
            img = ImageGrab.grab(window=widget.winfo_id())
            if img.width > 50 and img.height > 50:
                return kirp(img)
        except Exception as e:
            self.hata_kaydet("Harita canvas penceresi doğrudan yakalanamadı, bbox yöntemine geçiliyor", e)

        x = widget.winfo_rootx() + sol
        y = widget.winfo_rooty() + ust
        w = max(1, genislik - sol - sag)
        h = max(1, yukseklik - ust - alt)
        return ImageGrab.grab(bbox=(x, y, x + w, y + h), all_screens=True)

    def yerbulduru_merkez_koordinati(self):
        points = getattr(self, "yuklu_kml_points", []) or []
        try:
            if points:
                latlar = [float(p[0]) for p in points]
                lonlar = [float(p[1]) for p in points]
                return ((min(latlar) + max(latlar)) / 2, (min(lonlar) + max(lonlar)) / 2)
        except Exception as e:
            self.hata_kaydet("Yerbulduru KML merkez koordinatı hesaplanamadı", e)

        for data in getattr(self, "harita_isaretleri", {}).values():
            if data.get("tip") in ("AÇ", "YN", "M"):
                try:
                    return (float(data["lat"]), float(data["lon"]))
                except Exception:
                    continue

        try:
            enlem = float(self.proje_degeri("ENLEM").replace(",", "."))
            boylam = float(self.proje_degeri("BOYLAM").replace(",", "."))
            if math.isfinite(enlem) and math.isfinite(boylam) and -90 <= enlem <= 90 and -180 <= boylam <= 180:
                return enlem, boylam
        except (TypeError, ValueError):
            pass

        try:
            konum = self.map_widget.get_position()
            if konum and len(konum) >= 2:
                enlem, boylam = float(konum[0]), float(konum[1])
                if math.isfinite(enlem) and math.isfinite(boylam):
                    return enlem, boylam
        except Exception:
            pass
        raise ValueError("Yerbulduru haritası için geçerli proje veya harita koordinatı bulunamadı.")

    def yerbulduru_kml_kuzey_koordinati(self):
        points = getattr(self, "yuklu_kml_points", []) or []
        try:
            if points:
                temiz_points = [(float(p[0]), float(p[1])) for p in points]
                return max(temiz_points, key=lambda p: p[0])
        except Exception as e:
            self.hata_kaydet("Yerbulduru KML kuzey noktası hesaplanamadı", e)
        return None

    def yerbulduru_isaret_koordinati(self):
        return self.yerbulduru_kml_kuzey_koordinati() or self.yerbulduru_merkez_koordinati()

    def yerbulduru_goruntu_bos_mu(self, img):
        ornek = img.convert("RGB").resize((64, 64), Image.Resampling.BILINEAR)
        veri = ornek.tobytes()
        toplam = max(1, len(veri) // 3)
        beyaz = 0
        for i in range(0, len(veri), 3):
            if veri[i] > 240 and veri[i + 1] > 240 and veri[i + 2] > 240:
                beyaz += 1
        gri = ornek.convert("L")
        ortalama = ImageStat.Stat(gri).mean[0]
        beyaz_orani = beyaz / toplam

        bloklu = img.convert("RGB").resize((96, 64), Image.Resampling.BILINEAR)
        eksik_blok = 0
        for satir in range(8):
            for sutun in range(12):
                kutu = (
                    sutun * bloklu.width // 12,
                    satir * bloklu.height // 8,
                    (sutun + 1) * bloklu.width // 12,
                    (satir + 1) * bloklu.height // 8,
                )
                blok = bloklu.crop(kutu)
                blok_veri = blok.tobytes()
                blok_toplam = max(1, len(blok_veri) // 3)
                blok_beyaz = 0
                for i in range(0, len(blok_veri), 3):
                    if blok_veri[i] > 240 and blok_veri[i + 1] > 240 and blok_veri[i + 2] > 240:
                        blok_beyaz += 1
                if (blok_beyaz / blok_toplam) > 0.65:
                    eksik_blok += 1

        return beyaz_orani > 0.15 or (ortalama > 235 and beyaz_orani > 0.08) or eksik_blok > 0

    def yerbulduru_yuklenmeyen_karo_sayisi(self):
        bekleyen = 0
        toplam = 0
        not_loaded = getattr(self.map_widget, "not_loaded_tile_image", None)
        empty = getattr(self.map_widget, "empty_tile_image", None)
        for kolon in getattr(self.map_widget, "canvas_tile_array", []) or []:
            for tile in kolon:
                toplam += 1
                image = getattr(tile, "image", None)
                canvas_object = getattr(tile, "canvas_object", None)
                if image == not_loaded or image == empty or canvas_object is None:
                    bekleyen += 1
        return bekleyen, toplam

    def yerbulduru_durum_yaz(self, mesaj):
        try:
            if hasattr(self.app, "durum_mesaji_yaz"):
                self.durum_mesaji_yaz(mesaj)
        except Exception:
            pass

    def yerbulduru_kml_sinirini_goster(self):
        try:
            if getattr(self, "kml_polygon_obj", None):
                self.kml_polygon_obj.draw()
        except Exception as e:
            self.hata_kaydet("Yerbulduru KML sınırı yeniden çizilemedi", e)

    def yerbulduru_harita_goruntusu_al(
        self, merkez, zoom, kml_sinir_goster=False, zaman_asimi=20.0, ortali_kirp=False
    ):
        self.map_widget.set_position(merkez[0], merkez[1])
        self.map_widget.set_zoom(zoom)
        son_img = None
        baslangic = time.monotonic()
        while time.monotonic() - baslangic < zaman_asimi:
            self.map_widget.update_idletasks()
            self.root.update()
            if kml_sinir_goster:
                self.yerbulduru_kml_sinirini_goster()
            bekleyen, toplam = self.yerbulduru_yuklenmeyen_karo_sayisi()
            if toplam:
                self.yerbulduru_durum_yaz(f"Yerbulduru karoları yükleniyor... {max(0, toplam - bekleyen)}/{toplam}")
            if toplam and bekleyen == 0:
                time.sleep(0.15)
                self.map_widget.update_idletasks()
                self.root.update()
                if kml_sinir_goster:
                    self.yerbulduru_kml_sinirini_goster()
                son_img = self.harita_goruntusu_yakala(ortali_kirp=ortali_kirp)
                if not self.yerbulduru_goruntu_bos_mu(son_img):
                    self.yerbulduru_durum_yaz("Yerbulduru harita görüntüsü hazır.")
                    return son_img
            time.sleep(0.08)

        son_img = self.harita_goruntusu_yakala(ortali_kirp=ortali_kirp)
        return son_img

    def yerbulduru_genis_gorunum_ayarlari(self, merkez=None):
        return merkez or self.yerbulduru_merkez_koordinati(), 9

    def yerbulduru_mercator_pixel(self, lat, lon, zoom):
        lat = max(-85.05112878, min(85.05112878, float(lat)))
        lon = float(lon)
        sin_lat = math.sin(math.radians(lat))
        dunya_boyutu = 256 * (2 ** int(zoom))
        x = (lon + 180.0) / 360.0 * dunya_boyutu
        y = (0.5 - math.log((1 + sin_lat) / (1 - sin_lat)) / (4 * math.pi)) * dunya_boyutu
        return x, y

    def yerbulduru_ekran_noktasi_hesapla(self, koordinat, merkez, zoom):
        widget = self.harita_yakalama_widgeti()
        genislik = max(1, widget.winfo_width())
        yukseklik = max(1, widget.winfo_height())
        sol, ust, _, _ = self.harita_kirpma_miktarlari(genislik, yukseklik)

        merkez_x, merkez_y = self.yerbulduru_mercator_pixel(merkez[0], merkez[1], zoom)
        nokta_x, nokta_y = self.yerbulduru_mercator_pixel(koordinat[0], koordinat[1], zoom)
        return (
            (genislik / 2) + (nokta_x - merkez_x) - sol,
            (yukseklik / 2) + (nokta_y - merkez_y) - ust,
        )

    def yerbulduru_panel_hazirla(self, img, hedef_boyut, nokta=None):
        hedef_w, hedef_h = hedef_boyut
        img = img.convert("RGB")
        oran = max(hedef_w / img.width, hedef_h / img.height)
        yeni_boyut = (max(1, int(img.width * oran)), max(1, int(img.height * oran)))
        img = img.resize(yeni_boyut, Image.Resampling.LANCZOS)
        sol = max(0, (img.width - hedef_w) // 2)
        ust = max(0, (img.height - hedef_h) // 2)
        panel = img.crop((sol, ust, sol + hedef_w, ust + hedef_h))

        panel_noktasi = None
        if nokta:
            x = int((nokta[0] * oran) - sol)
            y = int((nokta[1] * oran) - ust)
            panel_noktasi = (
                max(35, min(hedef_w - 35, x)),
                max(105, min(hedef_h - 20, y)),
            )

        return panel, panel_noktasi

    def yerbulduru_resmi_panele_uydur(self, img, hedef_boyut):
        return self.yerbulduru_panel_hazirla(img, hedef_boyut)[0]

    def yerbulduru_fontlari(self):
        try:
            return (
                ImageFont.truetype("arialbd.ttf", 46),
                ImageFont.truetype("arialbd.ttf", 38),
                ImageFont.truetype("arial.ttf", 24),
            )
        except Exception:
            font = ImageFont.load_default()
            return font, font, font

    def yerbulduru_etiket_yaz(self, draw, konum, metin, font, fill):
        x, y = konum
        golge = "black" if str(fill).lower() == "white" else "white"
        for dx, dy in [(-2, -2), (-2, 2), (2, -2), (2, 2), (0, -2), (0, 2), (-2, 0), (2, 0)]:
            draw.text((x + dx, y + dy), metin, font=font, fill=golge)
        draw.text((x, y), metin, font=font, fill=fill)

    def yerbulduru_pin_ciz(self, draw, x, y, label, label_renk="red", uc_noktasi=False):
        if uc_noktasi:
            y -= 38
        govde_renk = "#ff1f1f"
        dis_renk = "white"
        r = 17
        draw.ellipse((x - r, y - r, x + r, y + r), fill=govde_renk, outline=dis_renk, width=5)
        draw.ellipse((x - 6, y - 6, x + 6, y + 6), fill=dis_renk)
        draw.polygon([(x - 9, y + 11), (x + 9, y + 11), (x, y + 38)], fill=govde_renk, outline=dis_renk)
        _, font_label, _ = self.yerbulduru_fontlari()
        try:
            bbox = draw.textbbox((0, 0), label, font=font_label)
            label_w = bbox[2] - bbox[0]
        except Exception:
            label_w = len(label) * 18
        label_x = max(18, x - label_w // 2)
        self.yerbulduru_etiket_yaz(draw, (label_x, y - 58), label, font_label, label_renk)

    def yerbulduru_kuzey_oku_ciz(self, draw, x, y, size=82):
        """İki panelli yerbulduru çıktısında yön bilgisini görünür tut."""
        half = size // 2
        draw.rectangle(
            (x - half - 12, y - half - 18, x + half + 12, y + half + 24),
            fill="white",
            outline="black",
            width=3,
        )
        draw.line((x, y + half - 4, x, y - half + 12), fill="black", width=5)
        draw.polygon(
            ((x, y - half), (x - 14, y - half + 28), (x + 14, y - half + 28)),
            fill="#c51f2d",
            outline="black",
        )
        _, _, font_small = self.yerbulduru_fontlari()
        draw.text((x - 8, y + half - 3), "K", fill="black", font=font_small)

    def yerbulduru_iki_panel_resim_olustur(self, genis_img, yakin_img, kayit_yolu, genis_pin=None, yakin_pin=None):
        # A4 portre oranına yakın çıktı üret; Word'de sayfanın kullanılabilir
        # alanına büyütüldüğünde harita küçük kalmasın. Paneller kendi oranları
        # bozulmadan merkezden kırpılır.
        canvas_w = 2000
        margin_x = 70
        margin_y = 50
        header_h = 95
        footer_h = 60
        panel_gap = 45
        panel_w = canvas_w - (2 * margin_x)
        panel_h = 1270
        top_y = margin_y + header_h
        bottom_y = top_y + panel_h + panel_gap
        canvas_h = bottom_y + panel_h + margin_y + footer_h

        canvas = Image.new("RGB", (canvas_w, canvas_h), "white")
        draw = ImageDraw.Draw(canvas)
        font_title, font_panel, font_small = self.yerbulduru_fontlari()
        draw.text(
            (margin_x, margin_y // 2),
            "YERBULDURU HARİTASI",
            fill="#111111",
            font=font_title,
        )
        top_panel, top_pin = self.yerbulduru_panel_hazirla(genis_img, (panel_w, panel_h), genis_pin)
        bottom_panel, bottom_pin = self.yerbulduru_panel_hazirla(yakin_img, (panel_w, panel_h), yakin_pin)

        canvas.paste(top_panel, (margin_x, top_y))
        canvas.paste(bottom_panel, (margin_x, bottom_y))
        draw.rectangle([margin_x, top_y, margin_x + panel_w, top_y + panel_h], outline="black", width=4)
        draw.rectangle([margin_x, bottom_y, margin_x + panel_w, bottom_y + panel_h], outline="black", width=4)

        label_w = 300
        draw.rectangle(
            (margin_x + 18, top_y + 18, margin_x + label_w, top_y + 64),
            fill="white",
            outline="black",
            width=2,
        )
        draw.text((margin_x + 30, top_y + 23), "GENİŞ KONUM", fill="black", font=font_panel)
        draw.rectangle(
            (margin_x + 18, bottom_y + 18, margin_x + label_w, bottom_y + 64),
            fill="white",
            outline="black",
            width=2,
        )
        draw.text((margin_x + 30, bottom_y + 23), "YAKIN ÇALIŞMA ALANI", fill="black", font=font_panel)
        self.yerbulduru_kuzey_oku_ciz(
            draw, margin_x + panel_w - 95, top_y + 82, size=72
        )
        self.yerbulduru_kuzey_oku_ciz(
            draw, margin_x + panel_w - 95, bottom_y + 82, size=72
        )

        if top_pin is None:
            top_pin = (panel_w // 2, panel_h // 2)
        if bottom_pin is None:
            bottom_pin = (panel_w // 2, panel_h // 2)

        top_pin_x = margin_x + top_pin[0]
        top_pin_y = top_y + top_pin[1]
        bottom_pin_x = margin_x + bottom_pin[0]
        bottom_pin_y = bottom_y + bottom_pin[1]
        draw.line([(top_pin_x, top_pin_y), (bottom_pin_x, bottom_pin_y)], fill="black", width=4)

        self.yerbulduru_pin_ciz(draw, top_pin_x, top_pin_y, "ÇALIŞMA ALANI", label_renk="white", uc_noktasi=True)
        self.yerbulduru_pin_ciz(draw, bottom_pin_x, bottom_pin_y, "ÇALIŞMA ALANI", label_renk="red", uc_noktasi=True)
        footer_y = bottom_y + panel_h + 18
        draw.line(
            (margin_x, footer_y - 8, canvas_w - margin_x, footer_y - 8),
            fill="#555555",
            width=2,
        )
        draw.text(
            (margin_x, footer_y),
            "Kaynak: Google Uydu altlığı",
            fill="#222222",
            font=font_small,
        )
        canvas.save(kayit_yolu, quality=95)

    def parsel_geometri_hashi(self, noktalar=None):
        return parsel_noktalari_hashi(
            noktalar if noktalar is not None else getattr(self, "yuklu_kml_points", [])
        )

    def parsel_haritasi_etiketi(self):
        ada = self.proje_degeri("ADA", "").strip()
        parsel = self.proje_degeri("PARSEL", "").strip()
        if parsel and ada and ada != "0":
            return f"{ada} / {parsel}"
        return parsel or ada or "PARSEL"

    def parsel_haritasi_kaynak_metni(self):
        path = str(getattr(self, "yuklu_kml_yolu", "") or "")
        kaynak_url = str(getattr(self, "parsel_haritasi_kaynak_url", "") or "").strip()
        if not kaynak_url.startswith(("http://", "https://")):
            if "tkgm" in os.path.basename(path).casefold():
                kaynak_url = TKGM_API_BASE
        kaynak = kaynak_url if kaynak_url else "Yüklenen parsel KML geometrisi"
        return f"Kaynak: {kaynak} · Altlık: Google Uydu"

    @staticmethod
    def _poligon_etiket_noktasi(points):
        if len(points) < 3:
            return points[0] if points else (0, 0)
        alan = 0.0
        cx = 0.0
        cy = 0.0
        for index, (x1, y1) in enumerate(points):
            x2, y2 = points[(index + 1) % len(points)]
            carpi = (x1 * y2) - (x2 * y1)
            alan += carpi
            cx += (x1 + x2) * carpi
            cy += (y1 + y2) * carpi
        if abs(alan) < 1e-9:
            return (
                sum(point[0] for point in points) / len(points),
                sum(point[1] for point in points) / len(points),
            )
        return cx / (3.0 * alan), cy / (3.0 * alan)

    def parsel_haritasi_resmi_olustur(self, map_img, noktalar, merkez, zoom):
        temiz = parsel_noktalarini_dogrula(noktalar)
        image = map_img.convert("RGBA")
        ekran_points = [
            self.yerbulduru_ekran_noktasi_hesapla(point, merkez, zoom)
            for point in temiz
        ]
        overlay = Image.new("RGBA", image.size, (0, 0, 0, 0))
        draw = ImageDraw.Draw(overlay)
        draw.polygon(ekran_points, fill=(190, 20, 45, 78))
        kapali = ekran_points + [ekran_points[0]]
        draw.line(kapali, fill=(145, 0, 35, 255), width=6, joint="curve")

        label = self.parsel_haritasi_etiketi()
        label_x, label_y = self._poligon_etiket_noktasi(ekran_points)
        font_size = max(18, min(34, int(min(image.size) * 0.045)))
        try:
            font = ImageFont.truetype("arialbd.ttf", font_size)
        except Exception:
            font = ImageFont.load_default()
        try:
            bbox = draw.textbbox((0, 0), label, font=font, stroke_width=3)
            text_w, text_h = bbox[2] - bbox[0], bbox[3] - bbox[1]
        except Exception:
            text_w, text_h = len(label) * font_size * 0.6, font_size
        draw.text(
            (label_x - text_w / 2, label_y - text_h / 2),
            label,
            font=font,
            fill="white",
            stroke_width=3,
            stroke_fill="black",
        )
        image = Image.alpha_composite(image, overlay).convert("RGB")

        target_width = 1200
        scale = target_width / float(image.width)
        target_height = max(1, int(round(image.height * scale)))
        image = image.resize((target_width, target_height), Image.Resampling.LANCZOS)
        footer_height = 70
        result = Image.new("RGB", (target_width, target_height + footer_height), "white")
        result.paste(image, (0, 0))
        footer = ImageDraw.Draw(result)
        kaynak_metni = self.parsel_haritasi_kaynak_metni()
        try:
            footer_punto = 22
            footer_font = ImageFont.truetype("arial.ttf", footer_punto)
            while footer_punto > 14:
                bbox = footer.textbbox((0, 0), kaynak_metni, font=footer_font)
                if (bbox[2] - bbox[0]) <= target_width - 44:
                    break
                footer_punto -= 1
                footer_font = ImageFont.truetype("arial.ttf", footer_punto)
        except Exception:
            footer_font = ImageFont.load_default()
        try:
            bbox = footer.textbbox((0, 0), kaynak_metni, font=footer_font)
            kaynak_w = bbox[2] - bbox[0]
        except Exception:
            kaynak_w = 0
        if kaynak_w > target_width - 44:
            tam_kaynak_metni = kaynak_metni
            kalan_uzunluk = len(tam_kaynak_metni)
            while kalan_uzunluk > 8:
                kalan_uzunluk -= 2
                kaynak_metni = tam_kaynak_metni[:kalan_uzunluk].rstrip() + "…"
                try:
                    bbox = footer.textbbox((0, 0), kaynak_metni, font=footer_font)
                    if (bbox[2] - bbox[0]) <= target_width - 44:
                        break
                except Exception:
                    break
        footer.line((0, target_height, target_width, target_height), fill="#444444", width=2)
        footer.text((22, target_height + 22), kaynak_metni, font=footer_font, fill="#222222")
        footer.rectangle((1, 1, result.width - 2, result.height - 2), outline="#222222", width=3)
        return result

    def parsel_haritasi_onizleme_onayi(self, image):
        pencere = self.animasyonlu_pencere()
        pencere.title("Parsel Haritası Önizlemesi")
        pencere.transient(self.root)
        pencere.grab_set()
        preview = image.copy()
        preview.thumbnail((820, 600), Image.Resampling.LANCZOS)
        photo = ImageTk.PhotoImage(preview, master=pencere)
        label = tk.Label(pencere, image=photo, background="white")
        label.image = photo
        label.pack(fill="both", expand=True, padx=12, pady=12)
        result = {"approved": False}
        buttons = tk.Frame(pencere)
        buttons.pack(fill="x", padx=12, pady=(0, 12))
        tk.Button(buttons, text="İptal", command=pencere.destroy).pack(side="right")

        def approve():
            result["approved"] = True
            pencere.destroy()

        tk.Button(buttons, text="Görüntüyü Kullan", command=approve).pack(side="right", padx=(0, 8))
        pencere.protocol("WM_DELETE_WINDOW", pencere.destroy)
        self.root.wait_window(pencere)
        return result["approved"]

    def parsel_haritasi_hazirla(self):
        if getattr(self.app, "_harita_disari_aktarim_aktif", False):
            messagebox.showwarning("İşlem Sürüyor", "Önce devam eden harita dışa aktarımının bitmesini bekleyin.")
            return False
        try:
            points = parsel_noktalarini_dogrula(getattr(self, "yuklu_kml_points", []))
        except ValueError as exc:
            messagebox.showwarning("Parsel Haritası", f"Önce geçerli parsel KML'sini yükleyin.\n\n{exc}")
            return False

        try:
            old_position = self.map_widget.get_position()
        except Exception:
            old_position = None
        try:
            old_zoom = float(getattr(self.map_widget, "zoom", 15))
        except Exception:
            old_zoom = 15.0

        self._harita_disari_aktarim_aktif = True
        image = None
        try:
            self.jeoloji_kutuphanesi_islemleri().jeoloji_harita_ciktisi_baslat()
            self.sil_ve_yeniden_ciz(False, False, False)
            if getattr(self, "kml_polygon_obj", None):
                self.kml_polygon_obj.delete()
                self.kml_polygon_obj = None
            self.map_widget.set_tile_server(
                "https://mt0.google.com/vt/lyrs=y&hl=tr&x={x}&y={y}&z={z}&s=Ga",
                max_zoom=22,
            )
            widget = self.harita_yakalama_widgeti()
            width = max(320, widget.winfo_width())
            height = max(240, widget.winfo_height())
            left, top, _, _ = self.harita_kirpma_miktarlari(width, height)
            capture_width = max(200, width - (2 * left))
            capture_height = max(160, height - (2 * top))
            center, zoom = parsel_gorunum_hesapla(points, capture_width, capture_height)
            map_img = self.yerbulduru_harita_goruntusu_al(
                center,
                zoom,
                zaman_asimi=20.0,
                ortali_kirp=True,
            )
            if self.yerbulduru_goruntu_bos_mu(map_img):
                raise RuntimeError("Uydu haritasının bazı karoları yüklenemedi.")
            image = self.parsel_haritasi_resmi_olustur(map_img, points, center, zoom)
        except Exception as exc:
            self.hata_kaydet("Parsel haritası hazırlanamadı", exc)
            messagebox.showerror("Parsel Haritası", f"Parsel haritası hazırlanamadı:\n{exc}")
            return False
        finally:
            try:
                self.map_widget.set_tile_server(
                    "https://mt0.google.com/vt/lyrs=s&hl=en&x={x}&y={y}&z={z}&s=Ga",
                    max_zoom=22,
                )
                if old_position:
                    self.map_widget.set_position(old_position[0], old_position[1])
                self.map_widget.set_zoom(old_zoom)
                self.kml_polygon_obj = self.map_widget.set_polygon(
                    points,
                    fill_color=None,
                    outline_color=CALISAN_PARSEL_SINIR_RENGI,
                    border_width=CALISAN_PARSEL_SINIR_KALINLIGI,
                )
                self.sil_ve_yeniden_ciz(True, True, True)
            except Exception as restore_error:
                self.hata_kaydet("Parsel haritası sonrası harita görünümü geri yüklenemedi", restore_error)
            self._harita_disari_aktarim_aktif = False
            self.jeoloji_kutuphanesi_islemleri().jeoloji_harita_ciktisi_bitir()

        if image is None or not self.parsel_haritasi_onizleme_onayi(image):
            return False
        folder = filedialog.askdirectory(title="Parsel Haritasının Kaydedileceği Klasörü Seçin")
        if not folder:
            return False
        path = os.path.join(folder, "Parsel_Haritasi.png")
        if not self.mevcut_dosyalar_icin_onay_al([path], "Mevcut Parsel Haritasının Üzerine Yazılsın mı?"):
            return False
        image.save(path, format="PNG", optimize=True)
        self.img_parsel_haritasi = path
        self.parsel_haritasi_geometri_hash = parsel_noktalari_hashi(points)
        self.parsel_haritasi_ada = self.proje_degeri("ADA", "")
        self.parsel_haritasi_parsel = self.proje_degeri("PARSEL", "")
        if hasattr(self, "proje_durum_seridi_guncelle"):
            self.proje_durum_seridi_guncelle()
        if hasattr(self, "durum_mesaji_yaz"):
            self.durum_mesaji_yaz(f"Parsel haritası kaydedildi: {os.path.basename(path)}")
        return True

    def resim_cek_ve_isaj_ekle(self, path, ac_yn_goster, ss_goster, baslik="MÜHENDİSLİK JEOLOJİSİ HARİTASI", a4_format=True):
        self.sil_ve_yeniden_ciz(ac_yn_goster, ss_goster, m_goster=False) # Resimlerde merkez noktasını GİZLE
        self.map_widget.update_idletasks(); self.root.update()
        time.sleep(0.1)
        
        map_img = self.harita_goruntusu_yakala()
        
        # OKUNABİLİRLİĞİ ARTIRMAK İÇİN GÖRÜNTÜYÜ x3 BÜYÜT (Yüksek Çözünürlük)
        scale = 3.0
        w, h = map_img.size
        w = int(w * scale)
        h = int(h * scale)
        map_img = map_img.resize((w, h), Image.Resampling.LANCZOS)
        
        map_draw = ImageDraw.Draw(map_img)
        
        # 1. Harita Çerçevesi
        try:
            font_title = ImageFont.truetype("arialbd.ttf", int(46 * scale))
            font_buyuk = ImageFont.truetype("arialbd.ttf", int(38 * scale))
            font_normal = ImageFont.truetype("arial.ttf", int(34 * scale))
            font_bold = ImageFont.truetype("arialbd.ttf", int(34 * scale))
            font_dev = ImageFont.truetype("arialbd.ttf", int(70 * scale))
        except:
            font_title = font_buyuk = font_normal = font_bold = font_dev = ImageFont.load_default()
            
        secim = self.combo_formasyon.get()
        kisa_kod = ""; tam_ad = ""
        if secim and "(" in secim:
            tam_ad = secim.split("(")[0].strip()
            kisa_kod = secim.split("(")[1].replace(")", "").strip()
        zemin_sinifi = self.proje_degeri("YEREL_ZEMIN_SINIFI", "-") or "-"

        # 2. Harita Üzerine Formasyon Metni (Sol Üst Köşeye)
        if ac_yn_goster and kisa_kod and "MÜHENDİSLİK" in baslik: 
            yass = self.proje_yass_degeri()
            metin = f"{kisa_kod}\n{zemin_sinifi}\nYASS= {yass}"
            lines_m = metin.split('\n')
            ty = int(50 * scale)
            tx = int(50 * scale)
            for lm in lines_m:
                 try: bbox = map_draw.textbbox((0,0), lm, font=font_dev)
                 except: bbox = (0, 0, len(lm)*30, 50)
                 try: map_draw.text((tx, ty), lm, fill="white", font=font_dev, stroke_width=5, stroke_fill="black")
                 except: map_draw.text((tx, ty), lm, fill="white", font=font_dev) # pillow eski versiyon koruması
                 ty += (bbox[3]-bbox[1]) + int(20 * scale)

        # 3. Puanları topla
        puanlar = []
        for isim, data in self.harita_isaretleri.items():
            if data["tip"] == "M": continue
            gosterim_isim = isim.replace("SS", "SİS")
            if data["tip"] in ["AÇ", "YN"] and ac_yn_goster:
                puanlar.append((gosterim_isim, data["lat"], data["lon"]))
            if data["tip"] == "SS" and ss_goster:
                puanlar.append((gosterim_isim, data["n1"][0], data["n1"][1]))
        
        try: puanlar.sort(key=lambda x: int(''.join(filter(str.isdigit, x[0]))))
        except: puanlar.sort(key=lambda x: x[0])

        if not a4_format:
            # Lokasyon haritalarında yalnız tek ve ince bir dış sınır kullan.
            map_draw.rectangle(
                [0, 0, w - 1, h - 1],
                outline="black",
                width=max(2, int(2 * scale)),
            )
            map_img.save(path, quality=95)
            return

        # Build legend items early to dynamically calculate footer right-side height
        legend_items = []
        if ac_yn_goster and kisa_kod and "MÜHENDİSLİK" in baslik:
            legend_items.append((kisa_kod, tam_ad, "box"))
            legend_items.append((zemin_sinifi, "Zemin Sınıfı", "box"))
        if ac_yn_goster:
            ac_yn_str = ""
            isimler = [p[0] for p in puanlar]
            if any("AÇ" in n for n in isimler): ac_yn_str += "Araştırma Çukuru "
            if any("YN" in n for n in isimler): ac_yn_str += "/ Yüzey Numunesi"
            if not ac_yn_str: ac_yn_str = "Araştırma Çukuru / Yüzey Numunesi"
            ac_yn_str = ac_yn_str.strip(" /")
            sample_isim = next((n for n in isimler if "AÇ" in n or "YN" in n), "YN/AÇ")
            legend_items.append((sample_isim, ac_yn_str, "box"))
        if ss_goster:
            sample_ss = next((p[0] for p in puanlar if "SİS" in p[0]), "SİS")
            legend_items.append((sample_ss, "Sismik Profil / Serim", "ss_line"))
            
        if self.kml_polygon_obj:
            legend_items.append(("", "Çalışma Alanı Sınırı", "line"))
            
        header_h = int(105 * scale)
        coord_title_h = int(60 * scale)
        coord_header_h = int(52 * scale)
        coord_row_h = int(64 * scale)
        legend_row_h = int(84 * scale)
        left_h = coord_title_h + coord_header_h + len(puanlar) * coord_row_h + int(28 * scale)
        right_h = coord_title_h + len(legend_items) * legend_row_h + int(28 * scale)
        footer_h = max(int(260 * scale), left_h, right_h)
        
        new_w = max(w, int(800 * scale))
        new_h = header_h + h + footer_h
        
        final_img = Image.new("RGB", (new_w, new_h), "white")
        final_draw = ImageDraw.Draw(final_img)
        
        # Haritayı yerleştir
        offset_x = (new_w - w) // 2
        final_img.paste(map_img, (offset_x, header_h))
        
        # --- HEADER ---
        try: t_bbox = final_draw.textbbox((0,0), baslik, font=font_title)
        except: t_bbox = (0, 0, len(baslik)*50, 100)
        title_h = t_bbox[3] - t_bbox[1]
        final_draw.text(
            (new_w//2 - (t_bbox[2]-t_bbox[0])//2, (header_h - title_h)//2 - t_bbox[1]),
            baslik,
            fill="black",
            font=font_title,
        )
        ince_cizgi = max(2, int(1 * scale))
        final_draw.line([(0, header_h), (new_w, header_h)], fill="#666666", width=ince_cizgi)
        
        # --- FOOTER ---
        fy = header_h + h
        final_draw.line([(0, fy), (new_w, fy)], fill="#666666", width=ince_cizgi)
        
        middle_x = new_w // 2
        final_draw.line([(middle_x, fy), (middle_x, new_h)], fill="#888888", width=ince_cizgi)

        def ortali_x(sol, sag, metin, font):
            try:
                bbox = final_draw.textbbox((0, 0), metin, font=font)
                genislik = bbox[2] - bbox[0]
            except Exception:
                genislik = len(metin) * int(20 * scale)
            return int(sol + ((sag - sol) - genislik) / 2)
        
        # Left Side (Koordinatlar)
        l_title = "Koordinatlar (WGS84)"
        try: lt_bbox = final_draw.textbbox((0,0), l_title, font=font_buyuk)
        except: lt_bbox = (0, 0, len(l_title)*36, 60)
        final_draw.text(
            ((middle_x)//2 - (lt_bbox[2]-lt_bbox[0])//2, fy + int(10*scale)),
            l_title,
            fill="black",
            font=font_buyuk,
        )
        
        koord_baslik_alt = fy + coord_title_h
        koord_kolon_alt = koord_baslik_alt + coord_header_h
        final_draw.line([(0, koord_baslik_alt), (middle_x, koord_baslik_alt)], fill="#888888", width=ince_cizgi)
        
        col1_w = middle_x * 0.30
        col2_w = middle_x * 0.35
        col3_w = middle_x * 0.35
        
        alt_baslik_y = koord_baslik_alt + int(7*scale)
        final_draw.text((ortali_x(col1_w, col1_w + col2_w, "Enlem", font_bold), alt_baslik_y), "Enlem", fill="black", font=font_bold)
        final_draw.text((ortali_x(col1_w + col2_w, middle_x, "Boylam", font_bold), alt_baslik_y), "Boylam", fill="black", font=font_bold)
        
        final_draw.line([(0, koord_kolon_alt), (middle_x, koord_kolon_alt)], fill="#888888", width=ince_cizgi)
        final_draw.line([(col1_w, koord_baslik_alt), (col1_w, new_h)], fill="#aaaaaa", width=ince_cizgi)
        final_draw.line([(col1_w + col2_w, koord_baslik_alt), (col1_w + col2_w, new_h)], fill="#aaaaaa", width=ince_cizgi)
        
        ry = koord_kolon_alt + int(10*scale)
        for p in puanlar:
            final_draw.text((int(22*scale), ry), p[0], fill="black", font=font_normal)
            final_draw.text((col1_w + int(12*scale), ry), f"{p[1]:.6f}°", fill="black", font=font_normal)
            final_draw.text((col1_w + col2_w + int(12*scale), ry), f"{p[2]:.6f}°", fill="black", font=font_normal)
            ry += coord_row_h
            
        # Right Side (Açıklamalar)
        r_title = "Açıklamalar"
        try: rt_bbox = final_draw.textbbox((0,0), r_title, font=font_buyuk)
        except: rt_bbox = (0, 0, len(r_title)*36, 60)
        final_draw.text(
            (middle_x + (middle_x)//2 - (rt_bbox[2]-rt_bbox[0])//2, fy + int(10*scale)),
            r_title,
            fill="black",
            font=font_buyuk,
        )
        
        final_draw.line([(middle_x, koord_baslik_alt), (new_w, koord_baslik_alt)], fill="#888888", width=ince_cizgi)
        
        ly = koord_baslik_alt + int(12*scale)
        lx = middle_x + int(32*scale)
        for item in legend_items:
            k_kod, k_ad, k_tip = item
            box_w = int(135*scale)
            box_h = int(60*scale)
            if k_tip == "box":
                final_draw.rectangle([lx, ly, lx + box_w, ly + box_h], outline="black", width=int(2*scale))
                try: t_b = final_draw.textbbox((0,0), k_kod, font=font_normal)
                except: t_b = (0, 0, len(k_kod)*30, 48)
                final_draw.text(
                    (
                        lx + box_w//2 - (t_b[2]-t_b[0])//2,
                        ly + (box_h - (t_b[3]-t_b[1]))//2 - t_b[1],
                    ),
                    k_kod,
                    fill="black",
                    font=font_normal,
                )
            elif k_tip == "ss_line":
                final_draw.rectangle([lx, ly, lx + box_w, ly + box_h], outline="black", width=int(2*scale))
                x = lx + int(12 * scale)
                x_son = lx + box_w - int(12 * scale)
                y = ly + box_h // 2
                kesik = int(16 * scale)
                bosluk = int(9 * scale)
                while x < x_son:
                    final_draw.line(
                        [(x, y), (min(x + kesik, x_son), y)],
                        fill=SS_SERIM_RENGI,
                        width=max(4, int(SS_SERIM_KALINLIGI * scale)),
                    )
                    x += kesik + bosluk
            elif k_tip == "line":
                final_draw.rectangle([lx, ly, lx + box_w, ly + box_h], outline="black", width=int(2*scale))
                final_draw.line(
                    [
                        (lx + int(12*scale), ly + box_h - int(12*scale)),
                        (lx + box_w - int(12*scale), ly + int(12*scale)),
                    ],
                    fill=CALISAN_PARSEL_SINIR_RENGI,
                    width=max(4, int(CALISAN_PARSEL_SINIR_KALINLIGI * scale)),
                )

            try:
                ad_bbox = final_draw.textbbox((0, 0), k_ad, font=font_normal)
                ad_y = ly + (box_h - (ad_bbox[3] - ad_bbox[1])) // 2 - ad_bbox[1]
            except Exception:
                ad_y = ly + int(8 * scale)
            final_draw.text((lx + box_w + int(24*scale), ad_y), k_ad, fill="black", font=font_normal)
            ly += legend_row_h

        # Pafta, içerik bittikten sonra sona erer; tek ve ince bir dış çerçeve kullanılır.
        cerceve_payi = max(3, int(3 * scale))
        final_draw.rectangle(
            [cerceve_payi, cerceve_payi, new_w - cerceve_payi - 1, new_h - cerceve_payi - 1],
            outline="black",
            width=max(2, int(2 * scale)),
        )
             
        final_img.save(path, quality=95)

    def haritalari_hazirla(self):
        # YEREL HARİTALAR (Klasör de soralım ki MJH yi vs de kaydetsin)
        if getattr(self.app, "_harita_disari_aktarim_aktif", False):
            messagebox.showwarning("İşlem Sürüyor", "Önce devam eden harita dışa aktarımının bitmesini bekleyin.")
            return
        if not messagebox.askyesno("Haritaları Kaydet", "Haritalar rapor için JPG kalitesinde dışarı aktarılacaktır.\nTamam dedikten sonra klasör seçin ve işlem bitene kadar farenizi hareket ettirmeyin."): return
        kayit_klasoru = filedialog.askdirectory(title="Harita Görüntülerinin Kaydedileceği Klasörü Seçin")
        if not kayit_klasoru: return
        
        img_mjh = os.path.join(kayit_klasoru, "Mühendislik_Jeolojisi_Haritasi.jpg")
        img_jeofizik_lok = os.path.join(kayit_klasoru, "Jeofizik_Lokasyon_Haritasi.jpg")
        img_jeoloji_lok = os.path.join(kayit_klasoru, "Jeoloji_Lokasyon_Haritasi.jpg")
        yollar = [img_mjh, img_jeofizik_lok, img_jeoloji_lok]
        secim = self.mevcut_harita_dosyalari_secimi(yollar)
        if secim == "iptal":
            return
        self.img_mjh = img_mjh
        self.img_jeofizik_lok = img_jeofizik_lok
        self.img_jeoloji_lok = img_jeoloji_lok
        if secim == "kullan":
            if hasattr(self, "durum_mesaji_yaz"):
                self.durum_mesaji_yaz("Mevcut 3 harita projeye bağlandı")
            return

        self._harita_disari_aktarim_aktif = True
        try:
            self.jeoloji_kutuphanesi_islemleri().jeoloji_harita_ciktisi_baslat()
            self.resim_cek_ve_isaj_ekle(self.img_mjh, True, True, "MÜHENDİSLİK JEOLOJİSİ HARİTASI", a4_format=True)
            self.resim_cek_ve_isaj_ekle(self.img_jeofizik_lok, False, True, "JEOFİZİK LOKASYON HARİTASI", a4_format=False)
            self.resim_cek_ve_isaj_ekle(self.img_jeoloji_lok, True, False, "JEOLOJİ LOKASYON HARİTASI", a4_format=False)
            
            self.sil_ve_yeniden_ciz(True, True, True)
            if hasattr(self, "durum_mesaji_yaz"):
                self.durum_mesaji_yaz("3 yerel harita hazırlandı", kayit_klasoru)
        except Exception as e:
            self.hata_kaydet("Harita görüntüleri hazırlanırken hata oluştu", e)
            self.sil_ve_yeniden_ciz(True, True, True)
            messagebox.showerror("Hata", f"Harita görüntüleri hazırlanamadı:\n{e}")
        finally:
            self._harita_disari_aktarim_aktif = False
            self.jeoloji_kutuphanesi_islemleri().jeoloji_harita_ciktisi_bitir()

    def yerbulduru_hazirla(self):
        if getattr(self.app, "_harita_disari_aktarim_aktif", False):
            messagebox.showwarning("İşlem Sürüyor", "Önce devam eden harita dışa aktarımının bitmesini bekleyin.")
            return
        if not messagebox.askyesno("Yerbulduru Haritası", "Yerbulduru haritası iki panelli olarak üretilecektir: üstte geniş konum, altta yakın çalışma alanı görünümü.\nTamam dedikten sonra kaydedilecek yeri seçin ve fareyi oynatmayın."): return
        kayit_klasoru = filedialog.askdirectory(title="Yerbulduru Haritasının Kaydedileceği Klasörü Seçin")
        if not kayit_klasoru: return
        
        img_yerbulduru = os.path.join(kayit_klasoru, "Yerbulduru_Haritasi.jpg")
        secim = self.mevcut_harita_dosyalari_secimi(
            [img_yerbulduru],
            "Mevcut Yerbulduru Haritası Bulundu",
        )
        if secim == "iptal":
            return
        self.img_yerbulduru = img_yerbulduru
        if secim == "kullan":
            if hasattr(self, "durum_mesaji_yaz"):
                self.durum_mesaji_yaz("Mevcut yerbulduru haritası projeye bağlandı")
            return
        
        try:
            mevcut_konum = self.map_widget.get_position()
        except Exception:
            mevcut_konum = None
        mevcut_zoom = getattr(self.map_widget, "zoom", 15)
        try:
            mevcut_zoom_int = int(mevcut_zoom)
        except Exception:
            mevcut_zoom_int = 15
        mevcut_zoom_int = min(22, max(1, mevcut_zoom_int))

        self._harita_disari_aktarim_aktif = True
        try:
            self.jeoloji_kutuphanesi_islemleri().jeoloji_harita_ciktisi_baslat()
            merkez = self.yerbulduru_merkez_koordinati()
            isaret_koordinati = self.yerbulduru_isaret_koordinati()
            genis_merkez, genis_zoom = self.yerbulduru_genis_gorunum_ayarlari(merkez)
            yakin_merkez = mevcut_konum if mevcut_konum else merkez

            self.sil_ve_yeniden_ciz(False, False, False)
            genis_img = self.yerbulduru_harita_goruntusu_al(genis_merkez, genis_zoom)
            if self.yerbulduru_goruntu_bos_mu(genis_img):
                raise RuntimeError("Proje merkezli geniş görünüm yüklenemedi. İnternet bağlantısını kontrol edip işlemi tekrar deneyin.")
            genis_pin = self.yerbulduru_ekran_noktasi_hesapla(isaret_koordinati, genis_merkez, genis_zoom)

            yakin_img = self.yerbulduru_harita_goruntusu_al(yakin_merkez, mevcut_zoom_int, kml_sinir_goster=True)
            if self.yerbulduru_goruntu_bos_mu(yakin_img):
                raise RuntimeError("Yakın yerbulduru görünümü yüklenemedi. Harita görüntüsü geldikten sonra işlemi tekrar deneyin.")
            yakin_pin = self.yerbulduru_ekran_noktasi_hesapla(isaret_koordinati, yakin_merkez, mevcut_zoom_int)

            self.yerbulduru_iki_panel_resim_olustur(
                genis_img,
                yakin_img,
                self.img_yerbulduru,
                genis_pin=genis_pin,
                yakin_pin=yakin_pin,
            )
            
            if mevcut_konum:
                self.map_widget.set_position(mevcut_konum[0], mevcut_konum[1])
                # Export uses a full tile zoom for deterministic raster
                # output; restore the exact interactive quarter-step view.
                self.map_widget.set_zoom(mevcut_zoom)
            self.sil_ve_yeniden_ciz(True, True, True)
            if hasattr(self, "durum_mesaji_yaz"):
                self.durum_mesaji_yaz("Yerbulduru haritası hazırlandı", kayit_klasoru)
        except Exception as e:
            self.hata_kaydet("Yerbulduru haritası hazırlanırken hata oluştu", e)
            if mevcut_konum:
                try:
                    self.map_widget.set_position(mevcut_konum[0], mevcut_konum[1])
                    self.map_widget.set_zoom(mevcut_zoom)
                except Exception:
                    pass
            self.sil_ve_yeniden_ciz(True, True, True)
            messagebox.showerror("Hata", f"Yerbulduru haritası hazırlanamadı:\n{e}")
        finally:
            self._harita_disari_aktarim_aktif = False
            self.jeoloji_kutuphanesi_islemleri().jeoloji_harita_ciktisi_bitir()

    def tum_loglari_ciz(self):
        kayitlar = self.ac_yn_sekme_kayitlari()
        if not kayitlar:
            messagebox.showwarning("Uyarı", "Henüz hiçbir Araştırma Çukuru oluşturulmamış.")
            return
            
        kayit_klasoru = filedialog.askdirectory(title="Log JPG Çıktılarının Kaydedileceği Klasörü Seçin")
        if not kayit_klasoru: return

        ciktilar = []
        kullanilan_adlar = set()
        for sira, kayit in enumerate(kayitlar, start=1):
            isim = kayit["isim"]
            guvenli_ad = guvenli_dosya_adi(isim, varsayilan=f"Kayit_{sira}")
            aday = guvenli_ad
            sayac = 2
            while aday.casefold() in kullanilan_adlar:
                aday = f"{guvenli_ad}_{sayac}"
                sayac += 1
            kullanilan_adlar.add(aday.casefold())
            ciktilar.append((kayit, isim, os.path.join(kayit_klasoru, f"{aday}_Logu.jpg")))
        if not self.mevcut_dosyalar_icin_onay_al(
            [yol for _, _, yol in ciktilar],
            "Mevcut Log Görsellerinin Üzerine Yazılsın mı?",
        ):
            return

        basarili = 0
        hatali = []
        for kayit, isim, res_path in ciktilar:
            try:
                self.tekil_log_ciz(kayit, isim, res_path)
                basarili += 1
            except Exception as e:
                self.hata_kaydet(f"{isim} log görseli çizilirken hata oluştu", e)
                hatali.append(f"{isim}: {e}")

        if hatali:
            messagebox.showwarning(
                "Log Üretimi Kısmen Tamamlandı",
                f"{basarili} log oluşturuldu, {len(hatali)} log oluşturulamadı.\n\n"
                + "\n".join(hatali[:8]),
            )
        else:
            if hasattr(self, "durum_mesaji_yaz"):
                self.durum_mesaji_yaz(f"{basarili} log görüntüsü hazırlandı", kayit_klasoru)

    def tekil_log_ciz(self, kayit, isim, kayit_yolu):
        derinlik = kayit["derinlik_entry"].get()
        enlem = kayit["enlem_entry"].get()
        boylam = kayit["boylam_entry"].get()
        tarih = kayit["tarih_entry"].get()
        satirlar = self.ac_yn_satirlari(kayit)

        try:
            girilen_derinlik = float(str(derinlik).replace(",", "."))
        except (TypeError, ValueError):
            girilen_derinlik = 0.0
        if not math.isfinite(girilen_derinlik) or girilen_derinlik < 0:
            raise ValueError(f"{isim} için çukur derinliği geçerli bir pozitif sayı olmalıdır.")

        gecerli_satirlar = []
        gecersiz_derinlikler = []
        for satir in satirlar:
            aralik = derinlik_araligi_oku(satir[0] if satir else "")
            if aralik is None:
                gecersiz_derinlikler.append(str(satir[0] if satir else ""))
                continue
            gecerli_satirlar.append((satir, aralik[0], aralik[1]))
        if gecersiz_derinlikler:
            raise ValueError(
                f"{isim} içinde geçersiz derinlik aralığı var: " + ", ".join(gecersiz_derinlikler[:5])
            )
        veri_derinligi = max((alt for _, _, alt in gecerli_satirlar), default=0.0)
        total_depth = max(girilen_derinlik, veri_derinligi)
        if total_depth <= 0:
            raise ValueError(f"{isim} için çizilecek derinlik bulunamadı.")
        if total_depth > 100:
            raise ValueError(f"{isim} derinliği 100 m üst sınırını aşıyor.")

        img_w = 1600
        y_bas = 100 + 7 * 45  
        baslik_h = 80
        y_data_start = y_bas + baslik_h 
        data_h = max(45, min(260, int(1800 / total_depth)))
        footer_y = y_data_start + int(total_depth * data_h)
        img_h = max(1700, footer_y + 500)

        img = Image.new("RGB", (img_w, img_h), "white")
        draw = ImageDraw.Draw(img)
        
        try:
            font_title = ImageFont.truetype("arialbd.ttf", 34)
            font_h = ImageFont.truetype("arialbd.ttf", 22)
            font_bold = ImageFont.truetype("arialbd.ttf", 18)
            font_n = ImageFont.truetype("arial.ttf", 20)
            font_k = ImageFont.truetype("arial.ttf", 16)
            font_s = ImageFont.truetype("arial.ttf", 14)
        except:
            font_title = font_h = font_bold = font_n = font_k = font_s = ImageFont.load_default()

        renk_mavi = "#87CEEB"
        
        draw.rectangle([50, 50, img_w-50, 100], fill="white", outline="black", width=2)
        
        log_basligi = "YÜZEY NUMUNESİ LOGU" if str(isim).upper().startswith("YN") else "ARAŞTIRMA ÇUKURU LOGU"
        try: t_bbox = draw.textbbox((0,0), log_basligi, font=font_title); t_w = t_bbox[2]-t_bbox[0]
        except: t_w = 400
        draw.text(((img_w - t_w)//2, 55), log_basligi, fill="black", font=font_title)

        proje_ad = self.proje_degeri("PROJE_ADI")
        firma_adi = self.proje_degeri("FIRMA_ADI", "-") or "-"
        il = self.proje_degeri("IL")
        ilce = self.proje_degeri("ILCE")
        ada = self.proje_degeri("ADA")
        parsel = self.proje_degeri("PARSEL")
            
        yass_val = "-"
        for s in satirlar:
            if len(s)>2 and str(s[2]).strip() and str(s[2]).strip() != "-":
                yass_val = str(s[2]).strip()
                break
        
        y_offs = 100
        row_h = 45
        
        kunye_sol = [("Firma Adı", firma_adi), ("Proje Adı", proje_ad), ("İl", il), ("İlçe", ilce), ("Ada", ada), ("Parsel", parsel), ("Yeraltı Suyu (m)", yass_val)]
        kunye_sag = [("Çukur No:", isim), ("Koordinatlar", f"X: {boylam}    Y: {enlem}"), ("Tarih", tarih), ("Çukur Derinliği", f"{total_depth:g}"), ("Kazıcı Tipi", ""), ("Zemin Kotu", ""), ("Operatör", "")]
        
        for idx in range(7):
            draw.rectangle([50, y_offs, 250, y_offs+row_h], fill=renk_mavi, outline="black", width=2)
            draw.text((60, y_offs+10), kunye_sol[idx][0], fill="black", font=font_bold)
            
            draw.rectangle([250, y_offs, 800, y_offs+row_h], fill="white", outline="black", width=2)
            draw.text((260, y_offs+10), kunye_sol[idx][1], fill="black", font=font_n)
            
            draw.rectangle([800, y_offs, 1100, y_offs+row_h], fill=renk_mavi, outline="black", width=2)
            draw.text((810, y_offs+10), kunye_sag[idx][0], fill="black", font=font_bold)
            
            draw.rectangle([1100, y_offs, img_w-50, y_offs+row_h], fill="white", outline="black", width=2)
            draw.text((1110, y_offs+10), kunye_sag[idx][1], fill="black", font=font_n)
            
            y_offs += row_h

        y_bas = y_offs
        baslik_h = 80
        
        c_w = [80, 100, 200, 80, 300, 80, 70, 60] 
        lab_w_total = (img_w - 50) - (50 + sum(c_w)) 
        lab_cols = ["Wn %", "LL %", "PL %", "PI %", "+4 %", "-200 %"]
        lw = lab_w_total // 6 
        
        cols = ["Derinlik\n(m)", "Örnek Tipi", "Profil", "YASS", "Zemin Tanımlaması", "USCS", "NEM", "PP"]
        x_c = 50
        
        for idx, col in enumerate(cols):
            draw.rectangle([x_c, y_bas, x_c+c_w[idx], y_bas+baslik_h], fill="white", outline="black", width=2)
            try: tb = draw.textbbox((0,0), col, font=font_k); tw = tb[2]-tb[0]; th = tb[3]-tb[1]
            except: tw = 50; th = 15
            lines = col.split('\n')
            ty = y_bas + (baslik_h//2) - (len(lines)*10)
            for l in lines:
                try: lb = draw.textbbox((0,0), l, font=font_k); lw_txt = lb[2]-lb[0]
                except: lw_txt = 40
                draw.text((x_c + (c_w[idx]-lw_txt)//2, ty), l, fill="black", font=font_k)
                ty += 20
            x_c += c_w[idx]
            
        draw.rectangle([x_c, y_bas, img_w-50, y_bas+40], fill="white", outline="black", width=2)
        try: tb = draw.textbbox((0,0), "Laboratuvar Sonuçları", font=font_k); tw = tb[2]-tb[0]
        except: tw=100
        draw.text((x_c + (lab_w_total-tw)//2, y_bas + 10), "Laboratuvar Sonuçları", fill="black", font=font_k)
        
        lx = x_c
        for lcol in lab_cols:
            draw.rectangle([lx, y_bas+40, lx+lw, y_bas+baslik_h], fill="white", outline="black", width=2)
            try: ltb = draw.textbbox((0,0), lcol, font=font_k); ltw = ltb[2]-ltb[0]
            except: ltw = 30
            draw.text((lx + (lw-ltw)//2, y_bas + 50), lcol, fill="black", font=font_k)
            lx += lw

        x_c = 50
        draw.rectangle([x_c, y_bas, img_w-50, footer_y], outline="black", width=3)
        for cw in c_w:
            draw.line([x_c, y_bas, x_c, footer_y], fill="black", width=2)
            x_c += cw
        lx = x_c
        for _ in lab_cols:
            draw.line([lx, y_bas+40, lx, footer_y], fill="black", width=1)
            lx += lw
            
        y_curr = y_data_start
        k = 0
        while y_curr <= footer_y:
            draw.line([50, y_curr, img_w-50, y_curr], fill="black", width=1) 
            if y_curr + (data_h//2) <= footer_y:
                draw.line([50, y_curr+(data_h//2), 80, y_curr+(data_h//2)], fill="black", width=2)
            draw.text((90, y_curr-10), f"{float(k)}", fill="black", font=font_bold)
            y_curr += data_h
            k += 1

        for satir, ust_d, alt_d in gecerli_satirlar:
            y_ust_row = y_data_start + int(ust_d * data_h)
            y_alt_row = y_data_start + int(alt_d * data_h)
            row_pix_h = y_alt_row - y_ust_row
            if row_pix_h <= 0: row_pix_h = data_h
            
            if y_alt_row <= footer_y: 
                draw.line([50, y_alt_row, img_w-50, y_alt_row], fill="black", width=2)
            
            x_c = 50
            x_c += c_w[0]
            val_ornek = str(satir[1]) if len(satir)>1 else ""
            draw.text((x_c+10, y_ust_row + (row_pix_h//2) - 10), val_ornek, fill="black", font=font_n)
            x_c += c_w[1]
            
            zemin_tanimi = str(satir[3]) if len(satir)>3 else ""
            son_kelime = ""
            if zemin_tanimi.strip():
                kelimeler = [k for k in zemin_tanimi.split() if len(k)>1]
                if kelimeler: son_kelime = kelimeler[-1].upper()
            if son_kelime: self.ciz_tarama_deseni(draw, son_kelime, x_c, y_ust_row, x_c+c_w[2], y_alt_row)
            x_c += c_w[2]
            
            val_yass = str(satir[2]) if len(satir)>2 else ""
            draw.text((x_c+10, y_ust_row + (row_pix_h//2) - 10), val_yass, fill="black", font=font_n)
            x_c += c_w[3]
            
            if zemin_tanimi:
                words = zemin_tanimi.split()
                lines = []; cl = ""
                for w in words:
                    if len(cl+w)*10 < c_w[4]-20: cl += w + " "
                    else: lines.append(cl); cl = w + " "
                lines.append(cl)
                ty = y_ust_row + (row_pix_h//2) - (len(lines)*12)
                for l in lines:
                    draw.text((x_c+10, ty), l.strip(), fill="black", font=font_n); ty += 25
            x_c += c_w[4]
            
            val_uscs = str(satir[4]) if len(satir)>4 else ""
            draw.text((x_c+15, y_ust_row + (row_pix_h//2) - 10), val_uscs, fill="black", font=font_n)
            x_c += c_w[5]
            
            x_c += c_w[6] 
            x_c += c_w[7] 
                
            lx = x_c
            yaz_labs = ["", "", "", "", "", ""] 
            if len(satir) > 5: yaz_labs[0] = str(satir[5])
            if len(satir) > 6: yaz_labs[1] = str(satir[6])
            if len(satir) > 7: yaz_labs[2] = str(satir[7])
            if len(satir) > 8: yaz_labs[3] = str(satir[8])
            if len(satir) > 9: yaz_labs[4] = str(satir[9])
            if len(satir) > 10: yaz_labs[5] = str(satir[10])
            
            for lv in yaz_labs:
                try: lb = draw.textbbox((0,0), lv, font=font_n); lw_txt = lb[2]-lb[0]
                except: lw_txt = 10
                draw.text((lx + (lw-lw_txt)//2, y_ust_row + (row_pix_h//2) - 10), lv, fill="black", font=font_n)
                lx += lw

        draw.line([50, footer_y, img_w-50, footer_y], fill="black", width=3)
        
        y_curr = footer_y
        
        draw.rectangle([50, y_curr, img_w//2 + 100, y_curr+40], fill=renk_mavi, outline="black", width=2)
        draw.text((60, y_curr+10), "Açıklamalar", fill="black", font=font_bold)
        
        draw.rectangle([img_w//2 + 100, y_curr, img_w-50, y_curr+40], fill=renk_mavi, outline="black", width=2)
        draw.text((img_w//2 + 110, y_curr+10), "Araştırma Çukuru Fotoğrafı", fill="black", font=font_bold)
        
        y_curr += 40
        f_box_top = y_curr
        
        aciklamalar = [
            ("PP", "Cep Penetrometresi", "VM", "Çok Nemli"),
            ("V", "Veyn Deneyi", "SM", "Az Nemli"),
            ("UD", "Örselenmemiş Örnek", "W", "Islak"),
            ("DS", "Örselenmiş Örnek", "Wn", "Doğal Su İçeriği"),
            ("BN", "Blok Örnek", "LL", "Likit Limit"),
            ("SN", "Silindir Örnek", "PL", "Plastisite İndeksi"),
            ("TN", "Torba Örnek", "+4", "4 nolu elekte kalan"),
            ("D", "Kuru", "-200", "200 nolu elekten geçen")
        ]
        
        c1 = 70; c2 = 250; c3 = 100; c4 = 300
        for ro in aciklamalar:
            draw.rectangle([50, y_curr, img_w//2 + 100, y_curr+25], fill="white", outline="black", width=1)
            draw.line([50+c1, y_curr, 50+c1, y_curr+25], fill="black", width=1)
            draw.line([50+c1+c2, y_curr, 50+c1+c2, y_curr+25], fill="black", width=1)
            draw.line([50+c1+c2+c3, y_curr, 50+c1+c2+c3, y_curr+25], fill="black", width=1)
            draw.text((55, y_curr+3), ro[0], fill="black", font=font_s)
            draw.text((55+c1, y_curr+3), ro[1], fill="black", font=font_s)
            draw.text((55+c1+c2, y_curr+3), ro[2], fill="black", font=font_s)
            draw.text((55+c1+c2+c3, y_curr+3), ro[3], fill="black", font=font_s)
            y_curr += 25
            
        draw.rectangle([50, y_curr, img_w//2 + 100, y_curr+100], fill="white", outline="black", width=2)
        dusunce_text = ""
        aciklama_widgeti = kayit.get("aciklama_text")
        if aciklama_widgeti is not None:
            dusunce_text = aciklama_widgeti.get("1.0", tk.END).strip()
        draw.text((60, y_curr+10), "Düşünceler:\n" + dusunce_text, fill="black", font=font_k)
        
        y_curr += 100
        
        draw.rectangle([img_w//2 + 100, f_box_top, img_w-50, y_curr], fill="white", outline="black", width=2)
        draw.text(((img_w//2 + 100 + img_w-50)//2 - 150, f_box_top + (y_curr-f_box_top)//2), "(Resim Yapıştırma Alanı)", fill="grey", font=font_n)
        
        draw.rectangle([img_w//2 + 100, y_curr, img_w-400, y_curr+50], fill=renk_mavi, outline="black", width=2)
        draw.text((img_w//2 + 110, y_curr+15), "Logu Hazırlayan", fill="black", font=font_bold)
        
        draw.rectangle([img_w-400, y_curr, img_w-50, y_curr+50], fill=renk_mavi, outline="black", width=2)
        draw.text((img_w-390, y_curr+15), "Kontrol", fill="black", font=font_bold)
        
        draw.rectangle([img_w//2 + 100, y_curr+50, img_w-400, y_curr+150], fill="white", outline="black", width=2)
        draw.rectangle([img_w-400, y_curr+50, img_w-50, y_curr+150], fill="white", outline="black", width=2)

        draw.rectangle([10, 10, img_w-10, img_h-10], outline="black", width=4)

        img.save(kayit_yolu, quality=95)

    # ==========================================
    # 4. YENİ EKLENEN JEO. KESİT GÖRSEL MOTORU
    # ==========================================
    def _kesit_ciz_olustur_eski(self, ac_sekmeleri, kayit_yolu):
        """Basit 800x400 boyutunda yatay Jeolojik Kesit şablonu çizip kaydeder."""
        img_w, img_h = 800, 400
        img = Image.new('RGB', (img_w, img_h), 'white')
        draw = ImageDraw.Draw(img)
        
        try:
            font_s = ImageFont.truetype("arial.ttf", 12)
            font_m = ImageFont.truetype("arialbd.ttf", 16)
            font_l = ImageFont.truetype("arialbd.ttf", 26)
        except:
            font_s = font_m = font_l = ImageFont.load_default()
            
        # Ana Çerçeveler
        draw.rectangle([0, 0, img_w-1, img_h-1], outline="black", width=2)
        
        y_baslik = 30
        draw.line([0, y_baslik, img_w, y_baslik], fill="black", width=2)
        
        # 1. Sol Sütun (Derinlik Cetveli) [Genislik: 60]
        x_der = 60
        draw.line([x_der, 0, x_der, img_h], fill="black", width=2)
        try: tb = draw.textbbox((0,0), "Derinlik", font=font_s); tw = tb[2]-tb[0]
        except: tw = 45
        draw.text(((x_der-tw)//2, y_baslik-20), "Derinlik", fill="black", font=font_s)
        
        m_artis = 0.5
        px_per_m = (img_h - y_baslik) / 3.0  # 3 metre toplam yükseklik 
        
        d = 0.0
        while d <= 3.0:
            y_cizgi = y_baslik + (d * px_per_m)
            if y_cizgi >= img_h: y_cizgi = img_h - 1
            draw.line([0, y_cizgi, x_der, y_cizgi], fill="black", width=1 if d%1!=0 else 2)
            
            # Orta tırnaklar
            if d < 3.0:
                for t in range(1, 5):
                    ty = y_cizgi + (px_per_m/5)*t
                    draw.line([x_der-10, ty, x_der, ty], fill="black", width=1)
            
            txt = f"{int(d)}" if d%1==0 else f"{d:.1f}"
            try: lb = draw.textbbox((0,0), txt, font=font_s); lw = lb[2]-lb[0]
            except: lw=10
            draw.text(((x_der-lw)//2, y_cizgi+5), txt, fill="black", font=font_s)
            d += m_artis

        # 2. YN1 ve YN2 Dikey Kuyuları ve Tarama
        x_yn1 = x_der + 40
        x_yn2 = img_w - 40
        
        def _katman_bas(zm_adi, s_kelime, x_sol, x_sag, u_d, a_d):
            y_ust = y_baslik + (u_d * px_per_m)
            y_alt = y_baslik + (a_d * px_per_m)
            self.ciz_tarama_deseni(draw, s_kelime, x_sol, y_ust, x_sag, y_alt)
            
            # Zemin Adını Merkeze Yaz
            mx, my = x_sol + (x_sag - x_sol)//2, y_ust + (y_alt - y_ust)//2
            try: 
                tb = draw.textbbox((0,0), zm_adi, font=font_m)
                tw = tb[2] - tb[0]; th = tb[3] - tb[1]
            except: tw = 50; th = 15
            
            # Yazı etrafına beyaz glow vererek okunurluğu artır
            for offset in [(1,1), (-1,-1), (1,-1), (-1,1), (2,0), (-2,0), (0,2), (0,-2)]:
                draw.text((mx - tw//2 + offset[0], my - th//2 + offset[1]), zm_adi, fill="white", font=font_m)
            draw.text((mx - tw//2, my - th//2), zm_adi, fill="black", font=font_m)
            
            # Derinlik Cetveline Bitiş Derinliğini (alt_d) Yaz
            txt_d = f"{a_d:.2f}"
            try: tbd = draw.textbbox((0,0), txt_d, font=font_s); twd = tbd[2]-tbd[0]
            except: twd = 20
            
            draw.line([x_der-15, y_alt, x_der, y_alt], fill="black", width=2)
            draw.rectangle([x_der - twd - 20, y_alt - 7, x_der - 17, y_alt + 7], fill="white")
            draw.text((x_der - twd - 18, y_alt - 6), txt_d, fill="black", font=font_s)

        if not ac_sekmeleri:
            # Sekme yoksa boş bir taslak çiz
            draw.line([x_yn1, 0, x_yn1, img_h], fill="black", width=2)
            draw.line([x_yn2, 0, x_yn2, img_h], fill="black", width=2)
            draw.text((x_der+2, y_baslik-22), "Kuyu 1", fill="black", font=font_m)
            draw.text((x_yn2+2, y_baslik-22), "Kuyu 2", fill="black", font=font_m)
        elif len(ac_sekmeleri) == 1:
            k1 = ac_sekmeleri[0]
            for satir in k1["satirlar"]:
                try: 
                    derinlikler = str(satir[0]).replace(",", ".").split("-"); ust_d = float(derinlikler[0].strip()); alt_d = float(derinlikler[1].strip())
                except: continue
                if ust_d >= 3.0: continue
                if alt_d > 3.0: alt_d = 3.0
                zemin_adi = str(satir[3]).strip().upper()
                son_kelime = zemin_adi.split()[-1] if zemin_adi else ""
                _katman_bas(zemin_adi, son_kelime, x_der, img_w, ust_d, alt_d)

            # Tarama üzerine dikey kuyu çizgileri ve isimler
            draw.line([x_yn1, 0, x_yn1, img_h], fill="black", width=2)
            draw.line([x_yn2, 0, x_yn2, img_h], fill="black", width=2)
            draw.text((x_der+2, y_baslik-22), k1["isim"], fill="black", font=font_m)
            draw.text((x_yn2+2, y_baslik-22), k1["isim"], fill="black", font=font_m)
        else:
            k1 = ac_sekmeleri[0]
            k2 = ac_sekmeleri[1]
            mid_x = x_der + (img_w - x_der) // 2
            
            # Kuyu 1 Tabakaları (Sol Yarı)
            for satir in k1["satirlar"]:
                try: 
                    derinlikler = str(satir[0]).replace(",", ".").split("-"); ust_d = float(derinlikler[0].strip()); alt_d = float(derinlikler[1].strip())
                except: continue
                if ust_d >= 3.0: continue
                if alt_d > 3.0: alt_d = 3.0
                zemin_adi = str(satir[3]).strip().upper()
                son_kelime = zemin_adi.split()[-1] if zemin_adi else ""
                _katman_bas(zemin_adi, son_kelime, x_der, mid_x, ust_d, alt_d)
                
            # Kuyu 2 Tabakaları (Sağ Yarı)
            for satir in k2["satirlar"]:
                try: 
                    derinlikler = str(satir[0]).replace(",", ".").split("-"); ust_d = float(derinlikler[0].strip()); alt_d = float(derinlikler[1].strip())
                except: continue
                if ust_d >= 3.0: continue
                if alt_d > 3.0: alt_d = 3.0
                zemin_adi = str(satir[3]).strip().upper()
                son_kelime = zemin_adi.split()[-1] if zemin_adi else ""
                _katman_bas(zemin_adi, son_kelime, mid_x, img_w, ust_d, alt_d)

            # Tarama üzerine dikey kuyu çizgileri ve isimler
            draw.line([x_yn1, 0, x_yn1, img_h], fill="black", width=2)
            draw.line([x_yn2, 0, x_yn2, img_h], fill="black", width=2)
            draw.line([mid_x, y_baslik, mid_x, img_h], fill="black", width=2) # Kuyu ayrım çizgisi
            draw.text((x_der+2, y_baslik-22), k1["isim"], fill="black", font=font_m)
            draw.text((x_yn2+2, y_baslik-22), k2["isim"], fill="black", font=font_m)

        img.save(kayit_yolu, "JPEG", quality=90)
        return kayit_yolu

    def kesit_ciz_olustur(self, ac_sekmeleri, kayit_yolu):
        """Tüm AÇ/YN kayıtlarını ve veri derinliğini kırpmadan gösteren dinamik kesit üretir."""
        kayitlar = list(ac_sekmeleri or [])
        hazir_kayitlar = []
        azami_derinlik = 0.0
        for kayit in kayitlar:
            katmanlar = []
            gecersiz = []
            for satir in kayit.get("satirlar", []) or []:
                aralik = derinlik_araligi_oku(satir[0] if satir else "")
                if aralik is None:
                    gecersiz.append(str(satir[0] if satir else ""))
                    continue
                katmanlar.append((satir, aralik[0], aralik[1]))
                azami_derinlik = max(azami_derinlik, aralik[1])
            hazir_kayitlar.append((kayit, katmanlar, gecersiz))

        cizim_derinligi = max(3.0, azami_derinlik)
        if cizim_derinligi > 100:
            raise ValueError("Kesit derinliği 100 m üst sınırını aşıyor.")
        kuyu_sayisi = max(1, len(hazir_kayitlar))
        x_der = 70
        bant_genisligi = 260
        img_w = max(800, x_der + kuyu_sayisi * bant_genisligi)
        baslik_h = 58
        px_per_m = max(25, min(120, int(1600 / cizim_derinligi)))
        img_h = max(480, baslik_h + int(cizim_derinligi * px_per_m) + 35)

        img = Image.new("RGB", (img_w, img_h), "white")
        draw = ImageDraw.Draw(img)
        try:
            font_s = ImageFont.truetype("arial.ttf", 12)
            font_m = ImageFont.truetype("arialbd.ttf", 16)
            font_l = ImageFont.truetype("arialbd.ttf", 23)
        except Exception:
            font_s = font_m = font_l = ImageFont.load_default()

        draw.rectangle([0, 0, img_w - 1, img_h - 1], outline="black", width=2)
        draw.line([0, baslik_h, img_w, baslik_h], fill="black", width=2)
        draw.line([x_der, 0, x_der, img_h], fill="black", width=2)
        draw.text((8, 8), "Derinlik\n(m)", fill="black", font=font_s)
        baslik = "JEOLOJİK KESİT"
        try:
            bbox = draw.textbbox((0, 0), baslik, font=font_l)
            baslik_w = bbox[2] - bbox[0]
        except Exception:
            baslik_w = 180
        draw.text((x_der + (img_w - x_der - baslik_w) // 2, 14), baslik, fill="black", font=font_l)

        if cizim_derinligi <= 5:
            cetvel_adimi = 0.5
        elif cizim_derinligi <= 15:
            cetvel_adimi = 1.0
        elif cizim_derinligi <= 40:
            cetvel_adimi = 2.0
        else:
            cetvel_adimi = 5.0
        derinlik_degeri = 0.0
        while derinlik_degeri <= cizim_derinligi + 1e-9:
            y = baslik_h + derinlik_degeri * px_per_m
            draw.line([0, y, img_w, y], fill="#b7b7b7", width=1)
            etiket = f"{derinlik_degeri:g}"
            draw.text((8, y + 2), etiket, fill="black", font=font_s)
            derinlik_degeri += cetvel_adimi
        alt_y = baslik_h + cizim_derinligi * px_per_m
        draw.line([0, alt_y, img_w, alt_y], fill="black", width=2)
        if not math.isclose((cizim_derinligi / cetvel_adimi) % 1, 0.0, abs_tol=1e-6):
            draw.text((8, alt_y - 15), f"{cizim_derinligi:g}", fill="black", font=font_s)

        if not hazir_kayitlar:
            draw.text((x_der + 20, baslik_h + 25), "Kesit oluşturacak AÇ/YN kaydı bulunmuyor.", fill="black", font=font_m)
        for indeks, (kayit, katmanlar, gecersiz) in enumerate(hazir_kayitlar):
            x1 = x_der + indeks * bant_genisligi
            x2 = x_der + (indeks + 1) * bant_genisligi
            draw.line([x1, baslik_h, x1, img_h], fill="black", width=2)
            draw.line([x2, baslik_h, x2, img_h], fill="black", width=2)
            isim = str(kayit.get("isim", f"Kayıt {indeks + 1}"))
            try:
                bbox = draw.textbbox((0, 0), isim, font=font_m)
                isim_w = bbox[2] - bbox[0]
            except Exception:
                isim_w = len(isim) * 9
            draw.text((x1 + max(4, (bant_genisligi - isim_w) // 2), 18), isim, fill="black", font=font_m)

            for satir, ust_d, alt_d in katmanlar:
                y1 = baslik_h + ust_d * px_per_m
                y2 = baslik_h + alt_d * px_per_m
                zemin_adi = str(satir[3]).strip().upper() if len(satir) > 3 else ""
                self.ciz_tarama_deseni(draw, zemin_adi, x1, y1, x2, y2)
                if zemin_adi:
                    kelimeler = zemin_adi.split()
                    satir_metni = []
                    aktif = ""
                    for kelime in kelimeler:
                        if len(aktif) + len(kelime) + 1 <= 24:
                            aktif = (aktif + " " + kelime).strip()
                        else:
                            satir_metni.append(aktif)
                            aktif = kelime
                    if aktif:
                        satir_metni.append(aktif)
                    metin_y = y1 + max(3, ((y2 - y1) - len(satir_metni) * 16) / 2)
                    for metin_satiri in satir_metni:
                        draw.text((x1 + 8, metin_y), metin_satiri, fill="black", font=font_s, stroke_width=2, stroke_fill="white")
                        metin_y += 16
                draw.text((x1 + 4, y2 - 14), f"{alt_d:g} m", fill="black", font=font_s, stroke_width=2, stroke_fill="white")

            if gecersiz:
                uyari = "Geçersiz aralık: " + ", ".join(gecersiz[:3])
                draw.text((x1 + 6, baslik_h + 6), uyari, fill="red", font=font_s)

        img.save(kayit_yolu, "JPEG", quality=90)
        return kayit_yolu

    def ciz_tarama_deseni(self, draw, zemin_tipi, x1, y1, x2, y2):
        # Arka planı beyaz yap
        draw.rectangle([x1, y1, x2, y2], fill="white")
        zemin_adi = str(zemin_tipi).upper().strip()
        
        # 1. BİTKİSEL TOPRAK (Ot Şekli - Siyah)
        if "BITKISEL" in zemin_adi or "BİTKİSEL" in zemin_adi or zemin_adi == "NEBATİ":
            for y in range(int(y1)+10, int(y2), 20):
                for x in range(int(x1)+10, int(x2)-10, 25):
                    # şaşırtmaca (offset)
                    ox = x if (y//20)%2==0 else x + 12
                    if ox > x2-10: continue
                    draw.line([ox, y, ox, y-6], fill="black", width=2)
                    draw.line([ox, y, ox-4, y-4], fill="black", width=2)
                    draw.line([ox, y, ox+4, y-4], fill="black", width=2)

        # 2. KİL - Şaşırtmalı Kesik Çizgiler (Gri)
        elif ("KİL" in zemin_adi or "KIL" in zemin_adi) and "TAŞI" not in zemin_adi and "TASI" not in zemin_adi:
            for y in range(int(y1)+10, int(y2), 12):
                offset = 0 if (y//12)%2 == 0 else 8
                for x in range(int(x1)+4+offset, int(x2), 16):
                    if x+8 < x2: draw.line([x, y, x+8, y], fill="grey", width=2)

        # 3. SİLT - Kesik ve Noktalar
        elif ("SİLT" in zemin_adi or "SILT" in zemin_adi) and "TAŞI" not in zemin_adi and "TASI" not in zemin_adi:
            for y in range(int(y1)+10, int(y2), 15):
                ol_x = int(x1)+5
                is_line = True if (y//15)%2 == 0 else False
                while ol_x < int(x2)-5:
                    if is_line:
                        if ol_x+10 < x2: draw.line([ol_x, y, ol_x+10, y], fill="grey", width=2)
                        ol_x += 15; is_line = False
                    else:
                        for _ in range(3):
                            if ol_x+2 < x2: draw.ellipse([ol_x, y-1, ol_x+2, y+1], fill="grey")
                            ol_x += 5
                        is_line = True

        # 4. KUM - Düzensiz Noktalar (Sarı)
        elif "KUM" in zemin_adi and not "TAŞI" in zemin_adi and not "TASI" in zemin_adi:
            rastgele = _deterministik_rng(zemin_adi, x1, y1, x2, y2)
            if int(x2) - int(x1) > 4 and int(y2) - int(y1) > 4:
                for _ in range(int((x2-x1)*(y2-y1)/80)):
                    px = rastgele.randint(int(x1)+2, int(x2)-2)
                    py = rastgele.randint(int(y1)+2, int(y2)-2)
                    draw.ellipse([px-1, py-1, px+1, py+1], fill="gold")

        # 5. ÇAKIL - İçi Boş Yuvarlaklar (Hardal)
        elif "ÇAKIL" in zemin_adi or "CAKIL" in zemin_adi:
            rastgele = _deterministik_rng(zemin_adi, x1, y1, x2, y2)
            if int(x2) - int(x1) > 8 and int(y2) - int(y1) > 8:
                for _ in range(int((x2-x1)*(y2-y1)/150)):
                    px = rastgele.randint(int(x1)+4, int(x2)-4)
                    py = rastgele.randint(int(y1)+4, int(y2)-4)
                    r = rastgele.randint(2, 4)
                    draw.ellipse([px-r, py-r, px+r, py+r], outline="darkgoldenrod", width=2)

        # 6. KİLTAŞI
        elif "KİLTAŞI" in zemin_adi or "KILTASI" in zemin_adi:
            for y in range(int(y1)+10, int(y2), 12):
                satir = (y//12)%3
                if satir == 0:
                    for x in range(int(x1)+4, int(x2), 16):
                        if x+8 < x2: draw.line([x, y, x+8, y], fill="grey", width=2)
                elif satir == 1:
                    ol_x = int(x1)+4
                    while ol_x < int(x2)-5:
                        if ol_x+8 < x2: draw.line([ol_x, y, ol_x+8, y], fill="grey", width=2)
                        ol_x += 12
                        for _ in range(3):
                            if ol_x < x2: draw.ellipse([ol_x-1, y-1, ol_x+1, y+1], fill="grey")
                            ol_x += 5
                elif satir == 2:
                    for x in range(int(x1)+12, int(x2), 16):
                        if x+8 < x2: draw.line([x, y, x+8, y], fill="grey", width=2)

        # 7. SİLTTAŞI (Turuncu Tonları)
        elif "SİLTTAŞI" in zemin_adi or "SILTTASI" in zemin_adi:
            renk = "darkorange"
            for y in range(int(y1)+10, int(y2), 12):
                satir = (y//12)%4
                if satir == 1 or satir == 3:
                    draw.line([x1, y, x2, y], fill=renk, width=1)
                elif satir == 0:
                    ol_x = int(x1)+5
                    while ol_x < int(x2):
                        if ol_x+12 < x2: draw.line([ol_x, y, ol_x+12, y], fill=renk, width=1)
                        ol_x += 16
                        for _ in range(2):
                            if ol_x < x2: draw.ellipse([ol_x-1, y-1, ol_x+1, y+1], fill=renk)
                            ol_x += 6
                elif satir == 2:
                    ol_x = int(x1)+5
                    while ol_x < int(x2):
                        for _ in range(2):
                            if ol_x < x2: draw.ellipse([ol_x-1, y-1, ol_x+1, y+1], fill=renk)
                            ol_x += 6
                        if ol_x+12 < x2: draw.line([ol_x, y, ol_x+12, y], fill=renk, width=1)
                        ol_x += 16

        # 8. ÇAMURTAŞI
        elif "ÇAMURTAŞI" in zemin_adi or "CAMURTASI" in zemin_adi:
            renk = "chocolate"
            for y in range(int(y1)+10, int(y2), 12):
                satir = (y//12)%4
                if satir == 1 or satir == 3:
                    draw.line([x1, y, x2, y], fill=renk, width=2)
                else:
                    offset = 0 if satir == 0 else 10
                    ol_x = int(x1)+offset
                    while ol_x < int(x2):
                        if ol_x+15 < x2: draw.line([ol_x, y, ol_x+15, y], fill=renk, width=2)
                        ol_x += 18
                        for _ in range(3):
                            if ol_x < x2: draw.ellipse([ol_x-1, y-1, ol_x+1, y+1], fill=renk)
                            ol_x += 6

        # 9. KUMTAŞI
        elif "KUMTAŞI" in zemin_adi or "KUMTASI" in zemin_adi:
            renk = "orange"
            for y in range(int(y1)+10, int(y2), 12):
                satir = (y//12)%2
                if satir == 1:
                    draw.line([x1, y, x2, y], fill=renk, width=2)
                else:
                    offset = 0 if (y//24)%2 == 0 else 5
                    for x in range(int(x1)+3+offset, int(x2), 10):
                        draw.ellipse([x-1, y-1, x+1, y+1], fill=renk)

        # 10. KİREÇTAŞI (Mavi Duvar Deseni)
        elif "KİREÇTAŞI" in zemin_adi or "KİRECTASI" in zemin_adi or "KİREÇ" in zemin_adi:
            renk = "dodgerblue"
            for y in range(int(y1)+15, int(y2), 15):
                draw.line([x1, y, x2, y], fill=renk, width=2) # Yatay Çizgi
                offset = 0 if (y//15)%2 == 0 else 15
                for x in range(int(x1)+offset, int(x2), 30):
                    draw.line([x, y, x, y-15], fill=renk, width=2) # Dikey Çizgiler

        # 11. GRANİT (Kırmızı Artı + Şekilleri)
        elif "GRANİT" in zemin_adi or "GRANIT" in zemin_adi:
            renk = "red"
            for y in range(int(y1)+12, int(y2), 20):
                offset = 0 if (y//20)%2 == 0 else 15
                for x in range(int(x1)+10+offset, int(x2)-10, 30):
                    draw.line([x-4, y, x+4, y], fill=renk, width=2)
                    draw.line([x, y-4, x, y+4], fill=renk, width=2)

        # 12. ANDEZİT (Mor Balık / Alfa Şekilleri)
        elif "ANDEZİT" in zemin_adi or "ANDEZIT" in zemin_adi:
            renk = "purple"
            for y in range(int(y1)+15, int(y2), 25):
                offset = 0 if (y//25)%2 == 0 else 15
                for x in range(int(x1)+10+offset, int(x2)-10, 30):
                    draw.arc([x-6, y-4, x, y+4], start=90, end=270, fill=renk, width=2)
                    draw.line([x, y-4, x+4, y], fill=renk, width=2)
                    draw.line([x, y+4, x+4, y], fill=renk, width=2)

        # 13. AGLOMERA (Pembe Çakıllar ve Alfa Karişimi)
        elif "AGLOMERA" in zemin_adi:
            renk = "hotpink"
            rastgele = _deterministik_rng(zemin_adi, x1, y1, x2, y2)
            if int(x2) - int(x1) > 10 and int(y2) - int(y1) > 10:
                for _ in range(int((x2-x1)*(y2-y1)/300)):
                    px = rastgele.randint(int(x1)+5, int(x2)-5)
                    py = rastgele.randint(int(y1)+5, int(y2)-5)
                    secim = rastgele.randint(1, 3)
                    if secim == 1:
                        r = rastgele.randint(3, 5)
                        draw.ellipse([px-r, py-r, px+r, py+r], outline=renk, width=3) # Çakıl
                    elif secim == 2:
                        draw.arc([px-5, py-3, px, py+3], start=90, end=270, fill=renk, width=2)
                        draw.line([px, py-3, px+3, py], fill=renk, width=2)
                        draw.line([px, py+3, px+3, py], fill=renk, width=2) # Alfa
                    else:
                        draw.ellipse([px-1, py-1, px+1, py+1], fill=renk) # Nokta

        # 14. BAZALT (Koyu Mor Beta)
        elif "BAZALT" in zemin_adi:
            renk = "indigo"
            for y in range(int(y1)+15, int(y2), 25):
                offset = 0 if (y//25)%2 == 0 else 15
                for x in range(int(x1)+10+offset, int(x2)-10, 30):
                    draw.line([x-2, y-6, x-2, y+6], fill=renk, width=2)
                    draw.arc([x-2, y-6, x+4, y], start=270, end=90, fill=renk, width=2)
                    draw.arc([x-2, y, x+4, y+6], start=270, end=90, fill=renk, width=2)

        else:
            pass # Tanımsız zemin ise sadece beyaz kutu kalır
        
        # Son Olarak Sütun Kutusu Sınırlarını Çiz
        draw.rectangle([x1, y1, x2, y2], outline="black", width=1)

