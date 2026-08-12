"""K-1 ile verilen 1/100.000 paftalar için doğrulanmış lejant sıra bilgileri."""

from __future__ import annotations

from pathlib import Path


HAZIR_PROFILLER = {
    "Ayvalik I16 ve J16": (
        ("Qal", "Alüvyon"),
        ("Tplb", "Bayramiç Formasyonu"),
        ("Tmplg", "Gülpınar Formasyonu"),
        ("Tçt", "Taştepe Bazaltı"),
        ("Tmçk", "Çanakkale Formasyonu"),
        ("Tmal", "Alçıtepe Üyesi"),
        ("Tmki", "Kirazlı Üyesi"),
        ("Tmçd", "Çamrakdere Üyesi"),
        ("Tmg", "Gazhanedere Formasyonu"),
        ("Tmi", "İlyasbaşı Formasyonu"),
        ("Tmbb", "Babadere Dasiti"),
        ("Tmç", "Çamkabalak İgnimbiriti"),
        ("Tmçt", "Tüf Üyesi"),
        ("Tmay", "Ayvacık Volkaniti"),
        ("Tmayt", "Tüf Üyesi"),
        ("Tmar", "Arıklı İgnimbiriti"),
        ("Tmhü", "Hüseyinfakı Volkaniti"),
        ("Tmk", "Küçükkuyu Formasyonu"),
        ("Tme", "Ezine Volkaniti"),
        ("Tmbe", "Behramkale Volkaniti"),
        ("Tmb", "Bademli Volkaniti"),
        ("Tmo", "Ortatepe Volkaniti"),
        ("Tmba", "Babakale Volkaniti"),
        ("Tma", "Araplar Volkaniti"),
        ("Tmh", "Halazadağ Volkaniti"),
        ("Toh", "Hallaçlar Volkaniti"),
        ("Tg", "Oligo-Miyosen Granitoyidleri"),
        ("Teç", "Ceylan Formasyonu"),
        ("Tes", "Soğucak Formasyonu"),
        ("Tef", "Fıçıtepe Formasyonu"),
        ("Kd", "Denizgören Ofiyoliti"),
        ("Kç", "Çetmi Melanjı"),
        ("Kça", "Çamlıca Metamorfitleri"),
        ("Kçam", "Mermer Üyesi"),
        ("Kças", "Metaserpantinit Üyesi"),
        ("Pç", "Çamköy Formasyonu"),
        ("Pb", "Bozalan Formasyonu"),
        ("Єg", "Geyikli Formasyonu"),
    ),
    "Ayvalik I17": (
        ("Qal", "Alüvyon"),
        ("Qym", "Yamaç Molozu"),
        ("plb", "Bayramiç Formasyonu"),
        ("Tmi", "İlyasbaşı Formasyonu"),
        ("Tmay", "Ayvacık Volkaniti"),
        ("Tmar", "Arıklı İgnimbiriti"),
        ("Tmhü", "Hüseyinfakı Volkaniti"),
        ("Tmça", "Çan Formasyonu"),
        ("Tmk", "Küçükkuyu Formasyonu"),
        ("Tme", "Ezine Volkaniti"),
        ("Toh", "Hallaçlar Volkaniti"),
        ("Tg", "Oligo-Miyosen Granitoyidleri"),
        ("Tos", "Saraycık Volkaniti"),
        ("Toy", "Yeniköy Volkaniti"),
        ("Ted", "Dededağ Volkanitleri"),
        ("Tedk", "Korudere İgnimbirit Üyesi"),
        ("Tedh", "Hacıbekirler Üyesi"),
        ("Teşa", "Şahinli Formasyonu"),
        ("Teşab", "Bilaller Üyesi"),
        ("Kç", "Çetmi Melanjı"),
        ("s", "Serpantinit Bloğu"),
        ("e", "Eklojit"),
        ("Tr", "Üst Triyas Yaşlı Kireçtaşı Bloğu"),
        ("kçb", "Tanımlanmamış Kireçtaşı Bloğu"),
        ("m", "Mermer Bloğu"),
        ("Kça", "Çamlıca Metamorfitleri"),
        ("Kçap", "Palamut Fillit Üyesi"),
        ("Kças", "Serpantinit Üyesi"),
        ("JKb", "Bilecik Formasyonu"),
        ("Tkk", "Karakaya Formasyonu (Ayrılmamış)"),
        ("Tkm", "Mehmetalan Formasyonu"),
        ("pk", "Permiyen Yaşlı Kireçtaşı Bloğu"),
        ("Pzs", "Sazak Formasyonu"),
        ("Pzsm", "Mermer Üyesi"),
        ("Pzt", "Torasan Formasyonu"),
        ("AMZ", "Alakeçili Milonit Zonu"),
        ("Cs", "Sütüven Formasyonu"),
        ("Csm", "Mermer Üyesi"),
        ("Gg", "Granitik Gnays"),
        ("Mzs", "Sarıkız Mermeri"),
        ("Mzt", "Tozlu Formasyonu"),
        ("Tf", "Fındıklı Formasyonu"),
        ("Trfb", "Babadağ Mermer Üyesi"),
        ("Trfa", "Altınoluk Mermer Üyesi"),
    ),
    "Balikesir I18": (
        ("Qal", "Alüvyon"),
        ("Tmi", "İlyasbaşı Formasyonu"),
        ("Tmso", "Soma Formasyonu"),
        ("Tme", "Ezine Volkaniti"),
        ("Tmş", "Şapçı Volkaniti"),
        ("Tmy", "Yürekli Dasiti"),
        ("Tmk", "Küçükkuyu Formasyonu"),
        ("Toh", "Hallaçlar Volkaniti"),
        ("Tob", "Bağburun Formasyonu"),
        ("Tg", "Oligo-Miyosen Granitoyidleri"),
        ("Tu", "Erdağ Volkaniti"),
        ("Teşa", "Şahinli Formasyonu"),
        ("Teşab", "Bilaller Üyesi"),
        ("Kç", "Çetmi Melanjı"),
        ("s", "Serpantinit Bloğu"),
        ("Kp", "Pınar Formasyonu"),
        ("JKb", "Bilecik Formasyonu"),
        ("Jba", "Bayırköy Formasyonu"),
        ("Trb", "Balya Formasyonu"),
        ("Trkk", "Karakaya Formasyonu (Ayrılmamış)"),
        ("Trkc", "Camialan Kireçtaşı"),
        ("Trkç", "Çal Formasyonu"),
        ("Trkm", "Mehmetalan Formasyonu"),
        ("Trko", "Orhanlar Grovağı"),
        ("Trka", "Arkozik Kumtaşı"),
        ("pk", "Permiyen Yaşlı Kireçtaşı Bloğu"),
        ("C", "Karbonifer Yaşlı Kireçtaşı Bloğu"),
        ("kb", "Tanımlanmamış Kireçtaşı Blokları"),
        ("Pzs", "Sazak Formasyonu"),
        ("Pzsm", "Mermer Üyesi"),
        ("Pzt", "Torasan Formasyonu"),
        ("Pzçm", "Mermer Üyesi"),
        ("Pzts", "Metaserpantinit Üyesi"),
        ("Pzç", "Çamlık Metagranodiyoriti"),
        ("Cs", "Sütüven Formasyonu"),
        ("Gg", "Granitik Gnays"),
        ("Csm", "Mermer Üyesi"),
        ("Tf", "Fındıklı Formasyonu"),
        ("Trfa", "Altınoluk Mermer Üyesi"),
    ),
    "Canakkale H17": (
        ("Qal", "Alüvyon"),
        ("Qea", "Eski Akarsu Çökelleri"),
        ("Tplb", "Bayramiç Formasyonu"),
        ("Tmçk", "Çanakkale Formasyonu"),
        ("Tmal", "Alçıtepe Üyesi"),
        ("Tmçd", "Çamrakdere Üyesi"),
        ("Tmki", "Kirazlı Üyesi"),
        ("Tmg", "Gazhanedere Formasyonu"),
        ("Tmis", "Işıkeli Riyoliti"),
        ("Tmça", "Çan Formasyonu"),
        ("Tme", "Ezine Volkaniti"),
        ("Toh", "Hallaçlar Volkaniti"),
        ("Tg", "Oligo-Miyosen Granitoyidleri"),
        ("Toa", "Atikhisar Volkaniti"),
        ("Toy", "Yeniköy Volkaniti"),
        ("Tom", "Mezardere Formasyonu"),
        ("Tomk", "Kanlıbent Üyesi"),
        ("Teer", "Erdağ Volkaniti"),
        ("Tebe", "Beybaşlı Formasyonu"),
        ("Teç", "Ceylan Formasyonu"),
        ("Teck", "Korudağ Üyesi"),
        ("Ted", "Dededağ Volkanitleri"),
        ("Tedka", "Kazmalı Tüf Üyesi"),
        ("Tedk", "Korudere İgnimbirit Üyesi"),
        ("Tedh", "Hacıbekirler Üyesi"),
        ("Tes", "Soğucak Formasyonu"),
        ("Teşa", "Şahinli Formasyonu"),
        ("Teşab", "Bilaller Üyesi"),
        ("Tef", "Fıçıtepe Formasyonu"),
        ("Teb", "Beyçayır Volkaniti"),
        ("Teg", "Eosen Granitoyidleri"),
        ("Kç", "Çetmi Melanjı"),
        ("Kd", "Denizgören Ofiyoliti"),
        ("Kça", "Çamlıca Metamorfitleri"),
        ("Kçap", "Palamut Fillit Üyesi"),
    ),
    "Bandirma H19": (
        ("Qal", "Alüvyon"),
        ("Tplb", "Bayramiç Formasyonu"),
        ("Tmş", "Şapçı Volkaniti"),
        ("Toh", "Hallaçlar Volkaniti"),
        ("Tg", "Oligo-Miyosen Granitoyidleri"),
        ("Tee", "Edincik Volkaniti"),
        ("Teg", "Eosen Granitoyidleri"),
        ("Ky", "Yayla Melanjı"),
        ("Kp", "Pınar Formasyonu"),
        ("JKb", "Bilecik Formasyonu"),
        ("Jba", "Bayırköy Formasyonu"),
        ("Tb", "Balya Formasyonu"),
        ("Tkk", "Karakaya Formasyonu (Ayrılmamış)"),
        ("Tkc", "Camialan Kireçtaşı"),
        ("Tkm", "Mehmetalan Formasyonu"),
        ("Pk", "Permiyen Yaşlı Kireçtaşı Bloğu"),
        ("C", "Karbonifer Yaşlı Kireçtaşı Bloğu"),
        ("Pzs", "Sazak Formasyonu"),
        ("Pzsm", "Mermer Üyesi"),
        ("Pzt", "Torasan Formasyonu"),
        ("Tf", "Fındıklı Formasyonu"),
        ("Trfa", "Altınoluk Mermer Üyesi"),
    ),
    "Canakkale H15-H16": (
        ("Qal", "Alüvyon"),
        ("Tplb", "Bayramiç Formasyonu"),
        ("Tmçk", "Çanakkale Formasyonu"),
        ("Tmal", "Alçıtepe Üyesi"),
        ("Tmçd", "Çamrakdere Üyesi"),
        ("Tmki", "Kirazlı Üyesi"),
        ("Tmg", "Gazhanedere Formasyonu"),
        ("Tmay", "Ayvacık Volkaniti"),
        ("Toma", "Armutburnu Formasyonu"),
        ("Toa", "Atikhisar Volkaniti"),
        ("Toy", "Yeniköy Volkaniti"),
        ("Tom", "Mezardere Formasyonu"),
        ("Tomk", "Kanlıbent Üyesi"),
        ("Teç", "Ceylan Formasyonu"),
        ("Teck", "Korudağ Üyesi"),
        ("Tes", "Soğucak Formasyonu"),
        ("Teşa", "Şahinli Formasyonu"),
        ("Tef", "Fıçıtepe Formasyonu"),
        ("Tek", "Karaağaç Formasyonu"),
        ("Kpl", "Lört Formasyonu"),
        ("Kd", "Denizgören Ofiyoliti"),
        ("Kça", "Çamlıca Metamorfitleri"),
    ),
    "Bandirma H18": (
        ("Qal", "Alüvyon"),
        ("Qym", "Yamaç Molozu"),
        ("plb", "Bayramiç Formasyonu"),
        ("Tmis", "Işıkeli Riyoliti"),
        ("Tmça", "Çan Formasyonu"),
        ("Tmş", "Şapçı Volkaniti"),
        ("Toh", "Hallaçlar Volkaniti"),
        ("Tg", "Oligo-Miyosen Granitoyidleri"),
        ("Toa", "Atikhisar Volkaniti"),
        ("Teç", "Ceylan Formasyonu"),
        ("Ted", "Dededağ Volkanitleri"),
        ("Tedka", "Kazmalı Tüf Üyesi"),
        ("Tes", "Soğucak Formasyonu"),
        ("Teşa", "Şahinli Formasyonu"),
        ("Teb", "Beyçayır Volkaniti"),
        ("Teg", "Eosen Granitoyidleri"),
        ("Kba", "Balıkkaya Formasyonu"),
        ("JK", "Jura-Kretase Yaşlı Kireçtaşı Bloğu"),
        ("Pk", "Permiyen Yaşlı Kireçtaşı Bloğu"),
        ("kb", "Tanımlanmamış Kireçtaşı Bloğu"),
        ("Kç", "Çetmi Melanjı"),
        ("s", "Serpantinit Bloğu"),
        ("Tr", "Triyas Kireçtaşı Bloğu"),
        ("m", "Mermer Bloğu"),
        ("Kça", "Çamlıca Metamorfitleri"),
        ("Kçap", "Palamut Fillit Üyesi"),
        ("Kçam", "Mermer Üyesi"),
        ("JKb", "Bilecik Formasyonu"),
        ("Jba", "Bayırköy Formasyonu"),
        ("Tkk", "Karakaya Formasyonu (Ayrılmamış)"),
        ("Tkc", "Camialan Kireçtaşı"),
        ("Tkç", "Çal Formasyonu"),
        ("Tkm", "Mehmetalan Formasyonu"),
        ("Tka", "Arkozik Kumtaşı"),
        ("Pk", "Permiyen Yaşlı Kireçtaşı Bloğu"),
        ("Pzs", "Sazak Formasyonu"),
        ("Pzt", "Torasan Formasyonu"),
        ("Pzçm", "Mermer Üyesi"),
        ("Pzts", "Metaserpantinit Üyesi"),
        ("Pzç", "Çamlık Metagranodiyoriti"),
    ),
}


def _anahtar(value):
    import unicodedata

    text = unicodedata.normalize("NFKD", str(value or "").casefold())
    text = "".join(char for char in text if not unicodedata.combining(char)).replace("ı", "i")
    return " ".join("".join(char if char.isalnum() else " " for char in text).split())


def _gorsel_oku(path):
    try:
        import cv2
        import numpy as np
    except Exception:
        return None
    try:
        return cv2.imdecode(np.fromfile(str(path), dtype=np.uint8), cv2.IMREAD_COLOR)
    except Exception:
        return None


def _ayni_kutulari_birlestir(rects, width, height):
    result = []
    for rect in sorted(rects, key=lambda item: (item[0], item[1], -(item[2] * item[3]))):
        x, y, w, h = rect
        center_x, center_y = x + w / 2, y + h / 2
        if any(
            abs(center_x - (a + c / 2)) < width * 0.004
            and abs(center_y - (b + d / 2)) < height * 0.004
            for a, b, c, d in result
        ):
            continue
        result.append(rect)
    return result


def _lejant_kutularini_bul(image, expected_count):
    import cv2
    import numpy as np

    height, width = image.shape[:2]
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    candidates = []
    for threshold in (130, 165, 195, 220):
        binary = cv2.threshold(gray, threshold, 255, cv2.THRESH_BINARY_INV)[1]
        contours, _ = cv2.findContours(binary, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
        for contour in contours:
            x, y, w, h = cv2.boundingRect(contour)
            if not (width * 0.48 < x < width * 0.92 and height * 0.08 < y < height * 0.92):
                continue
            if not (width * 0.007 < w < width * 0.055 and height * 0.004 < h < height * 0.035):
                continue
            if 1.0 < w / max(h, 1) < 6.5:
                candidates.append((x, y, w, h))
    candidates = _ayni_kutulari_birlestir(candidates, width, height)

    clusters = []
    for rect in sorted(candidates, key=lambda item: item[0] + item[2] / 2):
        center = rect[0] + rect[2] / 2
        cluster = next((item for item in clusters if abs(item[0] - center) < width * 0.018), None)
        if cluster is None:
            cluster = [center, []]
            clusters.append(cluster)
        cluster[1].append(rect)
        cluster[0] = sum(item[0] + item[2] / 2 for item in cluster[1]) / len(cluster[1])
    eligible = [item for item in clusters if len(item[1]) >= min(12, expected_count)]
    if not eligible:
        return []
    chosen = max(eligible, key=lambda item: item[0])

    groups = []
    tolerance = height * 0.006
    for rect in sorted(chosen[1], key=lambda item: item[1] + item[3] / 2):
        center_y = rect[1] + rect[3] / 2
        group = next((item for item in groups if abs(item[0] - center_y) < tolerance), None)
        if group is None:
            group = [center_y, []]
            groups.append(group)
        group[1].append(rect)
        group[0] = sum(item[1] + item[3] / 2 for item in group[1]) / len(group[1])
    median_width = float(np.median([item[2] for item in chosen[1]]))
    rows = []
    for _center, group in groups:
        plausible = [item for item in group if item[2] >= median_width * 0.72]
        rows.append(max(plausible or group, key=lambda item: item[2] * item[3]))
    return rows[:expected_count]


def hazir_profil_ogeleri(jpeg_path):
    path = Path(jpeg_path)
    wanted = _anahtar(path.stem)
    profile_name = next((name for name in HAZIR_PROFILLER if _anahtar(name) == wanted), None)
    if profile_name is None:
        return []
    definitions = HAZIR_PROFILLER[profile_name]
    image = _gorsel_oku(path)
    if image is None:
        return []
    height, width = image.shape[:2]
    rows = _lejant_kutularini_bul(image, len(definitions))
    if len(rows) != len(definitions):
        return []
    result = []
    for index, ((code, name), (x, y, w, h)) in enumerate(zip(definitions, rows), start=1):
        inset_x = max(1, round(w * 0.025))
        inset_y = max(1, round(h * 0.04))
        result.append(
            {
                "sira": index,
                "kod": code,
                "ad": name,
                "rect": [
                    (x + inset_x) / width,
                    (y + inset_y) / height,
                    (x + w - inset_x) / width,
                    (y + h - inset_y) / height,
                ],
            }
        )
    return result


__all__ = ["HAZIR_PROFILLER", "hazir_profil_ogeleri"]
