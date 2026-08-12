import re
import unicodedata

from docx.oxml.ns import qn


WD_FIELD_EMPTY = -1
WD_STYLE_CAPTION = -35

_ETIKETLER = {
    "sekil": "Şekil",
    "tablo": "Tablo",
    "cizelge": "Çizelge",
}


def _metin_normalize(metin):
    metin = unicodedata.normalize("NFKD", str(metin or ""))
    metin = "".join(ch for ch in metin if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]+", "_", metin.casefold().replace("ı", "i")).strip("_")


def _etiket_normalize(etiket):
    return _ETIKETLER.get(_metin_normalize(etiket), str(etiket or "").strip())


def _baslik_etiketi(metin):
    eslesme = re.match(
        r"^\s*(Şekil|Sekil|Tablo|Çizelge|Cizelge)\b",
        str(metin or "").rstrip("\r\n\x07"),
        re.IGNORECASE,
    )
    if not eslesme:
        return None, None
    return _etiket_normalize(eslesme.group(1)), eslesme


def _alan_seq_etiketi(kod):
    eslesme = re.search(
        r"\bSEQ\s+(?:\"([^\"]+)\"|([^\s\\]+))",
        str(kod or ""),
        re.IGNORECASE,
    )
    if not eslesme:
        return ""
    return _etiket_normalize(eslesme.group(1) or eslesme.group(2))


def _alan_ref_mi(kod):
    return bool(re.search(r"\bREF\s+", str(kod or ""), re.IGNORECASE))


def _word_stili_resim_yazisi_mi(paragraf):
    try:
        stil = paragraf.Range.Style
        stil_adi = getattr(stil, "NameLocal", stil)
    except Exception:
        return False
    normal = _metin_normalize(stil_adi)
    return "caption" in normal or "resim_yaz" in normal


def _docx_stili_resim_yazisi_mi(paragraf):
    stil = getattr(paragraf, "style", None)
    normal = _metin_normalize(
        " ".join(
            str(deger or "")
            for deger in (
                getattr(stil, "style_id", ""),
                getattr(stil, "name", ""),
            )
        )
    )
    return "caption" in normal or "resim_yaz" in normal


def _word_paragraf_alanlari(paragraf):
    alanlar = []
    try:
        for index in range(1, paragraf.Range.Fields.Count + 1):
            alanlar.append(paragraf.Range.Fields(index))
    except Exception:
        pass
    return alanlar


def _word_bookmark_adlari(paragraf):
    adlar = []
    try:
        for index in range(1, paragraf.Range.Bookmarks.Count + 1):
            ad = str(paragraf.Range.Bookmarks(index).Name or "")
            if ad:
                adlar.append(ad)
    except Exception:
        pass
    return list(dict.fromkeys(adlar))


def _word_bookmark_var_mi(belge, ad):
    try:
        return bool(belge.Bookmarks.Exists(ad))
    except Exception:
        try:
            belge.Bookmarks(ad)
            return True
        except Exception:
            return False


def _word_caption_yeniden_kur(
    belge,
    paragraf,
    etiket,
    baslik,
    sayi_bookmarklari=(),
):
    for ad in sayi_bookmarklari:
        if _word_bookmark_var_mi(belge, ad):
            try:
                belge.Bookmarks(ad).Delete()
            except Exception:
                pass

    icerik = paragraf.Range.Duplicate
    if icerik.End > icerik.Start:
        icerik.End -= 1
    icerik.Text = f"{etiket}  {baslik}"
    ekleme_konumu = paragraf.Range.Start + len(etiket) + 1
    yeni_alan = belge.Fields.Add(
        belge.Range(Start=ekleme_konumu, End=ekleme_konumu),
        WD_FIELD_EMPTY,
        f"SEQ {etiket} \\* ARABIC",
        True,
    )
    yeni_alan.Locked = False
    yeni_alan.Update()
    for ad in sayi_bookmarklari:
        try:
            belge.Bookmarks.Add(ad, yeni_alan.Result.Duplicate)
        except Exception:
            pass
    return yeni_alan


def _word_caption_duzelt(word, belge, paragraf):
    metin = str(paragraf.Range.Text or "").rstrip("\r\n\x07")
    eski_genel_jeoloji_caption = (
        _metin_normalize(metin)
        == "genel_jeoloji_haritasi_olcek_1_100_000"
    )
    if not _word_stili_resim_yazisi_mi(paragraf) and not eski_genel_jeoloji_caption:
        return False
    if eski_genel_jeoloji_caption:
        try:
            paragraf.Range.Style = WD_STYLE_CAPTION
        except Exception:
            pass
        _word_caption_yeniden_kur(
            belge,
            paragraf,
            "Şekil",
            "Genel jeoloji haritası (Ölçek 1/100.000)",
        )
        return True

    etiket, etiket_eslesmesi = _baslik_etiketi(metin)
    if not etiket_eslesmesi:
        return False

    # Eski şablonda koordinat tablosu yanlışlıkla "Şekil" dizisine ve bazı
    # çıktılarda başka bir tablonun REF alanına bağlıydı.
    numune_koordinat_basligi = (
        "numune_lokasyonlari_koordinatlari" in _metin_normalize(metin)
    )
    if numune_koordinat_basligi:
        etiket = "Tablo"

    alanlar = _word_paragraf_alanlari(paragraf)
    seq_alanlari = []
    for alan in alanlar:
        kod = str(alan.Code.Text or "")
        if _alan_seq_etiketi(kod):
            seq_alanlari.append(alan)
    if len(seq_alanlari) > 1:
        raise ValueError(f"Bir resim yazısında birden fazla SEQ alanı bulundu: {metin[:120]}")

    bookmark_adlari = _word_bookmark_adlari(paragraf)
    sayi_bookmarklari = []
    for ad in bookmark_adlari:
        try:
            bookmark_metni = str(belge.Bookmarks(ad).Range.Text or "").strip()
        except Exception:
            bookmark_metni = ""
        if ad.startswith("K1_") or re.fullmatch(r"\d+", bookmark_metni):
            sayi_bookmarklari.append(ad)

    ref_alani_var = any(
        _alan_ref_mi(str(alan.Code.Text or ""))
        for alan in alanlar
    )
    if numune_koordinat_basligi and (
        len(seq_alanlari) != 1
        or _alan_seq_etiketi(str(seq_alanlari[0].Code.Text or "")) != "Tablo"
        or ref_alani_var
    ):
        _word_caption_yeniden_kur(
            belge,
            paragraf,
            "Tablo",
            "Numune Lokasyonları Koordinatları",
            sayi_bookmarklari,
        )
        return True

    degisti = False
    mevcut_etiket = _etiket_normalize(etiket_eslesmesi.group(1))
    if mevcut_etiket != etiket:
        etiket_araligi = belge.Range(
            Start=paragraf.Range.Start + etiket_eslesmesi.start(1),
            End=paragraf.Range.Start + etiket_eslesmesi.end(1),
        )
        etiket_araligi.Text = etiket
        degisti = True
        metin = str(paragraf.Range.Text or "").rstrip("\r\n\x07")
        _, etiket_eslesmesi = _baslik_etiketi(metin)

    if seq_alanlari:
        alan = seq_alanlari[0]
        yeni_kod = f" SEQ {etiket} \\* ARABIC "
        if " ".join(str(alan.Code.Text or "").split()) != " ".join(yeni_kod.split()):
            alan.Code.Text = yeni_kod
            degisti = True
        alan.Locked = False
        alan.Update()
        return degisti

    # Eski numune koordinat başlığında başka bir tablonun REF alanı başlık
    # numarası gibi kullanılmıştı. Sadece başın içindeki REF alanını metne çevir.
    for alan in reversed(alanlar):
        kod = str(alan.Code.Text or "")
        if _alan_ref_mi(kod) and alan.Result.Start <= paragraf.Range.Start + 24:
            alan.Unlink()
            degisti = True

    for ad in sayi_bookmarklari:
        if _word_bookmark_var_mi(belge, ad):
            try:
                belge.Bookmarks(ad).Delete()
            except Exception:
                pass

    metin = str(paragraf.Range.Text or "").rstrip("\r\n\x07")
    _, etiket_eslesmesi = _baslik_etiketi(metin)
    if not etiket_eslesmesi:
        raise ValueError(f"Resim yazısının etiketi okunamadı: {metin[:120]}")

    etiket_bitis = etiket_eslesmesi.end(1)
    sayi_eslesmesi = re.match(r"\s*(\d+)\b", metin[etiket_bitis:])
    if sayi_eslesmesi:
        sayi_baslangic = etiket_bitis + sayi_eslesmesi.start(1)
        sayi_bitis = etiket_bitis + sayi_eslesmesi.end(1)
        sayi_araligi = belge.Range(
            Start=paragraf.Range.Start + sayi_baslangic,
            End=paragraf.Range.Start + sayi_bitis,
        )
        ekleme_konumu = sayi_araligi.Start
        sayi_araligi.Text = ""
    else:
        bosluk_eslesmesi = re.match(r"\s*", metin[etiket_bitis:])
        bosluk_bitis = etiket_bitis + len(bosluk_eslesmesi.group(0))
        bosluk_araligi = belge.Range(
            Start=paragraf.Range.Start + etiket_bitis,
            End=paragraf.Range.Start + bosluk_bitis,
        )
        bosluk_araligi.Text = " "
        ekleme_konumu = bosluk_araligi.End
        sonraki = str(paragraf.Range.Text or "").rstrip("\r\n\x07")
        sonraki_karakter = sonraki[ekleme_konumu - paragraf.Range.Start:][:1]
        if sonraki_karakter and sonraki_karakter not in ".,:;-–—)":
            belge.Range(Start=ekleme_konumu, End=ekleme_konumu).Text = " "

    alan_araligi = belge.Range(Start=ekleme_konumu, End=ekleme_konumu)
    yeni_alan = belge.Fields.Add(
        alan_araligi,
        WD_FIELD_EMPTY,
        f"SEQ {etiket} \\* ARABIC",
        True,
    )
    yeni_alan.Locked = False
    yeni_alan.Update()

    for ad in sayi_bookmarklari:
        try:
            belge.Bookmarks.Add(ad, yeni_alan.Result.Duplicate)
        except Exception:
            pass
    return True


def word_baslik_numaralarini_normallestir(word, belge):
    """Caption paragraflarını tek bir SEQ alanına bağla ve sabit sayıları kaldır."""
    try:
        belge.Bookmarks.ShowHidden = True
    except Exception:
        pass
    degisen = 0
    for index in range(1, belge.Paragraphs.Count + 1):
        if _word_caption_duzelt(word, belge, belge.Paragraphs(index)):
            degisen += 1
    return degisen


def _docx_paragraf_gorunen_metni(paragraf):
    return "".join(node.text or "" for node in paragraf._p.xpath(".//w:t"))


def _docx_paragraf_alan_kodlari(paragraf):
    kodlar = [str(node.text or "").strip() for node in paragraf._p.xpath(".//w:instrText")]
    for node in paragraf._p.xpath(".//w:fldSimple"):
        kod = str(node.get(qn("w:instr")) or "").strip()
        if kod:
            kodlar.append(kod)
    return [kod for kod in kodlar if kod]


def docx_baslik_numaralandirma_hatalari(doc):
    """Son Word çıktısındaki sabit, yanlış ve sıra dışı şekil/tablo numaralarını bul."""
    hatalar = []
    siralar = {"Şekil": [], "Tablo": [], "Çizelge": []}
    for index, paragraf in enumerate(doc.paragraphs, 1):
        if not _docx_stili_resim_yazisi_mi(paragraf):
            continue
        gorunen = _docx_paragraf_gorunen_metni(paragraf).strip()
        etiket, _ = _baslik_etiketi(gorunen)
        if not etiket:
            continue
        if "numune_lokasyonlari_koordinatlari" in _metin_normalize(gorunen) and etiket != "Tablo":
            hatalar.append(f"{index}. paragraftaki koordinat başlığı Tablo olmalıdır.")

        kodlar = _docx_paragraf_alan_kodlari(paragraf)
        seq_etiketleri = [deger for deger in (_alan_seq_etiketi(kod) for kod in kodlar) if deger]
        if len(seq_etiketleri) != 1:
            hatalar.append(
                f"{index}. paragraftaki '{gorunen[:80]}' başlığında tam bir SEQ alanı bulunmalıdır."
            )
            continue
        if seq_etiketleri[0] != etiket:
            hatalar.append(
                f"{index}. paragraftaki {etiket} başlığı yanlış {seq_etiketleri[0]} sırasına bağlı."
            )
            continue

        sayi_eslesmesi = re.match(
            rf"^\s*{re.escape(etiket)}\s+(\d+)\b",
            gorunen,
            re.IGNORECASE,
        )
        if not sayi_eslesmesi:
            hatalar.append(f"{index}. paragraftaki {etiket} numarası Word tarafından güncellenmemiş.")
            continue
        siralar.setdefault(etiket, []).append(int(sayi_eslesmesi.group(1)))

    for etiket, bulunan in siralar.items():
        if not bulunan:
            continue
        beklenen = list(range(1, len(bulunan) + 1))
        if bulunan != beklenen:
            hatalar.append(
                f"{etiket} numaraları sıralı değil; bulunan: {bulunan}, beklenen: {beklenen}."
            )
    return hatalar
