"""DOCX şablonlarındaki kişisel/makine verilerini denetle ve temizle.

Bu yardımcı, şablon düzenini mümkün olduğunca korumak için OOXML metin
düğümlerini yerinde günceller. Varsayılan çalışma salt-okunur denetimdir;
``--apply`` verilmeden hiçbir dosya değiştirilmez.
"""

from __future__ import annotations

import argparse
import os
import re
import tempfile
import zipfile
from pathlib import Path

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC_NS = "http://purl.org/dc/elements/1.1/"

HASSAS_DENETIM = re.compile(
    r"(?:[A-Z]:\\|file:/+|\\\\|\b(?:tel(?:efon)?|adres|sicil|mersis)\b|"
    r"\b\d{3}[ .()-]*\d{3}[ .-]*\d{2}[ .-]*\d{2}\b|"
    r"\b(?:mühendis|müh\.|mim\.|inş\.|ltd\.|şti\.)\b|"
    r"\b(?:ada|parsel|pafta)\b)",
    re.IGNORECASE,
)

# Değerler yalnız mevcut örnek şablonların anonimleştirilmesi içindir. Yeni
# projelerde gerçek bilgi ilgili proje/taahhüt alanlarından gelir.
RAPOR_DEGISIMLERI = (
    (re.compile(r"UB\s+ZEM[Iİ]N\s+M[ÜU]H\.?\s*M[Iİ]M\.?\s*[İI]N[ŞS]\.?\s*SAN\.?\s*T[Iİ]C\.?\s*LTD\.?\s*[ŞS]T[Iİ]\.?", re.I), "[FIRMA_ADI]"),
    (re.compile(r"UB\s+ZEM[Iİ]N\s+M[ÜU]HEND[Iİ]SL[Iİ]K", re.I), "[FIRMA_ADI]"),
    (re.compile(r"[İI]smetpa[şs]a\s+M(?:ah|h)\.?[^\n]{0,120}?Çanakkale", re.I), "[FIRMA_ADRESI]"),
    (re.compile(r"Tel\s*:\s*\(?286\)?\s*213\s*06\s*69", re.I), "[FIRMA_TELEFON]"),
    (re.compile(r"Mersis\s+No\s*:[^\n]{1,80}", re.I), ""),
    (re.compile(r"Hasan\s+ALANYALI", re.I), "[PROJE_ADI]"),
    (re.compile(r"G[öo]kalp\s+DO[ĞG]AN", re.I), "[JEOLOJI_MUH_AD]"),
    (re.compile(r"Suat\s+ERG[ÜU]N", re.I), "[JEOFIZIK_MUH_AD]"),
    (re.compile(r"Oda\s+Sicil\s+No\s*:\s*7400", re.I), "Oda Sicil No: [JEOLOJI_MUH_SICIL]"),
    (re.compile(r"Oda\s+Sicil\s+No\s*:\s*1982", re.I), "Oda Sicil No: [JEOFIZIK_MUH_SICIL]"),
    (
        re.compile(r"İmar\s+Bilgileri\s*:\s*Çanakkale\s+İli,.+?\b2\s+Parsel", re.I),
        "İmar Bilgileri: [IL] İli, [ILCE] İlçesi, [KOY] Köyü, [PAFTA] Pafta, [ADA] Ada, [PARSEL] Parsel",
    ),
    (re.compile(r"\[IL\]\s+İl\s+Özel\s+İdaresi", re.I), "[ILGILI_IDARE]"),
    (re.compile(r"Aral[ıi]k\s+2025", re.I), "[RAPOR_TARIHI]"),
)

TAAHHUT_DEGISIMLERI = (
    (re.compile(r"G[öo]kalp\s+DO[ĞG]AN", re.I), "[JEOLOJI_MUH_AD]"),
    (re.compile(r"Cemal\s+Bu[ğg]ra\s+[ŞS]ENEL", re.I), "[JEOFIZIK_MUH_AD]"),
    (re.compile(r"Suat\s+ERG[ÜU]N", re.I), "[JEOFIZIK_MUH_AD]"),
    (re.compile(r"\b14\.06\.2026\b"), "[TARIH]"),
)

RAPOR_HASSAS_PARCA_ADLARI = {
    "word/media/image1.jpeg",
    "word/media/image2.jpg",
    "word/media/image3.png",
    "word/media/image5.png",
    "word/media/image7.jpg",
    "word/embeddings/oleObject1.bin",
}


def _xml_mi(ad: str) -> bool:
    return ad.endswith((".xml", ".rels"))


def _metinler(xml: bytes) -> list[str]:
    try:
        kok = etree.fromstring(xml)
    except etree.XMLSyntaxError:
        return []
    return [
        "".join(dugum.itertext()).strip()
        for dugum in kok.xpath(".//*[local-name()='p']")
        if not dugum.xpath(".//*[local-name()='p']")
    ]


def denetle(yol: Path) -> list[str]:
    bulgular: list[str] = []
    with zipfile.ZipFile(yol) as paket:
        for ad in paket.namelist():
            if not _xml_mi(ad):
                continue
            veri = paket.read(ad)
            for metin in _metinler(veri):
                if HASSAS_DENETIM.search(metin):
                    bulgular.append(f"{ad}: {metin}")
            duz = veri.decode("utf-8", "ignore")
            if re.search(r"(?:[A-Z]:\\|file:/+|\\\\)", duz, re.I):
                bulgular.append(f"{ad}: yerel/harici dosya yolu")
    return bulgular


def _metni_dagit(dugumler: list[etree._Element], yeni: str) -> None:
    kalan = yeni
    for dugum in dugumler[:-1]:
        eski_uzunluk = len(dugum.text or "")
        dugum.text, kalan = kalan[:eski_uzunluk], kalan[eski_uzunluk:]
    dugumler[-1].text = kalan


def _metni_degistir(kok: etree._Element, degisimler) -> int:
    sayi = 0
    paragraflar = kok.xpath(".//*[local-name()='p']")
    for paragraf in paragraflar:
        if paragraf.xpath(".//*[local-name()='p']"):
            continue
        metin_dugumleri = paragraf.xpath(".//*[local-name()='t']")
        if not metin_dugumleri:
            continue
        eski = "".join(dugum.text or "" for dugum in metin_dugumleri)
        yeni = eski
        for desen, karsilik in degisimler:
            yeni, adet = desen.subn(karsilik, yeni)
            sayi += adet
        if yeni != eski:
            _metni_dagit(metin_dugumleri, yeni)
    return sayi


def _metadata_temizle(kok: etree._Element) -> None:
    for xpath in (
        ".//*[local-name()='creator']",
        ".//*[local-name()='lastModifiedBy']",
        ".//*[local-name()='company']",
        ".//*[local-name()='manager']",
    ):
        for dugum in kok.xpath(xpath):
            dugum.text = ""
    for dugum in kok.iter():
        for nitelik in list(dugum.attrib):
            if etree.QName(nitelik).localname.lower() == "gfxdata":
                del dugum.attrib[nitelik]
        for nitelik in list(dugum.attrib):
            if etree.QName(nitelik).localname.lower().startswith("rsid"):
                del dugum.attrib[nitelik]


def _iliski_yollarini_temizle(kok: etree._Element) -> None:
    for iliski in kok.xpath(".//*[local-name()='Relationship']"):
        hedef = iliski.get("Target", "")
        if re.match(r"(?i)(?:file:/+|[A-Z]:\\|\\\\)", hedef):
            iliski.set("Target", "about:blank")
            iliski.set("TargetMode", "External")


def _icerik_turlerinden_custom_sil(kok: etree._Element) -> None:
    for dugum in list(kok):
        parca = dugum.get("PartName", "")
        tur = dugum.get("ContentType", "")
        if parca == "/docProps/custom.xml" or "custom-properties" in tur:
            kok.remove(dugum)


def _kok_iliskilerinden_custom_sil(kok: etree._Element) -> None:
    for dugum in list(kok):
        if dugum.get("Type", "").endswith("/custom-properties"):
            kok.remove(dugum)


def _iliski_sahibi(rels_adi: str) -> str | None:
    if rels_adi == "_rels/.rels":
        return None
    on_ek, ayirici, son = rels_adi.rpartition("/_rels/")
    if not ayirici or not son.endswith(".rels"):
        return None
    return f"{on_ek}/{son[:-5]}"


def _hassas_iliskileri_bul(
    parcalar: dict[str, bytes], hassas_parcalar: set[str]
) -> tuple[dict[str, set[str]], dict[str, bytes]]:
    sahip_kimlikleri: dict[str, set[str]] = {}
    guncel_rels: dict[str, bytes] = {}
    for ad, veri in parcalar.items():
        if not ad.endswith(".rels"):
            continue
        try:
            kok = etree.fromstring(veri)
        except etree.XMLSyntaxError:
            continue
        sahip = _iliski_sahibi(ad)
        degisti = False
        for iliski in list(kok):
            hedef = iliski.get("Target", "").replace("\\", "/")
            if sahip:
                sahip_klasoru = str(Path(sahip).parent).replace("\\", "/")
                cozulmus = str(Path(sahip_klasoru, hedef)).replace("\\", "/")
            else:
                cozulmus = hedef.lstrip("/")
            while "/../" in f"/{cozulmus}":
                once = cozulmus
                cozulmus = str(Path(cozulmus)).replace("\\", "/")
                if cozulmus == once:
                    break
            if cozulmus not in hassas_parcalar:
                continue
            if sahip:
                sahip_kimlikleri.setdefault(sahip, set()).add(iliski.get("Id", ""))
            kok.remove(iliski)
            degisti = True
        if degisti:
            guncel_rels[ad] = etree.tostring(
                kok, xml_declaration=True, encoding="UTF-8", standalone=True
            )
    return sahip_kimlikleri, guncel_rels


def _bagli_nesneleri_sil(kok: etree._Element, kimlikler: set[str]) -> None:
    if not kimlikler:
        return
    adaylar = []
    for dugum in kok.iter():
        if any(deger in kimlikler for deger in dugum.attrib.values()):
            adaylar.append(dugum)
    for dugum in adaylar:
        hedef = dugum
        while hedef.getparent() is not None and etree.QName(hedef).localname not in {
            "drawing",
            "pict",
            "object",
        }:
            hedef = hedef.getparent()
        ebeveyn = hedef.getparent()
        if ebeveyn is not None:
            ebeveyn.remove(hedef)


def temizle(yol: Path) -> tuple[int, int]:
    parcalar_kucuk = {parca.lower() for parca in yol.parts}
    ad_kucuk = yol.stem.lower()
    rapor_sablonu = "rapor" in parcalar_kucuk or ad_kucuk == "taslak"
    taahhut_sablonu = "taahhutname" in parcalar_kucuk or "taahhut" in ad_kucuk
    if rapor_sablonu:
        degisimler = RAPOR_DEGISIMLERI
    elif taahhut_sablonu:
        degisimler = TAAHHUT_DEGISIMLERI
    else:
        degisimler = ()

    fd, gecici_ad = tempfile.mkstemp(prefix=f".{yol.stem}_", suffix=".docx", dir=yol.parent)
    os.close(fd)
    gecici = Path(gecici_ad)
    degisim_sayisi = 0
    yol_sayisi = 0
    try:
        with zipfile.ZipFile(yol) as kaynak:
            parcalar = {bilgi.filename: kaynak.read(bilgi.filename) for bilgi in kaynak.infolist()}
            bilgiler = list(kaynak.infolist())
        hassas_parcalar = RAPOR_HASSAS_PARCA_ADLARI if rapor_sablonu else set()
        sahip_kimlikleri, guncel_rels = _hassas_iliskileri_bul(parcalar, hassas_parcalar)
        with zipfile.ZipFile(
            gecici, "w", compression=zipfile.ZIP_DEFLATED
        ) as hedef:
            for bilgi in bilgiler:
                ad = bilgi.filename
                if ad == "docProps/custom.xml" or ad in hassas_parcalar:
                    continue
                veri = guncel_rels.get(ad, parcalar[ad])
                if _xml_mi(ad):
                    try:
                        kok = etree.fromstring(veri)
                    except etree.XMLSyntaxError:
                        hedef.writestr(bilgi, veri)
                        continue
                    _metadata_temizle(kok)
                    if ad.startswith("word/") and ad.endswith(".xml"):
                        _bagli_nesneleri_sil(kok, sahip_kimlikleri.get(ad, set()))
                        degisim_sayisi += _metni_degistir(kok, degisimler)
                    if ad.endswith(".rels"):
                        once = etree.tostring(kok)
                        _iliski_yollarini_temizle(kok)
                        if ad == "_rels/.rels":
                            _kok_iliskilerinden_custom_sil(kok)
                        yol_sayisi += int(etree.tostring(kok) != once)
                    elif ad == "[Content_Types].xml":
                        _icerik_turlerinden_custom_sil(kok)
                    veri = etree.tostring(
                        kok, xml_declaration=True, encoding="UTF-8", standalone=True
                    )
                hedef.writestr(bilgi, veri)
        os.replace(gecici, yol)
    except Exception:
        gecici.unlink(missing_ok=True)
        raise
    return degisim_sayisi, yol_sayisi


def docx_yollari(kok: Path) -> list[Path]:
    if kok.is_file():
        return [kok]
    return sorted(kok.rglob("*.docx"))


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("path", type=Path)
    parser.add_argument("--apply", action="store_true")
    args = parser.parse_args()

    yollar = docx_yollari(args.path)
    if not yollar:
        parser.error("DOCX bulunamadı")
    for yol in yollar:
        if args.apply:
            metin, iliski = temizle(yol)
            print(f"TEMIZLENDI {yol} metin={metin} iliski={iliski}")
        bulgular = denetle(yol)
        print(f"DENETIM {yol} bulgu={len(bulgular)}")
        for bulgu in bulgular:
            print(f"  {bulgu}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
