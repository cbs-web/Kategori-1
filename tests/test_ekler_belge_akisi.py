from pathlib import Path

import pytest
from PIL import Image
from pypdf import PdfReader, PdfWriter

from ekler import ekleri_denetle, ekler_pdf_olustur


def _pdf(path, pages=1):
    writer = PdfWriter()
    for _ in range(pages):
        writer.add_blank_page(width=200, height=300)
    with open(path, "wb") as f:
        writer.write(f)


def test_derin_ek_denetimi_bozuk_pdfyi_yakalar(tmp_path):
    bozuk = tmp_path / "bozuk.pdf"
    bozuk.write_bytes(b"PDF degil")
    denetim = ekleri_denetle(
        {"EVRAKLAR": [{"baslik": "Bozuk", "yol": str(bozuk)}]},
        ["EVRAKLAR"],
        derin=True,
    )
    assert len(denetim["gecersiz_dosyalar"]) == 1
    assert denetim["pdf_hazir"] == 0


def test_pdf_birlestirme_gorseli_dogrular_ve_atomik_yazar(tmp_path):
    kaynak_pdf = tmp_path / "kaynak.pdf"
    _pdf(kaynak_pdf, pages=2)
    gorsel = tmp_path / "seffaf.png"
    Image.new("RGBA", (20, 10), (255, 0, 0, 0)).save(gorsel)
    hedef = tmp_path / "birlesik.pdf"
    hedef.write_bytes(b"eski")
    ekler = {
        "EVRAKLAR": [
            {"baslik": "PDF", "yol": str(kaynak_pdf)},
            {"baslik": "Görsel", "yol": str(gorsel)},
        ]
    }

    sonuc = ekler_pdf_olustur(ekler, ["EVRAKLAR"], str(hedef))

    assert sonuc["sayfa"] == 3
    assert len(PdfReader(hedef).pages) == 3
    assert not list(tmp_path.glob(".*.tmp.pdf"))


def test_pdf_hedefi_kaynak_dosyayla_ayni_olamaz(tmp_path):
    kaynak_pdf = tmp_path / "kaynak.pdf"
    _pdf(kaynak_pdf)
    ekler = {"EVRAKLAR": [{"baslik": "PDF", "yol": str(kaynak_pdf)}]}
    with pytest.raises(ValueError, match="aynı olamaz"):
        ekler_pdf_olustur(ekler, ["EVRAKLAR"], str(kaynak_pdf))


def test_docx_eki_otomatik_pdfye_cevrilip_siraya_eklenir(tmp_path):
    kaynak_pdf = tmp_path / "once.pdf"
    _pdf(kaynak_pdf, pages=1)
    word = tmp_path / "Profil_1_MASW_degerlendirme.docx"
    word.write_bytes(b"orijinal-word")
    hedef = tmp_path / "ekler.pdf"
    donusumler = []

    def sahte_docx_donusturucu(docx_yolu, pdf_yolu):
        donusumler.append((Path(docx_yolu).name, Path(pdf_yolu).name))
        _pdf(pdf_yolu, pages=2)
        return True, ""

    ekler = {
        "JEOFİZİK": [
            {"baslik": "Önce", "yol": str(kaynak_pdf)},
            {"baslik": "Profil 1 MASW", "yol": str(word)},
        ]
    }
    denetim = ekleri_denetle(ekler, ["JEOFİZİK"], derin=True)

    assert not denetim["donusum_gerekenler"]
    assert len(denetim["otomatik_donusumler"]) == 1

    sonuc = ekler_pdf_olustur(
        ekler,
        ["JEOFİZİK"],
        str(hedef),
        docx_donusturucu=sahte_docx_donusturucu,
    )

    assert sonuc["dosya_sayisi"] == 2
    assert sonuc["donusturulen_dosya_sayisi"] == 1
    assert sonuc["sayfa"] == 3
    assert len(PdfReader(hedef).pages) == 3
    assert donusumler == [(word.name, "ek_word_002.pdf")]
    assert word.read_bytes() == b"orijinal-word"
