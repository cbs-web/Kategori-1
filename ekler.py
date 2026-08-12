import os
import tempfile
import unicodedata

from PIL import Image, ImageOps


PDF_UYUMLU_UZANTILAR = {".pdf", ".jpg", ".jpeg", ".png"}
OTOMATIK_PDF_DONUSUM_UZANTILARI = {".docx"}


def ek_taahhutname_mi(ek):
    """Taahhütname çıktısını rapor eklerinden ayırt et."""
    if not isinstance(ek, dict):
        return False
    metin = " ".join(
        str(ek.get(alan) or "")
        for alan in ("baslik", "yol")
    ).casefold()
    metin = unicodedata.normalize("NFKD", metin)
    metin = "".join(ch for ch in metin if not unicodedata.combining(ch))
    return "taahhut" in metin


def ek_dosya_turu(yol):
    uzanti = os.path.splitext(yol)[1].replace(".", "").upper()
    return uzanti or "DOSYA"


def ek_pdf_uyumlu_mu(yol):
    return os.path.splitext(yol)[1].lower() in PDF_UYUMLU_UZANTILAR


def ek_dosya_dogrulama_hatasi(yol):
    if not os.path.isfile(yol):
        return "Dosya bulunamadı"
    uzanti = os.path.splitext(yol)[1].lower()
    try:
        if uzanti == ".pdf":
            PdfReader, _ = pdf_yazici_siniflarini_al()
            reader = PdfReader(yol)
            if getattr(reader, "is_encrypted", False):
                sonuc = reader.decrypt("")
                if not sonuc:
                    return "PDF parola korumalı"
            if len(reader.pages) == 0:
                return "PDF sayfa içermiyor"
            for sayfa in reader.pages:
                _ = sayfa.mediabox
        elif uzanti in {".jpg", ".jpeg", ".png"}:
            with Image.open(yol) as img:
                img.verify()
        else:
            return "PDF'e dönüştürülmesi gerekiyor"
    except Exception as exc:
        return f"Okunamadı: {exc}"
    return ""


def ek_durumunu_hazirla(ek, derin=False):
    yol = ek.get("yol", "")
    if not os.path.isfile(yol):
        return "Dosya yok", "eksik"
    if os.path.splitext(yol)[1].lower() in OTOMATIK_PDF_DONUSUM_UZANTILARI:
        return "PDF'e otomatik çevrilecek", "otomatik_donusum"
    if not ek_pdf_uyumlu_mu(yol):
        return "PDF'e çevrilmeli", "donusum"
    if derin:
        hata = ek_dosya_dogrulama_hatasi(yol)
        if hata:
            return hata, "gecersiz"
    return "PDF hazır", "var"


def ekleri_sirali_listele(ekler, kategoriler, derin=False):
    sirali = []
    for kategori in kategoriler:
        for sira, ek in enumerate(ekler.get(kategori, []), start=1):
            if ek_taahhutname_mi(ek):
                continue
            yol = ek.get("yol", "")
            durum, tag = ek_durumunu_hazirla(ek, derin=derin)
            sirali.append({
                "kategori": kategori,
                "sira": sira,
                "baslik": ek.get("baslik", ""),
                "yol": yol,
                "tur": ek_dosya_turu(yol),
                "durum": durum,
                "tag": tag,
            })
    return sirali


def ekleri_denetle(ekler, kategoriler, derin=False):
    sirali = ekleri_sirali_listele(ekler, kategoriler, derin=derin)
    bos_kategoriler = [
        kategori
        for kategori in kategoriler
        if not any(
            not ek_taahhutname_mi(ek)
            for ek in (ekler.get(kategori, []) or [])
        )
    ]
    eksik_dosyalar = [ek for ek in sirali if ek["tag"] == "eksik"]
    otomatik_donusumler = [ek for ek in sirali if ek["tag"] == "otomatik_donusum"]
    donusum_gerekenler = [ek for ek in sirali if ek["tag"] == "donusum"]
    gecersiz_dosyalar = [ek for ek in sirali if ek["tag"] == "gecersiz"]
    pdf_hazir = [ek for ek in sirali if ek["tag"] == "var"]
    return {
        "toplam": len(sirali),
        "pdf_hazir": len(pdf_hazir),
        "bos_kategoriler": bos_kategoriler,
        "eksik_dosyalar": eksik_dosyalar,
        "otomatik_donusumler": otomatik_donusumler,
        "donusum_gerekenler": donusum_gerekenler,
        "gecersiz_dosyalar": gecersiz_dosyalar,
        "sirali": sirali,
    }


def ek_kategori_durumunu_hazirla(ekler, kategori):
    liste = [
        ek for ek in (ekler.get(kategori, []) or [])
        if not ek_taahhutname_mi(ek)
    ]
    if not liste:
        return "İsteğe bağlı / eklenmedi", "secondary"
    denetim = ekleri_denetle({kategori: liste}, [kategori])
    if denetim["eksik_dosyalar"]:
        return f"{len(liste)} dosya / {len(denetim['eksik_dosyalar'])} yok", "warning"
    if denetim["donusum_gerekenler"]:
        return f"{len(liste)} dosya / {len(denetim['donusum_gerekenler'])} çevrilecek", "warning"
    if denetim["otomatik_donusumler"]:
        return (
            f"PDF hazır / {len(denetim['otomatik_donusumler'])} Word otomatik çevrilecek",
            "info",
        )
    return f"PDF hazır ({len(liste)})", "success"


def pdf_yazici_siniflarini_al():
    try:
        from pypdf import PdfReader, PdfWriter
        return PdfReader, PdfWriter
    except ImportError:
        try:
            from PyPDF2 import PdfReader, PdfWriter
            return PdfReader, PdfWriter
        except ImportError as exc:
            raise RuntimeError(
                "Sıralı EKLER PDF oluşturmak için pypdf veya PyPDF2 kütüphanesi gerekli."
            ) from exc


def gorseli_gecici_pdf_yap(gorsel_yolu, gecici_klasor):
    with tempfile.NamedTemporaryFile(
        prefix="ek_gorsel_",
        suffix=".pdf",
        dir=gecici_klasor,
        delete=False,
    ) as f:
        pdf_yolu = f.name
    with Image.open(gorsel_yolu) as img:
        img = ImageOps.exif_transpose(img)
        if "A" in img.getbands():
            arka_plan = Image.new("RGB", img.size, "white")
            arka_plan.paste(img, mask=img.getchannel("A"))
            img = arka_plan
        elif img.mode != "RGB":
            img = img.convert("RGB")
        img.save(pdf_yolu, "PDF", resolution=150.0)
    return pdf_yolu


def ekler_pdf_olustur(ekler, kategoriler, hedef_yol, docx_donusturucu=None):
    denetim = ekleri_denetle(ekler, kategoriler, derin=True)
    if denetim["toplam"] == 0:
        raise ValueError("PDF oluşturmak için ek dosya bulunmuyor.")
    if denetim["eksik_dosyalar"]:
        raise ValueError("Dosya yolu bulunamayan ekler var. Önce eksikleri düzeltin.")
    if denetim["donusum_gerekenler"]:
        raise ValueError("PDF'e çevrilmesi gereken Word/Excel veya desteklenmeyen ekler var.")
    if denetim["gecersiz_dosyalar"]:
        raise ValueError("Bozuk, parola korumalı veya okunamayan ek dosyalar var.")

    hedef_mutlak = os.path.normcase(os.path.abspath(hedef_yol))
    kaynak_yollar = {
        os.path.normcase(os.path.abspath(ek["yol"]))
        for ek in denetim["sirali"]
    }
    if hedef_mutlak in kaynak_yollar:
        raise ValueError("Birleşik PDF hedefi kaynak ek dosyalarından biriyle aynı olamaz.")

    PdfReader, PdfWriter = pdf_yazici_siniflarini_al()
    writer = PdfWriter()
    eklenen_sayfa = 0
    donusturulen_dosya = 0

    with tempfile.TemporaryDirectory() as gecici_klasor:
        for sira, ek in enumerate(denetim["sirali"], start=1):
            yol = ek["yol"]
            uzanti = os.path.splitext(yol)[1].lower()
            kaynak_pdf = yol
            if uzanti in {".jpg", ".jpeg", ".png"}:
                kaynak_pdf = gorseli_gecici_pdf_yap(yol, gecici_klasor)
            elif uzanti == ".docx":
                if docx_donusturucu is None:
                    # Yerel içe aktarma, taahhutname_islemleri -> ekler bağımlılığıyla
                    # modül yüklenirken döngü oluşmasını önler.
                    from taahhutname_islemleri import taahhut_docx_pdfye_cevir

                    donusturucu = taahhut_docx_pdfye_cevir
                else:
                    donusturucu = docx_donusturucu
                kaynak_pdf = os.path.join(gecici_klasor, f"ek_word_{sira:03d}.pdf")
                basarili, hata = donusturucu(yol, kaynak_pdf)
                if not basarili:
                    baslik = ek.get("baslik") or os.path.basename(yol)
                    raise ValueError(
                        f"Word eki PDF'ye dönüştürülemedi: {baslik} "
                        f"({os.path.basename(yol)}). {hata or 'Dönüştürücü ayrıntı vermedi.'}"
                    )
                pdf_hatasi = ek_dosya_dogrulama_hatasi(kaynak_pdf)
                if pdf_hatasi:
                    baslik = ek.get("baslik") or os.path.basename(yol)
                    raise ValueError(
                        f"Dönüştürülen Word eki geçerli bir PDF üretmedi: {baslik}. {pdf_hatasi}"
                    )
                donusturulen_dosya += 1

            reader = PdfReader(kaynak_pdf)
            for sayfa in reader.pages:
                writer.add_page(sayfa)
                eklenen_sayfa += 1

        if eklenen_sayfa == 0:
            raise ValueError("Ek dosyalardan PDF sayfası üretilemedi.")

        hedef_klasor = os.path.dirname(os.path.abspath(hedef_yol))
        if not os.path.isdir(hedef_klasor):
            raise ValueError(f"Hedef klasör bulunamadı: {hedef_klasor}")
        gecici_hedef = ""
        try:
            with tempfile.NamedTemporaryFile(
                prefix=f".{os.path.basename(hedef_yol)}.",
                suffix=".tmp.pdf",
                dir=hedef_klasor,
                delete=False,
            ) as f:
                gecici_hedef = f.name
                writer.write(f)
            kontrol = PdfReader(gecici_hedef)
            if len(kontrol.pages) != eklenen_sayfa:
                raise ValueError("Birleşik PDF doğrulamasında sayfa sayısı uyuşmadı.")
            os.replace(gecici_hedef, hedef_yol)
            gecici_hedef = ""
        finally:
            if gecici_hedef:
                try:
                    os.remove(gecici_hedef)
                except OSError:
                    pass

    return {
        "dosya": hedef_yol,
        "sayfa": eklenen_sayfa,
        "dosya_sayisi": denetim["toplam"],
        "donusturulen_dosya_sayisi": donusturulen_dosya,
    }
