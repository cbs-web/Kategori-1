import os
import shutil
import subprocess
import tempfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from taahhutname import TaahhutnameUretici
from ekler import ek_dosya_dogrulama_hatasi
from taahhutname_raporpro import (
    taahhut_xlsx_kaydet,
    taahhut_xlsx_pdfye_cevir,
    taahhutname_dosya_adi,
)


TAAHHUT_ALAN_ETIKETLERI = [
    ("JEOFIZIK_MUH_AD", "Jeofizik mühendisi adı"),
    ("JEOFIZIK_MUH_SICIL", "Jeofizik oda sicil no"),
    ("JEOFIZIK_MUH_ADRES", "Jeofizik adresi"),
    ("JEOFIZIK_MUH_TELEFON", "Jeofizik telefonu"),
    ("JEOLOJI_MUH_AD", "Jeoloji mühendisi adı"),
    ("JEOLOJI_MUH_SICIL", "Jeoloji oda sicil no"),
    ("JEOLOJI_MUH_ADRES", "Jeoloji adresi"),
    ("JEOLOJI_MUH_TELEFON", "Jeoloji telefonu"),
]


def benzersiz_taahhut_yolu(yol):
    kok, uzanti = os.path.splitext(yol)
    aday = yol
    sira = 2
    while os.path.exists(aday):
        aday = f"{kok}_{sira}{uzanti}"
        sira += 1
    return aday


def benzersiz_taahhut_koku(kok):
    """Aynı ada sahip XLSX/PDF çiftinin ikisi için de boş bir kök bul."""
    aday = kok
    sira = 2
    while any(os.path.exists(f"{aday}{uzanti}") for uzanti in (".xlsx", ".pdf")):
        aday = f"{kok}_{sira}"
        sira += 1
    return aday


def atomik_taahhut_docx_kaydet(doc, hedef_yol):
    hedef_yol = os.path.abspath(hedef_yol)
    klasor = os.path.dirname(hedef_yol)
    gecici_yol = ""
    try:
        with tempfile.NamedTemporaryFile(
            prefix=f".{os.path.basename(hedef_yol)}.",
            suffix=".tmp.docx",
            dir=klasor,
            delete=False,
        ) as f:
            gecici_yol = f.name
        doc.save(gecici_yol)
        from docx import Document

        Document(gecici_yol)
        os.replace(gecici_yol, hedef_yol)
        gecici_yol = ""
    finally:
        if gecici_yol:
            try:
                os.remove(gecici_yol)
            except OSError:
                pass


def _word_docx_pdf_bir_kez(docx_yolu, pdf_yolu, yontem="export"):
    """Yeni bir Word oturumunda PDF aktar; COM kapanış hatasını çıktından ayır."""
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        return False, f"Microsoft Word PDF dönüştürücüsü için pywin32 bulunamadı: {exc}"

    word = None
    belge = None
    com_hazir = False
    aktarim_hatasi = None
    try:
        pythoncom.CoInitialize()
        com_hazir = True
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            word.AutomationSecurity = 3
        except Exception:
            pass
        belge = word.Documents.Open(
            FileName=os.path.abspath(docx_yolu),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
            OpenAndRepair=True,
            NoEncodingDialog=True,
        )
        if yontem == "saveas":
            belge.SaveAs2(
                FileName=os.path.abspath(pdf_yolu),
                FileFormat=17,
                AddToRecentFiles=False,
            )
        else:
            belge.ExportAsFixedFormat(
                OutputFileName=os.path.abspath(pdf_yolu),
                ExportFormat=17,
                OpenAfterExport=False,
                OptimizeFor=0,
                CreateBookmarks=1,
            )
    except Exception as exc:
        aktarim_hatasi = exc
    finally:
        # Word, büyük bir raporun aktarımından sonra kendi COM nesnesini ayırabilir.
        # Kapanış hatası geçerli PDF'i başarısız saydırmamalıdır.
        if belge is not None:
            try:
                belge.Close(SaveChanges=False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        if com_hazir:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    dogrulama_hatasi = ek_dosya_dogrulama_hatasi(pdf_yolu)
    if not dogrulama_hatasi:
        return True, ""
    if aktarim_hatasi is not None:
        return False, f"{aktarim_hatasi}; {dogrulama_hatasi}"
    return False, dogrulama_hatasi


def taahhut_docx_pdfye_cevir(docx_yolu, pdf_yolu):
    """LibreOffice veya Microsoft Word ile DOCX'i doğrulanmış PDF'e çevir."""
    docx_yolu = os.path.abspath(docx_yolu)
    pdf_yolu = os.path.abspath(pdf_yolu)
    hedef_klasor = os.path.dirname(pdf_yolu)
    if not os.path.isfile(docx_yolu):
        return False, "Dönüştürülecek Word belgesi bulunamadı."
    try:
        with tempfile.TemporaryDirectory(prefix="taahhut_pdf_", dir=hedef_klasor) as gecici_klasor:
            gecici_pdf = os.path.join(
                gecici_klasor,
                os.path.splitext(os.path.basename(docx_yolu))[0] + ".pdf",
            )
            hatalar = []
            soffice = shutil.which("soffice") or shutil.which("libreoffice")
            if soffice:
                try:
                    subprocess.run(
                        [soffice, "--headless", "--convert-to", "pdf", "--outdir", gecici_klasor, docx_yolu],
                        check=True,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE,
                        text=True,
                        timeout=120,
                    )
                except Exception as exc:
                    hatalar.append(f"LibreOffice: {exc}")
                else:
                    hata = ek_dosya_dogrulama_hatasi(gecici_pdf)
                    if not hata:
                        os.replace(gecici_pdf, pdf_yolu)
                        return True, ""
                    hatalar.append(f"LibreOffice çıktısı: {hata}")

            for yontem in ("export", "saveas"):
                if os.path.exists(gecici_pdf):
                    try:
                        os.remove(gecici_pdf)
                    except OSError:
                        pass
                basarili, hata = _word_docx_pdf_bir_kez(
                    docx_yolu,
                    gecici_pdf,
                    yontem=yontem,
                )
                if basarili:
                    os.replace(gecici_pdf, pdf_yolu)
                    return True, ""
                hatalar.append(f"Microsoft Word ({yontem}): {hata}")

            hata_ozeti = " | ".join(hata for hata in hatalar if hata)
            return False, hata_ozeti or "LibreOffice veya Microsoft Word PDF dönüştürücüsü bulunamadı."
    except Exception as exc:
        return False, str(exc)


class TaahhutnameIslemleri:
    def __init__(self, app):
        object.__setattr__(self, "app", app)

    def __getattr__(self, name):
        return getattr(self.app, name)

    def __setattr__(self, name, value):
        if name == "app":
            object.__setattr__(self, name, value)
        else:
            setattr(self.app, name, value)

    def varsayilan_taahhut_word_sablonu(self):
        return self.sablon_dosyasi_bul("taahhutname", ["örnek taahütname.docx"], "taah")

    def taahhut_word_sablonu_sec(self):
        baslangic_klasoru = (
            os.path.dirname(self.taahhut_word_sablon_yolu)
            if self.taahhut_word_sablon_yolu
            else self.sablon_alt_klasoru("taahhutname")
        )
        yol = filedialog.askopenfilename(
            initialdir=baslangic_klasoru,
            filetypes=[("Word Dosyaları", "*.docx"), ("Tüm Dosyalar", "*.*")],
            title="Taahhütname Word şablonunu seçin",
        )
        if yol:
            self.taahhut_word_sablon_yolu = yol
            self.taahhut_word_sablon_label_guncelle()

    def taahhut_word_sablon_label_guncelle(self):
        if hasattr(self, "lbl_taahhut_word"):
            text = self.taahhut_word_sablon_yolu if self.taahhut_word_sablon_yolu else "Seçilmedi"
            self.lbl_taahhut_word.config(text=text)

    def taahhutname_paneli_ekle(self, parent):
        frame = ttk.LabelFrame(parent, text="Taahhütname Üretimi", padding=8, bootstyle="info")
        frame.pack(fill="x", pady=(0, 8))

        ttk.Button(
            frame,
            text="Jeoloji ve Jeofizik Taahhütnamelerini Oluştur",
            command=self.taahhutnameleri_olustur,
            bootstyle="success",
        ).pack(side="left")
        ttk.Button(
            frame,
            text="Mühendis Bilgilerini Düzenle",
            command=self.taahhut_bilgilerini_duzenle,
            bootstyle="secondary",
        ).pack(side="left", padx=(8, 0))

    def taahhut_bilgilerini_topla(self):
        kaynak = getattr(self, "taahhut_bilgileri", None)
        if not isinstance(kaynak, dict):
            kaynak = getattr(self, "taahhut_varsayilanlari", {})
        return {kod: str(kaynak.get(kod, "")).strip() for kod, _ in TAAHHUT_ALAN_ETIKETLERI}

    def taahhut_bilgilerini_yerlestir(self, veriler):
        if not isinstance(veriler, dict):
            return
        mevcut = self.taahhut_bilgilerini_topla()
        for kod, _ in TAAHHUT_ALAN_ETIKETLERI:
            if kod in veriler:
                mevcut[kod] = str(veriler[kod]).strip()
        self.taahhut_bilgileri = mevcut

    def taahhut_bilgilerini_kalici_kaydet(self, veriler):
        """Bilgileri açık projeye ve sonraki projelerin varsayılanına uygula."""
        yeni = {
            kod: str((veriler or {}).get(kod, "") or "").strip()
            for kod, _etiket in TAAHHUT_ALAN_ETIKETLERI
        }
        kaydedilen = self.kayit_yoneticisi().taahhut_varsayilanlarini_kaydet(yeni)
        self.taahhut_varsayilanlari = dict(kaydedilen)
        self.taahhut_bilgileri = dict(kaydedilen)

        varsayilan_proje = getattr(self, "varsayilan_proje_verisi", None)
        if isinstance(varsayilan_proje, dict):
            varsayilan_proje["_TAAHHUT_BILGILERI_"] = dict(kaydedilen)
        self._proje_kirli = True
        return kaydedilen

    def taahhut_eksik_muhendis_alanlari(self):
        bilgiler = self.taahhut_bilgilerini_topla()
        return [
            (kod, etiket)
            for kod, etiket in TAAHHUT_ALAN_ETIKETLERI
            if not bilgiler.get(kod, "").strip()
        ]

    def taahhut_bilgilerini_duzenle(self, kaydedilince=None):
        pencere = self.animasyonlu_pencere()
        pencere.title("Taahhütname Mühendis Bilgileri")
        pencere.transient(self.root)
        pencere.grab_set()
        govde = ttk.Frame(pencere, padding=12)
        govde.pack(fill="both", expand=True)
        mevcut = self.taahhut_bilgilerini_topla()
        entryler = {}
        for satir, (kod, etiket) in enumerate(TAAHHUT_ALAN_ETIKETLERI):
            ttk.Label(govde, text=etiket).grid(row=satir, column=0, padx=(0, 8), pady=4, sticky="w")
            entry = ttk.Entry(govde, width=58)
            entry.insert(0, mevcut.get(kod, ""))
            entry.grid(row=satir, column=1, pady=4, sticky="ew")
            entryler[kod] = entry
        govde.columnconfigure(1, weight=1)

        def kaydet():
            yeni = {kod: entry.get().strip() for kod, entry in entryler.items()}
            eksikler = [etiket for kod, etiket in TAAHHUT_ALAN_ETIKETLERI if not yeni[kod]]
            if eksikler:
                messagebox.showwarning("Eksik Bilgi", "Boş bırakılamayan alanlar:\n- " + "\n- ".join(eksikler), parent=pencere)
                return
            try:
                self.taahhut_bilgilerini_kalici_kaydet(yeni)
            except Exception as exc:
                self.hata_kaydet("Taahhütname mühendis bilgileri kalıcı kaydedilemedi", exc)
                messagebox.showerror(
                    "Kayıt Hatası",
                    f"Mühendis bilgileri kalıcı olarak kaydedilemedi:\n{exc}",
                    parent=pencere,
                )
                return
            pencere.destroy()
            if hasattr(self, "durum_mesaji_yaz"):
                self.durum_mesaji_yaz(
                    "Taahhütname mühendis bilgileri kaydedildi",
                    "Yeni projelerde otomatik kullanılacak",
                )
            if callable(kaydedilince):
                kaydedilince()

        butonlar = ttk.Frame(govde)
        butonlar.grid(row=len(TAAHHUT_ALAN_ETIKETLERI), column=0, columnspan=2, pady=(10, 0), sticky="e")
        ttk.Button(butonlar, text="İptal", command=pencere.destroy, bootstyle="secondary").pack(side="left", padx=(0, 6))
        ttk.Button(butonlar, text="Kaydet", command=kaydet, bootstyle="success").pack(side="left")

        eksikler = {kod for kod, _etiket in self.taahhut_eksik_muhendis_alanlari()}
        ilk_eksik = next((entryler[kod] for kod, _ in TAAHHUT_ALAN_ETIKETLERI if kod in eksikler), None)
        if ilk_eksik is not None:
            ilk_eksik.focus_set()

    def taahhutnameleri_olustur(self):
        eksikler = self.taahhut_eksik_muhendis_alanlari()
        if eksikler:
            messagebox.showinfo(
                "Mühendis Bilgileri Gerekli",
                "Taahhütname oluşturulmadan önce eksik mühendis bilgilerini doldurun. "
                "Kaydettiğinizde oluşturma işlemi otomatik olarak devam edecek.",
                parent=self.root,
            )
            self.taahhut_bilgilerini_duzenle(
                kaydedilince=self.taahhutnameleri_olustur
            )
            return

        uretici = TaahhutnameUretici(self.app)
        if not uretici.taahhut_ilgili_idare().strip():
            messagebox.showwarning(
                "İl Bilgisi Eksik",
                "İlgili idarenin belirlenebilmesi için 1. Proje Bilgileri sekmesindeki "
                "İl alanını doldurun.",
                parent=self.root,
            )
            return

        klasor = filedialog.askdirectory(title="Taahhütnamelerin kaydedileceği klasörü seçin")
        if not klasor:
            return
        try:
            hedefler = []
            for tur in ("jeofizik", "jeoloji"):
                xlsx_adi = taahhutname_dosya_adi(uretici, tur, ".xlsx")
                kok = os.path.join(klasor, os.path.splitext(xlsx_adi)[0])
                hedefler.append(
                    {
                        "tur": tur,
                        "kok": kok,
                        "xlsx": f"{kok}.xlsx",
                        "pdf": f"{kok}.pdf",
                    }
                )

            mevcut_var = any(
                os.path.exists(hedef[tur])
                for hedef in hedefler
                for tur in ("xlsx", "pdf")
            )
            if mevcut_var:
                karar = messagebox.askyesnocancel(
                    "Mevcut Taahhütname",
                    "RaporPro düzenindeki taahhütname dosyalarından bazıları zaten var.\n\n"
                    "Evet: güvenli biçimde üzerlerine yaz\n"
                    "Hayır: numaralı yeni dosyalar oluştur",
                )
                if karar is None:
                    return
                if not karar:
                    for hedef in hedefler:
                        yeni_kok = benzersiz_taahhut_koku(hedef["kok"])
                        hedef["kok"] = yeni_kok
                        hedef["xlsx"] = f"{yeni_kok}.xlsx"
                        hedef["pdf"] = f"{yeni_kok}.pdf"

            with tempfile.TemporaryDirectory(prefix="k1_taahhut_raporpro_", dir=klasor) as gecici:
                for hedef in hedefler:
                    gecici_xlsx = os.path.join(gecici, os.path.basename(hedef["xlsx"]))
                    gecici_pdf = os.path.join(gecici, os.path.basename(hedef["pdf"]))
                    taahhut_xlsx_kaydet(uretici, hedef["tur"], gecici_xlsx)
                    pdf_basarili, pdf_hatasi = taahhut_xlsx_pdfye_cevir(gecici_xlsx, gecici_pdf)
                    if not pdf_basarili:
                        raise RuntimeError(
                            f"{hedef['tur'].title()} taahhütnamesi PDF'e çevrilemedi: {pdf_hatasi}"
                        )
                    hedef["gecici_xlsx"] = gecici_xlsx
                    hedef["gecici_pdf"] = gecici_pdf

                for hedef in hedefler:
                    os.replace(hedef["gecici_xlsx"], hedef["xlsx"])
                    os.replace(hedef["gecici_pdf"], hedef["pdf"])

            dosya_listesi = "\n".join(
                f"- {os.path.basename(hedef['xlsx'])}\n- {os.path.basename(hedef['pdf'])}"
                for hedef in hedefler
            )
            messagebox.showinfo(
                "Başarılı",
                "Jeoloji ve jeofizik taahhütnameleri RaporPro düzeninde oluşturuldu. "
                "Dosyalar bağımsız olarak seçilen klasöre kaydedildi; rapor eklerine "
                "otomatik eklenmedi:\n\n"
                f"{dosya_listesi}",
            )
            if hasattr(self, "durum_mesaji_yaz"):
                self.durum_mesaji_yaz("RaporPro düzeninde taahhütnameler oluşturuldu")
        except Exception as e:
            self.hata_kaydet("Taahhütname oluşturulamadı", e)
            messagebox.showerror("Hata", f"Taahhütname oluşturulamadı:\n{e}")
