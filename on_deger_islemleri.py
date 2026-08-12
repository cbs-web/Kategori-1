import copy
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from on_deger import (
    IS_DURUMLARI,
    ZEMIN_SINIFLARI,
    is_durumu_degistir,
    normalize_is_akisi,
    normalize_on_deger,
    normalize_tdth,
    on_deger_durumu,
    on_deger_revizyonu_ekle,
    tdth_kaydi_etkinlestir,
    tdth_pdf_bilgilerini_oku,
    tdth_zemin_sinifi_guncelle,
)


class OnDegerIslemleri:
    def __init__(self, app):
        object.__setattr__(self, "app", app)

    def __getattr__(self, name):
        return getattr(self.app, name)

    def __setattr__(self, name, value):
        if name == "app":
            object.__setattr__(self, name, value)
        else:
            setattr(self.app, name, value)

    def on_deger_paneli_olustur(self, parent):
        panel = ttk.LabelFrame(parent, text="Zemin Ön Değerleri", padding=(12, 10), bootstyle="secondary")
        panel.pack(fill="x", pady=(0, 12))
        for col in (1, 3):
            panel.grid_columnconfigure(col, weight=1, minsize=190)

        self.on_deger_durumu_var = tk.StringVar(value="Ön Değer Verilmedi")
        self.on_qt_var = tk.StringVar()
        self.on_ks_var = tk.StringVar()
        self.on_zemin_sinifi_var = tk.StringVar()
        self.on_aciklama_var = tk.StringVar()
        self.tdth_ozet_var = tk.StringVar(value="TDTH PDF: Seçilmedi")
        self.pga_haritasi_ozet_var = tk.StringVar(value="PGA haritası: Seçilmedi")
        self.on_deger_karsilastirma_var = tk.StringVar(value="")

        ttk.Label(panel, text="Ön Değer Durumu", style="Muted.TLabel").grid(row=0, column=0, sticky="w", pady=4)
        self.lbl_on_deger_durumu = ttk.Label(panel, textvariable=self.on_deger_durumu_var, style="AltBaslik.TLabel")
        self.lbl_on_deger_durumu.grid(row=0, column=1, sticky="w", pady=4)

        ttk.Separator(panel).grid(row=1, column=0, columnspan=4, sticky="ew", pady=(5, 8))

        ttk.Label(panel, text="Yaklaşık qₜ [t/m²]", style="Muted.TLabel").grid(row=2, column=0, sticky="w", pady=4)
        self.entry_on_qt = ttk.Entry(panel, textvariable=self.on_qt_var)
        self.entry_on_qt.grid(row=2, column=1, sticky="ew", padx=(0, 18), pady=4)
        ttk.Label(panel, text="Yaklaşık kₛ [t/m³]", style="Muted.TLabel").grid(row=2, column=2, sticky="w", pady=4)
        self.entry_on_ks = ttk.Entry(panel, textvariable=self.on_ks_var)
        self.entry_on_ks.grid(row=2, column=3, sticky="ew", pady=4)

        ttk.Label(panel, text="Ön Yerel Zemin Sınıfı", style="Muted.TLabel").grid(row=3, column=0, sticky="w", pady=4)
        self.combo_on_zemin = ttk.Combobox(
            panel,
            textvariable=self.on_zemin_sinifi_var,
            values=ZEMIN_SINIFLARI,
            state="readonly",
        )
        self.combo_on_zemin.grid(row=3, column=1, sticky="ew", padx=(0, 18), pady=4)
        self.combo_on_zemin.bind("<<ComboboxSelected>>", self.on_zemin_sinifi_degisti)
        ttk.Label(panel, text="Açıklama", style="Muted.TLabel").grid(row=3, column=2, sticky="w", pady=4)
        self.entry_on_aciklama = ttk.Entry(panel, textvariable=self.on_aciklama_var)
        self.entry_on_aciklama.grid(row=3, column=3, sticky="ew", pady=4)

        tdth_satiri = ttk.Frame(panel)
        tdth_satiri.grid(row=4, column=0, columnspan=4, sticky="ew", pady=(8, 4))
        self.lbl_tdth_ozet = ttk.Label(tdth_satiri, textvariable=self.tdth_ozet_var, wraplength=720, justify="left")
        self.lbl_tdth_ozet.pack(side="left", fill="x", expand=True)
        self.btn_tdth_sec = ttk.Button(tdth_satiri, text="TDTH PDF Seç", command=self.tdth_pdf_sec, bootstyle="primary")
        self.btn_tdth_sec.pack(side="left", padx=(8, 6))
        self.btn_tdth_ac = ttk.Button(tdth_satiri, text="PDF'yi Aç", command=self.tdth_pdf_ac, bootstyle="secondary")
        self.btn_tdth_ac.pack(side="left")

        pga_satiri = ttk.Frame(panel)
        pga_satiri.grid(row=5, column=0, columnspan=4, sticky="ew", pady=(2, 4))
        ttk.Label(
            pga_satiri,
            textvariable=self.pga_haritasi_ozet_var,
            style="Muted.TLabel",
        ).pack(side="left", fill="x", expand=True)
        self.btn_pga_haritasi_sec = ttk.Button(
            pga_satiri,
            text="PGA Haritası Seç",
            command=self.pga_haritasi_sec,
            bootstyle="secondary outline",
        )
        self.btn_pga_haritasi_sec.pack(side="left", padx=(8, 0))

        alt = ttk.Frame(panel)
        alt.grid(row=6, column=0, columnspan=4, sticky="ew", pady=(8, 0))
        self.btn_on_deger_kaydet = ttk.Button(
            alt, text="Ön Değeri Kaydet", command=self.on_deger_kaydet, bootstyle="success"
        )
        self.btn_on_deger_kaydet.pack(side="left", padx=(0, 6))
        self.btn_on_deger_gecmis = ttk.Button(
            alt, text="Ön Değer Geçmişi", command=self.on_deger_gecmisini_goster, bootstyle="secondary"
        )
        self.btn_on_deger_gecmis.pack(side="left")
        self.btn_hesaptan_on_deger = ttk.Button(
            alt,
            text="Hesap Sonuçlarını Ön Değere Al",
            command=self.hesap_sonuclarini_on_degere_al,
            bootstyle="info outline",
        )
        self.btn_hesaptan_on_deger.pack(side="left", padx=(6, 0))
        ttk.Label(alt, textvariable=self.on_deger_karsilastirma_var, style="Muted.TLabel").pack(side="right")

        self.salt_okunurda_acik_widgetlar = [self.btn_tdth_ac, self.btn_on_deger_gecmis]
        self.on_deger_ekranini_guncelle()
        aktif = self.tdth_verisi.get("aktif") or {}
        if aktif:
            self._tdth_rapor_alanlarini_uygula(aktif)
        return panel

    def pga_haritasi_sec(self):
        if not self.degisiklik_izni_kontrol_et("PGA haritası seçimi"):
            return
        yol = filedialog.askopenfilename(
            title="Çanakkale PGA haritası görselini seç",
            filetypes=[
                ("Harita Görselleri", "*.jpg *.jpeg *.png"),
                ("Tüm Dosyalar", "*.*"),
            ],
        )
        if not yol:
            return
        kaynak = os.path.abspath(yol)
        proje_yolu = str(getattr(self, "guncel_dosya_yolu", "") or "").strip()
        if proje_yolu:
            proje_klasoru = os.path.dirname(os.path.abspath(proje_yolu))
            haritalar_klasoru = os.path.join(proje_klasoru, "Haritalar")
            hedef = os.path.join(haritalar_klasoru, "PGA_Haritasi.jpg")
            try:
                from PIL import Image, ImageOps

                os.makedirs(haritalar_klasoru, exist_ok=True)
                with Image.open(kaynak) as image:
                    image = ImageOps.exif_transpose(image)
                    if image.mode not in {"RGB", "L"}:
                        image = image.convert("RGB")
                    image.save(hedef, "JPEG", quality=95)
                kaynak = hedef
            except Exception as exc:
                self.hata_kaydet("PGA haritası proje klasörüne kopyalanamadı", exc)
                messagebox.showerror(
                    "PGA Haritası",
                    f"Seçilen görsel proje klasörüne kaydedilemedi:\n{exc}",
                )
                return
        self.img_pga_haritasi = kaynak
        self.on_deger_ekranini_guncelle()
        if hasattr(self, "durum_mesaji_yaz"):
            self.durum_mesaji_yaz("PGA haritası seçildi", os.path.basename(yol))

    def degisiklik_izni_kontrol_et(self, eylem="Bu işlem"):
        if not getattr(self, "proje_salt_okunur", False):
            return True
        messagebox.showwarning("İzleme Modu", f"{eylem}, izleme modunda kullanılamaz.")
        return False

    def on_deger_verisini_topla(self):
        sonuc = normalize_on_deger(getattr(self, "on_deger_verisi", {}))
        if hasattr(self, "on_qt_var"):
            sonuc["guncel"] = {
                "qt": self.on_qt_var.get().strip(),
                "ks": self.on_ks_var.get().strip(),
                "zemin_sinifi": self.on_zemin_sinifi_var.get().strip().upper(),
                "aciklama": self.on_aciklama_var.get().strip(),
            }
        self.on_deger_verisi = sonuc
        return copy.deepcopy(sonuc)

    def on_deger_verisini_yerlestir(self, on_deger, tdth, is_akisi):
        self.on_deger_verisi = normalize_on_deger(on_deger)
        self.tdth_verisi = normalize_tdth(tdth)
        self.is_akisi_verisi = normalize_is_akisi(is_akisi)
        if not hasattr(self, "on_qt_var"):
            return
        guncel = self.on_deger_verisi.get("guncel", {})
        self.on_qt_var.set(guncel.get("qt", ""))
        self.on_ks_var.set(guncel.get("ks", ""))
        self.on_zemin_sinifi_var.set(guncel.get("zemin_sinifi", ""))
        self.on_aciklama_var.set(guncel.get("aciklama", ""))
        self.on_deger_ekranini_guncelle()
        aktif = self.tdth_verisi.get("aktif") or {}
        if aktif:
            self._tdth_rapor_alanlarini_uygula(aktif)

    def on_deger_ekranini_guncelle(self):
        if not hasattr(self, "on_deger_durumu_var"):
            return
        akis = normalize_is_akisi(getattr(self, "is_akisi_verisi", {}))
        self.is_akisi_verisi = akis
        durum = akis.get("durum", "yeni")
        on_durum = on_deger_durumu(getattr(self, "on_deger_verisi", {}))
        self.on_deger_durumu_var.set("Ön Değer Verildi" if on_durum == "verildi" else "Ön Değer Verilmedi")

        tdth = normalize_tdth(getattr(self, "tdth_verisi", {}))
        self.tdth_verisi = tdth
        aktif = tdth.get("aktif") or {}
        if aktif:
            ozet = (
                f"TDTH: {aktif.get('orijinal_dosya_adi') or os.path.basename(aktif.get('pdf_yolu', ''))} · "
                f"{aktif.get('sayfa_sayisi', 0)} sayfa · "
                f"{aktif.get('zemin_sinifi') or 'Zemin sınıfı okunamadı'} · Durum: {tdth.get('durum')}"
            )
            if tdth.get("uyarilar"):
                ozet += " · " + "; ".join(tdth["uyarilar"][:2])
        else:
            ozet = "TDTH PDF: Seçilmedi"
        self.tdth_ozet_var.set(ozet)

        pga_yolu = str(getattr(self, "img_pga_haritasi", "") or "").strip()
        if pga_yolu and os.path.isfile(pga_yolu):
            self.pga_haritasi_ozet_var.set(
                f"PGA haritası: {os.path.basename(pga_yolu)}"
            )
        else:
            self.pga_haritasi_ozet_var.set("PGA haritası: Seçilmedi")

        ilk = (getattr(self, "on_deger_verisi", {}) or {}).get("ilk") or {}
        qt_nihai = self.entry_qt_nihai.get().strip() if hasattr(self, "entry_qt_nihai") else ""
        ks_nihai = self.entry_ks_nihai.get().strip() if hasattr(self, "entry_ks_nihai") else ""
        if ilk:
            self.on_deger_karsilastirma_var.set(
                f"İlk: qₜ {ilk.get('qt', '-')} / kₛ {ilk.get('ks', '-')} · "
                f"Nihai: qₜ {qt_nihai or '-'} / kₛ {ks_nihai or '-'}"
            )
        else:
            self.on_deger_karsilastirma_var.set("Henüz ön değer kaydı yok")

        if getattr(self, "proje_salt_okunur", False):
            return
        self.btn_hesaptan_on_deger.configure(state="normal" if on_durum == "verildi" else "disabled")
        if hasattr(self, "btn_asama_degistir"):
            self.btn_asama_degistir.configure(state="disabled" if durum == "bitti" else "normal")
        if hasattr(self, "proje_durumu_yenilemeyi_planla"):
            self.proje_durumu_yenilemeyi_planla()

    def _proje_metni_normalize(self, deger):
        return re.sub(r"[^a-z0-9]", "", str(deger or "").lower().translate(str.maketrans("çğıöşü", "cgiosu")))

    def _tdth_proje_uyarilari(self, kayit):
        uyarilar = []
        baslik = self._proje_metni_normalize(kayit.get("rapor_basligi"))
        ada = self._proje_metni_normalize(self.proje_deger("ADA", ""))
        parsel = self._proje_metni_normalize(self.proje_deger("PARSEL", ""))
        if ada and parsel and (ada not in baslik or parsel not in baslik):
            uyarilar.append("PDF rapor başlığı mevcut ada/parseli açıkça doğrulamıyor.")

        for kod, pdf_anahtari in (("ENLEM", "enlem"), ("BOYLAM", "boylam")):
            proje_degeri = self.proje_deger(kod, "")
            pdf_degeri = kayit.get(pdf_anahtari, "")
            if not proje_degeri or not pdf_degeri:
                continue
            try:
                fark = abs(float(str(proje_degeri).replace(",", ".")) - float(str(pdf_degeri).replace(",", ".")))
                if fark > 0.01:
                    uyarilar.append(f"PDF {kod.lower()} değeri mevcut projeyle uyuşmuyor.")
            except ValueError:
                pass

        secili = self.on_zemin_sinifi_var.get().strip().upper() if hasattr(self, "on_zemin_sinifi_var") else ""
        pdf_sinifi = str(kayit.get("zemin_sinifi") or "").upper()
        if secili and pdf_sinifi and secili != pdf_sinifi:
            uyarilar.append(f"PDF zemin sınıfı {pdf_sinifi}, seçilen ön zemin sınıfı {secili}.")
        return uyarilar

    def _benzersiz_tdth_hedefi(self, kaynak_yolu, kaynak_hash):
        proje_yolu = getattr(self, "guncel_dosya_yolu", "")
        proje_klasoru = os.path.dirname(os.path.abspath(proje_yolu))
        hedef_klasor = os.path.join(proje_klasoru, "On_Degerler", "TDTH")
        os.makedirs(hedef_klasor, exist_ok=True)
        dosya_adi = re.sub(r"[^\w .()\-]", "_", os.path.basename(kaynak_yolu), flags=re.UNICODE).strip() or "TDTH.pdf"
        kok, uzanti = os.path.splitext(dosya_adi)
        uzanti = uzanti or ".pdf"
        hedef = os.path.join(hedef_klasor, kok + uzanti)
        sira = 2
        while os.path.exists(hedef):
            try:
                from on_deger import dosya_sha256
                if dosya_sha256(hedef) == kaynak_hash:
                    return hedef
            except OSError:
                pass
            hedef = os.path.join(hedef_klasor, f"{kok}_{sira}{uzanti}")
            sira += 1
        return hedef

    def tdth_pdf_sec(self):
        if not self.degisiklik_izni_kontrol_et("TDTH PDF seçimi"):
            return
        if not getattr(self, "guncel_dosya_yolu", None):
            if not messagebox.askyesno(
                "Proje Kaydı Gerekli",
                "TDTH PDF'nin proje klasörüne kopyalanabilmesi için proje önce kaydedilmelidir. Şimdi kaydedilsin mi?",
            ):
                return
            if not self.farkli_kaydet():
                return
        kaynak = filedialog.askopenfilename(title="Sismik Tehlike Haritası Raporunu Seç", filetypes=[("PDF", "*.pdf")])
        if not kaynak:
            return
        try:
            kayit = tdth_pdf_bilgilerini_oku(kaynak)
            uyarilar = self._tdth_proje_uyarilari(kayit)
            if uyarilar:
                if not messagebox.askyesno(
                    "TDTH Uyuşma Uyarısı",
                    "\n".join(f"• {x}" for x in uyarilar) + "\n\nDosya yine de kullanılsın mı?",
                ):
                    return
                gerekce = simpledialog.askstring(
                    "TDTH Uyarı Gerekçesi",
                    "Uyuşmazlığa rağmen bu PDF'nin kullanılma gerekçesini yazın:",
                    parent=self.root,
                )
                if gerekce is None or not gerekce.strip():
                    messagebox.showwarning("TDTH PDF", "Uyarılı PDF için gerekçe girilmeden aktarım yapılamaz.")
                    return
                uyarilar.append("Kullanıcı gerekçesi: " + gerekce.strip())
            hedef = self._benzersiz_tdth_hedefi(kaynak, kayit["sha256"])
            if os.path.normcase(os.path.abspath(kaynak)) != os.path.normcase(os.path.abspath(hedef)):
                self.atomik_dosya_kopyala(kaynak, hedef)
            eski_aktif = (normalize_tdth(getattr(self, "tdth_verisi", {})).get("aktif") or {}).get("pdf_yolu", "")
            kayit["pdf_yolu"] = os.path.abspath(hedef)
            self.tdth_verisi = tdth_kaydi_etkinlestir(self.tdth_verisi, kayit, uyarilar)

            pdf_sinifi = kayit.get("zemin_sinifi", "")
            if not self.on_zemin_sinifi_var.get().strip() and pdf_sinifi in ZEMIN_SINIFLARI:
                self.on_zemin_sinifi_var.set(pdf_sinifi)
            self._tdth_rapor_alanlarini_uygula(kayit)

            if eski_aktif and os.path.normcase(eski_aktif) != os.path.normcase(kayit["pdf_yolu"]):
                self.ekler["TDTH"] = [
                    ek for ek in self.ekler.get("TDTH", [])
                    if os.path.normcase(ek.get("yol", "")) != os.path.normcase(eski_aktif)
                ]
            self.ek_dosyayi_listeye_ekle("TDTH", "Sismik Tehlike Haritası Detay Raporu", kayit["pdf_yolu"])
            self.on_deger_ekranini_guncelle()
            self.durum_mesaji_yaz("TDTH PDF projeye aktarıldı", os.path.basename(kayit["pdf_yolu"]))
        except Exception as exc:
            self.hata_kaydet("TDTH PDF içe aktarılamadı", exc)
            messagebox.showerror("TDTH PDF", f"PDF içe aktarılamadı:\n{exc}")

    def _tdth_rapor_alanlarini_uygula(self, kayit):
        degerler = kayit.get("degerler", {})
        eslesmeler = {
            "PGA": "PGA", "PGV": "PGV", "SS": "SS", "S1": "S1", "SDS": "SDS", "SD1": "SD1"
        }
        for kaynak, hedef in eslesmeler.items():
            entry = self.veri_alanlari.get(hedef)
            if entry and degerler.get(kaynak):
                entry.delete(0, tk.END)
                entry.insert(0, degerler[kaynak])
        zemin = kayit.get("zemin_sinifi", "")
        entry = self.veri_alanlari.get("YEREL_ZEMIN_SINIFI")
        if entry and zemin:
            entry.delete(0, tk.END)
            entry.insert(0, zemin)

    def tdth_pdf_ac(self):
        aktif = normalize_tdth(getattr(self, "tdth_verisi", {})).get("aktif") or {}
        yol = aktif.get("pdf_yolu", "")
        if not yol or not os.path.isfile(yol):
            messagebox.showwarning("TDTH PDF", "Aktif TDTH PDF dosyası bulunamadı.")
            return
        try:
            os.startfile(yol)
        except OSError as exc:
            messagebox.showerror("TDTH PDF", f"PDF açılamadı:\n{exc}")

    def on_zemin_sinifi_degisti(self, _event=None):
        if getattr(self, "proje_salt_okunur", False):
            return
        self.tdth_verisi = tdth_zemin_sinifi_guncelle(self.tdth_verisi, self.on_zemin_sinifi_var.get())
        self.on_deger_ekranini_guncelle()

    def nihai_zemin_sinifi_degisti(self, _event=None):
        if getattr(self, "proje_salt_okunur", False):
            return
        entry = self.veri_alanlari.get("YEREL_ZEMIN_SINIFI")
        if not entry:
            return
        self.tdth_verisi = tdth_zemin_sinifi_guncelle(self.tdth_verisi, entry.get())
        self.on_deger_ekranini_guncelle()

    def on_deger_kaydet(self):
        if not self.degisiklik_izni_kontrol_et("Ön değer kaydı"):
            return
        tdth = normalize_tdth(getattr(self, "tdth_verisi", {}))
        aktif = tdth.get("aktif") or {}
        if tdth.get("durum") in {"eksik", "yenilenmeli"}:
            messagebox.showwarning("Ön Değer", "Önce seçilen zemin sınıfıyla uyumlu TDTH PDF eklenmelidir.")
            return
        try:
            self.on_deger_verisi, revizyon = on_deger_revizyonu_ekle(
                self.on_deger_verisini_topla(),
                self.on_qt_var.get(),
                self.on_ks_var.get(),
                self.on_zemin_sinifi_var.get(),
                self.on_aciklama_var.get(),
                aktif.get("sha256", ""),
            )
            durum = self.is_akisi_verisi.get("durum", "yeni")
            if durum == "yeni":
                self.is_akisi_verisi = is_durumu_degistir(
                    self.is_akisi_verisi, "on_deger_verildi", "Ön değer kaydı oluşturuldu."
                )
            self.on_deger_ekranini_guncelle()
            self.durum_mesaji_yaz("Ön değer revizyonu kaydedildi")
            messagebox.showinfo("Ön Değer", f"Ön değer kaydedildi. Toplam revizyon: {len(self.on_deger_verisi['revizyonlar'])}")
        except ValueError as exc:
            messagebox.showwarning("Ön Değer", str(exc))

    def hesap_sonuclarini_on_degere_al(self):
        if not self.degisiklik_izni_kontrol_et("Hesap sonuçlarını ön değere aktarma"):
            return
        if on_deger_durumu(getattr(self, "on_deger_verisi", {})) != "verildi":
            messagebox.showwarning(
                "Ön Değer",
                "Bu işte daha önce ön değer verilmediği için nihai sonuçlar geriye dönük ön değer olarak kaydedilemez.",
            )
            return
        qt = self.entry_qt_nihai.get().strip() if hasattr(self, "entry_qt_nihai") else ""
        ks = self.entry_ks_nihai.get().strip() if hasattr(self, "entry_ks_nihai") else ""
        if not qt or not ks:
            messagebox.showwarning("Ön Değer", "Önce nihai qₜ ve kₛ değerlerini oluşturun.")
            return
        self.on_qt_var.set(qt)
        self.on_ks_var.set(ks)
        nihai_zemin = self.proje_deger("YEREL_ZEMIN_SINIFI", "").strip().upper()
        if nihai_zemin in ZEMIN_SINIFLARI:
            self.on_zemin_sinifi_var.set(nihai_zemin)
            self.on_zemin_sinifi_degisti()
        if not self.on_aciklama_var.get().strip():
            self.on_aciklama_var.set("Hesap sonuçlarına göre güncellendi.")
        messagebox.showinfo(
            "Ön Değer",
            "Hesap sonuçları ön değer alanlarına alındı. Geçmişe eklemek için Ön Değeri Kaydet'e basın.",
        )

    def on_deger_gecmisini_goster(self):
        pencere = tk.Toplevel(self.root)
        pencere.title("Ön Değer Geçmişi")
        pencere.geometry("900x430")
        pencere.transient(self.root)
        kolonlar = ("rev", "tarih", "qt", "ks", "zemin", "aciklama")
        tree = ttk.Treeview(pencere, columns=kolonlar, show="headings")
        basliklar = {"rev": "Rev.", "tarih": "Tarih", "qt": "qₜ", "ks": "kₛ", "zemin": "Zemin", "aciklama": "Açıklama"}
        for kod in kolonlar:
            tree.heading(kod, text=basliklar[kod])
            tree.column(kod, width=70 if kod in {"rev", "qt", "ks", "zemin"} else 190, anchor="center" if kod != "aciklama" else "w")
        for index, item in enumerate(normalize_on_deger(self.on_deger_verisi).get("revizyonlar", []), start=1):
            tree.insert("", "end", values=(index, item.get("tarih", ""), item.get("qt", ""), item.get("ks", ""), item.get("zemin_sinifi", ""), item.get("aciklama", "")))
        tree.pack(fill="both", expand=True, padx=12, pady=12)
        ttk.Button(pencere, text="Kapat", command=pencere.destroy).pack(pady=(0, 12))

    def is_asamasini_belirle(self):
        if not self.degisiklik_izni_kontrol_et("İş aşaması seçimi"):
            return
        mevcut_durum = normalize_is_akisi(self.is_akisi_verisi).get("durum", "yeni")
        secim = self._asama_secim_penceresi()
        if not secim or secim == mevcut_durum:
            return
        if mevcut_durum == "belirlenmedi":
            self.is_akisi_verisi = is_durumu_degistir(
                self.is_akisi_verisi, secim, "Eski proje için başlangıç aşaması seçildi.", zorla=True
            )
            self.on_deger_ekranini_guncelle()
        elif secim == "yazim_asamasinda":
            self.yazim_asamasina_gec()
        elif secim == "bitti":
            self.projeyi_bitir()

    def eski_proje_asama_secimi(self):
        return self._asama_secim_penceresi(mevcut="belirlenmedi")

    def _asama_secim_penceresi(self, mevcut=None):
        if mevcut is None:
            mevcut = normalize_is_akisi(getattr(self, "is_akisi_verisi", {})).get("durum", "yeni")
        secenek_kodlari = {
            "belirlenmedi": ("yeni", "on_deger_verildi", "yazim_asamasinda", "bitti"),
            "yeni": ("yazim_asamasinda",),
            "on_deger_verildi": ("yazim_asamasinda",),
            "yazim_asamasinda": ("bitti",),
            "duzeltme_asamasinda": ("bitti",),
            "bitti": (),
        }.get(mevcut, ())
        if not secenek_kodlari:
            messagebox.showinfo("İş Aşaması", "Bu aşamadan yapılabilecek bir geçiş bulunmuyor.")
            return None

        pencere = tk.Toplevel(self.root)
        pencere.title("Aşamayı Değiştir")
        pencere.geometry("430x180")
        pencere.resizable(False, False)
        pencere.transient(self.root)
        pencere.grab_set()
        sonuc = {"deger": None}
        ttk.Label(pencere, text="Geçilecek iş aşamasını seçin:", wraplength=390).pack(anchor="w", padx=18, pady=(18, 10))
        secenek_etiketleri = tuple(IS_DURUMLARI[kod] for kod in secenek_kodlari)
        combo = ttk.Combobox(
            pencere,
            state="readonly",
            values=secenek_etiketleri,
        )
        combo.pack(fill="x", padx=18)
        combo.set(secenek_etiketleri[0])
        ters = {v: k for k, v in IS_DURUMLARI.items()}

        def onayla():
            sonuc["deger"] = ters.get(combo.get())
            pencere.destroy()

        ttk.Button(pencere, text="Seç", command=onayla, bootstyle="primary").pack(pady=16)
        pencere.protocol("WM_DELETE_WINDOW", pencere.destroy)
        self.root.wait_window(pencere)
        return sonuc["deger"]

    def yazim_asamasina_gec(self):
        if not self.degisiklik_izni_kontrol_et("Aşama değişikliği"):
            return
        mevcut_durum = self.is_akisi_verisi.get("durum", "yeni")
        if mevcut_durum == "yeni" and not messagebox.askyesno(
            "Doğrudan Yazım",
            "Bu işte ön değer verilmeden doğrudan yazım aşamasına geçilecek. Devam edilsin mi?",
        ):
            return
        try:
            self.is_akisi_verisi = is_durumu_degistir(
                self.is_akisi_verisi,
                "yazim_asamasinda",
                "Ön değer verilmeden doğrudan yazım başlatıldı."
                if mevcut_durum == "yeni"
                else "Yazım çalışması başlatıldı.",
            )
            self.on_deger_ekranini_guncelle()
        except ValueError as exc:
            messagebox.showwarning("İş Aşaması", str(exc))

    def projeyi_bitir(self):
        if not self.degisiklik_izni_kontrol_et("Projeyi bitirme"):
            return
        eksikler = []
        tdth = normalize_tdth(self.tdth_verisi)
        if tdth.get("durum") in {"eksik", "yenilenmeli"}:
            eksikler.append("Geçerli ve güncel TDTH PDF")
        if hasattr(self, "entry_qt_nihai") and not self.entry_qt_nihai.get().strip():
            eksikler.append("Nihai qₜ")
        if hasattr(self, "entry_ks_nihai") and not self.entry_ks_nihai.get().strip():
            eksikler.append("Nihai kₛ")
        rapor_yolu = self.is_akisi_verisi.get("son_rapor_yolu", "")
        if not rapor_yolu or not os.path.isfile(rapor_yolu):
            eksikler.append("Oluşturulmuş Word raporu")
        nihai_pdf = self.is_akisi_verisi.get("son_nihai_pdf_yolu", "")
        if not nihai_pdf or not os.path.isfile(nihai_pdf):
            eksikler.append("TDTH ekini içeren nihai PDF")
        if eksikler:
            messagebox.showwarning("Proje Tamamlanamadı", "Eksik bilgiler:\n• " + "\n• ".join(eksikler))
            return
        if not messagebox.askyesno("Projeyi Bitir", "Proje bitmiş olarak işaretlenecek ve sonraki açılışta kilitlenecek. Devam edilsin mi?"):
            return
        onceki_akis = copy.deepcopy(self.is_akisi_verisi)
        try:
            self.is_akisi_verisi = is_durumu_degistir(self.is_akisi_verisi, "bitti", "Proje tamamlandı.")
            self.on_deger_ekranini_guncelle()
            if not self.kaydet():
                self.is_akisi_verisi = onceki_akis
                self.on_deger_ekranini_guncelle()
                return
            self.proje_salt_okunur_ayarla(True)
            messagebox.showinfo("İş Aşaması", "Proje Bitti durumuna geçirildi ve izleme modunda kilitlendi.")
        except ValueError as exc:
            messagebox.showwarning("İş Aşaması", str(exc))

    def bitmis_proje_acilis_secimi(self):
        pencere = tk.Toplevel(self.root)
        pencere.title("Bitmiş Proje")
        pencere.geometry("520x220")
        pencere.resizable(False, False)
        pencere.transient(self.root)
        pencere.grab_set()
        sonuc = {"secim": None}
        ttk.Label(pencere, text="Bu proje tamamlanmış ve değişikliklere karşı kilitlenmiştir.", style="AltBaslik.TLabel", wraplength=470).pack(padx=22, pady=(24, 8))
        ttk.Label(pencere, text="İzleme yalnızca görüntüleme sağlar. Düzeltme yeni bir revizyon başlatır.", style="Muted.TLabel", wraplength=470).pack(padx=22, pady=(0, 22))
        butonlar = ttk.Frame(pencere)
        butonlar.pack()

        def sec(deger):
            sonuc["secim"] = deger
            pencere.destroy()

        ttk.Button(butonlar, text="İzleme Modunda Aç", command=lambda: sec("izleme"), bootstyle="secondary", width=20).pack(side="left", padx=8)
        ttk.Button(butonlar, text="Düzeltme Başlat", command=lambda: sec("duzeltme"), bootstyle="warning", width=20).pack(side="left", padx=8)
        pencere.protocol("WM_DELETE_WINDOW", pencere.destroy)
        self.root.wait_window(pencere)
        if sonuc["secim"] == "duzeltme":
            neden = simpledialog.askstring("Düzeltme Nedeni", "Düzeltme nedenini yazın:", parent=self.root)
            if neden is None:
                return None, ""
            neden = neden.strip()
            if not neden:
                messagebox.showwarning("Düzeltme", "Düzeltme nedeni boş bırakılamaz.")
                return None, ""
            return "duzeltme", neden
        return sonuc["secim"], ""

    def proje_salt_okunur_ayarla(self, aktif):
        aktif = bool(aktif)
        onceki = bool(getattr(self, "proje_salt_okunur", False))
        if aktif == onceki and hasattr(self, "_salt_okunur_widget_states"):
            return
        izinli = set(getattr(self, "salt_okunurda_acik_widgetlar", []))
        if aktif:
            self._salt_okunur_widget_states = {}
            duzenlenebilir_siniflar = (
                ttk.Entry, ttk.Combobox, ttk.Button, ttk.Checkbutton, ttk.Radiobutton, ttk.Scale,
                tk.Entry, tk.Text, tk.Button, tk.Checkbutton, tk.Radiobutton, tk.Scale, tk.Listbox,
            )

            def kilitle(widget):
                for child in widget.winfo_children():
                    if isinstance(child, duzenlenebilir_siniflar) and child not in izinli:
                        try:
                            self._salt_okunur_widget_states[str(child)] = (child, str(child.cget("state")))
                            child.configure(state="disabled")
                        except (tk.TclError, KeyError):
                            pass
                    kilitle(child)

            kilitle(self.root)
            for widget in izinli:
                try:
                    widget.configure(state="normal")
                except tk.TclError:
                    pass
        else:
            for widget, durum in getattr(self, "_salt_okunur_widget_states", {}).values():
                try:
                    if widget.winfo_exists():
                        widget.configure(state=durum)
                except tk.TclError:
                    pass
            self._salt_okunur_widget_states = {}
        self.proje_salt_okunur = aktif
        menu = getattr(self, "menu_cubugu", None)
        if menu is not None:
            try:
                menu.entryconfigure("Araçlar", state="disabled" if aktif else "normal")
            except tk.TclError:
                pass
        dosya_menusu = getattr(self, "dosya_menusu", None)
        if dosya_menusu is not None:
            for etiket in ("Kaydet", "Farklı Kaydet"):
                try:
                    dosya_menusu.entryconfigure(etiket, state="disabled" if aktif else "normal")
                except tk.TclError:
                    pass
        if hasattr(self, "on_deger_ekranini_guncelle"):
            self.on_deger_ekranini_guncelle()
        if hasattr(self, "durum_mesaji_yaz"):
            self.durum_mesaji_yaz("Proje izleme modunda" if aktif else "Proje düzenleme modunda")
        if hasattr(self, "proje_durum_seridi_guncelle"):
            self.proje_durum_seridi_guncelle()
