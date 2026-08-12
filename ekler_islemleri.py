import os
from tkinter import ttk, filedialog, messagebox, simpledialog

from ekler import (
    ek_dosya_turu as ekler_dosya_turu,
    ek_durumunu_hazirla,
    ek_taahhutname_mi,
    ek_kategori_durumunu_hazirla as ekler_kategori_durumunu_hazirla,
    ekleri_denetle,
    ekler_pdf_olustur as ekler_pdf_birlestir,
)


class EklerIslemleri:
    def __init__(self, app):
        object.__setattr__(self, "app", app)

    def __getattr__(self, name):
        return getattr(self.app, name)

    def __setattr__(self, name, value):
        if name == "app":
            object.__setattr__(self, name, value)
        else:
            setattr(self.app, name, value)

    def sekme10_ekler(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="10. Ekler")
        page = ttk.Frame(frame, padding=16)
        page.pack(fill="both", expand=True)

        ust = ttk.Frame(page)
        ust.pack(fill="x", pady=(0, 10))
        ttk.Label(ust, text="Ekler", style="Baslik.TLabel").pack(side="left")
        self.ek_ozet_label = ttk.Label(ust, text="", style="AltBaslik.TLabel")
        self.ek_ozet_label.pack(side="right")
        ttk.Separator(page).pack(fill="x", pady=(0, 12))

        akis = ttk.LabelFrame(page, text="İş Akışı", padding=(10, 8), bootstyle="secondary")
        akis.pack(fill="x", pady=(0, 10))
        ttk.Button(akis, text="Ekleri Kontrol Et", command=self.ek_kontrol_ozeti_goster, bootstyle="info").pack(side="left", padx=(0, 8))
        ttk.Button(akis, text="Mühendis Bilgileri", command=self.taahhut_bilgilerini_duzenle, bootstyle="secondary outline").pack(side="left", padx=(0, 8))
        ttk.Button(akis, text="Taahhütnameleri Oluştur", command=self.taahhutnameleri_olustur, bootstyle="success outline").pack(side="left", padx=(0, 8))
        ttk.Button(akis, text="Sıralı EKLER PDF Oluştur", command=self.ekler_pdf_olustur, bootstyle="success").pack(side="left")

        durum_frame = ttk.LabelFrame(page, text="Kategori Durumu", padding=(8, 8), bootstyle="secondary")
        durum_frame.pack(fill="x", pady=(0, 10))
        for kategori in self.ek_kategorileri:
            lbl = ttk.Label(durum_frame, text=f"{kategori}: Eksik", padding=(8, 4), relief="groove", bootstyle="secondary")
            lbl.pack(side="left", padx=(0, 6))
            self.ek_durum_etiketleri[kategori] = lbl

        self.ekler_notebook = ttk.Notebook(page)
        self.ekler_notebook.pack(fill="both", expand=True)
        self.ek_kategori_frame = {}

        for kategori in self.ek_kategorileri:
            kat_frame = ttk.Frame(self.ekler_notebook, padding=8)
            self.ekler_notebook.add(kat_frame, text=f"{kategori} (0)")
            self.ek_kategori_frame[kategori] = kat_frame

            toolbar = ttk.Frame(kat_frame)
            toolbar.pack(fill="x", pady=(0, 8))
            ttk.Button(toolbar, text="Dosya Ekle", command=lambda k=kategori: self.ek_dosya_ekle(k), bootstyle="primary").pack(side="left", padx=(0, 6))
            ttk.Button(toolbar, text="Seçiliyi Sil", command=lambda k=kategori: self.ek_secili_sil(k), bootstyle="secondary").pack(side="left", padx=(0, 6))
            ttk.Button(toolbar, text="Yukarı", command=lambda k=kategori: self.ek_secili_tasi(k, -1), bootstyle="secondary").pack(side="left", padx=(0, 6))
            ttk.Button(toolbar, text="Aşağı", command=lambda k=kategori: self.ek_secili_tasi(k, 1), bootstyle="secondary").pack(side="left", padx=(0, 6))
            ttk.Button(toolbar, text="Başlık Düzenle", command=lambda k=kategori: self.ek_baslik_duzenle(k), bootstyle="info").pack(side="left", padx=(0, 6))

            kolonlar = ("Sıra", "Başlık", "Tür", "Durum", "Dosya")
            scroll_y = ttk.Scrollbar(kat_frame, orient="vertical")
            scroll_y.pack(side="right", fill="y")
            scroll_x = ttk.Scrollbar(kat_frame, orient="horizontal")
            scroll_x.pack(side="bottom", fill="x")

            tree = ttk.Treeview(
                kat_frame,
                columns=kolonlar,
                show="headings",
                height=10,
                style="Ekler.Treeview",
                yscrollcommand=scroll_y.set,
                xscrollcommand=scroll_x.set,
            )
            scroll_y.config(command=tree.yview)
            scroll_x.config(command=tree.xview)

            genislikler = {"Sıra": 55, "Başlık": 260, "Tür": 80, "Durum": 110, "Dosya": 620}
            for col in kolonlar:
                tree.heading(col, text=col)
                anchor = "center" if col in ("Sıra", "Tür", "Durum") else "w"
                tree.column(col, width=genislikler[col], anchor=anchor, stretch=(col in ("Başlık", "Dosya")))
            tree.tag_configure("eksik", background="#ffecec")
            tree.tag_configure("donusum", background="#fff6dc")
            tree.tag_configure("otomatik_donusum", background="#eaf4ff")
            tree.tag_configure("gecersiz", background="#ffd6d6")
            tree.tag_configure("var", background="#ffffff")
            tree.pack(fill="both", expand=True)
            tree.bind("<Double-1>", lambda e, k=kategori: self.ek_baslik_duzenle(k))
            self.ek_treeviewler[kategori] = tree

        self.ek_listeleri_guncelle()

    def taahhutname_paneli_ekle(self, parent):
        return self.app.taahhutname_paneli_ekle(parent)

    def ek_dosyayi_listeye_ekle(self, kategori, baslik, yol):
        if ek_taahhutname_mi({"baslik": baslik, "yol": yol}):
            return
        liste = self.ekler.setdefault(kategori, [])
        for ek in liste:
            if os.path.normcase(ek.get("yol", "")) == os.path.normcase(yol):
                ek["baslik"] = baslik
                self.ek_listeleri_guncelle(kategori)
                return
        liste.append({"baslik": baslik, "yol": yol})
        self.ek_listeleri_guncelle(kategori)

    def ek_dosya_turu(self, yol):
        return ekler_dosya_turu(yol)

    def ek_kategori_durumunu_hazirla(self, kategori):
        return ekler_kategori_durumunu_hazirla(self.ekler, kategori)

    def ek_dosya_ekle(self, kategori):
        if hasattr(self, "degisiklik_izni_kontrol_et") and not self.degisiklik_izni_kontrol_et("Ek dosya ekleme"):
            return
        dosyalar = filedialog.askopenfilenames(
            title=f"{kategori} dosyası ekle",
            filetypes=[
                ("Ek Dosyaları", "*.pdf *.jpg *.jpeg *.png *.docx *.doc *.xlsx *.xls"),
                ("PDF", "*.pdf"),
                ("Görseller", "*.jpg *.jpeg *.png"),
                ("Word", "*.docx *.doc"),
                ("Excel", "*.xlsx *.xls"),
                ("Tüm Dosyalar", "*.*"),
            ],
        )
        if not dosyalar:
            return
        for yol in dosyalar:
            baslik = os.path.splitext(os.path.basename(yol))[0]
            self.ek_dosyayi_listeye_ekle(kategori, baslik, os.path.abspath(yol))
        self.ek_listeleri_guncelle(kategori)

    def ek_secili_index(self, kategori):
        tree = self.ek_treeviewler.get(kategori)
        if not tree:
            return None
        secim = tree.selection()
        if not secim:
            messagebox.showwarning("Uyarı", "Lütfen önce bir ek dosya seçin.")
            return None
        return tree.index(secim[0])

    def ek_secili_sil(self, kategori):
        if hasattr(self, "degisiklik_izni_kontrol_et") and not self.degisiklik_izni_kontrol_et("Ek dosya silme"):
            return
        index = self.ek_secili_index(kategori)
        if index is None:
            return
        if messagebox.askyesno("Ek Sil", "Seçili ek dosya listeden kaldırılsın mı?"):
            del self.ekler[kategori][index]
            self.ek_listeleri_guncelle(kategori)

    def ek_secili_tasi(self, kategori, yon):
        if hasattr(self, "degisiklik_izni_kontrol_et") and not self.degisiklik_izni_kontrol_et("Ek sırası değiştirme"):
            return
        index = self.ek_secili_index(kategori)
        if index is None:
            return
        yeni_index = index + yon
        if yeni_index < 0 or yeni_index >= len(self.ekler.get(kategori, [])):
            return
        liste = self.ekler[kategori]
        liste[index], liste[yeni_index] = liste[yeni_index], liste[index]
        self.ek_listeleri_guncelle(kategori, secili_index=yeni_index)

    def ek_baslik_duzenle(self, kategori):
        if hasattr(self, "degisiklik_izni_kontrol_et") and not self.degisiklik_izni_kontrol_et("Ek başlığı düzenleme"):
            return
        index = self.ek_secili_index(kategori)
        if index is None:
            return
        ek = self.ekler[kategori][index]
        yeni_baslik = simpledialog.askstring("Başlık Düzenle", "Ek başlığı:", initialvalue=ek.get("baslik", ""))
        if yeni_baslik is None:
            return
        ek["baslik"] = yeni_baslik.strip() or os.path.splitext(os.path.basename(ek.get("yol", "")))[0]
        self.ek_listeleri_guncelle(kategori, secili_index=index)

    def ek_listeleri_guncelle(self, kategori=None, secili_index=None):
        kategoriler = [kategori] if kategori else self.ek_kategorileri
        for kat in kategoriler:
            tree = self.ek_treeviewler.get(kat)
            if not tree:
                continue
            for item in tree.get_children():
                tree.delete(item)
            for i, ek in enumerate(self.ekler.get(kat, []), start=1):
                yol = ek.get("yol", "")
                durum, tag = ek_durumunu_hazirla(ek)
                item = tree.insert(
                    "",
                    "end",
                    values=(i, ek.get("baslik", ""), self.ek_dosya_turu(yol), durum, yol),
                    tags=(tag,),
                )
                if secili_index is not None and i - 1 == secili_index:
                    tree.selection_set(item)
                    tree.focus(item)
                    tree.see(item)

            if hasattr(self, "ekler_notebook") and hasattr(self, "ek_kategori_frame"):
                self.ekler_notebook.tab(self.ek_kategori_frame[kat], text=f"{kat} ({len(self.ekler.get(kat, []))})")
            durum_metni, bootstyle = self.ek_kategori_durumunu_hazirla(kat)
            lbl = self.ek_durum_etiketleri.get(kat)
            if lbl:
                lbl.config(text=f"{kat}: {durum_metni}", bootstyle=bootstyle)

        denetim = ekleri_denetle(self.ekler, self.ek_kategorileri)
        if hasattr(self, "ek_ozet_label"):
            if denetim["eksik_dosyalar"]:
                self.ek_ozet_label.config(text=f"Toplam {denetim['toplam']} dosya / {len(denetim['eksik_dosyalar'])} dosya bulunamadı", bootstyle="warning")
            elif denetim["donusum_gerekenler"]:
                self.ek_ozet_label.config(text=f"Toplam {denetim['toplam']} dosya / {len(denetim['donusum_gerekenler'])} dosya PDF'e çevrilmeli", bootstyle="warning")
            elif denetim["bos_kategoriler"]:
                self.ek_ozet_label.config(text=f"Toplam {denetim['toplam']} dosya / Eksik kategori: {len(denetim['bos_kategoriler'])}", bootstyle="danger")
            elif denetim["otomatik_donusumler"]:
                self.ek_ozet_label.config(
                    text=(
                        f"Toplam {denetim['toplam']} dosya / "
                        f"{len(denetim['otomatik_donusumler'])} Word otomatik PDF'e çevrilecek"
                    ),
                    bootstyle="info",
                )
            else:
                self.ek_ozet_label.config(text=f"Toplam {denetim['toplam']} dosya / PDF hazır", bootstyle="success")

    def ek_denetim_mesaji(self, denetim):
        satirlar = [
            f"Toplam ek dosya: {denetim['toplam']}",
            f"PDF'e hazır dosya: {denetim['pdf_hazir']}",
        ]
        if denetim["bos_kategoriler"]:
            satirlar.append("Boş kategoriler: " + ", ".join(denetim["bos_kategoriler"]))
        if denetim["eksik_dosyalar"]:
            satirlar.append("")
            satirlar.append("Dosya bulunamayan ekler:")
            for ek in denetim["eksik_dosyalar"][:12]:
                satirlar.append(f"- {ek['kategori']} {ek['sira']}: {ek['baslik']}")
            if len(denetim["eksik_dosyalar"]) > 12:
                satirlar.append(f"... ve {len(denetim['eksik_dosyalar']) - 12} ek daha")
        if denetim["donusum_gerekenler"]:
            satirlar.append("")
            satirlar.append("PDF'e çevrilmesi gereken ekler:")
            for ek in denetim["donusum_gerekenler"][:12]:
                satirlar.append(f"- {ek['kategori']} {ek['sira']}: {ek['baslik']} ({ek['tur']})")
            if len(denetim["donusum_gerekenler"]) > 12:
                satirlar.append(f"... ve {len(denetim['donusum_gerekenler']) - 12} ek daha")
        if denetim.get("otomatik_donusumler"):
            satirlar.append("")
            satirlar.append("EKLER PDF hazırlanırken otomatik çevrilecek Word ekleri:")
            for ek in denetim["otomatik_donusumler"][:12]:
                satirlar.append(f"- {ek['kategori']} {ek['sira']}: {ek['baslik']} ({ek['tur']})")
            if len(denetim["otomatik_donusumler"]) > 12:
                satirlar.append(f"... ve {len(denetim['otomatik_donusumler']) - 12} ek daha")
        if denetim.get("gecersiz_dosyalar"):
            satirlar.append("")
            satirlar.append("Bozuk veya okunamayan ekler:")
            for ek in denetim["gecersiz_dosyalar"][:12]:
                satirlar.append(f"- {ek['kategori']} {ek['sira']}: {ek['baslik']} ({ek['durum']})")
        if (
            not denetim["eksik_dosyalar"]
            and not denetim["donusum_gerekenler"]
            and not denetim.get("gecersiz_dosyalar")
        ):
            satirlar.append("")
            satirlar.append("PDF üretimi için dosya durumları uygun.")
        return "\n".join(satirlar)

    def ek_kontrol_ozeti_goster(self):
        denetim = ekleri_denetle(self.ekler, self.ek_kategorileri, derin=True)
        mesaj = self.ek_denetim_mesaji(denetim)
        if (
            denetim["eksik_dosyalar"]
            or denetim["donusum_gerekenler"]
            or denetim.get("gecersiz_dosyalar")
            or denetim["bos_kategoriler"]
        ):
            messagebox.showwarning("EKLER Kontrolü", mesaj)
        else:
            messagebox.showinfo("EKLER Kontrolü", mesaj)
        if hasattr(self, "durum_mesaji_yaz"):
            self.durum_mesaji_yaz("Ekler kontrol edildi")

    def ekler_pdf_olustur(self):
        denetim = ekleri_denetle(self.ekler, self.ek_kategorileri, derin=True)
        if denetim["toplam"] == 0:
            messagebox.showwarning("EKLER PDF", "PDF oluşturmak için ek dosya bulunmuyor.")
            return
        kullanilacak_etiketler = {"var", "otomatik_donusum"}
        if (
            denetim["eksik_dosyalar"]
            or denetim["donusum_gerekenler"]
            or denetim.get("gecersiz_dosyalar")
        ):
            devam = messagebox.askyesno(
                "EKLER PDF",
                self.ek_denetim_mesaji(denetim)
                + "\n\nEksik, dönüştürülemeyen veya bozuk ekler atlanarak "
                "geçerli bulunan eklerle devam edilsin mi?",
            )
            if not devam:
                return
        kullanilabilir_ekler = {kategori: [] for kategori in self.ek_kategorileri}
        for ek in denetim["sirali"]:
            if ek["tag"] in kullanilacak_etiketler:
                kullanilabilir_ekler[ek["kategori"]].append(
                    {"baslik": ek["baslik"], "yol": ek["yol"]}
                )
        if not any(kullanilabilir_ekler.values()):
            messagebox.showwarning(
                "EKLER PDF",
                "Geçerli bir ek dosyası bulunmadığı için PDF oluşturulamadı.",
            )
            return
        if denetim["bos_kategoriler"]:
            devam = messagebox.askyesno(
                "Boş Kategoriler",
                "Bazı ek kategorileri boş:\n"
                + ", ".join(denetim["bos_kategoriler"])
                + "\n\nYine de sıralı EKLER PDF oluşturulsun mu?"
            )
            if not devam:
                return

        kayit_yolu = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile="K-1_EKLER.pdf",
            title="Sıralı EKLER PDF dosyasını kaydet",
            confirmoverwrite=False,
        )
        if not kayit_yolu:
            return
        if os.path.exists(kayit_yolu):
            karar = messagebox.askyesnocancel(
                "Mevcut PDF",
                "Seçilen PDF zaten var.\n\n"
                "Evet: güvenli biçimde üzerine yaz\n"
                "Hayır: numaralı yeni bir dosya oluştur",
            )
            if karar is None:
                return
            if not karar:
                kok, uzanti = os.path.splitext(kayit_yolu)
                sira = 2
                aday = f"{kok}_{sira}{uzanti}"
                while os.path.exists(aday):
                    sira += 1
                    aday = f"{kok}_{sira}{uzanti}"
                kayit_yolu = aday

        try:
            sonuc = ekler_pdf_birlestir(
                kullanilabilir_ekler,
                self.ek_kategorileri,
                kayit_yolu,
            )
            messagebox.showinfo(
                "EKLER PDF",
                f"Sıralı EKLER PDF oluşturuldu:\n{sonuc['dosya']}\n\n"
                f"Eklenen dosya: {sonuc['dosya_sayisi']}\n"
                f"Otomatik dönüştürülen Word: {sonuc.get('donusturulen_dosya_sayisi', 0)}\n"
                f"Toplam sayfa: {sonuc['sayfa']}"
            )
            if hasattr(self, "durum_mesaji_yaz"):
                self.durum_mesaji_yaz("EKLER PDF oluşturuldu")
        except RuntimeError as e:
            messagebox.showerror(
                "PDF Kütüphanesi Eksik",
                f"{e}\n\nKurulum için: pip install pypdf"
            )
        except Exception as e:
            self.hata_kaydet("EKLER PDF oluşturulamadı", e)
            messagebox.showerror("EKLER PDF", f"PDF oluşturulamadı:\n{e}")

    def ekler_verisini_topla(self):
        return {
            kategori: [
                {"baslik": ek.get("baslik", ""), "yol": ek.get("yol", "")}
                for ek in self.ekler.get(kategori, [])
                if not ek_taahhutname_mi(ek)
            ]
            for kategori in self.ek_kategorileri
        }

    def ekler_verisini_yerlestir(self, veriler):
        self.ekler = {kategori: [] for kategori in self.ek_kategorileri}
        if isinstance(veriler, dict):
            for kategori in self.ek_kategorileri:
                for ek in veriler.get(kategori, []):
                    if isinstance(ek, dict):
                        temiz_ek = {
                            "baslik": ek.get("baslik", ""),
                            "yol": ek.get("yol", ""),
                        }
                        if not ek_taahhutname_mi(temiz_ek):
                            self.ekler[kategori].append(temiz_ek)
                    elif isinstance(ek, str):
                        temiz_ek = {
                            "baslik": os.path.splitext(os.path.basename(ek))[0],
                            "yol": ek,
                        }
                        if not ek_taahhutname_mi(temiz_ek):
                            self.ekler[kategori].append(temiz_ek)
        self.ek_listeleri_guncelle()
