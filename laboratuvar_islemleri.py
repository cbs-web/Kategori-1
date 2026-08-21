import os
import math
import re
from tkinter import ttk, filedialog, messagebox

from laboratuvar import (
    EksikLaboratuvarSutunu,
    LAB_AC_KOLONLARI,
    LAB_YN_KOLONLARI,
    laboratuvar_dosyasi_oku,
    laboratuvar_numune_anahtari,
    laboratuvar_numune_etiketlerini_uyarla,
    laboratuvar_pano_verisini_donustur,
    laboratuvar_satirlarini_birlestir,
)


LAB_LOG_KOLON_ESLEME = {
    "USCS": "Sınıflama",
    "Wn": "Wn",
    "LL": "LL",
    "PL": "PL",
    "PI": "PI",
    "+4 %": "Çakıl",
    "-200 %": "SiltKil",
}


def lab_isim_normalize(isim):
    return laboratuvar_numune_anahtari(isim)


def lab_derinlik_araligi(derinlik):
    metin = str(derinlik or "").strip().replace(",", ".")
    eslesme = re.fullmatch(
        r"\s*(-?\d+(?:\.\d+)?)\s*[-–—]\s*(-?\d+(?:\.\d+)?)\s*",
        metin,
    )
    if not eslesme:
        return None
    baslangic, bitis = float(eslesme.group(1)), float(eslesme.group(2))
    if not math.isfinite(baslangic) or not math.isfinite(bitis) or bitis <= baslangic:
        return None
    return baslangic, bitis


def lab_derinlik_eslesiyor_mu(birinci, ikinci, tolerans=1e-6):
    aralik_1 = lab_derinlik_araligi(birinci)
    aralik_2 = lab_derinlik_araligi(ikinci)
    if aralik_1 is None or aralik_2 is None:
        return str(birinci).strip() == str(ikinci).strip()
    return all(
        math.isclose(a, b, rel_tol=0.0, abs_tol=tolerans)
        for a, b in zip(aralik_1, aralik_2)
    )


def lab_eslesen_satiri_bul(lab_satirlari, hedef_derinlik):
    return next(
        (
            satir
            for satir in lab_satirlari or []
            if lab_derinlik_eslesiyor_mu(hedef_derinlik, satir.get("orijinal_derinlik", ""))
        ),
        None,
    )


class LaboratuvarIslemleri:
    def __init__(self, app):
        object.__setattr__(self, "app", app)

    def __getattr__(self, name):
        return getattr(self.app, name)

    def __setattr__(self, name, value):
        if name == "app":
            object.__setattr__(self, name, value)
        else:
            setattr(self.app, name, value)

    def sekme9_lab(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="9. Laboratuvar")

        page = ttk.Frame(frame, padding=16)
        page.pack(fill='both', expand=True)

        ust_kontrol_frame = ttk.Frame(page)
        ust_kontrol_frame.pack(fill='x', pady=(0, 10))
        ttk.Label(ust_kontrol_frame, text="Laboratuvar Verileri", style="Baslik.TLabel").pack(side="left")
        ttk.Button(ust_kontrol_frame, text="Laboratuvar Excel'i Yükle", command=self.lab_excel_yukle, bootstyle="primary").pack(side="right")
        ttk.Separator(page).pack(fill="x", pady=(0, 10))

        dosya_frame = ttk.Frame(page)
        dosya_frame.pack(fill='x', pady=(0, 8))
        ttk.Label(dosya_frame, text="Yüklü Dosya:", style="Muted.TLabel").pack(side="left")
        self.lbl_lab_excel = ttk.Label(dosya_frame, text="Yok", style="AltBaslik.TLabel")
        self.lbl_lab_excel.pack(side="left", padx=(6, 0))
        ttk.Label(
            dosya_frame,
            text="LAB_1 veri satırlarını kopyalayıp AÇ veya YN tablosunda Ctrl+V yapabilirsiniz.",
            style="Muted.TLabel",
        ).pack(side="right")

        paned = ttk.PanedWindow(page, orient="vertical")
        paned.pack(fill='both', expand=True)
        
        # --- ÜST YARI: AÇ TABLOSU ---
        frame_ac = ttk.LabelFrame(paned, text="Araştırma Çukuru (AÇ) Laboratuvar Verileri", padding=8, bootstyle="info")
        paned.add(frame_ac, weight=1)
        
        kolonlar_ac = LAB_AC_KOLONLARI
        
        scroll_x_ac = ttk.Scrollbar(frame_ac, orient="horizontal")
        scroll_x_ac.pack(side="bottom", fill="x")
        scroll_y_ac = ttk.Scrollbar(frame_ac, orient="vertical")
        scroll_y_ac.pack(side="right", fill="y")
        
        self.tree_lab_ac = ttk.Treeview(frame_ac, columns=kolonlar_ac, show="headings", height=8, style="Lab.Treeview",
                                     xscrollcommand=scroll_x_ac.set, yscrollcommand=scroll_y_ac.set)
        for col in kolonlar_ac:
            self.tree_lab_ac.heading(col, text=col)
            if col == "Sınıflama":
                width = 110
            elif col in ("No", "Derinlik"):
                width = 95
            else:
                width = 72
            self.tree_lab_ac.column(col, width=width, anchor="center")
            
        scroll_x_ac.config(command=self.tree_lab_ac.xview)
        scroll_y_ac.config(command=self.tree_lab_ac.yview)
        self.tree_lab_ac.pack(expand=True, fill='both')
        self.tree_lab_ac.tag_configure('oddrow', background="white")
        self.tree_lab_ac.tag_configure('evenrow', background="#eaf7ff")
        self.tree_lab_ac.bind("<Double-1>", lambda e, t=self.tree_lab_ac: self.hucre_duzenle(e, t))
        
        def _senkronize_ac_tablo(temizle_eslesmeyen=False):
            """Lab verisini yalnız isim ve gerçek derinlik aralığı eşleştiğinde loga uygular."""
            try:
                lab_data = {}
                for item in self.tree_lab_ac.get_children():
                    vals = self.tree_lab_ac.item(item)["values"]
                    if not vals or len(vals) < 12: continue
                    kuyu_no = str(vals[0]).strip()
                    derinlik = str(vals[1]).strip()
                    if kuyu_no and kuyu_no != "-":
                        t_kuyu = lab_isim_normalize(kuyu_no)
                        if t_kuyu not in lab_data: lab_data[t_kuyu] = []
                        lab_data[t_kuyu].append({
                            "orijinal_derinlik": derinlik,
                            "Çakıl": str(vals[2]).strip(),
                            "SiltKil": str(vals[4]).strip(),
                            "LL": str(vals[5]).strip(),
                            "PL": str(vals[6]).strip(),
                            "PI": str(vals[7]).strip(),
                            "Wn": str(vals[8]).strip(),
                            "Sınıflama": str(vals[11]).strip()
                        })
                        
                for kayit in self.ac_yn_sekme_kayitlari():
                    sekme_ismi = kayit["isim"]
                    t_sekme_ismi = lab_isim_normalize(sekme_ismi)
                    tree_ac = kayit["tree"]
                    ac_verisi_listesi = lab_data.get(t_sekme_ismi, [])
                    for row_id in tree_ac.get_children():
                        vals = tree_ac.item(row_id)["values"]
                        if not vals:
                            continue
                        d_key = str(vals[0]).strip()
                        hedef_veri = lab_eslesen_satiri_bul(ac_verisi_listesi, d_key)
                        if temizle_eslesmeyen or hedef_veri:
                            for log_kolonu in LAB_LOG_KOLON_ESLEME:
                                tree_ac.set(row_id, log_kolonu, "")
                        if hedef_veri:
                            for log_kolonu, lab_anahtari in LAB_LOG_KOLON_ESLEME.items():
                                deger = hedef_veri[lab_anahtari]
                                tree_ac.set(row_id, log_kolonu, deger if deger != "-" else "")
                    self.stripe_tree(tree_ac)
            except Exception as e:
                self.hata_kaydet("Laboratuvar AÇ senkronizasyon hatası", e)
                
        self.senkronize_ac_tablo = _senkronize_ac_tablo
        
        self.tree_lab_ac.bind(
            "<Control-v>", lambda event: self.lab_panodan_yapistir("ac", event)
        )
        self.tree_lab_ac.bind(
            "<Control-V>", lambda event: self.lab_panodan_yapistir("ac", event)
        )
        
        def satir_ekle_ac():
            item = self.tree_lab_ac.insert("", "end", values=(["-"] * len(kolonlar_ac)))
            self.tree_lab_ac.selection_set(item); self.tree_lab_ac.focus(item); self.tree_lab_ac.see(item); self.stripe_tree(self.tree_lab_ac)
            self.lab_sayaclari_guncelle()
        def satir_sil_ac():
            [self.tree_lab_ac.delete(i) for i in self.tree_lab_ac.selection()]; self.stripe_tree(self.tree_lab_ac)
            self.senkronize_ac_tablo(temizle_eslesmeyen=True)
            self.lab_sayaclari_guncelle()
        def temizle_ac():
            if messagebox.askyesno("Uyarı", "Emin misiniz?"):
                [self.tree_lab_ac.delete(i) for i in self.tree_lab_ac.get_children()]
                self.senkronize_ac_tablo(temizle_eslesmeyen=True)
                self.lab_sayaclari_guncelle()

        btn_ac = ttk.Frame(frame_ac)
        btn_ac.pack(fill='x', pady=(6, 0))
        ttk.Button(btn_ac, text="Satır Ekle", command=satir_ekle_ac, bootstyle="success outline").pack(side="left", padx=(0, 6))
        ttk.Button(btn_ac, text="Seçili Sil", command=satir_sil_ac, bootstyle="danger outline").pack(side="left")
        ttk.Button(btn_ac, text="Yukarı", command=lambda: self.tree_secili_satirlari_tasi(self.tree_lab_ac, -1), bootstyle="secondary outline").pack(side="left", padx=(6, 0))
        ttk.Button(btn_ac, text="Aşağı", command=lambda: self.tree_secili_satirlari_tasi(self.tree_lab_ac, 1), bootstyle="secondary outline").pack(side="left", padx=(6, 0))
        ttk.Button(btn_ac, text="Temizle", command=temizle_ac, bootstyle="warning outline").pack(side="right")
        self.lbl_lab_ac_sayac = ttk.Label(btn_ac, text="0 satır", style="Muted.TLabel")
        self.lbl_lab_ac_sayac.pack(side="right", padx=(0, 12))

        # --- ALT YARI: YN TABLOSU ---
        frame_yn = ttk.LabelFrame(paned, text="Yüzey Numunesi (YN) Laboratuvar Verileri", padding=8, bootstyle="success")
        paned.add(frame_yn, weight=1)
        
        kolonlar_yn = LAB_YN_KOLONLARI
        
        scroll_x_yn = ttk.Scrollbar(frame_yn, orient="horizontal")
        scroll_x_yn.pack(side="bottom", fill="x")
        scroll_y_yn = ttk.Scrollbar(frame_yn, orient="vertical")
        scroll_y_yn.pack(side="right", fill="y")
        
        self.tree_lab_yn = ttk.Treeview(frame_yn, columns=kolonlar_yn, show="headings", height=8, style="Lab.Treeview",
                                     xscrollcommand=scroll_x_yn.set, yscrollcommand=scroll_y_yn.set)
        for col in kolonlar_yn:
            self.tree_lab_yn.heading(col, text=col)
            self.tree_lab_yn.column(col, width=130, anchor="center") 
            
        scroll_x_yn.config(command=self.tree_lab_yn.xview)
        scroll_y_yn.config(command=self.tree_lab_yn.yview)
        self.tree_lab_yn.pack(expand=True, fill='both')
        self.tree_lab_yn.tag_configure('oddrow', background="white")
        self.tree_lab_yn.tag_configure('evenrow', background="#eaf7ff")
        self.tree_lab_yn.bind("<Double-1>", lambda e, t=self.tree_lab_yn: self.hucre_duzenle(e, t))
        
        self.tree_lab_yn.bind(
            "<Control-v>", lambda event: self.lab_panodan_yapistir("yn", event)
        )
        self.tree_lab_yn.bind(
            "<Control-V>", lambda event: self.lab_panodan_yapistir("yn", event)
        )
        
        def satir_ekle_yn():
            item = self.tree_lab_yn.insert("", "end", values=(["-"] * len(kolonlar_yn)))
            self.tree_lab_yn.selection_set(item); self.tree_lab_yn.focus(item); self.tree_lab_yn.see(item); self.stripe_tree(self.tree_lab_yn)
            self.lab_sayaclari_guncelle()
        def satir_sil_yn():
            [self.tree_lab_yn.delete(i) for i in self.tree_lab_yn.selection()]; self.stripe_tree(self.tree_lab_yn)
            self.lab_sayaclari_guncelle()
        def temizle_yn():
            if messagebox.askyesno("Uyarı", "Emin misiniz?"):
                [self.tree_lab_yn.delete(i) for i in self.tree_lab_yn.get_children()]
                self.lab_sayaclari_guncelle()
        
        btn_yn = ttk.Frame(frame_yn)
        btn_yn.pack(fill='x', pady=(6, 0))
        ttk.Button(btn_yn, text="Satır Ekle", command=satir_ekle_yn, bootstyle="success outline").pack(side="left", padx=(0, 6))
        ttk.Button(btn_yn, text="Seçili Sil", command=satir_sil_yn, bootstyle="danger outline").pack(side="left")
        ttk.Button(btn_yn, text="Yukarı", command=lambda: self.tree_secili_satirlari_tasi(self.tree_lab_yn, -1), bootstyle="secondary outline").pack(side="left", padx=(6, 0))
        ttk.Button(btn_yn, text="Aşağı", command=lambda: self.tree_secili_satirlari_tasi(self.tree_lab_yn, 1), bootstyle="secondary outline").pack(side="left", padx=(6, 0))
        ttk.Button(btn_yn, text="Temizle", command=temizle_yn, bootstyle="warning outline").pack(side="right")
        self.lbl_lab_yn_sayac = ttk.Label(btn_yn, text="0 satır", style="Muted.TLabel")
        self.lbl_lab_yn_sayac.pack(side="right", padx=(0, 12))
        self.lab_sayaclari_guncelle()

    def lab_sayaclari_guncelle(self):
        if hasattr(self, "lbl_lab_ac_sayac") and hasattr(self, "tree_lab_ac"):
            self.lbl_lab_ac_sayac.config(text=f"{len(self.tree_lab_ac.get_children())} satır")
        if hasattr(self, "lbl_lab_yn_sayac") and hasattr(self, "tree_lab_yn"):
            self.lbl_lab_yn_sayac.config(text=f"{len(self.tree_lab_yn.get_children())} satır")

    def _lab_projedeki_numune_etiketleri(self, hedef):
        """AÇ/YN sekmelerindeki kullanıcıya görünen numune adlarını döndür."""
        onek = "AC" if hedef == "ac" else "YN"
        etiketler = []
        try:
            kayitlar = self.ac_yn_sekme_kayitlari()
        except Exception:
            kayitlar = []
        for kayit in kayitlar:
            etiket = str(kayit.get("isim", "") or "").strip()
            if laboratuvar_numune_anahtari(etiket).startswith(onek):
                etiketler.append(etiket)
        return etiketler

    def _lab_satir_etiketlerini_projeye_uyarla(self, hedef, satirlar):
        return laboratuvar_numune_etiketlerini_uyarla(
            satirlar,
            self._lab_projedeki_numune_etiketleri(hedef),
        )

    def _lab_standart_satirlari_yapistir(self, hedef, satirlar):
        satirlar = self._lab_satir_etiketlerini_projeye_uyarla(hedef, satirlar)
        tree = self.tree_lab_ac if hedef == "ac" else self.tree_lab_yn
        children = list(tree.get_children())
        selection = tree.selection()
        start_index = children.index(selection[0]) if selection else len(children)
        last_item = None
        for offset, row in enumerate(satirlar):
            target_index = start_index + offset
            if target_index < len(children):
                last_item = children[target_index]
                tree.item(last_item, values=row)
            else:
                last_item = tree.insert("", "end", values=row)
        if last_item:
            tree.selection_set(last_item)
            tree.focus(last_item)
            tree.see(last_item)
        self.stripe_tree(tree)
        if hedef == "ac":
            self.senkronize_ac_tablo()
        self.lab_sayaclari_guncelle()

    def _lab1_satirlarini_yapistir(self, sonuc):
        ac_satirlari = self._lab_satir_etiketlerini_projeye_uyarla(
            "ac", sonuc.get("ac_satirlari", [])
        )
        yn_satirlari = self._lab_satir_etiketlerini_projeye_uyarla(
            "yn", sonuc.get("yn_satirlari", [])
        )
        ac_sonuc = laboratuvar_satirlarini_birlestir(
            self.lab_ac_satirlari_al(), ac_satirlari
        )
        yn_sonuc = laboratuvar_satirlarini_birlestir(
            self.lab_yn_satirlari_al(), yn_satirlari
        )

        if sonuc.get("ac_satirlari"):
            self.lab_ac_satirlari_yerlestir(ac_sonuc["satirlar"])
        if sonuc.get("yn_satirlari"):
            self.lab_yn_satirlari_yerlestir(yn_sonuc["satirlar"])
        if sonuc.get("ac_satirlari"):
            self.senkronize_ac_tablo()
        self.lab_sayaclari_guncelle()

        tekrar = ac_sonuc["pano_tekrari"] + yn_sonuc["pano_tekrari"]
        lines = [
            f"AÇ: {ac_sonuc['eklenen']} eklendi, {ac_sonuc['guncellenen']} güncellendi.",
            f"YN: {yn_sonuc['eklenen']} eklendi, {yn_sonuc['guncellenen']} güncellendi.",
        ]
        if sonuc.get("atlanan"):
            lines.append(f"{sonuc['atlanan']} başlık/uygunsuz satır atlandı.")
        if tekrar:
            lines.append(f"Panoda {tekrar} yinelenen No + Derinlik kaydı bulundu; son değer kullanıldı.")
        messagebox.showinfo("Laboratuvar Yapıştırma", "\n".join(lines))
        if hasattr(self, "durum_mesaji_yaz"):
            self.durum_mesaji_yaz("LAB_1 verileri panodan aktarıldı")

    def lab_panodan_yapistir(self, hedef, _event=None):
        try:
            clipboard_text = self.root.clipboard_get()
        except Exception:
            messagebox.showwarning("Laboratuvar Yapıştırma", "Panoda yapıştırılabilir Excel verisi bulunamadı.")
            return "break"

        try:
            sonuc = laboratuvar_pano_verisini_donustur(clipboard_text, hedef)
        except Exception as exc:
            self.hata_kaydet("Laboratuvar pano verisi okunamadı", exc)
            messagebox.showerror("Laboratuvar Yapıştırma", f"Pano verisi okunamadı:\n{exc}")
            return "break"

        if sonuc["format"] == "bos":
            messagebox.showwarning("Laboratuvar Yapıştırma", "Panoda veri satırı bulunamadı.")
            return "break"
        if sonuc["format"] == "lab1":
            if not sonuc.get("ac_satirlari") and not sonuc.get("yn_satirlari"):
                messagebox.showwarning(
                    "Laboratuvar Yapıştırma",
                    "LAB_1 düzeni tanındı ancak AÇ veya YN veri satırı bulunamadı.",
                )
                return "break"
            self._lab1_satirlarini_yapistir(sonuc)
        else:
            self._lab_standart_satirlari_yapistir(hedef, sonuc["standart_satirlar"])
        return "break"

    def lab_excel_yukle(self):
        """Laboratuvar Excel/CSV dosyasını standart tablo satırlarına çevirir."""
        dosya_yolu = filedialog.askopenfilename(
            initialdir=self.sablon_alt_klasoru("excel"),
            filetypes=[("Excel ve CSV Dosyaları", "*.xlsx *.xls *.csv")]
        )
        if not dosya_yolu:
            return

        try:
            sonuc = laboratuvar_dosyasi_oku(dosya_yolu)

            ac_dolu = len(self.tree_lab_ac.get_children()) > 0
            yn_dolu = hasattr(self, "tree_lab_yn") and len(self.tree_lab_yn.get_children()) > 0
            temizleyerek_yukle = False
            if ac_dolu or yn_dolu:
                cevap = messagebox.askyesnocancel(
                    "Tabloyu Temizle",
                    "Mevcut laboratuvar verileri silinip, sadece Excel'dekiler eklensin mi?\n"
                    "(Evet: Sil ve Ekle | Hayır: Silmeden Altına Ekle | İptal: Vazgeç)"
                )
                if cevap is None:
                    return
                if cevap is True:
                    temizleyerek_yukle = True
                    for child in self.tree_lab_ac.get_children():
                        self.tree_lab_ac.delete(child)
                    if hasattr(self, "tree_lab_yn"):
                        for child in self.tree_lab_yn.get_children():
                            self.tree_lab_yn.delete(child)

            # Excel'deki AÇ-1 ile tabloda bulunan AÇ1 aynı No+Derinlik
            # kaydıdır. Dosya yükleme yolu da pano yapıştırmasıyla aynı
            # birleştirme yardımcısını kullanmalıdır.
            ac_satirlari = self._lab_satir_etiketlerini_projeye_uyarla(
                "ac", sonuc["ac_satirlari"]
            )
            yn_satirlari = self._lab_satir_etiketlerini_projeye_uyarla(
                "yn", sonuc["yn_satirlari"]
            )
            ac_birlestirme = laboratuvar_satirlarini_birlestir(
                self.lab_ac_satirlari_al(), ac_satirlari
            )
            yn_birlestirme = laboratuvar_satirlarini_birlestir(
                self.lab_yn_satirlari_al(), yn_satirlari
            )
            if sonuc["ac_satirlari"]:
                self.lab_ac_satirlari_yerlestir(ac_birlestirme["satirlar"])
            if hasattr(self, "tree_lab_yn") and sonuc["yn_satirlari"]:
                self.lab_yn_satirlari_yerlestir(yn_birlestirme["satirlar"])

            self.lbl_lab_excel.config(text=os.path.basename(dosya_yolu))
            self.stripe_tree(self.tree_lab_ac)
            if hasattr(self, "tree_lab_yn"):
                self.stripe_tree(self.tree_lab_yn)
            self.lab_sayaclari_guncelle()

            self.senkronize_ac_tablo(temizle_eslesmeyen=temizleyerek_yukle)
            if hasattr(self, "durum_mesaji_yaz"):
                self.durum_mesaji_yaz("Laboratuvar dosyası yüklendi")

            messagebox.showinfo(
                "Mükemmel Otomasyon",
                "Excel başarıyla okundu!\n\n"
                "Değerler belirlenen standartlara yuvarlandı ve kısa kodlar (PL vb.) hatasız olarak cımbızlandı.\n"
                f"{sonuc['eklenen_ac']} adet AÇ ve {sonuc.get('eklenen_yn', 0)} adet YN verisi içe aktarıldı."
            )
        except EksikLaboratuvarSutunu as e:
            messagebox.showwarning("Eksik Sütun", str(e))
        except Exception as e:
            self.hata_kaydet("Laboratuvar Excel dosyası okunamadı", e)
            messagebox.showerror("Hata", f"Dosya okunurken bir sorun oluştu:\n{e}")

    def lab_tree_satirlari_al(self, tree_attr):
        tree = getattr(self, tree_attr, None)
        if not tree:
            return []
        return [tree.item(item)["values"] for item in tree.get_children()]

    def lab_tree_satirlari_yerlestir(self, tree_attr, satirlar):
        tree = getattr(self, tree_attr, None)
        if not tree:
            return
        for child in tree.get_children():
            tree.delete(child)
        for satir in satirlar or []:
            tree.insert("", "end", values=satir)
        self.stripe_tree(tree)
        self.lab_sayaclari_guncelle()

    def lab_ac_satirlari_al(self):
        return self.lab_tree_satirlari_al("tree_lab_ac")

    def lab_yn_satirlari_al(self):
        return self.lab_tree_satirlari_al("tree_lab_yn")

    def lab_ac_satirlari_yerlestir(self, satirlar):
        return self.lab_tree_satirlari_yerlestir("tree_lab_ac", satirlar)

    def lab_yn_satirlari_yerlestir(self, satirlar):
        return self.lab_tree_satirlari_yerlestir("tree_lab_yn", satirlar)

    def lab_numaralari_al(self, tree_attr):
        return [
            str(satir[0])
            for satir in self.lab_tree_satirlari_al(tree_attr)
            if satir and len(satir) > 0
        ]

    def lab_ac_numaralari_al(self):
        return self.lab_numaralari_al("tree_lab_ac")

    def lab_yn_numaralari_al(self):
        return self.lab_numaralari_al("tree_lab_yn")

    def lab_bos_satir_ekle(self, tree_attr, no):
        tree = getattr(self, tree_attr, None)
        if not tree:
            return None
        item = tree.insert("", "end", values=([no] + ["-"] * (len(tree["columns"]) - 1)))
        tree.selection_set(item)
        tree.focus(item)
        tree.see(item)
        self.stripe_tree(tree)
        self.lab_sayaclari_guncelle()
        return item

    def lab_ac_bos_satir_ekle(self, no):
        return self.lab_bos_satir_ekle("tree_lab_ac", no)

    def lab_yn_bos_satir_ekle(self, no):
        return self.lab_bos_satir_ekle("tree_lab_yn", no)
