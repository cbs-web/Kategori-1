import datetime
import tkinter as tk
from tkinter import ttk, messagebox


class AcYnIslemleri:
    def __init__(self, app):
        object.__setattr__(self, "app", app)

    def __getattr__(self, name):
        return getattr(self.app, name)

    def __setattr__(self, name, value):
        if name == "app":
            object.__setattr__(self, name, value)
        else:
            setattr(self.app, name, value)

    def sekme4_cukur(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="4. Araştırma Çukuru / Yüzey Numunesi")
        page = ttk.Frame(frame, padding=16)
        page.pack(fill="both", expand=True)

        ust_frame = ttk.Frame(page)
        ust_frame.pack(fill="x", pady=(0, 10))
        ttk.Label(ust_frame, text="Araştırma Çukuru / Yüzey Numunesi", style="Baslik.TLabel").pack(side="left")
        ttk.Button(ust_frame, text="Log Görüntülerini (JPG) Oluştur", command=self.tum_loglari_ciz, bootstyle="primary").pack(side="right")
        ttk.Separator(page).pack(fill="x", pady=(0, 12))

        self.tablo_alani = ttk.Notebook(page)
        self.tablo_alani.pack(expand=True, fill="both")

    def cukur_sekmesi_ekle(self, isim, enlem="", boylam="", tarih=""):
        sekme = ttk.Frame(self.tablo_alani)
        self.tablo_alani.add(sekme, text=isim)
        bilgi_frame = ttk.LabelFrame(sekme, text="Numune Bilgileri", padding=(12, 10), bootstyle="secondary")
        bilgi_frame.pack(fill="x", padx=8, pady=(8, 8))
        for col in (1, 3, 5, 7):
            bilgi_frame.grid_columnconfigure(col, weight=1)

        ttk.Label(bilgi_frame, text="Derinlik (m)").grid(row=0, column=0, sticky="w", padx=(0, 6), pady=3)
        derinlik_entry = ttk.Entry(bilgi_frame, width=12)
        derinlik_entry.grid(row=0, column=1, sticky="ew", padx=(0, 12), pady=3)
        derinlik_entry.insert(0, "3.0")

        ttk.Label(bilgi_frame, text="Enlem").grid(row=0, column=2, sticky="w", padx=(0, 6), pady=3)
        enlem_entry = ttk.Entry(bilgi_frame, width=15)
        enlem_entry.grid(row=0, column=3, sticky="ew", padx=(0, 12), pady=3)
        if enlem:
            enlem_entry.insert(0, enlem)

        ttk.Label(bilgi_frame, text="Boylam").grid(row=0, column=4, sticky="w", padx=(0, 6), pady=3)
        boylam_entry = ttk.Entry(bilgi_frame, width=15)
        boylam_entry.grid(row=0, column=5, sticky="ew", padx=(0, 12), pady=3)
        if boylam:
            boylam_entry.insert(0, boylam)

        ttk.Label(bilgi_frame, text="Tarih").grid(row=0, column=6, sticky="w", padx=(0, 6), pady=3)
        tarih_entry = ttk.Entry(bilgi_frame, width=12)
        tarih_entry.grid(row=0, column=7, sticky="ew", pady=3)
        if tarih:
            tarih_entry.insert(0, tarih)
        else:
            tarih_entry.insert(0, datetime.datetime.now().strftime("%d/%m/%Y"))

        kolonlar = ("Derinlik", "Örnek Tipi", "YASS", "Zemin Tanımı", "USCS", "Wn", "LL", "PL", "PI", "+4 %", "-200 %")
        tablo_frame = ttk.LabelFrame(sekme, text="Numune Tablosu", padding=(8, 8), bootstyle="info")
        tablo_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        btn_frame = ttk.Frame(tablo_frame)
        btn_frame.pack(fill="x", pady=(0, 8))
        satir_sayac = ttk.Label(btn_frame, text="0 satır", style="Muted.TLabel")
        satir_sayac.pack(side="right")

        tree_frame = ttk.Frame(tablo_frame)
        tree_frame.pack(fill="both", expand=True)
        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)
        scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal")
        scroll_y = ttk.Scrollbar(tree_frame, orient="vertical")
        tree = ttk.Treeview(
            tree_frame,
            columns=kolonlar,
            show="headings",
            height=10,
            selectmode="extended",
            style="AcYn.Treeview",
            xscrollcommand=scroll_x.set,
            yscrollcommand=scroll_y.set,
        )
        for col in kolonlar:
            tree.heading(col, text=col)
            w = 320 if col == "Zemin Tanımı" else 82
            anchor = "w" if col == "Zemin Tanımı" else "center"
            tree.column(col, width=w, anchor=anchor)
        tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        scroll_y.config(command=tree.yview)
        scroll_x.config(command=tree.xview)
        tree.tag_configure("oddrow", background="white")
        tree.tag_configure("evenrow", background="#f1f8ff")
        tree.bind("<Double-1>", lambda e, t=tree: self.hucre_duzenle(e, t))

        start_rows = [("0.0 - 3.0", "DS1", "-", "", "", "", "", "", "", "", "")]
        for vals in start_rows:
            tree.insert("", "end", values=vals)
        self.stripe_tree(tree)
        satir_sayac.config(text=f"{len(tree.get_children())} satır")

        def satir_sayac_guncelle():
            satir_sayac.config(text=f"{len(tree.get_children())} satır")

        def satir_ekle():
            cocuklar = tree.get_children()
            yeni_derinlik = "3.0 - 4.0"
            yeni_ds = "DS4"
            if cocuklar:
                son_deger = tree.item(cocuklar[-1])["values"]
                if son_deger and len(son_deger) > 0:
                    derinlik_str = str(son_deger[0])
                    if "-" in derinlik_str:
                        parts = derinlik_str.split("-")
                        try:
                            bas = float(parts[1].strip())
                            bit = bas + 1.0
                            yeni_derinlik = f"{bas:.1f} - {bit:.1f}"
                        except (ValueError, IndexError):
                            pass
                    ornek_str = str(son_deger[1])
                    if ornek_str.startswith("DS"):
                        try:
                            ds_no = int(ornek_str.replace("DS", "")) + 1
                            yeni_ds = f"DS{ds_no}"
                        except ValueError:
                            pass
            tree.insert("", "end", values=(yeni_derinlik, yeni_ds, "-", "", "", "", "", "", "", "", ""))
            self.stripe_tree(tree)
            satir_sayac_guncelle()

        def satir_sil():
            secili = tree.selection()
            if secili and messagebox.askyesno(
                "Satırları Sil",
                f"Seçili {len(secili)} numune satırı silinsin mi? Bu işlem geri alınamaz.",
            ):
                for item in secili:
                    tree.delete(item)
                self.stripe_tree(tree)
                satir_sayac_guncelle()

        ttk.Button(btn_frame, text="Satır Ekle", command=satir_ekle, bootstyle="success outline").pack(side="left", padx=(0, 6))
        ttk.Button(btn_frame, text="Seçili Satırları Sil", command=satir_sil, bootstyle="danger outline").pack(side="left")
        ttk.Button(btn_frame, text="Yukarı", command=lambda t=tree: self.tree_secili_satirlari_tasi(t, -1), bootstyle="secondary outline").pack(side="left", padx=(6, 0))
        ttk.Button(btn_frame, text="Aşağı", command=lambda t=tree: self.tree_secili_satirlari_tasi(t, 1), bootstyle="secondary outline").pack(side="left", padx=(6, 0))

        aciklama_frame = ttk.LabelFrame(sekme, text="Açıklama", padding=(8, 8), bootstyle="secondary")
        aciklama_frame.pack(fill="x", padx=8, pady=(0, 8))
        aciklama_text = tk.Text(aciklama_frame, height=4, width=80, wrap="word", font=("Segoe UI", 10), relief="solid", borderwidth=1)
        aciklama_text.pack(fill="x")

        self.ac_yn_sekme_bilgileri[str(sekme)] = {
            "sekme": sekme,
            "derinlik_entry": derinlik_entry,
            "enlem_entry": enlem_entry,
            "boylam_entry": boylam_entry,
            "tarih_entry": tarih_entry,
            "tree": tree,
            "aciklama_text": aciklama_text,
        }
        return sekme

    def ac_yn_sekme_bilgisi(self, sekme):
        bilgi = getattr(self, "ac_yn_sekme_bilgileri", {}).get(str(sekme))
        if not bilgi:
            return None
        bilgi["sekme"] = sekme
        bilgi["isim"] = self.tablo_alani.tab(sekme, "text")
        return bilgi

    def ac_yn_sekme_kayitlari(self):
        if not hasattr(self, "tablo_alani"):
            return []
        kayitlar = []
        for tab_id in self.tablo_alani.tabs():
            try:
                sekme = self.tablo_alani.nametowidget(tab_id)
            except KeyError:
                continue
            bilgi = self.ac_yn_sekme_bilgisi(sekme)
            if bilgi:
                kayitlar.append(bilgi)
        return kayitlar

    def ac_yn_satirlari(self, kayit):
        tree = kayit["tree"]
        return [tree.item(item)["values"] for item in tree.get_children()]

    def ac_yn_kaydi_verisini_oku(self, kayit):
        return {
            "isim": kayit["isim"],
            "derinlik": kayit["derinlik_entry"].get(),
            "enlem": kayit["enlem_entry"].get(),
            "boylam": kayit["boylam_entry"].get(),
            "tarih": kayit["tarih_entry"].get(),
            "satirlar": self.ac_yn_satirlari(kayit),
            "aciklama": kayit["aciklama_text"].get("1.0", tk.END).strip(),
        }

    def ac_yn_sekmelerini_temizle(self):
        if hasattr(self, "tablo_alani"):
            for widget in self.tablo_alani.winfo_children():
                widget.destroy()
        self.ac_yn_sekme_bilgileri = {}
