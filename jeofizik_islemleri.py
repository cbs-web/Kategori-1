import math
import os
from tkinter import ttk, filedialog, messagebox


JEOFON_DIZILIM_VARSAYILAN = {
    "jeofon_sayisi": "12",
    "jeofon_araligi": "2",
    "duz_offset": "0",
    "ters_offset": "2",
    "jeofizik_profil_adedi": "1",
}


def jeofon_dizilim_bilgilerini_dogrula(bilgiler):
    ham = {**JEOFON_DIZILIM_VARSAYILAN, **(bilgiler or {})}
    sonuc = {}
    hatalar = []

    for kod, etiket, minimum, maksimum in (
        ("jeofon_sayisi", "Jeofon sayısı", 1, 48),
        ("jeofizik_profil_adedi", "Jeofizik profil adedi", 1, 99),
    ):
        try:
            sayi_float = float(str(ham[kod]).replace(",", "."))
            sayi = int(sayi_float)
            if not math.isfinite(sayi_float) or sayi_float != sayi or not minimum <= sayi <= maksimum:
                raise ValueError
            sonuc[kod] = str(sayi)
        except (TypeError, ValueError, OverflowError):
            sonuc[kod] = JEOFON_DIZILIM_VARSAYILAN[kod]
            hatalar.append(f"{etiket} {minimum}-{maksimum} arasında bir tam sayı olmalıdır.")

    for kod, etiket, sifir_olabilir in (
        ("jeofon_araligi", "Jeofon aralığı", False),
        ("duz_offset", "Düz vuruş başlangıcı", True),
        ("ters_offset", "Ters vuruş offseti", True),
    ):
        try:
            sayi = float(str(ham[kod]).replace(",", "."))
            if not math.isfinite(sayi) or sayi < 0 or (not sifir_olabilir and sayi == 0):
                raise ValueError
            sonuc[kod] = f"{sayi:g}"
        except (TypeError, ValueError, OverflowError):
            sonuc[kod] = JEOFON_DIZILIM_VARSAYILAN[kod]
            kosul = "sıfır veya pozitif" if sifir_olabilir else "sıfırdan büyük"
            hatalar.append(f"{etiket} {kosul} sonlu bir sayı olmalıdır.")
    return sonuc, hatalar


class JeofizikIslemleri:
    def __init__(self, app):
        object.__setattr__(self, "app", app)

    def __getattr__(self, name):
        return getattr(self.app, name)

    def __setattr__(self, name, value):
        if name == "app":
            object.__setattr__(self, name, value)
        else:
            setattr(self.app, name, value)

    def sekme5_jeofizik(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="5. Jeofizik")
        ttk.Button(
            frame,
            text="Jeofizik Parametre Excel'i Yükle",
            command=self.excel_yukle,
            bootstyle="info",
        ).pack(pady=10)
        self.lbl_excel = ttk.Label(frame, text="Yüklenen Excel: Yok")
        self.lbl_excel.pack()

        dizilim_frame = ttk.LabelFrame(frame, text="Jeofon Dizilim ve Çalışma Özeti")
        dizilim_frame.pack(fill="x", padx=16, pady=(12, 6))
        self.jeofon_dizilim_entryleri = {}
        alanlar = [
            ("jeofon_sayisi", "Jeofon Sayısı", 12),
            ("jeofon_araligi", "Jeofon Aralığı (m)", 2),
            ("duz_offset", "Düz Vuruş Başlangıç (m)", 0),
            ("ters_offset", "Ters Vuruş Offset (m)", 2),
            ("jeofizik_profil_adedi", "Jeofizik Profil Adedi", 1),
        ]
        for idx, (kod, etiket, varsayilan) in enumerate(alanlar):
            ttk.Label(dizilim_frame, text=etiket).grid(row=0, column=idx, padx=6, pady=(8, 2), sticky="w")
            entry = ttk.Entry(dizilim_frame, width=14)
            entry.insert(0, str(varsayilan))
            entry.grid(row=1, column=idx, padx=6, pady=(0, 8), sticky="ew")
            dizilim_frame.columnconfigure(idx, weight=1)
            self.jeofon_dizilim_entryleri[kod] = entry

        ttk.Label(frame, text="Haritadan Aktarılan Koordinatlar").pack(pady=(20, 5))

        kolonlar = ("Çalışma No", "Enlem", "Boylam")
        self.tree_sis = ttk.Treeview(frame, columns=kolonlar, show="headings", height=5)
        for col in kolonlar:
            self.tree_sis.heading(col, text=col)
        self.tree_sis.pack(pady=5)
        self.tree_sis.bind("<Double-1>", lambda e, t=self.tree_sis: self.hucre_duzenle(e, t))

    def excel_yukle(self):
        dosya_yolu = filedialog.askopenfilename(
            initialdir=self.sablon_alt_klasoru("excel"),
            filetypes=[("Excel Dosyaları", "*.xlsx *.xls")],
        )
        if dosya_yolu:
            uzanti = os.path.splitext(dosya_yolu)[1].lower()
            if uzanti not in (".xlsx", ".xls"):
                messagebox.showwarning(
                    "Desteklenmeyen Dosya",
                    "Jeofizik rapor okuyucusu yalnız .xlsx ve .xls dosyalarını destekler.",
                )
                return
            if not os.path.isfile(dosya_yolu):
                messagebox.showwarning("Dosya Bulunamadı", "Seçilen jeofizik dosyası bulunamadı.")
                return
            if os.path.getsize(dosya_yolu) > 50 * 1024 * 1024:
                messagebox.showwarning("Dosya Çok Büyük", "Jeofizik Excel dosyası 50 MB sınırını aşıyor.")
                return
            self.jeofizik_excel_yolu_ayarla(dosya_yolu)

    def jeofizik_excel_yolu_al(self):
        excel_yolu = getattr(self.app, "_jeofizik_excel_yolu", "")
        if excel_yolu:
            return excel_yolu
        if not hasattr(self, "lbl_excel"):
            return ""
        etiket = self.lbl_excel.cget("text")
        onek = "Yüklenen Excel: "
        return etiket[len(onek):].strip() if etiket.startswith(onek) and etiket != onek + "Yok" else ""

    def jeofizik_excel_yolu_ayarla(self, yol):
        if not hasattr(self, "lbl_excel"):
            return
        if yol and yol != "Yok":
            yol = str(yol)
            self._jeofizik_excel_yolu = yol
            durum = "" if os.path.isfile(yol) else " [DOSYA BULUNAMADI]"
            self.lbl_excel.config(text=f"Yüklenen Excel: {yol}{durum}")
        else:
            self._jeofizik_excel_yolu = ""
            self.lbl_excel.config(text="Yüklenen Excel: Yok")

    def jeofizik_koordinatlari_al(self):
        if not hasattr(self, "tree_sis"):
            return []
        return [self.tree_sis.item(item)["values"] for item in self.tree_sis.get_children()]

    def jeofizik_koordinatlarini_temizle(self):
        if not hasattr(self, "tree_sis"):
            return
        for child in self.tree_sis.get_children():
            self.tree_sis.delete(child)

    def jeofizik_koordinat_ekle(self, calisma_no, enlem, boylam):
        if not hasattr(self, "tree_sis"):
            return None
        return self.tree_sis.insert("", "end", values=(calisma_no, enlem, boylam))

    def jeofizik_koordinatlari_yerlestir(self, satirlar):
        if not hasattr(self, "tree_sis"):
            return
        self.jeofizik_koordinatlarini_temizle()
        for satir in satirlar or []:
            self.tree_sis.insert("", "end", values=satir)

    def jeofon_dizilim_bilgileri_al(self):
        if not hasattr(self, "jeofon_dizilim_entryleri"):
            return dict(JEOFON_DIZILIM_VARSAYILAN)
        ham_bilgiler = {}
        for kod, varsayilan in JEOFON_DIZILIM_VARSAYILAN.items():
            entry = self.jeofon_dizilim_entryleri.get(kod)
            ham_bilgiler[kod] = entry.get().strip() if entry else varsayilan
        bilgiler, hatalar = jeofon_dizilim_bilgilerini_dogrula(ham_bilgiler)
        if hatalar:
            messagebox.showwarning(
                "Jeofon Dizilim Girişleri Düzeltildi",
                "Geçersiz alanlar güvenli varsayılanlara döndürüldü:\n\n"
                + "\n".join(f"- {hata}" for hata in hatalar),
            )
            self.jeofon_dizilim_bilgileri_yerlestir(bilgiler)
        return bilgiler

    def jeofon_dizilim_bilgileri_yerlestir(self, bilgiler):
        if not hasattr(self, "jeofon_dizilim_entryleri"):
            return
        bilgiler, _ = jeofon_dizilim_bilgilerini_dogrula(bilgiler)
        for kod, varsayilan in JEOFON_DIZILIM_VARSAYILAN.items():
            entry = self.jeofon_dizilim_entryleri.get(kod)
            if not entry:
                continue
            entry.delete(0, "end")
            entry.insert(0, str(bilgiler.get(kod, varsayilan)))
