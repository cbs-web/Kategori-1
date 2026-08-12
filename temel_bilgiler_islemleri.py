import re
import tkinter as tk
from tkinter import ttk


VARSAYILAN_ALAN_DEGERLERI = {}

PROJE_ALANLARI = [
    ("Proje Sahibi", "PROJE_ADI"),
    ("İl", "IL"),
    ("İlçe", "ILCE"),
    ("Köy", "KOY"),
    ("Mevkii", "MEVKII"),
    ("Pafta", "PAFTA"),
    ("ADA", "ADA"),
    ("PARSEL", "PARSEL"),
]

ARAZI_ALANLARI = [
    ("ENLEM", "ENLEM"),
    ("BOYLAM", "BOYLAM"),
    ("KOT", "KOT"),
    ("EĞİM yönü", "EGIM_YONU"),
    ("Eğim Derecesi", "EGIM"),
    ("PGA", "PGA"),
    ("PGV", "PGV"),
    ("Ss", "SS"),
    ("S1", "S1"),
    ("Sds", "SDS"),
    ("Sd1", "SD1"),
    ("İmar Durumu", "IMAR_DURUMU"),
    ("Plan Adı", "PLAN_ADI"),
    ("Yerel Zemin Sınıfı", "YEREL_ZEMIN_SINIFI"),
]

BINA_ETIKETLERI = [
    "Bina Kullanım Amacı",
    "Bina Kullanım Sınıfı",
    "Bina Önem Katsayısı",
    "Yapı Malzemesi",
    "Bodrum Kat Adedi",
    "Toplam Kat Adedi",
    "Yapı Yüksekliği",
    "Bina Yükseklik Sınıfı",
    "Yapı Boyutları",
    "Temel Derinliği",
]


class TemelBilgilerIslemleri:
    def __init__(self, app):
        object.__setattr__(self, "app", app)

    def __getattr__(self, name):
        return getattr(self.app, name)

    def __setattr__(self, name, value):
        if name == "app":
            object.__setattr__(self, name, value)
        else:
            setattr(self.app, name, value)

    def sayfa_olustur(self, sekme_adi, baslik):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text=sekme_adi)
        page = ttk.Frame(frame, padding=16)
        page.pack(fill="both", expand=True)
        ttk.Label(page, text=baslik, style="Baslik.TLabel").pack(anchor="w", pady=(0, 10))
        ttk.Separator(page).pack(fill="x", pady=(0, 12))
        return page

    def form_grubu_olustur(self, parent, baslik, alanlar, hedef_sozluk):
        grup = ttk.LabelFrame(parent, text=baslik, padding=(12, 10), bootstyle="secondary")
        grup.pack(fill="x", pady=(0, 12))
        for col in (1, 3):
            grup.grid_columnconfigure(col, weight=1, minsize=230)

        for index, (metin, kod) in enumerate(alanlar):
            row = index // 2
            col = (index % 2) * 2
            ttk.Label(grup, text=f"{metin} [{kod}]", style="Muted.TLabel").grid(
                row=row, column=col, sticky="w", padx=(0, 8), pady=5
            )
            entry = ttk.Entry(grup)
            entry.grid(row=row, column=col + 1, sticky="ew", padx=(0, 18 if col == 0 else 0), pady=5)
            if kod in VARSAYILAN_ALAN_DEGERLERI:
                entry.insert(0, VARSAYILAN_ALAN_DEGERLERI[kod])
            if kod == "YEREL_ZEMIN_SINIFI":
                entry.bind("<FocusOut>", self.nihai_zemin_sinifi_degisti)
            hedef_sozluk[kod] = entry
        return grup

    def varsayilan_alan_degerlerini_yerlestir(self, yalniz_bos=True):
        for kod, deger in VARSAYILAN_ALAN_DEGERLERI.items():
            entry = self.veri_alanlari.get(kod)
            if not entry:
                continue
            if yalniz_bos and entry.get().strip():
                continue
            entry.delete(0, tk.END)
            entry.insert(0, deger)

    def bina_form_grubu_olustur(self, parent):
        grup = ttk.LabelFrame(parent, text="Bina Parametreleri", padding=(12, 10), bootstyle="secondary")
        grup.pack(fill="x", pady=(0, 12))
        for col in (1, 3):
            grup.grid_columnconfigure(col, weight=1, minsize=230)

        for index, etiket in enumerate(BINA_ETIKETLERI):
            row = index // 2
            col = (index % 2) * 2
            ttk.Label(grup, text=etiket, style="Muted.TLabel").grid(
                row=row, column=col, sticky="w", padx=(0, 8), pady=5
            )
            entry = ttk.Entry(grup)
            entry.grid(row=row, column=col + 1, sticky="ew", padx=(0, 18 if col == 0 else 0), pady=5)
            self.bina_alanlari[etiket] = entry
        return grup

    def metin_alani_olustur(self, parent, baslik, yukseklik):
        grup = ttk.LabelFrame(parent, text=baslik, padding=(10, 8), bootstyle="secondary")
        text = tk.Text(grup, height=yukseklik, wrap="word", font=("Segoe UI", 10), relief="solid", borderwidth=1)
        scroll = ttk.Scrollbar(grup, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=scroll.set)
        text.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")
        grup.grid_columnconfigure(0, weight=1)
        grup.grid_rowconfigure(0, weight=1)
        return grup, text

    def formasyon_degisti(self, event=None):
        secili = self.combo_formasyon.get()
        metin = self.formasyon_metinleri.get(secili, "")
        if hasattr(self, "txt_formasyon_rapor"):
            self.txt_formasyon_rapor.delete("1.0", tk.END)
            self.txt_formasyon_rapor.insert("1.0", metin)
        if hasattr(self, "txt_muhendislik_jeolojisi"):
            self.txt_muhendislik_jeolojisi.delete("1.0", tk.END)

    def formasyon_bilgilerini_hazirla(self):
        secim = self.combo_formasyon.get().strip() if hasattr(self, "combo_formasyon") else ""
        muhendislik_metni = self.txt_muhendislik_jeolojisi.get("1.0", tk.END).strip() if hasattr(self, "txt_muhendislik_jeolojisi") else ""
        if not secim or secim == "Seçiniz...":
            formasyon_adi = ""
            formasyon_kisa = ""
            # Eski proje kayıtlarında formasyon seçimi tutulmamış olabilir. Bu durumda
            # kütüphaneden gelmiş mühendislik jeolojisi cümlesindeki ad/kodu geri kazan.
            kod_eslesmesi = re.search(r"[“\"]?([A-Za-zÇĞİÖŞÜçğıöşü0-9]+)\s+simgesiyle", muhendislik_metni, re.IGNORECASE)
            ad_eslesmesi = re.search(r"simgesiyle\s+gösterilen\s+[“\"]([^”\"]+)", muhendislik_metni, re.IGNORECASE)
            if kod_eslesmesi:
                formasyon_kisa = kod_eslesmesi.group(1).strip()
            if ad_eslesmesi:
                formasyon_adi = ad_eslesmesi.group(1).strip()
        elif "(" in secim and ")" in secim:
            formasyon_adi = secim.rsplit("(", 1)[0].strip()
            formasyon_kisa = secim.rsplit("(", 1)[1].replace(")", "").strip()
        else:
            formasyon_adi = secim
            formasyon_kisa = ""

        formasyon_metni = self.txt_formasyon_rapor.get("1.0", tk.END).strip() if hasattr(self, "txt_formasyon_rapor") else ""
        return {
            "secim": secim,
            "adi": formasyon_adi,
            "kisa": formasyon_kisa,
            "birim_tanimi": formasyon_adi,
            "formasyon_metni": formasyon_metni,
            "muhendislik_metni": muhendislik_metni,
        }

    def sekme1_proje(self):
        page = self.sayfa_olustur("1. Proje Bilgileri", "Proje Bilgileri")
        self.form_grubu_olustur(page, "Taşınmaz ve Proje", PROJE_ALANLARI, self.veri_alanlari)
        self.on_deger_paneli_olustur(page)

    def sekme2_arazi(self):
        page = self.sayfa_olustur("2. Arazi Bilgileri", "Arazi Bilgileri")
        self.form_grubu_olustur(page, "Konum, Deprem ve İmar", ARAZI_ALANLARI, self.veri_alanlari)

        formasyon = ttk.LabelFrame(page, text="Jeolojik Formasyon", padding=(12, 10), bootstyle="secondary")
        formasyon.pack(fill="x", pady=(0, 12))
        formasyon.grid_columnconfigure(1, weight=1)
        ttk.Label(formasyon, text="Jeolojik Formasyon (MJH)", style="Muted.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, 10), pady=4
        )
        self.combo_formasyon = ttk.Combobox(formasyon, values=self.formasyonlar, state="readonly")
        self.combo_formasyon.grid(row=0, column=1, sticky="ew", pady=4)
        self.combo_formasyon.set("Seçiniz...")
        self.combo_formasyon.bind("<<ComboboxSelected>>", self.formasyon_degisti)

        ttk.Button(
            formasyon,
            text="Çanakkale Jeoloji Kütüphanesi",
            command=self.jeoloji_kutuphanesi_penceresi,
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))

        metinler = ttk.Frame(page)
        metinler.pack(fill="both", expand=True)
        metinler.grid_columnconfigure(0, weight=1)
        metinler.grid_columnconfigure(1, weight=1)
        metinler.grid_rowconfigure(0, weight=1)

        frm_birim, self.txt_formasyon_rapor = self.metin_alani_olustur(
            metinler, "Formasyon Metni ([FORMASYON])", 7
        )
        frm_birim.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        frm_muh, self.txt_muhendislik_jeolojisi = self.metin_alani_olustur(
            metinler, "Kısa Mühendislik Jeolojisi Metni ([MUHENDISLIK_JEOLOJISI_METNI])", 7
        )
        frm_muh.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

    def sekme3_bina(self):
        page = self.sayfa_olustur("3. Bina Bilgileri", "Bina Bilgileri")
        ttk.Label(page, text="[BINA] etiketi bu bilgilerden tablo olarak üretilir.", style="Muted.TLabel").pack(
            anchor="w", pady=(0, 8)
        )
        self.bina_form_grubu_olustur(page)

    def proje_deger(self, kod, varsayilan=""):
        entry = self.veri_alanlari.get(kod)
        deger = entry.get().strip() if entry else ""
        return deger or varsayilan

    def bina_deger(self, etiket, varsayilan=""):
        entry = self.bina_alanlari.get(etiket)
        deger = entry.get().strip() if entry else ""
        return deger or varsayilan
