import os
from pathlib import Path
import queue
import re
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from docx import Document

from jeoloji_kutuphanesi import (
    AyniJeolojiKaydiHatasi,
    JeolojiKutuphanesi,
    JeolojiKutuphanesiHatasi,
    jeoloji_anahtari,
)
from harita_islemleri import kml_poligonlarini_oku
from harita_renkleri import (
    CALISAN_PARSEL_SINIR_KALINLIGI,
    CALISAN_PARSEL_SINIR_RENGI,
    JEOLOJI_HARITA_RENK_ACIKLAMASI,
    JEOLOJI_ONAYLI_DOLGU_RENGI,
    JEOLOJI_ONAYLI_SINIR_RENGI,
    JEOLOJI_SECILI_SINIR_RENGI,
    JEOLOJI_TASLAK_DOLGU_RENGI,
    JEOLOJI_TASLAK_SINIR_RENGI,
)
from jeoloji_klasor_aktarimi import kml_adaylarini_sirala, proje_klasorunu_incele
from jeoloji_toplu_aktarim import (
    TopluTaramaIptalEdildi,
    ilce_klasorunu_tara,
    toplu_proje_formasyonunu_belirle,
    toplu_kayitlari_aktar,
)
from jeoloji_word_aktarimi import word_raporunu_oku
from word_jeoloji_birlestirme import (
    wordde_stratigrafik_kesit_var_mi,
    yapisal_jeoloji_basligi_mi,
)


class JeolojiKutuphanesiIslemleri:
    """KATEGORI_1 için Çanakkale Jeoloji Kütüphanesi arayüzü."""

    def __init__(self, app):
        object.__setattr__(self, "app", app)

    def __getattr__(self, name):
        return getattr(self.app, name)

    def __setattr__(self, name, value):
        if name == "app":
            object.__setattr__(self, name, value)
        else:
            setattr(self.app, name, value)

    def _veritabani_yolu(self):
        klasor = self.kullanici_veri_klasoru_bul()
        if not klasor:
            klasor = os.path.join(os.path.expanduser("~"), ".k1")
        return os.path.join(klasor, "jeoloji", "canakkale_jeoloji.db")

    def _kutuphane(self):
        yol = os.path.abspath(self._veritabani_yolu())
        mevcut = getattr(self, "_jeoloji_kutuphanesi_nesnesi", None)
        if mevcut is None or os.path.normcase(str(mevcut.db_path)) != os.path.normcase(yol):
            mevcut = JeolojiKutuphanesi(yol)
            self._jeoloji_kutuphanesi_nesnesi = mevcut
        return mevcut

    def jeoloji_kutuphanesi_penceresi(self):
        pencere = getattr(self, "_jeoloji_kutuphanesi_penceresi", None)
        try:
            if pencere is not None and pencere.winfo_exists():
                pencere.deiconify()
                pencere.lift()
                self._listeyi_yenile()
                return pencere
        except tk.TclError:
            pass

        pencere = tk.Toplevel(self.root)
        self._jeoloji_kutuphanesi_penceresi = pencere
        pencere.title("Çanakkale Jeoloji Kütüphanesi")
        pencere.geometry("1180x760")
        pencere.minsize(900, 620)
        pencere.protocol("WM_DELETE_WINDOW", pencere.destroy)

        self._jeoloji_kutuphanesi_filtreleri = {
            "arama": tk.StringVar(master=pencere),
            "ilce": tk.StringVar(master=pencere),
            "formasyon": tk.StringVar(master=pencere),
            "durum": tk.StringVar(master=pencere, value=""),
        }
        self._jeoloji_kutuphanesi_form_vars = {}
        self._jeoloji_kutuphanesi_metinler = {}
        self._jeoloji_kutuphanesi_secili_id = None
        self._jeoloji_bekleyen_geometriler = None
        self._jeoloji_bekleyen_kml_path = ""
        self._jeoloji_kml_var = tk.StringVar(master=pencere, value="")

        header = ttk.Frame(pencere, padding=(12, 10))
        header.pack(fill="x")
        ttk.Label(
            header,
            text="Çanakkale Jeoloji Kütüphanesi",
            font=("Segoe UI", 14, "bold"),
        ).pack(side="left")
        ttk.Label(
            header,
            text=f"Veritabanı: {Path(self._kutuphane().db_path).name}",
        ).pack(side="left", padx=(16, 0))
        ttk.Button(header, text="Kütüphaneyi Yedekle", command=self._yedekle).pack(
            side="right", padx=(6, 0)
        )
        ttk.Button(
            header,
            text="İlçe Klasöründen Toplu Aktar",
            command=self._ilce_klasorunden_toplu_aktar,
        ).pack(side="right", padx=(6, 0))
        ttk.Button(header, text="Kapat", command=pencere.destroy).pack(side="right")

        ana = ttk.PanedWindow(pencere, orient="horizontal")
        ana.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        sol = ttk.Frame(ana, padding=(0, 0, 8, 0))
        sag = ttk.Frame(ana, padding=(8, 0, 0, 0))
        ana.add(sol, weight=1)
        ana.add(sag, weight=2)

        self._filtreleri_olustur(sol)
        self._listeyi_olustur(sol)
        self._formu_olustur(sag)
        self._yeni_kayit(bildir=False)
        self._listeyi_yenile()
        return pencere

    def _filtreleri_olustur(self, parent):
        frame = ttk.LabelFrame(parent, text="Kayıtları Filtrele", padding=8)
        frame.pack(fill="x", pady=(0, 8))
        specs = (("Ara", "arama"), ("İlçe", "ilce"), ("Formasyon", "formasyon"))
        for row, (label, key) in enumerate(specs):
            ttk.Label(frame, text=label).grid(row=row, column=0, sticky="e", padx=(0, 6), pady=3)
            entry = ttk.Entry(frame, textvariable=self._jeoloji_kutuphanesi_filtreleri[key])
            entry.grid(row=row, column=1, sticky="ew", pady=3)
            entry.bind("<Return>", lambda _event: self._listeyi_yenile())
        ttk.Label(frame, text="Durum").grid(row=3, column=0, sticky="e", padx=(0, 6), pady=3)
        durum = ttk.Combobox(
            frame,
            textvariable=self._jeoloji_kutuphanesi_filtreleri["durum"],
            values=("", "taslak", "onayli"),
            state="readonly",
        )
        durum.grid(row=3, column=1, sticky="ew", pady=3)
        durum.bind("<<ComboboxSelected>>", lambda _event: self._listeyi_yenile())
        ttk.Button(frame, text="Filtrele", command=self._listeyi_yenile).grid(
            row=4, column=0, columnspan=2, sticky="ew", pady=(6, 0)
        )
        frame.columnconfigure(1, weight=1)

    def _listeyi_olustur(self, parent):
        ttk.Button(parent, text="Yeni Kayıt", command=self._yeni_kayit).pack(fill="x", pady=(0, 6))
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)
        kolonlar = ("id", "ilce", "yerlesim", "ada_parsel", "formasyon", "kml", "durum", "rev")
        self._jeoloji_kutuphanesi_tree = ttk.Treeview(frame, columns=kolonlar, show="headings")
        basliklar = {
            "id": "No",
            "ilce": "İlçe",
            "yerlesim": "Köy / Mahalle",
            "ada_parsel": "Ada / Parsel",
            "formasyon": "Formasyon",
            "kml": "KML",
            "durum": "Durum",
            "rev": "Rev.",
        }
        genislikler = {"id": 48, "ilce": 110, "yerlesim": 140, "ada_parsel": 100, "formasyon": 180, "kml": 42, "durum": 70, "rev": 45}
        for column in kolonlar:
            self._jeoloji_kutuphanesi_tree.heading(column, text=basliklar[column])
            self._jeoloji_kutuphanesi_tree.column(
                column,
                width=genislikler[column],
                anchor="center" if column in ("id", "kml", "durum", "rev") else "w",
                stretch=column in ("yerlesim", "formasyon"),
            )
        scroll = ttk.Scrollbar(frame, orient="vertical", command=self._jeoloji_kutuphanesi_tree.yview)
        self._jeoloji_kutuphanesi_tree.configure(yscrollcommand=scroll.set)
        self._jeoloji_kutuphanesi_tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self._jeoloji_kutuphanesi_tree.bind("<<TreeviewSelect>>", self._secim_degisti)

    def _formu_olustur(self, parent):
        ust = ttk.Frame(parent)
        ust.pack(fill="x", pady=(0, 8))
        self._jeoloji_kutuphanesi_form_baslik = tk.StringVar(master=parent, value="Yeni kayıt")
        ttk.Label(ust, textvariable=self._jeoloji_kutuphanesi_form_baslik, font=("Segoe UI", 11, "bold")).pack(side="left")
        ttk.Button(ust, text="Word'den Doldur", command=self._worddan_doldur).pack(side="right")
        ttk.Button(ust, text="Klasörden Rapor Ekle", command=self._klasorden_rapor_ekle).pack(side="right", padx=(0, 6))
        ttk.Button(ust, text="Projeden Doldur", command=self._formu_projeden_doldur).pack(side="right", padx=(0, 6))

        canvas = tk.Canvas(parent, highlightthickness=0)
        scroll = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        form = ttk.Frame(canvas, padding=(4, 2, 12, 12))
        form.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=form, anchor="nw")
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        form.columnconfigure(1, weight=1)

        row = 0
        fields = (
            ("İl", "il", "entry"),
            ("İlçe", "ilce", "entry"),
            ("Köy / Mahalle", "yerlesim", "entry"),
            ("Ada", "ada", "entry"),
            ("Parsel", "parsel", "entry"),
            ("Formasyon", "formasyon", "formation"),
            ("Onay Durumu", "onay_durumu", "status"),
        )
        for label, key, kind in fields:
            ttk.Label(form, text=label).grid(row=row, column=0, sticky="e", padx=(0, 8), pady=4)
            if kind == "formation":
                values = list(getattr(self, "formasyonlar", []))
                widget = ttk.Combobox(form, values=values, state="normal")
            elif kind == "status":
                widget = ttk.Combobox(form, values=("taslak", "onayli"), state="readonly")
            else:
                widget = ttk.Entry(form)
            widget.grid(row=row, column=1, sticky="ew", pady=4)
            self._jeoloji_kutuphanesi_form_vars[key] = widget
            row += 1

        for label, key, height in (
            ("Genel Jeoloji Metni", "genel_jeoloji_metni", 7),
            ("İnceleme Alanı Jeolojisi", "inceleme_alani_jeolojisi", 7),
            ("Notlar", "notlar", 4),
        ):
            ttk.Label(form, text=label).grid(row=row, column=0, columnspan=2, sticky="w", pady=(8, 3))
            text = tk.Text(form, height=height, wrap="word", undo=True)
            text.grid(row=row + 1, column=0, columnspan=2, sticky="nsew")
            self._jeoloji_kutuphanesi_metinler[key] = text
            row += 2

        for label, key in (("Jeoloji Haritası", "harita_path"), ("Kaynak Rapor", "kaynak_rapor_path")):
            ttk.Label(form, text=label).grid(row=row, column=0, sticky="e", padx=(0, 8), pady=4)
            entry = ttk.Entry(form)
            entry.grid(row=row, column=1, sticky="ew", pady=4)
            self._jeoloji_kutuphanesi_form_vars[key] = entry
            ttk.Button(form, text="Seç", command=lambda k=key: self._dosya_sec(k)).grid(
                row=row, column=2, padx=(6, 0), pady=4
            )
            row += 1

        ttk.Label(form, text="Parsel KML", anchor="e").grid(
            row=row, column=0, sticky="e", padx=(0, 8), pady=4
        )
        ttk.Entry(form, textvariable=self._jeoloji_kml_var, state="readonly").grid(
            row=row, column=1, sticky="ew", pady=4
        )
        ttk.Button(form, text="KML Bağla", command=self._kml_bagla).grid(
            row=row, column=2, padx=(6, 0), pady=4
        )
        row += 1

        for label, key in (
            ("Harita Açıklaması", "harita_aciklamasi"),
            ("Harita Kaynağı", "harita_kaynagi"),
            ("Harita Ölçeği", "harita_olcegi"),
        ):
            ttk.Label(form, text=label).grid(row=row, column=0, sticky="e", padx=(0, 8), pady=4)
            entry = ttk.Entry(form)
            entry.grid(row=row, column=1, columnspan=2, sticky="ew", pady=4)
            self._jeoloji_kutuphanesi_form_vars[key] = entry
            row += 1

        actions = ttk.Frame(form)
        actions.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(12, 0))
        for column in range(4):
            actions.columnconfigure(column, weight=1)
        ttk.Button(actions, text="Kaydet", command=self._kaydet).grid(row=0, column=0, sticky="ew", padx=(0, 4))
        ttk.Button(actions, text="Seçili Kaydı Uygula", command=self._secili_kaydi_uygula).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(actions, text="Proje Önerisini Uygula", command=self.jeoloji_onerisini_projeye_uygula).grid(row=0, column=2, sticky="ew", padx=4)
        ttk.Button(actions, text="Arşivle", command=self._arsivle).grid(row=0, column=3, sticky="ew", padx=(4, 0))
        self._jeoloji_kutuphanesi_durum = tk.StringVar(master=parent, value="")
        ttk.Label(parent, textvariable=self._jeoloji_kutuphanesi_durum, wraplength=650).pack(fill="x", pady=(8, 0))

    @staticmethod
    def _widget_yaz(widget, value):
        if isinstance(widget, ttk.Combobox):
            widget.set(str(value or ""))
            return
        widget.delete(0, tk.END)
        if value:
            widget.insert(0, str(value))

    @staticmethod
    def _text_yaz(widget, value):
        widget.delete("1.0", tk.END)
        if value:
            widget.insert("1.0", str(value))

    def _durum(self, text):
        durum_degiskeni = getattr(self, "_jeoloji_kutuphanesi_durum", None)
        if durum_degiskeni is not None:
            try:
                durum_degiskeni.set(text)
            except tk.TclError:
                pass
        try:
            self.durum_mesaji_yaz(text)
        except Exception:
            pass

    def _yeni_kayit(self, bildir=True):
        self._jeoloji_kutuphanesi_secili_id = None
        self._jeoloji_bekleyen_geometriler = None
        self._jeoloji_bekleyen_kml_path = ""
        self._jeoloji_kml_var.set("")
        for widget in self._jeoloji_kutuphanesi_form_vars.values():
            self._widget_yaz(widget, "")
        for widget in self._jeoloji_kutuphanesi_metinler.values():
            self._text_yaz(widget, "")
        self._widget_yaz(self._jeoloji_kutuphanesi_form_vars["il"], "Çanakkale")
        self._widget_yaz(self._jeoloji_kutuphanesi_form_vars["onay_durumu"], "taslak")
        self._formu_projeden_doldur(metinleri_doldur=False)
        self._jeoloji_kutuphanesi_form_baslik.set("Yeni kayıt")
        if bildir:
            self._durum("Yeni jeoloji kütüphanesi kaydı hazırlandı.")

    def _formu_projeden_doldur(self, metinleri_doldur=True):
        values = {
            "il": self.proje_deger("IL", "Çanakkale"),
            "ilce": self.proje_deger("ILCE", ""),
            "yerlesim": self.proje_deger("KOY", ""),
            "ada": self.proje_deger("ADA", ""),
            "parsel": self.proje_deger("PARSEL", ""),
            "formasyon": self.combo_formasyon.get().strip() if hasattr(self, "combo_formasyon") else "",
        }
        if values["formasyon"] == "Se\u00e7iniz...":
            values["formasyon"] = ""
        for key, value in values.items():
            self._widget_yaz(self._jeoloji_kutuphanesi_form_vars[key], value)
        if metinleri_doldur:
            if hasattr(self, "txt_formasyon_rapor"):
                self._text_yaz(
                    self._jeoloji_kutuphanesi_metinler["genel_jeoloji_metni"],
                    self.txt_formasyon_rapor.get("1.0", tk.END).strip(),
                )
            if hasattr(self, "txt_muhendislik_jeolojisi"):
                self._text_yaz(
                    self._jeoloji_kutuphanesi_metinler["inceleme_alani_jeolojisi"],
                    self.txt_muhendislik_jeolojisi.get("1.0", tk.END).strip(),
                )
            if getattr(self, "img_mjh", ""):
                self._widget_yaz(self._jeoloji_kutuphanesi_form_vars["harita_path"], self.img_mjh)

    def _worddan_doldur(self):
        path = filedialog.askopenfilename(
            parent=self._jeoloji_kutuphanesi_penceresi,
            title="Jeoloji bilgisi alinacak Word raporunu sec",
            filetypes=[("Word raporlari", "*.docx"), ("Tum dosyalar", "*.*")],
        )
        if not path:
            return False
        try:
            result = word_raporunu_oku(path)
        except Exception as exc:
            messagebox.showerror(
                "Word aktarimi",
                f"Word raporu okunamadi:\n{exc}",
                parent=self._jeoloji_kutuphanesi_penceresi,
            )
            return False
        if result.hata:
            messagebox.showerror(
                "Word aktarimi",
                result.hata,
                parent=self._jeoloji_kutuphanesi_penceresi,
            )
            return False

        self._formu_yukle(result.kutuphane_kaydi())
        self._jeoloji_kutuphanesi_secili_id = None
        self._jeoloji_kutuphanesi_form_baslik.set(
            f"Word'den yeni kayit - {Path(path).name}"
        )
        if result.uyarilar:
            self._durum(
                "Word bilgileri forma alindi; kontrol edilmesi gerekenler: "
                + "; ".join(result.uyarilar)
            )
        else:
            self._durum("Word raporundan jeoloji bilgileri forma alindi; kaydetmeden once kontrol edin.")
        return True

    def _aday_sec(self, baslik, aciklama, adaylar, etiket_uret):
        if not adaylar:
            return None
        if len(adaylar) == 1:
            return adaylar[0]
        pencere = tk.Toplevel(self._jeoloji_kutuphanesi_penceresi)
        pencere.title(baslik)
        pencere.geometry("760x380")
        pencere.transient(self._jeoloji_kutuphanesi_penceresi)
        pencere.grab_set()
        ttk.Label(pencere, text=aciklama, wraplength=720, padding=(12, 12, 12, 6)).pack(fill="x")
        frame = ttk.Frame(pencere, padding=(12, 0, 12, 8))
        frame.pack(fill="both", expand=True)
        liste = tk.Listbox(frame, activestyle="dotbox")
        kaydirma = ttk.Scrollbar(frame, orient="vertical", command=liste.yview)
        liste.configure(yscrollcommand=kaydirma.set)
        liste.pack(side="left", fill="both", expand=True)
        kaydirma.pack(side="right", fill="y")
        for aday in adaylar:
            liste.insert(tk.END, etiket_uret(aday))
        liste.selection_set(0)
        secim = {"aday": None}

        def kabul():
            indeksler = liste.curselection()
            if indeksler:
                secim["aday"] = adaylar[int(indeksler[0])]
                pencere.destroy()

        alt = ttk.Frame(pencere, padding=(12, 0, 12, 12))
        alt.pack(fill="x")
        ttk.Button(alt, text="İptal", command=pencere.destroy).pack(side="right")
        ttk.Button(alt, text="Seç", command=kabul).pack(side="right", padx=(0, 6))
        liste.bind("<Double-Button-1>", lambda _event: kabul())
        pencere.protocol("WM_DELETE_WINDOW", pencere.destroy)
        pencere.wait_window()
        return secim["aday"]

    def _klasor_onizleme_onayi(self, klasor, word_sonucu, kml_adayi):
        pencere = tk.Toplevel(self._jeoloji_kutuphanesi_penceresi)
        pencere.title("Klasör Aktarımını Kontrol Et")
        pencere.geometry("820x620")
        pencere.transient(self._jeoloji_kutuphanesi_penceresi)
        pencere.grab_set()
        bilgi = (
            f"Klasör: {klasor}\n"
            f"Word: {Path(word_sonucu.dosya_yolu).name}\n"
            f"Konum: {word_sonucu.ilce} / {word_sonucu.yerlesim} · "
            f"{word_sonucu.ada} ada {word_sonucu.parsel} parsel\n"
            f"Formasyon: {word_sonucu.formasyon or '-'}\n"
            f"KML: {Path(kml_adayi['dosya_yolu']).name if kml_adayi else 'Bulunamadı'}\n"
            f"Parsel poligonu: {len(kml_adayi['poligonlar']) if kml_adayi else 0}"
        )
        ttk.Label(pencere, text=bilgi, justify="left", padding=12, wraplength=780).pack(fill="x")
        canvas = tk.Canvas(pencere, height=340, background="#f3f5f7", highlightthickness=1)
        canvas.pack(fill="both", expand=True, padx=12, pady=(0, 10))
        if kml_adayi:
            noktalar = [
                nokta
                for poligon in kml_adayi["poligonlar"]
                for nokta in poligon.get("noktalar", ())
            ]
            if noktalar:
                min_lat = min(point[0] for point in noktalar)
                max_lat = max(point[0] for point in noktalar)
                min_lon = min(point[1] for point in noktalar)
                max_lon = max(point[1] for point in noktalar)

                def ciz(_event=None):
                    canvas.delete("all")
                    width = max(canvas.winfo_width(), 100)
                    height = max(canvas.winfo_height(), 100)
                    lon_span = max(max_lon - min_lon, 1e-9)
                    lat_span = max(max_lat - min_lat, 1e-9)
                    scale = min((width - 40) / lon_span, (height - 40) / lat_span)
                    for poligon in kml_adayi["poligonlar"]:
                        coords = []
                        for lat, lon in poligon["noktalar"]:
                            coords.extend(
                                (
                                    20 + (lon - min_lon) * scale,
                                    height - 20 - (lat - min_lat) * scale,
                                )
                            )
                        if len(coords) >= 6:
                            canvas.create_polygon(
                                coords, outline="#1565c0", fill="#90caf9", stipple="gray25", width=3
                            )
                    canvas.create_text(12, 12, anchor="nw", text="Parsel KML önizlemesi", fill="#263238")

                canvas.bind("<Configure>", ciz)
        else:
            canvas.create_text(
                20, 20, anchor="nw", text="Bu kayıt KML olmadan eklenecek ve haritada gösterilmeyecek."
            )
        sonuc = {"onay": False}

        def onayla():
            sonuc["onay"] = True
            pencere.destroy()

        alt = ttk.Frame(pencere, padding=(12, 0, 12, 12))
        alt.pack(fill="x")
        ttk.Button(alt, text="İptal", command=pencere.destroy).pack(side="right")
        ttk.Button(alt, text="Forma Aktar", command=onayla).pack(side="right", padx=(0, 6))
        pencere.protocol("WM_DELETE_WINDOW", pencere.destroy)
        pencere.wait_window()
        return sonuc["onay"]

    def _klasorden_rapor_ekle(self, klasor=None):
        if not klasor:
            klasor = filedialog.askdirectory(
                parent=self._jeoloji_kutuphanesi_penceresi,
                title="Word raporu ve parsel KML'si bulunan proje klasörünü seç",
            )
        if not klasor:
            return False
        pencere = self._jeoloji_kutuphanesi_penceresi
        try:
            pencere.configure(cursor="watch")
            pencere.update_idletasks()
            inceleme = proje_klasorunu_incele(klasor)
        except Exception as exc:
            messagebox.showerror("Klasörden Rapor Ekle", f"Klasör incelenemedi:\n{exc}", parent=pencere)
            return False
        finally:
            try:
                pencere.configure(cursor="")
            except tk.TclError:
                pass
        word_adayi = self._aday_sec(
            "Ana Word Raporunu Seç",
            "Klasörde birden fazla Word adayı bulundu. Jeoloji kütüphanesine alınacak ana raporu seçin.",
            inceleme["word_adaylari"],
            lambda item: f"Puan {item['puan']:>3} · {Path(item['sonuc'].dosya_yolu).name}",
        )
        if word_adayi is None:
            messagebox.showwarning("Klasörden Rapor Ekle", "Klasörde okunabilir ana Word raporu bulunamadı.", parent=pencere)
            return False
        word_sonucu = word_adayi["sonuc"]
        kml_adaylari = kml_adaylarini_sirala(inceleme["kml_adaylari"], word_sonucu)
        kml_adayi = self._aday_sec(
            "Parsel KML'sini Seç",
            "Birden fazla poligon içeren KML bulundu. Ana rapora ait parsel KML'sini seçin.",
            kml_adaylari,
            lambda item: f"Puan {item['puan']:>3} · {Path(item['dosya_yolu']).name} · {len(item['poligonlar'])} poligon",
        )
        if not self._klasor_onizleme_onayi(klasor, word_sonucu, kml_adayi):
            return False
        self._formu_yukle(word_sonucu.kutuphane_kaydi())
        self._jeoloji_kutuphanesi_secili_id = None
        self._jeoloji_bekleyen_geometriler = kml_adayi["poligonlar"] if kml_adayi else None
        self._jeoloji_bekleyen_kml_path = kml_adayi["dosya_yolu"] if kml_adayi else ""
        self._jeoloji_kml_var.set(self._jeoloji_bekleyen_kml_path)
        self._jeoloji_kutuphanesi_form_baslik.set(
            f"Klasörden yeni kayıt · {Path(word_sonucu.dosya_yolu).name}"
        )
        uyari = ""
        if inceleme["kml_hatalari"]:
            uyari = f" · {len(inceleme['kml_hatalari'])} geçersiz KML atlandı"
        if kml_adayi:
            self._durum(
                f"Word ve {len(kml_adayi['poligonlar'])} parsel poligonu forma alındı; kaydetmeden önce kontrol edin{uyari}."
            )
        else:
            self._durum(f"Word forma alındı; parsel KML'si bulunamadı{uyari}.")
        return True

    def _arka_plan_gorevi(self, baslik, islem, tamamlandi):
        """Uzun dosya taramalarını Tk ana döngüsünü kilitlemeden çalıştırır."""
        pencere = tk.Toplevel(self._jeoloji_kutuphanesi_penceresi)
        pencere.title(baslik)
        pencere.geometry("560x170")
        pencere.resizable(False, False)
        pencere.transient(self._jeoloji_kutuphanesi_penceresi)
        pencere.grab_set()
        mesaj = tk.StringVar(master=pencere, value="Hazırlanıyor…")
        ttk.Label(pencere, textvariable=mesaj, wraplength=520, padding=(16, 18, 16, 8)).pack(fill="x")
        ilerleme = ttk.Progressbar(pencere, mode="determinate", maximum=1)
        ilerleme.pack(fill="x", padx=16, pady=(0, 14))
        iptal_event = threading.Event()
        kanal = queue.Queue()

        def rapor(mevcut, toplam, metin):
            kanal.put(("ilerleme", mevcut, toplam, metin))

        def calistir():
            try:
                kanal.put(("sonuc", islem(rapor, iptal_event)))
            except Exception as exc:
                kanal.put(("hata", exc))

        def iptal():
            iptal_event.set()
            mesaj.set("İptal isteği alındı; açık dosya işleminin bitmesi bekleniyor…")
            iptal_dugmesi.configure(state="disabled")

        iptal_dugmesi = ttk.Button(pencere, text="İptal", command=iptal)
        iptal_dugmesi.pack(side="right", padx=16, pady=(0, 14))
        pencere.protocol("WM_DELETE_WINDOW", iptal)

        def kontrol():
            try:
                while True:
                    item = kanal.get_nowait()
                    if item[0] == "ilerleme":
                        _, mevcut, toplam, metin = item
                        ilerleme.configure(maximum=max(int(toplam), 1))
                        ilerleme["value"] = min(int(mevcut), max(int(toplam), 1))
                        mesaj.set(metin)
                    elif item[0] == "sonuc":
                        pencere.grab_release()
                        pencere.destroy()
                        tamamlandi(item[1])
                        return
                    elif item[0] == "hata":
                        pencere.grab_release()
                        pencere.destroy()
                        hata = item[1]
                        if isinstance(hata, TopluTaramaIptalEdildi) or iptal_event.is_set():
                            self._durum("Toplu klasör işlemi iptal edildi.")
                        else:
                            messagebox.showerror(baslik, str(hata), parent=self._jeoloji_kutuphanesi_penceresi)
                        return
            except queue.Empty:
                pass
            try:
                pencere.after(100, kontrol)
            except tk.TclError:
                iptal_event.set()

        threading.Thread(target=calistir, name="jeoloji-toplu-aktarim", daemon=True).start()
        pencere.after(100, kontrol)

    def _ilce_klasorunden_toplu_aktar(self):
        root = filedialog.askdirectory(
            parent=self._jeoloji_kutuphanesi_penceresi,
            title="AYVACIK gibi ilçe klasörünü seç",
        )
        if not root:
            return False
        kutuphane = self._kutuphane()
        self._arka_plan_gorevi(
            "İlçe Klasörü Taranıyor",
            lambda rapor, iptal: ilce_klasorunu_tara(
                root, kutuphane=kutuphane, ilerleme=rapor, iptal_event=iptal
            ),
            self._toplu_onizleme_ac,
        )
        return True

    @staticmethod
    def _toplu_dosya_adi(aday, tur):
        if not aday:
            return "—"
        if tur == "word":
            return Path(aday["sonuc"].dosya_yolu).name
        return Path(aday["dosya_yolu"]).name

    def _toplu_onizleme_ac(self, tarama):
        projeler = tarama.get("projeler", [])
        if not projeler:
            messagebox.showinfo(
                "Toplu Jeoloji Aktarımı",
                "Seçilen ilçe klasöründe ana rapor veya parsel KML adayı bulunamadı.",
                parent=self._jeoloji_kutuphanesi_penceresi,
            )
            return
        eski = getattr(self, "_jeoloji_toplu_onizleme_penceresi", None)
        try:
            if eski is not None and eski.winfo_exists():
                eski.destroy()
        except tk.TclError:
            pass
        pencere = tk.Toplevel(self._jeoloji_kutuphanesi_penceresi)
        self._jeoloji_toplu_onizleme_penceresi = pencere
        self._jeoloji_toplu_projeler = projeler
        pencere.title("İlçe Klasöründen Toplu Jeoloji Aktarımı")
        pencere.geometry("1320x760")
        pencere.minsize(980, 620)
        pencere.transient(self._jeoloji_kutuphanesi_penceresi)

        ust = ttk.Frame(pencere, padding=12)
        ust.pack(fill="x")
        hazir = sum(1 for proje in projeler if proje["durum_kodu"] == "hazir")
        formasyon_gerekli = sum(
            1 for proje in projeler if proje["durum_kodu"] == "formasyon_gerekli"
        )
        self._jeoloji_toplu_ozet_var = tk.StringVar(
            master=pencere,
            value=(
                f"{len(projeler)} proje klasörü bulundu · {hazir} kesin eşleşme hazır"
                + (f" · {formasyon_gerekli} formasyon seçimi bekliyor" if formasyon_gerekli else "")
            ),
        )
        ttk.Label(
            ust,
            textvariable=self._jeoloji_toplu_ozet_var,
            font=("Segoe UI", 11, "bold"),
        ).pack(side="left")
        self._jeoloji_toplu_onayli_var = tk.BooleanVar(master=pencere, value=True)
        ttk.Checkbutton(
            ust, text="Aktarılanları onaylı kaydet", variable=self._jeoloji_toplu_onayli_var
        ).pack(side="right")

        ana = ttk.PanedWindow(pencere, orient="vertical")
        ana.pack(fill="both", expand=True, padx=12, pady=(0, 10))
        liste_frame = ttk.Frame(ana)
        detay_frame = ttk.Frame(ana)
        ana.add(liste_frame, weight=3)
        ana.add(detay_frame, weight=2)
        kolonlar = ("sec", "klasor", "word", "kml", "konum", "ada", "formasyon", "durum")
        tree = ttk.Treeview(liste_frame, columns=kolonlar, show="headings")
        self._jeoloji_toplu_tree = tree
        basliklar = {
            "sec": "Seç", "klasor": "Proje klasörü", "word": "Ana Word", "kml": "Parsel KML",
            "konum": "İlçe / Yerleşim", "ada": "Ada / Parsel", "formasyon": "Formasyon", "durum": "Durum",
        }
        widths = {"sec": 42, "klasor": 245, "word": 210, "kml": 210, "konum": 145, "ada": 90, "formasyon": 170, "durum": 145}
        for column in kolonlar:
            tree.heading(column, text=basliklar[column])
            tree.column(column, width=widths[column], anchor="center" if column in ("sec", "ada", "durum") else "w")
        yscroll = ttk.Scrollbar(liste_frame, orient="vertical", command=tree.yview)
        xscroll = ttk.Scrollbar(liste_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        liste_frame.rowconfigure(0, weight=1)
        liste_frame.columnconfigure(0, weight=1)
        tree.tag_configure("hazir", background="#e8f5e9")
        tree.tag_configure("uyari", background="#fff8e1")
        tree.tag_configure("mevcut", foreground="#607d8b")
        tree.tag_configure("hata", background="#ffebee")
        self._toplu_tree_yenile()

        self._jeoloji_toplu_detay = tk.Text(detay_frame, height=9, wrap="word", state="disabled")
        self._jeoloji_toplu_detay.pack(side="left", fill="both", expand=True)
        self._jeoloji_toplu_canvas = tk.Canvas(
            detay_frame, width=360, background="#f3f5f7", highlightthickness=1
        )
        self._jeoloji_toplu_canvas.pack(side="right", fill="both", padx=(8, 0))
        tree.bind("<<TreeviewSelect>>", self._toplu_secim_detayi)
        tree.bind("<Double-Button-1>", self._toplu_satir_secimini_degistir)
        if tree.get_children():
            tree.selection_set(tree.get_children()[0])
            self._toplu_secim_detayi()

        alt = ttk.Frame(pencere, padding=(12, 0, 12, 12))
        alt.pack(fill="x")
        ttk.Button(alt, text="Kapat", command=pencere.destroy).pack(side="right")
        ttk.Button(alt, text="Seçili Hazır Kayıtları Aktar", command=self._toplu_aktarimi_baslat).pack(side="right", padx=(0, 6))
        ttk.Button(alt, text="Formasyonu Belirle", command=self._toplu_formasyonu_belirle).pack(side="left")
        ttk.Button(alt, text="Sorunlu Kaydı Tek Proje Olarak İncele", command=self._toplu_sorunlu_kaydi_incele).pack(side="left")
        ttk.Button(alt, text="Tüm Hazırları Seç", command=lambda: self._toplu_hazirlari_sec(True)).pack(side="left", padx=(6, 0))
        ttk.Button(alt, text="Seçimi Kaldır", command=lambda: self._toplu_hazirlari_sec(False)).pack(side="left", padx=(6, 0))

    def _toplu_tree_yenile(self):
        tree = getattr(self, "_jeoloji_toplu_tree", None)
        if tree is None:
            return
        for item in tree.get_children():
            tree.delete(item)
        for index, proje in enumerate(self._jeoloji_toplu_projeler):
            record = proje.get("record") or {}
            durum_kodu = proje.get("durum_kodu", "")
            if durum_kodu == "hazir" and proje.get("eslesme_durumu") == "duzeltildi":
                tag = "uyari"
            else:
                tag = "hazir" if durum_kodu == "hazir" else "mevcut" if durum_kodu in ("mevcut", "aktarildi") else "uyari" if durum_kodu in ("revizyon", "word_belirsiz", "kml_belirsiz", "formasyon_gerekli") else "hata"
            tree.insert(
                "", "end", iid=str(index), tags=(tag,),
                values=(
                    "✓" if proje.get("secili") else "",
                    proje.get("goreli_klasor", ""),
                    self._toplu_dosya_adi(proje.get("word_adayi"), "word"),
                    self._toplu_dosya_adi(proje.get("kml_adayi"), "kml"),
                    " / ".join(part for part in (record.get("ilce", ""), record.get("yerlesim", "")) if part),
                    " / ".join(part for part in (record.get("ada", ""), record.get("parsel", "")) if part),
                    record.get("formasyon", ""),
                    proje.get("durum", ""),
                ),
            )
        ozet = getattr(self, "_jeoloji_toplu_ozet_var", None)
        if ozet is not None:
            projeler = self._jeoloji_toplu_projeler
            hazir = sum(1 for proje in projeler if proje.get("durum_kodu") == "hazir")
            formasyon_gerekli = sum(
                1 for proje in projeler if proje.get("durum_kodu") == "formasyon_gerekli"
            )
            metin = f"{len(projeler)} proje klasörü bulundu · {hazir} kesin eşleşme hazır"
            if formasyon_gerekli:
                metin += f" · {formasyon_gerekli} formasyon seçimi bekliyor"
            ozet.set(metin)

    def _toplu_secili_proje(self):
        tree = getattr(self, "_jeoloji_toplu_tree", None)
        secim = tree.selection() if tree is not None else ()
        if not secim:
            return None
        return self._jeoloji_toplu_projeler[int(secim[0])]

    def _toplu_secim_detayi(self, _event=None):
        proje = self._toplu_secili_proje()
        if not proje:
            return
        word = proje.get("word_adayi")
        kml = proje.get("kml_adayi")
        record = proje.get("record") or {}
        eslesme = proje.get("eslesme") or {}
        metin = (
            f"Klasör: {proje.get('klasor', '')}\n"
            f"Durum: {proje.get('durum', '')}\n"
            f"Ana Word: {word['sonuc'].dosya_yolu if word else '—'}\n"
            f"Word puanı: {word['puan'] if word else '—'}\n"
            f"Parsel KML: {kml['dosya_yolu'] if kml else '—'}\n"
            f"Konum: {record.get('ilce', '')} / {record.get('yerlesim', '')} · "
            f"{record.get('ada', '')} ada {record.get('parsel', '')} parsel\n"
            f"Formasyon: {record.get('formasyon') or '—'}"
        )
        formasyon_adaylari = (
            getattr(word["sonuc"], "formasyon_adaylari", ()) if word else ()
        )
        if formasyon_adaylari and not record.get("formasyon"):
            metin += "\nFormasyon adayları: " + ", ".join(formasyon_adaylari)
        if eslesme:
            kaynaklar = eslesme.get("kaynaklar", {})
            canonical = eslesme.get("kanonik", {})
            source_labels = (
                ("Word içeriği", "word_icerigi"),
                ("Proje klasörü", "proje_klasoru"),
                ("Word dosya adı", "word_dosya_adi"),
                ("KML", "kml"),
            )
            metin += "\n\nKÜNYE KAYNAKLARI"
            for label, key in source_labels:
                data = kaynaklar.get(key, {})
                metin += (
                    f"\n{label}: {data.get('yerlesim') or '—'} · "
                    f"{data.get('ada') or '—'} / {data.get('parsel') or data.get('tek_parsel') or '—'}"
                )
            metin += (
                f"\nKullanılacak: {canonical.get('yerlesim') or '—'} · "
                f"{canonical.get('ada') or '—'} / {canonical.get('parsel') or '—'}"
            )
            if eslesme.get("uyarilar"):
                metin += "\n\nDÜZELTMELER / UYARILAR\n- " + "\n- ".join(eslesme["uyarilar"])
            if eslesme.get("celiskiler"):
                metin += "\n\nÇELİŞKİLER\n- " + "\n- ".join(eslesme["celiskiler"])
        text = self._jeoloji_toplu_detay
        text.configure(state="normal")
        text.delete("1.0", tk.END)
        text.insert("1.0", metin)
        text.configure(state="disabled")
        self._toplu_canvas_ciz(kml["poligonlar"] if kml else [])

    def _toplu_canvas_ciz(self, poligonlar):
        canvas = self._jeoloji_toplu_canvas
        canvas.delete("all")
        noktalar = [point for polygon in poligonlar for point in polygon.get("noktalar", ())]
        if not noktalar:
            canvas.create_text(16, 16, anchor="nw", text="KML önizlemesi yok")
            return
        width = max(canvas.winfo_width(), 340)
        height = max(canvas.winfo_height(), 150)
        min_lat, max_lat = min(p[0] for p in noktalar), max(p[0] for p in noktalar)
        min_lon, max_lon = min(p[1] for p in noktalar), max(p[1] for p in noktalar)
        scale = min((width - 30) / max(max_lon - min_lon, 1e-9), (height - 30) / max(max_lat - min_lat, 1e-9))
        for polygon in poligonlar:
            coords = []
            for lat, lon in polygon["noktalar"]:
                coords.extend((15 + (lon - min_lon) * scale, height - 15 - (lat - min_lat) * scale))
            if len(coords) >= 6:
                canvas.create_polygon(coords, outline="#1565c0", fill="#90caf9", stipple="gray25", width=2)

    def _toplu_satir_secimini_degistir(self, _event=None):
        proje = self._toplu_secili_proje()
        if proje and proje.get("durum_kodu") == "hazir":
            proje["secili"] = not proje.get("secili")
            secim = self._jeoloji_toplu_tree.selection()
            self._toplu_tree_yenile()
            if secim:
                self._jeoloji_toplu_tree.selection_set(secim[0])

    def _toplu_hazirlari_sec(self, secili):
        for proje in self._jeoloji_toplu_projeler:
            if proje.get("durum_kodu") == "hazir":
                proje["secili"] = bool(secili)
        self._toplu_tree_yenile()

    def _toplu_formasyonu_belirle(self):
        proje = self._toplu_secili_proje()
        if not proje:
            return
        if proje.get("durum_kodu") != "formasyon_gerekli":
            messagebox.showinfo(
                "Formasyonu Belirle",
                "Bu işlem yalnız formasyon seçimi bekleyen satırlarda kullanılabilir.",
                parent=self._jeoloji_toplu_onizleme_penceresi,
            )
            return

        word_adayi = proje.get("word_adayi") or {}
        word = word_adayi.get("sonuc")
        adaylar = list(getattr(word, "formasyon_adaylari", ()) or ())
        katalog = [
            value for value in getattr(self, "formasyonlar", ())
            if jeoloji_anahtari(value) not in ("", "seciniz")
        ]
        degerler = list(dict.fromkeys((*adaylar, *katalog)))

        pencere = tk.Toplevel(self._jeoloji_toplu_onizleme_penceresi)
        pencere.title("Formasyonu Belirle")
        pencere.geometry("590x205")
        pencere.resizable(False, False)
        pencere.transient(self._jeoloji_toplu_onizleme_penceresi)
        pencere.grab_set()
        ttk.Label(
            pencere,
            text=(
                "Word raporunda formasyon kesinleştirilemedi. Listeden seçin; "
                "listede yoksa adını yazabilirsiniz."
            ),
            wraplength=550,
            padding=(18, 18, 18, 8),
        ).pack(fill="x")
        secim_var = tk.StringVar(master=pencere, value="")
        combo = ttk.Combobox(
            pencere,
            textvariable=secim_var,
            values=degerler,
            state="normal",
            width=65,
        )
        combo.pack(fill="x", padx=18, pady=(4, 12))
        combo.focus_set()
        if adaylar:
            ttk.Label(
                pencere,
                text="Rapordan bulunan adaylar: " + ", ".join(adaylar),
                wraplength=550,
                foreground="#8a5a00",
            ).pack(fill="x", padx=18, pady=(0, 8))

        dugmeler = ttk.Frame(pencere, padding=(18, 0, 18, 14))
        dugmeler.pack(fill="x", side="bottom")

        def uygula():
            try:
                toplu_proje_formasyonunu_belirle(
                    proje,
                    secim_var.get(),
                    kutuphane=self._kutuphane(),
                )
            except (ValueError, JeolojiKutuphanesiHatasi) as exc:
                messagebox.showerror("Formasyonu Belirle", str(exc), parent=pencere)
                return
            secim = self._jeoloji_toplu_tree.selection()
            pencere.grab_release()
            pencere.destroy()
            self._toplu_tree_yenile()
            if secim:
                self._jeoloji_toplu_tree.selection_set(secim[0])
            self._toplu_secim_detayi()

        ttk.Button(dugmeler, text="İptal", command=pencere.destroy).pack(side="right")
        ttk.Button(dugmeler, text="Uygula", command=uygula).pack(side="right", padx=(0, 6))
        pencere.bind("<Return>", lambda _event: uygula())

    def _toplu_sorunlu_kaydi_incele(self):
        proje = self._toplu_secili_proje()
        if not proje:
            return
        self._klasorden_rapor_ekle(proje["klasor"])

    def _toplu_aktarimi_baslat(self):
        secilen = [
            proje for proje in self._jeoloji_toplu_projeler
            if proje.get("secili") and proje.get("durum_kodu") == "hazir"
        ]
        if not secilen:
            messagebox.showinfo("Toplu Jeoloji Aktarımı", "Aktarılacak hazır kayıt seçilmedi.")
            return
        onayli = self._jeoloji_toplu_onayli_var.get()
        kutuphane = self._kutuphane()
        self._arka_plan_gorevi(
            "Jeoloji Kayıtları Aktarılıyor",
            lambda rapor, iptal: toplu_kayitlari_aktar(
                self._jeoloji_toplu_projeler,
                kutuphane,
                onayli=onayli,
                ilerleme=rapor,
                iptal_event=iptal,
            ),
            self._toplu_aktarim_tamamlandi,
        )

    def _toplu_aktarim_tamamlandi(self, sonuc):
        self._toplu_tree_yenile()
        self._listeyi_yenile()
        self.jeoloji_harita_katmanini_yenile(zorla=True)
        mesaj = (
            f"Başarılı: {len(sonuc['basarili'])}\n"
            f"Atlanan: {len(sonuc['atlanan'])}\n"
            f"Hatalı: {len(sonuc['hatali'])}"
        )
        if sonuc["hatali"]:
            mesaj += "\n\n" + "\n".join(
                f"{item['proje']['goreli_klasor']}: {item['hata']}" for item in sonuc["hatali"][:8]
            )
        messagebox.showinfo("Toplu Jeoloji Aktarımı", mesaj, parent=self._jeoloji_toplu_onizleme_penceresi)

    def _kml_bagla(self):
        path = filedialog.askopenfilename(
            parent=self._jeoloji_kutuphanesi_penceresi,
            title="Kayda bağlanacak parsel KML'sini seç",
            filetypes=[("KML dosyaları", "*.kml")],
        )
        if not path:
            return False
        try:
            poligonlar = kml_poligonlarini_oku(path)
        except (OSError, ValueError) as exc:
            messagebox.showerror("Parsel KML", str(exc), parent=self._jeoloji_kutuphanesi_penceresi)
            return False
        self._jeoloji_bekleyen_geometriler = poligonlar
        self._jeoloji_bekleyen_kml_path = path
        self._jeoloji_kml_var.set(path)
        self._durum(f"{len(poligonlar)} parsel poligonu hazırlandı; bağlamak için kaydı kaydedin.")
        return True

    def _formu_topla(self):
        result = {
            key: widget.get().strip()
            for key, widget in self._jeoloji_kutuphanesi_form_vars.items()
        }
        result.update({
            key: widget.get("1.0", tk.END).strip()
            for key, widget in self._jeoloji_kutuphanesi_metinler.items()
        })
        if self._jeoloji_kutuphanesi_secili_id:
            previous = self._kutuphane().getir(self._jeoloji_kutuphanesi_secili_id)
            if previous:
                for key in (
                    "bolum_docx_path",
                    "bolum_hash",
                    "kaynak_klasor_path",
                    "kaynak_rapor_hash",
                    "kunye_kaynaklari_json",
                    "kunye_duzeltme_notu",
                ):
                    result[key] = previous.get(key, "")
        return result

    def _kaydet(self):
        try:
            record = self._formu_topla()
            map_path = record.get("harita_path", "")
            if map_path and os.path.isfile(map_path):
                record["harita_path"] = self._kutuphane().harita_dosyasi_ekle(map_path)
            record_id = self._kutuphane().kaydet(record, self._jeoloji_kutuphanesi_secili_id)
            if self._jeoloji_bekleyen_geometriler is not None:
                self._kutuphane().geometrileri_degistir(
                    record_id,
                    self._jeoloji_bekleyen_geometriler,
                    self._jeoloji_bekleyen_kml_path,
                )
        except AyniJeolojiKaydiHatasi as exc:
            messagebox.showwarning(
                "Jeoloji Kütüphanesi",
                f"Aynı konum ve formasyon için #{exc.kayit_id} numaralı kayıt zaten var.",
                parent=self._jeoloji_kutuphanesi_penceresi,
            )
            return
        except JeolojiKutuphanesiHatasi as exc:
            messagebox.showerror("Jeoloji Kütüphanesi", str(exc), parent=self._jeoloji_kutuphanesi_penceresi)
            return
        self._jeoloji_kutuphanesi_secili_id = record_id
        self._jeoloji_bekleyen_geometriler = None
        self._jeoloji_bekleyen_kml_path = ""
        self._listeyi_yenile()
        record = self._kutuphane().getir(record_id)
        if record:
            self._formu_yukle(record)
        self._durum(f"Jeoloji kütüphanesi kaydı kaydedildi: #{record_id}")

    def _formu_yukle(self, record):
        self._jeoloji_kutuphanesi_secili_id = record.get("id")
        for key, widget in self._jeoloji_kutuphanesi_form_vars.items():
            self._widget_yaz(widget, record.get(key, ""))
        for key, widget in self._jeoloji_kutuphanesi_metinler.items():
            self._text_yaz(widget, record.get(key, ""))
        self._jeoloji_bekleyen_geometriler = None
        self._jeoloji_bekleyen_kml_path = ""
        geometriler = self._kutuphane().geometrileri_getir(record["id"]) if record.get("id") else []
        self._jeoloji_kml_var.set(geometriler[0].get("kml_path", "") if geometriler else "")
        self._jeoloji_kutuphanesi_form_baslik.set(
            f"Kayıt #{record.get('id')} · Revizyon {record.get('revizyon_no', 1)}"
        )

    def _listeyi_yenile(self):
        tree = getattr(self, "_jeoloji_kutuphanesi_tree", None)
        if tree is None:
            return
        filters = self._jeoloji_kutuphanesi_filtreleri
        records = self._kutuphane().listele(
            ilce=filters["ilce"].get(),
            formasyon=filters["formasyon"].get(),
            onay_durumu=filters["durum"].get(),
            arama=filters["arama"].get(),
        )
        for item in tree.get_children():
            tree.delete(item)
        for record in records:
            ada_parsel = " / ".join(
                part for part in (record.get("ada", ""), record.get("parsel", "")) if part
            )
            tree.insert(
                "",
                "end",
                iid=str(record["id"]),
                values=(
                    record["id"],
                    record.get("ilce", ""),
                    record.get("yerlesim", ""),
                    ada_parsel,
                    record.get("formasyon", ""),
                    "Var" if record.get("geometri_sayisi") else "—",
                    record.get("onay_durumu", "taslak"),
                    record.get("revizyon_no", 1),
                ),
            )

    def _secim_degisti(self, _event=None):
        tree = self._jeoloji_kutuphanesi_tree
        selection = tree.selection()
        if not selection:
            return
        record = self._kutuphane().getir(int(selection[0]))
        if record:
            self._formu_yukle(record)

    def _secili_kaydi_uygula(self):
        record_id = self._jeoloji_kutuphanesi_secili_id
        if not record_id:
            messagebox.showinfo("Jeoloji Kütüphanesi", "Önce uygulanacak kaydı seçin.")
            return False
        record = self._kutuphane().getir(record_id)
        return self._icerigi_projeye_uygula(record) if record else False

    def _kutuphane_bolum_baglantisini_temizle(self):
        self.jeoloji_kutuphanesi_bolumu_aktif = False
        self.jeoloji_kutuphanesi_kayit_id = None
        self.jeoloji_kutuphanesi_bolum_yolu = ""
        self.jeoloji_kutuphanesi_bolum_hash = ""
        self.jeoloji_kutuphanesi_uygulanan_genel = ""
        self.jeoloji_kutuphanesi_uygulanan_inceleme = ""

    def jeoloji_kutuphanesi_bolum_yolunu_coz(self):
        """Projeye bağlı kütüphane bölümünü kayıt kimliği ve özetle yeniden bulur."""
        if not getattr(self, "jeoloji_kutuphanesi_bolumu_aktif", False):
            return ""
        expected_hash = str(getattr(self, "jeoloji_kutuphanesi_bolum_hash", "") or "")
        record_id = getattr(self, "jeoloji_kutuphanesi_kayit_id", None)
        candidates = []
        if record_id:
            try:
                record = self._kutuphane().getir(int(record_id), aktif_olmayan=True)
            except (OSError, ValueError, TypeError):
                record = None
            if record:
                record_hash = str(record.get("bolum_hash") or "")
                if not expected_hash or not record_hash or record_hash == expected_hash:
                    candidates.append(str(record.get("bolum_docx_path") or ""))
        candidates.append(str(getattr(self, "jeoloji_kutuphanesi_bolum_yolu", "") or ""))
        for path in candidates:
            if path and os.path.isfile(path):
                self.jeoloji_kutuphanesi_bolum_yolu = os.path.abspath(path)
                return self.jeoloji_kutuphanesi_bolum_yolu
        return ""

    def _kutuphane_bolumunu_projeye_bagla(self, record, general, site):
        self._kutuphane_bolum_baglantisini_temizle()
        direct_id = record.get("id")
        general_id = record.get("genel_jeoloji_metni_kayit_id")
        site_id = record.get("inceleme_alani_jeolojisi_kayit_id")
        if direct_id:
            record_id = int(direct_id)
        elif general and site and general_id and general_id == site_id:
            record_id = int(general_id)
        else:
            if record.get("kayit_idleri"):
                return False, "Önerinin metinleri farklı kayıtlardan geldiği için tam Word bağlanmadı."
            return False, "Bu kayıtta bağlanabilir tam Word bölümü bulunmuyor."

        package_path = str(record.get("bolum_docx_path") or "").strip()
        if not package_path or not os.path.isfile(package_path):
            return False, "Kayıtlı JEOLOJİ Word paketi bulunamadığı için yalnız metinler uygulandı."

        self.jeoloji_kutuphanesi_bolumu_aktif = True
        self.jeoloji_kutuphanesi_kayit_id = record_id
        self.jeoloji_kutuphanesi_bolum_yolu = os.path.abspath(package_path)
        self.jeoloji_kutuphanesi_bolum_hash = str(record.get("bolum_hash") or "")
        self.jeoloji_kutuphanesi_uygulanan_genel = general
        self.jeoloji_kutuphanesi_uygulanan_inceleme = site
        self.jeoloji_sablon_yolu = ""
        return True, ""

    def _jeoloji_proje_konumu(self):
        formasyon = self.combo_formasyon.get().strip() if hasattr(self, "combo_formasyon") else ""
        if formasyon == "Seçiniz...":
            formasyon = ""
        return {
            "il": self.proje_deger("IL", "Çanakkale"),
            "ilce": self.proje_deger("ILCE", ""),
            "yerlesim": self.proje_deger("KOY", ""),
            "ada": self.proje_deger("ADA", ""),
            "parsel": self.proje_deger("PARSEL", ""),
            "formasyon": formasyon,
        }

    @staticmethod
    def _jeoloji_kayit_ozeti(record):
        konum = " / ".join(
            value for value in (
                str(record.get("ilce") or "").strip(),
                str(record.get("yerlesim") or "").strip(),
            ) if value
        ) or "Konum belirtilmemiş"
        ada_parsel = " / ".join(
            value for value in (
                str(record.get("ada") or "").strip(),
                str(record.get("parsel") or "").strip(),
            ) if value
        ) or "—"
        formasyon = str(record.get("formasyon") or "").strip() or "Formasyon belirtilmemiş"
        return f"Kayıt #{record.get('id')} · {konum} · Ada/Parsel: {ada_parsel}\n{formasyon}"

    @staticmethod
    def _jeoloji_formasyon_eslesiyor(kayit_formasyonu, proje_formasyonu):
        kayit_anahtari = jeoloji_anahtari(kayit_formasyonu)
        proje_anahtari = jeoloji_anahtari(proje_formasyonu)
        if not proje_anahtari:
            return False
        if kayit_anahtari == proje_anahtari:
            return True
        kayit_adi = jeoloji_anahtari(str(kayit_formasyonu or "").split("(", 1)[0])
        proje_adi = jeoloji_anahtari(str(proje_formasyonu or "").split("(", 1)[0])
        return bool(kayit_adi and kayit_adi == proje_adi)

    def _benzer_onayli_jeoloji_kayitlari(
        self,
        konum,
        *,
        word_zorunlu=False,
        ayni_formasyon_zorunlu=False,
    ):
        records = self._kutuphane().listele(
            il=konum.get("il") or "Çanakkale",
            onay_durumu="onayli",
        )
        if word_zorunlu:
            records = [
                record for record in records
                if str(record.get("bolum_docx_path") or "").strip()
                and os.path.isfile(str(record.get("bolum_docx_path") or ""))
            ]
        else:
            records = [
                record for record in records
                if any(
                    str(record.get(field) or "").strip()
                    for field in (
                        "genel_jeoloji_metni",
                        "inceleme_alani_jeolojisi",
                        "harita_path",
                        "bolum_docx_path",
                    )
                )
            ]
        if not records:
            return []

        keys = {
            field: jeoloji_anahtari(konum.get(field, ""))
            for field in ("ilce", "yerlesim", "ada", "parsel", "formasyon")
        }
        if keys["formasyon"]:
            same_formation = [
                record for record in records
                if self._jeoloji_formasyon_eslesiyor(
                    record.get("formasyon", ""), konum.get("formasyon", "")
                )
            ]
            if same_formation:
                records = same_formation
            elif ayni_formasyon_zorunlu:
                return []

        def rank(record):
            return (
                self._jeoloji_formasyon_eslesiyor(
                    record.get("formasyon", ""), konum.get("formasyon", "")
                ),
                bool(keys["ilce"] and record.get("ilce_key") == keys["ilce"]),
                bool(keys["yerlesim"] and record.get("yerlesim_key") == keys["yerlesim"]),
                bool(keys["ada"] and record.get("ada_key") == keys["ada"]),
                bool(keys["parsel"] and record.get("parsel_key") == keys["parsel"]),
                str(record.get("guncelleme_tarihi") or ""),
                int(record.get("id") or 0),
            )

        return sorted(records, key=rank, reverse=True)

    def _word_2_1_1_iceriyor(self, path):
        try:
            document = Document(path)
        except Exception as exc:
            self.hata_kaydet("Kütüphane Word paketi okunamadı", exc)
            return False
        return any(
            yapisal_jeoloji_basligi_mi(paragraph.text)
            for paragraph in document.paragraphs
        )

    def _word_2_1_1_ve_kesit_iceriyor(self, path):
        return bool(
            self._word_2_1_1_iceriyor(path)
            and wordde_stratigrafik_kesit_var_mi(path)
        )

    def jeoloji_2_1_1_kaynagini_bagla(self, parent=None, bolgesel_de_kullanilacak=False):
        """Projeye, mevcut 2.1 metnini değiştirmeden uygun eski Word paketini bağla."""
        konum = self._jeoloji_proje_konumu()
        candidates = self._benzer_onayli_jeoloji_kayitlari(
            konum,
            word_zorunlu=True,
            ayni_formasyon_zorunlu=bool(konum.get("formasyon")),
        )
        record = next(
            (
                candidate for candidate in candidates
                if self._word_2_1_1_ve_kesit_iceriyor(
                    candidate.get("bolum_docx_path", "")
                )
            ),
            None,
        )
        if not record:
            hedef = (
                f"'{konum['formasyon']}' formasyonu için "
                if konum.get("formasyon") else "bu proje için "
            )
            messagebox.showinfo(
                "2.1.1 Kaynağı",
                f"Kütüphanede {hedef}2.1.1 ve stratigrafik kesit içeren onaylı "
                "bir Word paketi bulunamadı.\n\n"
                "Kütüphaneden başka bir kaydı elle seçip 'Seçili Kaydı Uygula' düğmesini kullanabilirsiniz.",
                parent=parent,
            )
            return ""

        kapsam = (
            "Bu Word'deki 2.1 ve 2.1.1 kullanılacak."
            if bolgesel_de_kullanilacak
            else "Yalnız 2.1.1 bu Word'den alınacak; programın hazırladığı 2.1 değişmeyecek."
        )
        if not messagebox.askyesno(
            "2.1.1 Kaynak Raporu",
            "Birebir ada/parsel kaydı bağlı değil. En uygun onaylı kaynak bulundu:\n\n"
            f"{self._jeoloji_kayit_ozeti(record)}\n\n{kapsam}\n\nBu kayıt kullanılsın mı?",
            parent=parent,
        ):
            return ""

        general = str(record.get("genel_jeoloji_metni") or "").strip()
        site = str(record.get("inceleme_alani_jeolojisi") or "").strip()
        linked, note = self._kutuphane_bolumunu_projeye_bagla(record, general, site)
        if not linked:
            messagebox.showwarning("2.1.1 Kaynağı", note, parent=parent)
            return ""
        if hasattr(self, "lbl_jeoloji_sablon"):
            self.lbl_jeoloji_sablon.config(text=self.jeoloji_sablon_etiket_metni())
        self._durum(
            f"2.1.1 için {record.get('id')} numaralı kütüphane Word paketi projeye bağlandı."
        )
        return self.jeoloji_kutuphanesi_bolum_yolunu_coz()

    def jeoloji_onerisini_projeye_uygula(self):
        konum = self._jeoloji_proje_konumu()
        if not konum["ilce"]:
            messagebox.showinfo("Jeoloji Kütüphanesi", "Öneri için önce İlçe alanını doldurun.")
            return False
        record = self._kutuphane().uygun_icerigi_bul(**konum)
        degistirme_onayi_alindi = False
        if not record:
            candidates = self._benzer_onayli_jeoloji_kayitlari(
                konum,
                ayni_formasyon_zorunlu=bool(konum.get("formasyon")),
            )
            if not candidates:
                hedef = (
                    f"'{konum['formasyon']}' formasyonu için"
                    if konum.get("formasyon") else "bu ilçe için"
                )
                messagebox.showinfo(
                    "Jeoloji Kütüphanesi",
                    f"Kütüphanede {hedef} kullanılabilir onaylı kayıt bulunamadı.",
                )
                return False
            record = candidates[0]
            if not messagebox.askyesno(
                "Benzer Jeoloji Kaydı Bulundu",
                "Bu ada/parsel için birebir kayıt yok; aynı formasyondaki en uygun onaylı rapor bulundu:\n\n"
                f"{self._jeoloji_kayit_ozeti(record)}\n\n"
                "Formasyon bilgileri ve tam Jeoloji Word paketi bu kayıttan uygulansın mı?",
            ):
                return False
            degistirme_onayi_alindi = True
        return self._icerigi_projeye_uygula(
            record,
            degistirme_onayi_alindi=degistirme_onayi_alindi,
        )

    def _kisa_muhendislik_jeolojisi_metni(self, metin):
        """Kütüphane bölümünden yalnız sahada gözlenen birimi anlatan kısa cümleyi al."""
        temiz = " ".join(str(metin or "").split())
        if not temiz:
            return ""
        kalip = re.compile(
            r"((?:Çalışma|İnceleme)\s+alanında\s+(?:birim|zemin|kayaçlar?)\b.*?"
            r"(?:gözlenmektedir|izlenmektedir|gözlenmiştir|izlenmiştir)\s*[.!?]?)",
            re.IGNORECASE,
        )
        eslesme = kalip.search(temiz)
        if not eslesme:
            return ""
        cumle = eslesme.group(1).strip()
        return cumle if cumle.endswith((".", "!", "?")) else cumle + "."

    def _icerigi_projeye_uygula(self, record, degistirme_onayi_alindi=False):
        if not record:
            return False
        general = str(record.get("genel_jeoloji_metni") or "").strip()
        site = str(record.get("inceleme_alani_jeolojisi") or "").strip()
        kisa_muhendislik = self._kisa_muhendislik_jeolojisi_metni(site)
        kayit_formasyonu = str(record.get("formasyon") or "").strip()
        map_path = str(record.get("harita_path") or "").strip()
        package_path = str(record.get("bolum_docx_path") or "").strip()
        if not any((general, site, map_path, package_path and os.path.isfile(package_path))):
            messagebox.showinfo("Jeoloji Kütüphanesi", "Seçilen kayıtta uygulanabilir içerik yok.")
            return False
        mevcut_metin = ""
        if hasattr(self, "txt_formasyon_rapor"):
            mevcut_metin = self.txt_formasyon_rapor.get("1.0", tk.END).strip()
        if hasattr(self, "txt_muhendislik_jeolojisi"):
            mevcut_metin = mevcut_metin or self.txt_muhendislik_jeolojisi.get("1.0", tk.END).strip()
        if mevcut_metin and not degistirme_onayi_alindi and not messagebox.askyesno(
            "Jeoloji Önerisini Uygula",
            "Mevcut jeoloji metni kütüphane içeriğiyle değiştirilecek. Devam edilsin?",
        ):
            return False
        if kayit_formasyonu and hasattr(self, "combo_formasyon"):
            self.combo_formasyon.set(kayit_formasyonu)
        if general and hasattr(self, "txt_formasyon_rapor"):
            self.txt_formasyon_rapor.delete("1.0", tk.END)
            self.txt_formasyon_rapor.insert("1.0", general)
        if kisa_muhendislik and hasattr(self, "txt_muhendislik_jeolojisi"):
            self.txt_muhendislik_jeolojisi.delete("1.0", tk.END)
            self.txt_muhendislik_jeolojisi.insert("1.0", kisa_muhendislik)
        if map_path and os.path.isfile(map_path):
            self.img_mjh = map_path
        word_bagli, word_notu = self._kutuphane_bolumunu_projeye_bagla(record, general, site)
        self.jeoloji_kutuphanesi_uygulanan_inceleme = kisa_muhendislik
        if hasattr(self, "lbl_jeoloji_sablon"):
            self.lbl_jeoloji_sablon.config(text=self.jeoloji_sablon_etiket_metni())
        if word_bagli:
            durum = "Resim ve tabloları içeren tam Jeoloji Word bölümü projeye bağlandı."
            if kisa_muhendislik:
                durum += " Kısa mühendislik jeolojisi cümlesi ayrıca alındı."
        else:
            durum = f"Kullanılabilir jeoloji bilgileri projeye uygulandı. {word_notu}"
        self._durum(durum + " Değişiklikleri kaydetmeyi unutmayın.")
        return True

    def _jeoloji_harita_nesnelerini_temizle(self):
        for polygon in list(getattr(self, "jeoloji_kutuphane_polygonlari", []) or []):
            try:
                polygon.delete()
            except (AttributeError, tk.TclError):
                pass
        self.jeoloji_kutuphane_polygonlari = []
        self.jeoloji_kutuphane_secili_polygon = None

    def jeoloji_harita_ciktisi_baslat(self):
        """Rapor haritası yakalanırken kütüphane parsellerini geçici olarak gizle."""
        after_id = getattr(self, "_jeoloji_harita_after_id", None)
        if after_id:
            try:
                self.root.after_cancel(after_id)
            except (tk.TclError, ValueError):
                pass
        self._jeoloji_harita_after_id = None
        self._jeoloji_harita_nesnelerini_temizle()
        self._jeoloji_harita_son_gorunum = None

    def jeoloji_harita_ciktisi_bitir(self):
        """Çıktı sonrasında, kullanıcı katmanı açıksa etkileşimli görünümü geri yükle."""
        self._jeoloji_harita_son_gorunum = None
        if (
            hasattr(self, "map_widget")
            and hasattr(self, "jeoloji_kutuphane_harita_var")
            and self.jeoloji_kutuphane_harita_var.get()
        ):
            self.jeoloji_harita_katmanini_yenile(zorla=True)

    def _jeoloji_harita_gorunum_kutusu(self):
        width = max(int(self.map_widget.winfo_width()), 2)
        height = max(int(self.map_widget.winfo_height()), 2)
        sol_ust = self.map_widget.convert_canvas_coords_to_decimal_coords(0, 0)
        sag_alt = self.map_widget.convert_canvas_coords_to_decimal_coords(width, height)
        return {
            "min_enlem": min(sol_ust[0], sag_alt[0]),
            "max_enlem": max(sol_ust[0], sag_alt[0]),
            "min_boylam": min(sol_ust[1], sag_alt[1]),
            "max_boylam": max(sol_ust[1], sag_alt[1]),
        }

    def _jeoloji_harita_izlemeyi_programla(self):
        onceki = getattr(self, "_jeoloji_harita_after_id", None)
        if onceki:
            try:
                self.root.after_cancel(onceki)
            except (tk.TclError, ValueError):
                pass
        try:
            self._jeoloji_harita_after_id = self.root.after(
                1200, lambda: self.jeoloji_harita_katmanini_yenile(zorla=False)
            )
        except tk.TclError:
            self._jeoloji_harita_after_id = None

    def jeoloji_harita_katmanini_degistir(self):
        if not self.jeoloji_kutuphane_harita_var.get():
            self._jeoloji_harita_nesnelerini_temizle()
            self._jeoloji_harita_son_gorunum = None
            self.jeoloji_harita_durum.set("Jeoloji kütüphanesi katmanı kapalı")
            return
        self.jeoloji_harita_katmanini_yenile(zorla=True)

    def jeoloji_harita_katmanini_yenile(self, zorla=True):
        self._jeoloji_harita_after_id = None
        if getattr(self.app, "_harita_disari_aktarim_aktif", False):
            self._jeoloji_harita_nesnelerini_temizle()
            self._jeoloji_harita_son_gorunum = None
            return
        if not hasattr(self, "map_widget") or not self.jeoloji_kutuphane_harita_var.get():
            return
        try:
            kutu = self._jeoloji_harita_gorunum_kutusu()
            gorunum = (
                round(kutu["min_enlem"], 5), round(kutu["max_enlem"], 5),
                round(kutu["min_boylam"], 5), round(kutu["max_boylam"], 5),
                int(round(float(getattr(self.map_widget, "zoom", 0)))),
                bool(self.jeoloji_kutuphane_taslak_var.get()),
            )
            if not zorla and gorunum == getattr(self, "_jeoloji_harita_son_gorunum", None):
                self._jeoloji_harita_izlemeyi_programla()
                return
            records = self._kutuphane().harita_kayitlari(
                **kutu,
                taslaklari_goster=self.jeoloji_kutuphane_taslak_var.get(),
            )
            self._jeoloji_harita_nesnelerini_temizle()
            for record in records:
                taslak = record.get("onay_durumu") != "onayli"
                polygon = self.map_widget.set_polygon(
                    record["noktalar"],
                    fill_color=JEOLOJI_TASLAK_DOLGU_RENGI if taslak else JEOLOJI_ONAYLI_DOLGU_RENGI,
                    outline_color=JEOLOJI_TASLAK_SINIR_RENGI if taslak else JEOLOJI_ONAYLI_SINIR_RENGI,
                    border_width=3,
                    command=self._jeoloji_harita_kaydi_sec,
                    name=f"Jeoloji kaydı #{record['id']}",
                    data=record,
                )
                self.jeoloji_kutuphane_polygonlari.append(polygon)
            mevcut = getattr(self, "kml_polygon_obj", None)
            if mevcut is not None and getattr(mevcut, "canvas_polygon", None):
                self.map_widget.canvas.itemconfig(
                    mevcut.canvas_polygon,
                    outline=CALISAN_PARSEL_SINIR_RENGI,
                    width=CALISAN_PARSEL_SINIR_KALINLIGI,
                    fill="",
                )
                self.map_widget.canvas.tag_raise(mevcut.canvas_polygon)
            self._jeoloji_harita_son_gorunum = gorunum
            rapor_sayisi = len({int(record["id"]) for record in records})
            self.jeoloji_harita_durum.set(
                f"{JEOLOJI_HARITA_RENK_ACIKLAMASI} · "
                f"{rapor_sayisi} rapora ait {len(records)} kütüphane parseli"
            )
        except Exception as exc:
            self.jeoloji_harita_durum.set(f"Jeoloji katmanı yüklenemedi: {exc}")
            try:
                self.hata_kaydet("Jeoloji kütüphanesi harita katmanı yüklenemedi", exc)
            except Exception:
                pass
        finally:
            if self.jeoloji_kutuphane_harita_var.get():
                self._jeoloji_harita_izlemeyi_programla()

    def _jeoloji_harita_kaydi_sec(self, polygon):
        onceki = getattr(self, "jeoloji_kutuphane_secili_polygon", None)
        if onceki is not None and getattr(onceki, "canvas_polygon", None):
            onceki_record = getattr(onceki, "data", {}) or {}
            taslak = onceki_record.get("onay_durumu") != "onayli"
            self.map_widget.canvas.itemconfig(
                onceki.canvas_polygon,
                outline=JEOLOJI_TASLAK_SINIR_RENGI if taslak else JEOLOJI_ONAYLI_SINIR_RENGI,
                width=3,
            )
        self.jeoloji_kutuphane_secili_polygon = polygon
        if getattr(polygon, "canvas_polygon", None):
            self.map_widget.canvas.itemconfig(
                polygon.canvas_polygon, outline=JEOLOJI_SECILI_SINIR_RENGI, width=5
            )
            self.map_widget.canvas.tag_raise(polygon.canvas_polygon)
        self._jeoloji_harita_kaydi_penceresi((getattr(polygon, "data", {}) or {}).get("id"))

    def _jeoloji_harita_kaydi_penceresi(self, record_id):
        record = self._kutuphane().getir(record_id) if record_id else None
        if not record:
            return
        eski = getattr(self, "_jeoloji_harita_detay_penceresi", None)
        try:
            if eski is not None and eski.winfo_exists():
                eski.destroy()
        except tk.TclError:
            pass
        pencere = tk.Toplevel(self.root)
        self._jeoloji_harita_detay_penceresi = pencere
        pencere.title(f"Jeoloji Raporu #{record['id']}")
        pencere.geometry("680x560")
        pencere.transient(self.root)
        ada_parsel = " / ".join(
            part for part in (record.get("ada", ""), record.get("parsel", "")) if part
        ) or "—"
        bilgi = (
            f"{record.get('ilce', '')} / {record.get('yerlesim', '')}\n"
            f"Ada / Parsel: {ada_parsel}\n"
            f"Formasyon: {record.get('formasyon') or '—'}\n"
            f"Durum: {record.get('onay_durumu', 'taslak')} · Revizyon {record.get('revizyon_no', 1)}"
        )
        ttk.Label(pencere, text=bilgi, justify="left", padding=12).pack(fill="x")
        notebook = ttk.Notebook(pencere)
        notebook.pack(fill="both", expand=True, padx=12, pady=(0, 10))
        for baslik, alan in (
            ("Genel Jeoloji", "genel_jeoloji_metni"),
            ("İnceleme Alanı", "inceleme_alani_jeolojisi"),
        ):
            frame = ttk.Frame(notebook)
            notebook.add(frame, text=baslik)
            text = tk.Text(frame, wrap="word")
            scroll = ttk.Scrollbar(frame, orient="vertical", command=text.yview)
            text.configure(yscrollcommand=scroll.set)
            text.insert("1.0", record.get(alan, "") or "Bu bölüm için metin bulunmuyor.")
            text.configure(state="disabled")
            text.pack(side="left", fill="both", expand=True)
            scroll.pack(side="right", fill="y")
        alt = ttk.Frame(pencere, padding=(12, 0, 12, 12))
        alt.pack(fill="x")
        ttk.Button(alt, text="Kapat", command=pencere.destroy).pack(side="right")
        ttk.Button(
            alt,
            text="Kütüphanede Aç",
            command=lambda: self._jeoloji_kaydini_kutuphanede_ac(record["id"]),
        ).pack(side="right", padx=(0, 6))
        ttk.Button(
            alt,
            text="Jeolojiyi Projeye Uygula",
            command=lambda: self._icerigi_projeye_uygula(record),
        ).pack(side="right", padx=(0, 6))
        if record.get("kaynak_rapor_path"):
            ttk.Button(
                alt,
                text="Kaynak Raporu Aç",
                command=lambda: self._kaynak_raporu_ac(record),
            ).pack(side="left")

    def _jeoloji_kaydini_kutuphanede_ac(self, record_id):
        self.jeoloji_kutuphanesi_penceresi()
        tree = getattr(self, "_jeoloji_kutuphanesi_tree", None)
        if tree is not None and tree.exists(str(record_id)):
            tree.selection_set(str(record_id))
            tree.focus(str(record_id))
            tree.see(str(record_id))
            self._secim_degisti()

    def _kaynak_raporu_ac(self, record):
        path = str(record.get("kaynak_rapor_path") or "")
        if not path or not os.path.isfile(path):
            messagebox.showwarning("Kaynak Rapor", "Kaynak Word dosyası artık bulunduğu yerde değil.")
            return
        try:
            os.startfile(path)
        except OSError as exc:
            messagebox.showerror("Kaynak Rapor", f"Dosya açılamadı:\n{exc}")

    def _dosya_sec(self, key):
        if key == "harita_path":
            path = filedialog.askopenfilename(
                parent=self._jeoloji_kutuphanesi_penceresi,
                title="Jeoloji haritası seç",
                filetypes=[("Görseller", "*.png *.jpg *.jpeg *.tif *.tiff *.bmp"), ("Tüm dosyalar", "*.*")],
            )
        else:
            path = filedialog.askopenfilename(
                parent=self._jeoloji_kutuphanesi_penceresi,
                title="Kaynak Word raporu seç",
                filetypes=[("Word dosyaları", "*.docx"), ("Tüm dosyalar", "*.*")],
            )
        if path:
            self._widget_yaz(self._jeoloji_kutuphanesi_form_vars[key], path)

    def _yedekle(self):
        yol = filedialog.asksaveasfilename(
            parent=self._jeoloji_kutuphanesi_penceresi,
            title="Jeoloji kütüphanesini yedekle",
            initialfile="canakkale_jeoloji_yedek.zip",
            defaultextension=".zip",
            filetypes=[("ZIP arşivi", "*.zip")],
        )
        if not yol:
            return
        try:
            sonuc = self._kutuphane().yedek_paketi_olustur(yol)
        except OSError as exc:
            messagebox.showerror("Jeoloji Kütüphanesi", f"Yedek oluşturulamadı:\n{exc}")
            return
        self._durum(f"Jeoloji kütüphanesi yedeklendi: {sonuc}")

    def _arsivle(self):
        record_id = self._jeoloji_kutuphanesi_secili_id
        if not record_id:
            messagebox.showinfo("Jeoloji Kütüphanesi", "Önce arşivlenecek kaydı seçin.")
            return
        if not messagebox.askyesno(
            "Kaydı Arşivle",
            "Seçili kayıt etkin listeden kaldırılacak; revizyon geçmişi korunacak. Devam edilsin mi?",
            parent=self._jeoloji_kutuphanesi_penceresi,
        ):
            return
        self._kutuphane().arsivle(record_id)
        self._yeni_kayit(bildir=False)
        self._listeyi_yenile()
        self._durum(f"Jeoloji kütüphanesi kaydı arşivlendi: #{record_id}")
