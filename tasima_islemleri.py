import logging
import math
import tkinter as tk
from tkinter import ttk, messagebox

from tasima import BOWLES_ZEMIN_KATSAYISI, TBDY2018TasimaGucu


logger = logging.getLogger("ZeminRaporPro")


KILONEWTON_PER_TON_FORCE = 9.81


def tasima_qt_asagi_yuvarla(qt, ondalik=2):
    """Hesaplanan q_t değerini rapor hassasiyetinde yukarı çıkarmadan yuvarlar."""
    qt = float(qt)
    if not math.isfinite(qt):
        raise ValueError("Hesaplanan q_t sonlu bir sayı olmalıdır.")
    carpan = 10 ** int(ondalik)
    return math.floor(qt * carpan + 1e-12) / carpan


class TasimaIslemleri:
    def __init__(self, app):
        object.__setattr__(self, "app", app)

    def __getattr__(self, name):
        return getattr(self.app, name)

    def __setattr__(self, name, value):
        if name == "app":
            object.__setattr__(self, name, value)
        else:
            setattr(self.app, name, value)

    def sekme6_tasima(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="6. Taşıma Gücü")
        
        self.zemin_kaya_var = tk.StringVar(value="zemin")
        ttk.Radiobutton(frame, text="Zemin", variable=self.zemin_kaya_var, value="zemin", command=self.tasima_ekran_guncelle).pack(pady=5)
        ttk.Radiobutton(frame, text="Kaya", variable=self.zemin_kaya_var, value="kaya", command=self.tasima_ekran_guncelle).pack(pady=5)
        
        self.input_frame_zemin = ttk.LabelFrame(frame, text="TBDY 2018 Zemin Parametreleri", bootstyle="info")
        self.input_frame_kaya = ttk.LabelFrame(frame, text="Kaya Taşıma Gücü Parametreleri", bootstyle="info")
        
        self.tg_girdiler = {}
        self.tasima_dayanim_23_uygulandi = tk.BooleanVar(value=False)
        self.tasima_dayanim_23_kaynak_c = ""
        self.tasima_dayanim_23_kaynak_phi = ""
        
        parametreler_zemin = [
            ("Efektif Kohezyon (c) [kPa]", "c", ""),
            ("İçsel Sürtünme Açısı (phi) [°]", "phi", ""),
            ("Doğal Birim Hacim Ağırlık [gr/cm3]", "gn", ""),
            ("Doygun Birim Hacim Ağırlık [gr/cm3]", "gsat", ""),
            ("YASS var; derinliği [m]", "yass", ""),
            ("Etkili Temel Genişliği (B′, kısa boyut) [m]", "B", ""),
            ("Etkili Temel Uzunluğu (L′) [m]", "L", ""),
            ("Temel Derinliği (Df) [m]", "Df", ""),
            ("TBDY Dayanım Katsayısı (γRv)", "RvGk", "1.4")
        ]
        self.tasima_yass_var = tk.BooleanVar(value=False)
        for i, (etiket, kod, varsayilan) in enumerate(parametreler_zemin):
            if kod == "yass":
                ttk.Checkbutton(
                    self.input_frame_zemin,
                    text=etiket,
                    variable=self.tasima_yass_var,
                    command=self.tasima_yass_durum_guncelle,
                    bootstyle="info",
                ).grid(row=i, column=0, sticky="e", padx=5, pady=2)
            else:
                ttk.Label(self.input_frame_zemin, text=etiket).grid(row=i, column=0, sticky="e", padx=5, pady=2)
            entry = ttk.Entry(self.input_frame_zemin)
            entry.grid(row=i, column=1, padx=5, pady=2)
            entry.insert(0, varsayilan)
            self.tg_girdiler[kod] = entry

        dayanim_23_frame = ttk.Frame(self.input_frame_zemin)
        dayanim_23_frame.grid(row=0, column=2, rowspan=2, sticky="w", padx=(12, 5), pady=2)
        self.btn_tasima_dayanim_23 = ttk.Button(
            dayanim_23_frame,
            text="c′ ve φ′ Değerlerini 2/3'e Düşür",
            command=self.tasima_dayanim_23_degistir,
            bootstyle="warning-outline",
        )
        self.btn_tasima_dayanim_23.pack(anchor="w")
        self.lbl_tasima_dayanim_23 = ttk.Label(dayanim_23_frame, text="2/3 indirgeme uygulanmadı")
        self.lbl_tasima_dayanim_23.pack(anchor="w", pady=(3, 0))
        self.tasima_yass_durum_guncelle(temizle=False)

        self.tasima_varsayim_onayi = tk.BooleanVar(value=False)
        self.tasima_rapor_imzasi = None
        ttk.Checkbutton(
            self.input_frame_zemin,
            text=(
                "Yükün düşey ve merkezî; temel tabanı ile zemin yüzeyinin yatay olduğunu "
                "doğruluyorum (i = g = b = 1)."
            ),
            variable=self.tasima_varsayim_onayi,
            bootstyle="warning",
        ).grid(row=len(parametreler_zemin), column=0, columnspan=2, sticky="w", padx=8, pady=(8, 5))
            
        parametreler_kaya = [
            ("KAYA: Karakteristik Taşıma Gücü (qk) [t/m2]", "qt", ""),
            ("KAYA: Dayanım/Güvenlik Katsayısı (Gk)", "Gk", "")
        ]
        for i, (etiket, kod, varsayilan) in enumerate(parametreler_kaya):
            ttk.Label(self.input_frame_kaya, text=etiket).grid(row=i, column=0, sticky="e", padx=5, pady=2)
            entry = ttk.Entry(self.input_frame_kaya)
            entry.grid(row=i, column=1, padx=5, pady=2)
            entry.insert(0, varsayilan)
            self.tg_girdiler[kod] = entry
            
        self.tasima_ekran_guncelle()
        
        # Sonuçlar ve Nihai Çıktılar
        sonuc_frame = ttk.LabelFrame(frame, text="Nihai Manuel Veri Girişi ve Word Raporu Metni", bootstyle="info")
        sonuc_frame.pack(padx=10, pady=5, fill='x')
        
        ttk.Label(sonuc_frame, text="Hesaplanmış qk (Bilgi):").grid(row=0, column=0, sticky="e")
        self.lbl_sonuc = ttk.Label(sonuc_frame, text="-", font=("Arial", 10, "bold"), foreground="blue")
        self.lbl_sonuc.grid(row=0, column=1, pady=2, sticky="w")
        
        ttk.Label(sonuc_frame, text="Raporda kullanılacak qt (t/m2):").grid(row=1, column=0, pady=2, sticky="e")
        self.entry_qt_nihai = ttk.Entry(sonuc_frame)
        self.entry_qt_nihai.grid(row=1, column=1, pady=2, sticky="w")

        self.lbl_ks_carpani = ttk.Label(
            sonuc_frame, text="Bowles zemin katsayısı (sabit):"
        )
        self.lbl_ks_carpani.grid(row=2, column=0, pady=2, sticky="e")
        self.entry_ks_carpani = ttk.Entry(sonuc_frame)
        self.entry_ks_carpani.grid(row=2, column=1, pady=2, sticky="w")
        self.entry_ks_carpani.insert(0, f"{BOWLES_ZEMIN_KATSAYISI:.0f}")
        self.tg_girdiler["ks_carpani"] = self.entry_ks_carpani
        self.tasima_ks_carpani_durum_guncelle()
        
        ttk.Label(sonuc_frame, text="Nihai ks (t/m3):").grid(row=3, column=0, pady=2, sticky="e")
        self.entry_ks_nihai = ttk.Entry(sonuc_frame)
        self.entry_ks_nihai.grid(row=3, column=1, pady=2, sticky="w")
        
        btn_f = ttk.Frame(frame)
        btn_f.pack(pady=5)
        ttk.Button(btn_f, text="1. Formülleri Hesapla", command=self.tasima_hesapla, style="Secondary.TButton").pack(side="left", padx=5)
        ttk.Button(btn_f, text="2. Word Metnini Oluştur", command=self.tasima_metni_olustur, style="Primary.TButton").pack(side="left", padx=5)
        
        txt_rf = ttk.LabelFrame(
            frame,
            text="Düzenlenebilir Rapor Metni",
        )
        txt_rf.pack(padx=10, pady=5, fill='both', expand=True)
        ttk.Label(
            txt_rf,
            text="Tablonun konumunu korumak için [TABLO_BURADA] satırını silmeyin.",
            style="Muted.TLabel",
        ).pack(anchor="w", padx=5, pady=(0, 3))
        self.txt_tasima_rapor = tk.Text(txt_rf, height=13, wrap='word')
        self.txt_tasima_rapor.pack(padx=5, pady=5, fill='both', expand=True)

    def tasima_yass_var_mi(self):
        variable = getattr(self, "tasima_yass_var", None)
        if variable is not None:
            return bool(variable.get())
        # Eski proje kayıtlarında ayrı tik yoktu; 999 değeri "YASS yok" olarak
        # kullanılıyordu. Eski projeleri açarken bu anlamı koru.
        entry = getattr(self, "tg_girdiler", {}).get("yass")
        raw = str(entry.get()).strip().replace(",", ".") if entry is not None else ""
        if not raw:
            return False
        try:
            return float(raw) < 999.0
        except ValueError:
            return True

    def tasima_dayanim_23_uygulandi_mi(self):
        variable = getattr(self, "tasima_dayanim_23_uygulandi", None)
        return bool(variable.get()) if variable is not None else False

    def tasima_dayanim_23_arayuz_guncelle(self):
        uygulandi = self.tasima_dayanim_23_uygulandi_mi()
        button = getattr(self, "btn_tasima_dayanim_23", None)
        if button is not None:
            button.config(
                text=(
                    "Orijinal c′ ve φ′ Değerlerine Dön"
                    if uygulandi
                    else "c′ ve φ′ Değerlerini 2/3'e Düşür"
                )
            )
        label = getattr(self, "lbl_tasima_dayanim_23", None)
        if label is not None:
            label.config(
                text=(
                    "2/3 indirgeme uygulandı; değerler elle düzenlenebilir"
                    if uygulandi
                    else "2/3 indirgeme uygulanmadı"
                )
            )

    @staticmethod
    def tasima_dayanim_degerini_yaz(deger):
        return f"{float(deger):.6f}".rstrip("0").rstrip(".")

    def tasima_dayanim_23_degistir(self):
        if self.tasima_dayanim_23_uygulandi_mi():
            kaynak_c = str(getattr(self, "tasima_dayanim_23_kaynak_c", "")).strip()
            kaynak_phi = str(getattr(self, "tasima_dayanim_23_kaynak_phi", "")).strip()
            if not kaynak_c or not kaynak_phi:
                messagebox.showwarning(
                    "2/3 Dayanım İndirgemesi",
                    "Orijinal kohezyon veya içsel sürtünme açısı kaydı bulunamadı.",
                )
                return
            yeni_degerler = {"c": kaynak_c, "phi": kaynak_phi}
            self.tasima_dayanim_23_uygulandi.set(False)
            self.tasima_dayanim_23_kaynak_c = ""
            self.tasima_dayanim_23_kaynak_phi = ""
        else:
            hatalar = []
            degerler = {}
            for kod, etiket in (
                ("c", "Efektif kohezyon c′"),
                ("phi", "İçsel sürtünme açısı φ′"),
            ):
                try:
                    degerler[kod] = self.sayisal_tasima_girdisi_oku(kod, etiket)
                except ValueError as exc:
                    hatalar.append(str(exc))
            if not hatalar and degerler["c"] < 0:
                hatalar.append("Efektif kohezyon c′ negatif olamaz.")
            if not hatalar and not 0 <= degerler["phi"] <= 45:
                hatalar.append("İçsel sürtünme açısı φ′ 0 ile 45 derece arasında olmalıdır.")
            if hatalar:
                self.tasima_giris_hatasi_goster(hatalar)
                return

            self.tasima_dayanim_23_kaynak_c = self.tg_girdiler["c"].get().strip()
            self.tasima_dayanim_23_kaynak_phi = self.tg_girdiler["phi"].get().strip()
            yeni_degerler = {
                "c": self.tasima_dayanim_degerini_yaz(degerler["c"] * 2.0 / 3.0),
                "phi": self.tasima_dayanim_degerini_yaz(degerler["phi"] * 2.0 / 3.0),
            }
            self.tasima_dayanim_23_uygulandi.set(True)

        for kod, deger in yeni_degerler.items():
            entry = self.tg_girdiler[kod]
            entry.delete(0, tk.END)
            entry.insert(0, deger)
        self.tasima_dayanim_23_arayuz_guncelle()
        if hasattr(self, "entry_qt_nihai"):
            self._tasima_sonuclarini_temizle(raporu_temizle=True)

    def tasima_yass_durum_guncelle(self, temizle=True):
        entry = getattr(self, "tg_girdiler", {}).get("yass")
        if entry is not None:
            entry.configure(state="normal" if self.tasima_yass_var_mi() else "disabled")
        if temizle and hasattr(self, "entry_qt_nihai"):
            self._tasima_sonuclarini_temizle(raporu_temizle=True)

    def tasima_ks_carpani_durum_guncelle(self):
        entry = getattr(self, "entry_ks_carpani", None)
        label = getattr(self, "lbl_ks_carpani", None)
        if entry is None:
            return
        if self.zemin_kaya_var.get() == "zemin":
            entry.configure(state="normal")
            entry.delete(0, tk.END)
            entry.insert(0, f"{BOWLES_ZEMIN_KATSAYISI:.0f}")
            entry.configure(state="disabled")
            if label is not None:
                label.configure(text="Bowles zemin katsayısı (sabit):")
        else:
            entry.configure(state="normal")
            if label is not None:
                label.configure(text="Kaya yatak katsayısı çarpanı:")

    def tasima_ekran_guncelle(self):
        if hasattr(self, "input_frame_zemin"):
            if self.zemin_kaya_var.get() == "zemin":
                self.input_frame_kaya.forget()
                self.input_frame_zemin.pack(padx=10, pady=5, fill='x')
            else:
                self.input_frame_zemin.forget()
                self.input_frame_kaya.pack(padx=10, pady=5, fill='x')
        self.tasima_yass_durum_guncelle(temizle=False)
        self.tasima_ks_carpani_durum_guncelle()
        self.tasima_dayanim_23_arayuz_guncelle()
        if hasattr(self, "entry_qt_nihai"):
            self._tasima_sonuclarini_temizle(raporu_temizle=True)

    def sayisal_tasima_girdisi_oku(self, kod, etiket):
        ham = self.tg_girdiler[kod].get().strip().replace(",", ".")
        if ham == "":
            raise ValueError(f"{etiket} boş bırakılamaz.")
        try:
            deger = float(ham)
        except ValueError:
            raise ValueError(f"{etiket} sayısal olmalıdır.")
        if not math.isfinite(deger):
            raise ValueError(f"{etiket} sonlu bir sayı olmalıdır.")
        return deger

    def tasima_giris_hatasi_goster(self, hatalar):
        mesaj = "Lütfen aşağıdaki taşıma gücü girişlerini kontrol edin:\n\n" + "\n".join(f"- {h}" for h in hatalar)
        logger.warning("Taşıma gücü giriş kontrolü başarısız: %s", "; ".join(hatalar))
        messagebox.showwarning("Taşıma Gücü Giriş Kontrolü", mesaj)

    def kaya_tasima_girdilerini_oku(self):
        hatalar = []
        degerler = {}
        for kod, etiket in [
            ("qt", "Kaya karakteristik taşıma gücü qk"),
            ("Gk", "Kaya dayanım/güvenlik katsayısı Gk"),
            ("ks_carpani", "Yatak katsayısı çarpanı"),
        ]:
            try:
                degerler[kod] = self.sayisal_tasima_girdisi_oku(kod, etiket)
            except ValueError as e:
                hatalar.append(str(e))

        if not hatalar:
            if degerler["qt"] <= 0:
                hatalar.append("Kaya karakteristik taşıma gücü qk sıfırdan büyük olmalıdır.")
            if degerler["Gk"] <= 0:
                hatalar.append("Kaya güvenlik katsayısı Gk sıfırdan büyük olmalıdır.")
            if degerler["ks_carpani"] <= 0:
                hatalar.append("Yatak katsayısı çarpanı sıfırdan büyük olmalıdır.")

        if hatalar:
            self.tasima_giris_hatasi_goster(hatalar)
            return None
        return degerler

    def zemin_tasima_girdilerini_oku(self):
        alanlar = [
            ("c", "Efektif kohezyon c"),
            ("phi", "İçsel sürtünme açısı phi"),
            ("gn", "Doğal birim hacim ağırlık"),
            ("gsat", "Doygun birim hacim ağırlık"),
            ("B", "Temel genişliği B"),
            ("L", "Temel uzunluğu L"),
            ("Df", "Temel derinliği Df"),
            ("RvGk", "Dayanım/güvenlik katsayısı Rv"),
        ]
        hatalar = []
        degerler = {}
        for kod, etiket in alanlar:
            try:
                degerler[kod] = self.sayisal_tasima_girdisi_oku(kod, etiket)
            except ValueError as e:
                hatalar.append(str(e))

        yass_var = self.tasima_yass_var_mi()
        degerler["yass_var"] = yass_var
        if yass_var:
            try:
                degerler["yass"] = self.sayisal_tasima_girdisi_oku(
                    "yass", "YASS derinliği"
                )
            except ValueError as e:
                hatalar.append(str(e))
        elif "Df" in degerler and "B" in degerler:
            # Su seviyesi temel tabanından B kadar veya daha derindeyse taşıma
            # gücü hesabındaki birim hacim ağırlık düzeltmesi etkisizdir.
            degerler["yass"] = degerler["Df"] + degerler["B"]
        degerler["ks_carpani"] = BOWLES_ZEMIN_KATSAYISI

        if not hatalar:
            if degerler["c"] < 0:
                hatalar.append("Efektif kohezyon c negatif olamaz.")
            if not 0 <= degerler["phi"] <= 45:
                hatalar.append("İçsel sürtünme açısı phi 0 ile 45 derece arasında olmalıdır.")
            if not 0.5 <= degerler["gn"] <= 3.5:
                hatalar.append("Doğal birim hacim ağırlık 0.5 ile 3.5 gr/cm3 arasında olmalıdır.")
            if not 1.0 < degerler["gsat"] <= 3.5:
                hatalar.append("Doygun birim hacim ağırlık 1.0 ile 3.5 gr/cm3 arasında olmalıdır.")
            if degerler["gsat"] < degerler["gn"]:
                hatalar.append("Doygun birim hacim ağırlık doğal birim hacim ağırlıktan küçük olamaz.")
            if yass_var and degerler["yass"] < 0:
                hatalar.append("YASS derinliği negatif olamaz.")
            if degerler["B"] <= 0:
                hatalar.append("Temel genişliği B sıfırdan büyük olmalıdır.")
            if degerler["L"] <= 0:
                hatalar.append("Temel uzunluğu L sıfırdan büyük olmalıdır.")
            if degerler["Df"] < 0:
                hatalar.append("Temel derinliği Df negatif olamaz.")
            if not math.isclose(degerler["RvGk"], 1.4, rel_tol=0.0, abs_tol=1e-9):
                hatalar.append("TBDY Tablo 16.2 uyarınca temel taşıma gücü için γRv 1.4 olmalıdır.")
            if degerler["B"] > degerler["L"]:
                hatalar.append("B etkili kısa boyuttur ve L'den büyük olamaz (B ≤ L).")
            if not getattr(self, "tasima_varsayim_onayi", None) or not self.tasima_varsayim_onayi.get():
                hatalar.append("i=g=b=1 kabulü için düşey/merkezî yük ve yatay yüzey varsayımını doğrulayın.")

        if hatalar:
            self.tasima_giris_hatasi_goster(hatalar)
            return None
        return degerler

    def nihai_qt_oku(self):
        ham = self.entry_qt_nihai.get().strip().replace(",", ".") if hasattr(self, "entry_qt_nihai") else ""
        if ham == "":
            raise ValueError("Nihai qt değeri boş bırakılamaz.")
        try:
            qt_val = float(ham)
        except ValueError:
            raise ValueError("Nihai qt değeri sayısal olmalıdır.")
        if not math.isfinite(qt_val):
            raise ValueError("Nihai qt değeri sonlu bir sayı olmalıdır.")
        if qt_val <= 0:
            raise ValueError("Nihai qt değeri sıfırdan büyük olmalıdır.")
        if not math.isclose(qt_val, round(qt_val, 2), rel_tol=0.0, abs_tol=1e-9):
            raise ValueError("Nihai qt değeri en fazla iki ondalık basamak içermelidir.")
        return qt_val

    def hesaplanmis_qk_oku(self):
        qk_str = self.lbl_sonuc.cget("text") if hasattr(self, "lbl_sonuc") else ""
        qk_ilk = qk_str.split()[0] if " " in qk_str else qk_str
        qk_ilk = qk_ilk.replace(",", ".")
        try:
            qk_val = float(qk_ilk)
        except ValueError:
            raise ValueError("Zemin rapor metni için önce 'Formülleri Hesapla' butonuyla qk değeri hesaplanmalıdır.")
        if not math.isfinite(qk_val):
            raise ValueError("Hesaplanmış qk değeri sonlu bir sayı olmalıdır.")
        if qk_val <= 0:
            raise ValueError("Hesaplanmış qk değeri sıfırdan büyük olmalıdır.")
        return qk_val

    def _zemin_tasima_sonuclari(self, degerler):
        analizci = TBDY2018TasimaGucu(
            degerler["c"],
            degerler["phi"],
            degerler["gn"] * KILONEWTON_PER_TON_FORCE,
            degerler["gsat"] * KILONEWTON_PER_TON_FORCE,
            degerler["yass"],
        )
        qk_kn, qt_kn = analizci.analiz_yap(
            degerler["B"],
            degerler["L"],
            degerler["Df"],
            gamma_Rv=degerler["RvGk"],
        )
        return analizci, qk_kn / KILONEWTON_PER_TON_FORCE, qt_kn / KILONEWTON_PER_TON_FORCE

    def _tasima_sonuclarini_yaz(self, qk, qt, ks):
        self.son_qk = qk
        self.son_qt = qt
        self.lbl_sonuc.config(
            text=f"{qk:.2f}",
            foreground="blue",
        )
        self.entry_qt_nihai.delete(0, tk.END)
        self.entry_qt_nihai.insert(0, f"{qt:.2f}")
        self.entry_ks_nihai.delete(0, tk.END)
        self.entry_ks_nihai.insert(0, f"{ks:.2f}")

    def _tasima_sonuclarini_temizle(self, raporu_temizle=False):
        self.son_qk = None
        self.son_qt = None
        self.tasima_rapor_imzasi = None
        if hasattr(self, "lbl_sonuc"):
            self.lbl_sonuc.config(text="-", foreground="blue")
        for ad in ("entry_qt_nihai", "entry_ks_nihai"):
            entry = getattr(self, ad, None)
            if entry is not None:
                entry.delete(0, tk.END)
        if raporu_temizle and hasattr(self, "txt_tasima_rapor"):
            self.txt_tasima_rapor.delete("1.0", tk.END)

    def tasima_girdi_imzasi_olustur(self):
        tur = self.zemin_kaya_var.get() if hasattr(self, "zemin_kaya_var") else ""
        if tur == "kaya":
            aktif_kodlar = {"qt", "Gk", "ks_carpani"}
            yass_var = False
        else:
            aktif_kodlar = {"c", "phi", "gn", "gsat", "yass", "B", "L", "Df", "RvGk"}
            yass_var = self.tasima_yass_var_mi()
            if not yass_var:
                aktif_kodlar.discard("yass")
        girdiler = tuple(
            (kod, entry.get().strip())
            for kod, entry in sorted(getattr(self, "tg_girdiler", {}).items())
            if kod in aktif_kodlar
        )
        ks_nihai = self.entry_ks_nihai.get().strip() if hasattr(self, "entry_ks_nihai") else ""
        girdiler += (("__ks_nihai__", ks_nihai), ("__yass_var__", "1" if yass_var else "0"))
        if tur != "kaya":
            girdiler += ((
                "__dayanim_23__",
                "1" if self.tasima_dayanim_23_uygulandi_mi() else "0",
            ),)
        qt_nihai = self.entry_qt_nihai.get().strip() if hasattr(self, "entry_qt_nihai") else ""
        varsayim = (
            bool(self.tasima_varsayim_onayi.get())
            if tur != "kaya" and hasattr(self, "tasima_varsayim_onayi")
            else False
        )
        return tur, varsayim, girdiler, qt_nihai

    def tasima_raporu_guncel_mi(self):
        metin = self.txt_tasima_rapor.get("1.0", tk.END).strip() if hasattr(self, "txt_tasima_rapor") else ""
        return bool(metin) and self.tasima_rapor_imzasi == self.tasima_girdi_imzasi_olustur()

    def tasima_hesapla(self):
        self._tasima_sonuclarini_temizle(raporu_temizle=True)
        try:
            if getattr(self, "zemin_kaya_var", None) and self.zemin_kaya_var.get() == "kaya":
                degerler = self.kaya_tasima_girdilerini_oku()
                if not degerler:
                    return
                qk = degerler["qt"]
                qt = qk / degerler["Gk"]
            else:
                degerler = self.zemin_tasima_girdilerini_oku()
                if not degerler:
                    return
                _analizci, qk, qt = self._zemin_tasima_sonuclari(degerler)

            qt_rapor = tasima_qt_asagi_yuvarla(qt)
            if qt_rapor <= 0:
                raise ValueError("Hesaplanan q_t, iki ondalık rapor hassasiyetinde sıfırdan büyük olmalıdır.")
            if self.zemin_kaya_var.get() == "kaya":
                ks = qt_rapor * degerler["ks_carpani"]
            else:
                ks = qt_rapor * BOWLES_ZEMIN_KATSAYISI * degerler["RvGk"]
            self._tasima_sonuclarini_yaz(qk, qt_rapor, ks)
        except ValueError as e:
            self.tasima_giris_hatasi_goster([str(e)])
        except Exception as e:
            self.hata_kaydet("Taşıma gücü hesabı sırasında hata oluştu", e)
            messagebox.showerror("Beklenmeyen Hata", f"Hesaplama sırasında hata oluştu:\n{str(e)}")

    def tasima_metni_olustur(self):
        if hasattr(self, "txt_tasima_rapor"):
            self.txt_tasima_rapor.delete("1.0", tk.END)
        try:
            qt_val = self.nihai_qt_oku()
            if self.zemin_kaya_var.get() == "kaya":
                degerler = self.kaya_tasima_girdilerini_oku()
                if not degerler:
                    return
                qk_val = degerler["qt"]
                hesaplanan_qt = qk_val / degerler["Gk"]
                ks_val = qt_val * degerler["ks_carpani"]
                kaynak_aciklamasi = (
                    "Karakteristik kaya taşıma gücü q_k, sorumlu proje/rapor müellifince "
                    "belirlenerek programa girilmiştir. Bu değer; kullanılan saha, laboratuvar "
                    "ve kaya kütlesi değerlendirmeleriyle birlikte ayrıca doğrulanmalıdır."
                )
                metin = (
                    "Kaya Birimi Taşıma Gücü Kontrolü:\n\n"
                    f"{kaynak_aciklamasi}\n\n"
                    "Tasarım taşıma gücü q_t = q_k / G_k bağıntısıyla belirlenmiştir.\n\n"
                    f"q_t = {qk_val:.2f} / {degerler['Gk']:.2f} = {hesaplanan_qt:.2f} t/m²\n\n"
                    f"Raporda kullanılan muhafazakâr q_t = {qt_val:.2f} t/m²'dir.\n\n"
                    "Statik ve deprem etkilerini içeren tasarım temel taban basıncı bu "
                    "uygulamanın girdileri arasında bulunmadığından taşıma gücü yeterliliği "
                    "hakkında hüküm üretilmemiştir. Hesaplanan q_t değeri, statik proje "
                    "müellifince kendi tasarım etkileriyle ayrıca kontrol edilmelidir.\n\n"
                    f"Yatak katsayısı k_s = q_t × {degerler['ks_carpani']:.2f} bağıntısıyla "
                    f"k_s = {ks_val:.2f} t/m³ alınmıştır. "
                    "Bu ampirik çarpan proje müellifince doğrulanmalıdır."
                )
            else:
                degerler = self.zemin_tasima_girdilerini_oku()
                if not degerler:
                    return
                analizci, qk_val, hesaplanan_qt = self._zemin_tasima_sonuclari(degerler)
                if degerler.get("yass_var"):
                    yass_aciklamasi = (
                        f"YASS, zemin yüzeyinden {degerler['yass']:.2f} m derinlikte "
                        "dikkate alınmıştır."
                    )
                else:
                    yass_aciklamasi = (
                        "YASS gözlenmediği kabul edilmiş ve taşıma gücü hesabında "
                        "yeraltı suyu düzeltmesi uygulanmamıştır."
                    )
                dayanim_23_aciklamasi = ""
                if self.tasima_dayanim_23_uygulandi_mi():
                    dayanim_23_aciklamasi = (
                        "Taşıma gücü hesabında kullanılan efektif dayanım parametreleri, "
                        "zemin koşullarındaki belirsizlikler dikkate alınarak güvenli tarafta "
                        "kalınması amacıyla başlangıç değerlerinin 2/3’üne düşürülmüş; "
                        f"hesaplarda c′={degerler['c'] / KILONEWTON_PER_TON_FORCE:.2f} t/m² "
                        f"ve φ′={degerler['phi']:.2f}° değerleri esas alınmıştır.\n\n"
                    )
                ks_val = (
                    qt_val
                    * BOWLES_ZEMIN_KATSAYISI
                    * degerler["RvGk"]
                )
                metin = (
                    "Taşıma Gücü Analizi:\n\n"
                    "Temel taşıma gücünün karakteristik dayanımı q_k, TBDY 2018 Bölüm 16 esas alınarak "
                    "aşağıdaki bağıntıyla hesaplanmıştır.\n\n"
                    "q_k = c N_c s_c d_c i_c g_c b_c + q N_q s_q d_q i_q g_q b_q + 0.5 γ B′ N_γ s_γ d_γ i_γ g_γ b_γ \n\n"
                    f"{yass_aciklamasi}\n\n"
                    f"{dayanim_23_aciklamasi}"
                    "[TABLO_BURADA]\n"
                    "Buna göre temel taşıma gücü karakteristik dayanımı;\n\n"
                    f"q_k = {qk_val:.2f} t/m² olarak hesaplanmıştır.\n\n"
                    "Temel taşıma gücü tasarım dayanımı;\n\n"
                    "q_t = q_k / γ_Rv olarak tanımlanmıştır (TBDY Denk. 16.7).\n\n"
                    "Temel taşıma gücü dayanım katsayısı "
                    f"γ_Rv = {degerler['RvGk']:.2f} (TBDY Tablo 16.2) olmak üzere;\n\n"
                    f"q_t = {qk_val:.2f} / {degerler['RvGk']:.2f} = {hesaplanan_qt:.2f} t/m² "
                    "olarak hesaplanmıştır.\n\n"
                    "Statik ve deprem etkilerini içeren tasarım temel taban basıncı bu "
                    "uygulamanın girdileri arasında bulunmadığından taşıma gücü yeterliliği "
                    "hakkında hüküm üretilmemiştir. Hesaplanan q_t değeri, statik proje "
                    "müellifince kendi tasarım etkileriyle ayrıca kontrol edilmelidir.\n\n"
                    "Yatak katsayısı hesabında raporda kullanılan tasarım dayanımı "
                    f"q_t = {qt_val:.2f} t/m² ve güvenlik katsayısı "
                    f"G_k = {degerler['RvGk']:.2f} kabul edilerek Bowles (1988) yaklaşımına göre;\n\n"
                    f"k_s = {BOWLES_ZEMIN_KATSAYISI:.0f} × q_t × G_k\n\n"
                    f"k_s = {BOWLES_ZEMIN_KATSAYISI:.0f} × {qt_val:.2f} × "
                    f"{degerler['RvGk']:.2f} = {ks_val:.2f} t/m³ olarak hesaplanmıştır."
                )

            if qt_val > hesaplanan_qt + max(1e-9, abs(hesaplanan_qt) * 1e-12):
                raise ValueError(
                    f"Raporda kullanılacak q_t ({qt_val:.2f} t/m²), hesaplanan tasarım "
                    f"dayanımından ({hesaplanan_qt:.2f} t/m²) büyük olamaz."
                )
            self._tasima_sonuclarini_yaz(qk_val, qt_val, ks_val)
            metin += (
                "\n\nTemel altı yerdeğiştirmeleri/oturmalar, yatayda kayma, genel stabilite, "
                "tabakalanma-süreksizlikler ve depremde oluşabilecek dayanım/rijitlik kayıpları "
                "ilgili proje koşulları için ayrıca incelenmelidir."
            )
            self.txt_tasima_rapor.delete("1.0", tk.END)
            self.txt_tasima_rapor.insert(tk.END, metin)
            self.tasima_rapor_imzasi = self.tasima_girdi_imzasi_olustur()
        except ValueError as e:
            self.tasima_giris_hatasi_goster([str(e)])
        except Exception as e:
            self.hata_kaydet("Taşıma gücü rapor metni oluşturulamadı", e)
            messagebox.showerror("Hata", f"Metin oluşturulamadı. Lütfen Nihai değer kutularını sayısal olarak doldurun.\n{e}")
