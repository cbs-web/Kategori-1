import os
import tkinter as tk
from tkinter import ttk

from PIL import Image, ImageTk

from arayuz_yardimcilari import ARAYUZ_RENKLERI
from on_deger import IS_DURUMLARI, normalize_is_akisi, on_deger_durumu


IS_DURUM_RENKLERI = {
    "belirlenmedi": "#64748b",
    "yeni": "#6b7280",
    "on_deger_verildi": "#d97706",
    "yazim_asamasinda": "#2563eb",
    "duzeltme_asamasinda": "#7e22ce",
    "bitti": "#15803d",
}

HARITALI_OZET_DURUMLARI = frozenset({"yazim_asamasinda", "duzeltme_asamasinda", "bitti"})


def turkce_buyuk_harf(deger):
    return str(deger or "").translate(str.maketrans({"i": "İ", "ı": "I"})).upper()


def proje_kimligi_hazirla(veriler):
    if not isinstance(veriler, dict):
        return "Yeni Proje"
    ilce = str(veriler.get("ILCE") or "").strip()
    koy = str(veriler.get("KOY") or "").strip()
    ada = str(veriler.get("ADA") or "").strip()
    parsel = str(veriler.get("PARSEL") or "").strip()
    parcalar = [x for x in (ilce, koy) if x]
    if ada or parsel:
        parcalar.append(f"{ada or '-'}-{parsel or '-'}")
    return " / ".join(parcalar) if parcalar else "Yeni Proje"


def parsel_haritasi_durumu_hazirla(
    durum_kodu,
    harita_yolu="",
    kml_noktalari=None,
    kayitli_hash="",
    guncel_hash="",
    kayitli_ada="",
    kayitli_parsel="",
    guncel_ada="",
    guncel_parsel="",
):
    if durum_kodu not in HARITALI_OZET_DURUMLARI:
        return {"kod": "asama_disinda", "goster": False, "mesaj": ""}

    if not kml_noktalari:
        return {
            "kod": "kml_yok",
            "goster": False,
            "mesaj": "Parsel KML'si bulunmuyor. Haritalar sekmesinden KML yükleyin veya TKGM'den alın.",
        }

    yol = os.path.abspath(str(harita_yolu or "")) if harita_yolu else ""
    if not yol or not os.path.isfile(yol):
        return {
            "kod": "harita_yok",
            "goster": False,
            "mesaj": "KML hazır; parselin uydu görüntüsü henüz oluşturulmadı.",
        }

    if kayitli_hash and guncel_hash and str(kayitli_hash) != str(guncel_hash):
        return {
            "kod": "geometri_degisti",
            "goster": False,
            "mesaj": "KML geometrisi değişmiş. Parsel haritasını yeniden hazırlayın.",
        }

    kayitli_kunye = (str(kayitli_ada or "").strip(), str(kayitli_parsel or "").strip())
    guncel_kunye = (str(guncel_ada or "").strip(), str(guncel_parsel or "").strip())
    if kayitli_kunye != guncel_kunye:
        return {
            "kod": "ada_parsel_degisti",
            "goster": False,
            "mesaj": "Ada/parsel bilgisi harita hazırlandıktan sonra değişmiş. Parsel haritasını yeniden hazırlayın.",
        }

    return {"kod": "hazir", "goster": True, "mesaj": "", "harita_yolu": yol}


def proje_durum_ozeti_hazirla(veriler, salt_okunur=False, kaydedilmedi=False):
    if not isinstance(veriler, dict):
        veriler = {}
    akis = normalize_is_akisi(veriler.get("_IS_AKISI_"))
    durum = akis.get("durum", "yeni")
    on_durum = on_deger_durumu(veriler.get("_ON_DEGER_"))
    kimlik = proje_kimligi_hazirla(veriler)
    asama = IS_DURUMLARI.get(durum, durum)
    mod = "İzleme" if salt_okunur else "Düzenleme"
    proje_adi = str(veriler.get("PROJE_ADI") or "").strip() or "Proje adı girilmedi"
    ilce = str(veriler.get("ILCE") or "").strip()
    koy = str(veriler.get("KOY") or "").strip()
    mevkii = str(veriler.get("MEVKII") or "").strip()
    ada = str(veriler.get("ADA") or "").strip()
    parsel = str(veriler.get("PARSEL") or "").strip()
    konum = " / ".join(parca for parca in (ilce, koy, mevkii) if parca) or "Konum bilgisi girilmedi"
    ada_parsel = f"ADA {ada or '-'} — PARSEL {parsel or '-'}"
    baslik_parcalari = ["K-1", turkce_buyuk_harf(asama), turkce_buyuk_harf(kimlik)]
    if kaydedilmedi:
        baslik_parcalari.append("KAYDEDİLMEDİ")
    return {
        "kimlik": kimlik,
        "proje_adi": proje_adi,
        "konum": konum,
        "ada": ada,
        "parsel": parsel,
        "ada_parsel": ada_parsel,
        "durum_kodu": durum,
        "asama": asama,
        "haritali_ozet": durum in HARITALI_OZET_DURUMLARI,
        "on_deger_kodu": on_durum,
        "on_deger": "Ön Değer: Verildi" if on_durum == "verildi" else "Ön Değer: Verilmedi",
        "revizyon_no": akis.get("revizyon_no", 1),
        "revizyon": f"Revizyon {akis.get('revizyon_no', 1)}",
        "mod": mod,
        "salt_okunur": bool(salt_okunur),
        "kaydedilmedi": bool(kaydedilmedi),
        "pencere_basligi": " — ".join(baslik_parcalari),
    }


class ProjeDurumuIslemleri:
    def __init__(self, app):
        object.__setattr__(self, "app", app)

    def __getattr__(self, name):
        return getattr(self.app, name)

    def __setattr__(self, name, value):
        if name == "app":
            object.__setattr__(self, name, value)
        else:
            setattr(self.app, name, value)

    def _serit_stillerini_hazirla(self):
        style = ttk.Style()
        renk = ARAYUZ_RENKLERI
        style.configure("ProjeOzet.TFrame", background=renk["zemin"])
        style.configure(
            "ProjeKimligi.TLabel",
            background=renk["zemin"],
            foreground=renk["metin_ikincil"],
            font=("Segoe UI", 12),
        )
        for kod in IS_DURUM_RENKLERI:
            style.configure(
                f"ProjeAsama.{kod}.TLabel",
                background=renk["yuzey"],
                foreground=renk["vurgu"],
                font=("Segoe UI", 20, "bold"),
                padding=(16, 12),
            )
        rozet_yazi = ("Segoe UI", 10, "bold")
        rozet_bosluk = (11, 7)
        style.configure("ProjeOnDegerVar.TLabel", background=renk["yuzey_ikincil"], foreground=renk["metin"], font=rozet_yazi, padding=rozet_bosluk)
        style.configure("ProjeOnDegerYok.TLabel", background=renk["yuzey_ikincil"], foreground=renk["metin_ikincil"], font=rozet_yazi, padding=rozet_bosluk)
        style.configure("ProjeRozet.TLabel", background=renk["yuzey_ikincil"], foreground=renk["metin_ikincil"], font=rozet_yazi, padding=rozet_bosluk)
        style.configure("ProjeAdi.TLabel", background=renk["zemin"], foreground=renk["metin"], font=("Segoe UI", 18, "bold"))
        style.configure("ProjeKonum.TLabel", background=renk["zemin"], foreground=renk["metin_ikincil"], font=("Segoe UI", 11))
        style.configure("ProjeAdaParsel.TLabel", background=renk["zemin"], foreground=renk["vurgu"], font=("Segoe UI", 16, "bold"))
        style.configure("ProjeHaritaDurum.TLabel", background=renk["zemin"], foreground=renk["metin_ikincil"], font=("Segoe UI", 10, "bold"))

    def proje_ozet_sekmesi_olustur(self):
        self._serit_stillerini_hazirla()
        self.proje_kimligi_var = tk.StringVar(value="Yeni Proje")
        self.proje_asama_var = tk.StringVar(value=IS_DURUMLARI["yeni"])
        self.proje_on_deger_ozet_var = tk.StringVar(value="Ön Değer: Verilmedi")
        self.proje_revizyon_var = tk.StringVar(value="Revizyon 1")
        self.proje_adi_ozet_var = tk.StringVar(value="")
        self.proje_konum_ozet_var = tk.StringVar(value="")
        self.proje_ada_parsel_ozet_var = tk.StringVar(value="")
        self.proje_harita_mesaj_var = tk.StringVar(value="")

        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="0. Özet")
        page = ttk.Frame(frame, padding=22, style="ProjeOzet.TFrame")
        page.pack(fill="both", expand=True)

        self.lbl_proje_asama = ttk.Label(
            page,
            textvariable=self.proje_asama_var,
            style="ProjeAsama.yeni.TLabel",
            anchor="center",
        )
        self.lbl_proje_asama.pack(fill="x", pady=(0, 4))
        self.lbl_proje_kimligi = ttk.Label(
            page,
            textvariable=self.proje_kimligi_var,
            style="ProjeKimligi.TLabel",
            anchor="center",
        )
        self.lbl_proje_kimligi.pack(fill="x", pady=(16, 18))

        self.proje_haritali_ozet_frame = ttk.Frame(page, style="ProjeOzet.TFrame")
        ttk.Label(
            self.proje_haritali_ozet_frame,
            textvariable=self.proje_adi_ozet_var,
            style="ProjeAdi.TLabel",
            anchor="center",
        ).pack(fill="x", pady=(12, 4))
        ttk.Label(
            self.proje_haritali_ozet_frame,
            textvariable=self.proje_konum_ozet_var,
            style="ProjeKonum.TLabel",
            anchor="center",
        ).pack(fill="x", pady=(0, 4))
        ttk.Label(
            self.proje_haritali_ozet_frame,
            textvariable=self.proje_ada_parsel_ozet_var,
            style="ProjeAdaParsel.TLabel",
            anchor="center",
        ).pack(fill="x", pady=(0, 10))

        self.proje_harita_cercevesi = ttk.LabelFrame(
            self.proje_haritali_ozet_frame,
            text="Parsel Uydu Görüntüsü",
            padding=(10, 8),
        )
        self.proje_harita_cercevesi.pack(fill="both", expand=True)
        self.lbl_proje_harita = tk.Label(
            self.proje_harita_cercevesi,
            text="",
            background="#e2e8f0",
            foreground="#334155",
            anchor="center",
        )
        self.lbl_proje_harita.pack(fill="both", expand=True)
        self.lbl_proje_harita_durum = ttk.Label(
            self.proje_harita_cercevesi,
            textvariable=self.proje_harita_mesaj_var,
            style="ProjeHaritaDurum.TLabel",
            anchor="center",
            justify="center",
            wraplength=780,
        )
        self.btn_proje_harita_islem = ttk.Button(
            self.proje_harita_cercevesi,
            text="Parsel Haritası Hazırla",
            command=self._ozetten_parsel_haritasi_hazirla,
            style="Secondary.TButton",
        )

        rozetler = ttk.Frame(page, style="ProjeOzet.TFrame")
        rozetler.pack(anchor="center", pady=(0, 20))
        self.lbl_proje_on_deger = ttk.Label(
            rozetler, textvariable=self.proje_on_deger_ozet_var, style="ProjeOnDegerYok.TLabel"
        )
        self.lbl_proje_on_deger.pack(side="left", padx=5)
        self.lbl_proje_revizyon = ttk.Label(
            rozetler, textvariable=self.proje_revizyon_var, style="ProjeRozet.TLabel"
        )

        self.btn_asama_degistir = ttk.Button(
            page,
            text="Aşamayı Değiştir",
            command=self.is_asamasini_belirle,
            style="Primary.TButton",
            width=22,
        )
        self.btn_asama_degistir.pack(anchor="center")
        self.proje_ozet_sekmesi = frame
        self._proje_ozet_harita_cache_key = None
        self._proje_ozet_harita_photo = None
        self.proje_durum_seridi_guncelle(kaydedilmedi=False)
        self._proje_durumu_eventlerini_bagla()
        return frame

    def proje_durum_seridi_olustur(self):
        """Eski çağrı adı; durum görünümü artık 0. Özet sekmesindedir."""
        return self.proje_ozet_sekmesi_olustur()

    def _mevcut_proje_ozet_verisi(self):
        veri = {
            "_IS_AKISI_": getattr(self, "is_akisi_verisi", {}),
            "_ON_DEGER_": self.on_deger_verisini_topla() if hasattr(self, "on_deger_verisini_topla") else getattr(self, "on_deger_verisi", {}),
        }
        for kod in ("PROJE_ADI", "ILCE", "KOY", "MEVKII", "ADA", "PARSEL"):
            entry = getattr(self, "veri_alanlari", {}).get(kod)
            veri[kod] = entry.get().strip() if entry is not None else ""
        return veri

    def _haritalar_sekmesine_git(self):
        for sekme_id in self.notebook.tabs():
            if "Haritalar" in str(self.notebook.tab(sekme_id, "text")):
                self.notebook.select(sekme_id)
                return

    def _ozetten_parsel_haritasi_hazirla(self):
        if self.parsel_haritasi_hazirla():
            self.proje_durum_seridi_guncelle()

    def _proje_ozet_harita_durumu(self, ozet):
        try:
            guncel_hash = self.cizim_uretici().parsel_geometri_hashi()
        except Exception:
            guncel_hash = ""
        return parsel_haritasi_durumu_hazirla(
            ozet["durum_kodu"],
            harita_yolu=getattr(self, "img_parsel_haritasi", "") or "",
            kml_noktalari=getattr(self, "yuklu_kml_points", []) or [],
            kayitli_hash=getattr(self, "parsel_haritasi_geometri_hash", "") or "",
            guncel_hash=guncel_hash,
            kayitli_ada=getattr(self, "parsel_haritasi_ada", "") or "",
            kayitli_parsel=getattr(self, "parsel_haritasi_parsel", "") or "",
            guncel_ada=ozet["ada"],
            guncel_parsel=ozet["parsel"],
        )

    def _proje_ozet_harita_gorselini_temizle(self):
        self.lbl_proje_harita.configure(image="", text="")
        self.lbl_proje_harita.image = None

    def _proje_ozet_harita_gorselini_yukle(self, harita_yolu):
        try:
            istatistik = os.stat(harita_yolu)
            cache_key = (os.path.normcase(os.path.abspath(harita_yolu)), istatistik.st_mtime_ns, istatistik.st_size)
            if cache_key != getattr(self, "_proje_ozet_harita_cache_key", None):
                with Image.open(harita_yolu) as kaynak:
                    kaynak.load()
                    gorsel = kaynak.convert("RGB")
                gorsel.thumbnail((860, 320), Image.Resampling.LANCZOS)
                self._proje_ozet_harita_photo = ImageTk.PhotoImage(gorsel, master=self.root)
                self._proje_ozet_harita_cache_key = cache_key
            self.lbl_proje_harita.configure(image=self._proje_ozet_harita_photo, text="")
            self.lbl_proje_harita.image = self._proje_ozet_harita_photo
            return True
        except (OSError, ValueError, tk.TclError):
            self._proje_ozet_harita_cache_key = None
            self._proje_ozet_harita_photo = None
            self._proje_ozet_harita_gorselini_temizle()
            return False

    def _proje_haritali_ozeti_guncelle(self, ozet):
        if not ozet["haritali_ozet"]:
            self.proje_haritali_ozet_frame.pack_forget()
            if not self.lbl_proje_kimligi.winfo_manager():
                self.lbl_proje_kimligi.pack(fill="x", pady=(16, 18), after=self.lbl_proje_asama)
            self._proje_ozet_harita_gorselini_temizle()
            return

        self.lbl_proje_kimligi.pack_forget()
        if not self.proje_haritali_ozet_frame.winfo_manager():
            self.proje_haritali_ozet_frame.pack(
                fill="both", expand=True, pady=(0, 14), after=self.lbl_proje_asama
            )
        self.proje_adi_ozet_var.set(ozet["proje_adi"])
        self.proje_konum_ozet_var.set(ozet["konum"])
        self.proje_ada_parsel_ozet_var.set(ozet["ada_parsel"])

        harita_durumu = self._proje_ozet_harita_durumu(ozet)
        if harita_durumu["goster"] and self._proje_ozet_harita_gorselini_yukle(harita_durumu["harita_yolu"]):
            self.lbl_proje_harita_durum.pack_forget()
            self.btn_proje_harita_islem.pack_forget()
            return

        self._proje_ozet_harita_gorselini_temizle()
        if harita_durumu["goster"]:
            harita_durumu = {
                "kod": "harita_okunamadi",
                "mesaj": "Kayıtlı parsel haritası okunamadı. Görüntüyü yeniden hazırlayın.",
            }
        self.proje_harita_mesaj_var.set(harita_durumu["mesaj"])
        if not self.lbl_proje_harita_durum.winfo_manager():
            self.lbl_proje_harita_durum.pack(fill="x", padx=12, pady=(12, 8))

        self.btn_proje_harita_islem.pack_forget()
        if ozet["salt_okunur"]:
            return
        if harita_durumu["kod"] == "kml_yok":
            self.btn_proje_harita_islem.configure(
                text="Haritalar Sekmesine Git",
                command=self._haritalar_sekmesine_git,
                style="Secondary.TButton",
            )
        else:
            self.btn_proje_harita_islem.configure(
                text="Parsel Haritası Hazırla",
                command=self._ozetten_parsel_haritasi_hazirla,
                style="Secondary.TButton",
            )
        self.btn_proje_harita_islem.pack(pady=(0, 10))

    def proje_durum_seridi_guncelle(self, kaydedilmedi=None):
        if kaydedilmedi is None:
            kaydedilmedi = bool(getattr(self, "_proje_kirli", False))
        ozet = proje_durum_ozeti_hazirla(
            self._mevcut_proje_ozet_verisi(),
            salt_okunur=getattr(self, "proje_salt_okunur", False),
            kaydedilmedi=kaydedilmedi,
        )
        if hasattr(self, "proje_kimligi_var"):
            self.proje_kimligi_var.set(ozet["kimlik"])
            self.proje_asama_var.set(turkce_buyuk_harf(ozet["asama"]))
            self.proje_on_deger_ozet_var.set(ozet["on_deger"])
            self.proje_revizyon_var.set(ozet["revizyon"])
            self.lbl_proje_asama.configure(style=f"ProjeAsama.{ozet['durum_kodu']}.TLabel")
            self._proje_haritali_ozeti_guncelle(ozet)
            self.lbl_proje_on_deger.configure(
                style="ProjeOnDegerVar.TLabel" if ozet["on_deger_kodu"] == "verildi" else "ProjeOnDegerYok.TLabel"
            )
            if ozet["revizyon_no"] > 1:
                if not self.lbl_proje_revizyon.winfo_manager():
                    self.lbl_proje_revizyon.pack(side="left", padx=5, after=self.lbl_proje_on_deger)
            else:
                self.lbl_proje_revizyon.pack_forget()
            if hasattr(self, "btn_asama_degistir") and not ozet["salt_okunur"]:
                self.btn_asama_degistir.configure(
                    state="disabled" if ozet["durum_kodu"] == "bitti" else "normal"
                )
        try:
            self.root.title(ozet["pencere_basligi"])
        except tk.TclError:
            pass

    def proje_durumu_yenilemeyi_planla(self, _event=None):
        olay_tipi = str(getattr(_event, "type", "")) if _event is not None else ""
        widget_sinifi = ""
        if _event is not None and getattr(_event, "widget", None) is not None:
            try:
                widget_sinifi = str(_event.widget.winfo_class())
            except (AttributeError, tk.TclError):
                pass
        sekme_tiklamasi = olay_tipi == "5" and widget_sinifi in {"TNotebook", "Notebook"}
        if olay_tipi not in {"", "35"} and not sekme_tiklamasi:
            self._proje_kirli = True
        onceki = getattr(self, "_proje_durumu_after_id", None)
        if onceki:
            try:
                self.root.after_cancel(onceki)
            except tk.TclError:
                pass
        try:
            self._proje_durumu_after_id = self.root.after(160, self._planlanan_proje_durumu_yenile)
        except tk.TclError:
            self._proje_durumu_after_id = None

    def _planlanan_proje_durumu_yenile(self):
        self._proje_durumu_after_id = None
        self.proje_durum_seridi_guncelle(kaydedilmedi=bool(getattr(self, "_proje_kirli", False)))

    def _proje_durumu_eventlerini_bagla(self):
        for olay in ("<KeyRelease>", "<ButtonRelease-1>", "<<ComboboxSelected>>", "<<NotebookTabChanged>>"):
            self.root.bind_all(olay, self.proje_durumu_yenilemeyi_planla, add="+")
