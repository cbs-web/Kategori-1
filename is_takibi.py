import datetime
import json
import os
import sqlite3
import tkinter as tk
import uuid
from tkinter import filedialog, messagebox, ttk

from on_deger import IS_DURUMLARI, normalize_is_akisi, normalize_on_deger, normalize_tdth, on_deger_durumu
from proje_durumu_islemleri import IS_DURUM_RENKLERI


class IsTakibiDeposu:
    def __init__(self, db_yolu):
        self.db_yolu = os.path.abspath(db_yolu)
        os.makedirs(os.path.dirname(self.db_yolu), exist_ok=True)
        self._hazirla()

    def _baglan(self):
        con = sqlite3.connect(self.db_yolu)
        con.row_factory = sqlite3.Row
        return con

    def _hazirla(self):
        with self._baglan() as con:
            con.execute(
                """
                CREATE TABLE IF NOT EXISTS projeler (
                    proje_id TEXT PRIMARY KEY,
                    dosya_yolu TEXT NOT NULL,
                    proje_adi TEXT NOT NULL DEFAULT '',
                    ilce TEXT NOT NULL DEFAULT '',
                    koy TEXT NOT NULL DEFAULT '',
                    ada TEXT NOT NULL DEFAULT '',
                    parsel TEXT NOT NULL DEFAULT '',
                    on_qt TEXT NOT NULL DEFAULT '',
                    on_ks TEXT NOT NULL DEFAULT '',
                    zemin_sinifi TEXT NOT NULL DEFAULT '',
                    on_deger_durumu TEXT NOT NULL DEFAULT 'verilmedi',
                    tdth_durumu TEXT NOT NULL DEFAULT 'eksik',
                    is_durumu TEXT NOT NULL DEFAULT 'yeni',
                    revizyon_no INTEGER NOT NULL DEFAULT 1,
                    son_guncelleme TEXT NOT NULL DEFAULT ''
                )
                """
            )
            mevcut_kolonlar = {row[1] for row in con.execute("PRAGMA table_info(projeler)")}
            if "on_deger_durumu" not in mevcut_kolonlar:
                con.execute(
                    "ALTER TABLE projeler ADD COLUMN on_deger_durumu TEXT NOT NULL DEFAULT 'verilmedi'"
                )
            con.execute("CREATE INDEX IF NOT EXISTS idx_projeler_durum ON projeler(is_durumu)")
            con.execute("CREATE INDEX IF NOT EXISTS idx_projeler_konum ON projeler(ilce, koy, ada, parsel)")

    def kaydet(self, proje_yolu, veriler):
        if not proje_yolu or not isinstance(veriler, dict):
            return
        akis = normalize_is_akisi(
            veriler.get("_IS_AKISI_"), eski_proje="_IS_AKISI_" not in veriler
        )
        if "_IS_AKISI_" not in veriler:
            akis["proje_id"] = str(uuid.uuid5(uuid.NAMESPACE_URL, os.path.normcase(os.path.abspath(proje_yolu))))
        on_deger = normalize_on_deger(veriler.get("_ON_DEGER_"))
        tdth = normalize_tdth(veriler.get("_TDTH_"))
        guncel = on_deger.get("guncel", {})
        satir = (
            akis["proje_id"], os.path.abspath(proje_yolu), str(veriler.get("PROJE_ADI", "")),
            str(veriler.get("ILCE", "")), str(veriler.get("KOY", "")), str(veriler.get("ADA", "")),
            str(veriler.get("PARSEL", "")), str(guncel.get("qt", "")), str(guncel.get("ks", "")),
            str(guncel.get("zemin_sinifi", "")), on_deger_durumu(on_deger),
            tdth.get("durum", "eksik"), akis.get("durum", "yeni"),
            int(akis.get("revizyon_no", 1)), datetime.datetime.now().astimezone().isoformat(timespec="seconds"),
        )
        with self._baglan() as con:
            con.execute(
                "DELETE FROM projeler WHERE dosya_yolu = ? AND proje_id <> ?",
                (os.path.abspath(proje_yolu), akis["proje_id"]),
            )
            con.execute(
                """
                INSERT INTO projeler (
                    proje_id, dosya_yolu, proje_adi, ilce, koy, ada, parsel,
                    on_qt, on_ks, zemin_sinifi, on_deger_durumu, tdth_durumu, is_durumu,
                    revizyon_no, son_guncelleme
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(proje_id) DO UPDATE SET
                    dosya_yolu=excluded.dosya_yolu,
                    proje_adi=excluded.proje_adi,
                    ilce=excluded.ilce,
                    koy=excluded.koy,
                    ada=excluded.ada,
                    parsel=excluded.parsel,
                    on_qt=excluded.on_qt,
                    on_ks=excluded.on_ks,
                    zemin_sinifi=excluded.zemin_sinifi,
                    on_deger_durumu=excluded.on_deger_durumu,
                    tdth_durumu=excluded.tdth_durumu,
                    is_durumu=excluded.is_durumu,
                    revizyon_no=excluded.revizyon_no,
                    son_guncelleme=excluded.son_guncelleme
                """,
                satir,
            )

    def listele(self, arama="", durum=""):
        kosullar = []
        params = []
        if arama.strip():
            like = f"%{arama.strip()}%"
            kosullar.append(
                "(proje_adi LIKE ? OR ilce LIKE ? OR koy LIKE ? OR ada LIKE ? OR parsel LIKE ? OR dosya_yolu LIKE ?)"
            )
            params.extend([like] * 6)
        if durum:
            kosullar.append("is_durumu = ?")
            params.append(durum)
        where = " WHERE " + " AND ".join(kosullar) if kosullar else ""
        with self._baglan() as con:
            return [dict(row) for row in con.execute(
                "SELECT * FROM projeler" + where + " ORDER BY son_guncelleme DESC, proje_adi", params
            )]


class IsTakibiIslemleri:
    def __init__(self, app):
        object.__setattr__(self, "app", app)

    def __getattr__(self, name):
        return getattr(self.app, name)

    def __setattr__(self, name, value):
        if name == "app":
            object.__setattr__(self, name, value)
        else:
            setattr(self.app, name, value)

    def is_takibi_deposu(self):
        return IsTakibiDeposu(os.path.join(self.kullanici_veri_klasoru_bul(), "is_takibi.db"))

    def is_takibi_kaydi_guncelle(self, proje_yolu=None, veriler=None):
        yol = proje_yolu or getattr(self, "guncel_dosya_yolu", "")
        if not yol:
            return
        try:
            self.is_takibi_deposu().kaydet(yol, veriler if isinstance(veriler, dict) else self.verileri_topla())
        except Exception as exc:
            if hasattr(self, "hata_kaydet"):
                self.hata_kaydet("İş takibi indeksi güncellenemedi", exc)

    def is_takibi_penceresi(self):
        pencere = self.animasyonlu_pencere()
        pencere.title("İş Takibi")
        pencere.geometry("1250x620")
        pencere.transient(self.root)

        ust = ttk.Frame(pencere, padding=12)
        ust.pack(fill="x")
        arama_var = tk.StringVar()
        durum_var = tk.StringVar(value="Tümü")
        ttk.Label(ust, text="Ara:").pack(side="left")
        arama = ttk.Entry(ust, textvariable=arama_var, width=35)
        arama.pack(side="left", padx=(6, 14))
        ttk.Label(ust, text="Aşama:").pack(side="left")
        durum = ttk.Combobox(ust, textvariable=durum_var, state="readonly", width=24, values=("Tümü",) + tuple(IS_DURUMLARI.values()))
        durum.pack(side="left", padx=(6, 14))

        kolonlar = ("proje", "konum", "ada_parsel", "on_deger", "qt", "ks", "zemin", "tdth", "durum", "rev", "tarih", "dosya")
        tree = ttk.Treeview(pencere, columns=kolonlar, show="headings", style="Treeview")
        basliklar = {
            "proje": "Proje Sahibi", "konum": "İlçe / Köy", "ada_parsel": "Ada / Parsel",
            "on_deger": "Ön Değer", "qt": "Ön qₜ", "ks": "Ön kₛ", "zemin": "Zemin", "tdth": "TDTH",
            "durum": "İş Aşaması", "rev": "Rev.", "tarih": "Son Değişiklik", "dosya": "Proje Dosyası",
        }
        genislik = {"proje": 170, "konum": 165, "ada_parsel": 95, "on_deger": 90, "qt": 65, "ks": 65, "zemin": 60, "tdth": 85, "durum": 145, "rev": 50, "tarih": 145, "dosya": 330}
        for kod in kolonlar:
            tree.heading(kod, text=basliklar[kod])
            tree.column(kod, width=genislik[kod], anchor="center" if kod not in {"proje", "konum", "dosya"} else "w")
        for kod, renk in IS_DURUM_RENKLERI.items():
            tree.tag_configure(kod, background=renk, foreground="white")
        tree.tag_configure("eksik", background="#fee2e2", foreground="#991b1b")
        scroll = ttk.Scrollbar(pencere, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y", padx=(0, 8), pady=(0, 8))
        tree.pack(fill="both", expand=True, padx=(12, 0), pady=(0, 8))

        kayitlar = {}
        ters = {v: k for k, v in IS_DURUMLARI.items()}

        def yenile(*_args):
            for item in tree.get_children():
                tree.delete(item)
            filtre = "" if durum_var.get() == "Tümü" else ters.get(durum_var.get(), "")
            kayitlar.clear()
            for satir in self.is_takibi_deposu().listele(arama_var.get(), filtre):
                yol_var = os.path.isfile(satir["dosya_yolu"])
                tag = satir["is_durumu"] if yol_var else "eksik"
                iid = tree.insert("", "end", values=(
                    satir["proje_adi"], f"{satir['ilce']} / {satir['koy']}", f"{satir['ada']} / {satir['parsel']}",
                    "Verildi" if satir["on_deger_durumu"] == "verildi" else "Verilmedi",
                    satir["on_qt"], satir["on_ks"], satir["zemin_sinifi"], satir["tdth_durumu"],
                    IS_DURUMLARI.get(satir["is_durumu"], satir["is_durumu"]), satir["revizyon_no"],
                    satir["son_guncelleme"].replace("T", " ")[:19], satir["dosya_yolu"] + ("" if yol_var else " [BULUNAMADI]"),
                ), tags=(tag,))
                kayitlar[iid] = satir

        def secili_ac():
            secim = tree.selection()
            if not secim:
                messagebox.showwarning("İş Takibi", "Önce bir proje seçin.")
                return
            satir = kayitlar.get(secim[0], {})
            yol = satir.get("dosya_yolu", "")
            if not os.path.isfile(yol):
                messagebox.showwarning("İş Takibi", "Proje dosyası bulunamadı.")
                return
            if self.proje_dosyasini_ac(yol):
                pencere.destroy()

        def klasor_tara():
            kok = filedialog.askdirectory(title="K-1 proje JSON dosyalarının bulunduğu klasörü seçin", parent=pencere)
            if not kok:
                return
            bulunan = 0
            for dizin, _klasorler, dosyalar in os.walk(kok):
                for ad in dosyalar:
                    if not ad.lower().endswith(".json"):
                        continue
                    yol = os.path.join(dizin, ad)
                    try:
                        with open(yol, "r", encoding="utf-8-sig") as f:
                            veri = json.load(f)
                        if not isinstance(veri, dict) or "PROJE_ADI" not in veri:
                            continue
                        self.is_takibi_deposu().kaydet(yol, veri)
                        bulunan += 1
                    except Exception:
                        continue
            yenile()
            messagebox.showinfo("İş Takibi", f"{bulunan} proje kaydı tarandı.")

        ttk.Button(ust, text="Yenile", command=yenile, bootstyle="secondary").pack(side="left", padx=(0, 6))
        ttk.Button(ust, text="Projeyi Aç", command=secili_ac, bootstyle="primary").pack(side="left", padx=(0, 6))
        ttk.Button(ust, text="Klasör Tara", command=klasor_tara, bootstyle="info").pack(side="left")
        arama.bind("<KeyRelease>", yenile)
        durum.bind("<<ComboboxSelected>>", yenile)
        tree.bind("<Double-1>", lambda _event: secili_ac())
        yenile()
