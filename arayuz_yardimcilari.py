import os
import time
import tkinter as tk
from tkinter import ttk, messagebox

import ttkbootstrap as tb
from PIL import Image, ImageDraw, ImageTk


ARAYUZ_RENKLERI = {
    "zemin": "#f5f7fa",
    "yuzey": "#ffffff",
    "yuzey_ikincil": "#eef2f6",
    "cizgi": "#d7dee7",
    "metin": "#172033",
    "metin_ikincil": "#526176",
    # Uygulamada zaten kullanılan ana mavi; yeni bir marka rengi icat edilmez.
    "vurgu": "#2563eb",
    "vurgu_koyu": "#1d4ed8",
    "tehlike": "#9f2d2d",
}


def yumusak_cikis_degeri(oran):
    """0..1 aralığında, ani bitişi olmayan ease-out cubic eğrisi."""
    oran = max(0.0, min(1.0, float(oran)))
    return 1.0 - ((1.0 - oran) ** 3)


class AnimasyonluToplevel(tk.Toplevel):
    """Windows yardımcı pencereleri için kısa, iptal edilebilir giriş/çıkış hareketi."""

    def __init__(self, master=None, *, animasyon_aktif=True, **kwargs):
        super().__init__(master, **kwargs)
        self._animasyon_aktif = bool(animasyon_aktif)
        self._animasyon_after_id = None
        self._animasyon_kapaniyor = False
        self._animasyon_hedef_y = None
        if self._animasyon_aktif:
            try:
                self.wm_attributes("-alpha", 0.0)
            except tk.TclError:
                self._animasyon_aktif = False
        self.after_idle(self._giris_animasyonunu_baslat)

    def _after_iptal(self):
        if self._animasyon_after_id is not None:
            try:
                self.after_cancel(self._animasyon_after_id)
            except tk.TclError:
                pass
            self._animasyon_after_id = None

    def _giris_animasyonunu_baslat(self):
        if not self._animasyon_aktif or self._animasyon_kapaniyor:
            return
        try:
            self.update_idletasks()
            self._animasyon_hedef_y = self.winfo_y()
            self.geometry(f"+{self.winfo_x()}+{self._animasyon_hedef_y + 8}")
        except tk.TclError:
            return
        self._animasyon_karesi(time.perf_counter(), 0.17, giris=True)

    def _animasyon_karesi(self, baslangic, sure, *, giris):
        try:
            oran = min(1.0, (time.perf_counter() - baslangic) / sure)
            yumusak = yumusak_cikis_degeri(oran)
            alpha = yumusak if giris else 1.0 - yumusak
            self.wm_attributes("-alpha", alpha)
            if giris and self._animasyon_hedef_y is not None:
                y = round(self._animasyon_hedef_y + (8 * (1.0 - yumusak)))
                self.geometry(f"+{self.winfo_x()}+{y}")
            if oran < 1.0:
                self._animasyon_after_id = self.after(
                    16, lambda: self._animasyon_karesi(baslangic, sure, giris=giris)
                )
            elif giris:
                self._animasyon_after_id = None
                self.wm_attributes("-alpha", 1.0)
            else:
                super().destroy()
        except tk.TclError:
            if not giris:
                try:
                    super().destroy()
                except tk.TclError:
                    pass

    def destroy(self):
        if not self._animasyon_aktif:
            return super().destroy()
        if self._animasyon_kapaniyor:
            return None
        try:
            if not self.winfo_exists():
                return None
        except tk.TclError:
            return None
        self._animasyon_kapaniyor = True
        self._after_iptal()
        try:
            if self.grab_current() == self:
                self.grab_release()
        except tk.TclError:
            pass
        self._animasyon_karesi(time.perf_counter(), 0.11, giris=False)
        return None


class ArayuzYardimcilari:
    def __init__(self, app):
        object.__setattr__(self, "app", app)

    def __getattr__(self, name):
        return getattr(self.app, name)

    def __setattr__(self, name, value):
        if name == "app":
            object.__setattr__(self, name, value)
        else:
            setattr(self.app, name, value)

    def animasyonlar_aktif_mi(self):
        ortam = str(os.environ.get("K1_REDUCED_MOTION", "")).strip().lower()
        test_ortami = bool(os.environ.get("PYTEST_CURRENT_TEST"))
        if ortam in {"1", "true", "evet", "yes", "on"} or test_ortami:
            return False
        degisken = getattr(self, "animasyonlar_aktif_var", None)
        if degisken is not None:
            try:
                return bool(degisken.get())
            except (AttributeError, tk.TclError):
                pass
        ayar = getattr(self, "animasyonlar_aktif", None)
        return True if ayar is None else bool(ayar)

    def animasyonlu_pencere(self, parent=None, **kwargs):
        return AnimasyonluToplevel(
            parent or self.root,
            animasyon_aktif=self.animasyonlar_aktif_mi(),
            **kwargs,
        )

    def genel_stilleri_hazirla(self):
        self.root.option_add("*Font", "{Segoe UI} 10")
        self.root.option_add("*TCombobox*Listbox.font", "{Segoe UI} 10")

        style = tb.Style()
        renk = ARAYUZ_RENKLERI
        self.root.configure(background=renk["zemin"])
        style.configure(".", font=("Segoe UI", 10), foreground=renk["metin"])
        style.configure("TFrame", background=renk["zemin"])
        style.configure("TLabel", font=("Segoe UI", 10), background=renk["zemin"], foreground=renk["metin"])
        style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=(11, 6))
        style.configure("Primary.TButton", background=renk["vurgu"], foreground="white", bordercolor=renk["vurgu"])
        style.map("Primary.TButton", background=[("active", renk["vurgu_koyu"]), ("pressed", renk["vurgu_koyu"])])
        style.configure("Secondary.TButton", background=renk["yuzey"], foreground=renk["metin"], bordercolor=renk["cizgi"])
        style.map("Secondary.TButton", background=[("active", renk["yuzey_ikincil"])])
        style.configure("Danger.TButton", background=renk["tehlike"], foreground="white", bordercolor=renk["tehlike"])
        style.configure("TEntry", padding=5, fieldbackground=renk["yuzey"])
        style.configure("TCombobox", padding=5, fieldbackground=renk["yuzey"])
        style.configure("TLabelframe", background=renk["zemin"], bordercolor=renk["cizgi"], relief="solid")
        style.configure("TLabelframe.Label", background=renk["zemin"], foreground=renk["metin"], font=("Segoe UI", 10, "bold"))
        style.configure("TNotebook", background=renk["zemin"], borderwidth=0, tabmargins=(0, 0, 0, 0))
        style.configure("TNotebook.Tab", font=("Segoe UI", 9, "bold"), padding=(11, 8), background=renk["yuzey_ikincil"], foreground=renk["metin_ikincil"])
        style.map("TNotebook.Tab", background=[("selected", renk["yuzey"])], foreground=[("selected", renk["vurgu"])])
        style.configure("Baslik.TLabel", font=("Segoe UI", 16, "bold"), foreground=renk["metin"])
        style.configure("AltBaslik.TLabel", font=("Segoe UI", 11, "bold"), foreground=renk["metin"])
        style.configure("Muted.TLabel", font=("Segoe UI", 9), foreground=renk["metin_ikincil"])
        style.configure("Status.TLabel", font=("Segoe UI", 9), background=renk["yuzey_ikincil"], foreground=renk["metin_ikincil"])
        style.configure("Panel.TFrame", background=renk["yuzey"])
        style.configure("Status.TFrame", background=renk["yuzey_ikincil"])
        style.configure("Treeview", font=("Segoe UI", 9), rowheight=32, background=renk["yuzey"], fieldbackground=renk["yuzey"], bordercolor=renk["cizgi"])
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), padding=(7, 7), background=renk["yuzey_ikincil"], foreground=renk["metin"])
        style.map("Treeview", background=[("selected", renk["vurgu"])], foreground=[("selected", "white")])

    def durum_cubugu_olustur(self):
        self.durum_var = tk.StringVar(value="Hazır")
        self.dosya_durum_var = tk.StringVar(value="Yeni proje")

        bar = ttk.Frame(self.root, padding=(10, 5), style="Status.TFrame")
        bar.pack(side="bottom", fill="x")
        ttk.Label(bar, textvariable=self.durum_var, style="Status.TLabel").pack(side="left")
        ttk.Label(bar, textvariable=self.dosya_durum_var, style="Status.TLabel").pack(side="right")
        self.durum_cubugu = bar

    def durum_mesaji_yaz(self, mesaj, dosya=None):
        if hasattr(self, "durum_var"):
            self.durum_var.set(mesaj)
        if dosya is not None and hasattr(self, "dosya_durum_var"):
            self.dosya_durum_var.set(dosya)

    def ikonlari_olustur(self):
        img_ac = Image.new("RGBA", (40, 40), (255, 255, 255, 0))
        draw_ac = ImageDraw.Draw(img_ac)
        draw_ac.rectangle((4, 4, 36, 36), fill="white", outline="black", width=4)
        self.ikon_ac = ImageTk.PhotoImage(img_ac)

        img_yn = Image.new("RGBA", (40, 40), (255, 255, 255, 0))
        draw_yn = ImageDraw.Draw(img_yn)
        draw_yn.ellipse((4, 4, 36, 36), fill="white", outline="black", width=4)
        self.ikon_yn = ImageTk.PhotoImage(img_yn)

        img_m = Image.new("RGBA", (40, 40), (255, 255, 255, 0))
        draw_m = ImageDraw.Draw(img_m)
        draw_m.ellipse((4, 4, 36, 36), fill="orange", outline="black", width=4)
        self.ikon_m = ImageTk.PhotoImage(img_m)

    def tablo_stillerini_hazirla(self):
        style = tb.Style()
        self.tablo_satir_yuksekligi = 32
        style.configure("Treeview", rowheight=self.tablo_satir_yuksekligi)
        style.configure("AcYn.Treeview", rowheight=self.tablo_satir_yuksekligi)
        style.configure("Lab.Treeview", rowheight=self.tablo_satir_yuksekligi)
        style.configure("Ekler.Treeview", rowheight=self.tablo_satir_yuksekligi)

    def hucre_duzenle(self, event, tree, set_row=None, set_col=None):
        if getattr(self, "proje_salt_okunur", False):
            return
        if event:
            region = tree.identify_region(event.x, event.y)
            if region != "cell":
                return
            column = tree.identify_column(event.x)
            row = tree.identify_row(event.y)
        else:
            row = set_row
            column = set_col
            if not row or not column:
                return

        if not tree.exists(row):
            return

        try:
            hucre_kutusu = tree.bbox(row, column)
            if not hucre_kutusu:
                tree.see(row)
                tree.update_idletasks()
                hucre_kutusu = tree.bbox(row, column)
        except tk.TclError:
            return
        if not hucre_kutusu:
            return

        x, y, width, height = hucre_kutusu
        try:
            mevcut_deger = tree.set(row, column)
        except tk.TclError:
            return

        entry = ttk.Entry(tree)
        entry.place(x=x, y=y, width=width, height=height)
        entry.insert(0, mevcut_deger)
        entry.selection_range(0, tk.END)
        entry.focus()

        col_idx = int(column.replace("#", "")) - 1
        items = tree.get_children()
        row_idx = items.index(row)

        def kaydet(e=None):
            if entry.winfo_exists():
                tree.set(row, column, entry.get())
                entry.destroy()

                cols = tree["columns"]
                if "Silt+Kil" in cols and "Çakıl" in cols:
                    if hasattr(self, "senkronize_ac_tablo"):
                        self.senkronize_ac_tablo()

        def hareket(e, dr, dc):
            kaydet()
            new_r = row_idx + dr
            new_c = col_idx + dc
            if 0 <= new_r < len(items) and 0 <= new_c < len(tree["columns"]):
                tree.selection_set(items[new_r])
                tree.focus(items[new_r])
                tree.see(items[new_r])
                tree.after(10, lambda: self.hucre_duzenle(None, tree, items[new_r], f"#{new_c+1}"))
            return "break"

        entry.bind("<Return>", lambda e: hareket(e, 1, 0))
        entry.bind("<Tab>", lambda e: hareket(e, 0, 1))
        entry.bind("<Shift-Tab>", lambda e: hareket(e, 0, -1))
        entry.bind("<Up>", lambda e: hareket(e, -1, 0))
        entry.bind("<Down>", lambda e: hareket(e, 1, 0))
        entry.bind("<FocusOut>", kaydet)

    def stripe_tree(self, tree):
        for index, item in enumerate(tree.get_children()):
            if index % 2 == 0:
                tree.item(item, tags=("evenrow",))
            else:
                tree.item(item, tags=("oddrow",))

    def tree_secili_satirlari_tasi(self, tree, yon):
        secili = sorted(list(tree.selection()), key=lambda item: tree.index(item))
        if not secili:
            messagebox.showwarning("Uyarı", "Lütfen önce taşınacak satırı seçin.")
            return

        secili_kumesi = set(secili)
        items = list(tree.get_children())
        if yon < 0:
            sirali_secili = secili
        else:
            sirali_secili = list(reversed(secili))

        tasindi = False
        for item in sirali_secili:
            index = items.index(item)
            yeni_index = index + yon
            if yeni_index < 0 or yeni_index >= len(items):
                continue
            komsu = items[yeni_index]
            if komsu in secili_kumesi:
                continue
            tree.move(item, "", yeni_index)
            items[index], items[yeni_index] = items[yeni_index], items[index]
            tasindi = True

        if tasindi:
            self.stripe_tree(tree)
            tree.selection_set(secili)
            tree.focus(secili[0])
            tree.see(secili[0])
            if tree is getattr(self, "tree_lab_ac", None) and hasattr(self, "senkronize_ac_tablo"):
                self.senkronize_ac_tablo()
