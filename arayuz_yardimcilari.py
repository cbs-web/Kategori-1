import tkinter as tk
from tkinter import ttk, messagebox

import ttkbootstrap as tb
from PIL import Image, ImageDraw, ImageTk


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

    def genel_stilleri_hazirla(self):
        self.root.option_add("*Font", "{Segoe UI} 10")
        self.root.option_add("*TCombobox*Listbox.font", "{Segoe UI} 10")

        style = tb.Style()
        style.configure(".", font=("Segoe UI", 10))
        style.configure("TLabel", font=("Segoe UI", 10))
        style.configure("TButton", font=("Segoe UI", 10), padding=(10, 5))
        style.configure("TEntry", padding=4)
        style.configure("TCombobox", padding=4)
        style.configure("TNotebook", tabmargins=(4, 4, 4, 0))
        style.configure("TNotebook.Tab", font=("Segoe UI", 9, "bold"), padding=(12, 7))
        style.configure("Baslik.TLabel", font=("Segoe UI", 14, "bold"), foreground="#1f2937")
        style.configure("AltBaslik.TLabel", font=("Segoe UI", 11, "bold"), foreground="#334155")
        style.configure("Muted.TLabel", font=("Segoe UI", 9), foreground="#64748b")
        style.configure("Status.TLabel", font=("Segoe UI", 9), foreground="#475569")
        style.configure("Panel.TFrame", background="#f8fafc")
        style.configure("Status.TFrame", background="#f1f5f9")
        style.configure("Treeview", font=("Segoe UI", 9), rowheight=42)
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), padding=(6, 6))

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
        self.tablo_satir_yuksekligi = 42
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
