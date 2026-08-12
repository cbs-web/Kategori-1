"""1/100.000 ölçekli jeoloji paftalarının K-1 arayüz bağlantıları."""

from __future__ import annotations

import datetime as dt
import os
from pathlib import Path
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageDraw, ImageTk

from jeoloji_pafta_kutuphanesi import (
    GORSEL_AZAMI_PIKSEL,
    JeolojiPaftaHatasi,
    JeolojiPaftaKutuphanesi,
    pafta_anahtari,
)
from jeoloji_pafta_tanima import paftada_birim_tahmin_et


class JeolojiPaftaIslemleri:
    def __init__(self, app):
        object.__setattr__(self, "app", app)

    def __getattr__(self, name):
        return getattr(self.app, name)

    def __setattr__(self, name, value):
        if name == "app":
            object.__setattr__(self, name, value)
        else:
            setattr(self.app, name, value)

    def _pafta_kutuphanesi(self):
        root = os.path.join(self.kullanici_veri_klasoru_bul(), "jeoloji_pafta_100k")
        current = getattr(self, "_jeoloji_pafta_kutuphanesi_nesnesi", None)
        if current is None or os.path.normcase(str(current.root_dir)) != os.path.normcase(root):
            current = JeolojiPaftaKutuphanesi(root)
            self._jeoloji_pafta_kutuphanesi_nesnesi = current
        return current

    def jeoloji_pafta_kutuphanesi_penceresi(self):
        window = getattr(self, "_jeoloji_pafta_penceresi", None)
        try:
            if window is not None and window.winfo_exists():
                window.deiconify()
                window.lift()
                self._pafta_listeyi_yenile()
                return window
        except tk.TclError:
            pass

        window = self.animasyonlu_pencere()
        self._jeoloji_pafta_penceresi = window
        window.title("1/100.000 Jeoloji Paftaları")
        window.geometry("1180x680")
        window.minsize(900, 520)

        header = ttk.Frame(window, padding=(12, 10))
        header.pack(fill="x")
        ttk.Label(header, text="1/100.000 Jeoloji Pafta Kütüphanesi", font=("Segoe UI", 14, "bold")).pack(side="left")
        ttk.Button(header, text="Kapat", command=window.destroy).pack(side="right")

        toolbar = ttk.Frame(window, padding=(12, 0, 12, 8))
        toolbar.pack(fill="x")
        ttk.Button(toolbar, text="KMZ + JPEG Klasörlerini Toplu Ekle", command=self._paftalari_toplu_ice_aktar).pack(side="left")
        ttk.Button(toolbar, text="Seçili Paftaya JPEG Bağla", command=self._paftaya_jpeg_bagla).pack(side="left", padx=(8, 0))
        ttk.Button(toolbar, text="Seçili JPEG Lejantını Düzenle", command=self._secili_lejanti_duzenle).pack(side="left", padx=(8, 0))
        ttk.Button(toolbar, text="Yenile", command=self._pafta_listeyi_yenile).pack(side="right")

        body = ttk.Frame(window, padding=(12, 0, 12, 12))
        body.pack(fill="both", expand=True)
        columns = ("pafta", "kmz", "jpeg", "eslesme", "lejant", "durum")
        tree = ttk.Treeview(body, columns=columns, show="headings", selectmode="browse")
        self._jeoloji_pafta_tree = tree
        headings = {
            "pafta": "Pafta / Katman",
            "kmz": "KMZ",
            "jpeg": "Açıklamalı JPEG",
            "eslesme": "Eşleştirme",
            "lejant": "Lejant Birimi",
            "durum": "Durum",
        }
        widths = {"pafta": 220, "kmz": 180, "jpeg": 190, "eslesme": 100, "lejant": 95, "durum": 180}
        for column in columns:
            tree.heading(column, text=headings[column])
            tree.column(column, width=widths[column], anchor="center" if column in ("eslesme", "lejant") else "w")
        yscroll = ttk.Scrollbar(body, orient="vertical", command=tree.yview)
        xscroll = ttk.Scrollbar(body, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)
        tree.bind("<Double-1>", lambda _event: self._secili_lejanti_duzenle())

        self._jeoloji_pafta_alt_durum = tk.StringVar(master=window, value="")
        ttk.Label(window, textvariable=self._jeoloji_pafta_alt_durum, padding=(12, 0, 12, 10)).pack(fill="x")
        self._pafta_listeyi_yenile()
        return window

    def _pafta_listeyi_yenile(self):
        tree = getattr(self, "_jeoloji_pafta_tree", None)
        if tree is None:
            return
        for iid in tree.get_children():
            tree.delete(iid)
        library = self._pafta_kutuphanesi()
        records = library.listele()
        ready = 0
        for record in records:
            profile = library.lejant_getir(record.get("lejant_id"))
            count = len((profile or {}).get("ogeler", []))
            jpeg = record.get("jpeg_path", "")
            jpeg_exists = bool(jpeg and os.path.isfile(jpeg))
            if not jpeg_exists:
                status = "JPEG bağlanmalı"
            elif not count:
                status = "Lejant tanımlanmalı"
            else:
                status = "Tanımaya hazır"
                ready += 1
            tree.insert(
                "",
                "end",
                iid=record["id"],
                values=(
                    record.get("ad", ""),
                    Path(record.get("kmz_path", "")).name,
                    Path(jpeg).name if jpeg else "—",
                    record.get("eslesme_yontemi", "") or "—",
                    count,
                    status,
                ),
            )
        if hasattr(self, "_jeoloji_pafta_alt_durum"):
            self._jeoloji_pafta_alt_durum.set(f"{len(records)} pafta kaydı • {ready} pafta tanımaya hazır")
        self.jeoloji_pafta_durumunu_guncelle()

    def _secili_pafta(self):
        tree = getattr(self, "_jeoloji_pafta_tree", None)
        selection = tree.selection() if tree is not None else ()
        if not selection:
            messagebox.showwarning("1/100.000 Paftalar", "Önce listeden bir pafta seçin.")
            return None
        return self._pafta_kutuphanesi().getir(selection[0])

    def _paftalari_toplu_ice_aktar(self):
        kmz_root = filedialog.askdirectory(title="KMZ dosyalarının bulunduğu klasörü seçin")
        if not kmz_root:
            return
        jpeg_root = filedialog.askdirectory(title="Açıklamalı JPEG paftalarının bulunduğu klasörü seçin")
        if not jpeg_root:
            return

        dialog = self.animasyonlu_pencere()
        dialog.title("Paftalar İçe Aktarılıyor")
        dialog.geometry("540x150")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.protocol("WM_DELETE_WINDOW", lambda: None)
        status_var = tk.StringVar(master=dialog, value="Dosyalar taranıyor...")
        ttk.Label(dialog, textvariable=status_var, wraplength=500, padding=(15, 15, 15, 8)).pack(fill="x")
        progress = ttk.Progressbar(dialog, mode="determinate", maximum=100)
        progress.pack(fill="x", padx=15, pady=(0, 15))
        messages = queue.Queue()
        library = self._pafta_kutuphanesi()

        def report(done, total, text):
            messages.put(("progress", done, total, text))

        def worker():
            try:
                result = library.toplu_ice_aktar(kmz_root, jpeg_root, ilerleme=report)
                messages.put(("done", result))
            except Exception as exc:
                messages.put(("error", exc))

        def poll():
            try:
                while True:
                    item = messages.get_nowait()
                    if item[0] == "progress":
                        _, done, total, text = item
                        progress["value"] = 100 * done / max(total, 1)
                        status_var.set(text)
                    elif item[0] == "done":
                        dialog.grab_release()
                        dialog.destroy()
                        result = item[1]
                        self._pafta_listeyi_yenile()
                        unmatched = sum(1 for record in result.get("paftalar", []) if not record.get("jpeg_path"))
                        messagebox.showinfo(
                            "Pafta Aktarımı Tamamlandı",
                            f"{len(result.get('paftalar', []))} pafta kaydı okundu.\n"
                            f"{result.get('jpeg_sayisi', 0)} açıklamalı JPEG tarandı.\n"
                            f"JPEG eşleşmesi bekleyen pafta: {unmatched}\n\n"
                            "Tanıma için her açıklamalı JPEG'in lejant örneklerini bir kez tanımlayın.",
                        )
                        return
                    elif item[0] == "error":
                        dialog.grab_release()
                        dialog.destroy()
                        self.hata_kaydet("Jeoloji paftaları toplu aktarılamadı", item[1])
                        messagebox.showerror("Pafta Aktarım Hatası", str(item[1]))
                        return
            except queue.Empty:
                pass
            try:
                dialog.after(120, poll)
            except tk.TclError:
                pass

        threading.Thread(target=worker, daemon=True).start()
        poll()

    def _paftaya_jpeg_bagla(self):
        record = self._secili_pafta()
        if not record:
            return
        path = filedialog.askopenfilename(
            title="Açıklamalı pafta JPEG'ini seçin",
            filetypes=(("JPEG", "*.jpg *.jpeg"), ("Tüm dosyalar", "*.*")),
        )
        if not path:
            return
        try:
            self._pafta_kutuphanesi().jpeg_bagla(record["id"], path)
            self._pafta_listeyi_yenile()
        except Exception as exc:
            messagebox.showerror("JPEG Bağlama Hatası", str(exc))

    def _secili_lejanti_duzenle(self):
        record = self._secili_pafta()
        if not record:
            return
        if not record.get("jpeg_path") or not os.path.isfile(record["jpeg_path"]):
            messagebox.showwarning("Lejant", "Önce seçili paftaya açıklamalı JPEG bağlayın.")
            return
        self._lejant_editorunu_ac(record)

    def _lejant_editorunu_ac(self, record):
        profile = self._pafta_kutuphanesi().lejant_getir(record.get("lejant_id"))
        if not profile:
            messagebox.showerror("Lejant", "Paftanın lejant profili bulunamadı.")
            return
        try:
            image = Image.open(profile["jpeg_path"])
            if image.width * image.height > GORSEL_AZAMI_PIKSEL:
                image.close()
                raise JeolojiPaftaHatasi("Açıklamalı JPEG güvenli piksel sınırını aşıyor.")
            image.load()
            image = image.convert("RGB")
        except Exception as exc:
            messagebox.showerror("Lejant", f"JPEG açılamadı:\n{exc}")
            return

        previous = getattr(self, "_jeoloji_pafta_lejant_editoru", None)
        try:
            if previous is not None and previous.winfo_exists():
                previous.destroy()
        except tk.TclError:
            pass
        window = self.animasyonlu_pencere()
        self._jeoloji_pafta_lejant_editoru = window
        window.title(f"Lejant Tanımlama — {Path(profile['jpeg_path']).name}")
        window.geometry("1320x820")
        window.minsize(980, 650)

        state = {
            "image": image,
            "zoom": min(0.32, 1500 / image.width, 1000 / image.height),
            "photo": None,
            "selection": None,
            "drag_start": None,
            "active_id": None,
            "profile": profile,
            "preview_photo": None,
        }
        state["zoom"] = max(0.06, state["zoom"])

        header = ttk.Frame(window, padding=(10, 8))
        header.pack(fill="x")
        ttk.Label(header, text="Lejant kutusunun yalnızca renk/tarama örneğini çerçeveleyin.", font=("Segoe UI", 11, "bold")).pack(side="left")
        ttk.Label(header, text="Kod ve adı yazıp Kaydet'e basın.").pack(side="left", padx=(12, 0))
        zoom_var = tk.StringVar(master=window)
        ttk.Button(header, text="−", width=3, command=lambda: change_zoom(0.8)).pack(side="right")
        ttk.Label(header, textvariable=zoom_var, width=8, anchor="center").pack(side="right")
        ttk.Button(header, text="+", width=3, command=lambda: change_zoom(1.25)).pack(side="right")

        pane = ttk.PanedWindow(window, orient="horizontal")
        pane.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        left = ttk.Frame(pane)
        right = ttk.Frame(pane, padding=(12, 0, 0, 0))
        pane.add(left, weight=4)
        pane.add(right, weight=1)

        canvas = tk.Canvas(left, background="#303030", cursor="crosshair", highlightthickness=0)
        xscroll = ttk.Scrollbar(left, orient="horizontal", command=canvas.xview)
        yscroll = ttk.Scrollbar(left, orient="vertical", command=canvas.yview)
        canvas.configure(xscrollcommand=xscroll.set, yscrollcommand=yscroll.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        left.rowconfigure(0, weight=1)
        left.columnconfigure(0, weight=1)

        code_var = tk.StringVar(master=window)
        name_var = tk.StringVar(master=window)
        ttk.Label(right, text="Birim kodu").pack(anchor="w")
        ttk.Entry(right, textvariable=code_var).pack(fill="x", pady=(2, 8))
        ttk.Label(right, text="Birim adı").pack(anchor="w")
        name_combo = ttk.Combobox(right, textvariable=name_var, state="normal", values=list(getattr(self, "formasyonlar", [])))
        name_combo.pack(fill="x", pady=(2, 8))
        preview = ttk.Label(right, text="Örnek seçilmedi", anchor="center")
        preview.pack(fill="x", pady=(2, 8), ipady=8)
        button_row = ttk.Frame(right)
        button_row.pack(fill="x", pady=(0, 8))

        item_tree = ttk.Treeview(right, columns=("kod", "ad"), show="headings", height=18, selectmode="browse")
        item_tree.heading("kod", text="Kod")
        item_tree.heading("ad", text="Birim")
        item_tree.column("kod", width=62, stretch=False)
        item_tree.column("ad", width=180)
        item_tree.pack(fill="both", expand=True)

        def current_profile():
            return self._pafta_kutuphanesi().lejant_getir(record.get("lejant_id")) or state["profile"]

        def fill_tree(select_id=None):
            for iid in item_tree.get_children():
                item_tree.delete(iid)
            state["profile"] = current_profile()
            for item in state["profile"].get("ogeler", []):
                item_tree.insert("", "end", iid=item["id"], values=(item.get("kod", ""), item.get("ad", "")))
            if select_id and item_tree.exists(select_id):
                item_tree.selection_set(select_id)
                item_tree.see(select_id)

        def update_preview():
            rect = state["selection"]
            if not rect:
                preview.configure(image="", text="Örnek seçilmedi")
                state["preview_photo"] = None
                return
            x1, y1, x2, y2 = rect
            box = (
                max(0, round(x1 * image.width)),
                max(0, round(y1 * image.height)),
                min(image.width, round(x2 * image.width)),
                min(image.height, round(y2 * image.height)),
            )
            crop = image.crop(box)
            crop.thumbnail((230, 115), Image.Resampling.LANCZOS)
            state["preview_photo"] = ImageTk.PhotoImage(crop, master=window)
            preview.configure(image=state["preview_photo"], text="")
            crop.close()

        def render():
            zoom = state["zoom"]
            display = image.resize((max(1, round(image.width * zoom)), max(1, round(image.height * zoom))), Image.Resampling.LANCZOS)
            state["photo"] = ImageTk.PhotoImage(display, master=window)
            display.close()
            canvas.delete("all")
            canvas.create_image(0, 0, anchor="nw", image=state["photo"])
            for item in state["profile"].get("ogeler", []):
                x1, y1, x2, y2 = item.get("rect", (0, 0, 0, 0))
                canvas.create_rectangle(x1 * image.width * zoom, y1 * image.height * zoom, x2 * image.width * zoom, y2 * image.height * zoom, outline="#20d875", width=2)
            if state["selection"]:
                x1, y1, x2, y2 = state["selection"]
                canvas.create_rectangle(x1 * image.width * zoom, y1 * image.height * zoom, x2 * image.width * zoom, y2 * image.height * zoom, outline="#ff2b38", width=3, tags="selection")
            canvas.configure(scrollregion=(0, 0, image.width * zoom, image.height * zoom))
            zoom_var.set(f"%{round(zoom * 100)}")

        def change_zoom(factor):
            state["zoom"] = max(0.04, min(1.5, state["zoom"] * factor))
            render()

        def begin_drag(event):
            state["drag_start"] = (canvas.canvasx(event.x), canvas.canvasy(event.y))

        def drag(event):
            if not state["drag_start"]:
                return
            x1, y1 = state["drag_start"]
            x2, y2 = canvas.canvasx(event.x), canvas.canvasy(event.y)
            canvas.delete("drag")
            canvas.create_rectangle(x1, y1, x2, y2, outline="#ff2b38", width=3, tags="drag")

        def end_drag(event):
            if not state["drag_start"]:
                return
            x1, y1 = state["drag_start"]
            x2, y2 = canvas.canvasx(event.x), canvas.canvasy(event.y)
            state["drag_start"] = None
            canvas.delete("drag")
            if abs(x2 - x1) < 5 or abs(y2 - y1) < 5:
                return
            zoom = state["zoom"]
            state["selection"] = [
                max(0.0, min(x1, x2) / (image.width * zoom)),
                max(0.0, min(y1, y2) / (image.height * zoom)),
                min(1.0, max(x1, x2) / (image.width * zoom)),
                min(1.0, max(y1, y2) / (image.height * zoom)),
            ]
            render()
            update_preview()

        def new_item():
            state["active_id"] = None
            state["selection"] = None
            code_var.set("")
            name_var.set("")
            item_tree.selection_remove(*item_tree.selection())
            render()
            update_preview()

        def choose_item(_event=None):
            selection = item_tree.selection()
            if not selection:
                return
            item_id = selection[0]
            item = next((entry for entry in state["profile"].get("ogeler", []) if entry.get("id") == item_id), None)
            if not item:
                return
            state["active_id"] = item_id
            state["selection"] = list(item.get("rect", []))
            code_var.set(item.get("kod", ""))
            name_var.set(item.get("ad", ""))
            render()
            update_preview()

        def formation_selected(_event=None):
            value = name_var.get().strip()
            if "(" in value and value.endswith(")"):
                name, code = value.rsplit("(", 1)
                name_var.set(name.strip())
                if not code_var.get().strip():
                    code_var.set(code[:-1].strip())

        def save_item():
            try:
                item = self._pafta_kutuphanesi().lejant_ogesi_kaydet(
                    record["lejant_id"], code_var.get(), name_var.get(), state["selection"], state["active_id"]
                )
            except Exception as exc:
                messagebox.showerror("Lejant Kaydı", str(exc), parent=window)
                return
            state["active_id"] = item["id"]
            fill_tree(item["id"])
            render()
            self._pafta_listeyi_yenile()

        def delete_item():
            if not state["active_id"]:
                return
            if not messagebox.askyesno("Lejant", "Seçili lejant örneği silinsin mi?", parent=window):
                return
            self._pafta_kutuphanesi().lejant_ogesi_sil(record["lejant_id"], state["active_id"])
            new_item()
            fill_tree()
            self._pafta_listeyi_yenile()

        ttk.Button(button_row, text="Yeni", command=new_item).pack(side="left")
        ttk.Button(button_row, text="Kaydet", command=save_item).pack(side="left", padx=(5, 0))
        ttk.Button(button_row, text="Sil", command=delete_item).pack(side="left", padx=(5, 0))
        canvas.bind("<ButtonPress-1>", begin_drag)
        canvas.bind("<B1-Motion>", drag)
        canvas.bind("<ButtonRelease-1>", end_drag)
        item_tree.bind("<<TreeviewSelect>>", choose_item)
        name_combo.bind("<<ComboboxSelected>>", formation_selected)

        def close():
            image.close()
            window.destroy()

        window.protocol("WM_DELETE_WINDOW", close)
        fill_tree()
        render()

    def jeoloji_pafta_durumunu_guncelle(self):
        variable = getattr(self, "jeoloji_pafta_durum", None)
        if variable is None:
            return
        result = getattr(self, "jeoloji_pafta_sonucu", {})
        if isinstance(result, dict) and result.get("birim_adi"):
            code = result.get("birim_kodu", "")
            suffix = f" ({code})" if code else ""
            variable.set(f"1/100.000: {result['birim_adi']}{suffix} • güven %{result.get('guven', 0)}")
            return
        points = getattr(self, "yuklu_kml_points", [])
        if not points:
            variable.set("1/100.000 formasyon tanıması için parsel KML'si yükleyin")
            return
        try:
            records = self._pafta_kutuphanesi().kapsayan_paftalar(points)
            ready = 0
            for record in records:
                profile = self._pafta_kutuphanesi().lejant_getir(record.get("lejant_id"))
                if record.get("jpeg_path") and (profile or {}).get("ogeler"):
                    ready += 1
            if ready:
                variable.set(f"Parsel hazır • {ready} adet 1/100.000 pafta tanımaya uygun")
            elif records:
                variable.set("Pafta bulundu; açıklamalı JPEG/lejant tanımı tamamlanmalı")
            else:
                variable.set("Parseli kapsayan kayıtlı 1/100.000 pafta bulunamadı")
        except Exception:
            variable.set("1/100.000 pafta durumu okunamadı")

    def formasyonu_jeoloji_haritasindan_bul(self):
        points = getattr(self, "yuklu_kml_points", [])
        if len(points) < 3:
            messagebox.showwarning("Formasyonu Haritadan Bul", "Önce çalışma parselinin KML dosyasını yükleyin.")
            return
        library = self._pafta_kutuphanesi()
        covering = library.kapsayan_paftalar(points)
        if not covering:
            messagebox.showwarning(
                "Formasyonu Haritadan Bul",
                "Parseli kapsayan 1/100.000 pafta bulunamadı. Önce pafta kütüphanesine KMZ ve JPEG dosyalarını ekleyin.",
            )
            return
        ready = []
        incomplete = []
        for record in covering:
            profile = library.lejant_getir(record.get("lejant_id"))
            if record.get("jpeg_path") and profile and profile.get("ogeler"):
                ready.append((record, profile))
            else:
                incomplete.append(record.get("ad", "Adsız pafta"))
        if not ready:
            messagebox.showwarning(
                "Formasyonu Haritadan Bul",
                "Parseli kapsayan pafta bulundu; ancak açıklamalı JPEG veya lejant örnekleri eksik.\n\n"
                "Araçlar > 1/100.000 Jeoloji Paftaları bölümünden paftayı hazırlayın.",
            )
            return

        self.root.configure(cursor="watch")
        self.root.update_idletasks()
        results = []
        errors = []
        try:
            for record, profile in ready:
                try:
                    results.append(paftada_birim_tahmin_et(record, profile, points))
                except Exception as exc:
                    errors.append(f"{record.get('ad', 'Pafta')}: {exc}")
        finally:
            self.root.configure(cursor="")
        if not results:
            messagebox.showerror("Formasyon Tanıma Hatası", "\n".join(errors) or "Tanıma tamamlanamadı.")
            return
        results.sort(
            key=lambda result: (
                -result["adaylar"][0]["guven"],
                -result["adaylar"][0]["oran"],
                -result["pafta"].get("kapsama_orani", 0),
            )
        )
        self._formasyon_sonuc_penceresi(results, errors, incomplete)

    def _formasyon_sonuc_penceresi(self, results, errors, incomplete):
        window = self.animasyonlu_pencere()
        window.title("1/100.000 Jeoloji Paftası — Formasyon Adayları")
        window.geometry("1280x720")
        window.minsize(980, 600)
        state = {
            "photo": None,
            "legend_photo": None,
            "legend_thumbnails": {},
            "selected": (0, 0),
        }

        header = ttk.Frame(window, padding=(12, 10))
        header.pack(fill="x")
        top = results[0]["adaylar"][0]
        warning = "Düşük güven — görseli ve adayları kontrol edin." if results[0]["dusuk_guven"] else "Sonuç kullanıcı onayı bekliyor."
        ttk.Label(header, text=f"En güçlü aday: {top['ad']} ({top['kod']})", font=("Segoe UI", 13, "bold")).pack(anchor="w")
        ttk.Label(header, text=warning).pack(anchor="w", pady=(3, 0))

        pane = ttk.PanedWindow(window, orient="horizontal")
        pane.pack(fill="both", expand=True, padx=12, pady=(0, 8))
        left = ttk.Frame(pane)
        right = ttk.Frame(pane, padding=(10, 0, 0, 0))
        pane.add(left, weight=3)
        pane.add(right, weight=3)
        preview = ttk.Label(left, anchor="center")
        preview.pack(fill="both", expand=True)

        legend_panel = ttk.LabelFrame(left, text="Seçili adayın pafta lejantı", padding=(10, 8))
        legend_panel.pack(fill="x", pady=(8, 0))
        legend_preview = ttk.Label(legend_panel, anchor="center")
        legend_preview.pack(side="left")
        legend_text_var = tk.StringVar(master=window, value="")
        ttk.Label(
            legend_panel,
            textvariable=legend_text_var,
            justify="left",
            wraplength=330,
            padding=(12, 0, 0, 0),
        ).pack(side="left", fill="x", expand=True)

        columns = ("kod", "ad", "ornek", "uyum", "guven", "pafta")
        tree_style_name = f"FormasyonAdaylari{id(window)}.Treeview"
        ttk.Style(window).configure(tree_style_name, rowheight=46)
        tree = ttk.Treeview(
            right,
            columns=columns,
            show=("tree", "headings"),
            selectmode="browse",
            style=tree_style_name,
        )
        tree.heading("#0", text="Lejant")
        tree.column("#0", width=86, minwidth=86, stretch=False, anchor="center")
        labels = {"kod": "Kod", "ad": "Birim", "ornek": "Örnek %", "uyum": "Görsel Uyum", "guven": "Güven", "pafta": "Pafta"}
        widths = {"kod": 55, "ad": 175, "ornek": 70, "uyum": 85, "guven": 65, "pafta": 130}
        for column in columns:
            tree.heading(column, text=labels[column])
            tree.column(column, width=widths[column], anchor="center" if column != "ad" else "w")
        tree.pack(fill="both", expand=True)

        def framed_legend(source, size):
            width, height = size
            canvas = Image.new("RGB", size, "white")
            sample = source.convert("RGB")
            available_width = max(1, width - 10)
            available_height = max(1, height - 10)
            scale = min(available_width / max(sample.width, 1), available_height / max(sample.height, 1))
            sample = sample.resize(
                (
                    max(1, round(sample.width * scale)),
                    max(1, round(sample.height * scale)),
                ),
                Image.Resampling.LANCZOS,
            )
            canvas.paste(sample, ((width - sample.width) // 2, (height - sample.height) // 2))
            sample.close()
            ImageDraw.Draw(canvas).rectangle((0, 0, width - 1, height - 1), outline="#60656d", width=1)
            return canvas

        for result_index, result in enumerate(results):
            for candidate_index, candidate in enumerate(result["adaylar"]):
                iid = f"{result_index}:{candidate_index}"
                previews = result.get("lejant_gorselleri", [])
                legend_image = previews[candidate_index] if candidate_index < len(previews) else None
                options = {
                    "iid": iid,
                    "values": (
                        candidate["kod"],
                        candidate["ad"],
                        candidate["oran"],
                        candidate["puan"],
                        candidate["guven"],
                        candidate["pafta_adi"],
                    ),
                }
                if legend_image is not None:
                    thumbnail = framed_legend(legend_image, (74, 36))
                    photo = ImageTk.PhotoImage(thumbnail, master=window)
                    thumbnail.close()
                    state["legend_thumbnails"][iid] = photo
                    options["image"] = photo
                tree.insert("", "end", **options)

        detail_var = tk.StringVar(master=window)
        ttk.Label(window, textvariable=detail_var, padding=(12, 0, 12, 8)).pack(fill="x")
        buttons = ttk.Frame(window, padding=(12, 0, 12, 12))
        buttons.pack(fill="x")
        ttk.Button(buttons, text="İptal", command=lambda: close()).pack(side="right")

        def show_result(result_index):
            evidence = results[result_index]["kanit_gorseli"]
            display = evidence.copy()
            display.thumbnail((650, 510), Image.Resampling.LANCZOS)
            state["photo"] = ImageTk.PhotoImage(display, master=window)
            display.close()
            preview.configure(image=state["photo"])

        def show_legend(result_index, candidate_index):
            candidate = results[result_index]["adaylar"][candidate_index]
            previews = results[result_index].get("lejant_gorselleri", [])
            legend_image = previews[candidate_index] if candidate_index < len(previews) else None
            legend_text_var.set(
                f"{candidate['kod']} — {candidate['ad']}\n"
                f"{candidate['pafta_adi']} paftasında kullanılan gerçek lejant örneği"
            )
            if legend_image is None:
                state["legend_photo"] = None
                legend_preview.configure(image="", text="Lejant örneği bulunamadı")
                return
            display = framed_legend(legend_image, (250, 96))
            state["legend_photo"] = ImageTk.PhotoImage(display, master=window)
            display.close()
            legend_preview.configure(image=state["legend_photo"], text="")

        def selection_changed(_event=None):
            selection = tree.selection()
            if not selection:
                return
            result_index, candidate_index = (int(value) for value in selection[0].split(":"))
            state["selected"] = (result_index, candidate_index)
            candidate = results[result_index]["adaylar"][candidate_index]
            show_result(result_index)
            show_legend(result_index, candidate_index)
            detail_var.set(
                f"{candidate['pafta_adi']} • {results[result_index]['ornek_sayisi']} parsel örneği • "
                f"Tahmin otomatik uygulanmaz; seçiminiz projeye kaydedilir."
            )

        def apply():
            result_index, candidate_index = state["selected"]
            candidate = results[result_index]["adaylar"][candidate_index]
            result = results[result_index]
            project_path = getattr(self, "guncel_dosya_yolu", None)
            target_dir = os.path.join(os.path.dirname(os.path.abspath(project_path)), "Haritalar") if project_path else None
            try:
                evidence_path = self._pafta_kutuphanesi().kanit_kaydet(result["kanit_gorseli"], target_dir)
            except Exception as exc:
                messagebox.showerror("Formasyon Sonucu", f"Kanıt görseli kaydedilemedi:\n{exc}", parent=window)
                return
            self._formasyon_adayini_uygula(candidate)
            self.jeoloji_pafta_sonucu = {
                "pafta_id": candidate.get("pafta_id", ""),
                "pafta_adi": candidate.get("pafta_adi", ""),
                "birim_kodu": candidate.get("kod", ""),
                "birim_adi": candidate.get("ad", ""),
                "oran": candidate.get("oran", 0),
                "puan": candidate.get("puan", 0),
                "guven": candidate.get("guven", 0),
                "kanit_yolu": evidence_path,
                "tarih": dt.datetime.now().isoformat(timespec="seconds"),
                "adaylar": [dict(item) for item in result["adaylar"][:5]],
            }
            self.jeoloji_pafta_durumunu_guncelle()
            if hasattr(self, "proje_durum_seridi_guncelle"):
                self.proje_durum_seridi_guncelle()
            close()
            messagebox.showinfo(
                "Formasyon Uygulandı",
                f"{candidate['ad']} ({candidate['kod']}) projeye aktarıldı.\n\nKanıt görseli: {evidence_path}",
            )

        ttk.Button(buttons, text="Seçili Adayı Projeye Uygula", command=apply).pack(side="right", padx=(0, 8))
        tree.bind("<<TreeviewSelect>>", selection_changed)
        first = tree.get_children()[0]
        tree.selection_set(first)
        tree.focus(first)
        selection_changed()

        def close():
            for result in results:
                try:
                    result["kanit_gorseli"].close()
                except Exception:
                    pass
                for legend_image in result.get("lejant_gorselleri", []):
                    try:
                        if legend_image is not None:
                            legend_image.close()
                    except Exception:
                        pass
            state["legend_thumbnails"].clear()
            state["legend_photo"] = None
            window.destroy()

        window.protocol("WM_DELETE_WINDOW", close)

    def _formasyon_adayini_uygula(self, candidate):
        name = str(candidate.get("ad", "")).strip()
        code = str(candidate.get("kod", "")).strip()
        display = f"{name} ({code})" if code else name
        wanted_name = pafta_anahtari(name)
        wanted_code = pafta_anahtari(code)
        for existing in getattr(self, "formasyonlar", []):
            existing_name = existing.rsplit("(", 1)[0].strip() if "(" in existing else existing
            existing_code = existing.rsplit("(", 1)[1].rstrip(")").strip() if "(" in existing else ""
            if pafta_anahtari(existing_name) == wanted_name and (
                not wanted_code or pafta_anahtari(existing_code) == wanted_code
            ):
                display = existing
                break
        if display not in self.formasyonlar:
            self.formasyonlar.append(display)
        self.formasyon_metinleri[display] = (
            f'"{code}" simgesiyle gösterilen "{name}" adı ile anılan' if code else f'"{name}" adı ile anılan'
        )
        self.combo_formasyon.configure(values=self.formasyonlar)
        self.combo_formasyon.set(display)
        self.formasyon_degisti()


__all__ = ["JeolojiPaftaIslemleri"]
