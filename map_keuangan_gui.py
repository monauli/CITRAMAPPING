"""GUI mapping keuangan — cek Excel rekap vs PDF laporan sumber.

Logic cek/banding dipakai dari map_keuangan.py (sama persis dengan versi
command-line), file ini cuma lapisan tampilan + penyimpanan:
- Pilih satu folder data (isinya 1 Excel rekap + PDF bulanan) — periode
  bulan terdeteksi otomatis dari Excel, gak perlu isi tahun manual.
- Hasil dikelompokkan per bulan.
- Export hasil satu run ke Excel/CSV.
- Riwayat tiap run tersimpan otomatis di map_keuangan_history.sqlite3.
"""
import csv
import sqlite3
import traceback
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk
import openpyxl

import auth
import map_keuangan as mk
from app_paths import CORRECTED_PDF_DIR, DB_PATH, ensure_app_dirs

STATUS_COLOR = {"BEDA": "#ff6b6b", "HANYA_EXCEL": "#e8c067", "HANYA_PDF": "#6bb3f0", "COCOK": "#8fd68f"}
PILL_BG = {"COCOK": "#1f3d1f", "BEDA": "#4a1f1f", "HANYA_EXCEL": "#4a3a1f", "HANYA_PDF": "#1f2f4a"}
STATUS_ICON = {"COCOK": "✓", "BEDA": "✕", "HANYA_EXCEL": "◐", "HANYA_PDF": "◑"}
VISIBLE_STATUSES = ("COCOK", "BEDA", "HANYA_EXCEL")

FONT_TITLE = ("Segoe UI", 18, "bold")
FONT_SUBTITLE = ("Segoe UI", 12)
FONT_SECTION = ("Segoe UI", 11, "bold")
FONT_BODY = ("Segoe UI", 12)
FONT_SMALL = ("Segoe UI", 11)
FONT_MONTH = ("Segoe UI", 13)
MUTED = "#9a9a9a"
CARD_BG = "#1f1f1f"
SHEET_LABELS = {
    "balance-sheet": "Balance Sheet",
    "profit-loss": "Profit & Loss",
    "cashflow": "Cash Flow",
}
SHEET_KINDS = {label: kind for kind, label in SHEET_LABELS.items()}


def ensure_db():
    ensure_app_dirs()
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""CREATE TABLE IF NOT EXISTS runs (
        id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT, excel_path TEXT, pdf_dir TEXT,
        cocok INTEGER, beda INTEGER, hanya_excel INTEGER, hanya_pdf INTEGER)""")
    conn.execute("""CREATE TABLE IF NOT EXISTS run_rows (
        run_id INTEGER, period TEXT, bulan TEXT, sheet TEXT, label TEXT,
        excel_value REAL, pdf_value REAL, selisih REAL, status TEXT)""")
    conn.execute("""CREATE TABLE IF NOT EXISTS corrections (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        run_id INTEGER, ts TEXT, username TEXT, period TEXT, sheet TEXT,
        label TEXT, correction_type TEXT, old_value REAL, new_value REAL,
        output_file TEXT, status TEXT)""")
    auth.init_auth_db(conn)
    conn.commit()
    return conn


def save_run(conn, excel_path, pdf_dir, rows):
    counts = {"COCOK": 0, "BEDA": 0, "HANYA_EXCEL": 0, "HANYA_PDF": 0}
    for r in rows:
        counts[r["status"]] += 1
    cur = conn.execute(
        "INSERT INTO runs(ts, excel_path, pdf_dir, cocok, beda, hanya_excel, hanya_pdf) VALUES (?,?,?,?,?,?,?)",
        (datetime.now().isoformat(timespec="seconds"), str(excel_path), str(pdf_dir),
         counts["COCOK"], counts["BEDA"], counts["HANYA_EXCEL"], counts["HANYA_PDF"]))
    run_id = cur.lastrowid
    conn.executemany(
        "INSERT INTO run_rows(run_id,period,bulan,sheet,label,excel_value,pdf_value,selisih,status) "
        "VALUES (?,?,?,?,?,?,?,?,?)",
        [(run_id, r["period"], r["month"], r["sheet"], r["label"], r["value"], r["pdf_value"], r["diff"], r["status"])
         for r in rows])
    conn.commit()
    return run_id, counts


def export_rows(rows, path):
    path = Path(path)
    headers = ["Periode", "Bulan", "Sheet", "Label", "Excel", "PDF", "Selisih", "Status"]
    data = [[r["period"], r["month"], r["sheet"], r["label"], r["value"], r["pdf_value"], r["diff"], r["status"]]
            for r in rows]
    if path.suffix.lower() == ".csv":
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(headers)
            writer.writerows(data)
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Hasil Cek"
        ws.append(headers)
        for row in data:
            ws.append(row)
        wb.save(path)


def summarize(rows):
    counts = {"COCOK": 0, "BEDA": 0, "HANYA_EXCEL": 0, "HANYA_PDF": 0}
    for r in rows:
        counts[r["status"]] += 1
    return counts


def save_corrections(conn, run_id, username, rows, correction_type, output_file="", skipped_labels=()):
    """Simpan audit trail koreksi tanpa menyimpan password atau file binary."""
    skipped = set(skipped_labels)
    now = datetime.now().isoformat(timespec="seconds")
    conn.executemany(
        "INSERT INTO corrections(run_id,ts,username,period,sheet,label,correction_type,"
        "old_value,new_value,output_file,status) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
        [(
            run_id, now, username, row["period"], row["sheet"], row["label"], correction_type,
            row.get("pdf_value") if correction_type == "excel" else row.get("pdf_value"),
            row.get("pdf_value") if correction_type == "pdf" else row.get("value"),
            output_file, "skipped_formula" if row["label"] in skipped else "applied",
        ) for row in rows],
    )
    conn.commit()


def apply_dark_treeview_style():
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Treeview", background="#242424", fieldbackground="#242424", foreground="#dcdcdc",
                     borderwidth=0, rowheight=28, font=FONT_SMALL)
    style.configure("Treeview.Heading", background="#161616", foreground="#ffffff", relief="flat",
                     font=FONT_SECTION, padding=(10, 8))
    style.map("Treeview", background=[("selected", "#2f6fed")], foreground=[("selected", "#ffffff")])
    style.map("Treeview.Heading", background=[("active", "#2a2a2a")])
    style.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])

    # Style terpisah buat panel bulan: baris lebih tinggi, gampang diklik.
    style.configure("Month.Treeview", background="#1c1c1c", fieldbackground="#1c1c1c",
                     foreground="#dcdcdc", borderwidth=0, rowheight=36, font=FONT_MONTH)
    style.map("Month.Treeview", background=[("selected", "#2f6fed")], foreground=[("selected", "#ffffff")])
    style.layout("Month.Treeview", [("Treeview.treearea", {"sticky": "nswe"})])


class Pill(ctk.CTkLabel):
    def __init__(self, parent, status, count):
        super().__init__(parent, text=f"{STATUS_ICON[status]}  {count}", fg_color=PILL_BG[status],
                          text_color=STATUS_COLOR[status], corner_radius=8, padx=14, pady=6,
                          font=("Segoe UI", 14, "bold"))

    def update_count(self, count):
        text = self.cget("text")
        icon = text.split()[0]
        self.configure(text=f"{icon}  {count}")


class App(ctk.CTk):
    def __init__(self, current_user):
        super().__init__()
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        apply_dark_treeview_style()

        self.title("Mapping Keuangan — Cek Excel vs PDF")
        self.geometry("1200x720")
        self.minsize(960, 600)

        self.conn = ensure_db()
        self.current_user = current_user
        self.excel_path = None
        self.pdf_dir = None
        self.rows = []          # semua baris, lintas bulan — sumber export/apply
        self.by_period = {}     # period -> rows bulan itu
        self.periods = []
        self.missing_periods = []
        self.current_period = None
        self.current_kind = None
        self.current_run_id = None
        self.viewing_history = False

        self._build_header()
        self._build_body()
        self.status_label = ctk.CTkLabel(self, text="Pilih folder data buat mulai.", anchor="w",
                                          text_color="#9a9a9a")
        self.status_label.pack(fill="x", padx=14, pady=(0, 10))

        self._refresh_history()

    # --- layout ---------------------------------------------------------

    def _build_header(self):
        top = ctk.CTkFrame(self, corner_radius=0, fg_color=CARD_BG)
        top.pack(fill="x")
        inner = ctk.CTkFrame(top, fg_color="transparent")
        inner.pack(fill="x", padx=18, pady=(16, 12))

        title_box = ctk.CTkFrame(inner, fg_color="transparent")
        title_box.pack(side="left", fill="y")
        ctk.CTkLabel(title_box, text="Mapping Keuangan", font=FONT_TITLE, anchor="w").pack(anchor="w")
        ctk.CTkLabel(title_box, text="Cek Excel rekap vs PDF laporan sumber, per bulan",
                     font=FONT_SUBTITLE, text_color=MUTED, anchor="w").pack(anchor="w")

        self.reload_btn = ctk.CTkButton(inner, text="↻  Jalankan Ulang", width=130, command=self._run_check,
                                         state="disabled")
        self.reload_btn.pack(side="right")
        if self.current_user["role"] == "admin":
            ctk.CTkButton(inner, text="⚙  User", width=90, command=self._open_users,
                          fg_color="transparent", border_width=1, border_color="#4a4a4a",
                          hover_color="#2a2a2a").pack(side="right", padx=(0, 10))
        ctk.CTkButton(inner, text="📁  Pilih Folder Data", width=170, command=self._pick_folder).pack(
            side="right", padx=(0, 10))

        folder_row = ctk.CTkFrame(top, fg_color="transparent")
        folder_row.pack(fill="x", padx=18, pady=(0, 14))
        self.folder_info = ctk.CTkLabel(folder_row, text="Belum pilih folder — isinya 1 Excel rekap + PDF bulanan.",
                                         anchor="w", font=FONT_SMALL, text_color=MUTED, justify="left")
        self.folder_info.pack(anchor="w", fill="x")

        self.summary_row = ctk.CTkFrame(top, fg_color="transparent")
        self.summary_row.pack(fill="x", padx=18, pady=(0, 16))
        self.pills = {}
        labels = {"COCOK": "Cocok", "BEDA": "Beda", "HANYA_EXCEL": "Hanya Excel", "HANYA_PDF": "Hanya PDF"}
        for status in VISIBLE_STATUSES:
            box = ctk.CTkFrame(self.summary_row, fg_color="transparent")
            box.pack(side="left", padx=(0, 16))
            pill = Pill(box, status, "–")
            pill.pack()
            ctk.CTkLabel(box, text=labels[status], font=FONT_SMALL, text_color=MUTED).pack(pady=(4, 0))
            self.pills[status] = pill

    def _build_body(self):
        self.tabs = ctk.CTkTabview(self)
        self.tabs.pack(fill="both", expand=True, padx=14, pady=(0, 6))
        self.tabs.add("Hasil per Bulan")
        self.tabs.add("Riwayat")
        self._build_result_tab(self.tabs.tab("Hasil per Bulan"))
        self._build_history_tab(self.tabs.tab("Riwayat"))

    def _build_result_tab(self, parent):
        parent.grid_columnconfigure(1, weight=1)
        parent.grid_rowconfigure(0, weight=1)
        parent.configure(fg_color="transparent")

        # Panel kiri: daftar bulan
        left = ctk.CTkFrame(parent, width=230, fg_color=CARD_BG)
        left.grid(row=0, column=0, sticky="ns", padx=(14, 10), pady=14)
        ctk.CTkLabel(left, text="BULAN", text_color=MUTED, font=FONT_SECTION).pack(
            anchor="w", padx=14, pady=(14, 6))
        self.month_tree = ttk.Treeview(left, columns=("bulan", "status"), show="headings", height=14,
                                        style="Month.Treeview")
        self.month_tree.heading("bulan", text="Bulan")
        self.month_tree.heading("status", text="")
        self.month_tree.column("bulan", width=140, anchor="w")
        self.month_tree.column("status", width=70, anchor="e")
        self.month_tree.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        self.month_tree.bind("<<TreeviewSelect>>", self._on_month_select)

        # Panel kanan: toolbar + tabel
        right = ctk.CTkFrame(parent, fg_color=CARD_BG)
        right.grid(row=0, column=1, sticky="nsew", padx=(0, 14), pady=14)
        right.grid_rowconfigure(1, weight=1)
        right.grid_columnconfigure(0, weight=1)

        toolbar = ctk.CTkFrame(right, fg_color="transparent")
        toolbar.grid(row=0, column=0, sticky="ew", padx=14, pady=14)
        self.month_title = ctk.CTkLabel(toolbar, text="Pilih bulan di kiri", font=FONT_TITLE)
        self.month_title.pack(side="left")
        self.category_segment = ctk.CTkSegmentedButton(
            toolbar,
            values=["Semua", "Balance Sheet", "Profit & Loss", "Cash Flow"],
            command=self._on_category_change,
        )
        self.category_segment.set("Semua")
        self.category_segment.pack(side="left", padx=(18, 0))
        ctk.CTkButton(toolbar, text="✕  Terapkan ke Excel", width=160, fg_color="#8a2d2d",
                      hover_color="#6e2424", command=self._apply_fixes).pack(side="right")
        ctk.CTkButton(toolbar, text="PDF Koreksi", width=120, fg_color="#315d8f",
                      hover_color="#274a72", command=self._make_corrected_pdfs).pack(side="right", padx=(0, 8))
        ctk.CTkButton(toolbar, text="💾  Simpan Excel/CSV...", width=170,
                      fg_color="transparent", border_width=1, border_color="#4a4a4a",
                      hover_color="#2a2a2a", command=self._export).pack(side="right", padx=(0, 8))

        table_frame = ctk.CTkFrame(right, fg_color="transparent")
        table_frame.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 14))
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        columns = ("sheet", "label", "excel", "pdf", "selisih", "status")
        headers = {"sheet": "Sheet", "label": "Label", "excel": "Excel", "pdf": "PDF",
                   "selisih": "Selisih", "status": "Status"}
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        for c in columns:
            self.tree.heading(c, text=headers[c], command=lambda c=c: self._sort_tree(self.tree, c))
            anchor = "e" if c in ("excel", "pdf", "selisih") else "w"
            width = 280 if c == "label" else (110 if c != "sheet" else 120)
            self.tree.column(c, width=width, anchor=anchor)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.grid(row=0, column=1, sticky="ns")

    def _build_history_tab(self, parent):
        parent.configure(fg_color="transparent")
        card = ctk.CTkFrame(parent, fg_color=CARD_BG)
        card.pack(fill="both", expand=True, padx=14, pady=14)

        header_row = ctk.CTkFrame(card, fg_color="transparent")
        header_row.pack(fill="x", padx=14, pady=14)
        ctk.CTkLabel(header_row, text="Riwayat Run", font=FONT_TITLE).pack(side="left")
        ctk.CTkButton(header_row, text="↻  Refresh", width=100, fg_color="transparent",
                      border_width=1, border_color="#4a4a4a", hover_color="#2a2a2a",
                      command=self._refresh_history).pack(side="right")

        columns = ("id", "waktu", "cocok", "beda", "hanya_excel")
        headers = {"id": "#", "waktu": "Waktu", "cocok": "Cocok", "beda": "Beda",
                   "hanya_excel": "Hanya Excel"}
        self.history_tree = ttk.Treeview(card, columns=columns, show="headings", height=14)
        for c in columns:
            self.history_tree.heading(c, text=headers[c])
            self.history_tree.column(c, width=130 if c == "waktu" else 100, anchor="center")
        self.history_tree.tag_configure("has_beda", foreground=STATUS_COLOR["BEDA"])
        self.history_tree.pack(fill="both", expand=True, padx=14, pady=(0, 6))
        self.history_tree.bind("<Double-1>", self._load_history_run)

        ctk.CTkLabel(card, text="Klik dua kali satu baris buat lihat detailnya lagi (read-only).",
                     font=FONT_SMALL, text_color=MUTED).pack(anchor="w", padx=14, pady=(0, 14))

    # --- actions ---------------------------------------------------------

    def _open_users(self):
        if self.current_user["role"] != "admin":
            return
        window = ctk.CTkToplevel(self)
        window.title("Kelola User")
        window.geometry("560x500")
        window.transient(self)
        window.grab_set()

        ctk.CTkLabel(window, text="Kelola User", font=FONT_TITLE).pack(anchor="w", padx=18, pady=(18, 8))
        rows_frame = ctk.CTkScrollableFrame(window, height=230)
        rows_frame.pack(fill="both", expand=True, padx=18, pady=(0, 12))

        def refresh():
            for child in rows_frame.winfo_children():
                child.destroy()
            for user_id, username, role, active, _created_at in auth.list_users(self.conn):
                row = ctk.CTkFrame(rows_frame, fg_color="transparent")
                row.pack(fill="x", pady=4)
                label = f"{username}  ·  {role}  ·  {'aktif' if active else 'nonaktif'}"
                ctk.CTkLabel(row, text=label, anchor="w").pack(side="left", fill="x", expand=True)
                if username != self.current_user["username"]:
                    ctk.CTkButton(row, text="Nonaktifkan" if active else "Aktifkan", width=110,
                                  command=lambda uid=user_id, value=not active: (auth.set_user_active(self.conn, uid, value), refresh())).pack(side="right")

        form = ctk.CTkFrame(window)
        form.pack(fill="x", padx=18, pady=(0, 18))
        username = ctk.CTkEntry(form, placeholder_text="Username")
        username.grid(row=0, column=0, padx=8, pady=10, sticky="ew")
        password = ctk.CTkEntry(form, placeholder_text="Password", show="*")
        password.grid(row=0, column=1, padx=8, pady=10, sticky="ew")
        role = ctk.CTkOptionMenu(form, values=["user", "admin"])
        role.grid(row=0, column=2, padx=8, pady=10)
        form.grid_columnconfigure(0, weight=1)
        form.grid_columnconfigure(1, weight=1)

        def add_user():
            try:
                auth.create_user(self.conn, username.get().strip(), password.get(), role.get())
            except Exception as exc:
                messagebox.showerror("User tidak dibuat", str(exc), parent=window)
                return
            username.delete(0, "end")
            password.delete(0, "end")
            refresh()

        ctk.CTkButton(form, text="Tambah User", command=add_user).grid(row=1, column=0, columnspan=3, pady=(0, 10))
        refresh()


    def _pick_folder(self):
        folder = filedialog.askdirectory(title="Pilih folder data (Excel rekap + PDF bulanan)")
        if not folder:
            return
        try:
            validation = mk.validate_data_folder(folder, mk.DEFAULT_PDF_PATTERN)
        except Exception as exc:
            messagebox.showerror("Folder gak valid", str(exc))
            return
        excel_path = validation["excel_path"]
        periods = validation["periods"]
        if not periods:
            messagebox.showerror("Gak ketemu periode", "Excel ketemu tapi gak ada kolom bulan yang terbaca.")
            return

        self.excel_path = excel_path
        self.pdf_dir = Path(folder)
        self.periods = validation["available_periods"]
        self.missing_periods = validation["missing_periods"]
        missing_text = ""
        if self.missing_periods:
            missing_text = f"  • PDF hilang: {', '.join(self.missing_periods)}"
        self.folder_info.configure(
            text=f"Excel: {excel_path.name}   •   Folder: {folder}   •   Periode siap: {len(self.periods)}"
                  f"{missing_text}",
            text_color="#dcdcdc")
        if self.missing_periods:
            messagebox.showwarning(
                "PDF belum lengkap",
                "Periode berikut tidak diproses karena PDF-nya tidak ditemukan:\n\n"
                + "\n".join(self.missing_periods),
            )
        if not self.periods:
            messagebox.showerror("Tidak ada PDF", "Tidak ada PDF yang cocok dengan periode di Excel.")
            return
        self.reload_btn.configure(state="normal")
        self._run_check()

    def _run_check(self):
        if not self.excel_path or not self.pdf_dir:
            return
        self.status_label.configure(text="Memproses...")
        self.reload_btn.configure(state="disabled")
        self.update_idletasks()
        try:
            validation = mk.validate_data_folder(self.pdf_dir, mk.DEFAULT_PDF_PATTERN)
            self.periods = validation["available_periods"]
            self.missing_periods = validation["missing_periods"]
            rows = mk.collect_rows_for_periods(
                self.excel_path, self.pdf_dir, self.periods, mk.DEFAULT_PDF_PATTERN)
        except Exception as exc:
            traceback.print_exc()
            messagebox.showerror("Gagal", f"Gagal membaca file:\n{exc}")
            self.status_label.configure(text="Gagal.")
            self.reload_btn.configure(state="normal")
            return

        self.rows = rows
        self.viewing_history = False
        self._group_by_period(rows)
        self._fill_month_list()
        run_id, counts = save_run(self.conn, self.excel_path, self.pdf_dir, rows)
        self.current_run_id = run_id
        self._update_summary(counts)
        self._refresh_history()
        self.status_label.configure(text=f"Run #{run_id} tersimpan di riwayat.")
        self.reload_btn.configure(state="normal")

        # auto-pilih bulan pertama yang ada BEDA, kalau semua aman pilih bulan pertama
        first_issue = next((p for p in sorted(self.by_period) if summarize(self.by_period[p])["BEDA"] > 0), None)
        target = first_issue or (sorted(self.by_period)[0] if self.by_period else None)
        if target:
            for item in self.month_tree.get_children():
                if self.month_tree.item(item, "values")[0] == self._period_label(target):
                    self.month_tree.selection_set(item)
                    self.month_tree.see(item)
                    break

    def _group_by_period(self, rows):
        self.by_period = {}
        for r in rows:
            if r["status"] == "HANYA_PDF":
                continue
            self.by_period.setdefault(r["period"], []).append(r)

    def _period_label(self, period):
        year, month = period.split("-")
        return f"{mk.MONTH_NAME_ID[int(month) - 1]} {year}"

    def _fill_month_list(self):
        self.month_tree.delete(*self.month_tree.get_children())
        for period in sorted(self.by_period):
            counts = summarize(self.by_period[period])
            has_issue = counts["BEDA"] > 0
            status_text = f"✕ {counts['BEDA']}" if has_issue else "✓"
            self.month_tree.insert("", "end", iid=period, values=(self._period_label(period), status_text))

    def _update_summary(self, counts):
        for status, pill in self.pills.items():
            pill.update_count(counts[status])

    def _on_month_select(self, _event):
        selection = self.month_tree.selection()
        if not selection:
            return
        period = selection[0]
        self.current_period = period
        self.month_title.configure(text=self._period_label(period))
        self._fill_tree(self.tree, self.by_period.get(period, []))

    def _on_category_change(self, category):
        self.current_kind = SHEET_KINDS.get(category)
        if self.current_period:
            self._fill_tree(self.tree, self.by_period.get(self.current_period, []))

    def _fill_tree(self, tree, rows):
        tree.delete(*tree.get_children())
        if self.current_kind:
            rows = [r for r in rows if r["sheet"] == self.current_kind]
        for r in rows:
            tree.insert("", "end", values=(
                SHEET_LABELS.get(r["sheet"], r["sheet"]), r["label"], mk.fmt(r["value"]), mk.fmt(r["pdf_value"]),
                mk.fmt(r["diff"]), f"{STATUS_ICON[r['status']]}  {r['status']}"))

    def _sort_tree(self, tree, col):
        items = [(tree.set(k, col), k) for k in tree.get_children("")]
        reverse = getattr(tree, "_sort_reverse", {}).get(col, False)
        items.sort(key=lambda t: t[0], reverse=reverse)
        for index, (_, k) in enumerate(items):
            tree.move(k, "", index)
        if not hasattr(tree, "_sort_reverse"):
            tree._sort_reverse = {}
        tree._sort_reverse[col] = not reverse

    def _export(self):
        if not self.rows:
            messagebox.showinfo("Belum ada data", "Jalankan cek dulu sebelum menyimpan.")
            return
        path = filedialog.asksaveasfilename(
            title="Simpan hasil", defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")])
        if not path:
            return
        try:
            export_rows([r for r in self.rows if r["status"] != "HANYA_PDF"], path)
        except Exception as exc:
            messagebox.showerror("Gagal simpan", str(exc))
            return
        messagebox.showinfo("Tersimpan", f"Hasil (semua bulan) disimpan ke:\n{path}")

    def _apply_fixes(self):
        if self.viewing_history:
            messagebox.showinfo("Read-only", "Ini tampilan riwayat lama, jalankan cek baru dulu buat menerapkan.")
            return
        beda_rows = [r for r in self.rows if r["status"] == "BEDA"]
        if not beda_rows:
            messagebox.showinfo("Tidak ada BEDA", "Tidak ada baris BEDA buat ditulis.")
            return
        preview = "\n".join(
            f"{r['period']} · {SHEET_LABELS.get(r['sheet'], r['sheet'])} · {r['label']}\n"
            f"  Excel {mk.fmt(r['value'])} → PDF {mk.fmt(r['pdf_value'])}"
            for r in beda_rows[:10]
        )
        if len(beda_rows) > 10:
            preview += f"\n... dan {len(beda_rows) - 10} baris lainnya"
        if not messagebox.askyesno(
                "Preview koreksi Excel",
                f"Akan mengubah {len(beda_rows)} cell berdasarkan nilai PDF.\n"
                "Cell formula dilewati otomatis. Backup bertimestamp dibuat sebelum menyimpan.\n\n"
                + preview + "\n\nLanjutkan?", parent=self):
            return
        try:
            written, skipped, backup_path = mk.apply_fixes(self.excel_path, beda_rows)
        except Exception as exc:
            messagebox.showerror("Gagal", str(exc))
            return
        msg = f"{written} cell ditulis ulang."
        if backup_path:
            msg += f"\nBackup: {backup_path}"
        if skipped:
            msg += "\n\nDilewati (cell formula), cek manual:\n" + "\n".join(
                f"- {label}: {formula}" for label, formula in skipped)
        if written:
            save_corrections(
                self.conn, self.current_run_id, self.current_user["username"], beda_rows,
                "excel", skipped_labels=[label for label, _formula in skipped])
        messagebox.showinfo("Selesai", msg)

    def _make_corrected_pdfs(self):
        if self.viewing_history:
            messagebox.showinfo("Read-only", "Jalankan cek baru dulu untuk membuat PDF koreksi.", parent=self)
            return
        beda_rows = [r for r in self.rows if r["status"] == "BEDA"]
        if not beda_rows:
            messagebox.showinfo("Tidak ada BEDA", "Tidak ada baris BEDA untuk dikoreksi.", parent=self)
            return
        periods = sorted({r["period"] for r in beda_rows})
        if not messagebox.askyesno(
                "Buat PDF koreksi",
                f"Akan membuat {len(periods)} PDF baru berdasarkan nilai Excel.\n"
                "PDF asli tidak akan diubah. Lanjutkan?", parent=self):
            return
        outputs, skipped = [], []
        for period in periods:
            year_str, month_str = period.split("-")
            source = mk.month_pdf_path(int(month_str) - 1, self.pdf_dir, int(year_str), mk.DEFAULT_PDF_PATTERN)
            if source is None:
                skipped.append((period, "PDF sumber tidak ditemukan"))
                continue
            output = CORRECTED_PDF_DIR / f"{source.stem}_corrected{source.suffix}"
            try:
                written, missed = mk.create_corrected_pdf(
                    source, [r for r in beda_rows if r["period"] == period], output)
            except Exception as exc:
                skipped.append((period, str(exc)))
                continue
            outputs.append((output, written))
            skipped_labels = [label for label, _reason in missed]
            save_corrections(
                self.conn, self.current_run_id, self.current_user["username"],
                [r for r in beda_rows if r["period"] == period], "pdf",
                output_file=str(output), skipped_labels=skipped_labels)
            skipped.extend((f"{period} · {label}", reason) for label, reason in missed)
        msg = "\n".join(f"{path.name}: {count} angka diganti" for path, count in outputs)
        if not msg:
            msg = "Tidak ada PDF yang berhasil dibuat."
        if skipped:
            msg += "\n\nDilewati:\n" + "\n".join(f"- {label}: {reason}" for label, reason in skipped)
        messagebox.showinfo("PDF koreksi selesai", msg, parent=self)

    def _refresh_history(self):
        self.history_tree.delete(*self.history_tree.get_children())
        cur = self.conn.execute(
            "SELECT id, ts, cocok, beda, hanya_excel FROM runs ORDER BY id DESC")
        for row in cur.fetchall():
            self.history_tree.insert("", "end", iid=str(row[0]), values=row,
                                      tags=("has_beda",) if row[3] > 0 else ())

    def _load_history_run(self, _event):
        selection = self.history_tree.selection()
        if not selection:
            return
        run_id = selection[0]
        cur = self.conn.execute(
            "SELECT period, bulan, sheet, label, excel_value, pdf_value, selisih, status "
            "FROM run_rows WHERE run_id=?", (run_id,))
        rows = [{"period": r[0], "month": r[1], "sheet": r[2], "label": r[3], "value": r[4],
                 "pdf_value": r[5], "diff": r[6], "status": r[7]} for r in cur.fetchall()]
        self.rows = rows
        self.viewing_history = True
        self._group_by_period(rows)
        self._fill_month_list()
        self._update_summary(summarize(rows))
        self.tabs.set("Hasil per Bulan")
        self.status_label.configure(text=f"Menampilkan riwayat run #{run_id} (read-only, 'Terapkan' nonaktif).")


def _prompt_credentials(title, setup=False):
    result = {"value": None}
    ctk.set_appearance_mode("light")
    window = ctk.CTk()
    window.withdraw()
    dialog = ctk.CTkToplevel(window)
    dialog.title(title)
    dialog.geometry("920x560")
    dialog.minsize(760, 500)
    dialog.resizable(False, False)
    dialog.protocol("WM_DELETE_WINDOW", lambda: (dialog.destroy(), window.destroy()))
    dialog.grab_set()

    dialog.grid_columnconfigure(0, weight=2)
    dialog.grid_columnconfigure(1, weight=3)
    dialog.grid_rowconfigure(0, weight=1)

    brand = ctk.CTkFrame(dialog, fg_color="#405CF5", corner_radius=0)
    brand.grid(row=0, column=0, sticky="nsew")
    ctk.CTkLabel(brand, text="MK", width=58, height=58, corner_radius=16,
                 fg_color="#ffffff", text_color="#405CF5",
                 font=("Segoe UI", 22, "bold")).pack(anchor="w", padx=34, pady=(42, 28))
    ctk.CTkLabel(brand, text="Mapping\nKeuangan", text_color="#ffffff",
                 font=("Segoe UI", 27, "bold"), justify="left", anchor="w").pack(
                     anchor="w", padx=34)
    ctk.CTkLabel(brand, text="Audit laporan keuangan\ndengan lebih mudah dan terkontrol.",
                 text_color="#E7EBFF", font=("Segoe UI", 14), justify="left", anchor="w").pack(
                     anchor="w", padx=34, pady=(14, 0))
    art = ctk.CTkFrame(brand, fg_color="transparent")
    art.pack(side="bottom", anchor="w", padx=34, pady=38)
    for width, height, color, xpad in (
            (130, 18, "#9EAFFF", 0), (90, 18, "#D8DEFF", 18), (160, 18, "#263CC1", 42)):
        ctk.CTkFrame(art, width=width, height=height, fg_color=color, corner_radius=9).pack(
            anchor="w", pady=6, padx=(xpad, 0))

    form = ctk.CTkFrame(dialog, fg_color="#FFFFFF", corner_radius=0)
    form.grid(row=0, column=1, sticky="nsew", padx=0)
    form.grid_columnconfigure(0, weight=1)
    form.grid_rowconfigure(5, weight=1)
    content = ctk.CTkFrame(form, fg_color="transparent", width=360)
    content.grid(row=0, column=0, sticky="n", padx=64, pady=(86, 30))
    ctk.CTkLabel(content, text="Buat admin pertama" if setup else "Selamat datang kembali",
                 text_color="#18213D", font=("Segoe UI", 25, "bold"), anchor="w").pack(
                     fill="x", pady=(0, 8))
    ctk.CTkLabel(content, text="Siapkan akun untuk mulai menggunakan aplikasi." if setup
                 else "Masuk untuk melanjutkan ke dashboard.", text_color="#68708A",
                 font=("Segoe UI", 13), anchor="w").pack(fill="x", pady=(0, 28))
    ctk.CTkLabel(content, text="Username", text_color="#18213D",
                 font=("Segoe UI", 12, "bold"), anchor="w").pack(fill="x", pady=(0, 7))
    username = ctk.CTkEntry(content, height=46, placeholder_text="Masukkan username",
                            fg_color="#F4F7FF", border_color="#D8DEEF", text_color="#18213D",
                            placeholder_text_color="#9AA3BB", corner_radius=10)
    username.pack(fill="x", pady=(0, 18))
    ctk.CTkLabel(content, text="Password", text_color="#18213D",
                 font=("Segoe UI", 12, "bold"), anchor="w").pack(fill="x", pady=(0, 7))
    password = ctk.CTkEntry(content, height=46, placeholder_text="Masukkan password", show="*",
                            fg_color="#F4F7FF", border_color="#D8DEEF", text_color="#18213D",
                            placeholder_text_color="#9AA3BB", corner_radius=10)
    password.pack(fill="x", pady=(0, 4))
    confirm = None
    if setup:
        ctk.CTkLabel(content, text="Ulangi password", text_color="#18213D",
                     font=("Segoe UI", 12, "bold"), anchor="w").pack(fill="x", pady=(14, 7))
        confirm = ctk.CTkEntry(content, height=46, placeholder_text="Ulangi password", show="*",
                               fg_color="#F4F7FF", border_color="#D8DEEF", text_color="#18213D",
                               placeholder_text_color="#9AA3BB", corner_radius=10)
        confirm.pack(fill="x", pady=(0, 4))
    error = ctk.CTkLabel(content, text="", text_color="#C43D4B", anchor="w",
                         font=("Segoe UI", 11))
    error.pack(fill="x", pady=(8, 0))

    def submit():
        if setup and password.get() != confirm.get():
            error.configure(text="Password tidak sama.")
            return
        result["value"] = (username.get().strip(), password.get())
        dialog.destroy()

    ctk.CTkButton(content, text="Buat Admin" if setup else "Masuk", height=46,
                  fg_color="#405CF5", hover_color="#3049D4", corner_radius=10,
                  font=("Segoe UI", 13, "bold"), command=submit).pack(fill="x", pady=(18, 0))
    dialog.bind("<Return>", lambda _event: submit())
    username.focus_set()
    window.wait_window(dialog)
    if window.winfo_exists():
        window.destroy()
    return result["value"]


def _login(conn):
    if auth.user_count(conn) == 0:
        credentials = _prompt_credentials("Buat Admin Pertama", setup=True)
        if not credentials:
            return None
        try:
            auth.create_user(conn, credentials[0], credentials[1], "admin")
        except Exception as exc:
            messagebox.showerror("Admin tidak dibuat", str(exc))
            return None

    while True:
        credentials = _prompt_credentials("Login Mapping Keuangan")
        if not credentials:
            return None
        user = auth.authenticate(conn, credentials[0], credentials[1])
        if user:
            return user
        messagebox.showerror("Login gagal", "Username atau password salah, atau user tidak aktif.")


def main():
    conn = ensure_db()
    user = _login(conn)
    conn.close()
    if not user:
        return
    app = App(user)
    app.mainloop()


if __name__ == "__main__":
    main()
