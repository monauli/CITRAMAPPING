"""Validasi Excel rekap keuangan bulanan vs PDF laporan sumber (CAM JO).

Dry-run (default): cetak baris yang beda (BEDA), yang cuma ada di satu sisi,
dan ringkasan jumlah per bulan/sheet.

--apply: tulis nilai PDF ke cell Excel untuk baris BEDA yang berupa angka
polos (bukan formula). Sebelum menulis, bikin backup *.bak.xlsx sekali.
"""
import argparse
import json
import os
import re
import shutil
import sys
import tempfile
from datetime import datetime
from pathlib import Path

import openpyxl
import pdfplumber

# exe hasil PyInstaller mengekstrak skrip ke folder temp — path default harus
# ikut lokasi exe/skrip aslinya, bukan folder temp itu.
BASE_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
DATA_DIR = BASE_DIR / "dataexcelpdfreportlapkeucamjoperiodejanuarisdag"
EXCEL_PATH = DATA_DIR / "Laporan Keuangan Bulanan Versi Ciputra Januari - Agustus 2026.xlsx"

MONTH_NAME_ID = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus",
                 "September", "Oktober", "November", "Desember"]

SHEET_KIND = {"Balance Sheet": "balance-sheet", "Profit & Loss": "profit-loss", "Cashflow": "cashflow"}

TOLERANCE = 1.0


# --- Excel ---------------------------------------------------------------

def find_header_row(ws):
    for r in range(1, 13):
        if ws.cell(r, 2).value == "Description":
            return r
    raise ValueError(f'Baris header "Description" tidak ketemu di sheet "{ws.title}"')


def month_columns(ws, header_row):
    columns = {}
    for c in range(3, ws.max_column + 1):
        val = ws.cell(header_row, c).value
        if not isinstance(val, str):
            continue
        try:
            d = datetime.strptime(val.strip(), "%d %B %Y")
        except ValueError:
            continue
        columns[f"{d.year}-{d.month:02d}"] = c
    return columns


def label_and_bold(ws, row, label_end_col):
    """Gabungkan teks label lintas kolom B..label_end_col-1.

    Sheet Cashflow menaruh label di kolom berbeda tergantung level indentasi
    (CASH INFLOW di B, Collection: di C, Residential di D, -/- Subsidy di D+E)
    — beda dengan Balance Sheet/P&L yang selalu di kolom B.
    """
    parts, bold = [], False
    for c in range(2, label_end_col):
        v = ws.cell(row, c).value
        if isinstance(v, str) and v.strip():
            parts.append(v.strip())
            if ws.cell(row, c).font.bold:
                bold = True
    return " ".join(parts), bold


def classify_excel_column(ws, header_row, col, label_end_col):
    """Baris -> (label, kind, value). kind: detail/subtotal/header.

    Bold + ada angka di cell = subtotal (termasuk grand-total yang menjumlah
    subtotal lain, bukan cuma baris detail langsung — mis. TOTAL ASSETS
    menjumlah TOTAL CURRENT ASSETS + TOTAL NON-CURRENT ASSETS, bukan baris
    detail). Bold + cell kosong = section banner murni (dibuang). Label
    "TOTAL"/"JUMLAH" TIDAK dipakai sebagai syarat — banyak baris total nyata
    di sheet ini gak pakai kata itu sama sekali (mis. "SURPLUS / (DEFICIT)
    Operational", "NET CASH FLOW").
    """
    lines = []
    for r in range(header_row + 1, ws.max_row + 1):
        label, bold = label_and_bold(ws, r, label_end_col)
        if not label:
            continue
        raw = ws.cell(r, col).value
        value = raw if isinstance(raw, (int, float)) else None
        if not bold:
            lines.append({"row": r, "label": label, "kind": "detail", "value": value if value is not None else 0.0})
        elif value is not None:
            lines.append({"row": r, "label": label, "kind": "subtotal", "value": value})
        else:
            lines.append({"row": r, "label": label, "kind": "header", "value": None})
    return lines


def load_excel_month(path, period):
    """period: 'YYYY-MM'. Returns {sheet_kind: [line, ...]} (hanya baris detail/subtotal)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    out = {}
    for sheet_name, kind in SHEET_KIND.items():
        ws = wb[sheet_name]
        header_row = find_header_row(ws)
        cols = month_columns(ws, header_row)
        if period not in cols:
            continue
        label_end_col = min(cols.values())
        lines = classify_excel_column(ws, header_row, cols[period], label_end_col)
        out[kind] = [l for l in lines if l["kind"] in ("detail", "subtotal")]
    return out


# --- PDF -------------------------------------------------------------------

PAGE_TITLE_KIND = [
    (re.compile(r"financial position", re.I), "balance-sheet"),
    (re.compile(r"profit or loss", re.I), "profit-loss"),
    (re.compile(r"cashflow", re.I), "cashflow"),
]

# Fragmen yang aman digabung lintas-kata: hasil kerning PDF memecah satu angka
# jadi beberapa "word" tanpa spasi asli di antaranya (lihat README di plan).
NUMERIC_CHARS = re.compile(r"^[0-9.,()%\-]+$")
NUMERIC_VALUE = re.compile(r"^\(?-?[0-9][0-9.,]*\)?%?$")
# Baris judul/tanggal halaman ("As of January 31, 2026", "For the Year Ended
# ...") ikut ditangkap kalau tidak dibuang di sini — angka tanggalnya terbaca
# seolah nilai akun.
JUNK_LINE = re.compile(r"^(as of|for the|project\s*:|statement|in rp\.?)", re.I)


def merge_letter_fragments(words):
    """Gabungkan huruf tunggal berurutan ("Y" "E" "A" "R") jadi satu kata.

    Beberapa laporan nyetak "YEAR" dengan tiap huruf sebagai word terpisah
    (artefak layout si exporter, bukan OCR) — tanpa ini baris "TOTAL
    COMPREHENSIVE INCOME ... YEAR" gak pernah ketemu pasangannya karena
    labelnya kepotong jadi "... Y".
    """
    ordered = sorted(words, key=lambda w: w["x0"])
    merged, i = [], 0
    while i < len(ordered):
        w = ordered[i]
        if len(w["text"]) == 1 and w["text"].isalpha():
            run = [w]
            j = i + 1
            while j < len(ordered) and len(ordered[j]["text"]) == 1 and ordered[j]["text"].isalpha():
                run.append(ordered[j])
                j += 1
            if len(run) >= 2:
                merged.append({"x0": run[0]["x0"], "x1": run[-1]["x1"], "text": "".join(r["text"] for r in run)})
                i = j
                continue
        merged.append(w)
        i += 1
    return merged


def merge_number_fragments(words):
    """Gabungkan word yang berdempetan (gap kecil) dan sama-sama berisi karakter angka."""
    merged = []
    for w in sorted(words, key=lambda w: w["x0"]):
        if merged and w["x0"] - merged[-1]["x1"] < 1.0 \
                and NUMERIC_CHARS.match(merged[-1]["text"]) and NUMERIC_CHARS.match(w["text"]):
            merged[-1] = {
                "x0": merged[-1]["x0"],
                "x1": w["x1"],
                "top": min(merged[-1].get("top", w.get("top", 0)), w.get("top", 0)),
                "bottom": max(merged[-1].get("bottom", w.get("bottom", 0)), w.get("bottom", 0)),
                "text": merged[-1]["text"] + w["text"],
            }
        else:
            merged.append(dict(w))
    return merged


def group_lines(words):
    lines = []
    for w in sorted(words, key=lambda w: w["top"]):
        if lines and abs(w["top"] - lines[-1][0]) <= 2:
            lines[-1][1].append(w)
        else:
            lines.append([w["top"], [w]])
    return [ws for _, ws in lines]


def parse_number(token):
    neg = token.startswith("(") and token.endswith(")")
    cleaned = token.strip("()%").replace(",", "")
    if cleaned in ("", "-"):
        return 0.0
    value = float(cleaned)
    return -value if neg else value


def parse_pdf_line(tokens, with_position=False):
    """tokens: list of merged word dicts untuk satu baris. -> (label, value) atau None."""
    # Kerning PDF sometimes attaches the currency marker to a negative amount:
    # "(R820,199,808)". Remove it only when followed by digits.
    tokens = [dict(token) for token in tokens]
    for i, token in enumerate(tokens):
        raw = token["text"]
        if i and tokens[i - 1]["text"].upper() == "YEA" and re.match(r"^\(?R(?=-?\d)", raw, re.I):
            # The final R of YEAR is glued to the amount token in this layout.
            tokens[i - 1]["text"] += "R"
        token["text"] = re.sub(r"^(\()R(?:p)?(?=-?\d)", r"\1", raw, flags=re.I)
        token["text"] = re.sub(r"^R(?=-?\d)", "", token["text"], flags=re.I)
    # "PSAK 72" itu nama akun (referensi standar akuntansi), bukan label+angka
    # — tanpa ini "72" kebaca sebagai nilai kolom pertama, label kepotong jadi
    # cuma "PSAK".
    search_from = 2 if (len(tokens) >= 2 and tokens[0]["text"].upper() == "PSAK"
                         and re.fullmatch(r"\d{1,3}", tokens[1]["text"])) else 0
    # Token "-" tunggal bisa berarti dua hal: nihil (nol) di kolom nilai, ATAU
    # cuma pemisah kata di dalam label itu sendiri (mis. "Afiliasi - Net").
    # Dianggap awal kolom nilai HANYA kalau SISA baris dari situ murni
    # angka/persen/dash — kalau masih ada kata (huruf) sesudahnya, itu masih
    # bagian label, bukan placeholder nihil.
    def is_value_region(idx):
        return all(t["text"] == "-" or "%" in t["text"] or NUMERIC_VALUE.match(t["text"])
                   for t in tokens[idx:])

    numeric_idx = next((i for i in range(search_from, len(tokens))
                         if NUMERIC_VALUE.match(tokens[i]["text"])
                         or (tokens[i]["text"] == "-" and is_value_region(i))), None)
    if numeric_idx is None:
        return None
    label = " ".join(t["text"] for t in tokens[:numeric_idx]).strip()
    if not label or JUNK_LINE.match(label):
        return None
    amount_word = next((t for t in tokens[numeric_idx:] if "%" not in t["text"]), None)
    amount_token = amount_word["text"] if amount_word else None
    if amount_token is None:
        return None
    if with_position:
        return label, parse_number(amount_token), amount_word
    return label, parse_number(amount_token)


def load_pdf_month(path):
    """Returns {sheet_kind: {normalized_label: [values]}} without dropping duplicate labels."""
    out = {}
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            kind = next((k for pat, k in PAGE_TITLE_KIND if pat.search(text)), None)
            if kind is None:
                continue
            page_lines = {}
            for line_words in group_lines(page.extract_words()):
                merged = merge_number_fragments(merge_letter_fragments(line_words))
                parsed = parse_pdf_line(merged)
                if parsed is None:
                    continue
                label, value = parsed
                page_lines.setdefault(normalize_label(label), []).append(value)
            out[kind] = page_lines
    return out


def load_pdf_records(path):
    """Ambil baris PDF beserta halaman dan koordinat angka untuk overlay koreksi."""
    records = []
    with pdfplumber.open(path) as pdf:
        for page_number, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            kind = next((k for pat, k in PAGE_TITLE_KIND if pat.search(text)), None)
            if kind is None:
                continue
            for line_words in group_lines(page.extract_words()):
                merged = merge_number_fragments(merge_letter_fragments(line_words))
                parsed = parse_pdf_line(merged, with_position=True)
                if parsed is None:
                    continue
                label, value, amount_word = parsed
                records.append({
                    "kind": kind,
                    "label": label,
                    "value": value,
                    "page": page_number,
                    "x0": amount_word["x0"],
                    "x1": amount_word["x1"],
                    "top": amount_word["top"],
                    "bottom": amount_word["bottom"],
                })
    return records


# --- Compare -----------------------------------------------------------------

def normalize_label(s):
    text = re.sub(r"\s+", " ", s.strip().lower()).rstrip(":").strip()
    # Exporter PDF suka nambahin akhiran "- net" yang gak ada di Excel
    # ("Estate Income" vs "estate income - net") — bukan beda data.
    return re.sub(r"\s*-\s*net$", "", text).strip()


# Pasangan label yang SUDAH diverifikasi manual (nilai sama, istilah beda) di
# data Citraland Megah Batam — bukan tebakan/aturan umum. Kalau bulan lain
# atau project lain nemu pasangan baru, tambah di sini.
ALIASES = [
    # Cashflow
    ("Biaya Management Fee", "Management Fee (khusus JO/JV)"),
    ("Biaya Royalty", "Royalty Fee (khusus JO/JV)"),
    ("Received from affiliated", "Received From/(Payment to) Afiliasi - net"),
    ("-/- Subsidy / KPR Fee", "-/- Subsidi / KPR Fee"),
    ("SURPLUS / (DEFICIT) Before Interest", "Surplus / (Deficit) Before Interest Expense"),
    ("SURPLUS / (DEFICIT) After Interes", "Surplus / (Deficit) After Interest"),
    ("CASH & BANK", "Cash/Bank"),
    ("DEPOSITO", "Bank & Deposit"),
    ("Penempatan (Pencairan) Deposito", "Placement (Proceed) Deposit"),
    # Balance Sheet
    ("Long-term employess benefit liability", "Long-term employee benefit liability"),
    ("Accumulated actuarial losses on Employee Benefit", "Accumulated actuarial losses on Emp Ben"),
    # Cashflow (investasi)
    ("Fixed Assets / Investment Property", "Fixed Assets / Investment Properti"),
    # Profit & Loss
    ("REVENUES", "Revenue From Apartment/House"),
    ("REVENUES TOTAL", "REVENUES"),
    ("SUBSIDI", "Return/Sales Discount"),
    ("Manajemen Fee & Royalty", "Management Fee & Royalty"),
    ("PROFIT FROM OPERATIONS", "Profit (Loss) From Operations"),
    ("PROFIT BEFORE INCOME TAX", "Profit (Loss) Before Income Tax"),
    ("PROFIT FOR THE PERIOD", "Profit (Loss) For The Period"),
    ("TOTAL COMPREHENSIVE INCOME FOR THE PERIOD", "Total Comprehensive Income (Loss) For The Year"),
]
ALIAS_MAP = {normalize_label(excel): normalize_label(pdf) for excel, pdf in ALIASES}


def _pdf_key_for_excel_label(label):
    normalized = normalize_label(label)
    return ALIAS_MAP.get(normalized, normalized)


def _pdf_number(value):
    if value is None:
        return "-"
    number = f"{abs(value):,.2f}".rstrip("0").rstrip(".")
    return f"({number})" if value < 0 else number


def create_corrected_pdf(pdf_path, beda_rows, output_path):
    """Buat PDF baru dengan angka BEDA diganti mengikuti nilai Excel."""
    import fitz

    pdf_path = Path(pdf_path)
    output_path = Path(output_path)
    if pdf_path.resolve() == output_path.resolve():
        raise ValueError("PDF koreksi harus memakai file output terpisah dari PDF asli.")
    records = load_pdf_records(pdf_path)
    document = fitz.open(pdf_path)
    used = set()
    replacements = []
    skipped = []
    for row in beda_rows:
        if row.get("value") is None or row.get("pdf_value") is None:
            skipped.append((row.get("label", "?"), "nilai kosong"))
            continue
        key = _pdf_key_for_excel_label(row["label"])
        candidates = [
            (index, record) for index, record in enumerate(records)
            if index not in used
            and record["kind"] == row["sheet"]
            and normalize_label(record["label"]) == key
            and abs(record["value"] - row["pdf_value"]) <= TOLERANCE
        ]
        if not candidates:
            skipped.append((row["label"], "koordinat PDF tidak ditemukan"))
            continue
        index, record = candidates[0]
        used.add(index)
        replacements.append((record, _pdf_number(row["value"])))

    by_page = {}
    for record, replacement in replacements:
        by_page.setdefault(record["page"], []).append((record, replacement))
    for page_number, page_replacements in by_page.items():
        page = document[page_number]
        for record, _replacement in page_replacements:
            rect = fitz.Rect(record["x0"] - 1, record["top"] - 1,
                             record["x1"] + 1, record["bottom"] + 1)
            page.add_redact_annot(rect, fill=(1, 1, 1))
        page.apply_redactions()
        for record, replacement in page_replacements:
            font_size = max(6, min(12, record["bottom"] - record["top"] + 2))
            available_width = max(20, record["x1"] - record["x0"])
            while font_size > 6 and fitz.get_text_length(replacement, fontname="helv", fontsize=font_size) > available_width:
                font_size -= 0.25
            text_width = fitz.get_text_length(replacement, fontname="helv", fontsize=font_size)
            page.insert_text(
                (record["x1"] - text_width, record["top"] + font_size - 1),
                replacement,
                fontsize=font_size,
                fontname="helv",
                color=(0, 0, 0),
            )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path, garbage=4, deflate=True)
    document.close()
    return len(replacements), skipped


def compare(excel_lines, pdf_by_label):
    """Pair each Excel row with one closest PDF value for the same label."""
    def match_key(label):
        key = normalize_label(label)
        # Alias diperiksa DULUAN: kadang teks ter-normalisasi excel KEBETULAN
        # sama persis dengan label PDF yang arti aslinya beda (mis. Excel
        # "REVENUES" = gross, sementara PDF juga punya baris "REVENUES" tapi
        # itu versi net — alias eksplisit (hasil verifikasi manual) harus
        # menang daripada kecocokan teks yang kebetulan tapi salah makna.
        aliased = ALIAS_MAP.get(key)
        if aliased in pdf_by_label:
            return aliased
        return key if key in pdf_by_label else None

    groups = {}
    rows = []
    for line in excel_lines:
        key = match_key(line["label"])
        if key is None:
            rows.append({**line, "pdf_value": None, "diff": None, "status": "HANYA_EXCEL"})
        else:
            groups.setdefault(key, []).append(line)

    for pdf_key, candidates in groups.items():
        pdf_values = pdf_by_label[pdf_key]
        # ponytail: greedy nearest pairs; switch to minimum-cost matching if duplicate labels become ambiguous.
        pairs = sorted(
            (abs((line["value"] or 0) - value), i, j)
            for i, line in enumerate(candidates)
            for j, value in enumerate(pdf_values)
        )
        matched_excel, matched_pdf = set(), set()
        for _, i, j in pairs:
            if i not in matched_excel and j not in matched_pdf:
                matched_excel.add(i)
                matched_pdf.add(j)
                line, pdf_value = candidates[i], pdf_values[j]
                diff = line["value"] - pdf_value
                status = "COCOK" if abs(diff) <= TOLERANCE else "BEDA"
                rows.append({**line, "pdf_value": pdf_value, "diff": diff, "status": status})
        for i, line in enumerate(candidates):
            if i not in matched_excel:
                rows.append({**line, "pdf_value": None, "diff": None, "status": "HANYA_EXCEL"})
        for j, value in enumerate(pdf_values):
            if j not in matched_pdf:
                rows.append({"row": None, "label": pdf_key, "kind": "detail", "value": None,
                             "pdf_value": value, "diff": None, "status": "HANYA_PDF"})

    for key, values in pdf_by_label.items():
        if key not in groups:
            for value in values:
                rows.append({"row": None, "label": key, "kind": "detail", "value": None,
                             "pdf_value": value, "diff": None, "status": "HANYA_PDF"})

    # Baris yang cuma ada di satu sisi DAN nilainya nol bukan beda data —
    # biasanya baris label pembantu (mis. "Collection:") atau akun nihil yang
    # cuma dicetak di salah satu sisi. Buang supaya ringkasan tidak berisik.
    return [r for r in rows if not (
        r["status"] in ("HANYA_EXCEL", "HANYA_PDF") and (r["value"] or r["pdf_value"] or 0) == 0
    )]


# --- Report / apply ----------------------------------------------------------

def fmt(n):
    return "-" if n is None else f"{n:,.2f}"


def month_pdf_path(month_index, pdf_dir, year, pattern):
    name = pattern.format(month=MONTH_NAME_ID[month_index], year=year)
    path = pdf_dir / name
    return path if path.exists() else None


def find_excel_in_folder(folder):
    """Satu-satunya .xlsx di folder (selain backup/lock file). Raise kalau 0 atau >1 ketemu."""
    candidates = [p for p in Path(folder).glob("*.xlsx")
                  if not p.name.endswith(".bak.xlsx") and not p.name.startswith("~$")]
    if not candidates:
        raise FileNotFoundError(f"Gak ada file .xlsx di folder: {folder}")
    if len(candidates) > 1:
        names = ", ".join(p.name for p in candidates)
        raise ValueError(f"Ada {len(candidates)} file .xlsx di folder itu, harus cuma 1: {names}")
    return candidates[0]


def detect_periods(excel_path):
    """List 'YYYY-MM' yang ketemu di Excel (gabungan 3 sheet), urut."""
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    periods = set()
    for sheet_name in SHEET_KIND:
        ws = wb[sheet_name]
        header_row = find_header_row(ws)
        periods.update(month_columns(ws, header_row).keys())
    return sorted(periods)


def validate_data_folder(folder, pdf_pattern=None):
    """Validasi satu folder input dan kembalikan periode tersedia/missing."""
    if pdf_pattern is None:
        pdf_pattern = DEFAULT_PDF_PATTERN
    folder = Path(folder)
    excel_path = find_excel_in_folder(folder)
    periods = detect_periods(excel_path)
    missing_periods = []
    available_periods = []
    for period in periods:
        year_str, month_str = period.split("-")
        pdf_path = month_pdf_path(int(month_str) - 1, folder, int(year_str), pdf_pattern)
        (available_periods if pdf_path else missing_periods).append(period)
    return {
        "folder": folder,
        "excel_path": excel_path,
        "periods": periods,
        "available_periods": available_periods,
        "missing_periods": missing_periods,
    }


def collect_rows_for_periods(excel_path, pdf_dir, periods, pdf_pattern):
    """List baris lintas periode ('YYYY-MM') + sheet yang filenya ketemu.

    Satu sumber dipakai CLI (run_report, lewat collect_rows) dan GUI, supaya
    logic loop bulan/sheet-nya gak dobel ditulis dan gak bisa menyimpang.
    """
    all_rows = []
    for period in periods:
        year_str, month_str = period.split("-")
        year, month_index = int(year_str), int(month_str) - 1
        pdf_path = month_pdf_path(month_index, pdf_dir, year, pdf_pattern)
        if pdf_path is None:
            continue
        excel_by_kind = load_excel_month(excel_path, period)
        pdf_by_kind = load_pdf_month(pdf_path)
        for kind, excel_lines in excel_by_kind.items():
            rows = compare(excel_lines, pdf_by_kind.get(kind, {}))
            for r in rows:
                all_rows.append({**r, "month": MONTH_NAME_ID[month_index], "period": period, "sheet": kind})
    return all_rows


def collect_rows(excel_path, pdf_dir, year, pdf_pattern):
    """Kompatibilitas CLI lama: satu tahun, 12 bulan."""
    periods = [f"{year}-{m:02d}" for m in range(1, 13)]
    return collect_rows_for_periods(excel_path, pdf_dir, periods, pdf_pattern)


def apply_fixes(excel_path, beda_rows):
    """Tulis nilai PDF ke cell Excel untuk tiap baris BEDA (punya 'period'/'sheet'/'row'/'pdf_value').

    Cell berisi formula dilewati (dilaporkan, tidak ditimpa). Backup *.bak.xlsx
    dibuat sekali sebelum tulisan pertama. Returns (written, skipped, backup_path).
    skipped: list of (label, formula_text).
    """
    wb = openpyxl.load_workbook(excel_path, data_only=False)
    backup_path = None
    temp_path = None
    written = 0
    skipped = []
    for r in beda_rows:
        sheet_name = next(name for name, k in SHEET_KIND.items() if k == r["sheet"])
        ws = wb[sheet_name]
        header_row = find_header_row(ws)
        col = month_columns(ws, header_row).get(r["period"])
        if col is None:
            continue
        cell = ws.cell(r["row"], col)
        if isinstance(cell.value, str) and cell.value.startswith("="):
            skipped.append((r["label"], cell.value))
            continue
        if backup_path is None:
            stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
            backup_path = excel_path.with_name(f"{excel_path.stem}.backup-{stamp}{excel_path.suffix}")
            shutil.copy2(excel_path, backup_path)
        cell.value = r["pdf_value"]
        written += 1
    if written:
        fd, temp_name = tempfile.mkstemp(prefix=f".{excel_path.stem}-", suffix=".xlsx", dir=excel_path.parent)
        os.close(fd)
        temp_path = Path(temp_name)
        try:
            wb.save(temp_path)
            os.replace(temp_path, excel_path)
        finally:
            if temp_path.exists():
                temp_path.unlink()
    return written, skipped, backup_path


def run_report(excel_path, pdf_dir, year, pdf_pattern, apply_changes):
    rows = collect_rows(excel_path, pdf_dir, year, pdf_pattern)
    total_summary = {"COCOK": 0, "BEDA": 0, "HANYA_EXCEL": 0, "HANYA_PDF": 0}
    groups = {}
    order = []
    for r in rows:
        key = (r["month"], r["sheet"])
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(r)
        total_summary[r["status"]] += 1

    all_beda = []
    for month, sheet in order:
        group_rows = groups[(month, sheet)]
        beda_rows = [r for r in group_rows if r["status"] == "BEDA"]
        all_beda.extend(beda_rows)
        print(f"\n=== {month} {year} / {sheet} ===")
        counts = {s: sum(1 for r in group_rows if r["status"] == s) for s in total_summary}
        print(f"  COCOK={counts['COCOK']} BEDA={counts['BEDA']} "
              f"HANYA_EXCEL={counts['HANYA_EXCEL']} HANYA_PDF={counts['HANYA_PDF']}")
        for r in beda_rows:
            print(f"  BEDA  {r['label']!r}: excel={fmt(r['value'])}  pdf={fmt(r['pdf_value'])}  "
                  f"selisih={fmt(r['diff'])}")

    print(f"\nTOTAL: {total_summary}")
    if apply_changes and all_beda:
        written, skipped, backup_path = apply_fixes(excel_path, all_beda)
        for label, formula in skipped:
            print(f"  SKIP tulis {label!r}: cell berisi formula ({formula}), cek manual dulu.")
        if written:
            print(f"\nBackup dibuat: {backup_path}")
            print(f"Excel diperbarui ({written} cell): {excel_path}")
        else:
            print("Tidak ada perubahan ditulis (semua BEDA ada di cell formula).")
    elif apply_changes:
        print("Tidak ada perubahan ditulis (tidak ada baris BEDA).")


def _selftest():
    """Cek temuan yang sudah divalidasi manual: Cash and cash equivalents Jan 2026 BEDA ~319,372,726.38.

    Cuma jalan kalau pakai data default (file Citraland bawaan folder ini) —
    temuan ini spesifik ke fixture itu, tidak berlaku untuk file lain.
    """
    excel_by_kind = load_excel_month(EXCEL_PATH, "2026-01")
    pdf_by_kind = load_pdf_month(month_pdf_path(0, DATA_DIR, 2026, DEFAULT_PDF_PATTERN))
    rows = compare(excel_by_kind["balance-sheet"], pdf_by_kind["balance-sheet"])
    cash = next(r for r in rows if normalize_label(r["label"]) == "cash and cash equivalents")
    assert cash["status"] == "BEDA", cash
    assert abs(cash["diff"] - (-319372726.38)) < 1, cash
    print("selftest OK: Cash and cash equivalents Jan 2026 terdeteksi BEDA, selisih", fmt(cash["diff"]))

    march_excel = load_excel_month(EXCEL_PATH, "2026-03")["profit-loss"]
    march_pdf = load_pdf_month(month_pdf_path(2, DATA_DIR, 2026, DEFAULT_PDF_PATTERN))["profit-loss"]
    march_rows = compare(march_excel, march_pdf)
    psak_rows = [r for r in march_rows if normalize_label(r["label"]) == "psak 72"]
    assert len(psak_rows) == 2 and all(r["status"] == "COCOK" for r in psak_rows), psak_rows
    comprehensive = next(r for r in march_rows
                         if normalize_label(r["label"]) == "total comprehensive income for the period")
    assert comprehensive["status"] == "COCOK", comprehensive
    assert abs(comprehensive["pdf_value"] - (-820199808)) < 1, comprehensive
    print("selftest OK: duplikat PSAK 72 dan comprehensive income terpetakan")


DEFAULT_PDF_PATTERN = "Report LapKeu CAM JO {month} {year}.pdf"


def json_report(excel_path, pdf_dir, pdf_pattern):
    """Return the audit rows needed by the local Next.js UI."""
    validation = validate_data_folder(pdf_dir, pdf_pattern)
    rows = [r for r in collect_rows_for_periods(
        validation["excel_path"], pdf_dir, validation["available_periods"], pdf_pattern
    ) if r["status"] != "HANYA_PDF"]
    summary = {status: sum(1 for row in rows if row["status"] == status)
               for status in ("COCOK", "BEDA", "HANYA_EXCEL")}
    return {
        "excel": validation["excel_path"].name,
        "periods": validation["available_periods"],
        "missing_periods": validation["missing_periods"],
        "summary": summary,
        "rows": rows,
    }

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--excel", type=Path, default=EXCEL_PATH, help="Path ke Excel rekap (xlsx).")
    parser.add_argument("--pdf-dir", type=Path, default=DATA_DIR, help="Folder berisi PDF bulanan.")
    parser.add_argument("--year", type=int, default=2026, help="Tahun data (dipakai cari kolom bulan & nama file PDF).")
    parser.add_argument("--pdf-pattern", default=DEFAULT_PDF_PATTERN,
                         help='Pola nama file PDF per bulan, pakai placeholder {month} (nama bulan Indonesia) dan {year}. Default: "%(default)s"')
    parser.add_argument("--apply", action="store_true", help="Tulis nilai PDF ke Excel untuk baris BEDA (angka polos saja).")
    parser.add_argument("--no-selftest", action="store_true", help="Lewati selftest sebelum laporan.")
    parser.add_argument("--json", action="store_true", help="Keluarkan hasil audit sebagai JSON untuk UI lokal.")
    args = parser.parse_args()

    is_default_data = args.excel == EXCEL_PATH and args.pdf_dir == DATA_DIR and args.year == 2026
    if args.json:
        print(json.dumps(json_report(args.excel, args.pdf_dir, args.pdf_pattern), ensure_ascii=False))
        raise SystemExit(0)
    if not args.no_selftest and is_default_data:
        _selftest()
    run_report(args.excel, args.pdf_dir, args.year, args.pdf_pattern, args.apply)
