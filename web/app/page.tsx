"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import type { User } from "@supabase/supabase-js";
import { supabase } from "../lib/supabase/client";

type Category = "Semua" | "Balance Sheet" | "Profit & Loss" | "Cash Flow";
type Status = "COCOK" | "BEDA" | "HANYA_EXCEL";
type SortKey = "label" | "category" | "value" | "pdf_value" | "diff" | "status";
type View = "audit" | "history" | "user";
type ApiRow = {
  month: string;
  period: string;
  sheet: string;
  label: string;
  value: number | null;
  pdf_value: number | null;
  diff: number | null;
  status: Status;
};
type AuditData = {
  excel: string;
  periods: string[];
  missing_periods: string[];
  summary: Record<Status, number>;
  rows: ApiRow[];
};
type HistoryRun = {
  id: string;
  name: string;
  source_excel: string | null;
  status: string;
  created_at: string;
};
const categories: Category[] = [
  "Semua",
  "Balance Sheet",
  "Profit & Loss",
  "Cash Flow",
];
const sheetNames: Record<string, Exclude<Category, "Semua">> = {
  "balance-sheet": "Balance Sheet",
  "profit-loss": "Profit & Loss",
  cashflow: "Cash Flow",
};
const statusLabel: Record<Status, string> = {
  COCOK: "Cocok",
  BEDA: "Beda",
  HANYA_EXCEL: "Hanya Excel",
};
const formatNumber = (value: number | null) =>
  value == null
    ? "—"
    : new Intl.NumberFormat("en-US", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
      }).format(value);
const monthLabel = (period: string) =>
  new Intl.DateTimeFormat("id-ID", { month: "long", year: "numeric" }).format(
    new Date(`${period}-01T00:00:00`),
  );
const FILE_BUCKET = "mapping-files";
const monthNumbers: Record<string, number> = {
  januari: 1,
  februari: 2,
  maret: 3,
  april: 4,
  mei: 5,
  juni: 6,
  juli: 7,
  agustus: 8,
  september: 9,
  oktober: 10,
  november: 11,
  desember: 12,
};
const fileNameOnly = (file: File) =>
  (file.webkitRelativePath || file.name)
    .replaceAll("\\", "/")
    .split("/")
    .pop() || file.name;
const periodFromFileName = (name: string) => {
  const match = name.toLowerCase().match(
    /(januari|februari|maret|april|mei|juni|juli|agustus|september|oktober|november|desember)[ -]+(20\d{2})/,
  );
  return match
    ? `${match[2]}-${String(monthNumbers[match[1]]).padStart(2, "0")}`
    : "";
};
const periodsFromFileNames = (names: string[]) =>
  [...new Set(names.map(periodFromFileName).filter(Boolean))].sort();

export default function Home() {
  const [user, setUser] = useState<User | null>(null),
    [authReady, setAuthReady] = useState(false),
    [email, setEmail] = useState(""),
    [password, setPassword] = useState(""),
    [authMode, setAuthMode] = useState<"login" | "signup">("login"),
    [authMessage, setAuthMessage] = useState(""),
    [authBusy, setAuthBusy] = useState(false);
  const folderInputRef = useRef<HTMLInputElement>(null),
    excelInputRef = useRef<HTMLInputElement>(null),
    pdfInputRef = useRef<HTMLInputElement>(null);
  const [files, setFiles] = useState<File[]>([]),
    [storedFileNames, setStoredFileNames] = useState<string[]>([]),
    [availablePeriods, setAvailablePeriods] = useState<string[]>([]),
    [data, setData] = useState<AuditData | null>(null),
    [period, setPeriod] = useState(""),
    [category, setCategory] = useState<Category>("Semua"),
    [busy, setBusy] = useState(false),
    [message, setMessage] = useState(""),
    [saved, setSaved] = useState(false),
    [activeView, setActiveView] = useState<View>("audit"),
    [history, setHistory] = useState<HistoryRun[]>([]),
    [sort, setSort] = useState<{ key: SortKey; direction: "asc" | "desc" }>({
      key: "label",
      direction: "asc",
    });
  useEffect(() => {
    if (!supabase) {
      setAuthReady(true);
      return;
    }
    supabase.auth.getUser().then(({ data }) => {
      setUser(data.user);
      setAuthReady(true);
    });
    const { data: listener } = supabase.auth.onAuthStateChange(
      (_event, session) => setUser(session?.user ?? null),
    );
    return () => listener.subscription.unsubscribe();
  }, []);
  useEffect(() => {
    if (!user || !supabase || activeView !== "history") return;
    supabase
      .from("mapping_runs")
      .select("id,name,source_excel,status,created_at")
      .order("created_at", { ascending: false })
      .limit(20)
      .then(({ data }) => setHistory((data as HistoryRun[] | null) ?? []));
  }, [user, activeView]);
  useEffect(() => {
    if (!user || !supabase) return;
    let cancelled = false;
    (async () => {
      const { data: objects, error } = await supabase.storage
        .from(FILE_BUCKET)
        .list(user.id, { limit: 100 });
      if (cancelled || error || !objects?.length) return;
      const names = objects.filter((object) => object.name).map((object) => object.name);
      const periods = periodsFromFileNames(names);
      setStoredFileNames(names);
      setAvailablePeriods(periods);
      const firstPeriod = periods[0] ?? "";
      const selectedNames = [
        names.find((name) => name.toLowerCase().endsWith(".xlsx")),
        ...names.filter(
          (name) =>
            name.toLowerCase().endsWith(".pdf") &&
            (!firstPeriod || periodFromFileName(name) === firstPeriod),
        ),
      ].filter((name): name is string => Boolean(name));
      const restored = (
        await Promise.all(
          selectedNames.map(async (name) => {
            const { data: blob } = await supabase!.storage
              .from(FILE_BUCKET)
              .download(`${user.id}/${name}`);
            return blob
              ? new File([blob], name, {
                  type: blob.type || "application/octet-stream",
                })
              : null;
          }),
        )
      ).filter((file): file is File => Boolean(file));
      if (!cancelled && restored.length) {
        setFiles(restored);
        setMessage(
          `${restored.length} file untuk ${firstPeriod ? monthLabel(firstPeriod) : "periode aktif"} dimuat. Audit dijalankan...`,
        );
        await runAudit(restored, false);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [user]);
  const submitAuth = async (event: React.FormEvent) => {
    event.preventDefault();
    setAuthBusy(true);
    setAuthMessage("");
    if (!supabase) {
      setAuthMessage("Konfigurasi Supabase belum tersedia.");
      setAuthBusy(false);
      return;
    }
    const result =
      authMode === "login"
        ? await supabase.auth.signInWithPassword({ email, password })
        : await supabase.auth.signUp({ email, password });
    if (result.error) setAuthMessage(result.error.message);
    else {
      setPassword("");
      setAuthMessage(
        authMode === "signup"
          ? "Akun berhasil dibuat. Cek email Anda jika konfirmasi email aktif."
          : "",
      );
    }
    setAuthBusy(false);
  };
  const rows = data?.rows ?? [],
    periods = availablePeriods.length ? availablePeriods : data?.periods ?? [],
    currentPeriod = period || periods[0] || "";
  const periodRows = useMemo(
    () => rows.filter((row) => row.period === currentPeriod),
    [rows, currentPeriod],
  );
  const currentSummary = useMemo(
    () => ({
      COCOK: periodRows.filter((row) => row.status === "COCOK").length,
      BEDA: periodRows.filter((row) => row.status === "BEDA").length,
      HANYA_EXCEL: periodRows.filter((row) => row.status === "HANYA_EXCEL")
        .length,
    }),
    [periodRows],
  );
  const onlyExcelLabels = useMemo(
    () =>
      periodRows
        .filter((row) => row.status === "HANYA_EXCEL")
        .map((row) => row.label),
    [periodRows],
  );
  const filteredRows = useMemo(
    () =>
      [
        ...periodRows.filter(
          (row) => category === "Semua" || sheetNames[row.sheet] === category,
        ),
      ].sort((left, right) => {
        const a =
            sort.key === "category" ? sheetNames[left.sheet] : left[sort.key],
          b =
            sort.key === "category" ? sheetNames[right.sheet] : right[sort.key];
        const result =
          typeof a === "number" && typeof b === "number"
            ? a - b
            : String(a ?? "").localeCompare(String(b ?? ""), "id");
        return sort.direction === "asc" ? result : -result;
      }),
    [periodRows, category, sort],
  );
  if (!authReady)
    return (
      <main className="auth-screen">
        <div className="auth-card">Memuat sesi...</div>
      </main>
    );
  if (!user)
    return (
      <main className="auth-screen">
        <section className="auth-layout">
          <div className="auth-visual">
            <span className="auth-visual-mark">MK</span>
            <div>
              <p className="auth-kicker">MAPPING KEUANGAN</p>
              <h1>
                Audit laporan
                <br />
                lebih sederhana.
              </h1>
              <p>
                Bandingkan data Excel dan PDF dalam satu workspace yang rapi.
              </p>
            </div>
            <div className="auth-tiles">
              <span>✓</span>
              <span>▤</span>
              <span>↗</span>
            </div>
          </div>
          <form className="auth-card" onSubmit={submitAuth}>
            <div className="auth-tabs">
              <button
                type="button"
                className={authMode === "login" ? "selected" : ""}
                onClick={() => {
                  setAuthMode("login");
                  setAuthMessage("");
                }}
              >
                Masuk
              </button>
              <button
                type="button"
                className={authMode === "signup" ? "selected" : ""}
                onClick={() => {
                  setAuthMode("signup");
                  setAuthMessage("");
                }}
              >
                Daftar
              </button>
            </div>
            <span className="brand-mark">MK</span>
            <h2>
              {authMode === "login"
                ? "Selamat datang kembali"
                : "Buat akun baru"}
            </h2>
            <p className="muted">
              {authMode === "login"
                ? "Masuk untuk melanjutkan audit keuangan Anda."
                : "Daftar untuk menyimpan riwayat audit Anda."}
            </p>
            <label>
              Email
              <input
                type="email"
                value={email}
                onChange={(event) => setEmail(event.target.value)}
                autoComplete="email"
                required
              />
            </label>
            <label>
              Password
              <input
                type="password"
                value={password}
                onChange={(event) => setPassword(event.target.value)}
                autoComplete={
                  authMode === "login" ? "current-password" : "new-password"
                }
                required
              />
            </label>
            {authMessage && (
              <p
                className={
                  authMessage.startsWith("Akun") ? "auth-success" : "auth-error"
                }
                role="status"
              >
                {authMessage}
              </p>
            )}
            <button className="button primary" disabled={authBusy}>
              {authBusy
                ? "Memproses..."
                : authMode === "login"
                  ? "Masuk"
                  : "Daftar"}
            </button>
            <p className="auth-switch">
              {authMode === "login"
                ? "Belum punya akun? "
                : "Sudah punya akun? "}
              <button
                type="button"
                onClick={() => {
                  setAuthMode(authMode === "login" ? "signup" : "login");
                  setAuthMessage("");
                }}
              >
                {authMode === "login" ? "Daftar sekarang" : "Masuk di sini"}
              </button>
            </p>
          </form>
        </section>
      </main>
    );
  const changeSort = (key: SortKey) =>
    setSort((current) =>
      current.key === key
        ? { key, direction: current.direction === "asc" ? "desc" : "asc" }
        : { key, direction: "asc" },
    );
  const sortLabel = (key: SortKey) =>
    sort.key === key ? (sort.direction === "asc" ? " ↑" : " ↓") : "";
  const selectView = (view: View) => {
    setActiveView(view);
    setMessage("");
  };
  const selectFolder = (event: React.ChangeEvent<HTMLInputElement>) => {
    const selected = Array.from(event.target.files ?? []).filter((file) =>
      /\.(xlsx|pdf)$/i.test(file.name),
    );
    setFiles(selected);
    setStoredFileNames(selected.map(fileNameOnly));
    setAvailablePeriods(periodsFromFileNames(selected.map(fileNameOnly)));
    setMessage(
      selected.length
        ? `${selected.length} file dari folder siap disimpan dan diaudit.`
        : "Pilih folder yang berisi Excel dan PDF.",
    );
  };
  const selectExcel = (event: React.ChangeEvent<HTMLInputElement>) => {
    const selected = Array.from(event.target.files ?? []).filter((file) =>
      file.name.toLowerCase().endsWith(".xlsx"),
    );
    setFiles((current) => [
      ...current.filter((file) => !file.name.toLowerCase().endsWith(".xlsx")),
      ...selected,
    ]);
    setStoredFileNames((current) => [
      ...current.filter((name) => !name.toLowerCase().endsWith(".xlsx")),
      ...selected.map(fileNameOnly),
    ]);
    setMessage(
      selected.length
        ? "File Excel siap disimpan dan diaudit."
        : "Pilih satu file Excel (.xlsx).",
    );
  };
  const selectPdfs = (event: React.ChangeEvent<HTMLInputElement>) => {
    const selected = Array.from(event.target.files ?? []).filter((file) =>
      file.name.toLowerCase().endsWith(".pdf"),
    );
    const selectedNames = selected.map(fileNameOnly);
    setFiles((current) => [
      ...current.filter((file) => !file.name.toLowerCase().endsWith(".pdf")),
      ...selected,
    ]);
    setStoredFileNames((current) => [
      ...current.filter((name) => !name.toLowerCase().endsWith(".pdf")),
      ...selectedNames,
    ]);
    setAvailablePeriods(
      periodsFromFileNames([
        ...storedFileNames.filter((name) => !name.toLowerCase().endsWith(".pdf")),
        ...selectedNames,
      ]),
    );
    setMessage(
      selected.length
        ? `${selected.length} file PDF siap disimpan dan diaudit.`
        : "Pilih file PDF.",
    );
  };
  const runAudit = async (inputFiles = files, persist = true) => {
    const hasExcel = inputFiles.some((file) =>
        file.name.toLowerCase().endsWith(".xlsx"),
      ),
      hasPdf = inputFiles.some((file) =>
        file.name.toLowerCase().endsWith(".pdf"),
      );
    if (!hasExcel || !hasPdf) {
      setMessage(
        !hasExcel
          ? "Pilih file Excel terlebih dahulu."
          : "Pilih minimal satu file PDF.",
      );
      return;
    }
    setBusy(true);
    setSaved(false);
    setMessage("Membaca Excel dan PDF...");
    try {
      const form = new FormData();
      inputFiles.forEach((file) =>
        form.append("files", file, fileNameOnly(file)),
      );
      const session = supabase
        ? (await supabase.auth.getSession()).data.session
        : null;
      const endpoint =
        process.env.NODE_ENV === "production"
          ? "/api/audit_serverless"
          : "/api/audit";
      const response = await fetch(endpoint, {
        method: "POST",
        headers: session
          ? { Authorization: `Bearer ${session.access_token}` }
          : undefined,
        body: form,
      });
      const result = await response.json();
      if (!response.ok)
        throw new Error(result.error || result.detail || "Audit gagal.");
      setData(result);
      setPeriod(result.periods[0] ?? "");
      setSaved(true);
      if (persist && supabase && user) {
        const uploads = await Promise.all(
          inputFiles.map((file) =>
            supabase!.storage
              .from(FILE_BUCKET)
              .upload(`${user.id}/${fileNameOnly(file)}`, file, {
                upsert: true,
                contentType: file.type || undefined,
              }),
          ),
        );
        const storageError = uploads.find((upload) => upload.error)?.error;
        setMessage(
          storageError
            ? `${result.rows.length} COA berhasil diaudit, tetapi file belum tersimpan ke Storage. Jalankan SQL Storage terlebih dahulu.`
            : `${result.rows.length} COA berhasil dimuat dan disimpan.`,
        );
      } else
        setMessage(
          `${result.rows.length} COA Excel berhasil dimuat dan disimpan.`,
        );
    } catch (error) {
      setMessage(
        error instanceof Error ? error.message : "Audit gagal dijalankan.",
      );
    } finally {
      setBusy(false);
    }
  };
  const selectPeriod = async (nextPeriod: string) => {
    setPeriod(nextPeriod);
    if (busy || !nextPeriod) return;
    const names = storedFileNames.length ? storedFileNames : files.map(fileNameOnly);
    const selectedNames = [
      names.find((name) => name.toLowerCase().endsWith(".xlsx")),
      ...names.filter(
        (name) =>
          name.toLowerCase().endsWith(".pdf") &&
          periodFromFileName(name) === nextPeriod,
      ),
    ].filter((name): name is string => Boolean(name));
    if (!selectedNames.some((name) => name.toLowerCase().endsWith(".pdf"))) {
      setMessage(`PDF untuk ${monthLabel(nextPeriod)} belum tersedia.`);
      return;
    }
    setBusy(true);
    try {
      const selectedFiles = await Promise.all(
        selectedNames.map(async (name) => {
          const local = files.find((file) => fileNameOnly(file) === name);
          if (local) return local;
          if (!supabase || !user) return null;
          const { data: blob } = await supabase.storage
            .from(FILE_BUCKET)
            .download(`${user.id}/${name}`);
          return blob
            ? new File([blob], name, { type: blob.type || "application/octet-stream" })
            : null;
        }),
      );
      const ready = selectedFiles.filter((file): file is File => Boolean(file));
      if (!ready.some((file) => file.name.toLowerCase().endsWith(".xlsx"))) {
        setBusy(false);
        setMessage("File Excel tersimpan tidak ditemukan.");
        return;
      }
      setFiles(ready);
      await runAudit(ready, false);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "Gagal memuat periode.");
      setBusy(false);
    }
  };
  return (
    <main className="app-shell">
      <aside className="sidebar">
        <div className="brand">
          <span className="brand-mark">MK</span>
          <span>
            Mapping
            <br />
            Keuangan
          </span>
        </div>
        <nav className="nav-list" aria-label="Navigasi utama">
          <button
            className={`nav-item ${activeView === "audit" ? "active" : ""}`}
            onClick={() => selectView("audit")}
          >
            <span>▦</span> Hasil Audit
          </button>
          <button
            className={`nav-item ${activeView === "history" ? "active" : ""}`}
            onClick={() => selectView("history")}
          >
            <span>◷</span> Riwayat
          </button>
          <button
            className={`nav-item ${activeView === "user" ? "active" : ""}`}
            onClick={() => selectView("user")}
          >
            <span>⚙</span> User
          </button>
        </nav>
        <div className="sidebar-footer">
          <span className="avatar">
            {(user.email?.[0] || "U").toUpperCase()}
          </span>
          <div>
            <strong>{user.email?.split("@")[0] || "User"}</strong>
            <small>Supabase account</small>
          </div>
          <button className="sidebar-logout" onClick={() => supabase?.auth.signOut()} title="Keluar">
            Keluar
          </button>
        </div>
      </aside>
      <section className="workspace">
        <header className="topbar">
          <div>
            <p className="eyebrow">WORKSPACE / {activeView.toUpperCase()}</p>
            <h1>
              {activeView === "audit"
                ? "Hasil Audit Keuangan"
                : activeView === "history"
                  ? "Riwayat Audit"
                  : "Pengaturan User"}
            </h1>
            <p className="muted">
              {activeView === "audit"
                ? "Bandingkan rekap Excel dengan laporan PDF per periode."
                : activeView === "history"
                  ? "Riwayat proses audit yang tersimpan."
                  : "Informasi akun yang sedang digunakan."}
            </p>
          </div>
          {activeView === "audit" && (
            <div className="top-actions">
              <input
                ref={folderInputRef}
                type="file"
                multiple
                webkitdirectory=""
                directory=""
                className="hidden-input"
                onChange={selectFolder}
              />
              <input
                ref={excelInputRef}
                type="file"
                accept=".xlsx"
                className="hidden-input"
                onChange={selectExcel}
              />
              <input
                ref={pdfInputRef}
                type="file"
                accept=".pdf"
                multiple
                className="hidden-input"
                onChange={selectPdfs}
              />
              <button
                className="button secondary"
                onClick={() => folderInputRef.current?.click()}
              >
                ↥ Upload Folder
              </button>
              <button
                className="button secondary"
                onClick={() => excelInputRef.current?.click()}
              >
                ↥ Upload Excel
              </button>
              <button
                className="button secondary"
                onClick={() => pdfInputRef.current?.click()}
              >
                ↥ Upload PDF
              </button>
              <button
                className="button primary"
                onClick={() => runAudit()}
                disabled={busy}
              >
                {busy ? "Memproses..." : "Jalankan Audit"}
              </button>
            </div>
          )}
        </header>
        {activeView === "audit" ? (
          <>
            <div className="file-strip">
              <span className="file-dot" />{" "}
              {data
                ? `${data.excel} • ${periods.length} periode siap`
                : files.length
                  ? `${files.length} file siap • klik Jalankan Audit`
                  : "Belum ada data. Upload folder, Excel, atau PDF."}
            </div>
            {message && (
              <div className="notice" role="status">
                {message}
              </div>
            )}
            <section className="summary-grid">
              <SummaryCard
                label="Cocok bulan ini"
                value={data ? currentSummary.COCOK : "—"}
                tone="green"
                icon="✓"
              />
              <SummaryCard
                label="Beda bulan ini"
                value={data ? currentSummary.BEDA : "—"}
                tone="red"
                icon="×"
              />
              <SummaryCard
                label="Hanya Excel bulan ini"
                value={data ? currentSummary.HANYA_EXCEL : "—"}
                tone="amber"
                icon="◐"
                detail={data ? onlyExcelLabels.join(", ") || "Tidak ada" : "—"}
              />
            </section>
            <section className="content-panel">
              <div className="panel-toolbar">
                <div>
                  <h2>
                    {currentPeriod
                      ? monthLabel(currentPeriod)
                      : "Belum ada hasil"}
                  </h2>
                  <p className="muted">
                    {data
                      ? `${filteredRows.length} COA ditampilkan dari Excel`
                      : "Upload data lalu jalankan audit"}
                  </p>
                </div>
                <div className="panel-controls">
                  <label className="period-control" htmlFor="period-select">
                    <span>Periode</span>
                    <select
                      id="period-select"
                      value={currentPeriod}
                      onChange={(event) => void selectPeriod(event.target.value)}
                      disabled={!periods.length}
                    >
                      <option value="">Belum ada periode</option>
                      {periods.map((item) => {
                        const beda = rows.filter(
                            (row) =>
                              row.period === item && row.status === "BEDA",
                          ).length,
                          only = rows.filter(
                            (row) =>
                              row.period === item &&
                              row.status === "HANYA_EXCEL",
                          ).length;
                        return (
                          <option key={item} value={item}>
                            {monthLabel(item)}
                            {beda
                              ? ` — ${beda} Beda`
                              : only
                                ? ` — ${only} Hanya Excel`
                                : " — Cocok"}
                          </option>
                        );
                      })}
                    </select>
                  </label>
                  <button
                    className="button secondary save-button"
                    disabled={!saved}
                    onClick={() =>
                      setMessage("Hasil audit sudah tersimpan di Supabase.")
                    }
                  >
                    {saved ? "✓ Tersimpan" : "Simpan Hasil"}
                  </button>
                </div>
              </div>
              <div className="category-tabs">
                {categories.map((item) => (
                  <button
                    key={item}
                    className={`category-tab ${category === item ? "selected" : ""}`}
                    onClick={() => setCategory(item)}
                  >
                    {item}
                  </button>
                ))}
              </div>
              <div className="table-wrap">
                <table>
                  <thead>
                    <tr>
                      {(
                        [
                          ["label", "Akun / COA Excel"],
                          ["category", "Kategori"],
                          ["value", "Excel"],
                          ["pdf_value", "PDF"],
                          ["diff", "Selisih"],
                          ["status", "Status"],
                        ] as const
                      ).map(([key, label]) => (
                        <th key={key}>
                          <button
                            className="sort-button"
                            onClick={() => changeSort(key)}
                          >
                            {label}
                            {sortLabel(key)}
                          </button>
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {filteredRows.length ? (
                      filteredRows.map((row, index) => (
                        <tr
                          key={`${row.period}-${row.sheet}-${row.label}-${index}`}
                        >
                          <td className="account">{row.label}</td>
                          <td className="category-cell">
                            {sheetNames[row.sheet] ?? row.sheet}
                          </td>
                          <td>{formatNumber(row.value)}</td>
                          <td>{formatNumber(row.pdf_value)}</td>
                          <td>{formatNumber(row.diff)}</td>
                          <td>
                            <span
                              className={`status status-${row.status.toLowerCase()}`}
                            >
                              <i />
                              {statusLabel[row.status]}
                            </span>
                          </td>
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan={6} className="empty-state">
                          Belum ada COA untuk ditampilkan.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </section>
          </>
        ) : activeView === "history" ? (
          <HistoryPanel history={history} />
        ) : (
          <UserPanel user={user} />
        )}
      </section>
    </main>
  );
}
function HistoryPanel({ history }: { history: HistoryRun[] }) {
  return (
    <section className="content-panel simple-page">
      <div className="panel-toolbar">
        <div>
          <h2>Riwayat Audit</h2>
          <p className="muted">
            {history.length
              ? `${history.length} proses terakhir`
              : "Belum ada audit tersimpan."}
          </p>
        </div>
      </div>
      <div className="history-list">
        {history.length ? (
          history.map((run) => (
            <div className="history-row" key={run.id}>
              <div>
                <strong>{run.name}</strong>
                <small>
                  {run.source_excel || "Sumber Excel tidak dicatat"}
                </small>
              </div>
              <span>{run.status}</span>
              <time>
                {new Intl.DateTimeFormat("id-ID", {
                  dateStyle: "medium",
                  timeStyle: "short",
                }).format(new Date(run.created_at))}
              </time>
            </div>
          ))
        ) : (
          <p className="empty-state">
            Riwayat akan muncul setelah audit dijalankan.
          </p>
        )}
      </div>
    </section>
  );
}
function UserPanel({ user }: { user: User }) {
  return (
    <section className="content-panel simple-page user-panel">
      <div className="panel-toolbar">
        <div>
          <h2>User</h2>
          <p className="muted">
            Akun yang sedang digunakan untuk menyimpan hasil audit.
          </p>
        </div>
      </div>
      <div className="user-details">
        <span className="avatar">{(user.email?.[0] || "U").toUpperCase()}</span>
        <div>
          <strong>{user.email}</strong>
          <small>Supabase account</small>
        </div>
      </div>
      <button
        className="button secondary"
        onClick={() => supabase?.auth.signOut()}
      >
        Keluar
      </button>
    </section>
  );
}
function SummaryCard({
  label,
  value,
  tone,
  icon,
  detail,
}: {
  label: string;
  value: string | number;
  tone: string;
  icon: string;
  detail?: string;
}) {
  return (
    <div className={`summary-card ${tone}`}>
      <span className="summary-icon">{icon}</span>
      <div>
        <strong>{value}</strong>
        <span>{label}</span>
        {detail && (
          <small className="summary-detail" title={detail}>
            {detail}
          </small>
        )}
      </div>
    </div>
  );
}
