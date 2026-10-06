"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import type { User } from "@supabase/supabase-js";
import { supabase } from "../lib/supabase/client";

type Category = "Semua" | "Balance Sheet" | "Profit & Loss" | "Cash Flow";
type Status = "COCOK" | "BEDA" | "HANYA_EXCEL";
type SortKey = "label" | "category" | "value" | "pdf_value" | "diff" | "status";
type View = "audit" | "history" | "user";
type ApiRow = { month: string; period: string; sheet: string; label: string; value: number | null; pdf_value: number | null; diff: number | null; status: Status };
type AuditData = { excel: string; periods: string[]; missing_periods: string[]; summary: Record<Status, number>; rows: ApiRow[] };

const categories: Category[] = ["Semua", "Balance Sheet", "Profit & Loss", "Cash Flow"];
const sheetNames: Record<string, Exclude<Category, "Semua">> = { "balance-sheet": "Balance Sheet", "profit-loss": "Profit & Loss", cashflow: "Cash Flow" };
const statusLabel: Record<Status, string> = { COCOK: "Cocok", BEDA: "Beda", HANYA_EXCEL: "Hanya Excel" };
const formatNumber = (value: number | null) => value == null ? "—" : new Intl.NumberFormat("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(value);
const monthLabel = (period: string) => new Intl.DateTimeFormat("id-ID", { month: "long", year: "numeric" }).format(new Date(`${period}-01T00:00:00`));

export default function Home() {
  const [user, setUser] = useState<User | null>(null);
  const [authReady, setAuthReady] = useState(false);
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [authMode, setAuthMode] = useState<"login" | "signup">("login");
  const [authMessage, setAuthMessage] = useState("");
  const [authBusy, setAuthBusy] = useState(false);
  useEffect(() => {
    if (!supabase) { setAuthReady(true); return; }
    supabase.auth.getUser().then(({ data }) => { setUser(data.user); setAuthReady(true); });
    const { data: listener } = supabase.auth.onAuthStateChange((_event, session) => setUser(session?.user ?? null));
    return () => listener.subscription.unsubscribe();
  }, []);
  const submitAuth = async (event: React.FormEvent) => {
    event.preventDefault(); setAuthBusy(true); setAuthMessage("");
    if (!supabase) { setAuthMessage("Konfigurasi Supabase belum tersedia."); setAuthBusy(false); return; }
    const result = authMode === "login"
      ? await supabase.auth.signInWithPassword({ email, password })
      : await supabase.auth.signUp({ email, password });
    if (result.error) setAuthMessage(result.error.message);
    else { setPassword(""); setAuthMessage(authMode === "signup" ? "Akun berhasil dibuat. Cek email Anda jika konfirmasi email aktif." : ""); }
    setAuthBusy(false);
  };
  if (!authReady) return <main className="auth-screen"><div className="auth-card">Memuat sesi...</div></main>;
  if (!user) return <main className="auth-screen"><section className="auth-layout"><div className="auth-visual"><span className="auth-visual-mark">MK</span><div><p className="auth-kicker">MAPPING KEUANGAN</p><h1>Audit laporan<br />lebih sederhana.</h1><p>Bandingkan data Excel dan PDF dalam satu workspace yang rapi.</p></div><div className="auth-tiles"><span>✓</span><span>▤</span><span>↗</span></div></div><form className="auth-card" onSubmit={submitAuth}><div className="auth-tabs"><button type="button" className={authMode === "login" ? "selected" : ""} onClick={() => { setAuthMode("login"); setAuthMessage(""); }}>Masuk</button><button type="button" className={authMode === "signup" ? "selected" : ""} onClick={() => { setAuthMode("signup"); setAuthMessage(""); }}>Daftar</button></div><span className="brand-mark">MK</span><h2>{authMode === "login" ? "Selamat datang kembali" : "Buat akun baru"}</h2><p className="muted">{authMode === "login" ? "Masuk untuk melanjutkan audit keuangan Anda." : "Daftar untuk menyimpan riwayat audit Anda."}</p><label>Email<input type="email" value={email} onChange={(event) => setEmail(event.target.value)} autoComplete="email" required /></label><label>Password<input type="password" value={password} onChange={(event) => setPassword(event.target.value)} autoComplete={authMode === "login" ? "current-password" : "new-password"} required /></label>{authMessage && <p className={authMessage.startsWith("Akun") ? "auth-success" : "auth-error"} role="status">{authMessage}</p>}<button className="button primary" disabled={authBusy}>{authBusy ? "Memproses..." : authMode === "login" ? "Masuk" : "Daftar"}</button><p className="auth-switch">{authMode === "login" ? "Belum punya akun? " : "Sudah punya akun? "}<button type="button" onClick={() => { setAuthMode(authMode === "login" ? "signup" : "login"); setAuthMessage(""); }}>{authMode === "login" ? "Daftar sekarang" : "Masuk di sini"}</button></p></form></section></main>;
  const inputRef = useRef<HTMLInputElement>(null);
  const [files, setFiles] = useState<File[]>([]);
  const [data, setData] = useState<AuditData | null>(null);
  const [period, setPeriod] = useState("");
  const [category, setCategory] = useState<Category>("Semua");
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState("");
  const [activeView, setActiveView] = useState<View>("audit");
  const [sort, setSort] = useState<{ key: SortKey; direction: "asc" | "desc" }>({ key: "label", direction: "asc" });
  const rows = data?.rows ?? [];
  const periods = data?.periods ?? [];
  const currentPeriod = period || periods[0] || "";
  const periodRows = useMemo(() => rows.filter((row) => row.period === currentPeriod), [rows, currentPeriod]);
  const currentSummary = useMemo(() => ({
    COCOK: periodRows.filter((row) => row.status === "COCOK").length,
    BEDA: periodRows.filter((row) => row.status === "BEDA").length,
    HANYA_EXCEL: periodRows.filter((row) => row.status === "HANYA_EXCEL").length,
  }), [periodRows]);
  const filteredRows = useMemo(() => {
    const visible = periodRows.filter((row) => category === "Semua" || sheetNames[row.sheet] === category);
    return [...visible].sort((left, right) => {
      const a = sort.key === "category" ? sheetNames[left.sheet] : left[sort.key];
      const b = sort.key === "category" ? sheetNames[right.sheet] : right[sort.key];
      const result = typeof a === "number" && typeof b === "number" ? a - b : String(a ?? "").localeCompare(String(b ?? ""), "id");
      return sort.direction === "asc" ? result : -result;
    });
  }, [periodRows, category, sort]);
  const changeSort = (key: SortKey) => setSort((current) => current.key === key ? { key, direction: current.direction === "asc" ? "desc" : "asc" } : { key, direction: "asc" });
  const sortLabel = (key: SortKey) => sort.key === key ? (sort.direction === "asc" ? " ↑" : " ↓") : "";
  const selectView = (view: View) => { setActiveView(view); setMessage(""); };
  const selectFiles = (event: React.ChangeEvent<HTMLInputElement>) => { const selected = Array.from(event.target.files ?? []); setFiles(selected); setMessage(selected.length ? `${selected.length} file siap diproses.` : ""); };
  const runAudit = async () => {
    if (!files.length) { setMessage("Pilih folder data terlebih dahulu."); inputRef.current?.click(); return; }
    setBusy(true); setMessage("Membaca Excel dan PDF...");
    try {
      const form = new FormData(); files.forEach((file) => form.append("files", file, file.webkitRelativePath || file.name));
      const session = supabase ? (await supabase.auth.getSession()).data.session : null;
      const response = await fetch("/api/audit", { method: "POST", headers: session ? { Authorization: `Bearer ${session.access_token}` } : undefined, body: form });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || "Audit gagal.");
      setData(result); setPeriod(result.periods[0] ?? ""); setMessage(`${result.rows.length} COA Excel berhasil dimuat.`);
    } catch (error) { setMessage(error instanceof Error ? error.message : "Audit gagal dijalankan."); }
    finally { setBusy(false); }
  };

  return <main className="app-shell">
    <aside className="sidebar"><div className="brand"><span className="brand-mark">MK</span><span>Mapping<br />Keuangan</span></div><nav className="nav-list" aria-label="Navigasi utama"><button className={`nav-item ${activeView === "audit" ? "active" : ""}`} onClick={() => selectView("audit")}><span>▦</span> Hasil Audit</button><button className={`nav-item ${activeView === "history" ? "active" : ""}`} onClick={() => selectView("history")}><span>◷</span> Riwayat</button><button className={`nav-item ${activeView === "user" ? "active" : ""}`} onClick={() => selectView("user")}><span>⚙</span> User</button></nav><div className="sidebar-footer"><span className="avatar">A</span><div><strong>Admin</strong><small>Local workspace</small></div><button className="more" aria-label="Menu akun">•••</button></div></aside>
    <section className="workspace"><header className="topbar"><div><p className="eyebrow">WORKSPACE / {activeView.toUpperCase()}</p><h1>{activeView === "audit" ? "Hasil Audit Keuangan" : activeView === "history" ? "Riwayat Audit" : "Pengaturan User"}</h1><p className="muted">{activeView === "audit" ? "Bandingkan rekap Excel dengan laporan PDF per periode." : activeView === "history" ? "Riwayat proses audit yang dijalankan di workspace ini." : "Informasi user yang sedang memakai workspace lokal."}</p></div>{activeView === "audit" && <div className="top-actions">
      {/* @ts-expect-error webkitdirectory is supported by Chromium but missing from React typings. */}
      <input ref={inputRef} type="file" multiple webkitdirectory="" directory="" className="hidden-input" onChange={selectFiles} />
      <button className="button secondary" onClick={() => inputRef.current?.click()}>↥ Upload Folder</button><button className="button primary" onClick={runAudit} disabled={busy}>{busy ? "Memproses..." : "Jalankan Audit"}</button></div>}</header>
      {activeView === "audit" ? <><div className="file-strip"><span className="file-dot" /> {data ? `${data.excel} • ${periods.length} periode siap` : files.length ? `${files.length} file dipilih • klik Jalankan Audit` : "Belum ada data. Upload folder Excel dan PDF."}</div>
      {message && <div className="notice" role="status">{message}</div>}
      <section className="summary-grid" aria-label="Ringkasan audit"><SummaryCard label="Cocok bulan ini" value={data ? currentSummary.COCOK : "—"} tone="green" icon="✓" /><SummaryCard label="Beda bulan ini" value={data ? currentSummary.BEDA : "—"} tone="red" icon="×" /><SummaryCard label="Hanya Excel bulan ini" value={data ? currentSummary.HANYA_EXCEL : "—"} tone="amber" icon="◐" /></section>
      <section className="content-panel"><div className="panel-toolbar"><div><h2>{currentPeriod ? monthLabel(currentPeriod) : "Belum ada hasil"}</h2><p className="muted">{data ? `${filteredRows.length} COA ditampilkan dari Excel` : "Upload data lalu jalankan audit"}</p></div></div><div className="category-tabs" role="tablist" aria-label="Kategori laporan">{categories.map((item) => <button key={item} className={`category-tab ${category === item ? "selected" : ""}`} onClick={() => setCategory(item)}>{item}</button>)}</div><div className="table-wrap"><table><thead><tr><th><button className="sort-button" onClick={() => changeSort("label")}>Akun / COA Excel{sortLabel("label")}</button></th><th><button className="sort-button" onClick={() => changeSort("category")}>Kategori{sortLabel("category")}</button></th><th><button className="sort-button" onClick={() => changeSort("value")}>Excel{sortLabel("value")}</button></th><th><button className="sort-button" onClick={() => changeSort("pdf_value")}>PDF{sortLabel("pdf_value")}</button></th><th><button className="sort-button" onClick={() => changeSort("diff")}>Selisih{sortLabel("diff")}</button></th><th><button className="sort-button" onClick={() => changeSort("status")}>Status{sortLabel("status")}</button></th></tr></thead><tbody>{filteredRows.length ? filteredRows.map((row, index) => <tr key={`${row.period}-${row.sheet}-${row.label}-${index}`}><td className="account">{row.label}</td><td className="category-cell">{sheetNames[row.sheet] ?? row.sheet}</td><td>{formatNumber(row.value)}</td><td>{formatNumber(row.pdf_value)}</td><td>{formatNumber(row.diff)}</td><td><span className={`status status-${row.status.toLowerCase()}`}><i />{statusLabel[row.status]}</span></td></tr>) : <tr><td colSpan={6} className="empty-state">Belum ada COA untuk ditampilkan.</td></tr>}</tbody></table></div></section></> : <section className="content-panel simple-page"><div className="panel-toolbar"><div><h2>{activeView === "history" ? "Belum ada riwayat tersimpan" : "Admin"}</h2><p className="muted">{activeView === "history" ? "Riwayat akan muncul setelah audit dijalankan." : "User lokal aktif pada workspace ini."}</p></div></div></section>}
    </section>
    <aside className="month-rail"><label className="period-label" htmlFor="period-select">PERIODE</label><select id="period-select" value={currentPeriod} onChange={(event) => setPeriod(event.target.value)} disabled={!periods.length}><option value="">Belum ada periode</option>{periods.map((item) => { const beda = rows.filter((row) => row.period === item && row.status === "BEDA").length; return <option key={item} value={item}>{monthLabel(item)}{beda ? ` — ${beda} Beda` : " — Cocok"}</option>; })}</select></aside>
  </main>;
}

function SummaryCard({ label, value, tone, icon }: { label: string; value: string | number; tone: string; icon: string }) { return <div className={`summary-card ${tone}`}><span className="summary-icon">{icon}</span><div><strong>{value}</strong><span>{label}</span></div></div>; }
