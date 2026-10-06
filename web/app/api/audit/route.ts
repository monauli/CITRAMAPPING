import { randomUUID } from "node:crypto";
import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { spawn } from "node:child_process";
import { NextResponse } from "next/server";
import { createClient } from "@supabase/supabase-js";

export const runtime = "nodejs";

function runPython(args: string[], cwd: string) {
  return new Promise<string>((resolve, reject) => {
    const command = process.env.MAPPING_PYTHON ?? "python";
    const child = spawn(command, args, { cwd, windowsHide: true });
    let stdout = "";
    let stderr = "";
    child.stdout.on("data", (chunk) => { stdout += chunk; });
    child.stderr.on("data", (chunk) => { stderr += chunk; });
    child.on("error", reject);
    child.on("close", (code) => {
      if (code === 0) resolve(stdout);
      else reject(new Error(stderr.trim() || `Python berhenti dengan kode ${code}`));
    });
  });
}

export async function POST(request: Request) {
  try {
    const token = request.headers.get("authorization")?.replace(/^Bearer\s+/i, "");
    const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL;
    const supabaseKey = process.env.NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY;
    if (!token || !supabaseUrl || !supabaseKey) return NextResponse.json({ error: "Login Supabase diperlukan sebelum menjalankan audit." }, { status: 401 });
    const supabase = createClient(supabaseUrl, supabaseKey, { global: { headers: { Authorization: `Bearer ${token}` } } });
    const { data: userData, error: userError } = await supabase.auth.getUser(token);
    if (userError || !userData.user) return NextResponse.json({ error: "Sesi login Supabase tidak valid." }, { status: 401 });
    const form = await request.formData();
    const files = form.getAll("files").filter((entry): entry is File => entry instanceof File);
    if (!files.length) return NextResponse.json({ error: "Pilih folder yang berisi Excel dan PDF." }, { status: 400 });

    const root = path.join(process.env.LOCALAPPDATA ?? process.cwd(), "MappingKeuangan", "uploads", randomUUID());
    await mkdir(root, { recursive: true });
    for (const file of files) {
      const relative = file.name.replaceAll("\\", "/").split("/").filter((part) => part && part !== "." && part !== "..").join("/");
      const destination = path.join(root, path.basename(relative || "upload.bin"));
      await mkdir(path.dirname(destination), { recursive: true });
      await writeFile(destination, Buffer.from(await file.arrayBuffer()));
    }

    const workspace = path.resolve(process.cwd(), "..");
    const output = await runPython([
      path.join(workspace, "map_keuangan.py"), "--json", "--no-selftest", "--pdf-dir", root,
    ], workspace);
    const result = JSON.parse(output) as { excel: string; periods: string[]; missing_periods: string[]; summary: Record<string, number>; rows: Array<{ month: string; period: string; sheet: string; label: string; value: number | null; pdf_value: number | null; diff: number | null; status: string }> };
    const { data: run, error: runError } = await supabase.from("mapping_runs").insert({ name: `Audit ${new Date().toISOString()}`, source_excel: result.excel, summary: result.summary, created_by: userData.user.id }).select("id").single();
    if (runError || !run) throw new Error(runError?.message || "Gagal menyimpan riwayat audit.");
    for (const period of result.periods) {
      const { data: periodRow, error: periodError } = await supabase.from("mapping_periods").insert({ run_id: run.id, period, month_name: period, summary: { rows: result.rows.filter((row) => row.period === period).length } }).select("id").single();
      if (periodError || !periodRow) throw new Error(periodError?.message || `Gagal menyimpan periode ${period}.`);
      const rows = result.rows.filter((row) => row.period === period).map((row) => ({ period_id: periodRow.id, sheet: row.sheet, label: row.label, excel_value: row.value, pdf_value: row.pdf_value, difference: row.diff, status: row.status }));
      for (let index = 0; index < rows.length; index += 500) {
        const { error: rowsError } = await supabase.from("mapping_rows").insert(rows.slice(index, index + 500));
        if (rowsError) throw new Error(rowsError.message);
      }
    }
    return NextResponse.json({ ...result, run_id: run.id });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Audit gagal dijalankan.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
