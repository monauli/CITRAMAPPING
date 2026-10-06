import json
import os
import shutil
import tempfile
import urllib.error
import urllib.request
from pathlib import Path

from fastapi import FastAPI, File, Header, HTTPException, UploadFile
from fastapi.responses import JSONResponse

from map_keuangan import json_report

app = FastAPI()


def supabase_request(url: str, method: str, token: str, payload=None):
    body = None if payload is None else json.dumps(payload).encode("utf-8")
    request = urllib.request.Request(url, data=body, method=method, headers={
        "apikey": os.environ["NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY"],
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "Prefer": "return=representation",
    })
    try:
        with urllib.request.urlopen(request, timeout=20) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as error:
        detail = error.read().decode("utf-8", errors="replace")
        raise HTTPException(status_code=502, detail=f"Supabase request gagal: {detail}") from error


def current_user(token: str):
    url = f"{os.environ['NEXT_PUBLIC_SUPABASE_URL']}/auth/v1/user"
    return supabase_request(url, "GET", token)


@app.post("/")
@app.post("/api/audit_serverless")
@app.post("/api/audit_serverless.py")
async def audit(files: list[UploadFile] = File(...), authorization: str | None = Header(default=None)):
    if not authorization or not authorization.lower().startswith("bearer "):
        raise HTTPException(status_code=401, detail="Login Supabase diperlukan sebelum menjalankan audit.")
    token = authorization.split(" ", 1)[1].strip()
    if not token or not os.environ.get("NEXT_PUBLIC_SUPABASE_URL") or not os.environ.get("NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY"):
        raise HTTPException(status_code=401, detail="Konfigurasi Supabase belum tersedia.")
    user = current_user(token)
    if not user.get("id"):
        raise HTTPException(status_code=401, detail="Sesi login Supabase tidak valid.")

    temp_dir = Path(tempfile.mkdtemp(prefix="mapping-audit-"))
    try:
        for upload in files:
            filename = Path(upload.filename or "upload.bin").name
            if filename in ("", ".", ".."):
                continue
            destination = temp_dir / filename
            destination.write_bytes(await upload.read())
        result = json_report(temp_dir / next(p.name for p in temp_dir.glob("*.xlsx")), temp_dir, "Report LapKeu CAM JO {month} {year}.pdf")

        base = os.environ["NEXT_PUBLIC_SUPABASE_URL"]
        run_rows = supabase_request(f"{base}/rest/v1/mapping_runs", "POST", token, {
            "name": f"Audit {result['excel']}", "source_excel": result["excel"], "summary": result["summary"], "created_by": user["id"]
        })
        run_id = run_rows[0]["id"]
        for period in result["periods"]:
            period_rows = supabase_request(f"{base}/rest/v1/mapping_periods", "POST", token, {
                "run_id": run_id, "period": period, "month_name": period,
                "summary": {"rows": sum(1 for row in result["rows"] if row["period"] == period)},
            })
            period_id = period_rows[0]["id"]
            rows = [{"period_id": period_id, "sheet": row["sheet"], "label": row["label"], "excel_value": row["value"], "pdf_value": row["pdf_value"], "difference": row["diff"], "status": row["status"]} for row in result["rows"] if row["period"] == period]
            for start in range(0, len(rows), 500):
                supabase_request(f"{base}/rest/v1/mapping_rows", "POST", token, rows[start:start + 500])
        return JSONResponse({**result, "run_id": run_id})
    except StopIteration as error:
        raise HTTPException(status_code=400, detail="Pilih folder yang berisi satu file Excel.") from error
    except HTTPException:
        raise
    except Exception as error:
        raise HTTPException(status_code=500, detail=f"Audit gagal dijalankan: {error}") from error
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
