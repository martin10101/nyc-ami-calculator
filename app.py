import io, os, time, sqlite3
from typing import Optional, Dict, Any
import pandas as pd
from fastapi import FastAPI, UploadFile, File, Form, Request
from fastapi.responses import StreamingResponse, HTMLResponse, JSONResponse
from fastapi.templating import Jinja2Templates

# NEW: allocate_with_scenarios now has fast/slow toggle
from ami_core import allocate_with_scenarios

app = FastAPI(title="NYC AMI Allocator", version="3.1")
templates = Jinja2Templates(directory="templates")

DB_PATH = os.environ.get("AMI_DB", "/data/ami_log.sqlite")

# ---------- DB for Master All Results ----------
def init_db():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    with sqlite3.connect(DB_PATH) as con:
        con.execute("""CREATE TABLE IF NOT EXISTS runs(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ts INTEGER, filename TEXT, scenario TEXT,
            aff_sf_total REAL, sf_at_40 REAL, pct40 REAL, wavg REAL
        )""")

def log_run(filename: str, metrics: Dict[str, Any]):
    init_db()
    with sqlite3.connect(DB_PATH) as con:
        for scen, m in metrics.items():
            con.execute(
                "INSERT INTO runs(ts, filename, scenario, aff_sf_total, sf_at_40, pct40, wavg) VALUES (?,?,?,?,?,?,?)",
                (int(time.time()), filename, scen, m["aff_sf_total"], m["sf_at_40"], m["pct40"], m["wavg"])
            )

# ---------- File loader ----------
def load_any_table(file_bytes: bytes, filename: str, sheet: Optional[str] = None) -> pd.DataFrame:
    name = filename.lower()
    bio = io.BytesIO(file_bytes)
    if name.endswith((".xlsx",".xlsm")):
        xls = pd.ExcelFile(bio)
        sh = sheet or xls.sheet_names[0]
        return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sh)
    if name.endswith(".xls"):
        xls = pd.ExcelFile(bio)
        sh = sheet or xls.sheet_names[0]
        return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sh, engine="xlrd")
    if name.endswith(".xlsb"):
        xls = pd.ExcelFile(bio, engine="pyxlsb")
        sh = sheet or xls.sheet_names[0]
        return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sh, engine="pyxlsb")
    if name.endswith(".csv"):
        return pd.read_csv(io.BytesIO(file_bytes))
    if name.endswith(".docx"):
        from docx import Document
        doc = Document(bio)
        best=None
        for t in doc.tables:
            rows = [[c.text.strip() for c in r.cells] for r in t.rows]
            if not rows or len(rows)<2: continue
            df = pd.DataFrame(rows[1:], columns=rows[0])
            if any("sf" in str(c).lower().replace(" ","") for c in df.columns):
                best = df; break
            best = best or df
        if best is None: raise ValueError("No readable tables in DOCX.")
        for c in best.columns: best[c] = pd.to_numeric(best[c], errors="ignore")
        return best
    raise ValueError("Unsupported file type. Upload .xlsx/.xlsm/.xls/.xlsb/.csv/.docx")

# ---------- UI ----------
@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.get("/health")
def health(): return {"ok": True}

# ---------- API ----------
@app.post("/preview")
async def preview(
    file: UploadFile = File(...),
    sheet: Optional[str] = Form(None),
    require_family_at_40: int = Form(0),
    spread_40_max_per_floor: Optional[int] = Form(None),
    exempt_top_k_floors: int = Form(0),
    top_k: int = Form(3),
    fast_preview: int = Form(1)  # NEW: 1 = heuristic only; 0 = allow MILP
):
    content = await file.read()
    try:
        df = load_any_table(content, file.filename, sheet)
        # Fast preview uses heuristic only to keep UI snappy
        full, aff_br, metrics, mirror, best = allocate_with_scenarios(
            df,
            require_family_at_40=bool(require_family_at_40),
            spread_40_max_per_floor=spread_40_max_per_floor,
            exempt_top_k_floors=exempt_top_k_floors,
            return_top_k=top_k,
            use_milp=bool(1 - fast_preview)  # 0 => heuristic; 1 => MILP
        )
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=400)

    # Compact JSON by band for each scenario
    labels = [c for c in full.columns if c.startswith("Assigned_AMI_")]
    aff_only = full.loc[full[labels[0]].notna()].copy()
    def buckets(label):
        col = f"Assigned_AMI_{label}"
        out = {"40":[],"60":[],"70":[],"80":[],"90":[],"100":[]}
        for _,r in aff_only.iterrows():
            v = float(r[col]); key = {0.4:"40",0.6:"60",0.7:"70",0.8:"80",0.9:"90",1.0:"100"}[round(v,1)]
            out[key].append({
                "APT": r.get("APT"),
                "FLOOR": None if pd.isna(r.get("FLOOR")) else int(r.get("FLOOR")),
                "BED": r.get("BED"),
                "NET_SF": float(r["NET SF"])
            })
        return out

    resp = {
        "filename": file.filename,
        "best": best,
        "metrics": metrics,
        "scenarios": {lab.replace("Assigned_AMI_",""): buckets(lab.replace("Assigned_AMI_","")) for lab in labels}
    }
    return resp

@app.post("/allocate")
async def allocate(
    file: UploadFile = File(...),
    sheet: Optional[str] = Form(None),
    write_back: str = Form("new"),  # "new" or "same"
    require_family_at_40: int = Form(0),
    spread_40_max_per_floor: Optional[int] = Form(None),
    exempt_top_k_floors: int = Form(0),
    top_k: int = Form(3)
):
    content = await file.read()
    try:
        df = load_any_table(content, file.filename, sheet)
        # Full solve (MILP) for the downloadable workbook
        full, aff_br, metrics, mirror, best = allocate_with_scenarios(
            df,
            require_family_at_40=bool(require_family_at_40),
            spread_40_max_per_floor=spread_40_max_per_floor,
            exempt_top_k_floors=exempt_top_k_floors,
            return_top_k=top_k,
            use_milp=True
        )
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=400)

    # Log Master All Results
    base = os.path.splitext(os.path.basename(file.filename))[0]
    try: log_run(base, metrics)
    except Exception: pass

    # Build workbook —— IMPORTANT: engine="xlsxwriter" (lowercase)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
        full.to_excel(w, index=False, sheet_name="Master")
        for lab in [c.replace("Assigned_AMI_","") for c in full.columns if c.startswith("Assigned_AMI_")]:
            # Write only the affordable rows with the scenario column
            cols = [c for c in aff_br.columns if not c.startswith("Assigned_AMI_")]
            br = aff_br[cols].copy()
            br[f"Assigned_AMI_{lab}"] = full.loc[full[f"Assigned_AMI_{lab}"].notna(), f"Assigned_AMI_{lab}"].to_numpy()
            br.to_excel(w, index=False, sheet_name=f"Breakdown_{lab}")
        pd.DataFrame([
            {"Scenario": k, **v} for k,v in metrics.items()
        ]).to_excel(w, index=False, sheet_name="Summary")

    buf.seek(0)
    out_name = (f"{base}.xlsx" if write_back=="same" else f"{base} - AMI Scenarios.xlsx")
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{out_name}"'}
    )

@app.get("/export_master")
def export_master():
    try:
        init_db()
        with sqlite3.connect(DB_PATH) as con:
            df = pd.read_sql_query("SELECT * FROM runs ORDER BY ts DESC", con)
    except Exception as e:
        return JSONResponse({"error": f"Cannot read DB: {e}"}, status_code=400)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:  # lowercase engine
        df.to_excel(w, index=False, sheet_name="All Runs")
    buf.seek(0)
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="AMI Master All Results.xlsx"'}
    )
