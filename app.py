import io, os, time, sqlite3
from typing import Optional, Dict, Any
import pandas as pd

# --- Excel writer engine registration (Railway-safe) ---
import xlsxwriter  # lowercase
from pandas.io.excel import register_writer
register_writer(xlsxwriter.Workbook)
# --------------------------------------------------------

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
            if any("sf" in st.lower() for st in df.columns): best=df
        if best is not None: return best
        raise ValueError("No table found in DOCX")
    raise ValueError("Unsupported file type")

# ---------- UI ----------
@app.get("/", response_class=HTMLResponse)
async def root(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

# ---------- Preview (fast heuristic) ----------
@app.post("/preview")
async def preview(
    file: UploadFile = File(...),
    sheet: Optional[str] = Form(None),
    require_family_at_40: Optional[str] = Form("0"),
    spread_40_max_per_floor: Optional[int] = Form(None),
    exempt_top_k_floors: Optional[int] = Form(0),
    top_k: Optional[int] = Form(3),
):
    content = await file.read()
    try:
        df = load_any_table(content, file.filename, sheet)
        _, _, metrics, _, best = allocate_with_scenarios(
            df,
            require_family_at_40=bool(int(require_family_at_40)),
            spread_40_max_per_floor=spread_40_max_per_floor,
            exempt_top_k_floors=exempt_top_k_floors,
            return_top_k=top_k,
            use_milp=False  # fast
        )
        return {"scenarios": {k: v for k, v in metrics.items()}, "best": best}
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=400)

# ---------- Full allocate (MILP + Excel) ----------
@app.post("/allocate")
async def allocate(
    file: UploadFile = File(...),
    sheet: Optional[str] = Form(None),
    require_family_at_40: Optional[str] = Form("0"),
    spread_40_max_per_floor: Optional[int] = Form(None),
    exempt_top_k_floors: Optional[int] = Form(0),
    top_k: Optional[int] = Form(3),
    write_back: Optional[str] = Form("same"),
):
    content = await file.read()
    try:
        df = load_any_table(content, file.filename, sheet)
        # Full solve (MILP) for the downloadable workbook
        full, aff_br, metrics, mirror_out, best = allocate_with_scenarios(
            df,
            require_family_at_40=bool(int(require_family_at_40)),
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

    # Build workbook to match sample
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
        mirror_out.to_excel(w, index=False, sheet_name="Sheet1")  # mirror with footer
        full.to_excel(w, index=False, sheet_name="Master (All Columns)")
        for lab in list(metrics.keys()):
            cols = [c for c in aff_br.columns if not c.startswith("Assigned_AMI_") or c == f"Assigned_AMI_{lab}"]
            br = aff_br[cols].copy()
            br = br.rename(columns={f"Assigned_AMI_{lab}": "Assigned_AMI"})
            br.to_excel(w, index=False, sheet_name=f"Scenario_{lab}")
        sum_df = pd.DataFrame([
            {"Scenario": k, "Affordable SF (total)": v["aff_sf_total"], "SF @ 40%": v["sf_at_40"], 
             "% @ 40%": v["pct40"], "Weighted Avg AMI": v["wavg"], "Score": v.get("score", 0)} 
            for k,v in metrics.items()
        ])
        sum_df.to_excel(w, index=False, sheet_name="Summary")

    buf.seek(0)
    out_name = f"{base} - AMI Scenarios.xlsx"
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
    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:  
        df.to_excel(w, index=False, sheet_name="All Runs")
    buf.seek(0)
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="AMI Master All Results.xlsx"'}
    )
