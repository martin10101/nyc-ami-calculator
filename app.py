import io, os, time, sqlite3, logging
from typing import Optional

import numpy as np
import pandas as pd
from fastapi import FastAPI, UploadFile, File, Form, Request
from fastapi.responses import StreamingResponse, JSONResponse, HTMLResponse
from fastapi.templating import Jinja2Templates

from ami_core import allocate_with_scenarios

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("ami")

app = FastAPI(title="NYC AMI Allocator", version="3.0")
templates = Jinja2Templates(directory="templates")

DB_PATH = os.environ.get("AMI_DB", "/data/ami_log.sqlite")  # attach a Railway volume at /data

# ---------- loaders ----------

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
        try:
            from docx import Document
        except Exception as e:
            raise ValueError("DOCX support requires python-docx") from e
        doc = Document(bio)
        best=None
        for t in doc.tables:
            rows = [[cell.text.strip() for cell in r.cells] for r in t.rows]
            if not rows or len(rows)<2: continue
            df = pd.DataFrame(rows[1:], columns=rows[0])
            if any("sf" in str(c).lower().replace(" ","") for c in df.columns):
                best = df; break
            best = best or df
        if best is None:
            raise ValueError("No readable tables found in DOCX.")
        for c in best.columns:
            best[c] = pd.to_numeric(best[c], errors="ignore")
        return best
    raise ValueError("Unsupported file type. Upload .xlsx, .xlsm, .xls, .xlsb, .csv, or .docx")

# ---------- DB (Master log) ----------

def init_db():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    with sqlite3.connect(DB_PATH) as con:
        con.execute("""CREATE TABLE IF NOT EXISTS runs(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ts INTEGER, filename TEXT, scenario TEXT,
            aff_sf_total REAL, sf_at_40 REAL, pct40 REAL, wavg REAL
        )""")

def log_run(filename: str, metrics: dict):
    init_db()
    with sqlite3.connect(DB_PATH) as con:
        for scen, m in metrics.items():
            con.execute(
                "INSERT INTO runs(ts, filename, scenario, aff_sf_total, sf_at_40, pct40, wavg) VALUES (?,?,?,?,?,?,?)",
                (int(time.time()), filename, scen, m["aff_sf_total"], m["sf_at_40"], m["pct40"], m["wavg"])
            )

# ---------- Excel writer ----------

def make_excel(
    full: pd.DataFrame,
    aff_br: pd.DataFrame,
    metrics: dict,
    mirror: Optional[pd.DataFrame] = None,
    mirror_sheet_name: str = "Sheet1",
    write_back_mode: str = "new",
) -> bytes:
    """
    Always produce:
      - Mirror (if write_back_mode == "same") under the original sheet name
      - Master (All Columns)  -> the full table with Assigned_AMI_S1..R3 columns
      - Scenario_S1..Scenario_R3 -> affordable-only breakdown for each scenario
      - Summary                -> metrics + scores

    This way the workbook always contains ALL six scenarios, regardless of write mode.
    """
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        wb = writer.book
        fmt_pct = wb.add_format({"num_format": "0%"})
        labels = list(metrics.keys())  # ["S1","S2","S3","R1","R2","R3"]

        # 1) Mirror sheet (only if write-back-to-same)
        if write_back_mode == "same" and mirror is not None:
            mirror.to_excel(writer, index=False, sheet_name=mirror_sheet_name)
            # Simple percent formatting if an AMI column is present
            if "AMI" in mirror.columns:
                ws = writer.sheets[mirror_sheet_name]
                col = mirror.columns.get_loc("AMI")
                ws.set_column(col, col, 14, fmt_pct)

        # 2) Master (all columns, with every Assigned_AMI_* column)
        master_name = "Master (All Columns)"
        master = full.copy()
        master.to_excel(writer, index=False, sheet_name=master_name)
        ws_master = writer.sheets[master_name]
        # Format each Assigned_AMI_* as %
        for c in master.columns:
            if str(c).startswith("Assigned_AMI_"):
                col = master.columns.get_loc(c)
                ws_master.set_column(col, col, 18, fmt_pct)

        # 3) Per-scenario sheets: affordable-only breakdown
        #    For each label, we rename Assigned_AMI_<label> -> Assigned_AMI
        base_cols = [c for c in ["FLOOR", "APT", "BED", "NET SF", "AMI_RAW"] if c in aff_br.columns]
        for label in labels:
            br = aff_br.copy()
            # Keep only the base cols + assigned col for this scenario
            keep = base_cols + [f"Assigned_AMI_{label}"]
            br = br[[c for c in keep if c in br.columns]]
            br = br.rename(columns={f"Assigned_AMI_{label}": "Assigned_AMI"})
            sheet = f"Scenario_{label}"
            br.to_excel(writer, index=False, sheet_name=sheet)
            ws = writer.sheets[sheet]
            # Format AMI and AMI_RAW (if present)
            if "Assigned_AMI" in br.columns:
                ws.set_column(br.columns.get_loc("Assigned_AMI"), br.columns.get_loc("Assigned_AMI"), 14, fmt_pct)
            if "AMI_RAW" in br.columns:
                # AMI_RAW may not be numeric, so do not apply pct format here
                ws.set_column(br.columns.get_loc("AMI_RAW"), br.columns.get_loc("AMI_RAW"), 12)

        # 4) Summary sheet with scores
        def grade(m):
            wavg = min(m["wavg"], 60.0)
            pct  = max(min(m["pct40"], 21.0), 20.0)
            return round(1000.0 - abs(60.0 - wavg)*50.0 - abs(20.0 - pct)*100.0, 3)

        rows = []
        for scen in labels:
            m = metrics[scen]
            rows.append({
                "Scenario": scen,
                "Affordable SF (total)": f"{m['aff_sf_total']:,.2f}",
                "SF @ 40%": f"{m['sf_at_40']:,.2f}",
                "% @ 40%": f"{m['pct40']:.3f}%",
                "Weighted Avg AMI": f"{m['wavg']:.3f}%",
                "Score": grade(m),
            })
        pd.DataFrame(rows).sort_values("Score", ascending=False).to_excel(
            writer, index=False, sheet_name="Summary"
        )

    buf.seek(0)
    return buf.getvalue()

# ---------- UI ----------

@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.get("/health")
def health(): return {"status": "ok"}

# ---------- API ----------

@app.post("/preview")
async def preview(
    file: UploadFile = File(...),
    sheet: Optional[str] = Form(None),
    require_family_at_40: int = Form(0),
    spread_40_max_per_floor: Optional[int] = Form(None),
    exempt_top_k_floors: int = Form(0)
):
    content = await file.read()
    try:
        df = load_any_table(content, file.filename, sheet=sheet)
        full, aff_br, metrics, mirror_out, best_label = allocate_with_scenarios(
            df,
            require_family_at_40=bool(require_family_at_40),
            spread_40_max_per_floor=spread_40_max_per_floor,
            exempt_top_k_floors=exempt_top_k_floors,
        )
        logger.info("Metrics: %s", metrics)
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=400)

    # Build per-scenario buckets (40/60/70/80/90/100)
    mask_aff = ~full["Assigned_AMI_S1"].isna()  # any scenario column works
    aff_only = full.loc[mask_aff].copy()

    def buckets(label):
        out = {"40":[], "60":[], "70":[], "80":[], "90":[], "100":[]}
        for _, r in aff_only.iterrows():
            v = float(r[f"Assigned_AMI_{label}"])
            row = {
                "APT": r.get("APT"),
                "FLOOR": None if pd.isna(r.get("FLOOR")) else int(r.get("FLOOR")),
                "BED": r.get("BED"),
                "NET_SF": float(r["NET SF"]),
            }
            if abs(v-0.40)<1e-9: out["40"].append(row)
            elif abs(v-0.60)<1e-9: out["60"].append(row)
            elif abs(v-0.70)<1e-9: out["70"].append(row)
            elif abs(v-0.80)<1e-9: out["80"].append(row)
            elif abs(v-0.90)<1e-9: out["90"].append(row)
            elif abs(v-1.00)<1e-9: out["100"].append(row)
        return out

    labels = list(metrics.keys())
    scen_payload = {}
    for label in labels:
        scen_payload[label] = {
            "metrics": metrics[label],
            "buckets": buckets(label),
        }

    return {
        "filename": file.filename,
        "best": best_label,
        "scenarios": scen_payload
    }

@app.post("/allocate")
async def allocate(
    file: UploadFile = File(...),
    sheet: Optional[str] = Form(None),
    write_back: str = Form("new"),  # "new" or "same"
    require_family_at_40: int = Form(0),
    spread_40_max_per_floor: Optional[int] = Form(None),
    exempt_top_k_floors: int = Form(0)
):
    content = await file.read()
    try:
        df = load_any_table(content, file.filename, sheet=sheet)
        full, aff_br, metrics, mirror_out, best_label = allocate_with_scenarios(
            df,
            require_family_at_40=bool(require_family_at_40),
            spread_40_max_per_floor=spread_40_max_per_floor,
            exempt_top_k_floors=exempt_top_k_floors,
        )
        logger.info("Metrics: %s", metrics)
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=400)

    base = os.path.splitext(os.path.basename(file.filename))[0]
    try: log_run(base, metrics)
    except Exception: pass

    mirror_sheet_name = sheet or "Sheet1"
    xlsx = make_excel(
        full, aff_br, metrics,
        mirror=mirror_out if write_back == "same" else None,
        mirror_sheet_name=mirror_sheet_name,
        write_back_mode=write_back
    )

    out_name = (f"{base}.xlsx" if write_back=="same" else f"{base} - AMI Scenarios (S1,S2,S3,R1,R2,R3).xlsx")
    return StreamingResponse(io.BytesIO(xlsx),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename=\"{out_name}\"'}
    )

@app.get("/export_master")
def export_master():
    try:
        init_db()
        with sqlite3.connect(DB_PATH) as con:
            df = pd.read_sql_query("SELECT * FROM runs ORDER BY ts DESC", con)
    except Exception as e:
        return JSONResponse({"error": f"Cannot read master DB: {e}"}, status_code=400)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
        df.to_excel(w, index=False, sheet_name="All Runs")
    buf.seek(0)
    return StreamingResponse(buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="AMI Master All Results.xlsx"'}
    )
