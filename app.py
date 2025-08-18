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

app = FastAPI(title="NYC AMI Allocator", version="2.2")
templates = Jinja2Templates(directory="templates")

DB_PATH = os.environ.get("AMI_DB", "/data/ami_log.sqlite")  # attach a Railway volume to /data for persistence

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

def make_excel(full: pd.DataFrame, aff_br: pd.DataFrame, metrics: dict, mirror: Optional[pd.DataFrame]=None, mirror_sheet_name: str="Sheet1", write_back_mode: str="new") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        wb = writer.book
        fmt_pct = wb.add_format({"num_format": "0%"})

        if write_back_mode == "same" and mirror is not None:
            mirror.to_excel(writer, index=False, sheet_name=mirror_sheet_name)
            # Summary
            def grade(m):
                wavg = min(m["wavg"], 60.0)
                pct  = max(min(m["pct40"], 21.0), 20.0)
                return round(1000.0 - abs(60.0 - wavg)*50.0 - abs(20.0 - pct)*100.0, 3)
            rows=[]
            for scen in ["A","B","C"]:
                m = metrics[scen]
                rows.append({
                    "Scenario": scen,
                    "Affordable SF (total)": f"{m['aff_sf_total']:,.2f}",
                    "SF @ 40%": f"{m['sf_at_40']:,.2f}",
                    "% @ 40%": f"{m['pct40']:.3f}%",
                    "Weighted Avg AMI": f"{m['wavg']:.3f}%",
                    "Score": grade(m),
                })
            pd.DataFrame(rows).sort_values("Score", ascending=False).to_excel(writer, index=False, sheet_name="Summary")
        else:
            master = full.copy()
            master.to_excel(writer, index=False, sheet_name="Master")
            if "Master" in writer.sheets:
                for label in ["A","B","C"]:
                    colname = f"Assigned_AMI_{label}"
                    if colname in master.columns:
                        col = master.columns.get_loc(colname)
                        writer.sheets["Master"].set_column(col, col, 16, fmt_pct)

            for label in ["A","B","C"]:
                br = aff_br.copy()
                if f"Assigned_AMI_{label}" in br.columns:
                    br = br.rename(columns={f"Assigned_AMI_{label}":"Assigned_AMI"})
                br.to_excel(writer, index=False, sheet_name=f"Breakdown_{label}")
                ws = writer.sheets[f"Breakdown_{label}"]
                for cname in ["AMI","Assigned_AMI"]:
                    if cname in br.columns:
                        c = br.columns.get_loc(cname)
                        ws.set_column(c, c, 16, fmt_pct)

            def grade(m):
                wavg = min(m["wavg"], 60.0)
                pct  = max(min(m["pct40"], 21.0), 20.0)
                return round(1000.0 - abs(60.0 - wavg)*50.0 - abs(20.0 - pct)*100.0, 3)
            rows=[]
            for scen in ["A","B","C"]:
                m = metrics[scen]
                rows.append({
                    "Scenario": scen,
                    "Affordable SF (total)": f"{m['aff_sf_total']:,.2f}",
                    "SF @ 40%": f"{m['sf_at_40']:,.2f}",
                    "% @ 40%": f"{m['pct40']:.3f}%",
                    "Weighted Avg AMI": f"{m['wavg']:.3f}%",
                    "Score": grade(m),
                })
            pd.DataFrame(rows).sort_values("Score", ascending=False).to_excel(writer, index=False, sheet_name="Summary")

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

    mask_aff = ~full["Assigned_AMI_A"].isna()
    aff_only = full.loc[mask_aff].copy()

    def rows(label):
        res=[]
        for _, r in aff_only.iterrows():
            res.append({
                "APT": r.get("APT"),
                "FLOOR": None if pd.isna(r.get("FLOOR")) else int(r.get("FLOOR")),
                "BED": r.get("BED"),
                "NET_SF": float(r["NET SF"]),
                "Original_AMI": None if pd.isna(r.get("AMI")) else float(r.get("AMI")),
                "Assigned_AMI": float(r[f"Assigned_AMI_{label}"]),
            })
        return res

    def buckets(label):
        items = rows(label)
        out = {"40":[], "60":[], "70":[], "80":[], "90":[], "100":[]}
        for it in items:
            v = it["Assigned_AMI"]
            if abs(v-0.40)<1e-9: out["40"].append(it)
            elif abs(v-0.60)<1e-9: out["60"].append(it)
            elif abs(v-0.70)<1e-9: out["70"].append(it)
            elif abs(v-0.80)<1e-9: out["80"].append(it)
            elif abs(v-0.90)<1e-9: out["90"].append(it)
            elif abs(v-1.00)<1e-9: out["100"].append(it)
        return out

    return {
        "filename": file.filename,
        "metrics": metrics,
        "best": best_label,
        "A": buckets("A"),
        "B": buckets("B"),
        "C": buckets("C"),
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

    out_name = (f"{base}.xlsx" if write_back=="same" else f"{base} - AMI Scenarios (A,B,C).xlsx")
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
