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
      - (optional) Mirror (if write_back_mode == "same") under original sheet name
      - Master (All Columns)  -> the full table with Assigned_AMI_* columns
      - Scenario_* (one per scenario, affordable-only subset)
      - Comparison            -> per-scenario counts/SF by AMI tier + family (2BR+) @ 40% check
      - Summary               -> metrics + scores (unchanged)

    This function is a drop-in; no other code changes are required.
    """
    import numpy as np
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        wb = writer.book
        fmt_pct = wb.add_format({"num_format": "0%"})
        fmt_int = wb.add_format({"num_format": "0"})
        fmt_sf  = wb.add_format({"num_format": "#,##0.00"})
        labels = list(metrics.keys())  # e.g., ["S1","S2","S3","R1","R2","R3"]

        # 1) Mirror sheet (only if write-back-to-same)
        if write_back_mode == "same" and mirror is not None:
            mirror.to_excel(writer, index=False, sheet_name=mirror_sheet_name)
            if "AMI" in mirror.columns:
                ws = writer.sheets[mirror_sheet_name]
                col = mirror.columns.get_loc("AMI")
                ws.set_column(col, col, 14, fmt_pct)

        # 2) Master (all columns, with every Assigned_AMI_* column)
        master_name = "Master (All Columns)"
        master = full.copy()
        master.to_excel(writer, index=False, sheet_name=master_name)
        ws_master = writer.sheets[master_name]
        for c in master.columns:
            if str(c).startswith("Assigned_AMI_"):
                col = master.columns.get_loc(c)
                ws_master.set_column(col, col, 18, fmt_pct)

        # 3) Per-scenario sheets: affordable-only breakdown
        base_cols = [c for c in ["FLOOR", "APT", "BED", "NET SF", "AMI_RAW"] if c in aff_br.columns]
        for label in labels:
            br = aff_br.copy()
            keep = base_cols + [f"Assigned_AMI_{label}"]
            keep = [c for c in keep if c in br.columns]
            br = br[keep]
            br = br.rename(columns={f"Assigned_AMI_{label}": "Assigned_AMI"})
            sheet = f"Scenario_{label}"
            br.to_excel(writer, index=False, sheet_name=sheet)
            ws = writer.sheets[sheet]
            if "Assigned_AMI" in br.columns:
                ws.set_column(br.columns.get_loc("Assigned_AMI"), br.columns.get_loc("Assigned_AMI"), 14, fmt_pct)

        # 4) Comparison (NEW): counts & SF by AMI tier + family@40 check
        comp_rows = []
        # Prepare arrays once per scenario
        has_bed = "BED" in aff_br.columns
        aff_sf = aff_br["NET SF"].astype(float).to_numpy() if "NET SF" in aff_br.columns else None
        for label in labels:
            col = f"Assigned_AMI_{label}"
            if col not in aff_br.columns or aff_sf is None:
                # minimal fallback if something is missing
                m = metrics[label]
                comp_rows.append({
                    "Scenario": label,
                    "WAvg_AMI_%": m["wavg"],
                    "Pct_40%_%": m["pct40"],
                    "Aff_SF_Total": m["aff_sf_total"],
                    "SF_at_40%": m["sf_at_40"],
                    "Units@40": None, "SF@40": None,
                    "Units@60": None, "SF@60": None,
                    "Units@70": None, "SF@70": None,
                    "Units@80": None, "SF@80": None,
                    "Units@90": None, "SF@90": None,
                    "Units@100": None, "SF@100": None,
                    "Family2BR+_Units@40": None
                })
                continue

            assigned = aff_br[col].astype(float).to_numpy()
            def cnt_and_sf(val):
                mask = np.isclose(assigned, val)
                return int(mask.sum()), float(aff_sf[mask].sum())

            u40, sf40 = cnt_and_sf(0.4)
            u60, sf60 = cnt_and_sf(0.6)
            u70, sf70 = cnt_and_sf(0.7)
            u80, sf80 = cnt_and_sf(0.8)
            u90, sf90 = cnt_and_sf(0.9)
            u100, sf100 = cnt_and_sf(1.0)

            fam_40 = None
            if has_bed:
                beds = pd.to_numeric(aff_br["BED"], errors="coerce").fillna(0).to_numpy()
                fam_40 = int(((beds >= 2) & np.isclose(assigned, 0.4)).sum())

            m = metrics[label]
            comp_rows.append({
                "Scenario": label,
                "WAvg_AMI_%": m["wavg"],
                "Pct_40%_%": m["pct40"],
                "Aff_SF_Total": m["aff_sf_total"],
                "SF_at_40%": m["sf_at_40"],
                "Units@40": u40, "SF@40": sf40,
                "Units@60": u60, "SF@60": sf60,
                "Units@70": u70, "SF@70": sf70,
                "Units@80": u80, "SF@80": sf80,
                "Units@90": u90, "SF@90": sf90,
                "Units@100": u100, "SF@100": sf100,
                "Family2BR+_Units@40": fam_40
            })

        comp_df = pd.DataFrame(comp_rows)
        comp_df.to_excel(writer, index=False, sheet_name="Comparison")
        ws_comp = writer.sheets["Comparison"]

        # Format numeric columns
        num_cols = [c for c in comp_df.columns if c not in ["Scenario"]]
        for c in num_cols:
            col_idx = comp_df.columns.get_loc(c)
            # percent-like columns
            if c in ["WAvg_AMI_%", "Pct_40%_%"]:
                ws_comp.set_column(col_idx, col_idx, 12, None)  # values already in %
            elif c.startswith("Units@") or c == "Family2BR+_Units@40":
                ws_comp.set_column(col_idx, col_idx, 12, fmt_int)
            else:
                ws_comp.set_column(col_idx, col_idx, 14, fmt_sf)

        # 5) Summary sheet (existing scoring logic)
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
