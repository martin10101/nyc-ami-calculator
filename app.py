import io
import os
import time
import sqlite3
from typing import Optional, Dict, Any

import pandas as pd

# --- Excel writer engine registration (Railway‑safe) ---
# Pandas uses openpyxl by default for .xlsx files, but we
# explicitly register the xlsxwriter engine because it is
# available and avoids case‑sensitivity issues in Railway.
try:
    import xlsxwriter as _xlsxwriter  # module name is lowercase
    from pandas.io.excel import _excel_writer
    _excel_writer._writers['xlsxwriter'] = _xlsxwriter
except Exception:
    pass
# ---------------------------------------------------------

from fastapi import FastAPI, UploadFile, File, Form, Request
from fastapi.responses import StreamingResponse, HTMLResponse, JSONResponse
from fastapi.templating import Jinja2Templates

# Import the core allocator.  It exposes allocate_with_scenarios,
# which returns the full dataset with assigned AMI columns,
# a subset with only affordable rows, a metrics dict, a mirror
# table with footer, and the label of the best scenario.
from ami_core import allocate_with_scenarios


# Initialise FastAPI and templates
app = FastAPI(title="NYC AMI Allocator", version="3.2")
templates = Jinja2Templates(directory="templates")

# Path for logging runs (Master All Results).  Railway
# mounts /data as a persistent volume if configured.
DB_PATH = os.environ.get("AMI_DB", "/data/ami_log.sqlite")


# ---------- DB helpers ----------
def init_db() -> None:
    """Initialise the sqlite database for recording runs."""
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    with sqlite3.connect(DB_PATH) as con:
        con.execute(
            """CREATE TABLE IF NOT EXISTS runs(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ts INTEGER, filename TEXT, scenario TEXT,
                aff_sf_total REAL, sf_at_40 REAL, pct40 REAL, wavg REAL
            )"""
        )


def log_run(filename: str, metrics: Dict[str, Any]) -> None:
    """Insert one row per scenario into the runs table."""
    init_db()
    with sqlite3.connect(DB_PATH) as con:
        for scen, m in metrics.items():
            con.execute(
                "INSERT INTO runs(ts, filename, scenario, aff_sf_total, sf_at_40, pct40, wavg)"
                " VALUES (?,?,?,?,?,?,?)",
                (
                    int(time.time()),
                    filename,
                    scen,
                    m.get("aff_sf_total", 0.0),
                    m.get("sf_at_40", 0.0),
                    m.get("pct40", 0.0),
                    m.get("wavg", 0.0),
                ),
            )


# ---------- File loader ----------
def load_any_table(
    file_bytes: bytes, filename: str, sheet: Optional[str] = None
) -> pd.DataFrame:
    """
    Load a table from various formats (.xlsx, .xlsm, .xlsb, .xls, .csv, .docx).
    """
    name = filename.lower()
    bio = io.BytesIO(file_bytes)
    if name.endswith((".xlsx", ".xlsm")):
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
        best: Optional[pd.DataFrame] = None
        for table in doc.tables:
            rows = [[c.text.strip() for c in r.cells] for r in table.rows]
            if not rows or len(rows) < 2:
                continue
            df = pd.DataFrame(rows[1:], columns=rows[0])
            # Prefer a table with a column that looks like a square footage column
            if any("sf" in str(c).lower().replace(" ", "") for c in df.columns):
                best = df
                break
            best = best or df
        if best is None:
            raise ValueError("No readable tables in DOCX.")
        # Convert numeric columns where possible
        for c in best.columns:
            try:
                best[c] = pd.to_numeric(best[c], errors="ignore")
            except Exception:
                pass
        return best
    raise ValueError(
        "Unsupported file type. Upload .xlsx/.xlsm/.xls/.xlsb/.csv/.docx"
    )


# ---------- UI routes ----------
@app.get("/", response_class=HTMLResponse)
async def root(request: Request):
    """Render the single‑page HTML interface."""
    return templates.TemplateResponse("index.html", {"request": request})


@app.get("/health")
async def health() -> Dict[str, bool]:
    """Simple health check endpoint."""
    return {"ok": True}


# ---------- Preview (fast heuristic or MILP) ----------
@app.post("/preview")
async def preview(
    file: UploadFile = File(...),
    sheet: Optional[str] = Form(None),
    require_family_at_40: int = Form(0),
    spread_40_max_per_floor: Optional[int] = Form(None),
    exempt_top_k_floors: int = Form(0),
    top_k: int = Form(3),
    fast_preview: int = Form(1),  # 1 = heuristic only, 0 = allow MILP
):
    """
    Produce a quick preview of up to `top_k` scenarios without generating an Excel file.
    """
    content = await file.read()
    try:
        df = load_any_table(content, file.filename, sheet)
        # Determine whether to run the exact solver
        use_milp = bool(1 - int(fast_preview))
        # Compute scenarios.  allocate_with_scenarios returns
        # (full, aff_br, metrics, mirror, best_label)
        full, aff_br, metrics, mirror, best = allocate_with_scenarios(
            df,
            require_family_at_40=bool(require_family_at_40),
            spread_40_max_per_floor=spread_40_max_per_floor,
            exempt_top_k_floors=exempt_top_k_floors,
            return_top_k=top_k,
            use_milp=use_milp,
        )
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=400)

    # Build per‑scenario buckets keyed by AMI band.  We only show
    # affordable units (those rows where Assigned_AMI_X is not NaN).
    labels = [c for c in full.columns if c.startswith("Assigned_AMI_")]
    # Filter the full DataFrame to affordable units using the first scenario column
    if not labels:
        return JSONResponse({"error": "No scenarios generated"}, status_code=400)
    
    aff_only = full.loc[full[labels[0]].notna()].copy()

    def buckets(label: str) -> Dict[str, List[Dict[str, Any]]]:
        col = f"Assigned_AMI_{label}"
        out = {"40": [], "60": [], "70": [], "80": [], "90": [], "100": []}
        for _, r in aff_only.iterrows():
            v = float(r[col])
            key = {
                0.4: "40",
                0.6: "60",
                0.7: "70",
                0.8: "80",
                0.9: "90",
                1.0: "100",
            }.get(round(v, 1), "60")  # Default to 60 if not found
            out[key].append(
                {
                    "APT": r.get("APT"),
                    "FLOOR": None
                    if pd.isna(r.get("FLOOR"))
                    else int(r.get("FLOOR")),
                    "BED": r.get("BED"),
                    "NET_SF": float(r["NET SF"]),
                }
            )
        return out

    # Assemble response.  Each scenario is keyed by its label (S1, S2, …).
    resp = {
        "filename": file.filename,
        "best": best,
        "metrics": metrics,  # This is the key fix - include metrics in response
        "scenarios": {
            lab.replace("Assigned_AMI_", ""): buckets(lab.replace("Assigned_AMI_", ""))
            for lab in labels
        },
    }
    return resp


# ---------- Full allocate (generate workbook) ----------
@app.post("/allocate")
async def allocate(
    file: UploadFile = File(...),
    sheet: Optional[str] = Form(None),
    write_back: str = Form("new"),  # "new" to rename, "same" to overwrite base filename
    require_family_at_40: int = Form(0),
    spread_40_max_per_floor: Optional[int] = Form(None),
    exempt_top_k_floors: int = Form(0),
    top_k: int = Form(3),
) -> StreamingResponse:
    """
    Solve the allocation problem (using MILP) and return an Excel workbook.
    """
    content = await file.read()
    try:
        df = load_any_table(content, file.filename, sheet)
        full, aff_br, metrics, mirror, best = allocate_with_scenarios(
            df,
            require_family_at_40=bool(require_family_at_40),
            spread_40_max_per_floor=spread_40_max_per_floor,
            exempt_top_k_floors=exempt_top_k_floors,
            return_top_k=top_k,
            use_milp=True,
        )
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=400)

    # Log metrics for master results
    base = os.path.splitext(os.path.basename(file.filename))[0]
    try:
        log_run(base, metrics)
    except Exception:
        pass

    # Build the Excel workbook.  We use xlsxwriter to avoid openpyxl
    # dependency issues on Railway.
    buf = io.BytesIO()
    try:
        with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
            # Master sheet: full dataset with Assigned_AMI_X columns
            full.to_excel(writer, index=False, sheet_name="Master")
            
            # Breakdown sheets: only affordable rows and one scenario column
            scenario_labels = [c.replace("Assigned_AMI_", "") for c in full.columns if c.startswith("Assigned_AMI_")]
            for lab in scenario_labels:
                # Get affordable rows (non-null assignments)
                aff_mask = full[f"Assigned_AMI_{lab}"].notna()
                if not aff_mask.any():
                    continue
                    
                # Create breakdown with base columns plus this scenario's assignments
                base_cols = [c for c in ["FLOOR","APT","BED","NET SF","AMI_RAW"] if c in full.columns]
                br = full.loc[aff_mask, base_cols].copy()
                br[f"Assigned_AMI_{lab}"] = full.loc[aff_mask, f"Assigned_AMI_{lab}"]
                br.to_excel(writer, index=False, sheet_name=f"Breakdown_{lab}")
            
            # Summary sheet: scenario metrics as rows
            summary_rows = []
            for k, v in metrics.items():
                row = {"Scenario": k}
                row.update(v)
                summary_rows.append(row)
            if summary_rows:
                pd.DataFrame(summary_rows).to_excel(
                    writer, index=False, sheet_name="Summary"
                )
    except Exception as e:
        return JSONResponse({"error": f"Excel generation failed: {str(e)}"}, status_code=500)

    buf.seek(0)
    out_name = (
        f"{base}.xlsx" if write_back == "same" else f"{base} - AMI Scenarios.xlsx"
    )
    return StreamingResponse(
        io.BytesIO(buf.read()),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{out_name}"'},
    )


@app.get("/export_master")
async def export_master() -> StreamingResponse:
    """
    Download the log of all runs (Master All Results).
    """
    try:
        init_db()
        with sqlite3.connect(DB_PATH) as con:
            df = pd.read_sql_query("SELECT * FROM runs ORDER BY ts DESC", con)
    except Exception as e:
        return JSONResponse({"error": f"Cannot read DB: {e}"}, status_code=400)
    
    buf = io.BytesIO()
    try:
        with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="All Runs")
    except Exception as e:
        return JSONResponse({"error": f"Excel generation failed: {str(e)}"}, status_code=500)
    
    buf.seek(0)
    return StreamingResponse(
        io.BytesIO(buf.read()),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="AMI Master All Results.xlsx"'},
    )

