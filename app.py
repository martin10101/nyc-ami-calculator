import io
import os
import time
import sqlite3
from typing import Optional, Dict, Any
from pathlib import Path

import pandas as pd

# --- Excel writer engine registration (Railway‑safe) ---
try:
    import xlsxwriter as _xlsxwriter
    from pandas.io.excel import _excel_writer
    _excel_writer._writers['xlsxwriter'] = _xlsxwriter
except Exception:
    pass
# ---------------------------------------------------------

from fastapi import FastAPI, UploadFile, File, Form, Request
from fastapi.responses import StreamingResponse, HTMLResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles

# Import the core allocator
from ami_core import allocate_with_scenarios

# Initialize FastAPI
app = FastAPI(title="NYC AMI Allocator", version="3.2")

# Setup templates directory - create if it doesn't exist
templates_dir = Path("templates")
templates_dir.mkdir(exist_ok=True)

# Create index.html if it doesn't exist
index_html_path = templates_dir / "index.html"
if not index_html_path.exists():
    index_html_content = '''<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1"/>
  <title>NYC AMI Allocator</title>
  <style>
    :root{--bg:#0f1115;--card:#161a22;--text:#e8ecf1;--muted:#a9b4c2;--accent:#6d9eff;--success:#4ade80;--error:#ef4444;--warning:#f59e0b}
    *{box-sizing:border-box} 
    body{margin:0;background:var(--bg);color:var(--text);font:15px/1.4 system-ui,Segoe UI,Roboto}
    .wrap{max-width:1200px;margin:40px auto;padding:0 16px}
    .card{background:var(--card);border-radius:14px;padding:18px;margin-bottom:16px;box-shadow:0 0 0 1px #1e2430}
    label{display:block;margin:.4rem 0 .2rem;color:var(--muted);font-weight:500}
    input,select{width:100%;padding:10px 12px;border-radius:10px;border:1px solid #293245;background:#0e1218;color:var(--text);transition:border-color 0.2s}
    input:focus,select:focus{outline:none;border-color:var(--accent)}
    .row{display:grid;gap:12px;grid-template-columns:repeat(auto-fit,minmax(250px,1fr))}
    .btn{display:inline-block;padding:12px 16px;border-radius:10px;border:none;background:var(--accent);color:#0b0f16;font-weight:600;cursor:pointer;transition:all 0.2s;text-decoration:none}
    .btn:hover{background:#5a8cff;transform:translateY(-1px)}
    .btn:disabled{background:#2b3550;color:var(--muted);cursor:not-allowed;transform:none}
    .btn.secondary{background:#20283a;color:var(--text)}
    .btn.secondary:hover{background:#2a3441}
    .grid{display:grid;gap:14px;grid-template-columns:repeat(auto-fit,minmax(320px,1fr))}
    .pill{display:inline-block;padding:6px 10px;margin:4px;border-radius:999px;background:#1e2533;border:1px solid #2b3550;color:#cfe0ff;font-size:12px;font-weight:500}
    .title{font-size:24px;font-weight:700;margin:0 0 10px;color:var(--text)}
    .subtitle{font-size:18px;font-weight:600;margin:12px 0 8px;color:var(--text)}
    .muted{color:var(--muted)}
    .success{color:var(--success)}
    .error{color:var(--error)}
    .warning{color:var(--warning)}
    .status{padding:12px;border-radius:8px;margin:8px 0;display:none}
    .status.loading{background:#1e3a8a;color:#93c5fd;display:block}
    .status.success{background:#065f46;color:#6ee7b7;display:block}
    .status.error{background:#7f1d1d;color:#fca5a5;display:block}
    .status.warning{background:#78350f;color:#fcd34d;display:block}
    .progress{width:100%;height:4px;background:#293245;border-radius:2px;overflow:hidden;margin:8px 0}
    .progress-bar{height:100%;background:var(--accent);width:0%;transition:width 0.3s;border-radius:2px}
    @media (max-width: 768px) {
      .wrap{padding:0 12px;margin:20px auto}
      .row{grid-template-columns:1fr}
      .grid{grid-template-columns:1fr}
      .btn{padding:10px 14px;font-size:14px}
    }
    .scenario-card{position:relative;overflow:hidden}
    .scenario-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px}
    .scenario-badge{background:var(--accent);color:#0b0f16;padding:4px 8px;border-radius:6px;font-size:11px;font-weight:600}
    .ami-section{margin-top:12px;padding-top:8px;border-top:1px solid #293245}
    .ami-header{font-size:13px;font-weight:600;color:var(--muted);margin-bottom:6px}
    input[type="file"]{padding:8px;background:#1a1f2e;border:2px dashed #293245;border-radius:8px}
    input[type="file"]:hover{border-color:var(--accent)}
    input[type="checkbox"]{width:auto;margin-right:8px}
    .btn-group{display:flex;gap:12px;flex-wrap:wrap;margin-top:16px}
    @media (max-width: 768px) {
      .btn-group{flex-direction:column}
      .btn-group .btn{margin-bottom:8px}
    }
  </style>
</head>
<body>
<div class="wrap">
  <h1 class="title">NYC AMI Allocator</h1>
  <p class="muted">Optimize affordable housing unit mix allocation for NYC development projects</p>
  
  <div class="card">
    <form id="f">
      <div class="row">
        <div>
          <label>Upload Unit Schedule File *</label>
          <input type="file" name="file" required accept=".xlsx,.xlsm,.xls,.xlsb,.csv,.docx"/>
          <small class="muted">Supported: Excel, CSV, Word documents</small>
        </div>
        <div>
          <label>Sheet Name (optional)</label>
          <input name="sheet" placeholder="Leave blank for first sheet"/>
        </div>
      </div>
      
      <div class="row">
        <div>
          <label>Require ≥1 family unit (2BR+) at 40% AMI</label>
          <select name="require_family_at_40">
            <option value="0">No</option>
            <option value="1">Yes</option>
          </select>
        </div>
        <div>
          <label>Max 40% AMI units per floor (optional)</label>
          <input name="spread_40_max_per_floor" type="number" min="1" placeholder="e.g., 2"/>
        </div>
      </div>
      
      <div class="row">
        <div>
          <label>Exempt top floors from 40% AMI (optional)</label>
          <input name="exempt_top_k_floors" type="number" min="0" placeholder="e.g., 2"/>
        </div>
        <div>
          <label>Number of scenarios to generate</label>
          <select name="top_k">
            <option value="3">3 scenarios</option>
            <option value="2">2 scenarios</option>
            <option value="1">1 scenario</option>
          </select>
        </div>
      </div>
      
      <div style="margin-top:12px">
        <label>
          <input type="checkbox" id="fast_preview" checked/> 
          Fast preview mode (uses heuristic algorithm for speed)
        </label>
      </div>
      
      <div id="status" class="status"></div>
      <div id="progress" class="progress" style="display:none">
        <div id="progress-bar" class="progress-bar"></div>
      </div>
      
      <div class="btn-group">
        <button class="btn" type="button" onclick="doPreview()" id="previewBtn">
          Preview Scenarios
        </button>
        <button class="btn secondary" type="button" onclick="doDownload()" id="downloadBtn">
          Download Excel Report
        </button>
        <a class="btn secondary" href="/export_master">
          Download Master Results
        </a>
      </div>
    </form>
  </div>

  <div id="out"></div>
</div>

<script>
function showStatus(message, type = 'loading') {
  const status = document.getElementById('status');
  status.textContent = message;
  status.className = `status ${type}`;
}

function hideStatus() {
  document.getElementById('status').style.display = 'none';
}

function showProgress() {
  document.getElementById('progress').style.display = 'block';
  animateProgress();
}

function hideProgress() {
  document.getElementById('progress').style.display = 'none';
  document.getElementById('progress-bar').style.width = '0%';
}

function animateProgress() {
  const bar = document.getElementById('progress-bar');
  let width = 0;
  const interval = setInterval(() => {
    if (width >= 90) {
      clearInterval(interval);
    } else {
      width += Math.random() * 10;
      bar.style.width = Math.min(width, 90) + '%';
    }
  }, 200);
}

function setButtonsEnabled(enabled) {
  document.getElementById('previewBtn').disabled = !enabled;
  document.getElementById('downloadBtn').disabled = !enabled;
}

function validateForm() {
  const fileInput = document.querySelector('input[type="file"]');
  if (!fileInput.files.length) {
    showStatus('Please select a file to upload', 'error');
    return false;
  }
  return true;
}

async function doPreview(){
  if (!validateForm()) return;
  
  setButtonsEnabled(false);
  showStatus('Analyzing unit data and generating scenarios...', 'loading');
  showProgress();
  
  try {
    const f = document.getElementById('f');
    const fd = new FormData(f);
    fd.append('fast_preview', document.getElementById('fast_preview').checked ? '1' : '0');
    
    const r = await fetch('/preview', {method:'POST', body:fd});
    const j = await r.json();
    
    if (j.error) { 
      showStatus(`Error: ${j.error}`, 'error');
      return; 
    }
    
    hideProgress();
    showStatus(`Successfully generated ${Object.keys(j.scenarios).length} scenario(s)`, 'success');
    
    const out = document.getElementById('out');
    out.innerHTML = '';
    
    const header = document.createElement('div');
    header.className = 'card';
    header.innerHTML = `
      <h2 class="subtitle">Generated Scenarios for ${j.filename}</h2>
      <p class="muted">Best scenario: <span class="success">${j.best}</span></p>
    `;
    out.appendChild(header);
    
    const grid = document.createElement('div'); 
    grid.className = 'grid'; 
    out.appendChild(grid);
    
    Object.keys(j.scenarios).forEach(k => {
      const card = document.createElement('div'); 
      card.className = 'card scenario-card'; 
      grid.appendChild(card);
      
      const m = j.metrics[k];
      const isBest = k === j.best;
      
      card.innerHTML = `
        <div class="scenario-header">
          <div class="title">Scenario ${k}</div>
          ${isBest ? '<div class="scenario-badge">BEST</div>' : ''}
        </div>
        <div class="muted">
          <div>Weighted Average AMI: <span class="success">${m.wavg.toFixed(2)}%</span></div>
          <div>40% AMI Share: <span class="success">${m.pct40.toFixed(2)}%</span></div>
          <div>Total Affordable SF: <span class="muted">${m.aff_sf_total.toFixed(0)} sq ft</span></div>
        </div>
      `;
      
      const b = j.scenarios[k];
      ["40","60","70","80","90","100"].forEach(t => {
        if (!b[t]?.length) return;
        
        const section = document.createElement('div'); 
        section.className = 'ami-section';
        section.innerHTML = `<div class="ami-header">${t}% AMI Units (${b[t].length})</div>`;
        
        b[t].forEach(u => {
          const p = document.createElement('span'); 
          p.className = 'pill';
          p.textContent = `${u.APT || 'N/A'} • Floor ${u.FLOOR || 'N/A'} • ${u.NET_SF.toFixed(0)} sf`;
          section.appendChild(p);
        });
        
        card.appendChild(section);
      });
    });
    
    setTimeout(hideStatus, 3000);
    
  } catch (error) {
    hideProgress();
    showStatus(`Network error: ${error.message}`, 'error');
  } finally {
    setButtonsEnabled(true);
  }
}

async function doDownload(){
  if (!validateForm()) return;
  
  setButtonsEnabled(false);
  showStatus('Generating detailed Excel report with MILP optimization...', 'loading');
  showProgress();
  
  try {
    const f = document.getElementById('f');
    const fd = new FormData(f);
    
    const r = await fetch('/allocate', {method:'POST', body:fd});
    
    if (!r.ok) { 
      const j = await r.json(); 
      showStatus(`Error: ${j.error || 'Download failed'}`, 'error');
      return; 
    }
    
    const blob = await r.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); 
    a.href = url; 
    a.download = 'AMI_Scenarios.xlsx'; 
    document.body.appendChild(a); 
    a.click(); 
    a.remove();
    URL.revokeObjectURL(url);
    
    hideProgress();
    showStatus('Excel report downloaded successfully!', 'success');
    setTimeout(hideStatus, 3000);
    
  } catch (error) {
    hideProgress();
    showStatus(`Download error: ${error.message}`, 'error');
  } finally {
    setButtonsEnabled(true);
  }
}

document.addEventListener('DOMContentLoaded', function() {
  hideStatus();
});
</script>
</body>
</html>'''
    
    with open(index_html_path, 'w') as f:
        f.write(index_html_content)

# Initialize templates
try:
    templates = Jinja2Templates(directory="templates")
except Exception as e:
    print(f"Template initialization error: {e}")
    templates = None

# Path for logging runs
DB_PATH = os.environ.get("AMI_DB", "/tmp/ami_log.sqlite")

# ---------- DB helpers ----------
def init_db() -> None:
    """Initialize the sqlite database for recording runs."""
    try:
        os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
        with sqlite3.connect(DB_PATH) as con:
            con.execute(
                """CREATE TABLE IF NOT EXISTS runs(
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    ts INTEGER, filename TEXT, scenario TEXT,
                    aff_sf_total REAL, sf_at_40 REAL, pct40 REAL, wavg REAL
                )"""
            )
    except Exception as e:
        print(f"Database initialization error: {e}")

def log_run(filename: str, metrics: Dict[str, Any]) -> None:
    """Insert one row per scenario into the runs table."""
    try:
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
    except Exception as e:
        print(f"Database logging error: {e}")

# ---------- File loader ----------
def load_any_table(file_bytes: bytes, filename: str, sheet: Optional[str] = None) -> pd.DataFrame:
    """Load a table from various formats."""
    name = filename.lower()
    bio = io.BytesIO(file_bytes)
    
    try:
        if name.endswith((".xlsx", ".xlsm")):
            xls = pd.ExcelFile(bio)
            sh = sheet or xls.sheet_names[0]
            return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sh)
        elif name.endswith(".xls"):
            xls = pd.ExcelFile(bio)
            sh = sheet or xls.sheet_names[0]
            return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sh, engine="xlrd")
        elif name.endswith(".xlsb"):
            xls = pd.ExcelFile(bio, engine="pyxlsb")
            sh = sheet or xls.sheet_names[0]
            return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sh, engine="pyxlsb")
        elif name.endswith(".csv"):
            return pd.read_csv(io.BytesIO(file_bytes))
        elif name.endswith(".docx"):
            try:
                from docx import Document
                doc = Document(bio)
                best: Optional[pd.DataFrame] = None
                for table in doc.tables:
                    rows = [[c.text.strip() for c in r.cells] for r in table.rows]
                    if not rows or len(rows) < 2:
                        continue
                    df = pd.DataFrame(rows[1:], columns=rows[0])
                    if any("sf" in str(c).lower().replace(" ", "") for c in df.columns):
                        best = df
                        break
                    best = best or df
                if best is None:
                    raise ValueError("No readable tables in DOCX.")
                for c in best.columns:
                    try:
                        best[c] = pd.to_numeric(best[c], errors="ignore")
                    except Exception:
                        pass
                return best
            except ImportError:
                raise ValueError("python-docx not available for DOCX files")
        else:
            raise ValueError("Unsupported file type. Upload .xlsx/.xlsm/.xls/.xlsb/.csv/.docx")
    except Exception as e:
        raise ValueError(f"Error reading file: {str(e)}")

# ---------- UI routes ----------
@app.get("/", response_class=HTMLResponse)
async def root(request: Request):
    """Render the single‑page HTML interface."""
    if templates:
        try:
            return templates.TemplateResponse("index.html", {"request": request})
        except Exception as e:
            print(f"Template error: {e}")
    
    # Fallback HTML response
    return HTMLResponse(content="""
    <!DOCTYPE html>
    <html>
    <head><title>NYC AMI Allocator</title></head>
    <body>
        <h1>NYC AMI Allocator</h1>
        <p>Application is running. Template system error - please check deployment.</p>
        <p>API endpoints available:</p>
        <ul>
            <li>POST /preview - Preview scenarios</li>
            <li>POST /allocate - Download Excel report</li>
            <li>GET /health - Health check</li>
        </ul>
    </body>
    </html>
    """)

@app.get("/health")
async def health() -> Dict[str, bool]:
    """Simple health check endpoint."""
    return {"ok": True}

# ---------- Preview endpoint ----------
@app.post("/preview")
async def preview(
    file: UploadFile = File(...),
    sheet: Optional[str] = Form(None),
    require_family_at_40: int = Form(0),
    spread_40_max_per_floor: Optional[int] = Form(None),
    exempt_top_k_floors: int = Form(0),
    top_k: int = Form(3),
    fast_preview: int = Form(1),
):
    """Produce a quick preview of scenarios."""
    try:
        content = await file.read()
        df = load_any_table(content, file.filename, sheet)
        use_milp = bool(1 - int(fast_preview))
        
        full, aff_br, metrics, mirror, best = allocate_with_scenarios(
            df,
            require_family_at_40=bool(require_family_at_40),
            spread_40_max_per_floor=spread_40_max_per_floor,
            exempt_top_k_floors=exempt_top_k_floors,
            return_top_k=top_k,
            use_milp=use_milp,
        )
        
        labels = [c for c in full.columns if c.startswith("Assigned_AMI_")]
        if not labels:
            return JSONResponse({"error": "No scenarios generated"}, status_code=400)
        
        aff_only = full.loc[full[labels[0]].notna()].copy()

        def buckets(label: str) -> Dict[str, list]:
            col = f"Assigned_AMI_{label}"
            out = {"40": [], "60": [], "70": [], "80": [], "90": [], "100": []}
            for _, r in aff_only.iterrows():
                v = float(r[col])
                key = {0.4: "40", 0.6: "60", 0.7: "70", 0.8: "80", 0.9: "90", 1.0: "100"}.get(round(v, 1), "60")
                out[key].append({
                    "APT": r.get("APT"),
                    "FLOOR": None if pd.isna(r.get("FLOOR")) else int(r.get("FLOOR")),
                    "BED": r.get("BED"),
                    "NET_SF": float(r["NET SF"]),
                })
            return out

        resp = {
            "filename": file.filename,
            "best": best,
            "metrics": metrics,
            "scenarios": {lab.replace("Assigned_AMI_", ""): buckets(lab.replace("Assigned_AMI_", "")) for lab in labels},
        }
        return resp
        
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=400)

# ---------- Full allocate endpoint ----------
@app.post("/allocate")
async def allocate(
    file: UploadFile = File(...),
    sheet: Optional[str] = Form(None),
    write_back: str = Form("new"),
    require_family_at_40: int = Form(0),
    spread_40_max_per_floor: Optional[int] = Form(None),
    exempt_top_k_floors: int = Form(0),
    top_k: int = Form(3),
) -> StreamingResponse:
    """Solve the allocation problem and return an Excel workbook."""
    try:
        content = await file.read()
        df = load_any_table(content, file.filename, sheet)
        
        full, aff_br, metrics, mirror, best = allocate_with_scenarios(
            df,
            require_family_at_40=bool(require_family_at_40),
            spread_40_max_per_floor=spread_40_max_per_floor,
            exempt_top_k_floors=exempt_top_k_floors,
            return_top_k=top_k,
            use_milp=True,
        )
        
        base = os.path.splitext(os.path.basename(file.filename))[0]
        log_run(base, metrics)

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
            full.to_excel(writer, index=False, sheet_name="Master")
            
            scenario_labels = [c.replace("Assigned_AMI_", "") for c in full.columns if c.startswith("Assigned_AMI_")]
            for lab in scenario_labels:
                aff_mask = full[f"Assigned_AMI_{lab}"].notna()
                if not aff_mask.any():
                    continue
                    
                base_cols = [c for c in ["FLOOR","APT","BED","NET SF","AMI_RAW"] if c in full.columns]
                br = full.loc[aff_mask, base_cols].copy()
                br[f"Assigned_AMI_{lab}"] = full.loc[aff_mask, f"Assigned_AMI_{lab}"]
                br.to_excel(writer, index=False, sheet_name=f"Breakdown_{lab}")
            
            summary_rows = []
            for k, v in metrics.items():
                row = {"Scenario": k}
                row.update(v)
                summary_rows.append(row)
            if summary_rows:
                pd.DataFrame(summary_rows).to_excel(writer, index=False, sheet_name="Summary")

        buf.seek(0)
        out_name = f"{base}.xlsx" if write_back == "same" else f"{base} - AMI Scenarios.xlsx"
        
        return StreamingResponse(
            io.BytesIO(buf.read()),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{out_name}"'},
        )
        
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=400)

@app.get("/export_master")
async def export_master() -> StreamingResponse:
    """Download the log of all runs."""
    try:
        init_db()
        with sqlite3.connect(DB_PATH) as con:
            df = pd.read_sql_query("SELECT * FROM runs ORDER BY ts DESC", con)
        
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="All Runs")
        
        buf.seek(0)
        return StreamingResponse(
            io.BytesIO(buf.read()),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": 'attachment; filename="AMI Master All Results.xlsx"'},
        )
        
    except Exception as e:
        return JSONResponse({"error": f"Cannot read DB: {e}"}, status_code=400)

# Initialize database on startup
@app.on_event("startup")
async def startup_event():
    init_db()

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))

