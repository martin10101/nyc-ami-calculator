import io
import os
import zipfile
from fastapi import FastAPI, UploadFile, Form, Request, File  # Added File
from fastapi.responses import StreamingResponse, HTMLResponse
from fastapi.templating import Jinja2Templates
import pandas as pd
from ami_core import generate_scenarios, build_outputs  # Import from your core logic

app = FastAPI()
templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=HTMLResponse)
async def root(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/assign_ami")
async def assign_ami(file: UploadFile = File(...), required_sf: float = Form(...), prefs: str = Form("")):
    try:
        contents = await file.read()
        # Handle Excel/PDF - add pdfplumber if PDF
        df = pd.read_excel(io.BytesIO(contents)) if file.filename.endswith(('.xlsx', '.xls', '.xlsb')) else pd.DataFrame()
        prefs_dict = {"required_40_pct": 0.20 if "40" in prefs else 0, "write_back": "write back" in prefs.lower()}
        scen, aff = generate_scenarios(df, required_sf, prefs_dict)
        outputs = build_outputs(df, scen, aff, prefs_dict)  # List of (name, bytes)
        
        # Create ZIP in memory
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zipf:
            for name, buf in outputs:
                zipf.writestr(name, buf)
        zip_buf.seek(0)
        
        return StreamingResponse(zip_buf, media_type="application/zip", headers={"Content-Disposition": "attachment; filename=optimized_ami.zip"})
    except Exception as e:
        return {"error": str(e)}
