import io
import os
from fastapi import FastAPI, UploadFile, Form, Request
from fastapi.responses import StreamingResponse, HTMLResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
import pandas as pd
from ami_core import generate_scenarios, build_outputs  # Import rewritten

app = FastAPI()
templates = Jinja2Templates(directory="templates")
app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/", response_class=HTMLResponse)
async def root(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/assign_ami")
async def assign_ami(file: UploadFile = File(...), required_sf: float = Form(...), prefs: str = Form("")):
    try:
        contents = await file.read()
        df = pd.read_excel(io.BytesIO(contents)) if file.filename.endswith(('.xlsx', '.xls', '.xlsb')) else pd.DataFrame()  # Add PDF via pdfplumber if needed
        prefs_dict = {"required_40_pct": 0.20 if "40" in prefs else 0, "write_back": "write back" in prefs.lower()}  # Parse prefs text
        scen, aff = generate_scenarios(df, required_sf, prefs_dict)
        outputs = build_outputs(df, scen, aff, prefs_dict)
        # Return ZIP or multiple downloads; for simplicity, stream first 2
        for name, buf in outputs:
            yield StreamingResponse(io.BytesIO(buf), media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f"attachment; filename={name}"})
    except Exception as e:
        return {"error": str(e)}

# Add static for CSS/JS in templates
