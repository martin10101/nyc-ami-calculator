import io
import os
import zipfile
from fastapi import FastAPI, UploadFile, Request, File
from fastapi.responses import StreamingResponse, HTMLResponse
from fastapi.templating import Jinja2Templates
import pandas as pd
from ami_core import generate_scenarios, build_outputs

app = FastAPI()
templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=HTMLResponse)
async def root(request: Request):
    try:
        return templates.TemplateResponse("index.html", {"request": request})
    except Exception as e:
        return HTMLResponse(content="<h1>Error: Template not found. Check deployment.</h1><p>" + str(e) + "</p>")

@app.post("/assign_ami")
async def assign_ami(file: UploadFile = File(...)):
    try:
        contents = await file.read()
        df = pd.read_excel(io.BytesIO(contents)) if file.filename.endswith(('.xlsx', '.xls', '.xlsb')) else pd.DataFrame()
        # Auto-detect required_sf from selected units
        required_sf = df[df["AMI"].notna()]["NET SF"].sum() if "NET SF" in df and "AMI" in df else 0
        prefs_dict = {}  # Bake rules in ami_core - no user prefs
        scen, aff = generate_scenarios(df, required_sf, prefs_dict)
        outputs = build_outputs(df, scen, aff, prefs_dict)
        
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zipf:
            for name, buf in outputs:
                zipf.writestr(name, buf)
        zip_buf.seek(0)
        
        return StreamingResponse(zip_buf, media_type="application/zip", headers={"Content-Disposition": "attachment; filename=optimized_ami.zip"})
    except Exception as e:
        return {"error": str(e)}
