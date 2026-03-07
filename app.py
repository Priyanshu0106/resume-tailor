import os
import uuid
import tempfile
from pathlib import Path

from fastapi import FastAPI, File, Form, UploadFile, HTTPException, Request
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles

from tailor import tailor_resume

app = FastAPI(title="Resume Tailor")

BASE_DIR = Path(__file__).parent
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))

UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/tailor")
async def tailor(
    resume: UploadFile = File(...),
    jd: str = Form(...),
):
    if not resume.filename.endswith(".docx"):
        raise HTTPException(status_code=400, detail="Only .docx files are supported.")
    if not jd.strip():
        raise HTTPException(status_code=400, detail="Job description cannot be empty.")

    session = uuid.uuid4().hex
    input_path = UPLOAD_DIR / f"{session}_input.docx"
    output_path = UPLOAD_DIR / f"{session}_tailored.docx"

    try:
        content = await resume.read()
        input_path.write_bytes(content)

        tailor_resume(str(input_path), str(output_path), jd)

        original_name = Path(resume.filename).stem
        download_name = f"{original_name}_tailored.docx"

        return FileResponse(
            path=str(output_path),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=download_name,
            background=None,
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        if input_path.exists():
            input_path.unlink()


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)
