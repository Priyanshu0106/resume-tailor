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
    if not resume.filename.lower().endswith((".docx", ".pdf")):
        raise HTTPException(status_code=400, detail="Only .docx and .pdf files are supported.")
    if not jd.strip():
        raise HTTPException(status_code=400, detail="Job description cannot be empty.")

    ext = Path(resume.filename).suffix.lower()
    session = uuid.uuid4().hex
    input_path = UPLOAD_DIR / f"{session}_input{ext}"
    # Output keeps same format as input (PDF→PDF, DOCX→DOCX)
    output_path = UPLOAD_DIR / f"{session}_tailored{ext}"

    try:
        content = await resume.read()
        input_path.write_bytes(content)

        tailor_resume(str(input_path), str(output_path), jd)

        original_name = Path(resume.filename).stem
        download_name = f"{original_name}_tailored{ext}"

        if ext == ".pdf":
            media_type = "application/pdf"
        else:
            media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

        return FileResponse(
            path=str(output_path),
            media_type=media_type,
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
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("app:app", host="0.0.0.0", port=port)
