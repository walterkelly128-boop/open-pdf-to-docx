import shutil
import tempfile
import uuid
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from .converter.docx_builder import convert_pdf_to_docx

BASE = Path(__file__).resolve().parents[2]
STATIC = BASE / "frontend"
app = FastAPI(title="Open PDF to DOCX", version="0.1.0")


@app.get("/api/health")
def health():
    return {"status": "ok"}


@app.post("/api/convert")
async def convert(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Please upload a PDF file.")

    job = uuid.uuid4().hex
    work = Path(tempfile.mkdtemp(prefix=f"open-pdf-to-docx-{job}-"))
    source = work / "input.pdf"
    target = work / f"{Path(file.filename).stem}.docx"
    try:
        with source.open("wb") as output:
            shutil.copyfileobj(file.file, output)
        convert_pdf_to_docx(source, target)
        response = FileResponse(
            target,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=target.name,
        )
        return response
    except Exception as exc:
        shutil.rmtree(work, ignore_errors=True)
        raise HTTPException(500, f"Conversion failed: {exc}") from exc


if STATIC.exists():
    app.mount("/", StaticFiles(directory=STATIC, html=True), name="frontend")
