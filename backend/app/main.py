import os
import re
import shutil
import tempfile
import uuid
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.background import BackgroundTask
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from .converter.docx_builder import convert_pdf_to_docx

BASE = Path(__file__).resolve().parents[2]
STATIC = BASE / "frontend"
VERSION = "1.0.0"
MAX_UPLOAD_MB = int(os.getenv("MAX_UPLOAD_MB", "100"))
MAX_UPLOAD_BYTES = MAX_UPLOAD_MB * 1024 * 1024

app = FastAPI(title="Open PDF to DOCX", version=VERSION)


@app.get("/api/health")
def health():
    return {"status": "ok", "version": VERSION}


def _safe_filename(name: str) -> str:
    stem = Path(name).stem
    stem = re.sub(r"[^\w\-. ]+", "_", stem, flags=re.UNICODE).strip(" .")
    return (stem or "converted")[:120]


@app.post("/api/convert")
async def convert(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Please upload a PDF file.")

    job = uuid.uuid4().hex
    work = Path(tempfile.mkdtemp(prefix=f"open-pdf-to-docx-{job}-"))
    source = work / "input.pdf"
    target = work / f"{_safe_filename(file.filename)}.docx"

    try:
        size = 0
        with source.open("wb") as output:
            while True:
                chunk = await file.read(1024 * 1024)
                if not chunk:
                    break
                size += len(chunk)
                if size > MAX_UPLOAD_BYTES:
                    raise HTTPException(413, f"PDF exceeds the {MAX_UPLOAD_MB} MB upload limit.")
                output.write(chunk)

        await file.close()
        with source.open("rb") as check:
            if check.read(5) != b"%PDF-":
                raise HTTPException(400, "The uploaded file is not a valid PDF.")

        convert_pdf_to_docx(source, target)
        cleanup = BackgroundTask(shutil.rmtree, work, ignore_errors=True)
        return FileResponse(
            target,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=target.name,
            background=cleanup,
        )
    except HTTPException:
        shutil.rmtree(work, ignore_errors=True)
        raise
    except Exception as exc:
        shutil.rmtree(work, ignore_errors=True)
        raise HTTPException(500, f"Conversion failed: {exc}") from exc


if STATIC.exists():
    app.mount("/", StaticFiles(directory=STATIC, html=True), name="frontend")
