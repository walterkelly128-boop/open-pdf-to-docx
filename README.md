# Open PDF to DOCX

A self-hosted open-source PDF to Word (DOCX) converter focused on layout preservation.

## Goals

- No `pdf2docx` dependency.
- Extract text, fonts, images and geometry from PDF pages.
- Reconstruct paragraphs, columns, tables and images in DOCX.
- Keep the conversion engine modular so OCR can be added for scanned PDFs.
- Simple web UI and Docker deployment.

## Architecture

```text
PDF
  -> PDF parser / geometry extraction
  -> layout analysis
  -> intermediate document model
  -> OOXML DOCX builder
  -> .docx
```

The project intentionally separates PDF extraction from layout reconstruction. This makes it possible to improve difficult documents such as resumes, multi-column pages and tables without replacing the whole converter.

## Current status

Early MVP. The first implementation targets text-based PDFs and basic layout reconstruction. Scanned-PDF OCR is an extension point rather than a hard dependency.

## Run locally

Python 3.11+ is recommended.

```bash
python -m venv .venv
# Linux/macOS
source .venv/bin/activate
# Windows PowerShell
# .venv\\Scripts\\Activate.ps1
pip install -r backend/requirements.txt
uvicorn backend.app.main:app --reload
```

Open `http://127.0.0.1:8000`.

## Docker

```bash
docker compose up --build
```

## License

MIT License.
