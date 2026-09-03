# Open PDF to DOCX

A self-hosted PDF to Word (DOCX) converter focused on producing editable documents while preserving the original page geometry as closely as practical.

## Version 1.0.0

This release is the first complete conversion engine rather than the earlier proof-of-concept.

### Included

- No `pdf2docx` dependency.
- Real PDF word extraction instead of treating whole text spans as words.
- Font family, size, bold, italic and text color preservation where available.
- Paragraph/block inference from PDF geometry.
- Basic two-column reading-order detection.
- Page dimensions and margins derived from the source PDF.
- Embedded PDF images copied into the DOCX with their source dimensions.
- Chinese/CJK font metadata written into DOCX runs when the source font is available.
- Browser drag-and-drop upload UI.
- 100 MB default upload limit, configurable with `MAX_UPLOAD_MB`.
- Automatic temporary-file cleanup after a successful download.
- Docker deployment.

### Best results

The engine is optimized for text-based PDFs such as articles, reports, resumes, manuals and business documents. PDFs that are scans, heavily composed of positioned text boxes, contain complex vector tables, or depend on unusual proprietary fonts may still require OCR or specialized layout handling.

## Architecture

```text
PDF
  -> word/font/image extraction
  -> geometry-based layout analysis
  -> editable document reconstruction
  -> OOXML DOCX builder
  -> .docx
```

The extraction and reconstruction layers are deliberately separated so future releases can add OCR, table detection, floating text boxes and more advanced reading-order analysis without replacing the web application.

## Run locally with Docker Desktop (recommended on Windows)

Make sure Docker Desktop is running, then open PowerShell or Command Prompt in the project directory and run:

```bash
docker compose up --build
```

Then open:

`http://localhost:8000`

Upload a PDF in the browser and download the generated DOCX. The default upload limit is 100 MB.

To stop the service:

```bash
docker compose down
```

For Windows, you can also double-click `start-docker.bat` in the project directory.

### Check the container health

```bash
docker compose ps
```

The application exposes `GET /api/health` for the container health check.

## Run locally with Python

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

## API

`GET /api/health` returns the service version.

`POST /api/convert` accepts a PDF as multipart field `file` and returns a DOCX file.

## Dependency licensing

The application source is provided under the MIT License. Third-party dependencies retain their own licenses. In particular, review the current PyMuPDF licensing terms before redistributing this project as a hosted or packaged commercial product.

## License

MIT License for the project source code. See `LICENSE`.
