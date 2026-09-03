# Open PDF to DOCX

A self-hosted PDF to Word (DOCX) converter focused on preserving the original PDF page appearance while also providing a separate editable reconstruction mode.

## Version 1.1.0

### Conversion modes

- **High fidelity (recommended):** renders each PDF page at high resolution and places it at the exact source page size in DOCX. This is the best choice for CVs, forms, brochures, certificates, multi-column layouts, logos and PDFs with heavily positioned elements.
- **Editable:** extracts words, fonts, images and basic geometry and rebuilds the content as normal editable Word paragraphs. This is useful when text editing is more important than pixel-level layout.

The high-fidelity mode deliberately does **not** use `pdf2docx`. It uses PyMuPDF to render the source page, avoiding the paragraph reflow problem where a complex PDF becomes a long, incorrectly positioned Word document.

### Included

- No `pdf2docx` dependency.
- PyMuPDF-based PDF parsing and rendering.
- High-fidelity page reproduction with configurable DPI (`FIDELITY_DPI`, default 180, allowed range 120–300).
- Editable text reconstruction mode.
- Font family, size, bold, italic and text color preservation where available in editable mode.
- Basic paragraph/block inference and two-column reading-order detection in editable mode.
- Source PDF page dimensions preserved.
- Embedded PDF images copied in editable mode.
- Chinese/CJK font metadata written into DOCX runs when the source font is available.
- Browser drag-and-drop upload UI with conversion-mode selection.
- 100 MB default upload limit, configurable with `MAX_UPLOAD_MB`.
- Automatic temporary-file cleanup after successful download.
- Docker deployment.

## Why high fidelity is the default

PDF is a fixed-position document format. Word is a reflowable document format. A PDF can place every word, image and line at an arbitrary coordinate, while a normal Word paragraph is expected to flow around other content. Trying to force a complex PDF into ordinary paragraphs often produces exactly the kind of result where headings, contact fields and columns drift vertically or horizontally.

High-fidelity mode solves that specific problem by treating each source PDF page as the visual canvas it actually is. The resulting DOCX preserves the original page composition instead of allowing Word to reflow it.

This is the most reliable mode when the requirement is **“make the Word file look as close to the PDF as possible.”**

## Run locally with Docker Desktop (recommended on Windows)

Make sure Docker Desktop is running, then open PowerShell or Command Prompt in the project directory:

```bash
docker compose up --build
```

Then open:

`http://localhost:8000`

Choose **High fidelity** for the closest visual match, or **Editable** when you need reconstructed text and images.

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

`POST /api/convert` accepts a PDF as multipart field `file` and an optional multipart field `mode`:

- `fidelity` (default)
- `editable`

Example:

```bash
curl -X POST http://127.0.0.1:8000/api/convert \
  -F "file=@sample.pdf" \
  -F "mode=fidelity" \
  -o sample.docx
```

## Limitations

High-fidelity mode is visual-first: the PDF page is represented by a high-resolution image inside the DOCX, so it is not equivalent to converting every PDF glyph into independently editable Word text. Editable mode provides actual Word text, but complex positioned layouts can still differ from the source because of the fundamental PDF-to-Word layout-model mismatch.

Scanned PDFs without a text layer still need OCR for a fully editable reconstruction. Complex vector drawings, unusual fonts, interactive PDF features and advanced forms may also require specialized handling.

## Dependency licensing

The application source is provided under the MIT License. Third-party dependencies retain their own licenses. In particular, review the current PyMuPDF licensing terms before redistributing this project as a hosted or packaged commercial product.

## License

MIT License for the project source code. See `LICENSE`.
