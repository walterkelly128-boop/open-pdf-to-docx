from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterator

import fitz


@dataclass
class PdfWord:
    text: str
    x0: float
    y0: float
    x1: float
    y1: float
    size: float
    font: str
    flags: int
    color: int
    block_no: int
    line_no: int
    word_no: int

    @property
    def bold(self) -> bool:
        return bool(self.flags & 16)

    @property
    def italic(self) -> bool:
        return bool(self.flags & 2)


@dataclass
class PdfImage:
    data: bytes
    x0: float
    y0: float
    x1: float
    y1: float
    ext: str


@dataclass
class PdfPage:
    number: int
    width: float
    height: float
    words: list[PdfWord] = field(default_factory=list)
    images: list[PdfImage] = field(default_factory=list)


def iter_pages(pdf_path: str | Path) -> Iterator[PdfPage]:
    """Read PDF content while keeping text runs tied to their original geometry.

    We intentionally build editable text from the spans returned by PyMuPDF's
    structured text dictionary instead of trying to join it back to
    ``get_text('words')`` output.  The two APIs do not guarantee identical block
    and line numbering (especially when image blocks are present), which could
    silently produce empty text boxes in the DOCX even though the PDF contains
    valid text.
    """
    doc = fitz.open(pdf_path)
    try:
        for page_number, page in enumerate(doc, start=1):
            raw = page.get_text("dict", flags=fitz.TEXTFLAGS_TEXT)
            words: list[PdfWord] = []

            for block_no, block in enumerate(raw.get("blocks", [])):
                if block.get("type") != 0:
                    continue
                for line_no, line in enumerate(block.get("lines", [])):
                    spans = line.get("spans", [])
                    for word_no, span in enumerate(spans):
                        text = str(span.get("text", ""))
                        if not text.strip():
                            continue
                        x0, y0, x1, y1 = span.get("bbox", (0, 0, 0, 0))
                        words.append(PdfWord(
                            text=text,
                            x0=float(x0),
                            y0=float(y0),
                            x1=float(x1),
                            y1=float(y1),
                            size=float(span.get("size", max(1.0, float(y1) - float(y0)))),
                            font=str(span.get("font", "Arial")),
                            flags=int(span.get("flags", 0)),
                            color=int(span.get("color", 0)),
                            block_no=block_no,
                            line_no=line_no,
                            word_no=word_no,
                        ))

            images: list[PdfImage] = []
            seen: set[tuple[int, int, int, int, int]] = set()
            for image in page.get_images(full=True):
                xref = image[0]
                try:
                    rects = page.get_image_rects(xref)
                    extracted = doc.extract_image(xref)
                    data = extracted.get("image")
                    ext = extracted.get("ext", "png")
                    if not data:
                        continue
                    for rect in rects:
                        key = (xref, round(rect.x0), round(rect.y0), round(rect.x1), round(rect.y1))
                        if key in seen:
                            continue
                        seen.add(key)
                        images.append(PdfImage(data, rect.x0, rect.y0, rect.x1, rect.y1, ext))
                except Exception:
                    continue

            yield PdfPage(page_number, page.rect.width, page.rect.height, words, images)
    finally:
        doc.close()
