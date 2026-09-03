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


def _span_for_word(spans: list[dict], bbox: tuple[float, float, float, float]) -> dict:
    cx = (bbox[0] + bbox[2]) / 2
    cy = (bbox[1] + bbox[3]) / 2
    best = {}
    best_score = -1.0
    for span in spans:
        sx0, sy0, sx1, sy1 = span.get("bbox", (0, 0, 0, 0))
        overlap_x = max(0.0, min(bbox[2], sx1) - max(bbox[0], sx0))
        overlap_y = max(0.0, min(bbox[3], sy1) - max(bbox[1], sy0))
        area = max(1.0, (bbox[2] - bbox[0]) * (bbox[3] - bbox[1]))
        score = (overlap_x * overlap_y) / area
        if sx0 <= cx <= sx1 and sy0 <= cy <= sy1:
            score += 2.0
        if score > best_score:
            best_score = score
            best = span
    return best


def iter_pages(pdf_path: str | Path) -> Iterator[PdfPage]:
    doc = fitz.open(pdf_path)
    try:
        for page_number, page in enumerate(doc, start=1):
            raw = page.get_text("dict", flags=fitz.TEXTFLAGS_TEXT)
            words: list[PdfWord] = []
            page_words = page.get_text("words")

            for block_no, block in enumerate(raw.get("blocks", [])):
                if block.get("type") != 0:
                    continue
                for line_no, line in enumerate(block.get("lines", [])):
                    spans = line.get("spans", [])
                    line_words = [
                        item for item in page_words
                        if int(item[5]) == block_no and int(item[6]) == line_no
                    ]
                    line_words.sort(key=lambda item: (float(item[0]), float(item[1])))
                    for word_no, item in enumerate(line_words):
                        x0, y0, x1, y1, text = item[:5]
                        if not str(text).strip():
                            continue
                        bbox = (float(x0), float(y0), float(x1), float(y1))
                        span = _span_for_word(spans, bbox)
                        words.append(PdfWord(
                            text=str(text), x0=float(x0), y0=float(y0),
                            x1=float(x1), y1=float(y1),
                            size=float(span.get("size", max(1.0, y1 - y0))),
                            font=str(span.get("font", "Arial")),
                            flags=int(span.get("flags", 0)),
                            color=int(span.get("color", 0)),
                            block_no=block_no, line_no=line_no, word_no=word_no,
                        ))

            images: list[PdfImage] = []
            seen: set[tuple[int, int, int, int]] = set()
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
                        key = (xref, round(rect.x0), round(rect.y0), round(rect.x1))
                        if key in seen:
                            continue
                        seen.add(key)
                        images.append(PdfImage(data, rect.x0, rect.y0, rect.x1, rect.y1, ext))
                except Exception:
                    continue

            yield PdfPage(page_number, page.rect.width, page.rect.height, words, images)
    finally:
        doc.close()
