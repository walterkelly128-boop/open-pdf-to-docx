from dataclasses import dataclass
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
class PdfPage:
    number: int
    width: float
    height: float
    words: list[PdfWord]
    images: list[bytes]


def iter_pages(pdf_path: str | Path) -> Iterator[PdfPage]:
    doc = fitz.open(pdf_path)
    try:
        for page_number, page in enumerate(doc, start=1):
            raw = page.get_text("dict", flags=fitz.TEXTFLAGS_TEXT | fitz.TEXTFLAGS_BLOCKS)
            words: list[PdfWord] = []
            for block_no, block in enumerate(raw.get("blocks", [])):
                if block.get("type") != 0:
                    continue
                for line_no, line in enumerate(block.get("lines", [])):
                    for word_no, span in enumerate(line.get("spans", [])):
                        text = span.get("text", "")
                        if not text.strip():
                            continue
                        x0, y0, x1, y1 = span["bbox"]
                        words.append(PdfWord(
                            text=text, x0=x0, y0=y0, x1=x1, y1=y1,
                            size=float(span.get("size", 10)),
                            font=span.get("font", "Arial"),
                            flags=int(span.get("flags", 0)),
                            color=int(span.get("color", 0)),
                            block_no=block_no, line_no=line_no, word_no=word_no,
                        ))
            images = []
            for image in page.get_images(full=True):
                try:
                    images.append(doc.extract_image(image[0])["image"])
                except Exception:
                    pass
            yield PdfPage(page_number, page.rect.width, page.rect.height, words, images)
    finally:
        doc.close()
