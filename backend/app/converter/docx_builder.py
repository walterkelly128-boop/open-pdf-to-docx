from pathlib import Path

from docx import Document
from docx.enum.text import WD_BREAK
from docx.shared import Inches, Pt

from .layout import build_layout
from .pdf_parser import iter_pages


_FONT_MAP = {
    "ArialMT": "Arial",
    "Arial-BoldMT": "Arial",
    "TimesNewRomanPSMT": "Times New Roman",
    "TimesNewRomanPS-BoldMT": "Times New Roman",
    "Calibri": "Calibri",
}


def safe_font(name: str) -> str:
    return _FONT_MAP.get(name, name.split("+")[-1] or "Arial")


def convert_pdf_to_docx(pdf_path: str | Path, output_path: str | Path) -> None:
    doc = Document()
    first_page = True

    for page in iter_pages(pdf_path):
        if not first_page:
            doc.add_page_break()
        first_page = False

        section = doc.sections[-1]
        section.page_width = Inches(page.width / 72)
        section.page_height = Inches(page.height / 72)
        section.top_margin = Inches(0.35)
        section.bottom_margin = Inches(0.35)
        section.left_margin = Inches(0.45)
        section.right_margin = Inches(0.45)

        blocks = build_layout(page)
        for block in blocks:
            paragraph = doc.add_paragraph()
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing = 1.0

            for index, line in enumerate(block.lines):
                if index:
                    paragraph.add_run().add_break()
                for word_index, word in enumerate(line.words):
                    if word_index:
                        paragraph.add_run(" ")
                    run = paragraph.add_run(word.text)
                    run.font.name = safe_font(word.font)
                    run.font.size = Pt(max(1, word.size))
                    run.bold = word.bold
                    run.italic = word.italic

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)
