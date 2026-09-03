from io import BytesIO
from pathlib import Path
from typing import Iterable

import fitz
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from .layout import LayoutBlock, LayoutLine, build_layout
from .pdf_parser import PdfImage, PdfPage, PdfWord, iter_pages

_FONT_MAP = {
    "ArialMT": "Arial", "Arial-BoldMT": "Arial", "TimesNewRomanPSMT": "Times New Roman",
    "TimesNewRomanPS-BoldMT": "Times New Roman", "Helvetica": "Arial", "Helvetica-Bold": "Arial",
    "Helvetica-Oblique": "Arial", "Courier": "Courier New", "Calibri": "Calibri",
    "SimSun": "SimSun", "宋体": "SimSun", "Microsoft YaHei": "Microsoft YaHei", "微软雅黑": "Microsoft YaHei",
}


def safe_font(name: str) -> str:
    base = name.split("+")[-1].strip()
    if base in _FONT_MAP:
        return _FONT_MAP[base]
    for suffix in ("-BoldItalic", "-Bold", "-Italic", "_Bold", "_Italic"):
        if base.endswith(suffix):
            base = base[: -len(suffix)]
            break
    return _FONT_MAP.get(base, base or "Arial")


def _set_run_font(run, word: PdfWord) -> None:
    name = safe_font(word.font)
    run.font.name = name
    run.font.size = Pt(max(1.0, word.size))
    run.bold = word.bold
    run.italic = word.italic
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
        rfonts.set(qn(f"w:{attr}"), name)
    value = max(0, min(0xFFFFFF, int(word.color)))
    run.font.color.rgb = RGBColor((value >> 16) & 255, (value >> 8) & 255, value & 255)


def _add_spacing(text: str, previous: PdfWord | None, current: PdfWord) -> str:
    if previous is None:
        return text
    gap = current.x0 - previous.x1
    threshold = max(1.5, min(previous.size, current.size) * 0.18)
    if gap <= threshold:
        return text
    if previous.text[-1:].isascii() and current.text[:1].isascii():
        return " " + text
    return text


def _write_line(paragraph, line: LayoutLine) -> None:
    previous = None
    for word in line.words:
        run = paragraph.add_run(_add_spacing(word.text, previous, word))
        _set_run_font(run, word)
        previous = word


def _configure_section(section, page: PdfPage) -> None:
    section.page_width = Inches(page.width / 72)
    section.page_height = Inches(page.height / 72)
    section.top_margin = section.bottom_margin = Inches(0.08)
    section.left_margin = section.right_margin = Inches(0.08)
    section.header_distance = Inches(0)
    section.footer_distance = Inches(0)


def _add_block(doc: Document, block: LayoutBlock, page: PdfPage) -> None:
    paragraph = doc.add_paragraph()
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(0)
    fmt.line_spacing = 1.0
    fmt.left_indent = Inches(max(0, block.x0) / 72)
    fmt.right_indent = Inches(max(0, page.width - block.x1) / 72)
    for index, line in enumerate(block.lines):
        if index:
            paragraph.add_run().add_break()
        _write_line(paragraph, line)


def _add_image(doc: Document, image: PdfImage, page: PdfPage) -> None:
    paragraph = doc.add_paragraph()
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(0)
    fmt.left_indent = Inches(max(0, image.x0) / 72)
    width_in = min(max(0.05, (image.x1 - image.x0) / 72), max(0.5, (page.width - image.x0 - 4) / 72))
    height_in = max(0.05, (image.y1 - image.y0) / 72)
    try:
        paragraph.add_run().add_picture(BytesIO(image.data), width=Inches(width_in), height=Inches(height_in))
    except Exception:
        paragraph._element.getparent().remove(paragraph._element)


def _content_items(page: PdfPage) -> Iterable[tuple[float, str, object]]:
    blocks = build_layout(page)
    items: list[tuple[float, str, object]] = [(b.y0, "text", b) for b in blocks]
    items.extend((i.y0, "image", i) for i in page.images)
    return sorted(items, key=lambda item: (item[0], 0 if item[1] == "image" else 1))


def _add_full_page_render(doc: Document, page, dpi: int = 180) -> None:
    """Place an exact rasterization of a PDF page at the source page size."""
    scale = dpi / 72.0
    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    image = BytesIO(pix.tobytes("png"))
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(0)
    fmt.line_spacing = 1.0
    paragraph.add_run().add_picture(
        image,
        width=Inches(page.rect.width / 72),
        height=Inches(page.rect.height / 72),
    )


def _convert_fidelity(pdf_path: str | Path, output_path: str | Path, dpi: int = 180) -> None:
    """Visual-first conversion for complex PDFs.

    Word's normal paragraph model cannot faithfully reproduce arbitrary PDF
    coordinates, columns, overlays, logos and form-like layouts. Fidelity mode
    therefore rasterizes each source page at high resolution and puts that
    page-sized image into Word. This prevents the cascading reflow visible in
    conventional paragraph reconstruction while keeping the source page size.
    """
    pdf = fitz.open(pdf_path)
    try:
        doc = Document()
        for index, page in enumerate(pdf):
            section = doc.sections[0] if index == 0 else doc.add_section(WD_SECTION.NEW_PAGE)
            section.page_width = Inches(page.rect.width / 72)
            section.page_height = Inches(page.rect.height / 72)
            section.top_margin = section.bottom_margin = Inches(0)
            section.left_margin = section.right_margin = Inches(0)
            section.header_distance = section.footer_distance = Inches(0)
            _add_full_page_render(doc, page, dpi=dpi)
        if len(doc.paragraphs) == 1 and not doc.paragraphs[0].text:
            p = doc.paragraphs[0]._element
            p.getparent().remove(p)
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        doc.save(output_path)
    finally:
        pdf.close()


def _convert_editable(pdf_path: str | Path, output_path: str | Path) -> None:
    doc = Document()
    first_page = True
    for page in iter_pages(pdf_path):
        if first_page:
            section = doc.sections[0]
            first_page = False
        else:
            section = doc.add_section(WD_SECTION.NEW_PAGE)
        _configure_section(section, page)
        for _, kind, item in _content_items(page):
            if kind == "text":
                _add_block(doc, item, page)
            else:
                _add_image(doc, item, page)
    if len(doc.paragraphs) == 1 and not doc.paragraphs[0].text:
        p = doc.paragraphs[0]._element
        p.getparent().remove(p)
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)


def convert_pdf_to_docx(
    pdf_path: str | Path,
    output_path: str | Path,
    mode: str = "fidelity",
    fidelity_dpi: int = 180,
) -> None:
    mode = (mode or "fidelity").lower().strip()
    if mode not in {"fidelity", "editable"}:
        raise ValueError("mode must be 'fidelity' or 'editable'")
    if mode == "editable":
        _convert_editable(pdf_path, output_path)
    else:
        dpi = max(120, min(300, int(fidelity_dpi)))
        _convert_fidelity(pdf_path, output_path, dpi=dpi)
