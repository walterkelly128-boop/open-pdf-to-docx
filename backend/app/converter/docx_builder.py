from io import BytesIO
from pathlib import Path
from typing import Iterable
from xml.sax.saxutils import escape

from docx import Document
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from .layout import LayoutBlock, build_layout
from .pdf_parser import PdfImage, PdfPage, PdfWord, iter_pages

_FONT_MAP = {
    "ArialMT": "Arial", "Arial-BoldMT": "Arial", "TimesNewRomanPSMT": "Times New Roman",
    "TimesNewRomanPS-BoldMT": "Times New Roman", "Helvetica": "Arial", "Helvetica-Bold": "Arial",
    "Helvetica-Oblique": "Arial", "Courier": "Courier New", "Calibri": "Calibri",
    "SimSun": "SimSun", "宋体": "SimSun", "Microsoft YaHei": "Microsoft YaHei", "微软雅黑": "Microsoft YaHei",
}
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
V_NS = "urn:schemas-microsoft-com:vml"
O_NS = "urn:schemas-microsoft-com:office:office"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def safe_font(name: str) -> str:
    base = name.split("+")[-1].strip()
    if base in _FONT_MAP:
        return _FONT_MAP[base]
    for suffix in ("-BoldItalic", "-Bold", "-Italic", "_Bold", "_Italic"):
        if base.endswith(suffix):
            base = base[: -len(suffix)]
            break
    return _FONT_MAP.get(base, base or "Arial")


def _add_text_box(doc: Document, block: LayoutBlock, shape_id: int) -> None:
    width = max(4.0, block.x1 - block.x0 + 2.0)
    height = max(4.0, block.y1 - block.y0 + 3.0)
    style = (
        f"position:absolute;margin-left:{block.x0:.2f}pt;margin-top:{block.y0:.2f}pt;"
        f"width:{width:.2f}pt;height:{height:.2f}pt;"
        "mso-position-horizontal-relative:page;"
        "mso-position-vertical-relative:page;"
        f"z-index:{100 + shape_id};mso-wrap-style:none;"
    )
    xml = (
        f'<w:pict xmlns:w="{W_NS}" xmlns:v="{V_NS}" xmlns:o="{O_NS}" xmlns:r="{R_NS}">'
        f'<v:shape id="TextBox{shape_id}" type="#_x0000_t202" style="{escape(style)}" stroked="f" filled="f">'
        '<v:textbox inset="0,0,0,0"><w:txbxContent/></v:textbox></v:shape></w:pict>'
    )
    pict = parse_xml(xml)
    # Do not use pict.xpath('.//w:...'): namespace declarations on the parsed
    # fragment are not reliably registered with lxml's XPath evaluator.
    txbx = next(pict.iter(f"{{{W_NS}}}txbxContent"), None)
    if txbx is None:
        raise RuntimeError("Unable to create Word text box content node")

    for line in block.lines:
        p = OxmlElement("w:p")
        ppr = OxmlElement("w:pPr")
        spacing = OxmlElement("w:spacing")
        line_height = max(line.y1 - line.y0, max((w.size for w in line.words), default=8) * 1.05)
        spacing.set(qn("w:line"), str(max(1, round(line_height * 20))))
        spacing.set(qn("w:lineRule"), "exact")
        ppr.append(spacing)
        p.append(ppr)
        previous = None
        for word in line.words:
            r = OxmlElement("w:r")
            rpr = OxmlElement("w:rPr")
            rfonts = OxmlElement("w:rFonts")
            name = safe_font(word.font)
            for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
                rfonts.set(qn(f"w:{attr}"), name)
            rpr.append(rfonts)
            sz = OxmlElement("w:sz")
            sz.set(qn("w:val"), str(max(2, round(word.size * 2))))
            rpr.append(sz)
            if word.bold:
                rpr.append(OxmlElement("w:b"))
            if word.italic:
                rpr.append(OxmlElement("w:i"))
            value = max(0, min(0xFFFFFF, int(word.color)))
            color = OxmlElement("w:color")
            color.set(qn("w:val"), f"{value:06X}")
            rpr.append(color)
            r.append(rpr)
            text = word.text
            if previous is not None:
                gap = word.x0 - previous.x1
                threshold = max(1.5, min(previous.size, word.size) * 0.18)
                if gap > threshold and previous.text[-1:].isascii() and word.text[:1].isascii():
                    text = " " + text
            t = OxmlElement("w:t")
            if text[:1].isspace() or text[-1:].isspace():
                t.set(qn("xml:space"), "preserve")
            t.text = text
            r.append(t)
            p.append(r)
            previous = word
        txbx.append(p)
    doc.paragraphs[-1]._p.append(pict)


def _add_positioned_image(doc: Document, image: PdfImage, shape_id: int) -> None:
    try:
        r_id, _ = doc.part.get_or_add_image(BytesIO(image.data))
    except Exception:
        return
    width = max(2.0, image.x1 - image.x0)
    height = max(2.0, image.y1 - image.y0)
    style = (
        f"position:absolute;margin-left:{image.x0:.2f}pt;margin-top:{image.y0:.2f}pt;"
        f"width:{width:.2f}pt;height:{height:.2f}pt;"
        "mso-position-horizontal-relative:page;mso-position-vertical-relative:page;"
        f"z-index:{shape_id};mso-wrap-style:none;"
    )
    xml = (
        f'<w:pict xmlns:w="{W_NS}" xmlns:v="{V_NS}" xmlns:o="{O_NS}" xmlns:r="{R_NS}">'
        f'<v:shape id="Image{shape_id}" type="#_x0000_t75" style="{escape(style)}" stroked="f" filled="f">'
        f'<v:imagedata r:id="{r_id}" o:title="PDF image" /></v:shape></w:pict>'
    )
    doc.paragraphs[-1]._p.append(parse_xml(xml))


def _configure_section(section, page: PdfPage) -> None:
    section.page_width = Inches(page.width / 72)
    section.page_height = Inches(page.height / 72)
    section.top_margin = section.bottom_margin = Inches(0)
    section.left_margin = section.right_margin = Inches(0)
    section.header_distance = section.footer_distance = Inches(0)


def _content_items(page: PdfPage) -> Iterable[tuple[float, int, object]]:
    blocks = build_layout(page)
    items: list[tuple[float, int, object]] = [(b.y0, 1, b) for b in blocks]
    items.extend((i.y0, 0, i) for i in page.images)
    return sorted(items, key=lambda item: (item[0], item[1]))


def _convert_positioned_editable(pdf_path: str | Path, output_path: str | Path) -> None:
    """Convert PDF pages to editable, coordinate-preserving DOCX objects."""
    doc = Document()
    first = True
    shape_id = 1
    for page in iter_pages(pdf_path):
        section = doc.sections[0] if first else doc.add_section(WD_SECTION.NEW_PAGE)
        first = False
        _configure_section(section, page)
        doc.add_paragraph().paragraph_format.space_after = Pt(0)
        for _, kind, item in _content_items(page):
            if kind == 0:
                _add_positioned_image(doc, item, shape_id)
            else:
                _add_text_box(doc, item, shape_id)
            shape_id += 1
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)


def _convert_editable(pdf_path: str | Path, output_path: str | Path) -> None:
    _convert_positioned_editable(pdf_path, output_path)


def convert_pdf_to_docx(pdf_path: str | Path, output_path: str | Path, mode: str = "fidelity", fidelity_dpi: int = 180) -> None:
    mode = (mode or "fidelity").lower().strip()
    if mode not in {"fidelity", "editable"}:
        raise ValueError("mode must be 'fidelity' or 'editable'")
    # fidelity_dpi is retained for API compatibility. Fidelity mode is text/vector
    # based and intentionally does not rasterize the PDF.
    _convert_positioned_editable(pdf_path, output_path)
