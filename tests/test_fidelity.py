from pathlib import Path
from zipfile import ZipFile

import fitz

from backend.app.converter.docx_builder import convert_pdf_to_docx


def make_pdf(path: Path) -> None:
    doc = fitz.open()
    for index in range(2):
        page = doc.new_page(width=595, height=842)
        page.insert_text((50, 80), f"Page {index + 1}", fontsize=24)
        page.insert_text((300, 180), "Positioned text", fontsize=12)
        pix = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 12, 12), False)
        pix.clear_with(255)
        page.insert_image(fitz.Rect(500, 50, 540, 90), pixmap=pix)
    doc.save(path)
    doc.close()


def test_fidelity_mode_preserves_page_count(tmp_path):
    source = tmp_path / "sample.pdf"
    target = tmp_path / "sample.docx"
    make_pdf(source)
    convert_pdf_to_docx(source, target, mode="fidelity", fidelity_dpi=120)

    with ZipFile(target) as archive:
        media = [name for name in archive.namelist() if name.startswith("word/media/")]
        document_xml = archive.read("word/document.xml").decode("utf-8")

    assert len(media) == 2
    assert document_xml.count("<w:sectPr") == 2
    assert "TextBox" in document_xml
    assert "Page 1" in document_xml and "Page 2" in document_xml
    assert "Positioned text" in document_xml
