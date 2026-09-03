from pathlib import Path
from zipfile import ZipFile

import fitz

from backend.app.converter.docx_builder import convert_pdf_to_docx


def make_pdf(path: Path) -> None:
    """Build a small fixture that mirrors the supplied CV: 3 pages, two images on page 1, text on all pages."""
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)
    page.insert_text((55, 70), "JAYDIP CV", fontsize=24)
    page.insert_text((55, 115), "ABOUT ME", fontsize=12)
    page.insert_text((55, 140), "Software professional with editable content.", fontsize=10)
    page.insert_text((55, 210), "WORK EXPERIENCE", fontsize=12)
    page.insert_text((55, 235), "Company Name — Software Engineer", fontsize=10)

    # Two different image assets, like the portrait + logo on the supplied CV.
    portrait = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 24, 24), False)
    portrait.clear_with(0xDDDDDD)
    logo = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 24, 24), False)
    logo.clear_with(0x6699CC)
    page.insert_image(fitz.Rect(470, 45, 545, 120), pixmap=portrait)
    page.insert_image(fitz.Rect(400, 45, 455, 100), pixmap=logo)

    page = doc.new_page(width=595, height=842)
    page.insert_text((55, 70), "WORK EXPERIENCE", fontsize=18)
    page.insert_text((55, 110), "Additional professional experience", fontsize=10)
    page.insert_text((55, 170), "EDUCATION AND TRAINING", fontsize=12)
    page.insert_text((55, 200), "Degree / Institution / Dates", fontsize=10)

    page = doc.new_page(width=595, height=842)
    page.insert_text((55, 70), "LANGUAGE SKILLS", fontsize=18)
    page.insert_text((55, 110), "English    C1    C1    B2    B2", fontsize=10)
    page.insert_text((55, 160), "DIGITAL SKILLS", fontsize=12)
    page.insert_text((55, 190), "Office tools, web technologies, communication", fontsize=10)
    doc.save(path)
    doc.close()


def test_fidelity_mode_preserves_cv_structure(tmp_path):
    source = tmp_path / "sample.pdf"
    target = tmp_path / "sample.docx"
    make_pdf(source)
    convert_pdf_to_docx(source, target, mode="fidelity", fidelity_dpi=120)

    with ZipFile(target) as archive:
        media = [name for name in archive.namelist() if name.startswith("word/media/")]
        document_xml = archive.read("word/document.xml").decode("utf-8")

    assert len(media) == 2
    assert document_xml.count("<w:sectPr") == 3
    assert document_xml.count("TextBox") >= 3
    assert "JAYDIP CV" in document_xml
    assert "ABOUT ME" in document_xml
    assert "WORK EXPERIENCE" in document_xml
    assert "EDUCATION AND TRAINING" in document_xml
    assert "LANGUAGE SKILLS" in document_xml
    assert "DIGITAL SKILLS" in document_xml
