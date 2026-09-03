from pathlib import Path
from zipfile import ZipFile

import fitz

from backend.app.converter.docx_builder import convert_pdf_to_docx


def make_pdf(path: Path) -> None:
    """Build a compact fixture based on the supplied Jaydipkumar Jani CV: 3 pages, two images on page 1, editable text throughout."""
    doc = fitz.open()

    # Page 1: personal information + ABOUT ME + WORK EXPERIENCE, matching the supplied CV.
    page = doc.new_page(width=595, height=842)
    page.insert_text((55, 70), "Jaydipkumar Jani", fontsize=24)
    page.insert_text((55, 105), "Passport: X5950476", fontsize=9)
    page.insert_text((55, 122), "Date of birth: 01/06/1992", fontsize=9)
    page.insert_text((55, 139), "Place of birth: Chappri, India", fontsize=9)
    page.insert_text((55, 156), "Nationality: Indian", fontsize=9)
    page.insert_text((55, 173), "Gender: Male", fontsize=9)
    page.insert_text((55, 190), "Phone number: (+91) 8140432035 (Mobile)", fontsize=9)
    page.insert_text((55, 207), "Email address: janijaydip1992@gmail.com", fontsize=9)
    page.insert_text((55, 224), "LinkedIn: https://www.linkedin.com/in/jaydip-jani-307686128", fontsize=8)
    page.insert_text((55, 241), "Whatsapp Messenger: +918140432035", fontsize=9)
    page.insert_text((55, 258), "Address: dharampur, district: Valsad, Gujarat, 396050, Valsad, India", fontsize=8)
    page.insert_text((55, 300), "ABOUT ME", fontsize=12)
    page.insert_text((55, 325), "I am a results-driven sales and business development professional with over 9 years of experience in the chemical and polymer", fontsize=8)
    page.insert_text((55, 340), "industry. Holding a Master’s degree in Polymer Science & Technology, I have successfully handled diverse product portfolios,", fontsize=8)
    page.insert_text((55, 355), "including raw materials for construction chemicals, personal care, paints and coatings, adhesives, FRP, and specialty polymers.", fontsize=8)
    page.insert_text((55, 370), "My expertise lies in developing new markets, generating leads, building strong client relationships, and providing technical solutions", fontsize=8)
    page.insert_text((55, 385), "in collaboration with R&D teams. I have worked with leading companies across India, managing B2B sales and technical support.", fontsize=8)
    page.insert_text((55, 420), "WORK EXPERIENCE", fontsize=12)
    page.insert_text((55, 445), "MARKETING & DEVELOPMENT MANAGER – TRUE CHEM – 28/07/2025 – 30/04/2026 – VALSAD, INDIA", fontsize=8)
    page.insert_text((55, 460), "• Lead marketing and business development for specialty chemical products.", fontsize=8)
    page.insert_text((55, 475), "• Identify and develop new market opportunities in construction chemicals, paints & coatings, personal care, and other industrial applications.", fontsize=7)
    page.insert_text((55, 490), "• Build and maintain strong relationships with key clients and distributors.", fontsize=8)

    # Two different image assets, matching the portrait + logo pattern of the supplied CV.
    portrait = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 24, 24), False)
    portrait.clear_with(0xDDDDDD)
    logo = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 24, 24), False)
    logo.clear_with(0x6699CC)
    page.insert_image(fitz.Rect(470, 45, 545, 120), pixmap=portrait)
    page.insert_image(fitz.Rect(400, 45, 455, 100), pixmap=logo)

    # Page 2: continued experience and education.
    page = doc.new_page(width=595, height=842)
    page.insert_text((55, 70), "WORK EXPERIENCE", fontsize=18)
    page.insert_text((55, 100), "AREA SALES MANAGER – WINWAYS CHEMTECH – 15/04/2024 – 26/07/2025 – VALSAD, INDIA", fontsize=8)
    page.insert_text((55, 125), "AREA SALES MANAGER – DEON TAPES INDUSTRIES PVT LTD – 26/09/2022 – 11/04/2024 – VALSAD, INDIA", fontsize=8)
    page.insert_text((55, 150), "BUSINESS DEVELOPMENT EXECUTIVE – WINDSON CHEMICALS PVT LTD – 08/06/2021 – 30/06/2022 – NAVSARI, INDIA", fontsize=8)
    page.insert_text((55, 175), "SALES EXECUTIVE – MACRO POLYMERS PVT LTD – 22/04/2019 – 01/01/2021 – AHMEDABAD, INDIA", fontsize=8)
    page.insert_text((55, 220), "EDUCATION AND TRAINING", fontsize=12)
    page.insert_text((55, 245), "10/06/2015 – 15/04/2017 Anand", fontsize=9)
    page.insert_text((55, 262), "M.SC. ( POLYMER SCIENCE & TECHNOLOGY ) Sardar Patel University", fontsize=9)
    page.insert_text((55, 280), "Key subjects, modules, and laboratory skills:", fontsize=8)
    page.insert_text((55, 297), "• Polymer synthesis, processing, and characterization techniques", fontsize=8)
    page.insert_text((55, 314), "• Polymer chemistry and physics fundamentals", fontsize=8)
    page.insert_text((55, 331), "• Polymer composites, blends, and nanocomposites", fontsize=8)
    page.insert_text((55, 348), "• Advanced polymer materials (engineering plastics, specialty polymers)", fontsize=8)
    page.insert_text((55, 365), "Website https://www.spuvvn.edu", fontsize=8)

    # Page 3: language and skills sections.
    page = doc.new_page(width=595, height=842)
    page.insert_text((55, 70), "LANGUAGE SKILLS", fontsize=18)
    page.insert_text((55, 105), "Mother tongue(s): GUJARATI", fontsize=9)
    page.insert_text((55, 125), "Other language(s):", fontsize=9)
    page.insert_text((55, 150), "UNDERSTANDING    SPEAKING    WRITING", fontsize=9)
    page.insert_text((55, 175), "ENGLISH    A2    A2    A2    A2    C2", fontsize=9)
    page.insert_text((55, 195), "HINDI    C2    C2    C2    C2    C2", fontsize=9)
    page.insert_text((55, 225), "Levels: A1 and A2: Basic user; B1 and B2: Independent user; C1 and C2: Proficient user", fontsize=8)
    page.insert_text((55, 270), "SKILLS", fontsize=12)
    page.insert_text((55, 295), "Organizational and planning skills", fontsize=9)
    page.insert_text((55, 312), "Team-work oriented", fontsize=9)
    page.insert_text((55, 329), "Decision-making", fontsize=9)
    page.insert_text((55, 346), "Good listener and communicator", fontsize=9)
    page.insert_text((55, 363), "Motivated", fontsize=9)
    page.insert_text((55, 405), "HOBBIES AND INTERESTS", fontsize=12)
    page.insert_text((55, 430), "playing valleyball in free time", fontsize=9)

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
    assert "Jaydipkumar Jani" in document_xml
    assert "ABOUT ME" in document_xml
    assert "WORK EXPERIENCE" in document_xml
    assert "EDUCATION AND TRAINING" in document_xml
    assert "LANGUAGE SKILLS" in document_xml
    assert "SKILLS" in document_xml
    assert "HOBBIES AND INTERESTS" in document_xml
