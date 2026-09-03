from backend.app.converter.layout import build_layout
from backend.app.converter.pdf_parser import PdfPage, PdfWord


def word(text, x0, y0, x1, y1, block=0, line=0):
    return PdfWord(text, x0, y0, x1, y1, 10, "Arial", 0, 0, block, line, 0)


def test_lines_are_merged_into_paragraphs():
    page = PdfPage(
        1,
        600,
        800,
        [
            word("Hello", 50, 50, 80, 60, 0, 0),
            word("world", 85, 50, 120, 60, 0, 0),
            word("Next", 50, 63, 75, 73, 0, 1),
            word("line", 80, 63, 100, 73, 0, 1),
        ],
    )
    blocks = build_layout(page)
    assert len(blocks) == 1
    assert len(blocks[0].lines) == 2
    assert blocks[0].lines[0].words[0].text == "Hello"


def test_large_vertical_gap_starts_new_block():
    page = PdfPage(
        1,
        600,
        800,
        [
            word("Heading", 50, 50, 110, 60, 0, 0),
            word("Body", 50, 100, 80, 110, 0, 1),
        ],
    )
    blocks = build_layout(page)
    assert len(blocks) == 2
