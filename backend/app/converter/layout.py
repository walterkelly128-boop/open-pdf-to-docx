from dataclasses import dataclass

from .pdf_parser import PdfPage, PdfWord


@dataclass
class LayoutLine:
    words: list[PdfWord]
    x0: float
    y0: float
    x1: float
    y1: float


@dataclass
class LayoutBlock:
    lines: list[LayoutLine]
    x0: float
    y0: float
    x1: float
    y1: float


def build_layout(page: PdfPage) -> list[LayoutBlock]:
    if not page.words:
        return []

    # Keep the PDF's original block/line information where possible.
    grouped: dict[tuple[int, int], list[PdfWord]] = {}
    for word in page.words:
        grouped.setdefault((word.block_no, word.line_no), []).append(word)

    lines_by_block: dict[int, list[LayoutLine]] = {}
    for (block_no, _), words in grouped.items():
        words.sort(key=lambda w: w.x0)
        lines_by_block.setdefault(block_no, []).append(
            LayoutLine(words, min(w.x0 for w in words), min(w.y0 for w in words),
                       max(w.x1 for w in words), max(w.y1 for w in words))
        )

    blocks: list[LayoutBlock] = []
    for lines in lines_by_block.values():
        lines.sort(key=lambda line: line.y0)
        blocks.append(LayoutBlock(
            lines=lines,
            x0=min(line.x0 for line in lines),
            y0=min(line.y0 for line in lines),
            x1=max(line.x1 for line in lines),
            y1=max(line.y1 for line in lines),
        ))
    blocks.sort(key=lambda block: (block.y0, block.x0))
    return blocks
