from dataclasses import dataclass

from .pdf_parser import PdfPage, PdfWord


@dataclass
class LayoutLine:
    words: list[PdfWord]
    x0: float
    y0: float
    x1: float
    y1: float

    @property
    def height(self) -> float:
        return max(1.0, self.y1 - self.y0)


@dataclass
class LayoutBlock:
    lines: list[LayoutLine]
    x0: float
    y0: float
    x1: float
    y1: float
    column: int = 0


def _make_lines(page: PdfPage) -> list[LayoutLine]:
    grouped: dict[tuple[int, int], list[PdfWord]] = {}
    for word in page.words:
        grouped.setdefault((word.block_no, word.line_no), []).append(word)

    lines: list[LayoutLine] = []
    for words in grouped.values():
        words.sort(key=lambda w: (w.x0, w.y0))
        lines.append(LayoutLine(
            words=words,
            x0=min(w.x0 for w in words), y0=min(w.y0 for w in words),
            x1=max(w.x1 for w in words), y1=max(w.y1 for w in words),
        ))
    return lines


def _merge_lines(lines: list[LayoutLine]) -> list[LayoutBlock]:
    """Merge PDF lines into editable paragraphs while keeping headings/list items separate."""
    lines = sorted(lines, key=lambda l: (l.y0, l.x0))
    blocks: list[LayoutBlock] = []
    current: list[LayoutLine] = []

    for line in lines:
        if not current:
            current = [line]
            continue
        prev = current[-1]
        gap = line.y0 - prev.y1
        size = sum(w.size for w in prev.words) / max(1, len(prev.words))
        x_shift = abs(line.x0 - prev.x0)
        # A large vertical gap, strong indentation change, or a very large font
        # usually means a new paragraph/heading. PDF blocks are also boundaries.
        same_block = line.words[0].block_no == prev.words[0].block_no
        continuation = same_block and gap <= max(size * 0.85, 7) and x_shift <= max(size * 2.2, 18)
        if continuation:
            current.append(line)
        else:
            blocks.append(_block(current))
            current = [line]
    if current:
        blocks.append(_block(current))
    return blocks


def _block(lines: list[LayoutLine]) -> LayoutBlock:
    return LayoutBlock(
        lines=lines,
        x0=min(l.x0 for l in lines), y0=min(l.y0 for l in lines),
        x1=max(l.x1 for l in lines), y1=max(l.y1 for l in lines),
    )


def _detect_columns(blocks: list[LayoutBlock], page_width: float) -> list[LayoutBlock]:
    """Assign two-column reading order when page geometry strongly suggests it."""
    if len(blocks) < 4:
        return blocks
    gutter = page_width * 0.055
    mid = page_width / 2
    left = [b for b in blocks if b.x1 <= mid + gutter]
    right = [b for b in blocks if b.x0 >= mid - gutter]
    if len(left) < 2 or len(right) < 2:
        return blocks
    left_width = sum(b.x1 - b.x0 for b in left) / len(left)
    right_width = sum(b.x1 - b.x0 for b in right) / len(right)
    if min(left_width, right_width) < page_width * 0.22:
        return blocks
    for b in left:
        b.column = 0
    for b in right:
        b.column = 1
    return sorted(blocks, key=lambda b: (b.column, b.y0, b.x0))


def build_layout(page: PdfPage) -> list[LayoutBlock]:
    if not page.words:
        return []
    blocks = _merge_lines(_make_lines(page))
    return _detect_columns(blocks, page.width)
