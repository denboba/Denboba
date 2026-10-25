"""Text documents (txt, docx, pdf with a text layer): read lines with their page, find a
labelled figure, verify that a quoted sentence is really in the document."""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path

from .numbers import parse_ro
from .workbook import norm

NUMBER = r"\(?-?\d{1,3}(?:\.\d{3})+(?:,\d+)?\)?|\(?-?\d+(?:,\d+)?\)?"


def read_lines(path: Path) -> list[tuple[int, str]]:
    """(page, line) pairs. Pages are 1-based; txt and docx count as page 1."""
    suffix = path.suffix.lower()
    if suffix == ".txt":
        return [(1, line) for line in path.read_text(encoding="utf-8").splitlines()]
    if suffix == ".docx":
        import docx

        return [(1, p.text) for p in docx.Document(str(path)).paragraphs]
    if suffix == ".pdf":
        import pdfplumber

        out = []
        with pdfplumber.open(str(path)) as pdf:
            for i, page in enumerate(pdf.pages, start=1):
                out += [(i, line) for line in (page.extract_text() or "").splitlines()]
        return out
    raise ValueError(f"unsupported document type: {path.name}")


@dataclass
class DocFigure:
    status: str  # verified | missing | uncertain
    value: Decimal | None = None
    file: str = ""
    page: int | None = None
    excerpt: str = ""
    why: list[str] = field(default_factory=list)
    reason_ro: str = ""


def _unit_factor(text: str, units: list[dict]) -> Decimal | None:
    t = norm(text)
    for u in sorted(units, key=lambda u: -len(u["phrase"])):
        if re.search(r"(^|[\s(])" + re.escape(norm(u["phrase"])) + r"($|[\s),])", t):
            return Decimal(u["factor"])
    return None


def table_row_figure(path: Path, row_label: str, column: str, units: list[dict]) -> DocFigure:
    """A row of a text table: 'Label  (15.000.000)  (12.000.000)' under a header line
    'Indicator (lei)  <year>  <prior year>'. The column is chosen by header text, the unit from the
    header, and parentheses mean a negative (outflow)."""
    fig = DocFigure(status="missing", file=path.name)
    lines = read_lines(path)
    for idx, (page, line) in enumerate(lines):
        if not norm(line).startswith(norm(row_label)):
            continue
        rest = line[len(row_label):]
        rest = re.sub(r"^[^()\d]*\([^)]*[A-Za-zăâîșțĂÂÎȘȚ][^)]*\)", "", rest)  # drop '(fluxuri ...)' qualifier
        numbers = re.findall(NUMBER, rest)
        header = next((h for p, h in reversed(lines[:idx]) if re.search(rf"(^|\s){re.escape(column)}(\s|$)", h)), None)
        if header is None:
            fig.status, fig.why = "uncertain", [f"no header line with column {column!r} above the row"]
            fig.reason_ro = f"coloana {column} nu este identificabilă în {path.name}"
            return fig
        cols = re.findall(r"(?<!\d)\d{4}(?!\d)", header)
        if column not in cols or len(cols) != len(numbers):
            fig.status = "uncertain"
            fig.why = [f"row {line!r}: {len(numbers)} numbers vs header columns {cols}"]
            fig.reason_ro = f"rândul „{row_label}” din {path.name} nu poate fi citit sigur"
            return fig
        factor = _unit_factor(header, units)
        if factor is None:
            fig.status = "uncertain"
            fig.why = [f"no unit in header {header!r}"]
            fig.reason_ro = f"unitatea de măsură lipsește în {path.name}"
            return fig
        raw = numbers[cols.index(column)]
        fig.value = parse_ro(raw) * factor
        fig.page, fig.excerpt, fig.status = page, line.strip(), "verified"
        fig.why = [
            f"{path.name} page {page}: row {row_label!r}, column {column} of header {header.strip()!r} "
            f"-> {raw} x {factor} lei"
        ]
        return fig
    fig.why = [f"no line starting with {row_label!r} in {path.name}"]
    fig.reason_ro = f"rândul „{row_label}” lipsește din {path.name}"
    return fig


def after_text_figure(path: Path, anchor: str, units: list[dict]) -> DocFigure:
    """The first number after an anchor phrase, with the unit word right after it."""
    fig = DocFigure(status="missing", file=path.name)
    for page, line in read_lines(path):
        hit = re.search(re.escape(anchor), unicodedata.normalize("NFC", line), re.IGNORECASE)
        if not hit:
            continue
        tail = unicodedata.normalize("NFC", line)[hit.end():]
        m = re.match(rf"\s*({NUMBER})\s*(\S+(?:\s\S+)?)", tail)
        factor = _unit_factor(m.group(2), units) if m else None
        if not m or factor is None:
            fig.status, fig.why = "uncertain", [f"anchor {anchor!r} found but no number + unit follows"]
            fig.reason_ro = f"suma după „{anchor}” din {path.name} nu poate fi citită"
            return fig
        fig.value = parse_ro(m.group(1)) * factor
        fig.page, fig.excerpt, fig.status = page, line.strip(), "verified"
        fig.why = [f"{path.name} page {page}: {m.group(1)} {m.group(2)!r} after {anchor!r}"]
        return fig
    fig.why = [f"anchor {anchor!r} not in {path.name}"]
    fig.reason_ro = f"textul „{anchor}” lipsește din {path.name}"
    return fig


def find_quote(path: Path, quote: str) -> int | None:
    """Page where the quote appears verbatim (after normalization), else None."""
    wanted = norm(quote)
    pages: dict[int, list[str]] = {}
    for page, line in read_lines(path):
        pages.setdefault(page, []).append(line)
    for page, lines in pages.items():
        if wanted in norm(" ".join(lines)):
            return page
    return None
