"""Locate a figure in a workbook by sheet + row label + column header, never by address.

Returns a Located record that says what was found and why, or why not. Formatting and
judgement (checks, decisions) happen elsewhere."""

from __future__ import annotations

import difflib
import re
import unicodedata
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter, range_boundaries

from .numbers import AmbiguousNumber, parse_ro, to_decimal


def norm(text: str) -> str:
    """Comparison key: NFC, legacy cedilla -> comma-below, case-fold, single spaces."""
    t = unicodedata.normalize("NFC", str(text))
    t = t.translate(str.maketrans("şŞţŢ", "șȘțȚ"))
    return re.sub(r"\s+", " ", t).strip().casefold()


@dataclass
class Located:
    status: str  # verified | missing | uncertain
    value: Decimal | None = None  # in lei (amounts) or as a fraction (ratios)
    raw: object = None
    file: str = ""
    sheet: str = ""
    cell: str = ""
    label: str = ""
    unit: str = ""
    formula: str | None = None
    factor: Decimal | None = None  # lei per workbook unit, from the sheet's unit note
    unit_src: str = ""
    number_format: str = ""
    why: list[str] = field(default_factory=list)  # English, for the reviewer
    reason_ro: str = ""  # short Romanian reason, used in [de completat: ...]


class Workbook:
    def __init__(self, path: Path, units: list[dict], header_rows: int, unit_prefix: str):
        self.path = path
        self.name = path.name
        # Two views: formulas (to see what a cell is) and cached values (what Excel computed).
        self.wb_f = openpyxl.load_workbook(path, data_only=False, read_only=False)
        self.wb_v = openpyxl.load_workbook(path, data_only=True, read_only=False)
        self.units = sorted(units, key=lambda u: -len(u["phrase"]))
        self.header_rows = header_rows
        self.unit_prefix = unit_prefix

    def hidden_sheets(self) -> list[str]:
        return [ws.title for ws in self.wb_f.worksheets if ws.sheet_state != "visible"]

    # -- unit -----------------------------------------------------------------------
    def sheet_unit(self, title: str) -> tuple[str | None, Decimal | None, str]:
        """Find the unit note (e.g. 'Valori în mii lei') in the top rows; longest known phrase wins."""
        ws = self.wb_f[title]
        for row in ws.iter_rows(min_row=1, max_row=self.header_rows):
            for c in row:
                if isinstance(c.value, str) and norm(c.value).startswith(norm(self.unit_prefix)):
                    text = norm(c.value)
                    for u in self.units:
                        if norm(u["phrase"]) in text:
                            return u["phrase"], Decimal(u["factor"]), f"{c.coordinate}: {c.value!r}"
                    return None, None, f"{c.coordinate}: {c.value!r}"
        return None, None, ""

    # -- formula evaluation (only + / - of same-sheet cells and SUM ranges) -----------
    def _eval(self, title: str, formula: str, depth: int = 0) -> Decimal:
        if depth > 5:
            raise ValueError("formula nesting too deep")
        expr = formula.lstrip("=").replace(" ", "").upper()
        total = Decimal(0)
        if not re.fullmatch(r"[+-]?(SUM\([A-Z]+\d+:[A-Z]+\d+\)|[A-Z]+\d+)([+-](SUM\([A-Z]+\d+:[A-Z]+\d+\)|[A-Z]+\d+))*", expr):
            raise ValueError(f"unsupported formula {formula!r}")
        for sign, term in re.findall(r"([+-]?)(SUM\([A-Z]+\d+:[A-Z]+\d+\)|[A-Z]+\d+)", expr):
            if term.startswith("SUM("):
                c1, r1, c2, r2 = range_boundaries(term[4:-1])
                refs = [f"{get_column_letter(c)}{r}" for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]
            else:
                refs = [term]
            part = Decimal(0)
            for ref in refs:
                v = self.wb_f[title][ref].value
                if isinstance(v, str) and v.startswith("="):
                    part += self._eval(title, v, depth + 1)
                elif v is None:
                    raise ValueError(f"{ref} is empty")
                else:
                    part += to_decimal(v)
            total += -part if sign == "-" else part
        return total

    # -- the locator ------------------------------------------------------------------
    def locate(self, sheet: str, label: str, header: str) -> Located:
        loc = Located(status="missing", file=self.name, sheet=sheet, label=label)
        titles = {ws.title: ws for ws in self.wb_f.worksheets}
        if sheet not in titles:
            loc.why.append(f"sheet {sheet!r} not found; sheets: {list(titles)}")
            loc.reason_ro = f"foaia „{sheet}” lipsește din {self.name}"
            return loc
        ws = titles[sheet]
        if ws.sheet_state != "visible":
            loc.status = "uncertain"
            loc.why.append(f"sheet {sheet!r} is {ws.sheet_state}; hidden sheets are never used")
            loc.reason_ro = f"foaia „{sheet}” din {self.name} este ascunsă"
            return loc

        label_cells, header_cells, all_labels = [], [], []
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str):
                    if norm(c.value) == norm(label):
                        label_cells.append(c)
                    if norm(c.value) == norm(header):
                        header_cells.append(c)
                    all_labels.append(c.value)

        if not label_cells:
            hints = difflib.get_close_matches(label, all_labels, n=3, cutoff=0.6)
            loc.why.append(f"no cell labelled {label!r} in {sheet}; closest labels (not used): {hints}")
            loc.reason_ro = f"rândul „{label}” nu există în {self.name}, foaia {sheet}"
            return loc
        if len(label_cells) > 1:
            loc.status = "uncertain"
            loc.why.append(f"label {label!r} appears {len(label_cells)} times: {[c.coordinate for c in label_cells]}")
            loc.reason_ro = f"rândul „{label}” apare de mai multe ori în {self.name}, foaia {sheet}"
            return loc
        if len(header_cells) != 1:
            loc.status = "uncertain"
            loc.why.append(f"column header {header!r} found {len(header_cells)} times in {sheet}")
            loc.reason_ro = f"coloana „{header}” nu poate fi identificată în {self.name}, foaia {sheet}"
            return loc

        lc, hc = label_cells[0], header_cells[0]
        coord = f"{hc.column_letter}{lc.row}"
        loc.cell = coord
        loc.why.append(
            f"label {lc.value!r} found at {sheet}!{lc.coordinate} (row {lc.row}); "
            f"header {hc.value!r} at {sheet}!{hc.coordinate} (column {hc.column_letter}) -> {sheet}!{coord}"
        )

        unit, factor, unit_src = self.sheet_unit(sheet)
        cell_f = ws[coord]
        raw_f = cell_f.value
        raw_v = self.wb_v[sheet][coord].value
        loc.raw = raw_v if raw_v is not None else raw_f

        if raw_f is None:
            notes = [f"{c.coordinate}: {c.value!r}" for c in ws[lc.row] if c.value is not None and c is not lc]
            loc.why.append(f"{sheet}!{coord} is empty" + (f"; other cells in the row: {notes}" if notes else ""))
            note_ro = "; ".join(str(c.value) for c in ws[lc.row] if isinstance(c.value, str) and c is not lc)
            loc.reason_ro = (
                f"{self.name}, foaia {sheet}, rândul „{lc.value}”: celula {coord} este goală"
                + (f" (notă în registru: „{note_ro}”)" if note_ro else "")
            )
            return loc

        try:
            if isinstance(raw_f, str) and raw_f.startswith("="):
                loc.formula = raw_f
                try:
                    evaluated = self._eval(sheet, raw_f)
                except ValueError as e:
                    if raw_v is None or isinstance(raw_v, (str, bool)):
                        raise ValueError(f"formula not evaluable here and no cached result ({e})") from e
                    evaluated = None
                if evaluated is None:
                    value = to_decimal(raw_v)
                    loc.why.append(f"formula {raw_f} is beyond the evaluator; Excel's cached result {raw_v} used")
                elif raw_v is None:
                    value = evaluated
                    loc.why.append(f"formula {raw_f} has no cached result; evaluated from its inputs = {evaluated}")
                else:
                    value = to_decimal(raw_v)
                    if value != evaluated:
                        loc.status = "uncertain"
                        loc.why.append(f"cached value {raw_v} differs from formula {raw_f} = {evaluated} (stale cache)")
                        loc.reason_ro = f"{self.name}, {sheet}!{coord}: valoarea salvată nu corespunde formulei"
                        return loc
                    loc.why.append(f"formula {raw_f}; cached value {raw_v} agrees with re-evaluation")
            elif isinstance(raw_f, bool):
                raise TypeError("boolean")
            elif isinstance(raw_f, (int, float)):
                value = to_decimal(raw_f)
                loc.why.append(f"typed constant {raw_f!r} (not a formula)")
            elif isinstance(raw_f, str):
                value = parse_ro(raw_f)
                loc.why.append(f"number stored as text {raw_f!r}, parsed with Romanian rules")
            else:
                raise TypeError(type(raw_f).__name__)
        except AmbiguousNumber:
            loc.status = "uncertain"
            loc.why.append(f"{sheet}!{coord} = {raw_f!r} is ambiguous")
            loc.reason_ro = f"{self.name}, {sheet}!{coord}: valoare ambiguă „{raw_f}”"
            return loc
        except (TypeError, ValueError) as e:
            loc.status = "uncertain"
            loc.why.append(f"{sheet}!{coord} = {raw_f!r} cannot be read as a number ({e})")
            loc.reason_ro = f"{self.name}, {sheet}!{coord}: valoarea „{raw_f}” nu este un număr"
            return loc

        loc.value = value
        loc.unit, loc.factor, loc.unit_src = unit or "", factor, unit_src
        loc.number_format = cell_f.number_format
        loc.status = "verified"
        return loc
