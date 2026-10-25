"""Exact parsing and Romanian report formatting. Pure functions, no I/O."""

from __future__ import annotations

import re
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation

MILLION = Decimal("1000000")


class AmbiguousNumber(ValueError):
    pass


def to_decimal(value) -> Decimal:
    """Exact Decimal from a workbook value. Floats go through their shortest repr,
    which is the text Excel stored, so 262449.6 stays 262449.6."""
    if isinstance(value, bool):
        raise TypeError("boolean is not a number")
    if isinstance(value, Decimal):
        return value
    if isinstance(value, (int, float)):
        return Decimal(repr(value)) if isinstance(value, float) else Decimal(value)
    raise TypeError(f"not a number: {value!r}")


_SPACES = re.compile(r"[\s  ]")


def parse_ro(text: str) -> Decimal:
    """Parse a Romanian-formatted number: '.' groups thousands, ',' is the decimal mark.

    '(15.000.000)' and a Unicode minus are negative. '1,234' and '1.234' alone are
    ambiguous only when they could be either convention; we reject the comma-three-digits
    case (1,234) because it is the common English thousands form."""
    s = _SPACES.sub("", text).replace("−", "-")
    negative = s.startswith("(") and s.endswith(")")
    if negative:
        s = s[1:-1]
    if s.startswith("-"):
        negative, s = True, s[1:]
    if not re.fullmatch(r"\d{1,3}(\.\d{3})*(,\d+)?|\d+(,\d+)?", s):
        raise ValueError(f"not a Romanian number: {text!r}")
    if re.fullmatch(r"\d{1,3},\d{3}", s):
        raise AmbiguousNumber(f"ambiguous number: {text!r}")
    d = Decimal(s.replace(".", "").replace(",", "."))
    return -d if negative else d


def _round(value: Decimal, places: int) -> Decimal:
    return value.quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_UP)


def _group(int_part: str) -> str:
    out = []
    while len(int_part) > 3:
        out.insert(0, int_part[-3:])
        int_part = int_part[:-3]
    out.insert(0, int_part)
    return ".".join(out)


def _ro(value: Decimal, places: int) -> str:
    r = _round(value, places)
    if r == 0:
        r = abs(r)  # never print "-0,0"
    sign = "-" if r < 0 else ""
    int_part, _, frac = f"{abs(r):f}".partition(".")
    return sign + _group(int_part) + ("," + frac if places else "")


def fmt_mil_lei(value_lei: Decimal, places: int = 1) -> str:
    """284600000 -> '284,6 mil. lei' (rounded half up, once, on the unrounded value)."""
    return f"{_ro(value_lei / MILLION, places)} mil. lei"


def fmt_pct(fraction: Decimal, places: int = 1) -> str:
    """1.874 -> '187,4 %'."""
    return f"{_ro(fraction * 100, places)} %"


def fmt_count(value: Decimal) -> str:
    """18420 -> '18.420'."""
    return _ro(value, 0)


def fmt_number_only(value_lei: Decimal, kind: str) -> str:
    """The bare figure as it appears in report text, used for mapping/trace comparisons."""
    if kind == "amount":
        return _ro(value_lei / MILLION, 1)
    if kind == "amount_abs":
        return _ro(abs(value_lei) / MILLION, 1)
    if kind == "ratio":
        return _ro(value_lei * 100, 1)
    if kind == "count":
        return _ro(value_lei, 0)
    raise ValueError(f"unknown kind {kind!r}")


def fmt_fact(value: Decimal, kind: str) -> str:
    if kind == "amount":
        return fmt_mil_lei(value)
    if kind == "ratio":
        return fmt_pct(value)
    if kind == "count":
        return fmt_count(value)
    raise ValueError(f"unknown kind {kind!r}")


def dec(text: str) -> Decimal:
    """Decimal from a config string; config numbers are always quoted strings."""
    try:
        return Decimal(str(text))
    except InvalidOperation as e:
        raise ValueError(f"not a decimal: {text!r}") from e
