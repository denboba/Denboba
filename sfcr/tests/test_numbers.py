from decimal import Decimal

import pytest

from sfcr.numbers import AmbiguousNumber, fmt_count, fmt_mil_lei, fmt_pct, parse_ro, to_decimal


def test_amounts_round_half_up_once_on_unrounded_value():
    assert fmt_mil_lei(Decimal("284600000")) == "284,6 mil. lei"
    assert fmt_mil_lei(Decimal("412350000")) == "412,4 mil. lei"   # x,x5 goes up
    assert fmt_mil_lei(Decimal("45570000")) == "45,6 mil. lei"
    assert fmt_mil_lei(Decimal("1234567890")) == "1.234,6 mil. lei"
    assert fmt_mil_lei(Decimal("-28500000")) == "-28,5 mil. lei"
    assert fmt_mil_lei(Decimal("-10000")) == "0,0 mil. lei"        # no "-0,0"


def test_percent_and_count():
    assert fmt_pct(Decimal("1.874")) == "187,4 %"
    assert fmt_pct(Decimal("293699.3") / Decimal("148763.2")) == "197,4 %"
    assert fmt_pct(Decimal("0.40845")) == "40,8 %"
    assert fmt_count(Decimal("18420")) == "18.420"


def test_total_comes_from_source_not_from_rounded_parts():
    t1, t2 = Decimal("262449600"), Decimal("31249700")
    assert fmt_mil_lei(t1 + t2) == "293,7 mil. lei"
    assert (Decimal("262.4") + Decimal("31.2")) != Decimal("293.7")


def test_parse_romanian():
    assert parse_ro("284,6") == Decimal("284.6")
    assert parse_ro("18.420") == Decimal("18420")
    assert parse_ro("1.234,56") == Decimal("1234.56")
    assert parse_ro("(15.000.000)") == Decimal("-15000000")
    assert parse_ro("−12,5") == Decimal("-12.5")
    assert parse_ro("1 234,5".replace(" ", "")) == Decimal("1234.5")
    with pytest.raises(AmbiguousNumber):
        parse_ro("1,234")
    with pytest.raises(ValueError):
        parse_ro("12,3,4")


def test_float_cells_become_exact_decimals():
    assert to_decimal(262449.6) == Decimal("262449.6")
    assert to_decimal(0.1) + to_decimal(0.2) == Decimal("0.3")
    with pytest.raises(TypeError):
        to_decimal(True)
