import copy
import os
from decimal import Decimal
from pathlib import Path

import pytest

from conftest import CONFIG, ROOT, SCR_CAUSE, config
from sfcr.render import similarity_check, trace_check
from sfcr.workbook import Workbook


def test_happy_path_moved_rows_units_and_uncached_formula(make_run):
    run, text, log = make_run()
    f = run.facts["eof_total"]
    assert f.status == "verified" and f.source["cell"] == "C8"          # row moved from C7 to C8
    assert f.value == Decimal("255000000")                               # mii lei -> lei
    assert any("no cached result" in w for w in f.why)                   # openpyxl writes no cache
    assert "255,0 mil. lei" in text and "nivel 1 de 230,0 mil. lei" in text
    assert "Societatea nu a deținut fonduri proprii de nivel 3." in text
    assert "255,0 %" in text                                             # 255 / 100, derived
    assert "Cerința de capital minim (MCR) a fost de 45,0 mil. lei și a fost acoperită integral." in text
    assert "în scădere față de 110,0 mil. lei" in text and "riscului de piață" in text
    assert trace_check(run, text, log) == []
    assert "[de completat" not in text


def test_hidden_sheet_is_never_read(tmp_path):
    from conftest import build
    build(tmp_path)
    wb = Workbook(tmp_path / "registru_2025.xlsx", CONFIG["workbook_layout"]["units"], 3, "Valori în")
    loc = wb.locate("Capital_vechi", "Fonduri proprii eligibile SCR, nivel 1", "31.12 an curent")
    assert loc.status == "uncertain" and loc.value is None
    assert wb.hidden_sheets() == ["Capital_vechi"]


def test_duplicate_label_is_uncertain(tmp_path):
    from conftest import build
    import openpyxl
    build(tmp_path)
    p = tmp_path / "registru_2025.xlsx"
    wb = openpyxl.load_workbook(p)
    wb["Capital"]["B20"] = "SCR"
    wb["Capital"]["C20"] = 1
    wb.save(p)
    loc = Workbook(p, CONFIG["workbook_layout"]["units"], 3, "Valori în").locate("Capital", "SCR", "31.12 an curent")
    assert loc.status == "uncertain"


def test_missing_mcr_is_a_gap_never_last_years_value(make_run):
    run, text, _ = make_run(capital={"mcr": None}, mcr_note="în lucru, actuariat")
    assert run.facts["mcr"].status == "missing"
    assert "[de completat:" in text and "în lucru, actuariat" in text
    assert "40,0 mil. lei" not in text                                   # the 2024 MCR


def test_ratio_conflict_without_matching_decision_is_a_gap(make_run):
    run, text, _ = make_run(capital={"ratio": 2.0})                     # stale: equals 2024
    f = run.facts["solvency_ratio"]
    assert f.status == "conflict" and f.decision is None                 # DEC-1 is pinned to other values
    assert "Rata de solvabilitate a fost de" not in text
    assert "[de completat: rata de solvabilitate" in text
    assert any(i.code == "POSSIBLE_CARRY_FORWARD" for i in run.issues)


def test_pinned_decision_applies_only_to_its_evidence(make_run):
    cfg = config()
    cfg["decision"] = [{**CONFIG["decision"][0], "when_workbook_value": "2.0", "when_computed": "255,0 %"}]
    run, text, _ = make_run(cfg=cfg, capital={"ratio": 2.0})
    assert run.facts["solvency_ratio"].decision == "DEC-1" and "255,0 %" in text
    run, text, _ = make_run(cfg=cfg, capital={"ratio": 2.0, "t1": 231000})   # computed changes -> not applied
    assert run.facts["solvency_ratio"].decision is None


def test_direction_word_follows_the_data_and_cause_needs_conditions(make_run):
    _, text, _ = make_run(capital={"scr": 120000}, modules={"op": 25000, "scr": 120000})
    assert "în creștere față de 110,0 mil. lei" in text
    assert "riscului de piață" not in text                              # the claim's condition fails


def test_claim_dropped_when_quote_not_in_source(make_run):
    _, text, _ = make_run(note=["Formula standard este folosită."])
    assert "în scădere față de 110,0 mil. lei la 31 decembrie 2024." in text
    assert "riscului de piață" not in text
    assert "formula standard și nu utilizează" not in text
    assert "[de completat: afirmația CLM-standard-formula" in text


def test_failed_tier_sum_keeps_figures_out_of_the_text(make_run):
    import openpyxl
    run, text, _ = make_run()
    p = run.base / "registru_2025.xlsx"
    wb = openpyxl.load_workbook(p)
    wb["Capital"]["C8"] = 999999                                         # total no longer the sum
    wb.save(p)
    from sfcr.engine import Run
    from sfcr.render import render
    run = Run(config(), run.base).run()
    text, _ = render(run)
    assert run.facts["eof_total"].status == "uncertain"
    assert "EOF_TIERS" in text and "999" not in text


def test_wrong_unit_is_caught_by_plausibility(make_run):
    run, text, _ = make_run(capital_unit="Valori în lei")               # values still in thousands
    assert run.facts["scr"].status == "uncertain"
    assert any(i.code == "UNIT_MISMATCH" for i in run.issues)
    assert "0,1 mil. lei" not in text


def test_dividends_must_be_corroborated(make_run):
    run, text, _ = make_run(aga_amount="16.000.000")
    assert run.facts["dividends_paid"].status == "conflict"
    assert "dividende de 15,0" not in text


def test_similarity_detects_copied_wording():
    other = "Poziția de solvabilitate a rămas robustă pe parcursul întregului exercițiu financiar."
    assert similarity_check("Noi scriem: poziția de solvabilitate a rămas robustă pe parcursul anului.", other, 6)
    assert similarity_check("SCR a fost de 148,8 mil. lei.", other, 6) == []


DATE = ROOT / "date"


@pytest.mark.skipif(not (DATE / "registru_2025.xlsx").exists(), reason="client materials not present in date/")
def test_real_materials(monkeypatch):
    from sfcr.engine import Run
    from sfcr.render import render
    run = Run(copy.deepcopy(CONFIG), DATE).run()
    text, log = render(run)
    assert "293,7 mil. lei" in text and "262,4 mil. lei" in text and "31,2 mil. lei" in text
    assert "148,8 mil. lei, în scădere față de 151,9 mil. lei" in text
    assert "197,4 %" in text and "187,4 %" not in text                   # stale C12 not used
    assert "45,6" not in text and "MCR" in text and "în lucru, actuariat" in text
    assert "dividende de 15,0 mil. lei" in text
    assert "programul de reasigurare" not in text and "program de reasigurare" not in text  # CFO request
    assert trace_check(run, text, log) == []
    other = (DATE / CONFIG["inputs"]["other_company_report"]).read_text(encoding="utf-8")
    assert similarity_check(text, other, 6) == []
    mapping_problems = {m["text"] for m in run.mapping_audit if m["problems"]}
    assert "12,6" in mapping_problems
