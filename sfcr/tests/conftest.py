"""Synthetic inputs shaped like the client's files, with invented numbers."""

import copy
import sys
import tomllib
from pathlib import Path

import docx
import openpyxl
import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

CONFIG = tomllib.loads((ROOT / "config" / "section_e.toml").read_text(encoding="utf-8"))

SCR_CAUSE = ("Cerința de capital de solvabilitate a scăzut la 31 decembrie 2025 față de anul anterior, "
             "în principal ca urmare a scăderii riscului de piață.")
NOTE = [SCR_CAUSE + " În trimestrul al treilea, Societatea a vândut integral portofoliul de acțiuni listate, "
        "iar sumele au fost plasate în titluri de stat.",
        "Calculele folosesc formula standard. Societatea nu utilizează parametri specifici proprii."]
FRAMEWORK = CONFIG["claim"][-1]["quote"]

# 2025 values in mii lei (Capital, SCR_module); 2024 in lei.
CAPITAL_2025 = {"t1": 230000, "t2": 25000, "t3": 0, "scr": 100000, "mcr": 45000, "ratio": "=C8/C10"}
MODULES_2025 = {"uw": 80000, "mkt": 30000, "cpty": 10000, "div": 25000, "basic": 95000, "op": 5000, "scr": 100000}


def _sheet(wb, title, unit, rows, prior=None):
    ws = wb.create_sheet(title)
    ws["A1"], ws["A2"] = title, unit
    ws["B3"], ws["C3"], ws["D3"] = "Indicator", "31.12 an curent", "31.12 an anterior"
    for r, (label, value) in rows.items():
        ws[f"B{r}"] = label
        if value is not None:
            ws[f"C{r}"] = value
        if isinstance(value, float) and 0 < value < 10 or label == "Rata de solvabilitate":
            ws[f"C{r}"].number_format = "0.0%"
    for r, v in (prior or {}).items():
        ws[f"D{r}"] = v
    return ws


def build(folder: Path, capital=None, modules=None, note=None, capital_unit="Valori în mii lei",
          mcr_note=None, hidden_copy=True, fs_dividends="(15.000.000)", aga_amount="15.000.000"):
    c = {**CAPITAL_2025, **(capital or {})}
    m = {**MODULES_2025, **(modules or {})}

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    _sheet(wb, "Capital", "Valori în lei", {
        5: ("Fonduri proprii eligibile SCR, nivel 1", 200000000), 6: ("Fonduri proprii eligibile SCR, nivel 2", 20000000),
        7: ("Fonduri proprii eligibile SCR, total", 220000000), 9: ("SCR", 110000000), 10: ("MCR", 40000000),
        11: ("Rata de solvabilitate", 2.0)})
    _sheet(wb, "SCR_module", "Valori în lei", {6: ("Risc de piață", 40000000)})
    wb.save(folder / "registru_2024.xlsx")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    if hidden_copy:
        old = _sheet(wb, "Capital_vechi", "Valori în lei", {5: ("Fonduri proprii eligibile SCR, nivel 1", 200000000)})
        old.sheet_state = "hidden"
    cap = _sheet(wb, "Capital", capital_unit, {
        5: ("Fonduri proprii eligibile SCR, nivel 1", c["t1"]), 6: ("Fonduri proprii eligibile SCR, nivel 2", c["t2"]),
        7: ("Fonduri proprii eligibile SCR, nivel 3", c["t3"]), 8: ("Fonduri proprii eligibile SCR, total", "=C5+C6+C7"),
        10: ("SCR", c["scr"]), 11: ("MCR", c["mcr"]), 12: ("Rata de solvabilitate", c["ratio"])}, prior={10: 110000})
    cap["C12"].number_format = "0.0%"
    if mcr_note:
        cap["E11"] = mcr_note
    _sheet(wb, "SCR_module", "Valori în mii lei", {
        5: ("Risc de subscriere asigurare generală", m["uw"]), 6: ("Risc de piață", m["mkt"]),
        7: ("Risc de contrapartidă", m["cpty"]), 8: ("Efect de diversificare (se scade)", m["div"]),
        9: ("SCR de bază", m["basic"]), 10: ("Risc operațional", m["op"]), 11: ("SCR", m["scr"])})
    wb.save(folder / "registru_2025.xlsx")

    (folder / "mapping_2024.toml").write_text("\n".join(
        f'[[figure]]\ntext = "{t}"\noccurrence = 1\nsheet = "Capital"\ncell = "{cell}"\nformat = "{fmt}"\n'
        for t, cell, fmt in [("200,0", "C5", "mil_lei_1"), ("20,0", "C6", "mil_lei_1"), ("220,0", "C7", "mil_lei_1"),
                             ("110,0", "C9", "mil_lei_1"), ("40,0", "C10", "mil_lei_1")]), encoding="utf-8")
    (folder / "raport_2024.txt").write_text(
        "E.1 Fondurile proprii au fost de 220,0 mil. lei, nivel 1 200,0 mil. lei, nivel 2 20,0 mil. lei.\n"
        f"{FRAMEWORK}\nSCR 110,0 mil. lei. MCR 40,0 mil. lei. Rata 200,0 %.\n", encoding="utf-8")
    (folder / "raport_Asigurari_Dunarea_2024_extras.txt").write_text(
        "Poziția de solvabilitate a rămas robustă pe parcursul întregului exercițiu financiar.\n", encoding="utf-8")
    (folder / "situatii_financiare.txt").write_text(
        "Indicator (lei) 2025 2024\n"
        f"Dividende plătite (fluxuri de numerar din activitatea de finanțare) {fs_dividends} (12.000.000)\n", encoding="utf-8")
    (folder / "proces_verbal_AGA_2025-04-24.txt").write_text(
        "Adunarea Generală Ordinară a Acționarilor aprobă, cu unanimitatea voturilor exprimate, distribuirea de "
        f"dividende în valoare totală de {aga_amount} lei din profitul net al exercițiului financiar 2024.\n", encoding="utf-8")
    d = docx.Document()
    for p in (NOTE if note is None else note):
        d.add_paragraph(p)
    d.save(folder / "nota_functia_actuariala_2026-02-02.docx")


def config(**overrides):
    cfg = copy.deepcopy(CONFIG)
    published = {"eof_t1": "200,0", "eof_t2": "20,0", "eof_total": "220,0", "scr": "110,0", "mcr": "40,0",
                 "solvency_ratio": "200,0"}  # the synthetic 2024 report's figures
    for f in cfg["fact"]:
        if f["id"] == "dividends_paid":
            f["file"] = "situatii_financiare.txt"
        if f["id"] in published:
            f["prior_mapping"] = {"text": published[f["id"]], "occurrence": 1}
    cfg.update(overrides)
    return cfg


@pytest.fixture
def make_run(tmp_path):
    from sfcr.engine import Run
    from sfcr.render import render

    def _make(cfg=None, **kw):
        build(tmp_path, **kw)
        run = Run(cfg or config(), tmp_path).run()
        text, log = render(run)
        return run, text, log

    return _make
