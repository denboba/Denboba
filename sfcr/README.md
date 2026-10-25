# SFCR roll-forward: section E (Asigurări Carpați S.A., 2025)

A small program that drafts section E of the 2025 Solvency and Financial Condition Report:
own funds, SCR, MCR and the solvency ratio. It starts from the 2024 report and this year's
workbook and documents. Every figure is traced to a cell or a document page, cross-checked
before it reaches the text, and left as a visible `[de completat: …]` gap when it can't be
filled correctly.

## Install

Python 3.11 or newer.

```bash
pip install -r requirements.txt          # openpyxl, pdfplumber, python-docx (pinned)
pip install -r requirements-dev.txt      # + pytest, for the tests
```

Put the client materials (the assessment's `date/` folder) in `date/`. They are not committed.

## Run: the one command

```bash
cd date && python3 ../src/run_section_e.py
```

It reads the inputs from the current folder and `../config/section_e.toml`, and writes
everything to `../output/`. Exit code 0 means the outputs were written; the review report
says whether the draft is ready. Exit code 2 means a configuration or input error, with a
message.

## Outputs (`output/`, all generated)

| File | What it is |
|---|---|
| `sectiunea_E_2025.txt` | the section E draft, in Romanian |
| `figuri_E_2025.json` | the evidence ledger: each figure's exact value, source cell/page, *why* that cell, its checks, decisions and the prior year |
| `verificari_E_2025.md` | the review report: figures, checks, issues, decisions to confirm, gaps, declined requests, an audit of last year's mapping |
| `manifest.json` | SHA-256 of every input and the config, library versions. Identical inputs give identical outputs |

## How it decides

- **Finding a figure.** Rows are found by **label** and columns by the header `31.12 an curent`,
  never by cell address. The unit is read from the sheet's "Valori în …" note, so rows that
  move or units that change are handled. Hidden sheets are never read.
- **Checks before text.** Tiers sum to the total; the SCR modules sum to the SCR; the SCR
  agrees on both sheets; ratio = own funds / SCR; MCR < SCR; prior-year values agree with the
  published 2024 report; dividends in the audited statements = the AGA decision; current/prior
  movements are plausible, which catches lei vs mii lei. A figure that fails a check stays out
  of the text.
- **Disagreements are never settled silently.** A conflict becomes a gap unless a
  `[[decision]]` in the config resolves it. A decision records source, reason, who and when,
  and is **pinned to the values it was taken on**: on a different workbook it doesn't apply,
  and the conflict is reported again.
- **Words follow the data.** "în creștere"/"în scădere" are chosen from unrounded values. A
  causal clause needs a `[[claim]]` whose quote is found verbatim in the company document and
  whose data conditions hold this year.
- **Number format.** `293,7 mil. lei`, `197,4 %`, `18.420`. Rounding is half up, once, on the
  exact `Decimal` value. Totals come from their source cell.

Everything that changes from year to year (file names, labels, headers, units, tolerances,
decisions, claims, sentence templates) is in `config/section_e.toml`.

## Tests

```bash
python3 -m pytest -q
```

The tests build synthetic workbooks: moved rows, unit changes, missing MCR, stale ratio,
pinned decisions, direction flips, missing quotes, failed sums, unit errors. One extra test
runs on `date/` when it is present.

## Known limits

- The formula evaluator handles only `+`/`-` of same-sheet cells and `SUM(range)`. Other
  formulas use Excel's cached result, and when there is none they become a gap.
- PDFs need a text layer. Scanned documents (e.g. the board minutes) are not read; they are
  not needed for section E.
- Section E only. See `AGENT.md` §10 for findings in sections A–D and for the open questions.
