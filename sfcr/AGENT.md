# AGENT.md: SFCR roll-forward, section E slice

This file tells agents and contributors how to work on this project. It is derived from the
assessment brief (`README.md` of the BlueberryX Engineering Assessment, below: **the brief**) and
from a review of every file in `date/`. Where this file and the brief disagree, **the brief
wins**; record the disagreement in §10.

Read §1 to §4 before you change anything. §10 lists the known gaps and inconsistencies; do not
resolve one by guessing.

---

## 1. What we are building (and what we are not)

The client is **Asigurări Carpați S.A.** (fictional). The goal is a roll-forward of the 2024
SFCR into a 2025 draft. The brief asks for three deliverables:

| # | Deliverable | Path | Who writes it | Acceptance criteria (from the brief) |
|---|---|---|---|---|
| 1 | The case | `DESIGN.md` | **the candidate** | ≤ 4 pages, ≥ 1 diagram, answers points 1–9 of brief §5, answer first |
| 2 | Working slice, **section E only** | `src/`, `config/`, `output/`, `tests/`, `README.md` | candidate + agents | see §5 below |
| 3 | AI usage | `ai/AI.md`, `ai/conversations/` | **the candidate** | 1 page, 4 questions; raw, complete, **unedited** exports |

Submission: one zip named `firstname_lastname.zip` with `README.md`, `DESIGN.md`, `ai/`,
`config/`, `output/`, `src/`, `tests/`, and the git history if git is used.

**Out of scope for the code:** sections A–D, the summary, DOCX rendering, OCR, multi-client
deployment, any LLM call from the program. Findings from those sections go in §10 and in
`DESIGN.md`, not in code. The brief says "Keep it small".

**Agents must not author `DESIGN.md` or `ai/AI.md` content, or touch `ai/conversations/`.**
Those files are the candidate's own judgement and raw record. Agents may create empty
skeletons with the required headings.

---

## 2. Setup

Requirements: Python **≥ 3.11** (the program uses `tomllib`). No system binaries are needed.

```bash
cd sfcr
python3 -m venv .venv && . .venv/bin/activate      # optional
pip install -r requirements.txt                    # runtime
pip install -r requirements-dev.txt                # + pytest
# put the client materials in sfcr/date/ (they are not committed, see §10 R1)
cd date && python3 ../src/run_section_e.py         # THE one command
cd .. && python3 -m pytest -q                      # tests
```

Pinned runtime dependencies (`requirements.txt`):

| Package | Version | Used for |
|---|---|---|
| `openpyxl` | 3.1.5 | workbook values, formulas, number formats, sheet visibility |
| `pdfplumber` | 0.11.10 | text lines of `situatii_financiare_2025_extras.pdf` (pypdf splits table rows) |
| `python-docx` | 1.2.0 | text of the actuarial note, for quote verification |

Dev: `pytest`. Do not add dependencies without a reason written in the commit message.

Never open and save the original workbooks in Excel or LibreOffice (brief §3): it can rewrite
cached values and formats. The program opens them read-only and records their SHA-256.

---

## 3. Repository layout

```text
sfcr/
├── AGENT.md            this file
├── README.md           install, the one command, outputs (user-facing)
├── DESIGN.md           deliverable 1 (candidate)
├── ai/AI.md            deliverable 3 (candidate)
├── ai/conversations/   raw AI exports (candidate, unedited)
├── config/section_e.toml   ALL year-, client- and layout-specific decisions
├── src/run_section_e.py    entry point (run from date/)
├── src/sfcr/           numbers.py · workbook.py · documents.py · engine.py · render.py · cli.py
├── tests/              synthetic fixtures only; one optional test on real data
├── date/               client materials, read-only, git-ignored
└── output/             generated only, git-ignored
```

The layers are kept apart: `numbers` (pure Decimal parsing and formatting), `workbook` and
`documents` (extraction, which never formats), `engine` (fact resolution, checks, decisions),
`render` (text built only from resolved facts), `cli` (I/O and outputs).

---

## 4. Input data: what each file is and what to expect

All files live in `date/`. Their hashes are recorded in `output/manifest.json` on every run.

| File | Used by the program for section E | What a reviewer found (verified on the 2026-10 copy) |
|---|---|---|
| `registru_2025.xlsx` | **primary source** for own funds, SCR, MCR, ratio | See the trap list below. |
| `registru_2024.xlsx` | prior-year values (comparatives, direction words) | Values in **lei**. Capital rows: T1 C5, T2 C6, total C7, SCR C9 (D9 = 2023), MCR C10, ratio C11. |
| `mapping_2024.toml` | lineage check (`prior_mapping` on a fact): 2024 text → 2024 cell → label | Has no E.1 dividend figure (`12,0`). The `12,6` entry points to `Investitii!C7` (a share, 0.784) when it should be `C6`. |
| `raport_2024.txt` | the structure of section E; static regulatory sentence | E.2 says the SCR rose "ca urmare a creșterii volumului de prime". That cause must not carry over. |
| `situatii_financiare_2025_extras.pdf` | dividends paid in 2025 (cash-flow line) | `(15.000.000)` in lei, 2025 column, page 1. Text layer present. |
| `proces_verbal_AGA_2025-04-24.txt` | corroborates dividends | 15.000.000 lei approved from 2024 profit, payable by 30.06.2025. The policy approved by the board in 2023 is unchanged. |
| `nota_functia_actuariala_2026-02-02.docx` | causal claim for the SCR; "formula standard" | SCR fell, mainly from market risk (listed equities sold in Q3). The reinsurance programme was renewed **on the same terms**. The MCR is still being calculated. |
| `email_director_financiar.txt` | nothing; its requests are logged as refused decisions | Asks to fill gaps with 2024 values, to credit reinsurance for the ratio, and to copy Dunărea wording. **All three are forbidden by the brief's rules 1, 2 and 4.** |
| `raport_Asigurari_Dunarea_2024_extras.txt` | similarity screen only (must share no wording) | Another company. Never a source of facts or text. |
| `proces_verbal_CA_2025-02-18_scanat.pdf` | nothing (section B) | Image only, no text layer. A sixth board member, effective on ASF approval. |
| `raport_2024_sinteza.txt` | nothing (summary out of scope) | — |

**`registru_2025.xlsx` traps the program must handle, not paper over:**

1. Sheet `Capital` gained a "nivel 3" row, so **every row below row 6 moved down by one**. A
   fixed-cell mapping from 2024 reads the wrong row. Locate rows by **label**, never by address.
2. `Capital`, `SCR_module`, `Provizioane` and `Daune` are now in **mii lei** (row 2
   "Valori în mii lei"); in 2024 they were in lei. Read the unit from the sheet on every run.
3. Hidden sheet `Capital_vechi` ("NU SE FOLOSEȘTE (copie 2024)") holds 2024 values. **Hidden
   sheets are never read.**
4. `Capital!C8` (total own funds) is a formula `=C5+C6+C7`. Another copy of the file may carry
   no cached value. The program then evaluates the simple formula itself and says so; it never
   treats a blank formula result as zero.
5. `Capital!C11` (MCR) is empty, with the note "în lucru, actuariat" in E11. The actuarial note
   confirms the MCR is pending. Result: a **visible gap**, never the 2024 value of 45,6.
6. `Capital!C12` (ratio) holds **1.874**, the 2024 ratio exactly, as a typed constant. But
   C8/C10 = 293.699,3 / 148.763,2 = **1,97427 → 197,4 %**. This is a **conflict**: the program
   must report it and may resolve it only through a pinned decision in config (§7).
7. `SCR_module!C8` diversification is **positive**, labelled "(se scade)". In 2024 it was
   negative, "(valoare negativă)". The sign convention comes from config, and the sum check
   catches any mismatch.

---

## 5. The program contract (deliverable 2)

**Command:** `cd date && python3 ../src/run_section_e.py [--config PATH]`. Inputs resolve
relative to the current folder (`date/`). Outputs go to `../output/`, configurable in
`[output] dir`. It must run **unchanged** on another version of `registru_2025.xlsx`.

**Outputs** (every file in `output/` is written by the program; nothing is hand-edited):

| File | Content |
|---|---|
| `sectiunea_E_<year>.txt` | Section E draft in Romanian, following the 2024 structure (E.1, E.2), with `[de completat: …]` gaps |
| `figuri_E_<year>.json` | Evidence ledger: for every fact, its exact value in lei, unit, status, source (file/sheet/cell/label/page), **why** that cell is right, and its checks and decisions |
| `verificari_E_<year>.md` | Human review report: figure table, check results, conflicts, decisions applied, gaps, refused requests, the 2024 mapping audit, the similarity screen |
| `manifest.json` | SHA-256 of every input and the config, library versions. No timestamp, so identical inputs give byte-identical outputs |

**Exit codes:** `0` when outputs were written, even with gaps (the draft says so); `2` for a
fatal error (config invalid, workbook unreadable), with a message naming the file and the problem.

**Fact statuses** (each fact gets exactly one): `verified`, `derived` (a calculation from
verified facts), `missing`, `uncertain` (failed a check, ambiguous location, unknown unit),
`conflict` (two sources disagree and no pinned decision applies). Only `verified` and `derived`
values may reach the text.

**How a workbook fact is found** (this is the "why" recorded for each figure):

1. Use the sheet named in config. If it is missing, or hidden, the fact is `missing` or
   `uncertain`.
2. Find exactly one cell whose normalized text equals the configured row label. Normalization
   is NFC, cedilla ş/ţ → comma ș/ț, case-fold and collapsed spaces. Zero matches means
   `missing`, and the closest labels are listed as hints but never used. Two or more matches
   mean `uncertain`.
3. Find exactly one header cell `31.12 an curent` (from config); its column gives the value.
4. Read the unit from the sheet's "Valori în …" note, matching the longest phrase first. If
   there is no unit note, or it can't be read, the fact is `uncertain`.
5. Read the value: a number, a simple formula (`+`/`-`/`SUM`: re-evaluated and compared with the
   cache, or evaluated when there is no cache; anything else: Excel's cached result, else a gap),
   or Romanian text that parses unambiguously. A blank is `missing`, and any note elsewhere in
   the row is quoted. Booleans, errors and unparseable text are `uncertain`.
6. Lineage: for a fact linked to `mapping_2024.toml`, check that the 2024 mapped cell carries
   the same label and formats to the 2024 report text.

---

## 6. Checks run before anything reaches the text

All tolerances and terms live in `config/section_e.toml`. Money uses `Decimal` only, built
from strings, never from binary-float arithmetic.

| ID | Check | On failure |
|---|---|---|
| `EOF_TIERS` | T1 + T2 + T3 = total own funds (source values, unrounded) | all four `uncertain` |
| `SCR_BASIC` | underwriting + market + counterparty − diversification = basic SCR | module facts `uncertain` |
| `SCR_TOTAL` | basic SCR + operational = SCR (SCR_module) | module facts `uncertain` |
| `SCR_CROSS` | `Capital` SCR = `SCR_module` SCR | SCR `conflict` |
| `RATIO` | workbook ratio = total own funds / SCR within tolerance | ratio `conflict` (unless a pinned decision applies) |
| `MCR_LT_SCR` | MCR < SCR | MCR `uncertain` |
| `PRIOR_AGREE` | prior value in the 2025 workbook ("an anterior") = the 2024 workbook | prior `conflict`, so no direction sentence |
| `PRIOR_PUBLISHED` | the 2024 workbook value formats to the figure published in the 2024 report (`prior_mapping`) | prior `conflict` |
| `PLAUSIBLE` | current / prior between 0,5 and 2,0, which catches lei vs mii lei | fact `uncertain` |
| `STALE` | current value identical to prior (possible carry-forward) | warning in the report |
| `CORROBORATE` | dividends in the audited statements = AGA amount | dividends `conflict` |
| `TRACE` | every number in the draft (outside gaps) equals a ledger value or a literal in the approved config template | run reports a blocking issue |
| `SIMILARITY` | the draft shares no 6-word sequence with the other insurer's report | blocking issue |
| `MAPPING_AUDIT` | each 2024 mapping entry formats to its 2024 text, and its occurrence exists on token boundaries | informational (lists errors such as `12,6`) |

---

## 7. Report rules, turned into guardrails

| Brief rule | Guardrail in this project |
|---|---|
| 1 Traceability; never last year's figure for this year's | Text is built only from ledger facts. There is no fallback to the prior year anywhere. Hidden sheets and the "an anterior" column never feed a current-year fact. |
| 2 No unsupported statements | Direction words ("în creștere"/"în scădere") are picked by comparing **unrounded** values. A causal clause appears only if it has a `[[claim]]` in config whose **quote is found verbatim** in the cited company document **and** whose conditions hold on this run's data. |
| 3 Visible gaps | A sentence that can't be filled becomes `[de completat: <concrete reason naming file/sheet/row>]`. |
| 4 No copied text | The `SIMILARITY` check runs against `raport_Asigurari_Dunarea_2024_extras.txt`. Templates are written fresh. |
| 5 Number format | `284,6 mil. lei` · `187,4 %` · `18.420`. ROUND_HALF_UP, once, on the unrounded value. A total comes from its source cell, never from adding rounded parts (e.g. 262,4 + 31,2 ≠ 293,7, and that is correct). |

**Human decisions** live in `config/section_e.toml`, each with `source`, `reason`,
`decided_by` and `date`:

- `[[decision]]` resolves a conflict. It is **pinned** to the values it was taken on
  (`when_workbook_value`, `when_computed`). On a different workbook the conflict reappears and
  is reported; the decision is not reused silently.
- `[[claim]]` is a supported narrative statement: source file, exact quote, data conditions.
- `[[refused_request]]` logs a request that was declined, with the rule it breaks. The CFO's
  three requests are logged here.

**Review points a person must sign off** (the program lists them in `verificari_*.md`):

1. Every decision in config, especially the ratio decision (C12 treated as stale).
2. Every gap, and who fills it (the MCR: the actuarial function).
3. The causal SCR clause taken from the actuarial note.
4. The static regulatory sentence carried from 2024 (compliance to confirm it is still accurate).
5. The data findings outside E in §10 (RCA premium, mapping error, provisions comparatives).

---

## 8. Tests

`python3 -m pytest -q` from `sfcr/`. Tests build **synthetic** workbooks with openpyxl. They
must not depend on `date/`; the single real-data test is skipped when `date/` is absent.
Required cases:

- formatting: half-up at x,x5; negatives; `18.420`; percentages from fractions.
- parsing: `284,6`, `18.420`, `1.234,56`, and `1,234` rejected as ambiguous.
- locator: a moved row is found by label; a duplicate label gives `uncertain`; a hidden sheet is
  ignored; a unit change from lei to mii lei is applied; a formula with no cache is evaluated.
- engine: tier-sum failure blocks the text; ratio conflict without a decision gives a gap; a
  pinned decision doesn't apply when the values differ; a missing MCR gives a gap, never the
  prior value; a direction word flips when the data flips; a claim is dropped when its quote
  is absent.
- end-to-end on `date/` when present: no `{`, no prior-year current figures, `TRACE` and
  `SIMILARITY` pass.

Never weaken or skip a test to get green. Report the commands you actually ran.

---

## 9. Working rules for agents

1. Before a change, re-run the command and the tests and read `output/verificari_*.md`.
2. Year-, client- and layout-specific values go in `config/`, never in `src/`. A grep for
   `2025`, `2024`, cell addresses (`C1[0-9]`) or sheet names in `src/` must come back empty.
3. Never modify `date/`. Never commit `date/` or `output/`.
4. Do not write narrative that the sources don't support. If you think a sentence is true but
   can't cite it, it is a gap.
5. Ask the user when a choice changes a published figure or a legal statement. Otherwise make
   a reversible choice and record it in §10.
6. Completion report: what changed, files, commands run with results, data findings, open
   decisions.

---

## 10. Known gaps and inconsistencies (do not resolve by assumption)

**Brief ↔ data**

- **D1.** The brief says the mapping covers "each figure in raport_2024.txt". It does not: the
  dividend `12,0` and the statutory technical provisions `214,6` (both from the financial
  statements), the employee count `412`, and the legal references are unmapped. The program's
  mapping audit lists them.
- **D2.** The mapping entry `12,6` points to `Investitii!C7` (0.784); the value is in `C6`
  (12.640.000). Outside E, but it proves that last year's mapping can't be trusted blindly.
- **D3.** The meaning of mapping `occurrence` is ambiguous: `12,4` also occurs inside `412,4`.
  We count on number-token boundaries.
- **D4.** Ratio in `Capital!C12` = 1.874 (stale) vs 197,4 % computed. Resolved by a pinned
  decision, **which the user must confirm**.
- **D5.** The MCR is missing for 2025, so the draft has a gap. "acoperită integral" can't be
  stated. The sources don't give MCR-eligible own funds, so if the MCR arrives we only claim
  coverage when tier 1 alone ≥ MCR.
- **D6.** Outside E: `Prime!C6` RCA = 142.017.800 in the workbook vs **124.017.800** in the audited
  statements (Note 4). The workbook lines sum to 456.912.400 ≠ the total of 438.912.400.
- **D7.** Outside E: the 2025 `Provizioane` "an anterior" column (184.100 / 9.800 / 193.900 mii lei)
  ≠ the published 2024 figures (182,3 / 9,8 / 192,1 mil. lei).
- **D8.** Outside E: listed equities are 0 in 2025 (`Investitii!C8`). The 2024 section C sentence
  "Expunerea pe riscul de piață provine în principal din acțiunile listate" is stale.
- **D9.** Outside E: 2024 reports headcount at 31 Dec (412); the 2025 statements give the *average*
  (437). These are different metrics, so the number can't simply be replaced.
- **D10.** Outside E: the board grows to six members only on ASF approval, and there is no
  evidence of that approval.
- **D11.** The 2024 E.1 says dividends were paid "conform politicii de dividende aprobate de
  Consiliul de Administrație". The 2025 sources support "approved by the AGA on 24.04.2025 from
  the 2024 profit", not that the payment followed the policy, so the wording changes.
- **D12.** The CFO email's requests 2, 3 and 4 conflict with the brief's rules 1, 2 and 4. The
  actuarial note directly contradicts the reinsurance explanation. They are logged as refused.

**Brief ambiguities (our reading)**

- **A1.** "run from `date/` … writes to `output/`": we write to `../output`, the sibling of `date/`,
  which matches the zip layout. This is configurable.
- **A2.** "another version of `registru_2025.xlsx`": the changes are unknown. We assume labels and
  the column header are the stable anchors; anything else unknown becomes a gap, not a guess.
- **A3.** The scope of section E follows the 2024 E.1/E.2: own funds, tiers, dividends,
  classification, SCR, MCR, ratio and method.

**Previous AGENT.md ↔ brief (why it was rewritten)**

- It listed no deliverables, page limit, zip name, `ai/` requirements, or the "section E only"
  scope.
- It proposed `setup.sh`/`run.sh`, DOCX rendering, OCR, regulatory profiles and multi-tenant
  isolation. The brief asks for one command run from `date/`, and to keep it small.
- "Do not overwrite outputs, retain run history" conflicts with "every file in `output/` must
  be produced by the program".
- It cited the data files as "planning names" only; they are now verified and described in §4.

**Repository**

- **R1.** This repository is the user's GitHub profile repo. The project lives in `sfcr/`.
  `date/` and `output/` are git-ignored so the assessment materials are not published. Copy the
  materials into `sfcr/date/` to run.
- **R2.** `DESIGN.md` and `ai/AI.md` are skeletons until the candidate writes them.
  `ai/conversations/` must hold the raw exports before submission.
