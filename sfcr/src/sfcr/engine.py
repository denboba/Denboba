"""Resolve every configured fact, run the cross-checks, apply pinned decisions and verify
narrative claims. Nothing here formats report text except via numbers.fmt_*."""

from __future__ import annotations

import re
import tomllib
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path

from . import documents
from .numbers import dec, fmt_fact, fmt_number_only, to_decimal
from .workbook import Located, Workbook, norm

OK = ("verified", "derived")


@dataclass
class Fact:
    id: str
    kind: str
    name_ro: str
    status: str = "missing"
    value: Decimal | None = None
    source: dict = field(default_factory=dict)
    why: list[str] = field(default_factory=list)
    reason_ro: str = ""
    prior: Decimal | None = None
    prior_source: dict = field(default_factory=dict)
    prior_status: str = "missing"
    checks: list[dict] = field(default_factory=list)
    decision: str | None = None

    @property
    def usable(self) -> bool:
        return self.status in OK and self.value is not None

    @property
    def display(self) -> str | None:
        return fmt_fact(self.value, self.kind) if self.usable else None

    @property
    def prior_display(self) -> str | None:
        return fmt_fact(self.prior, self.kind) if self.prior is not None and self.prior_status in OK else None


@dataclass
class Issue:
    code: str
    severity: str  # blocking | review | warning | info
    message: str
    facts: list[str] = field(default_factory=list)


class Run:
    def __init__(self, cfg: dict, base: Path):
        self.cfg, self.base = cfg, base
        self.year, self.prior_year = cfg["report"]["year"], cfg["report"]["prior_year"]
        lay = cfg["workbook_layout"]
        self.units = lay["units"]
        self.wb = Workbook(base / cfg["inputs"]["workbook"], self.units, lay["unit_search_rows"], lay["unit_note_prefix"])
        self.prior_wb = Workbook(base / cfg["inputs"]["prior_workbook"], self.units, lay["unit_search_rows"],
                                 lay["unit_note_prefix"])
        self.facts: dict[str, Fact] = {}
        self.issues: list[Issue] = []
        self.check_results: list[dict] = []
        self.claims: dict[str, dict] = {}
        self.mapping_audit: list[dict] = []
        self.used_files: set[str] = {self.wb.name, self.prior_wb.name, cfg["inputs"]["other_company_report"]}

    def fmt_year(self, text: str) -> str:
        return text.replace("{year}", str(self.year)).replace("{prior_year}", str(self.prior_year))

    def issue(self, code, severity, message, facts=()):
        self.issues.append(Issue(code, severity, message, list(facts)))

    # ---------------------------------------------------------------- facts ----------
    def _from_located(self, f: Fact, loc: Located, wb: Workbook) -> tuple[str, Decimal | None, dict, list[str], str]:
        src = {"file": loc.file, "sheet": loc.sheet, "cell": loc.cell, "label": loc.label,
               "unit": loc.unit, "unit_note": loc.unit_src, "formula": loc.formula, "raw": repr(loc.raw)}
        why = list(loc.why)
        if loc.status != "verified":
            return loc.status, None, src, why, loc.reason_ro
        if f.kind == "amount":
            if loc.factor is None:
                why.append(f"no unit note found on sheet {loc.sheet}")
                return "uncertain", None, src, why, f"unitatea de măsură a foii {loc.sheet} din {wb.name} nu este indicată"
            why.append(f"unit from {loc.unit_src} -> x{loc.factor} lei")
            return "verified", loc.value * loc.factor, src, why, ""
        if f.kind == "ratio":
            if "%" not in (loc.number_format or ""):
                why.append(f"number format {loc.number_format!r} is not a percentage: scale unknown")
                return "uncertain", None, src, why, f"{wb.name}, {loc.sheet}!{loc.cell}: scara procentului este nesigură"
            why.append(f"percentage format {loc.number_format!r}: stored as a fraction")
            return "verified", loc.value, src, why, ""
        return "verified", loc.value, src, why, ""  # count

    def resolve_facts(self):
        lay = self.cfg["workbook_layout"]
        for spec in self.cfg["fact"]:
            f = Fact(spec["id"], spec["kind"], spec.get("name_ro", spec["id"]))
            self.facts[f.id] = f
            if spec.get("source", "workbook") == "workbook":
                self.used_files.add(self.wb.name)
                loc = self.wb.locate(spec["sheet"], spec["label"], lay["current_column"])
                f.status, f.value, f.source, f.why, f.reason_ro = self._from_located(f, loc, self.wb)
                if f.status != "verified" and not f.reason_ro:
                    f.reason_ro = f"{f.name_ro} nu poate fi citit din {self.wb.name}"
                if spec.get("prior"):
                    self._resolve_prior(f, spec)
            else:
                self._resolve_document(f, spec)

    def _resolve_prior(self, f: Fact, spec: dict):
        lay = self.cfg["workbook_layout"]
        cands = []
        loc = self.prior_wb.locate(spec["sheet"], spec["label"], lay["current_column"])
        st, val, src, why, _ = self._from_located(f, loc, self.prior_wb)
        if st == "verified":
            cands.append((val, src, why))
        loc2 = self.wb.locate(spec["sheet"], spec["label"], lay["prior_column"])
        st2, val2, src2, why2, _ = self._from_located(f, loc2, self.wb)
        if st2 == "verified":
            cands.append((val2, src2, why2))
        if not cands:
            f.prior_status = "missing"
            f.why.append("prior-year value not found in either workbook")
            return
        f.prior, f.prior_source, f.prior_status = cands[0][0], cands[0][1], "verified"
        f.why.append(f"prior year from {cands[0][1]['file']} {cands[0][1]['sheet']}!{cands[0][1]['cell']}")
        if len(cands) == 2:
            agree = cands[0][0] == cands[1][0]
            self._check(f"PRIOR_AGREE:{f.id}", f"{f.id} prior: {self.prior_wb.name} vs {self.wb.name} '{lay['prior_column']}'",
                        agree, f"{cands[0][0]} vs {cands[1][0]}", [f.id])
            if not agree:
                f.prior_status = "conflict"
                self.issue("SOURCE_CONFLICT", "review", f"prior-year {f.id}: {cands[0][0]} ({cands[0][1]['file']}) vs "
                           f"{cands[1][0]} ({cands[1][1]['file']} {cands[1][1]['cell']})", [f.id])
        m = spec.get("prior_mapping")
        if m:
            shown = fmt_number_only(f.prior, f.kind)
            agree = shown == m["text"]
            self._check(f"PRIOR_PUBLISHED:{f.id}", f"{f.id} prior formats to the published figure {m['text']!r}",
                        agree, f"{shown} vs {m['text']}", [f.id])
            if not agree:
                f.prior_status = "conflict"
                self.issue("SOURCE_CONFLICT", "review", f"prior-year {f.id} {shown} differs from the published {m['text']}", [f.id])

    def _resolve_document(self, f: Fact, spec: dict):
        path = self.base / spec["file"]
        self.used_files.add(path.name)
        column = self.fmt_year(spec.get("column", ""))
        if spec["method"] == "table_row":
            fig = documents.table_row_figure(path, spec["row_label"], column, self.units)
        else:
            fig = documents.after_text_figure(path, spec["anchor"], self.units)
        f.status, f.why, f.reason_ro = fig.status, list(fig.why), fig.reason_ro
        f.source = {"file": fig.file, "page": fig.page, "excerpt": fig.excerpt}
        if fig.status != "verified":
            return
        f.value = abs(fig.value) if spec.get("absolute") else fig.value
        if spec.get("absolute") and fig.value < 0:
            f.why.append("value shown as an outflow (parentheses); the absolute amount is reported, per config")
        c = spec.get("corroborate")
        if c:
            cpath = self.base / c["file"]
            self.used_files.add(cpath.name)
            if c["method"] == "after_text":
                other = documents.after_text_figure(cpath, c["anchor"], self.units)
            else:
                other = documents.table_row_figure(cpath, c["row_label"], self.fmt_year(c.get("column", "")), self.units)
            agree = other.status == "verified" and abs(other.value) == f.value
            self._check(f"CORROBORATE:{f.id}", f"{f.id}: {fig.file} vs {other.file}", agree,
                        f"{f.value} vs {other.value}", [f.id])
            f.why += other.why
            if not agree:
                f.status = "conflict"
                f.reason_ro = f"{f.name_ro}: {fig.file} și {other.file} nu concordă"
                self.issue("SOURCE_CONFLICT", "blocking", f"{f.id}: {fig.file} {f.value} vs {other.file} {other.value}", [f.id])

    # ---------------------------------------------------------------- checks ---------
    def _check(self, cid, what, passed, detail, facts):
        res = {"id": cid, "what": what, "result": "pass" if passed else "FAIL", "detail": detail}
        self.check_results.append(res)
        for fid in facts:
            if fid in self.facts:
                self.facts[fid].checks.append({"id": cid, "result": res["result"]})
        return passed

    def _block(self, fids, reason_ro):
        for fid in fids:
            f = self.facts[fid]
            if f.status in OK:
                f.status = "uncertain"
                f.reason_ro = reason_ro

    def run_checks(self):
        tol = self.cfg["tolerances"]
        sum_tol = dec(tol["sum_lei"])
        for chk in self.cfg.get("check", []):
            names = [t.lstrip("+-") for t in chk.get("terms", [])] + [chk.get("equals") or chk.get("less"), chk.get("than")]
            names = [n for n in names if n]
            missing = [n for n in names if not self.facts[n].usable]
            if missing:
                self.check_results.append({"id": chk["id"], "what": chk["what"], "result": "not run",
                                           "detail": f"inputs not available: {missing}"})
                if chk.get("block_if_not_run"):
                    self._block(chk["on_fail"], f"verificarea {chk['id']} nu poate fi efectuată: lipsesc {', '.join(missing)}")
                    self.issue("CHECK_NOT_RUN", "blocking", f"{chk['id']} not run ({missing} unavailable); {chk['on_fail']} withheld", names)
                else:
                    self.issue("CHECK_NOT_RUN", "review", f"{chk['id']} not run: {missing} unavailable", names)
                continue
            if "terms" in chk:
                total = sum((-self.facts[t[1:]].value if t[0] == "-" else self.facts[t.lstrip('+')].value)
                            for t in chk["terms"])
                target = self.facts[chk["equals"]].value
                ok = abs(total - target) <= sum_tol
                self._check(chk["id"], chk["what"], ok, f"{total} vs {target} lei", names)
            else:
                a, b = self.facts[chk["less"]].value, self.facts[chk["than"]].value
                ok = a < b
                self._check(chk["id"], chk["what"], ok, f"{a} vs {b} lei", names)
            if not ok:
                self._block(chk["on_fail"], f"verificarea {chk['id']} a eșuat ({chk['what']}) în {self.wb.name}")
                self.issue("CALCULATION_MISMATCH", "blocking", f"{chk['id']} failed: {chk['what']}", chk["on_fail"])

        lo, hi = dec(tol["plausible_min"]), dec(tol["plausible_max"])
        for f in self.facts.values():
            if f.value is None or f.prior is None or f.prior_status not in OK:
                continue
            if f.value == f.prior:
                self._check(f"STALE:{f.id}", f"{f.id} differs from the prior year", False,
                            f"identical to prior {f.prior}", [f.id])
                self.issue("POSSIBLE_CARRY_FORWARD", "review", f"{f.id} is identical to the prior year ({f.prior})", [f.id])
            if f.prior != 0 and f.usable:
                r = f.value / f.prior
                ok = lo <= abs(r) <= hi
                self._check(f"PLAUSIBLE:{f.id}", f"{f.id} current/prior within [{lo}, {hi}]", ok, f"ratio {r:.4f}", [f.id])
                if not ok:
                    self._block([f.id], f"{f.name_ro}: variație neobișnuită față de {self.prior_year} "
                                f"(posibilă eroare de unitate în {self.wb.name})")
                    self.issue("UNIT_MISMATCH", "blocking", f"{f.id} moved x{r:.4f} vs prior: possible unit error", [f.id])

    def reconcile_computed(self):
        """Facts with a 'computed' definition: workbook value vs the calculation."""
        tol = dec(self.cfg["tolerances"]["ratio"])
        for spec in self.cfg["fact"]:
            comp = spec.get("computed")
            if not comp:
                continue
            f = self.facts[spec["id"]]
            num, den = self.facts[comp["numerator"]], self.facts[comp["denominator"]]
            calc = num.value / den.value if num.usable and den.usable and den.value != 0 else None
            if calc is None:
                why = f"{comp['numerator']} / {comp['denominator']} not computable"
                self.check_results.append({"id": f"RATIO:{f.id}", "what": f"{f.id} = {comp['numerator']} / {comp['denominator']}",
                                           "result": "not run", "detail": why})
                if f.usable:  # a workbook ratio we cannot confirm does not reach the text
                    f.status = "uncertain"
                    f.reason_ro = f"{f.name_ro} nu poate fi verificată: {num.name_ro} sau {den.name_ro} indisponibil"
                continue
            calc_shown = fmt_fact(calc, f.kind)
            if not f.usable:
                if f.status == "missing" or f.status == "uncertain":
                    f.why.append(f"workbook value unavailable; derived as {comp['numerator']} / {comp['denominator']} = {calc}")
                    f.status, f.value, f.reason_ro = "derived", calc, ""
                    f.source = {**f.source, "derived_from": [comp["numerator"], comp["denominator"]]}
                continue
            agree = abs(f.value - calc) <= tol and fmt_fact(f.value, f.kind) == calc_shown
            self._check(f"RATIO:{f.id}", f"{f.id} = {comp['numerator']} / {comp['denominator']}", agree,
                        f"workbook {f.value} vs computed {calc}", [f.id])
            if agree:
                continue
            dec_ = next((d for d in self.cfg.get("decision", []) if d["fact"] == f.id
                         and dec(d["when_workbook_value"]) == f.value and d["when_computed"] == calc_shown), None)
            if dec_:
                f.why.append(f"conflict workbook {f.value} vs computed {calc}; {dec_['id']} applies: choose {dec_['choose']}")
                f.decision = dec_["id"]
                if dec_["choose"] == "computed":
                    f.status, f.value = "derived", calc
                    f.source = {**f.source, "derived_from": [comp["numerator"], comp["denominator"]]}
                self.issue("DECISION_APPLIED", "review", f"{f.id}: {dec_['id']} ({dec_['decided_by']})", [f.id])
            else:
                f.status = "conflict"
                f.reason_ro = (f"{f.name_ro} din {f.source.get('file')} {f.source.get('sheet')}!{f.source.get('cell')} "
                               f"({fmt_fact(f.value, f.kind)}) diferă de {num.name_ro} / {den.name_ro} ({calc_shown})")
                self.issue("SOURCE_CONFLICT", "blocking", f"{f.id}: workbook {f.value} vs computed {calc}; no decision applies", [f.id])

    # ---------------------------------------------------------------- claims ---------
    def operand(self, token: str):
        if re.fullmatch(r"-?\d+(\.\d+)?", token):
            return dec(token), None
        fid, _, which = token.partition("@")
        f = self.facts[fid]
        if which == "prior":
            if f.prior is None or f.prior_status not in OK:
                return None, f"{f.name_ro} pentru {self.prior_year} indisponibil"
            return f.prior, None
        if not f.usable:
            return None, f.reason_ro or f"{f.name_ro} indisponibil"
        return f.value, None

    def condition(self, cond) -> tuple[bool | None, str | None]:
        left, op, right = cond
        a, ra = self.operand(left)
        b, rb = self.operand(right)
        if a is None or b is None:
            return None, ra or rb
        return {"<": a < b, ">": a > b, "==": a == b, "<=": a <= b, ">=": a >= b}[op], None

    def verify_claims(self):
        for c in self.cfg.get("claim", []):
            path = self.base / c["file"]
            self.used_files.add(path.name)
            page = documents.find_quote(path, c["quote"]) if path.exists() else None
            conds = [self.condition(x) for x in c.get("conditions", [])]
            ok = page is not None and all(r is True for r, _ in conds)
            detail = f"quote {'found on page ' + str(page) if page else 'NOT found'} in {c['file']}; conditions {[r for r, _ in conds]}"
            self.claims[c["id"]] = {**c, "supported": ok, "page": page, "detail": detail}
            self.check_results.append({"id": f"CLAIM:{c['id']}", "what": "quote present and conditions hold",
                                       "result": "pass" if ok else "FAIL", "detail": detail})
            if page is None:
                self.issue("UNSUPPORTED_CLAIM", "review", f"{c['id']}: quote not found in {c['file']}")
            if c.get("review") and ok:
                self.issue("REVIEW_REQUIRED", "review", f"{c['id']}: {c['review']}")

    # ---------------------------------------------------------------- mapping audit --
    def audit_mapping(self):
        """Re-check last year's mapping against last year's workbook and report."""
        cfgi = self.cfg["inputs"]
        mpath, rpath = self.base / cfgi["prior_mapping"], self.base / cfgi["prior_report"]
        self.used_files |= {mpath.name, rpath.name}
        entries = tomllib.loads(mpath.read_text(encoding="utf-8")).get("figure", [])
        report = rpath.read_text(encoding="utf-8")
        covered = set()
        for e in entries:
            pattern = rf"(?<![\d.,]){re.escape(e['text'])}(?![\d]|,\d)"
            hits = len(re.findall(pattern, report))
            row = {"text": e["text"], "occurrence": e["occurrence"], "cell": f"{e['sheet']}!{e['cell']}", "problems": []}
            if hits < e["occurrence"]:
                row["problems"].append(f"occurrence {e['occurrence']} not in report (found {hits} on token boundaries)")
            covered.add((e["text"], e["occurrence"]))
            try:
                ws = self.prior_wb.wb_v[e["sheet"]]
                raw = ws[e["cell"]].value
                _, factor, _ = self.prior_wb.sheet_unit(e["sheet"])
                if raw is None:
                    row["problems"].append("mapped cell is empty")
                else:
                    v = to_decimal(raw)
                    fmt = e["format"]
                    if e.get("derived") == "share":
                        v = v / to_decimal(ws[e["total_cell"]].value)
                        shown = fmt_number_only(v, "ratio")
                    elif fmt == "pct_1":
                        shown = fmt_number_only(v, "ratio")
                    elif fmt == "count":
                        shown = fmt_number_only(v, "count")
                    else:
                        shown = fmt_number_only(v * (factor or 1), "amount_abs" if fmt.endswith("_abs") else "amount")
                    row["workbook_shows"] = shown
                    if shown != e["text"]:
                        label = next((c.value for c in ws[ws[e["cell"]].row] if isinstance(c.value, str)), "")
                        row["problems"].append(f"cell formats to {shown!r}, not {e['text']!r} (row label {label!r})")
            except (KeyError, TypeError) as ex:
                row["problems"].append(f"not in workbook: {ex}")
            self.mapping_audit.append(row)
        # Decimal figures in last year's text that no mapping entry covers.
        seen: dict[str, int] = {}
        for m in re.finditer(r"(?<![\d.,])\d{1,3}(?:\.\d{3})*,\d+(?![\d])", report):
            seen[m.group()] = seen.get(m.group(), 0) + 1
            if (m.group(), seen[m.group()]) not in covered:
                line = report[:m.start()].count("\n") + 1
                self.mapping_audit.append({"text": m.group(), "occurrence": seen[m.group()], "cell": None,
                                           "problems": [f"figure on line {line} of {rpath.name} has no mapping entry"]})
        bad = [r for r in self.mapping_audit if r["problems"]]
        if bad:
            self.issue("MAPPING_DEFECT", "info", f"{len(bad)} problem(s) in {mpath.name}; see the mapping audit")

    def run(self):
        self.resolve_facts()
        self.run_checks()
        self.reconcile_computed()
        self.verify_claims()
        self.audit_mapping()
        hidden = self.wb.hidden_sheets()
        if hidden:
            self.issue("HIDDEN_SHEET", "info", f"{self.wb.name}: hidden sheet(s) {hidden} were not read")
        return self
