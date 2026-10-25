"""Build section text from resolved facts only, then check what was built."""

from __future__ import annotations

import re

from .engine import Run
from .workbook import norm

PLACEHOLDER = re.compile(r"\{([a-z_0-9]+(?:@prior)?)\}")
GAP = re.compile(r"\[de completat:[^\]]*\]")
FIGURE = re.compile(r"(?<![\d.,])\d{1,3}(?:\.\d{3})*,\d+(?![\d])")


def _fill(run: Run, text: str) -> tuple[str | None, list[str], list[str]]:
    """Fill placeholders. Returns (text or None, reasons if unfillable, fact ids used)."""
    rep = run.cfg["report"]
    fixed = {
        "year": str(run.year),
        "prior_year": str(run.prior_year),
        "date": rep["date_text"].replace("{year}", str(run.year)),
        "prior_date": rep["date_text"].replace("{year}", str(run.prior_year)),
    }
    reasons, used = [], []

    def sub(m):
        key = m.group(1)
        if key in fixed:
            return fixed[key]
        fid, _, which = key.partition("@")
        f = run.facts[fid]
        used.append(key)
        shown = f.prior_display if which == "prior" else f.display
        if shown is None:
            reasons.append(f"{f.name_ro} pentru {run.prior_year} indisponibil" if which == "prior"
                           else (f.reason_ro or f"{f.name_ro} indisponibil"))
            return "?"
        return shown

    out = PLACEHOLDER.sub(sub, text)
    return (None if reasons else out), reasons, used


def render(run: Run) -> tuple[str, list[dict]]:
    """Returns the section text and a per-sentence log (variant used, facts, claims, gap)."""
    blocks: list[list[str]] = []  # paragraphs; headings are their own paragraph
    log = []
    for i, s in enumerate(run.cfg.get("sentence", [])):
        if s.get("heading"):
            blocks += [[s["heading"]], []]
        elif s.get("new_paragraph") and blocks and blocks[-1]:
            blocks.append([])
        variants = s.get("variant", [])
        if not variants:
            continue
        reasons: list[str] = []
        chosen = None
        for vi, v in enumerate(variants):
            conds = [run.condition(c) for c in v.get("when", [])]
            if any(r is None for r, _ in conds):
                reasons += [why for r, why in conds if r is None]
                continue
            if not all(r for r, _ in conds):
                continue
            bad_claims = [c for c in v.get("claims", []) if not run.claims[c]["supported"]]
            if bad_claims:
                reasons += [f"afirmația {c} nu este susținută de {run.claims[c]['file']}" for c in bad_claims]
                continue
            text, why, used = _fill(run, v["text"])
            if text is None:
                reasons += why
                continue
            chosen = (vi, text, used, v.get("claims", []), v["text"])
            break
        if not blocks:
            blocks.append([])
        if chosen:
            vi, text, used, claims, template = chosen
            blocks[-1].append(text)
            log.append({"sentence": i, "variant": vi, "text": text, "facts": used, "claims": claims, "template": template})
        else:
            uniq = list(dict.fromkeys(reasons)) or ["nicio variantă a frazei nu se aplică datelor din acest an"]
            gap = f"[de completat: {'; '.join(uniq)}]"
            blocks[-1].append(gap)
            log.append({"sentence": i, "variant": None, "text": gap, "facts": [], "claims": [], "template": None})
            run.issue("GAP", "blocking", f"sentence {i} left as a gap: {'; '.join(uniq)}")
    paragraphs = [" ".join(b) for b in blocks if b]
    return "\n\n".join(paragraphs) + "\n", log


def trace_check(run: Run, text: str, log: list[dict]) -> list[str]:
    """Every decimal figure outside gaps must be a ledger value used by its sentence, or a
    literal of the approved template it came from."""
    problems = []
    for entry in log:
        if entry["variant"] is None:
            continue
        allowed = set(FIGURE.findall(PLACEHOLDER.sub("", entry["template"])))
        for key in entry["facts"]:
            fid, _, which = key.partition("@")
            f = run.facts[fid]
            shown = f.prior_display if which else f.display
            allowed |= set(FIGURE.findall(shown or ""))
        for fig in FIGURE.findall(GAP.sub("", entry["text"])):
            if fig not in allowed:
                problems.append(f"figure {fig} in sentence {entry['sentence']} has no ledger source")
    if "{" in text or "}" in text:
        problems.append("unfilled placeholder in the draft")
    return problems


def _words(text: str) -> list[str]:
    return re.findall(r"\w+", norm(GAP.sub(" ", text)))


def similarity_check(text: str, other: str, n: int) -> list[str]:
    a, b = _words(text), _words(other)
    grams = {tuple(b[i:i + n]) for i in range(len(b) - n + 1)}
    shared = {" ".join(a[i:i + n]) for i in range(len(a) - n + 1) if tuple(a[i:i + n]) in grams}
    return sorted(shared)
