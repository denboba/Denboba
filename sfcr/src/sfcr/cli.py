"""Entry point: validate config, run, write output/. Run from the folder holding the inputs."""

from __future__ import annotations

import argparse
import hashlib
import json
import platform
import sys
import tomllib
from importlib.metadata import version
from pathlib import Path

from .engine import Run
from .render import PLACEHOLDER, render, similarity_check, trace_check

VERSION = "0.1.0"


class ConfigError(Exception):
    pass


def validate(cfg: dict, base: Path) -> None:
    for key in ("report", "inputs", "output", "workbook_layout", "tolerances", "fact", "sentence"):
        if key not in cfg:
            raise ConfigError(f"missing [{key}]")
    ids = [f["id"] for f in cfg["fact"]]
    if len(ids) != len(set(ids)):
        raise ConfigError("duplicate fact ids")
    for f in cfg["fact"]:
        if f["kind"] not in ("amount", "ratio", "count"):
            raise ConfigError(f"fact {f['id']}: unknown kind {f['kind']!r}")
        if f.get("source", "workbook") == "workbook" and not (f.get("sheet") and f.get("label")):
            raise ConfigError(f"fact {f['id']}: workbook facts need sheet and label")
    known = set(ids)
    for c in cfg.get("check", []):
        refs = [t.lstrip("+-") for t in c.get("terms", [])] + [c.get(k) for k in ("equals", "less", "than") if c.get(k)]
        refs += c.get("on_fail", [])
        if set(refs) - known:
            raise ConfigError(f"check {c['id']}: unknown facts {set(refs) - known}")
    claims = {c["id"] for c in cfg.get("claim", [])}
    for d in cfg.get("decision", []):
        for k in ("fact", "choose", "source", "reason", "decided_by", "date"):
            if k not in d:
                raise ConfigError(f"decision {d.get('id')}: missing {k!r} (every decision needs a source and a reason)")
        if d["fact"] not in known:
            raise ConfigError(f"decision {d['id']}: unknown fact {d['fact']}")
    for s in cfg["sentence"]:
        for v in s.get("variant", []):
            for key in PLACEHOLDER.findall(v["text"]):
                if key not in ("year", "prior_year", "date", "prior_date") and key.partition("@")[0] not in known:
                    raise ConfigError(f"placeholder {{{key}}} is not a fact")
            for c in v.get("claims", []):
                if c not in claims:
                    raise ConfigError(f"unknown claim {c}")
    for name in cfg["inputs"].values():
        if not (base / name).exists():
            raise ConfigError(f"input not found: {base / name} (run the command from the folder holding the inputs)")


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def ledger(run: Run) -> list[dict]:
    out = []
    for f in run.facts.values():
        out.append({
            "id": f.id, "kind": f.kind, "status": f.status,
            "value_lei_or_fraction": None if f.value is None else str(f.value),
            "display": f.display, "source": f.source, "why": f.why,
            "gap_reason_ro": None if f.usable else f.reason_ro,
            "prior": None if f.prior is None else str(f.prior), "prior_status": f.prior_status,
            "prior_display": f.prior_display, "prior_source": f.prior_source,
            "checks": f.checks, "decision": f.decision,
        })
    return out


def _cell(s) -> str:
    return str(s).replace("|", "\\|").replace("\n", " ")


def review_report(run: Run, text: str, log, trace, similar, extra_files) -> str:
    y = run.year
    blocking = [i for i in run.issues if i.severity == "blocking"]
    lines = [f"# Section E {y}: review report", ""]
    state = "NOT READY: unresolved blocking items" if blocking else "Ready for human review"
    lines += [f"**Status: {state}.** A clean run is not an approval; every item under *Review points* needs a person.", ""]
    lines += ["## Figures", "", "| Fact | Status | Value | Prior | Source | Why this cell / reason |", "|---|---|---|---|---|---|"]
    for f in run.facts.values():
        src = f.source
        where = f"{src.get('file', '')} {src.get('sheet', '')}!{src.get('cell', '')}" if src.get("sheet") else \
            f"{src.get('file', '')} p.{src.get('page')}"
        why = f.reason_ro if not f.usable else "; ".join(f.why)
        lines.append(f"| {f.id} | {f.status}{' (' + f.decision + ')' if f.decision else ''} | {f.display or '—'} | "
                     f"{f.prior_display or '—'} | {_cell(where)} | {_cell(why)} |")
    lines += ["", "## Checks", "", "| Check | Result | Detail |", "|---|---|---|"]
    for c in run.check_results:
        lines.append(f"| {c['id']} | {c['result']} | {_cell(c['what'])}: {_cell(c['detail'])} |")
    lines.append(f"| TRACE | {'pass' if not trace else 'FAIL'} | {_cell('; '.join(trace)) or 'every figure in the draft has a ledger source'} |")
    lines.append(f"| SIMILARITY | {'pass' if not similar else 'FAIL'} | "
                 f"{_cell('; '.join(similar)) or 'no shared ' + str(run.cfg['tolerances']['similarity_ngram']) + '-word sequence with ' + run.cfg['inputs']['other_company_report']} |")
    lines += ["", "## Issues", ""]
    for sev in ("blocking", "review", "warning", "info"):
        for i in [x for x in run.issues if x.severity == sev]:
            lines.append(f"- **{sev}** `{i.code}`: {i.message}")
    if extra_files:
        lines.append(f"- **warning** `FOREIGN_OUTPUT`: files in output/ not written by this run: {extra_files}")
    lines += ["", "## Review points (a person must sign off)", ""]
    for d in run.cfg.get("decision", []):
        applied = any(f.decision == d["id"] for f in run.facts.values())
        lines.append(f"- Decision **{d['id']}** on `{d['fact']}` ({'applied' if applied else 'NOT applied: values differ from those it was taken on'}): "
                     f"{_cell(d['reason'])} *Source:* {_cell(d['source'])}. *By:* {d['decided_by']}, {d['date']}.")
    for e in log:
        for c in e["claims"]:
            cl = run.claims[c]
            lines.append(f"- Claim **{c}** used: “{cl['quote']}” ({cl['file']}, p.{cl['page']}).")
    for e in log:
        if e["variant"] is None:
            lines.append(f"- Gap to fill: {e['text']}")
    lines += ["", "## Requests declined", ""]
    for r in run.cfg.get("refused_request", []):
        lines.append(f"- *{r['source']}*: “{r['request']}”. {_cell(r['reason'])}")
    lines += ["", f"## {run.prior_year} mapping audit", "", "| Text | Occ. | Cell | Problem |", "|---|---|---|---|"]
    for m in run.mapping_audit:
        if m["problems"]:
            lines.append(f"| {m['text']} | {m['occurrence']} | {m['cell'] or '—'} | {_cell('; '.join(m['problems']))} |")
    ok = sum(1 for m in run.mapping_audit if not m["problems"])
    lines += ["", f"{ok} other mapping entries reproduce their {run.prior_year} text exactly.", ""]
    return "\n".join(lines)


def main(argv=None) -> int:
    root = Path(__file__).resolve().parents[2]
    ap = argparse.ArgumentParser(description="Roll forward SFCR section E. Run from the folder with the inputs (date/).")
    ap.add_argument("--config", type=Path, default=root / "config" / "section_e.toml")
    args = ap.parse_args(argv)
    base = Path.cwd()
    try:
        cfg = tomllib.loads(args.config.read_text(encoding="utf-8"))
        validate(cfg, base)
    except (OSError, tomllib.TOMLDecodeError, ConfigError, KeyError) as e:
        print(f"error: configuration: {e}", file=sys.stderr)
        return 2
    try:
        run = Run(cfg, base).run()
    except Exception as e:  # a broken input must stop the run loudly, never yield a partial draft
        print(f"error: could not process inputs: {type(e).__name__}: {e}", file=sys.stderr)
        return 2

    text, log = render(run)
    trace = trace_check(run, text, log)
    for p in trace:
        run.issue("TRACE", "blocking", p)
    similar = similarity_check(text, (base / cfg["inputs"]["other_company_report"]).read_text(encoding="utf-8"),
                               int(cfg["tolerances"]["similarity_ngram"]))
    if similar:
        run.issue("COPIED_TEXT", "blocking", f"shares wording with {cfg['inputs']['other_company_report']}: {similar}")

    out = (base / cfg["output"]["dir"]).resolve()
    out.mkdir(parents=True, exist_ok=True)
    y, sec = cfg["report"]["year"], cfg["report"]["section"]
    names = {
        "text": f"sectiunea_{sec}_{y}.txt",
        "ledger": f"figuri_{sec}_{y}.json",
        "review": f"verificari_{sec}_{y}.md",
        "manifest": "manifest.json",
    }
    extra = sorted(p.name for p in out.iterdir() if p.name not in names.values())
    (out / names["text"]).write_text(text, encoding="utf-8")
    (out / names["ledger"]).write_text(json.dumps({"facts": ledger(run), "sentences": log, "claims": run.claims},
                                                  ensure_ascii=False, indent=2, default=str) + "\n", encoding="utf-8")
    (out / names["review"]).write_text(review_report(run, text, log, trace, similar, extra), encoding="utf-8")
    files = sorted(run.used_files)
    manifest = {
        "program": f"sfcr section-E {VERSION}",
        "config": {"path": args.config.name, "sha256": sha256(args.config)},
        "inputs": {name: sha256(base / name) for name in files},
        "inputs_present_not_used": sorted(p.name for p in base.iterdir() if p.is_file() and p.name not in files),
        "environment": {"python": platform.python_version(),
                        **{pkg: version(pkg) for pkg in ("openpyxl", "pdfplumber", "python-docx")}},
        "outputs": {k: v for k, v in names.items() if k != "manifest"},
        "summary": {sev: sum(1 for i in run.issues if i.severity == sev) for sev in ("blocking", "review", "warning", "info")},
    }
    (out / names["manifest"]).write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    s = manifest["summary"]
    gaps = sum(1 for e in log if e["variant"] is None)
    print(f"wrote {len(names)} files to {out}: {gaps} gap(s), {s['blocking']} blocking, {s['review']} review item(s)")
    return 0
