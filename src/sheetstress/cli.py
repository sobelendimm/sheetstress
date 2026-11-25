"""sheetstress command line."""
import argparse
import json
import os
import sys

from . import __version__, cycle, lo, logical, recalc_checks
from .findings import Finding

ALL = ("logical", "cycle", "error", "text", "edge", "blankstr")
LO_CHECKS = {"error", "text", "edge", "blankstr"}


def run(path, checks, max_samples=30):
    findings = []
    if "logical" in checks:
        findings += logical.check(path)
    if "cycle" in checks:
        findings += cycle.check(path)
    lo_modes = [c for c in checks if c in LO_CHECKS]
    if lo_modes:
        soffice = lo.find_soffice()
        if not soffice:
            findings.append(Finding("libreoffice", "info", "-", "LibreOffice not found; skipped: " + ",".join(lo_modes)
                                    + " (set SHEETSTRESS_SOFFICE to its path)"))
        else:
            findings += recalc_checks.run(path, lo_modes, soffice, max_samples)
    return findings


def main(argv=None):
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass
    ap = argparse.ArgumentParser(prog="sheetstress", description="Check an .xlsx template for bugs that only show up in real use.")
    ap.add_argument("file", help=".xlsx file")
    ap.add_argument("--checks", default=",".join(ALL), help="comma list of: " + ",".join(ALL) + " (default: all)")
    ap.add_argument("--json", action="store_true", help="machine-readable output")
    ap.add_argument("--max-samples", type=int, default=30, help="max input cells to inject into (default 30)")
    ap.add_argument("--version", action="version", version="sheetstress " + __version__)
    a = ap.parse_args(argv)
    checks = list(dict.fromkeys(c.strip() for c in a.checks.split(",") if c.strip()))
    bad = [c for c in checks if c not in ALL]
    if bad or not checks:
        print("unknown check(s): " + ",".join(bad or ["<none>"]), file=sys.stderr)
        return 2
    if a.max_samples < 0:
        print("--max-samples must be >= 0", file=sys.stderr)
        return 2
    if not os.path.isfile(a.file):
        print("file not found: " + a.file, file=sys.stderr)
        return 2
    try:
        findings = run(a.file, checks, a.max_samples)
    except Exception as e:  # corrupt / not an xlsx
        print(f"cannot read workbook: {e}", file=sys.stderr)
        return 2
    errors = [f for f in findings if f.severity == "error"]
    if a.json:
        print(json.dumps({"file": a.file, "checks": checks, "pass": not errors,
                          "findings": [f.as_dict() for f in findings]}, ensure_ascii=False, indent=2))
    else:
        for f in findings:
            print(f"{f.severity.upper():7} [{f.check}] {f.location}: {f.message}")
        print("sheetstress: PASS" if not errors else f"sheetstress: FAIL ({len(errors)} error(s))")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
