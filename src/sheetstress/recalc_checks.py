"""LibreOffice-based checks: error values after recalculation, and input injection (text/edge/blankstr)."""
import re
import shutil
import tempfile
import zipfile
from html import unescape
from pathlib import Path
from xml.sax.saxutils import escape

import openpyxl
from openpyxl.utils import get_column_letter as L

from . import lo
from .findings import Finding

ERRORS = ("#N/A", "#VALUE!", "#DIV/0!", "#REF!", "#NAME?", "#NUM!", "#NULL!")
INJECT = {
    "text": [("abc", "text")],
    "edge": [("1E-307", "num"), ("1E+307", "num"), ("-1", "num"), ("-1E+307", "num")],
    "blankstr": [('""', "formula")],
}


def sheet_paths(z):
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    tgt = {}
    for m in re.finditer(r"<Relationship\b([^>]*)/?>", rels):
        i, t = re.search(r'Id="([^"]+)"', m.group(1)), re.search(r'Target="([^"]+)"', m.group(1))
        if i and t:
            p = t.group(1).lstrip("/")
            tgt[i.group(1)] = p if p.startswith("xl/") else "xl/" + p
    res = {}
    for m in re.finditer(r"<(?:\w+:)?sheet\b([^>]*)/>", wb):
        a = m.group(1)
        rid = re.search(r'r:id="([^"]+)"', a) or re.search(r'\bid="([^"]+)"', a)
        res[unescape(re.search(r'name="([^"]*)"', a).group(1))] = tgt[rid.group(1)]
    return res


def visible_errors(path, only_formulas=True):
    """{(sheet, ref): error} for visible formula cells (cached values)."""
    f = openpyxl.load_workbook(path)
    v = openpyxl.load_workbook(path, data_only=True)
    out = {}
    for ws in v.worksheets:
        if ws.sheet_state != "visible":
            continue
        fs = f[ws.title]
        hidden_r = {r for r, d in ws.row_dimensions.items() if d.hidden}
        hidden_c = set()
        for d in ws.column_dimensions.values():
            if d.hidden and d.min and d.max:
                hidden_c.update(range(d.min, d.max + 1))
        for row in ws.iter_rows():
            for c in row:
                if c.data_type == "e" or (isinstance(c.value, str) and c.value in ERRORS):
                    if c.row in hidden_r or c.column in hidden_c:
                        continue
                    fv = fs[c.coordinate].value
                    if only_formulas and not (isinstance(fv, str) and fv.startswith("=")) and type(fv).__name__ != "ArrayFormula":
                        continue
                    out[(ws.title, c.coordinate)] = str(c.value)
    return out


def input_cells(path, limit):
    """Sample unlocked input cells (round-robin across sheets): first and last of each column run, capped."""
    wb = openpyxl.load_workbook(path)
    per_sheet = []
    for ws in wb.worksheets:
        runs = []
        if ws.sheet_state != "visible":
            continue
        cols = {}
        for row in ws.iter_rows():
            for c in row:
                if c.protection.locked is False and not (isinstance(c.value, str) and c.value.startswith("=")):
                    cols.setdefault(c.column, []).append(c.row)
        for col, rows in sorted(cols.items()):
            rows.sort()
            for r in sorted({rows[0], rows[-1]}):
                runs.append((ws.title, f"{L(col)}{r}"))
        if runs:
            per_sheet.append(runs)
    out = []
    i = 0
    while len(out) < limit and any(i < len(r) for r in per_sheet):
        out += [r[i] for r in per_sheet if i < len(r)]
        i += 1
    return out[:max(limit, 0)]


def inject(src, dst, sheet_xml, ref, value, kind):
    pat = re.compile(r'<c r="%s"(?=[\s>/])([^>]*?)(?:/>|>(.*?)</c>)' % ref, re.S)
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == sheet_xml:
                xml = data.decode("utf-8")
                m = pat.search(xml)
                if not m:
                    return False
                s = re.search(r'\ss="\d+"', m.group(1))
                s = s.group(0) if s else ""
                if kind == "text":
                    new = f'<c r="{ref}"{s} t="inlineStr"><is><t>{escape(value)}</t></is></c>'
                elif kind == "formula":
                    new = f'<c r="{ref}"{s} t="str"><f>{escape(value)}</f><v></v></c>'
                else:
                    new = f'<c r="{ref}"{s}><v>{value}</v></c>'
                data = (xml[:m.start()] + new + xml[m.end():]).encode("utf-8")
            zout.writestr(it, data)
    return True


def run(path, modes, soffice, max_samples=30, timeout=600):
    """modes: subset of {'error','text','edge','blankstr'}. Returns findings."""
    out = []
    work0 = Path(tempfile.mkdtemp(prefix="sheetstress-in-"))
    try:
        base_copy = work0 / "base.xlsx"
        shutil.copy(path, base_copy)
        jobs = [(str(base_copy), None)]
        with zipfile.ZipFile(path) as z:
            paths = sheet_paths(z)
        inj = [m for m in modes if m in INJECT]
        if inj:
            cells = input_cells(path, max_samples)
            n = 0
            tried = failed = 0
            for mode in inj:
                for sheet, ref in cells:
                    for value, kind in INJECT[mode]:
                        n += 1
                        tried += 1
                        dst = work0 / f"inj{n}.xlsx"
                        if inject(path, dst, paths[sheet], ref, value, kind):
                            jobs.append((str(dst), (mode, sheet, ref, value)))
                        else:
                            failed += 1
            if not cells:
                out.append(Finding("input", "skipped", "-", "SKIPPED: no unlocked input cells found; "
                                   + ",".join(inj) + " checks not run "
                                   "(protect sheets and unlock input cells to enable them)"))
            elif failed:
                out.append(Finding("input", "skipped", "-", f"SKIPPED: could not inject {failed} of {tried} test value(s) "
                                   f"into the {len(cells)} sampled input cell(s); those " + ",".join(inj)
                                   + " checks did not run"))
        res, work1 = lo.recalc(soffice, [j[0] for j in jobs], timeout)
        try:
            if not res[jobs[0][0]]:
                out.append(Finding("error", "error", "-", "LibreOffice recalculation failed or timed out; no recalculation result"))
                return out
            base = visible_errors(res[jobs[0][0]])
            if "error" in modes:
                for (s, r), e in sorted(base.items()):
                    out.append(Finding("error", "error", f"{s}!{r}", f"error value {e} after recalculation"))
            no_out = []
            for job, meta in jobs[1:]:
                p = res[job]
                if not p:
                    no_out.append(meta)
                    continue
                for (s, r), e in sorted(visible_errors(p).items()):
                    if (s, r) not in base and (s, r) != (meta[1], meta[2]):
                        out.append(Finding(meta[0], "error", f"{s}!{r}",
                                           f"input {meta[1]}!{meta[2]} = {meta[3]} makes this visible formula return {e}"))
            if no_out:                                       # one aggregated finding, not one per injected copy
                ex = ", ".join(f"{m[1]}!{m[2]}={m[3]}" for m in no_out[:3])
                out.append(Finding("input", "skipped", "-", f"SKIPPED: recalculation produced no output for {len(no_out)} "
                                   f"of {len(jobs) - 1} injected cop{'y' if len(jobs) == 2 else 'ies'} (e.g. {ex}); "
                                   "those checks did not run"))
        finally:
            shutil.rmtree(work1, ignore_errors=True)
    finally:
        shutil.rmtree(work0, ignore_errors=True)
    return out
