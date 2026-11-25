"""Static circular-reference check using the formula dependency graph (LibreOffice not needed)."""
import re
from collections import defaultdict

import openpyxl
from openpyxl.formula import Tokenizer
from openpyxl.utils import column_index_from_string

from .findings import Finding

CELL = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")
DYN = ("OFFSET", "INDIRECT")


def _split(ref):
    sheet = None
    if "!" in ref:
        sheet, ref = ref.rsplit("!", 1)
        sheet = sheet.strip("'").replace("''", "'")
    return sheet, ref.replace("$", "")


def _area(ref, maxr, maxc):
    if ":" in ref:
        a, b = ref.split(":", 1)
        ma, mb = CELL.match(a), CELL.match(b)
        if ma and mb:
            r1, c1 = int(ma.group(2)), column_index_from_string(ma.group(1).upper())
            r2, c2 = int(mb.group(2)), column_index_from_string(mb.group(1).upper())
            return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)
        if a.isalpha() and b.isalpha():
            return 1, column_index_from_string(a.upper()), maxr, column_index_from_string(b.upper())
        if a.isdigit() and b.isdigit():
            return int(a), 1, int(b), maxc
        return None
    m = CELL.match(ref)
    if not m:
        return None
    r, c = int(m.group(2)), column_index_from_string(m.group(1).upper())
    return r, c, r, c


def check(path):
    wb = openpyxl.load_workbook(path)
    formulas = {}
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith("=") and len(v) > 1:
                    formulas[(ws.title, c.row, c.column)] = v
                elif type(v).__name__ == "ArrayFormula" and getattr(v, "text", None):
                    formulas[(ws.title, c.row, c.column)] = v.text if v.text.startswith("=") else "=" + v.text
    by_sheet = defaultdict(list)
    for k in formulas:
        by_sheet[k[0]].append(k)
    sizes = {ws.title: (ws.max_row or 1, ws.max_column or 1) for ws in wb.worksheets}
    names = {n.lower(): d.attr_text for n, d in wb.defined_names.items()}
    lower = {k.lower(): k for k in sizes}
    out, edges = [], defaultdict(set)
    seen_dyn = set()

    def refs(formula, sheet, seen):
        for t in Tokenizer(formula).items:
            if t.type == "FUNC" and t.subtype == "OPEN" and t.value.upper().rstrip("(") in DYN:
                key = (sheet, t.value.upper().rstrip("("))
                if key not in seen_dyn:
                    seen_dyn.add(key)
                    out.append(Finding("cycle", "info", sheet, f"dynamic reference ({key[1]}) present; not followed by the static scan (reported once per sheet)"))
            if t.type != "OPERAND":
                continue
            if t.subtype == "RANGE":
                v = t.value
                if "!" not in v and v.lower() in names and v.lower() not in seen:
                    yield from refs("=" + names[v.lower()], sheet, seen | {v.lower()})
                    continue
                s, r = _split(v)
                s = s or sheet
                s = lower.get(s.lower())
                if s is None:
                    continue
                a = _area(r, *sizes[s])
                if a:
                    yield s, a

    for k, f in formulas.items():
        try:
            for s, (r1, c1, r2, c2) in refs(f, k[0], frozenset()):
                if (r2 - r1 + 1) * (c2 - c1 + 1) <= len(by_sheet[s]):
                    cand = ((s, r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1))
                else:
                    cand = (x for x in by_sheet[s] if r1 <= x[1] <= r2 and c1 <= x[2] <= c2)
                for x in cand:
                    if x in formulas:
                        edges[k].add(x)
        except Exception as e:  # tokenizer failure on exotic syntax
            out.append(Finding("cycle", "info", f"{k[0]}", f"formula not parsed: {e}"))
    # iterative Tarjan SCC
    index, low, on, stack, idx = {}, {}, set(), [], [0]
    sccs = []
    for root in formulas:
        if root in index:
            continue
        work = [(root, iter(edges[root]))]
        index[root] = low[root] = idx[0]; idx[0] += 1
        stack.append(root); on.add(root)
        while work:
            v, it = work[-1]
            for w in it:
                if w not in index:
                    index[w] = low[w] = idx[0]; idx[0] += 1
                    stack.append(w); on.add(w)
                    work.append((w, iter(edges[w])))
                    break
                if w in on:
                    low[v] = min(low[v], index[w])
            else:
                work.pop()
                if work:
                    low[work[-1][0]] = min(low[work[-1][0]], low[v])
                if low[v] == index[v]:
                    comp = []
                    while True:
                        w = stack.pop(); on.discard(w); comp.append(w)
                        if w == v:
                            break
                    if len(comp) > 1 or v in edges[v]:
                        sccs.append(comp)
    from openpyxl.utils import get_column_letter as L
    for comp in sccs:
        cells = sorted(comp)
        names_ = ", ".join(f"{s}!{L(c)}{r}" for s, r, c in cells[:6]) + (" ..." if len(cells) > 6 else "")
        out.append(Finding("cycle", "error", f"{cells[0][0]}!{L(cells[0][2])}{cells[0][1]}",
                           f"circular reference among {len(cells)} cell(s): {names_}"))
    return out
