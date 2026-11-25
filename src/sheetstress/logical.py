"""Static check: formulas that compare a logical value with a number.

Excel ranks every logical above every number, so ISERROR(x)>0 is always TRUE in Excel.
LibreOffice treats FALSE as 0, so it passes there. Also flags SUMPRODUCT(IFERROR/IF(range))
that is not an array formula (Excel applies implicit intersection) and multi-range
conditional-format rules whose relative formula does not use the first cell.
"""
import re
import zipfile
from html import unescape

from openpyxl.formula import Tokenizer

from .findings import Finding

LOGICAL_FUNCS = frozenset((
    "ISERROR", "ISERR", "ISNA", "ISNUMBER", "ISTEXT", "ISBLANK", "ISLOGICAL", "ISNONTEXT",
    "ISREF", "ISEVEN", "ISODD", "ISFORMULA", "AND", "OR", "NOT", "EXACT",
))
COMPARE = frozenset(("=", "<>", "<", ">", "<=", ">="))
ARITH = frozenset(("+", "-", "*", "/", "^", "&"))
ARRAY = re.compile(r"SUMPRODUCT\((?:--)?(?:IFERROR|IF)\(")
_REF = r"\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+"  # A1:B9
_COL = r"\$?[A-Z]{1,3}:\$?[A-Z]{1,3}"  # A:A
_ROW = r"\$?\d+:\$?\d+"  # 1:1
RANGE = re.compile(rf"(?<![A-Za-z0-9_.$])(?:{_REF}|{_COL}|{_ROW})(?![A-Za-z0-9_(])")
CELL = re.compile(r'<c r="([A-Z]+[0-9]+)"[^>]*>\s*<f([^>]*)>([^<]*)</f>')


def needs_array(f):
    for m in ARRAY.finditer(f):
        depth, i = 1, m.end()
        while i < len(f) and depth:
            depth += {"(": 1, ")": -1}.get(f[i], 0)
            i += 1
        if RANGE.search(f, m.end(), i):
            return True
    return False


def _is_op(tok, values):
    return tok.type == "OPERATOR-INFIX" and tok.value in values


def _is_number(tok):
    return tok.type == "OPERAND" and tok.subtype == "NUMBER"


def bool_number_compares(formula):
    """Spans like ``ISERROR(x)>0`` or ``0<ISERROR(x)`` from a token stream.

    Only a numeric literal on the other side counts: comparing with a cell or TRUE/FALSE
    cannot be judged statically. Coerced uses (``--F()``, ``N(F())``, ``F()*1``) are skipped.
    """
    toks = [t for t in Tokenizer(formula).items if t.type != "WHITE-SPACE"]
    stack, spans = [], []
    for i, t in enumerate(toks):
        if t.type == "FUNC" and t.subtype == "OPEN":
            stack.append(i)
        elif t.type == "FUNC" and t.subtype == "CLOSE" and stack:
            start = stack.pop()
            name = toks[start].value[:-1].upper().removeprefix("_XLFN.")
            if name in LOGICAL_FUNCS:
                spans.append((start, i))
    out = []
    for start, end in spans:
        # gereksiz parantezleri aş: =(ISERROR(A2))>0
        while (start and end + 1 < len(toks) and toks[start - 1].type == "PAREN" and toks[start - 1].subtype == "OPEN"
               and toks[end + 1].type == "PAREN" and toks[end + 1].subtype == "CLOSE"):
            start, end = start - 1, end + 1
        prev = toks[start - 1] if start else None
        nxt = toks[end + 1] if end + 1 < len(toks) else None
        if prev is not None and (prev.type == "OPERATOR-PREFIX" or _is_op(prev, ARITH)):
            continue
        if nxt is not None and (_is_op(nxt, ARITH) or nxt.type == "OPERATOR-POSTFIX"):
            continue
        lo = hi = None
        if nxt is not None and _is_op(nxt, COMPARE):
            k = end + 2
            while k < len(toks) and toks[k].type == "OPERATOR-PREFIX":
                k += 1
            if k < len(toks) and _is_number(toks[k]):
                lo, hi = start, k
        if lo is None and prev is not None and _is_op(prev, COMPARE) and start >= 2 and _is_number(toks[start - 2]):
            lo, hi = start - 2, end
        if lo is not None:
            spans_text = "".join(x.value for x in toks[lo:hi + 1])
            out.append(spans_text)
    return out


def check(path):
    out = []
    with zipfile.ZipFile(path) as z:
        for n in sorted(z.namelist()):
            if not n.startswith("xl/worksheets/sheet") or not n.endswith(".xml"):
                continue
            sheet = n.rsplit("/", 1)[-1]
            xml = z.read(n).decode("utf8", "ignore")
            for ref, attr, f in CELL.findall(xml):
                f = unescape(f)
                try:
                    hits = bool_number_compares("=" + f.lstrip("="))
                except Exception:  # noqa: BLE001  # tek bozuk formül tüm taramayı düşürmesin
                    out.append(Finding("logical", "warning", f"{sheet}!{ref}", "formula could not be tokenized: " + f[:60]))
                    hits = []
                for span in hits:
                    out.append(Finding("logical", "error", f"{sheet}!{ref}",
                                       "logical value compared with a number: " + span[:80]))
                if 't="array"' not in attr and needs_array(f):
                    out.append(Finding("logical", "error", f"{sheet}!{ref}",
                                       "SUMPRODUCT(IFERROR/IF(range)) is not an array formula (implicit intersection in Excel)"))
            for m in re.finditer(r'<conditionalFormatting sqref="([^"]+)"[^>]*>(.*?)</conditionalFormatting>', xml, re.DOTALL):
                sq = m.group(1)
                if " " not in sq:
                    continue
                first = sq.split()[0].split(":")[0]
                for f in re.findall(r"<formula>([^<]*)</formula>", m.group(2)):
                    rel = [r for r in re.findall(r"(?<![A-Za-z$!])(\$?[A-Z]{1,3}\$?\d+)", f) if "$" not in r]
                    if rel and first not in rel:
                        out.append(Finding("logical", "warning", f"{sheet} sqref={sq[:50]}",
                                           f"multi-range conditional format: relative formula {unescape(f)[:50]} is anchored to {first}"))
    return out
