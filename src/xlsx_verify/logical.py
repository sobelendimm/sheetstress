"""Static check: formulas that compare a logical value with a number.

Excel ranks every logical above every number, so ISERROR(x)>0 is always TRUE in Excel.
LibreOffice treats FALSE as 0, so it passes there. Also flags SUMPRODUCT(IFERROR/IF(range))
that is not an array formula (Excel applies implicit intersection) and multi-range
conditional-format rules whose relative formula does not use the first cell.
"""
import re
import zipfile
from html import unescape

from .findings import Finding

PAT = re.compile(r"\b(ISERROR|ISERR|ISNA|ISNUMBER|ISTEXT|ISBLANK|ISLOGICAL|ISNONTEXT|ISREF|ISEVEN|ISODD|ISFORMULA|AND|OR|NOT|EXACT)\(((?:[^()]|\((?:[^()]|\([^()]*\))*\))*)\)\s*(?:>=?|<=?|<>|=)\s*-?[0-9]")
ARRAY = re.compile(r"SUMPRODUCT\((?:--)?(?:IFERROR|IF)\(")
RANGE = re.compile(r"\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+")
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
                for m in PAT.finditer(f):
                    before = f[max(0, m.start() - 2):m.start()]
                    if before.endswith("--") or before.endswith("N("):
                        continue
                    out.append(Finding("logical", "error", f"{sheet}!{ref}",
                                       "logical value compared with a number: " + m.group(0)[:80]))
                if 't="array"' not in attr and needs_array(f):
                    out.append(Finding("logical", "error", f"{sheet}!{ref}",
                                       "SUMPRODUCT(IFERROR/IF(range)) is not an array formula (implicit intersection in Excel)"))
            for m in re.finditer(r'<conditionalFormatting sqref="([^"]+)"[^>]*>(.*?)</conditionalFormatting>', xml, re.S):
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
