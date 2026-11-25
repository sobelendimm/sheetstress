"""Generate small synthetic workbooks used by the docs and the tests (no real product data)."""
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Protection


def _input(ws, ref, value=None):
    ws[ref] = value
    ws[ref].protection = Protection(locked=False)


def build(outdir):
    out = Path(outdir)
    out.mkdir(parents=True, exist_ok=True)

    def save(wb, name):
        for ws in wb.worksheets:
            ws.protection.sheet = True
        wb.save(out / name)

    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Calc"
    _input(ws, "A1", 10)
    ws["B1"] = '=IF(ISNUMBER(A1),IF(ABS(A1)<1E100,A1*2,"too large"),"enter a number")'
    save(wb, "clean.xlsx")

    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Calc"
    ws["A1"] = 5
    ws["B1"] = "=ISERROR(A1)>0"
    ws["C1"] = "=SUMPRODUCT(IFERROR(A1:A3/1,0))"
    save(wb, "bad_logical.xlsx")

    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Calc"
    ws["A1"] = "=B1+1"; ws["B1"] = "=A1+1"; ws["C1"] = "=C1"
    save(wb, "bad_cycle.xlsx")

    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Calc"
    ws["A1"] = 0
    ws["B1"] = "=1/A1"
    save(wb, "bad_error.xlsx")

    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Calc"
    _input(ws, "A1", 10)
    ws["B1"] = "=A1+1"
    save(wb, "fragile_input.xlsx")
    return out


if __name__ == "__main__":
    print(build(sys.argv[1] if len(sys.argv) > 1 else "examples/generated"))
