# Checks in detail

## logical
Excel orders values as numbers < text < logicals. `ISERROR(A1)>0` is therefore TRUE for every `A1`,
even when no error exists. LibreOffice treats FALSE as 0, so the test passes there. Wrap the logical in
`N()` or `--` to compare numerically. A second rule flags `SUMPRODUCT(IFERROR(range,..))` and
`SUMPRODUCT(IF(range,..))` that is not stored as an array formula: Excel (pre-dynamic-array) evaluates the
inner function with implicit intersection, so the counter silently returns 0. A third rule warns when a
conditional format over several ranges uses a relative formula that does not mention the first cell.
Limit: the scan is textual and covers up to two parenthesis levels inside the logical function.

## cycle
Builds a dependency graph from every formula (`openpyxl` tokenizer, cross-sheet references, defined
names, whole-column references bounded by the used area) and reports strongly connected components,
including a cell that references itself. `OFFSET` and `INDIRECT` are not followed and are listed as info.

## error
All sheets are recalculated by LibreOffice with "always recalculate" and threaded calculation disabled;
visible (not hidden sheet, not hidden row) formula cells that evaluate to an error are findings.

## text, edge, blankstr
For a sample of unlocked input cells (first and last of each column run, capped by `--max-samples`) a copy
of the workbook is written in which only that cell XML is replaced (text `abc`; extreme numbers; `=""`).
All copies are recalculated in one LibreOffice call. A finding is a visible formula cell that has an error
in the copy but not in the baseline. Not detected: wrong-but-error-free results (silent acceptance), and
inputs where an error is the intended answer.
