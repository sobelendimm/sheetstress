# xlsx-verify (Excel template verification kit)

Catches the bugs in `.xlsx` templates that pass in LibreOffice or in your own testing and then break
in a customer's Excel: logical-vs-number comparisons, circular references, error values, and
fragile input cells that turn a pasted text, an extreme number or an empty string into `#VALUE!`.

```
pip install .
xlsx-verify model.xlsx [--checks logical,cycle,error,text,edge,blankstr] [--json] [--max-samples N]
```

Exit codes: `0` pass, `1` at least one error finding, `2` usage / unreadable file.
Last line is `xlsx-verify: PASS` or `xlsx-verify: FAIL (n error(s))`.

| Check | Needs LibreOffice | What it catches |
|---|---|---|
| `logical` | no | `ISERROR(x)>0` style comparisons (always TRUE in Excel), non-array `SUMPRODUCT(IFERROR(range))`, multi-range conditional formats anchored wrong |
| `cycle` | no | static circular references (graph SCC, defined names followed, `OFFSET`/`INDIRECT` reported as info) |
| `error` | yes | visible formula cells showing `#DIV/0!`, `#REF!`, ... after a forced full recalculation |
| `text` | yes | text pasted into an unlocked input cell produces a new visible error elsewhere |
| `edge` | yes | `1E-307`, `1E+307`, `-1`, `-1E+307` in input cells produce a new visible error |
| `blankstr` | yes | an input cell holding `=""` produces a new visible error |

Input cells are the cells with `locked = False` (the usual convention for protected templates). If a
workbook has none, the injection checks report an info line and are skipped. LibreOffice is found on
`PATH`, in the usual per-OS locations, or via `XLSX_VERIFY_SOFFICE`; without it those checks are skipped
with an info line and never fail the run.

See [docs/checks.md](docs/checks.md) for details and [docs/design.md](docs/design.md) for the design notes.
Synthetic sample workbooks: `python examples/make_examples.py outdir`.

## Honest limits

LibreOffice is not Excel. Most notably it computes `TRUE` as `1` and `FALSE` as `0` in comparisons and
sums, while Excel ranks logicals above all numbers and ignores them in range sums. The static `logical`
check exists because the recalculation checks cannot see that class of bug. A clean run is evidence, not
proof; confirm important templates in real Excel.

## Development

```
pip install -e '.[test]' && pytest          # LibreOffice tests are marked `lo` and skipped without it
```

MIT licensed. Sponsorship link: to be added when the repository is published.
