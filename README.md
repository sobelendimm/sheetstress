# sheetstress (Excel template verification kit)

Catches the bugs in `.xlsx` templates that pass in LibreOffice or in your own testing and then break
in a customer's Excel: logical-vs-number comparisons, circular references, error values, and
fragile input cells that turn a pasted text, an extreme number or an empty string into `#VALUE!`.

```
pip install .
sheetstress model.xlsx [--checks logical,cycle,error,text,edge,blankstr] [--json] [--max-samples N] [--fail-on-skipped]
```

Exit codes: `0` pass, `1` at least one error finding, `2` usage / unreadable file,
`4` a check was skipped and `--fail-on-skipped` was given (errors still win: `1`).
Last line is `sheetstress: PASS`, `sheetstress: PASS (with skipped checks)` or `sheetstress: FAIL (n error(s))`.
`PASS (with skipped checks)` (still exit `0`; JSON `"skipped_checks": true`) means a requested check could not
run (no LibreOffice, no unlocked input cells for `text`/`edge`/`blankstr`, test values that could not be injected, or
LibreOffice returning no output for an injected copy); each is listed as a `SKIPPED:` finding,
so a plain `PASS` always means every requested check actually ran.

| Check | Needs LibreOffice | What it catches |
|---|---|---|
| `logical` | no | `ISERROR(x)>0` style comparisons (always TRUE in Excel), non-array `SUMPRODUCT(IFERROR(range))`, multi-range conditional formats anchored wrong |
| `cycle` | no | static circular references (graph SCC, defined names followed, `OFFSET`/`INDIRECT` reported as info) |
| `error` | yes | visible formula cells showing `#DIV/0!`, `#REF!`, ... after a forced full recalculation |
| `text` | yes | text pasted into an unlocked input cell produces a new visible error elsewhere |
| `edge` | yes | `1E-307`, `1E+307`, `-1`, `-1E+307` in input cells produce a new visible error |
| `blankstr` | yes | an input cell holding `=""` produces a new visible error |

Input cells are the cells with `locked = False` (the usual convention for protected templates). If a
workbook has none, the injection checks report a `SKIPPED` finding and do not run. LibreOffice is found on
`PATH`, in the usual per-OS locations, or via `SHEETSTRESS_SOFFICE`; without it those checks are skipped
with a `SKIPPED` finding and never fail the run.

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

MIT licensed. Issues and pull requests are welcome.
