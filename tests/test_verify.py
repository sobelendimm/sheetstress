import json

import pytest

from sheetstress import cli, lo

needs_lo = pytest.mark.skipif(lo.find_soffice() is None, reason="LibreOffice not installed")
pytestmark = []


def test_logical_flags_bool_compare(ex):
    f = cli.run(str(ex / "bad_logical.xlsx"), ["logical"])
    msgs = " ".join(x.message for x in f)
    assert "ISERROR" in msgs and "SUMPRODUCT" in msgs


def test_logical_clean(ex):
    assert cli.run(str(ex / "clean.xlsx"), ["logical"]) == []


def test_cycle_found(ex):
    f = [x for x in cli.run(str(ex / "bad_cycle.xlsx"), ["cycle"]) if x.severity == "error"]
    sizes = sorted(x.message.split("among ")[1].split(" ")[0] for x in f)
    assert sizes == ["1", "2"]  # C1 self reference + A1/B1 loop


def test_cycle_clean(ex):
    assert cli.run(str(ex / "clean.xlsx"), ["cycle"]) == []


def test_exit_codes_and_json(ex, capsys):
    assert cli.main([str(ex / "bad_cycle.xlsx"), "--checks", "cycle", "--json"]) == 1
    data = json.loads(capsys.readouterr().out)
    assert data["pass"] is False
    assert cli.main([str(ex / "clean.xlsx"), "--checks", "logical,cycle"]) == 0
    assert cli.main([str(ex / "nope.xlsx")]) == 2
    assert cli.main([str(ex / "clean.xlsx"), "--checks", "bogus"]) == 2


@needs_lo
@pytest.mark.lo
def test_error_after_recalc(ex):
    f = cli.run(str(ex / "bad_error.xlsx"), ["error"])
    assert any("#DIV/0!" in x.message for x in f)


@needs_lo
@pytest.mark.lo
def test_injection_fragile_vs_clean(ex):
    modes = ["text", "blankstr"]
    bad = [x for x in cli.run(str(ex / "fragile_input.xlsx"), modes) if x.severity == "error"]
    assert bad and all("Calc!B1" in x.location for x in bad)
    good = [x for x in cli.run(str(ex / "clean.xlsx"), ["text", "edge", "blankstr"]) if x.severity == "error"]
    assert good == []


def test_numeric_is_functions_not_flagged(tmp_path):
    import openpyxl
    wb = openpyxl.Workbook()
    wb.active["A1"] = "=ISOWEEKNUM(DATE(2026,1,5))>0"
    wb.save(tmp_path / "w.xlsx")
    assert cli.run(str(tmp_path / "w.xlsx"), ["logical"]) == []


def test_sheet_name_case_insensitive_cycle(tmp_path):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = "=data!B1"
    ws["B1"] = "=A1"
    wb.save(tmp_path / "c.xlsx")
    assert any(x.severity == "error" for x in cli.run(str(tmp_path / "c.xlsx"), ["cycle"]))


def test_module_exit_code(ex):
    import subprocess
    import sys
    r = subprocess.run([sys.executable, "-m", "sheetstress", str(ex / "bad_cycle.xlsx"), "--checks", "cycle"], capture_output=True)
    assert r.returncode == 1


def _logical(tmp_path, formulas):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    for i, f in enumerate(formulas, 1):
        ws.cell(i, 1, f)
    wb.save(tmp_path / "t.xlsx")
    return cli.run(str(tmp_path / "t.xlsx"), ["logical"])


@pytest.mark.parametrize("formula", [
    "=ISERROR(A2)>0",
    "=0<ISERROR(A2)",
    "=IF(0<ISERROR(A2),1,2)",
    "=ISNUMBER(A2) = 1",
    "=ISERROR(IF(A2,(B2+(C2*(D2+1))),1))>0",
    "=AND(A2>1,B2<3)=1",
    "=ISERROR(A2)>-1",
    "=(ISERROR(A2))>0",
    "=IF((ISNUMBER(A2))=1,1,2)",
    "=0<(ISERROR(A2))",
])
def test_logical_tokenizer_flags(tmp_path, formula):
    f = _logical(tmp_path, [formula])
    assert len(f) == 1 and f[0].severity == "error", formula


@pytest.mark.parametrize("formula", [
    '="ISERROR(x)>0"',
    "=--ISERROR(A2)>0",
    "=N(ISERROR(A2))>0",
    "=ISERROR(A2)*1>0",
    "=1+ISERROR(A2)>0",
    "=0<ISERROR(A2)+1",
    "=ISERROR(A2)=B2",
    "=ISNUMBER(A2)=TRUE",
    "=ISERROR(A2)",
    "=IF(ISERROR(A2),0,1)>0",
    "=A2>0",
    "=(--ISERROR(A2))>0",
])
def test_logical_tokenizer_no_false_alarm(tmp_path, formula):
    assert _logical(tmp_path, [formula]) == [], formula


def test_full_column_sumproduct_needs_array():
    from sheetstress.logical import needs_array
    assert needs_array("=SUMPRODUCT(IF(A:A>0,1,0))")
    assert needs_array("=SUMPRODUCT(--IF($A:$A>0,1,0))")
    assert needs_array("=SUMPRODUCT(IF(1:1>0,1,0))")
    assert needs_array("=SUMPRODUCT(IF(Sheet1!A:A>0,1,0))")
    assert not needs_array("=SUMPRODUCT(IF(A1>0,1,0))")


def _fake_recalc(monkeypatch, missing=()):
    """Replace the LibreOffice run: every file 'recalculates' to itself, except those whose stem is in `missing`."""
    import shutil
    import tempfile
    from pathlib import Path

    from sheetstress import recalc_checks

    def fake(soffice, files, timeout=300):
        work = Path(tempfile.mkdtemp(prefix="sheetstress-fake-"))
        res = {}
        for f in files:
            if Path(f).stem in missing:
                res[f] = None
                continue
            dst = work / Path(f).name
            shutil.copy(f, dst)
            res[f] = str(dst)
        return res, work
    monkeypatch.setattr(recalc_checks.lo, "recalc", fake)
    monkeypatch.setattr(cli.lo, "find_soffice", lambda: "/fake/soffice")


def test_no_unlocked_cells_reports_skipped(ex, monkeypatch):
    """bad_error.xlsx has no unlocked cells: the injection checks must say so (unconditional, LibreOffice stubbed)."""
    from sheetstress import recalc_checks
    assert recalc_checks.input_cells(str(ex / "bad_error.xlsx"), 30) == []
    _fake_recalc(monkeypatch)
    f = cli.run(str(ex / "bad_error.xlsx"), ["text"])
    sk = [x for x in f if x.severity == "skipped"]
    assert sk and sk[0].message.startswith("SKIPPED: no unlocked input cells found")


@needs_lo
@pytest.mark.lo
def test_no_unlocked_cells_reports_skipped_real_lo(ex):
    f = cli.run(str(ex / "bad_error.xlsx"), ["text", "edge", "blankstr"])
    sk = [x for x in f if x.severity == "skipped"]
    assert sk and sk[0].message.startswith("SKIPPED: no unlocked input cells found")


def test_inject_all_false_reports_skipped(ex, monkeypatch, capsys):
    from sheetstress import recalc_checks
    _fake_recalc(monkeypatch)
    monkeypatch.setattr(recalc_checks, "inject", lambda *a, **k: False)
    f = cli.run(str(ex / "clean.xlsx"), ["text"])
    sk = [x for x in f if x.severity == "skipped"]
    assert sk and sk[0].message.startswith("SKIPPED: could not inject")
    assert cli.main([str(ex / "clean.xlsx"), "--checks", "text"]) == 0
    assert capsys.readouterr().out.strip().splitlines()[-1] == "sheetstress: PASS (with skipped checks)"


def test_inject_partial_false_reports_skipped(ex, monkeypatch):
    from sheetstress import recalc_checks
    real = recalc_checks.inject
    _fake_recalc(monkeypatch)
    monkeypatch.setattr(recalc_checks, "inject", lambda src, dst, sx, ref, value, kind:
                        False if value == "-1" else real(src, dst, sx, ref, value, kind))
    f = cli.run(str(ex / "clean.xlsx"), ["edge"])
    sk = [x for x in f if x.severity == "skipped"]
    assert len(sk) == 1 and sk[0].message.startswith("SKIPPED: could not inject 1 of 4")


def test_lo_injected_recalc_missing_reports_skipped(ex, monkeypatch):
    """LibreOffice produced no output for an injected copy: that check did not run -> SKIPPED, not a quiet warning."""
    from sheetstress import recalc_checks
    assert recalc_checks.input_cells(str(ex / "clean.xlsx"), 30)   # fixture really has an unlocked cell
    _fake_recalc(monkeypatch, missing={"inj1"})   # first injected copy is named inj1 by run()
    f = cli.run(str(ex / "clean.xlsx"), ["text"])
    sk = [x for x in f if x.severity == "skipped"]
    assert len(sk) == 1 and sk[0].message.startswith("SKIPPED: recalculation produced no output for 1 of 1")
    assert not [x for x in f if x.severity == "warning"]


def test_lo_base_recalc_failure_is_error(ex, monkeypatch):
    _fake_recalc(monkeypatch, missing={"base"})
    f = cli.run(str(ex / "clean.xlsx"), ["text"])
    assert [x for x in f if x.severity == "error" and "recalculation failed" in x.message]


def test_fail_on_skipped(ex, capsys, monkeypatch):
    monkeypatch.setattr(cli.lo, "find_soffice", lambda: None)
    clean = str(ex / "clean.xlsx")
    assert cli.main([clean, "--checks", "text"]) == 0
    capsys.readouterr()
    assert cli.main([clean, "--checks", "text", "--fail-on-skipped"]) == 4
    assert capsys.readouterr().out.strip().splitlines()[-1] == "sheetstress: FAIL (skipped checks; --fail-on-skipped)"
    assert cli.main([clean, "--checks", "text", "--fail-on-skipped", "--json"]) == 4
    data = json.loads(capsys.readouterr().out)
    assert data["pass"] is False and data["skipped_checks"] is True
    assert cli.main([clean, "--checks", "logical,cycle", "--fail-on-skipped"]) == 0   # nothing skipped
    capsys.readouterr()
    assert cli.main([str(ex / "bad_cycle.xlsx"), "--checks", "cycle,text", "--fail-on-skipped"]) == 1   # errors win


def test_skipped_changes_last_line(ex, capsys, monkeypatch):
    monkeypatch.setattr(cli.lo, "find_soffice", lambda: None)
    assert cli.main([str(ex / "clean.xlsx"), "--checks", "text"]) == 0
    out = capsys.readouterr().out
    assert "SKIPPED" in out
    assert out.strip().splitlines()[-1] == "sheetstress: PASS (with skipped checks)"
    assert cli.main([str(ex / "clean.xlsx"), "--checks", "logical"]) == 0
    assert capsys.readouterr().out.strip().splitlines()[-1] == "sheetstress: PASS"


def test_needs_array_ignores_ranges_inside_string_literals():
    from sheetstress.logical import needs_array
    assert not needs_array('SUMPRODUCT(IF(TEXT(A1,"h:mm")="9:30",1,0))')
    assert not needs_array('SUMPRODUCT(IF(B1="A:B",1,0))')
    assert needs_array('SUMPRODUCT(IF(A:A>0,1,0))')
    assert needs_array('SUMPRODUCT(IF(A1:A9="x:y",1,0))')
