"""Locate LibreOffice and recalculate workbooks with it (stdlib only)."""
import glob
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

LINUX = ("/usr/lib/libreoffice/program/soffice", "/usr/lib64/libreoffice/program/soffice",
         "/opt/libreoffice*/program/soffice", "/snap/bin/libreoffice")
MAC = "/Applications/LibreOffice.app/Contents/MacOS/soffice"

SETTINGS = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Formula/Calculation"><prop oor:name="UseThreadedCalculationForFormulaGroups" oor:op="fuse"><value>false</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>
"""


def find_soffice():
    """Path to soffice or None. Honors XLSX_VERIFY_SOFFICE."""
    env = os.environ.get("XLSX_VERIFY_SOFFICE")
    if env and os.path.isfile(env):
        return env
    for name in ("soffice", "libreoffice", "soffice.exe"):
        p = shutil.which(name)
        if p:
            return p
    cands = []
    if sys.platform.startswith("win"):
        for v in ("ProgramFiles", "ProgramW6432", "ProgramFiles(x86)"):
            d = os.environ.get(v)
            if d:
                cands.append(os.path.join(d, "LibreOffice", "program", "soffice.exe"))
    elif sys.platform == "darwin":
        cands.append(MAC)
    else:
        for g in LINUX:
            cands += sorted(glob.glob(g), reverse=True)
    for c in cands:
        if os.path.isfile(c):
            return c
    return None


def recalc(soffice, files, timeout=300):
    """Recalculate each file in one LibreOffice run. Returns {src: recalculated path or None}."""
    work = Path(tempfile.mkdtemp(prefix="xlsx-verify-"))
    prof = work / "profile"
    (prof / "user").mkdir(parents=True)
    (prof / "user" / "registrymodifications.xcu").write_text(SETTINGS, encoding="utf-8")
    out = work / "out"
    out.mkdir()
    files = [str(f) for f in files]
    cmd = [soffice, "--headless", "--norestore", "--nolockcheck",
           "-env:UserInstallation=" + (prof.as_uri()),
           "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(out)] + files
    try:
        subprocess.run(cmd, capture_output=True, timeout=timeout, check=False)
    except (OSError, subprocess.TimeoutExpired):
        pass
    res = {}
    for f in files:
        p = out / (Path(f).stem + ".xlsx")
        res[f] = str(p) if p.exists() else None
    return res, work
