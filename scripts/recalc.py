"""Fills in formula results so the plan opens with numbers everywhere (also in the generator's upload).
openpyxl writes formulas without results; LibreOffice opens the file, calculates, and saves it back.
Usage: python3 scripts/recalc.py plan/Campaign_Plan_Dec2026.xlsx"""
import os, shutil, subprocess, sys, tempfile
src = os.path.abspath(sys.argv[1])
with tempfile.TemporaryDirectory() as out:
    subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", out, src],
                   check=True, capture_output=True, timeout=180)
    shutil.copy(os.path.join(out, os.path.basename(src)), src)
print("recalculated", src)
