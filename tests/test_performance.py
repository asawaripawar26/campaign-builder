"""Checks the performance workbook reaches the right calls on the Aug–Oct data.
Run: python3 tests/test_performance.py
With LibreOffice installed it also throws away every stored result and recalculates all formulas from scratch,
so the test proves the formulas work, not just that the saved numbers look right."""
import os, shutil, subprocess, sys, tempfile
from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BOOK = os.path.join(ROOT, "performance", "Performance_Workbook_AugOct2026.xlsx")
ERRORS = ("#VALUE!", "#DIV/0!", "#REF!", "#NAME?", "#N/A", "#NUM!", "#NULL!")


def rows(wb, tab):
    return [[v for v in r if v is not None] for r in wb[tab].iter_rows(values_only=True)]


def find(wb, tab, key):
    hits = [r for r in rows(wb, tab) if r and r[0] == key or (len(r) > 1 and r[1] == key)]
    assert hits, f"{key} not found on {tab}"
    return hits[0]


def check(path, label):
    wb = load_workbook(path, data_only=True)
    bad = [(ws.title, c.coordinate, c.value) for ws in wb for r in ws.iter_rows() for c in r
           if isinstance(c.value, str) and c.value.startswith(ERRORS)]
    assert not bad, f"formula errors: {bad[:5]}"
    calls = {
        "Search_NonBrand_UK": "Pause candidate",
        "Prospecting_Broad": "Reduce or refresh",
        "Retargeting_30d": "Reduce or refresh",
        "PMax_US": "Investigate",
        "Search_NonBrand_US": "Inconclusive",
    }
    for campaign, call in calls.items():
        r = find(wb, "Decisions", campaign)
        assert any(isinstance(v, str) and v.startswith(call) for v in r), (campaign, r)
    founder = find(wb, "Decisions", "FOUNDER_STORY")
    assert any(isinstance(v, str) and v.startswith("Winner") for v in founder), founder
    assert round(founder[2]) == 132 and round(founder[3]) == 200, founder           # 132 trials at ~$200
    assert 173 <= founder[4] <= 174 and 232 <= founder[5] <= 233, founder           # 90% range $173–233
    testimonial = find(wb, "Decisions", "TESTIMONIAL")
    assert any(isinstance(v, str) and v.startswith("Retire") for v in testimonial), testimonial
    weeks = find(wb, "Summary", "Weeks with Meta tracking diverging")
    assert weeks[1] == 4, weeks
    print(f"ok - {label}: UK pause, Broad/Retargeting reduce, PMax investigate, founder winner ($173–233), "
          f"testimonial retire, Meta tracking diverging 4 weeks, no formula errors")


check(BOOK, "saved workbook")

if shutil.which("soffice"):
    with tempfile.TemporaryDirectory() as tmp:
        stripped = os.path.join(tmp, "fresh.xlsx")
        load_workbook(BOOK).save(stripped)            # openpyxl drops every stored formula result
        assert load_workbook(stripped, data_only=True)["Decisions"]["K5"].value is None
        subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", os.path.join(tmp, "out"), stripped],
                       check=True, capture_output=True, timeout=300)
        check(os.path.join(tmp, "out", "fresh.xlsx"), "recalculated from scratch")
else:
    print("skip - LibreOffice not installed, so formulas were not recalculated")
