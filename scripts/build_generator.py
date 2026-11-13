"""Builds index.html (the generator) from src/generator_template.html with the plan workbook embedded
as the 'Load sample plan' data. Re-run after changing the plan or the template.
Usage: python3 scripts/build_generator.py [plan.xlsx]"""
import json, os, sys, datetime as dt
from openpyxl import load_workbook
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PLAN = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "plan", "Campaign_Plan_Dec2026.xlsx")
TABS = ["Settings", "Campaigns", "Concepts", "Claims", "Creative_Library", "Google_Keywords", "Google_Negatives",
        "Google_RSAs", "PMax_AssetGroups", "PMax_Signals", "Meta_AdSets", "Meta_Ads", "Variants"]
wb = load_workbook(PLAN, data_only=True)
if wb["Campaigns"]["K2"].value is None:
    sys.exit("The plan has no formula results yet. Run: python3 scripts/recalc.py " + PLAN)
plan = {}
for t in TABS:
    rows = list(wb[t].iter_rows(values_only=True)); hdr = rows[0]
    plan[t] = [{h: (v.date().isoformat() if isinstance(v, dt.datetime) else ("" if v is None else v))
                for h, v in zip(hdr, r) if h} for r in rows[1:]]
tpl = open(os.path.join(ROOT, "src", "generator_template.html"), encoding="utf-8").read()
assert "/*SAMPLE_JSON*/null" in tpl
html = tpl.replace("/*SAMPLE_JSON*/null", json.dumps(plan, ensure_ascii=False))
open(os.path.join(ROOT, "index.html"), "w", encoding="utf-8").write(html)
print("wrote index.html", len(html), "bytes")
