"""Builds the performance workbook: funnel and trial quality, pacing, decision engine, reallocation and decision log on top of perf_base.py.
Usage: python3 scripts/build_performance.py <data.xlsx>, then python3 scripts/recalc.py build/Performance_Workbook.xlsx"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import perf_base as m          # builds the base workbook in memory
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment
from openpyxl.styles import Font

wb, put, hdr, section, ST = m.wb, m.put, m.hdr, m.section, m.ST
CALC, INP, RED, AMB, GRN = m.CALC, m.INP, m.RED, m.AMB, m.GRN
MONEY, INT, PCT, DATE = m.MONEY, m.INT, m.PCT, m.DATE
P0, P1, TH, PC, MS, MODE, DATEF = m.P0, m.P1, m.TH, m.PC, m.MS, m.MODE, m.DATEF
OUT = m.OUT

# ---------------- Settings: 7b rows ----------------
ws = wb["Settings"]
new = [("As-of date (pacing)", "=B4", "Defaults to the period end. Set it to today mid-month."),
       ("Days in budget month (Meta)", 31, "Meta paces over the calendar month"),
       ("Google billing days per month", 30.4, "Google's monthly cap = daily budget × 30.4"),
       ("Cohort maturity (days)", 21, "PLACEHOLDER: trial length + payment window. Trials younger than this are excluded from trial-to-paid"),
       ("Average MRR per paid customer", None, "Optional. Fills the payback column"),
       ("Reallocation step (max share per week)", 0.2, "Agreed: move money in steps of at most 20% of a campaign's weekly budget"),
       ("Can absorb: lost IS (budget) above", 0.1, "Google Search losing more than this share of impressions to budget = can take more"),
       ("Can absorb: Meta pace at least", 0.95, "Meta spending ≥95% of its planned budget = budget-limited"),
       ("Poor activation: below × account average", 0.8, "Campaign activation rate under 80% of the account's = quality concern")]
r0 = ws.max_row + 1
for i, (a, b, c) in enumerate(new):
    r = r0 + i
    put(ws, f"A{r}", a, CALC); put(ws, f"B{r}", b, CALC if str(b).startswith("=") else INP); put(ws, f"C{r}", c, CALC, wrap=True)
    ST[a] = f"Settings!$B${r}"
ws[f"B{r0}"].number_format = DATE
for k in ("Reallocation step (max share per week)", "Can absorb: lost IS (budget) above", "Can absorb: Meta pace at least"):
    ws[ST[k].split("!")[1].replace("$", "")].number_format = "0%"
ws[ST["Average MRR per paid customer"].split("!")[1].replace("$", "")].number_format = MONEY
ASOF, MDAYS, GDAYS, MAT, MRR = ST["As-of date (pacing)"], ST["Days in budget month (Meta)"], ST["Google billing days per month"], ST["Cohort maturity (days)"], ST["Average MRR per paid customer"]
STEP, LIS_T, PACE_T, ACT_T = ST["Reallocation step (max share per week)"], ST["Can absorb: lost IS (budget) above"], ST["Can absorb: Meta pace at least"], ST["Poor activation: below × account average"]

# ---------------- IN_Funnel (optional input) ----------------
ws = wb.create_sheet("IN_Funnel"); hdr(ws, 1, ["Date", "utm_source", "Landing page sessions", "Signups started"], [12, 16, 18, 16])
ws["A1"].comment = Comment("Optional. PostHog/GA4 sessions on landing pages by source, per day. Empty in the historical test (no funnel data supplied).", "Workbook")
ws.sheet_properties.tabColor = "F4C430"; ws.freeze_panes = "A2"

# ---------------- Funnel ----------------
ws = wb.create_sheet("Funnel")
put(ws, "A1", "Funnel and trial quality by paid source (PostHog)", bold=True)
put(ws, "A2", f'="Trial-to-paid uses only cohorts at least "&{MAT}&" days old by "&TEXT({ASOF},"d mmm yyyy")&". Younger trials haven\'t had time to pay, so they\'d make quality look worse than it is."')
hdr(ws, 4, ["Source", "utm_source", "Spend", "Landing page sessions", "Session → trial", "Trials", "Activated", "Activation rate",
            "Mature trials", "Paid (mature cohorts)", "Trial-to-paid (mature)", "Immature share of trials", "Paid (all)", "CAC (spend ÷ paid)", "Payback (months)"],
    [12, 12, 12, 14, 11, 9, 10, 11, 11, 12, 12, 12, 9, 12, 11])
ND = '"no data"'
for i, (lab, src, spend) in enumerate([("Meta", ST["PostHog source for Meta"], "Platform!B4"), ("Google", ST["PostHog source for Google"], "Platform!B5")], 5):
    put(ws, f"A{i}", lab, CALC, bold=True); put(ws, f"B{i}", f"={src}", CALC); put(ws, f"C{i}", f"={spend}", CALC, MONEY)
    ph = lambda col, extra="": f'SUMIFS(IN_PostHog!{col}:{col},IN_PostHog!B:B,$B{i},{DATEF("IN_PostHog")}{extra})'
    has = lambda sheet, col: f'COUNTIFS({sheet}!B:B,$B{i},{sheet}!{col}:{col},"<>")>0'
    put(ws, f"D{i}", f'=IF({has("IN_Funnel","C")},SUMIFS(IN_Funnel!C:C,IN_Funnel!B:B,$B{i},{DATEF("IN_Funnel")}),{ND})', CALC, INT)
    put(ws, f"E{i}", f'=IF(ISNUMBER(D{i}),IF(D{i}=0,"",F{i}/D{i}),{ND})', CALC, "0.0%")
    put(ws, f"F{i}", f"={ph('E')}", CALC, INT)
    put(ws, f"G{i}", f'=IF({has("IN_PostHog","F")},{ph("F")},{ND})', CALC, INT)
    put(ws, f"H{i}", f'=IF(ISNUMBER(G{i}),IF(F{i}=0,"",G{i}/F{i}),{ND})', CALC, "0.0%")
    mature = f',IN_PostHog!A:A,"<="&({ASOF}-{MAT})'
    put(ws, f"I{i}", f"={ph('E', mature)}", CALC, INT)
    put(ws, f"J{i}", f'=IF({has("IN_PostHog","G")},{ph("G", mature)},{ND})', CALC, INT)
    put(ws, f"K{i}", f'=IF(ISNUMBER(J{i}),IF(I{i}=0,"no mature cohorts yet",J{i}/I{i}),{ND})', CALC, "0.0%")
    put(ws, f"L{i}", f'=IF(F{i}=0,"",1-I{i}/F{i})', CALC, "0%")
    put(ws, f"M{i}", f'=IF({has("IN_PostHog","G")},{ph("G")},{ND})', CALC, INT)
    put(ws, f"N{i}", f'=IF(ISNUMBER(M{i}),IF(M{i}=0,"no paid yet",C{i}/M{i}),{ND})', CALC, MONEY)
    put(ws, f"O{i}", f'=IF(AND(ISNUMBER(N{i}),N({MRR})>0),N{i}/{MRR},IF(ISNUMBER(N{i}),"add MRR in Settings",{ND}))', CALC, "0.0")
put(ws, "A7", "Total paid", CALC, bold=True)
put(ws, "C7", "=SUM(C5:C6)", CALC, MONEY, bold=True)
for col in "DFGIJM":
    put(ws, f"{col}7", f'=IF(COUNT({col}5:{col}6)=0,{ND},SUM({col}5:{col}6))', CALC, INT, bold=True)
put(ws, "E7", '=IF(ISNUMBER(D7),IF(D7=0,"",F7/D7),"no data")', CALC, "0.0%"); put(ws, "H7", '=IF(ISNUMBER(G7),IF(F7=0,"",G7/F7),"no data")', CALC, "0.0%")
put(ws, "K7", '=IF(ISNUMBER(J7),IF(I7=0,"no mature cohorts yet",J7/I7),"no data")', CALC, "0.0%"); put(ws, "L7", '=IF(F7=0,"",1-I7/F7)', CALC, "0%")
put(ws, "N7", '=IF(ISNUMBER(M7),IF(M7=0,"no paid yet",C7/M7),"no data")', CALC, MONEY)
put(ws, "O7", f'=IF(AND(ISNUMBER(N7),N({MRR})>0),N7/{MRR},IF(ISNUMBER(N7),"add MRR in Settings","no data"))', CALC, "0.0")
put(ws, "A9", "'no data' = that column wasn't supplied (Activated / Paid in IN_PostHog, sessions in IN_Funnel). The historical test has neither, so quality can't be judged from it.")
ws.conditional_formatting.add("D5:O7", FormulaRule(formula=['D5="no data"'], fill=AMB))

# ---------------- Pacing ----------------
ws = wb.create_sheet("Pacing")
put(ws, "A1", "Budget pacing (planned vs actual)", bold=True)
put(ws, "A2", f'="As of "&TEXT({ASOF},"d mmm yyyy")&": day "&(MIN({ASOF},{P1})-{P0}+1)&" of the budget month. Meta paces over "&{MDAYS}&" days, Google over "&{GDAYS}&" (its monthly cap)."')
hdr(ws, 4, ["Platform", "Campaign", "Planned monthly", "Spend to date", "Expected by now", "Pace", "Projected month", "Remaining", "Daily needed", "Status"],
    [9, 28, 13, 13, 13, 8, 14, 12, 12, 16])
for k in range(30):
    r, cr = 5 + k, 5 + k
    put(ws, f"A{r}", f'=Campaigns!A{cr}', CALC); put(ws, f"B{r}", f'=Campaigns!B{cr}', CALC)
    put(ws, f"C{r}", f'=IF(OR($B{r}="",IN_Plan!C{k+2}=""),"",IN_Plan!C{k+2})', CALC, MONEY)
    upto = lambda sh: f'SUMIFS({sh}!C:C,{sh}!$B:$B,$B{r},{sh}!$A:$A,">="&{P0},{sh}!$A:$A,"<="&MIN({ASOF},{P1}))'
    put(ws, f"D{r}", f'=IF($B{r}="","",IF($A{r}="Meta",{upto("IN_Meta_Campaign")},{upto("IN_Google")}))', CALC, MONEY)
    DM, EL = f'IF($A{r}="Google",{GDAYS},{MDAYS})', f'(MIN({ASOF},{P1})-{P0}+1)'
    put(ws, f"E{r}", f'=IF(N(C{r})=0,"",C{r}/{DM}*{EL})', CALC, MONEY)
    put(ws, f"F{r}", f'=IF(N(E{r})=0,"",D{r}/E{r})', CALC, "0%")
    put(ws, f"G{r}", f'=IF($B{r}="","",D{r}/{EL}*{DM})', CALC, MONEY)
    put(ws, f"H{r}", f'=IF(N(C{r})=0,"",C{r}-D{r})', CALC, MONEY)
    put(ws, f"I{r}", f'=IF(N(C{r})=0,"",IF({DM}-{EL}<=0,"month over",MAX(0,H{r})/({DM}-{EL})))', CALC, MONEY)
    put(ws, f"J{r}", f'=IF($B{r}="","",IF(N(C{r})=0,"No plan",IF(F{r}<0.9,"Under-pacing",IF(F{r}>1.1,"Over-pacing","On pace"))))', CALC)
ws.conditional_formatting.add("J5:J34", FormulaRule(formula=['OR(J5="Under-pacing",J5="Over-pacing")'], fill=AMB))
ws.conditional_formatting.add("J5:J34", FormulaRule(formula=['J5="On pace"'], fill=GRN))
put(ws, "A36", "Projected month = spend so far ÷ days elapsed × days in the budget month. Google can overspend on single days but not past daily budget × 30.4 for the month.")

# ---------------- Decisions (engine) ----------------
ws = wb.create_sheet("Decisions")
put(ws, "A1", "Recommended actions: rules, not autopilot. A person decides; record it in Decision_Log.", bold=True)
put(ws, "A2", f'="Rules: scale only if even the worst-case CPA ≤ $"&{TH}&", quality is OK and the campaign can absorb more · reduce/refresh if even the best case > $"&{TH}&" and worsening · pause candidate at 0 trials after "&{MS}&"× planning CPA · otherwise inconclusive."')
hdr(ws, 4, ["Platform", "Campaign", "Spend", "Trials", "CPA", "Best case", "Worst case", "Trend", "Can absorb more?", "Trial quality",
            "Recommendation", "Why", "Weekly budget", "Movable / week", "Receiver"],
    [9, 28, 12, 8, 10, 10, 12, 11, 30, 22, 26, 70, 12, 12, 14])
for k in range(30):
    r, cr = 5 + k, 5 + k
    for col, src in (("A", "A"), ("B", "B"), ("C", "C"), ("D", "H"), ("E", "J"), ("F", "K"), ("G", "L"), ("H", "Q")):
        put(ws, f"{col}{r}", f'=IF(Campaigns!$B{cr}="","",Campaigns!{src}{cr})', CALC, MONEY if col in "CEFG" else INT if col == "D" else None)
    lis = f'AVERAGEIFS(IN_Google!G:G,IN_Google!B:B,$B{r},{DATEF("IN_Google")})'
    put(ws, f"I{r}", (f'=IF($B{r}="","",IF($A{r}="Google",IF(COUNTIFS(IN_Google!B:B,$B{r},IN_Google!G:G,"<>")=0,"Unknown – add lost IS (budget)",'
                      f'IFERROR(IF({lis}>{LIS_T},"Yes – losing "&TEXT({lis},"0%")&" of impressions to budget","No – not budget-limited"),"Unknown – no data in period")),'
                      f'IF(NOT(ISNUMBER(Pacing!F{r})),"Unknown – no planned budget",IF(Pacing!F{r}>={PACE_T},"Likely – spending its full budget","No – not spending its full budget"))))'), CALC)
    act = f'SUMIFS(IN_PostHog!F:F,IN_PostHog!C:C,$B{r},{DATEF("IN_PostHog")})/SUMIFS(IN_PostHog!E:E,IN_PostHog!C:C,$B{r},{DATEF("IN_PostHog")})'
    put(ws, f"J{r}", (f'=IF($B{r}="","",IF(COUNTIFS(IN_PostHog!C:C,$B{r},IN_PostHog!F:F,"<>")=0,"n/a – no activation by campaign",'
                      f'IF(NOT(ISNUMBER(Funnel!H7)),"n/a",IF(IFERROR({act},0)<{ACT_T}*Funnel!H7,"Poor","OK"))))'), CALC)
    rec = (f'=IF($B{r}="","",IF(N(C{r})=0,"No spend",IF(N(D{r})=0,IF(C{r}>={MS}*{PC},"Pause candidate","Inconclusive – keep running"),'
           f'IF(AND(ISNUMBER(G{r}),G{r}<={TH}),IF(J{r}="Poor","Investigate trial quality",IF(OR(LEFT(I{r},3)="Yes",LEFT(I{r},6)="Likely"),"Scale candidate","Maintain")),'
           f'IF(F{r}>{TH},IF(H{r}="Worsening","Reduce or refresh","Investigate – above target"),IF(E{r}<={TH},"Maintain","Inconclusive – keep running"))))))')
    put(ws, f"K{r}", rec, CALC, bold=True)
    src = f'IF(LEFT(Campaigns!I{cr},4)="Meta"," Trials are Meta-reported"&IF(COUNTIF(Platform!F11:F26,"Diverging")>0," while Meta tracking diverges, so the true CPA is probably lower."," (platform-attributed)."),"")'
    rng_txt = f'TEXT(F{r},"$#,##0")&"–"&IF(ISNUMBER(G{r}),TEXT(G{r},"$#,##0"),"?")'
    parts = [
        ("Pause candidate", f'TEXT(C{r},"$#,##0")&" spent, no trials (≥ "&{MS}&"× planning CPA). Even at best the CPA is ≥ "&TEXT(F{r},"$#,##0")&". Check tracking and conversion lag before pausing."'),
        ("Scale candidate", f'"Even the worst-case CPA ("&TEXT(G{r},"$#,##0")&") is within target. "&I{r}&". Step up by ≤"&TEXT({STEP},"0%")&" a week."'),
        ("Maintain", f'IF(AND(ISNUMBER(G{r}),G{r}<={TH}),"Within target even at worst ("&TEXT(G{r},"$#,##0")&"), but no proven headroom: "&I{r}&".","CPA "&TEXT(E{r},"$#,##0")&" is on target but the range ("&{rng_txt}&") is still wide.")'),
        ("Reduce or refresh", f'"Even the best-case CPA ("&TEXT(F{r},"$#,##0")&") is above target, and it is worsening."'),
        ("Investigate – above target", f'"Best-case CPA "&TEXT(F{r},"$#,##0")&" is above target; trend "&LOWER(H{r})&". Diagnose before cutting."'),
        ("Investigate trial quality", '"Cheap trials but activation is below the account average."'),
        ("Inconclusive – keep running", f'IF(N(D{r})=0,"Spend is below the minimum to judge zero trials.","Range "&{rng_txt}&" straddles $"&{TH}&". A low-volume campaign isn\'t a loser.")'),
    ]
    why = f'=IF($B{r}="","",' + "&".join(f'IF(K{r}="{k}",{v},"")' for k, v in parts) + f'&{src})'
    put(ws, f"L{r}", why, CALC, wrap=True)
    put(ws, f"M{r}", f'=IF($B{r}="","",IF(N(Pacing!C{r})>0,Pacing!C{r}/IF($A{r}="Google",{GDAYS},{MDAYS})*7,C{r}/(MIN({ASOF},{P1})-{P0}+1)*7))', CALC, MONEY)
    put(ws, f"N{r}", f'=IF($B{r}="","",IF(K{r}="Pause candidate",M{r},IF(K{r}="Reduce or refresh",M{r}*{STEP},0)))', CALC, MONEY)
    put(ws, f"O{r}", f'=IF(K{r}="Scale candidate",B{r},"")', CALC)
for col_rule, fill in (('$K5="Scale candidate"', GRN), ('OR($K5="Reduce or refresh",$K5="Pause candidate")', RED), ('LEFT($K5,11)="Investigate"', AMB)):
    ws.conditional_formatting.add("K5:K34", FormulaRule(formula=[col_rule], fill=fill))

R = 37
section(ws, f"A{R}", "Budget reallocation (suggested, in steps)")
put(ws, f"A{R+1}", "Movable per week", CALC, bold=True); put(ws, f"C{R+1}", "=SUM(N5:N34)", CALC, MONEY, bold=True)
put(ws, f"A{R+2}", "Campaigns that can take it", CALC, bold=True); put(ws, f"C{R+2}", '=COUNTIF(K5:K34,"Scale candidate")', CALC, INT)
put(ws, f"A{R+3}", "Suggestion", CALC, bold=True)
put(ws, f"C{R+3}", (f'=IF(C{R+1}=0,"No reallocation suggested.",IF(C{R+2}=0,"Free up "&TEXT(C{R+1},"$#,##0")&"/week from the red rows. No campaign has proven headroom yet, so hold it in the reserve until one does (add lost IS / planned budgets to test headroom).",'
                    f'"Move up to "&TEXT(C{R+1},"$#,##0")&"/week, in ≤"&TEXT({STEP},"0%")&" weekly steps, toward: "&_xlfn.TEXTJOIN(", ",TRUE,O5:O34)&". Re-check CPA ranges each week; the last dollars cost more than the average."))'), CALC, wrap=True)
ws.merge_cells(f"C{R+3}:L{R+3}"); ws.row_dimensions[R + 3].height = 45
put(ws, f"A{R+4}", "Pause candidates move all of their weekly budget once tracking is confirmed; reduce/refresh move one step. The dashboard never pauses anything itself.")

C0 = R + 7
section(ws, f"A{C0}", "Concepts")
hdr(ws, C0 + 1, ["Concept", "", "Spend", "Trials", "CPA", "Best case", "Worst case", "", "", "", "Recommendation", "Why"])
for k in range(12):
    r, cr = C0 + 2 + k, 50 + k
    put(ws, f"A{r}", f'=IF(Creative!A{cr}="","",Creative!A{cr})', CALC)
    for col, src in (("C", "C"), ("D", "I"), ("E", "J"), ("F", "K"), ("G", "L")):
        put(ws, f"{col}{r}", f'=IF($A{r}="","",Creative!{src}{cr})', CALC, MONEY if col != "D" else INT)
    put(ws, f"K{r}", (f'=IF($A{r}="","",IF(N(C{r})=0,"Not launched",IF(N(D{r})=0,IF(C{r}>={MS}*{PC},"Pause candidate","Keep running"),'
                      f'IF(AND(ISNUMBER(G{r}),G{r}<={TH}),"Winner – scale into new executions",IF(F{r}>{TH},"Retire or rework","Inconclusive – keep testing")))))'), CALC, bold=True)
    put(ws, f"L{r}", (f'=IF(OR($A{r}="",K{r}="Not launched"),"",IF(LEFT(K{r},6)="Winner","Within target even at worst ("&TEXT(G{r},"$#,##0")&"). Make new executions and hooks of it.",'
                      f'IF(K{r}="Retire or rework","Above target even at best ("&TEXT(F{r},"$#,##0")&").",IF(LEFT(K{r},12)="Inconclusive","Range "&TEXT(F{r},"$#,##0")&"–"&IF(ISNUMBER(G{r}),TEXT(G{r},"$#,##0"),"?")&" straddles target.","Too little spend to judge.")))&'
                      f'IF(COUNTA(IN_PostHog!D2:D20000)=0," Meta-reported trials.",""))'), CALC, wrap=True)
ws.conditional_formatting.add(f"K{C0+2}:K{C0+13}", FormulaRule(formula=[f'LEFT($K{C0+2},6)="Winner"'], fill=GRN))
ws.conditional_formatting.add(f"K{C0+2}:K{C0+13}", FormulaRule(formula=[f'OR($K{C0+2}="Retire or rework",$K{C0+2}="Pause candidate")'], fill=RED))
ws.freeze_panes = "C5"

# ---------------- Decision_Log ----------------
ws = wb.create_sheet("Decision_Log")
put(ws, "A1", "Decision log: the human decides. Every override needs a reason.", bold=True)
hdr(ws, 3, ["Date", "Level", "Entity (exact name)", "Recommendation now", "Decision", "Follows or overrides?", "Reason (required for overrides)",
            "Evidence at the time (type it: numbers change)", "Owner", "Next review", "Outcome"], [12, 10, 28, 28, 12, 22, 40, 44, 12, 12, 30])
for r in range(4, 204):
    put(ws, f"D{r}", f'=IF(C{r}="","",IFERROR(INDEX(Decisions!K:K,MATCH(C{r},Decisions!B:B,0)),IFERROR(INDEX(Decisions!K:K,MATCH(C{r},Decisions!A:A,0)),"not found – check the name")))', CALC)
    put(ws, f"F{r}", f'=IF(OR(C{r}="",E{r}=""),"",IF(ISNUMBER(SEARCH(E{r},D{r})),"Follows","OVERRIDE – give a reason"))', CALC)
    for col in "ABCEGHIJK":
        ws[f"{col}{r}"].fill = INP; ws[f"{col}{r}"].font = Font(name=m.F)
    ws[f"A{r}"].number_format = ws[f"J{r}"].number_format = DATE
import datetime as _dt
ex = [_dt.date(2026,10,12), "Campaign", "Search_NonBrand_UK", None, "Pause", None, "", "EXAMPLE: $15,090 spent, 0 trials; best-case CPA ≥ $5,037. Tracking checked: UK conversion tag fires.", "Owner", _dt.date(2026,10,19), ""]
for col, v in zip("ABCDEFGHIJK", ex):
    if v is not None: ws[f"{col}4"] = v
dv = DataValidation(type="list", formula1='"Scale,Maintain,Keep,Investigate,Refresh,Reduce,Retire,Pause"', allow_blank=True); ws.add_data_validation(dv); dv.add("E4:E203")
dv2 = DataValidation(type="list", formula1='"Campaign,Concept,Ad"', allow_blank=True); ws.add_data_validation(dv2); dv2.add("B4:B203")
ws.conditional_formatting.add("G4:G203", FormulaRule(formula=['AND(LEFT($F4,8)="OVERRIDE",$G4="")'], fill=RED))
ws.freeze_panes = "D4"

# ---------------- Summary additions ----------------
ws = wb["Summary"]
ws["A20"] = "Yellow tabs (IN_…) are inputs: paste exports there. Grey cells are formulas. Decisions are recommendations; record what you decide in Decision_Log."
put(ws, "A22", "Decisions and quality", m.SEC, bold=True)
add = [("Scale candidates", '=COUNTIF(Decisions!K5:K34,"Scale candidate")', INT, "Even worst-case CPA within target and able to absorb more"),
       ("Reduce / refresh", '=COUNTIF(Decisions!K5:K34,"Reduce or refresh")', INT, "Even best-case CPA above target and worsening"),
       ("Pause candidates", '=COUNTIF(Decisions!K5:K34,"Pause candidate")', INT, "No trials after meaningful spend: check tracking first"),
       ("Investigate", '=COUNTIF(Decisions!K5:K34,"Investigate*")', INT, ""),
       ("Inconclusive", '=COUNTIF(Decisions!K5:K34,"Inconclusive*")', INT, "A low-volume campaign isn't a loser"),
       ("Reallocation", "=Decisions!C40", None, ""),
       ("Activation rate (paid)", "=Funnel!H7", "0.0%", "PostHog Activated ÷ trials"),
       ("Trial-to-paid (mature cohorts)", "=Funnel!K7", "0.0%", "Only cohorts old enough to have paid"),
       ("CAC (spend ÷ paid customers)", "=Funnel!N7", MONEY, ""),
       ("Overrides without a reason", '=COUNTIFS(Decision_Log!F4:F203,"OVERRIDE*",Decision_Log!G4:G203,"")', INT, "Should be 0")]
for i, (a, b, f, d) in enumerate(add, 23):
    put(ws, f"A{i}", a, CALC, bold=True); put(ws, f"B{i}", b, CALC, f); put(ws, f"D{i}", d, CALC, wrap=True)
ws.merge_cells("B28:D28"); from openpyxl.styles import Alignment; ws["B28"].alignment = Alignment(wrap_text=True, vertical="top"); ws.row_dimensions[28].height = 48
ws.conditional_formatting.add("B32", FormulaRule(formula=["B32>0"], fill=RED))

order = ["Summary", "Decisions", "Platform", "Campaigns", "Creative", "Funnel", "Pacing", "Decision_Log", "Settings",
         "IN_Plan", "IN_Meta_Campaign", "IN_Meta_Ads", "IN_Google", "IN_PostHog", "IN_Funnel", "IN_AdMap"]
wb._sheets = [wb[n] for n in order]
wb.save(OUT); print("saved", OUT)
