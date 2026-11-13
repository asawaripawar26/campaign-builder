"""Base of the performance workbook (inputs, summary, platform, campaigns, creative). Imported by build_performance.py."""
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule, CellIsRule

import os, sys
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# The source data: an .xlsx with tabs Meta_Daily, Meta_Ads_Weekly, Google_Daily, Trials_PostHog, Change_Log.
# The assessment data file is not included in this repo; pass your own exports in the same layout.
DATA = os.environ.get("PERF_DATA") or (sys.argv[1] if len(sys.argv) > 1 else sys.exit("Usage: python3 scripts/build_performance.py <data.xlsx>"))
OUT = os.path.join(ROOT, "build", "Performance_Workbook.xlsx"); os.makedirs(os.path.dirname(OUT), exist_ok=True)
F = "Arial"
HDR = PatternFill("solid", fgColor="1F2A44"); INP = PatternFill("solid", fgColor="FFF5CC"); CALC = PatternFill("solid", fgColor="E7E9EE")
SEC = PatternFill("solid", fgColor="D9E1F2"); RED = PatternFill("solid", fgColor="F8C9C4"); AMB = PatternFill("solid", fgColor="FFE7A8"); GRN = PatternFill("solid", fgColor="CDEBD6")
MONEY, INT, PCT, DATE = "$#,##0;($#,##0);-", "#,##0;(#,##0);-", "0.0%", "yyyy-mm-dd"
wb = Workbook(); wb.remove(wb.active)
x = pd.read_excel(DATA, sheet_name=None)


def hdr(ws, row, labels, widths=None):
    for i, l in enumerate(labels, 1):
        c = ws.cell(row, i, l); c.font = Font(name=F, bold=True, color="FFFFFF"); c.fill = HDR
        c.alignment = Alignment(wrap_text=True, vertical="center")
    if widths:
        for i, w in enumerate(widths, 1): ws.column_dimensions[ws.cell(1, i).column_letter].width = w
    ws.row_dimensions[row].height = 30


def put(ws, ref, v, fill=None, fmt=None, bold=False, wrap=False):
    c = ws[ref]; c.value = v; c.font = Font(name=F, bold=bold)
    if fill: c.fill = fill
    if fmt: c.number_format = fmt
    if wrap: c.alignment = Alignment(wrap_text=True, vertical="top")
    return c


def section(ws, ref, text):
    c = put(ws, ref, text, SEC, bold=True)


def input_tab(name, cols, widths, df, note):
    ws = wb.create_sheet(name); hdr(ws, 1, cols, widths); ws.freeze_panes = "A2"
    ws["A1"].comment = Comment(note, "Workbook")
    for r in df.itertuples(index=False):
        ws.append(list(r))
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for c in row:
            c.font = Font(name=F); c.fill = INP
            if isinstance(c.value, (pd.Timestamp,)) or "Date" in cols[c.column - 1] or "Week" in cols[c.column - 1]: c.number_format = DATE
    return ws


# ---------------- inputs (history as test data) ----------------
md, ma, gd, ph = x["Meta_Daily"], x["Meta_Ads_Weekly"], x["Google_Daily"], x["Trials_PostHog"]
input_tab("IN_Meta_Campaign", ["Date", "Campaign", "Spend", "Impressions", "Clicks", "Trials (Meta-reported)"], [12, 30, 12, 13, 10, 14],
          md[["Date", "Campaign", "Spend ($)", "Impressions", "Clicks", "Trials (Meta reported)"]].assign(Date=lambda d: d.Date.dt.date),
          "Paste Ads Manager → Campaigns, broken down by Day. One row per campaign per day.")
input_tab("IN_Meta_Ads", ["Date (or week start)", "Campaign", "Ad name", "Spend", "Impressions", "Clicks", "Trials (Meta-reported)", "Frequency"],
          [14, 28, 30, 12, 13, 10, 14, 10],
          ma[["Week starting", "Campaign", "Ad name", "Spend ($)", "Impressions", "Clicks", "Trials (Meta reported)", "Frequency"]].assign(**{"Week starting": lambda d: d["Week starting"].dt.date}),
          "Paste Ads Manager → Ads, by Day or Week. Ad names link to IN_AdMap.")
input_tab("IN_Google", ["Date", "Campaign", "Spend", "Impressions", "Clicks", "Conversions (Google-reported)", "Search lost IS (budget)"],
          [12, 26, 12, 13, 10, 16, 16],
          gd[["Date", "Campaign", "Spend ($)", "Impressions", "Clicks", "Conversions (Google reported)"]].assign(Date=lambda d: d.Date.dt.date, LIS=None),
          "Paste Google Ads → Campaigns by Day. 'Search lost IS (budget)' is optional (used in 7b to judge whether a campaign can absorb more budget).")
input_tab("IN_PostHog", ["Date", "utm_source", "utm_campaign", "utm_content", "Trials", "Activated", "Paid"], [12, 16, 30, 30, 9, 10, 8],
          ph[["Date", "Source (utm_source)"]].assign(Date=lambda d: d.Date.dt.date, c=None, a=None, T=ph["Trial signups"], act=None, paid=None),
          "Paste PostHog trial signups. utm_campaign / utm_content are optional: when present, campaign and ad trials come from PostHog instead of the platforms.")
plan = pd.DataFrame([["Meta", "Prospecting_AdvantagePlus"], ["Meta", "Prospecting_Broad"], ["Meta", "Retargeting_30d"],
                     ["Google", "Search_Brand_US"], ["Google", "Search_NonBrand_US"], ["Google", "PMax_US"], ["Google", "Search_NonBrand_UK"]], columns=["p", "c"])
plan["planned"] = None
input_tab("IN_Plan", ["Platform", "Campaign name (exact)", "Planned monthly budget"], [10, 34, 16], plan,
          "One row per campaign you want reported, named exactly as in the platform. Copy from the plan's Campaigns tab. Planned budget feeds pacing (7b). Historical test: names from the Aug–Oct data; no plan.")
admap = pd.DataFrame([["UGC_Founder_v2", "FOUNDER_STORY", "v2", "HA"], ["UGC_Agency_Owner", "AGENCY_OWNER", "v1", "HA"],
                      ["Static_AgentStack_Benefits", "BENEFITS_STATIC", "v1", "HA"], ["Demo_Screen_Walkthrough", "DEMO_WALKTHROUGH", "v1", "HA"],
                      ["Carousel_Case_Study", "CASE_STUDY", "v1", "HA"], ["Retargeting_Testimonial", "TESTIMONIAL", "v1", "HA"]],
                     columns=["ad", "con", "exe", "hook"])
input_tab("IN_AdMap", ["Ad name (exact)", "Concept ID", "Execution", "Hook code"], [30, 22, 10, 10], admap,
          "One row per ad. From the plan's Meta_Ads (incl. 'Previous ad name') and the generator's test ads. Historical test: old ad names mapped to concepts.")

# ---------------- Settings ----------------
ws = wb.create_sheet("Settings", 0); hdr(ws, 1, ["Setting", "Value", "Note"], [36, 22, 70])
S = [("Report name", "TEST: Aug 10 – Oct 11, 2026 history", "Change for each live report"),
     ("Period start", pd.Timestamp("2026-08-10").date(), ""), ("Period end", pd.Timestamp("2026-10-11").date(), ""),
     ("Days in period", "=B4-B3+1", "Calculated"),
     ("Target CPA – low", 250, "Brief"), ("Target CPA – high", 280, "Brief. Decisions compare against this upper bound"),
     ("Planning CPA", "=AVERAGE(B6:B7)", "Calculated"),
     ("Min trials for sufficient evidence", 20, "Agreed threshold (a volume threshold, not significance)"),
     ("Early-signal trials", 5, "From this many trials a read is 'early signal'"),
     ("Min spend = × planning CPA", 3, "Zero trials below this spend can easily be chance"),
     ("Confidence level for CPA ranges", 0.9, "90%: the CPA range covers the true CPA 9 times in 10"),
     ("PostHog source for Meta", "facebook", "utm_source value"), ("PostHog source for Google", "google", "utm_source value"),
     ("Trial attribution available", '=IF(COUNTA(IN_PostHog!C2:C20000)+COUNTA(IN_PostHog!D2:D20000)>0,"PostHog by campaign/ad (utm)","Platform-reported only")',
      "Auto: PostHog utm_campaign/utm_content present? If not, campaign and ad trials are labelled platform-attributed"),
     ("Tracking ratio – healthy band low", 0.8, "Platform-reported ÷ PostHog trials"), ("Tracking ratio – healthy band high", 1.25, ""),
     ("Baseline: spend", 249298, "Aug 10 – Oct 11, 2026 (63 days). From the audit"),
     ("Baseline: PostHog paid trials", 879, "Same 63 days (facebook + google)"),
     ("Baseline: blended CPA", "=B18/B19", "Same 63 days"),
     ("Baseline: paid trials per month", "=B19/63*30.4", "63-day total converted to a 30.4-day month")]
for i, (a, b, c) in enumerate(S, 2):
    put(ws, f"A{i}", a, CALC); put(ws, f"B{i}", b, CALC if str(b).startswith("=") else INP); put(ws, f"C{i}", c, CALC, wrap=True)
for r in (3, 4): ws[f"B{r}"].number_format = DATE
for r in (6, 7, 8, 18, 20): ws[f"B{r}"].number_format = MONEY
ws["B12"].number_format = "0%"; ws["B21"].number_format = INT; ws["B19"].number_format = INT
ST = {k: f"Settings!$B${i}" for i, (k, _, _) in enumerate(S, 2)}
P0, P1, TH, PC, MINT, EST, MS, CONF = ST["Period start"], ST["Period end"], ST["Target CPA – high"], ST["Planning CPA"], ST["Min trials for sufficient evidence"], ST["Early-signal trials"], ST["Min spend = × planning CPA"], ST["Confidence level for CPA ranges"]
MODE = ST["Trial attribution available"]
IN_P = lambda col: f"IN_PostHog!${col}:${col}"
DATEF = lambda sheet, col="A": f'{sheet}!${col}:${col},">="&{P0},{sheet}!${col}:${col},"<="&{P1}'


def ci(spend, trials):
    """CPA range from a 90% Poisson interval on the trial count (exact, chi-square)."""
    lo = f'=IF({spend}=0,"",{spend}/(CHIINV((1-{CONF})/2,2*{trials}+2)/2))'
    hi = f'=IF({spend}=0,"",IF({trials}=0,"no upper bound",{spend}/(CHIINV(1-(1-{CONF})/2,2*{trials})/2)))'
    return lo, hi


def stage(spend, trials):
    return (f'=IF({spend}=0,"Not launched",IF({trials}>={MINT},"Sufficient evidence",IF({spend}<{MS}*{PC},"Launched – below min spend",'
            f'IF({trials}>={EST},"Early signal","Min spend reached"))))')


def vs_target(trials, lo, hi):
    return f'=IF(OR({trials}="",{lo}=""),"",IF({trials}=0,"No trials",IF(AND(ISNUMBER({hi}),{hi}<={TH}),"Within target even at worst",IF({lo}>{TH},"Above target even at best","Straddles target – inconclusive"))))'


def colour_status(ws, rng):
    first = rng.split(":")[0]
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'ISNUMBER(SEARCH("within",{first}))'], fill=GRN))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'ISNUMBER(SEARCH("above target",{first}))'], fill=RED))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'OR(ISNUMBER(SEARCH("straddles",{first})),ISNUMBER(SEARCH("No trials",{first})))'], fill=AMB))


# ---------------- Platform ----------------
ws = wb.create_sheet("Platform", 1)
put(ws, "A1", "Platform performance: platform-reported conversions vs PostHog trials (never added together)", bold=True)
hdr(ws, 3, ["Platform", "Spend", "Platform-reported conversions", "PostHog trials (source)", "Platform CPA", "Source CPA (PostHog)",
            "Source CPA – best case", "Source CPA – worst case", "vs target", "Platform ÷ PostHog", "Tracking health"],
    [26, 12, 14, 14, 12, 14, 13, 13, 26, 12, 22])
rows = [("Meta", f"SUMIFS(IN_Meta_Campaign!C:C,{DATEF('IN_Meta_Campaign')})", f"SUMIFS(IN_Meta_Campaign!F:F,{DATEF('IN_Meta_Campaign')})", ST["PostHog source for Meta"]),
        ("Google", f"SUMIFS(IN_Google!C:C,{DATEF('IN_Google')})", f"SUMIFS(IN_Google!F:F,{DATEF('IN_Google')})", ST["PostHog source for Google"])]
for i, (pl, sp, cv, src) in enumerate(rows, 4):
    put(ws, f"A{i}", pl, CALC, bold=True); put(ws, f"B{i}", "=" + sp, CALC, MONEY); put(ws, f"C{i}", "=" + cv, CALC, INT)
    put(ws, f"D{i}", f"=SUMIFS(IN_PostHog!E:E,IN_PostHog!B:B,{src},{DATEF('IN_PostHog')})", CALC, INT)
    put(ws, f"E{i}", f'=IF(C{i}=0,"",B{i}/C{i})', CALC, MONEY); put(ws, f"F{i}", f'=IF(D{i}=0,"",B{i}/D{i})', CALC, MONEY)
    lo, hi = ci(f"B{i}", f"D{i}"); put(ws, f"G{i}", lo, CALC, MONEY); put(ws, f"H{i}", hi, CALC, MONEY)
    put(ws, f"I{i}", vs_target(f"D{i}", f"G{i}", f"H{i}"), CALC)
    put(ws, f"J{i}", f'=IF(D{i}=0,"",C{i}/D{i})', CALC, "0.00")
    put(ws, f"K{i}", f'=IF(J{i}="","",IF(AND(J{i}>={ST["Tracking ratio – healthy band low"]},J{i}<={ST["Tracking ratio – healthy band high"]}),"Consistent","Diverging – check tracking"))', CALC)
put(ws, "A6", "Total paid", CALC, bold=True)
for col, f in (("B", MONEY), ("C", INT), ("D", INT)):
    put(ws, f"{col}6", f"=SUM({col}4:{col}5)", CALC, f, bold=True)
put(ws, "F6", '=IF(D6=0,"",B6/D6)', CALC, MONEY, bold=True); lo, hi = ci("B6", "D6"); put(ws, "G6", lo, CALC, MONEY); put(ws, "H6", hi, CALC, MONEY)
put(ws, "I6", vs_target("D6", "G6", "H6"), CALC)
colour_status(ws, "I4:I6")
ws.conditional_formatting.add("K4:K5", FormulaRule(formula=['ISNUMBER(SEARCH("Diverging",K4))'], fill=RED))

section(ws, "A9", "Weekly: does each platform's count still match PostHog? (a divergence means tracking, not performance, changed)")
hdr(ws, 10, ["Week starting", "Meta spend", "Meta-reported", "PostHog facebook", "Meta ÷ PostHog", "Meta health",
             "Google spend", "Google-reported", "PostHog google", "Google ÷ PostHog", "Google health"])
for k in range(16):
    r = 11 + k
    put(ws, f"A{r}", f"={P0}-WEEKDAY({P0},3)" if k == 0 else f'=IF(A{r-1}="","",IF(A{r-1}+7>{P1},"",A{r-1}+7))', CALC, DATE)
    wk = lambda sheet, col: f'{sheet}!$A:$A,">="&$A{r},{sheet}!$A:$A,"<"&$A{r}+7'
    for col, f in (("B", f"SUMIFS(IN_Meta_Campaign!C:C,{wk('IN_Meta_Campaign','A')})"), ("C", f"SUMIFS(IN_Meta_Campaign!F:F,{wk('IN_Meta_Campaign','A')})"),
                   ("D", f"SUMIFS(IN_PostHog!E:E,IN_PostHog!B:B,{ST['PostHog source for Meta']},{wk('IN_PostHog','A')})"),
                   ("G", f"SUMIFS(IN_Google!C:C,{wk('IN_Google','A')})"), ("H", f"SUMIFS(IN_Google!F:F,{wk('IN_Google','A')})"),
                   ("I", f"SUMIFS(IN_PostHog!E:E,IN_PostHog!B:B,{ST['PostHog source for Google']},{wk('IN_PostHog','A')})")):
        put(ws, f"{col}{r}", f'=IF($A{r}="","",{f})', CALC, MONEY if col in "BG" else INT)
    for a, b, c, d in (("E", "C", "D", "F"), ("J", "H", "I", "K")):
        put(ws, f"{a}{r}", f'=IF(OR($A{r}="",N({c}{r})=0),"",{b}{r}/{c}{r})', CALC, "0.00")
        put(ws, f"{d}{r}", f'=IF({a}{r}="","",IF(AND({a}{r}>={ST["Tracking ratio – healthy band low"]},{a}{r}<={ST["Tracking ratio – healthy band high"]}),"Consistent","Diverging"))', CALC)
ws.conditional_formatting.add("F11:F26", FormulaRule(formula=['F11="Diverging"'], fill=RED))
ws.conditional_formatting.add("K11:K26", FormulaRule(formula=['K11="Diverging"'], fill=RED))
put(ws, "A28", "Note: weekly ratios swing with small numbers. One diverging week is a prompt to check, not a verdict; a run of them (e.g. Meta from late September) is a tracking problem.", wrap=False)

# ---------------- Campaigns ----------------
ws = wb.create_sheet("Campaigns", 2)
put(ws, "A1", "Campaign performance", bold=True)
put(ws, "A2", f'="Trials column source: "&IF({MODE}="Platform-reported only","PLATFORM-REPORTED (PostHog has no utm_campaign). Meta undercounted after Sep 22; read Meta campaign CPAs with care.","PostHog utm_campaign")')
cols = ["Platform", "Campaign", "Spend", "Impressions", "Clicks", "CTR", "CPC", "Trials", "Trial source", "CPA", "CPA – best case", "CPA – worst case",
        "Evidence stage", "vs target", "CPA last 14 days", "CPA prior 14 days", "Trend"]
hdr(ws, 4, cols, [9, 28, 12, 13, 10, 8, 8, 9, 16, 10, 12, 13, 24, 28, 12, 12, 12])
N = 30
for k in range(N):
    r = 5 + k; pr = 2 + k
    put(ws, f"A{r}", f'=IF(IN_Plan!A{pr}="","",IN_Plan!A{pr})', CALC); put(ws, f"B{r}", f'=IF(IN_Plan!B{pr}="","",IN_Plan!B{pr})', CALC)
    isM = f'$A{r}="Meta"'
    m = lambda col, extra="": f'SUMIFS(IN_Meta_Campaign!{col}:{col},IN_Meta_Campaign!$B:$B,$B{r},{DATEF("IN_Meta_Campaign")}{extra})'
    g = lambda col, extra="": f'SUMIFS(IN_Google!{col}:{col},IN_Google!$B:$B,$B{r},{DATEF("IN_Google")}{extra})'
    pick = lambda mc, gc: f'=IF($B{r}="","",IF({isM},{m(mc)},{g(gc)}))'
    put(ws, f"C{r}", pick("C", "C"), CALC, MONEY); put(ws, f"D{r}", pick("D", "D"), CALC, INT); put(ws, f"E{r}", pick("E", "E"), CALC, INT)
    put(ws, f"F{r}", f'=IF(N(D{r})=0,"",E{r}/D{r})', CALC, "0.00%"); put(ws, f"G{r}", f'=IF(N(E{r})=0,"",C{r}/E{r})', CALC, "$0.00")
    plat_tr = f'IF({isM},{m("F")},{g("F")})'
    ph_tr = f'SUMIFS(IN_PostHog!E:E,IN_PostHog!C:C,$B{r},{DATEF("IN_PostHog")})'
    put(ws, f"H{r}", f'=IF($B{r}="","",IF({MODE}="Platform-reported only",{plat_tr},{ph_tr}))', CALC, INT)
    put(ws, f"I{r}", f'=IF($B{r}="","",IF({MODE}="Platform-reported only",$A{r}&"-reported","PostHog utm"))', CALC)
    put(ws, f"J{r}", f'=IF(OR($B{r}="",N(H{r})=0),"",C{r}/H{r})', CALC, MONEY)
    lo, hi = ci(f"C{r}", f"H{r}"); put(ws, f"K{r}", f'=IF($B{r}="","",{lo[1:]})', CALC, MONEY); put(ws, f"L{r}", f'=IF($B{r}="","",{hi[1:]})', CALC, MONEY)
    put(ws, f"M{r}", f'=IF($B{r}="","",{stage(f"C{r}", f"H{r}")[1:]})', CALC)
    put(ws, f"N{r}", f'=IF($B{r}="","",{vs_target(f"H{r}", f"K{r}", f"L{r}")[1:]})', CALC)
    w14 = lambda a, b: f',$A$1:$A$1,"<>x"'  # placeholder (kept simple below)
    rng = lambda sheet, lo_, hi_: f',{sheet}!$A:$A,">"&({P1}-{hi_}),{sheet}!$A:$A,"<="&({P1}-{lo_})'
    for col, lo_, hi_ in (("O", 0, 14), ("P", 14, 28)):
        sp = f'IF({isM},SUMIFS(IN_Meta_Campaign!C:C,IN_Meta_Campaign!$B:$B,$B{r}{rng("IN_Meta_Campaign",lo_,hi_)}),SUMIFS(IN_Google!C:C,IN_Google!$B:$B,$B{r}{rng("IN_Google",lo_,hi_)}))'
        tr = f'IF({isM},SUMIFS(IN_Meta_Campaign!F:F,IN_Meta_Campaign!$B:$B,$B{r}{rng("IN_Meta_Campaign",lo_,hi_)}),SUMIFS(IN_Google!F:F,IN_Google!$B:$B,$B{r}{rng("IN_Google",lo_,hi_)}))'
        put(ws, f"{col}{r}", f'=IF($B{r}="","",IF({tr}=0,IF({sp}>0,"no trials",""),{sp}/{tr}))', CALC, MONEY)
    put(ws, f"Q{r}", f'=IF(OR($B{r}="",NOT(ISNUMBER(O{r})),NOT(ISNUMBER(P{r}))),"",IF(O{r}<P{r}*0.9,"Improving",IF(O{r}>P{r}*1.1,"Worsening","Flat")))', CALC)
colour_status(ws, f"N5:N{4+N}")
ws.conditional_formatting.add(f"Q5:Q{4+N}", FormulaRule(formula=['Q5="Worsening"'], fill=RED))
ws.conditional_formatting.add(f"Q5:Q{4+N}", FormulaRule(formula=['Q5="Improving"'], fill=GRN))
r = 6 + N
put(ws, f"A{r}", "Check", CALC, bold=True)
put(ws, f"B{r}", "Spend in the data but not in IN_Plan (an unplanned or misnamed campaign):", CALC)
put(ws, f"C{r}", f"=SUMIFS(IN_Meta_Campaign!C:C,{DATEF('IN_Meta_Campaign')})+SUMIFS(IN_Google!C:C,{DATEF('IN_Google')})-SUM(C5:C{4+N})", CALC, MONEY, bold=True)
ws.conditional_formatting.add(f"C{r}", CellIsRule(operator="greaterThan", formula=["1"], fill=RED))
put(ws, f"A{r+1}", "Trend compares the last 14 days of the period with the 14 before, using platform-reported trials (the only daily trial data). CPA ranges are 90% intervals from the trial count.")
ws.freeze_panes = "C5"

# ---------------- Creative ----------------
ws = wb.create_sheet("Creative", 3)
put(ws, "A1", "Creative performance: concept → execution → hook → ad", bold=True)
put(ws, "A2", f'="Trials source: "&IF(COUNTA(IN_PostHog!D2:D20000)=0,"META-REPORTED (platform-attributed; PostHog has no utm_content). Undercounted after Sep 22.","PostHog utm_content")')
CC = ["Spend", "Impressions", "Clicks", "CTR", "CPC", "Avg frequency", "Trials", "CPA", "CPA – best case", "CPA – worst case", "Evidence stage", "vs target"]


def metrics(ws, r, crit, freq=True, base="ad"):
    """crit: SUMIFS criteria pairs against IN_Meta_Ads (ad level) or against the Ads table (roll-ups)."""
    if base == "ad":
        S_ = lambda col: f'SUMIFS(IN_Meta_Ads!{col}:{col},{crit},{DATEF("IN_Meta_Ads")})'
        sp, im, cl = S_("D"), S_("E"), S_("F")
        tr = f'IF(COUNTA(IN_PostHog!D2:D20000)=0,{S_("G")},SUMIFS(IN_PostHog!E:E,IN_PostHog!D:D,$A{r},{DATEF("IN_PostHog")}))'
        fq = f'IFERROR(AVERAGEIFS(IN_Meta_Ads!H:H,{crit},{DATEF("IN_Meta_Ads")}),"")'
    else:
        S_ = lambda col: f"SUMIFS(${col}$6:${col}$45,{crit})"
        sp, im, cl, tr, fq = S_("E"), S_("F"), S_("G"), S_("K"), '""'
    c0 = 5 if base == "ad" else 3   # first metric column index (E for ads, C for roll-ups)
    L = lambda i: ws.cell(r, c0 + i).column_letter
    vals = [sp, im, cl, None, None, fq, tr]
    fm = [MONEY, INT, INT, "0.00%", "$0.00", "0.0", INT]
    for i, v in enumerate(vals):
        if v is not None:
            put(ws, f"{L(i)}{r}", f"=IF($A{r}=\"\",\"\",{v})", CALC, fm[i])
    put(ws, f"{L(3)}{r}", f'=IF(N({L(1)}{r})=0,"",{L(2)}{r}/{L(1)}{r})', CALC, "0.00%")
    put(ws, f"{L(4)}{r}", f'=IF(N({L(2)}{r})=0,"",{L(0)}{r}/{L(2)}{r})', CALC, "$0.00")
    put(ws, f"{L(7)}{r}", f'=IF(N({L(6)}{r})=0,"",{L(0)}{r}/{L(6)}{r})', CALC, MONEY)
    lo, hi = ci(f"{L(0)}{r}", f"{L(6)}{r}")
    put(ws, f"{L(8)}{r}", f'=IF($A{r}="","",{lo[1:]})', CALC, MONEY); put(ws, f"{L(9)}{r}", f'=IF($A{r}="","",{hi[1:]})', CALC, MONEY)
    put(ws, f"{L(10)}{r}", f'=IF($A{r}="","",{stage(f"{L(0)}{r}", f"{L(6)}{r}")[1:]})', CALC)
    put(ws, f"{L(11)}{r}", f'=IF($A{r}="","",{vs_target(f"{L(6)}{r}", f"{L(8)}{r}", f"{L(9)}{r}")[1:]})', CALC)
    return L(11)


section(ws, "A4", "Ads (one row per ad in IN_AdMap)")
hdr(ws, 5, ["Ad name", "Concept", "Execution", "Hook"] + CC, [30, 22, 10, 8, 12, 13, 10, 8, 8, 10, 9, 10, 12, 13, 24, 28])
for k in range(40):
    r = 6 + k; ar = 2 + k
    for col, src in zip("ABCD", "ABCD"):
        put(ws, f"{col}{r}", f'=IF(IN_AdMap!{src}{ar}="","",IN_AdMap!{src}{ar})', CALC)
    last = metrics(ws, r, f"IN_Meta_Ads!$C:$C,$A{r}")
colour_status(ws, "P6:P45")

R0 = 48
section(ws, f"A{R0}", "Concepts (judge here: hooks and ads within a concept are pooled)")
hdr(ws, R0 + 1, ["Concept", "Ads"] + CC)
concepts = ["FOUNDER_STORY", "AGENCY_OWNER", "BENEFITS_STATIC", "DEMO_WALKTHROUGH", "CASE_STUDY", "PROBLEM_SOLUTION", "BEFORE_AFTER_WORKFLOW", "AGENCY_RESULTS", "TESTIMONIAL", "", "", ""]
for k, cid in enumerate(concepts):
    r = R0 + 2 + k
    put(ws, f"A{r}", cid if cid else None, INP)
    put(ws, f"B{r}", f'=IF($A{r}="","",COUNTIF($B$6:$B$45,$A{r}))', CALC, INT)
    metrics(ws, r, f"$B$6:$B$45,$A{r}", base="roll")
colour_status(ws, f"N{R0+2}:N{R0+13}")

H0 = R0 + 17
section(ws, f"A{H0}", "Hooks within a concept (directional: ~6 trials per hook per month at plan)")
hdr(ws, H0 + 1, ["Concept", "Hook"] + CC)
pairs = [("FOUNDER_STORY", "HA"), ("FOUNDER_STORY", "HB")] + [("", "")] * 8
for k, (cid, hk) in enumerate(pairs):
    r = H0 + 2 + k
    put(ws, f"A{r}", cid or None, INP); put(ws, f"B{r}", hk or None, INP)
    metrics(ws, r, f"$B$6:$B$45,$A{r},$D$6:$D$45,$B{r}", base="roll")
colour_status(ws, f"N{H0+2}:N{H0+11}")
r = H0 + 13
put(ws, f"A{r}", "Check", CALC, bold=True)
put(ws, f"B{r}", "Ad spend not mapped in IN_AdMap:", CALC)
put(ws, f"E{r}", f"=SUMIFS(IN_Meta_Ads!D:D,{DATEF('IN_Meta_Ads')})-SUM(E6:E45)", CALC, MONEY, bold=True)
ws.conditional_formatting.add(f"E{r}", CellIsRule(operator="greaterThan", formula=["1"], fill=RED))
ws.freeze_panes = "B6"

# ---------------- Summary ----------------
ws = wb.create_sheet("Summary", 0)
ws.column_dimensions["A"].width = 34; ws.column_dimensions["B"].width = 18; ws.column_dimensions["C"].width = 18; ws.column_dimensions["D"].width = 60
put(ws, "A1", f"={ST['Report name']}", bold=True)
put(ws, "A2", f'="Period: "&TEXT({P0},"d mmm yyyy")&" – "&TEXT({P1},"d mmm yyyy")&" ("&{ST["Days in period"]}&" days)"')
hdr(ws, 4, ["Measure", "This period", "Baseline", "How to read it"])
lines = [
    ("Paid spend", "=Platform!B6", f"={ST['Baseline: spend']}", MONEY, "Meta + Google, from the platform exports"),
    ("PostHog paid trials", "=Platform!D6", f"={ST['Baseline: PostHog paid trials']}", INT, "facebook + google sources in PostHog: the source of truth"),
    ("Blended CPA (PostHog)", "=Platform!F6", f"={ST['Baseline: blended CPA']}", MONEY, "Target $250–280"),
    ("Blended CPA – 90% range", '=TEXT(Platform!G6,"$#,##0")&" – "&TEXT(Platform!H6,"$#,##0")', None, None, "With this many trials the true CPA is very likely inside this range"),
    ("vs target", "=Platform!I6", None, None, ""),
    ("Platform-reported conversions", "=Platform!C6", None, INT, "Shown separately; never added to PostHog trials"),
    ("Platform ÷ PostHog", '=IF(Platform!D6=0,"",Platform!C6/Platform!D6)', None, "0.00", "≈1.0 when tracking agrees"),
    ("Paid trials per month (equivalent)", f'=Platform!D6/{ST["Days in period"]}*30.4', f"={ST['Baseline: paid trials per month']}", INT, "Totals converted to a 30.4-day month so periods compare"),
    ("Trial attribution available", f"={MODE}", None, None, "Decides whether campaign and ad trials come from PostHog or the platforms"),
    ("Weeks with Meta tracking diverging", '=COUNTIF(Platform!F11:F26,"Diverging")', None, INT, "See Platform → weekly"),
    ("Weeks with Google tracking diverging", '=COUNTIF(Platform!K11:K26,"Diverging")', None, INT, ""),
    ("Campaigns with spend but zero trials", '=COUNTIFS(Campaigns!C5:C34,">0",Campaigns!H5:H34,0)', None, INT, "Each one: verify tracking, then consider pausing (decision engine in 7b)"),
    ("Unplanned / unmatched spend", "=Campaigns!C36", None, MONEY, "Spend on campaigns not listed in IN_Plan"),
]
for i, (a, b, c, f, d) in enumerate(lines, 5):
    put(ws, f"A{i}", a, CALC, bold=True); put(ws, f"B{i}", b, CALC, f); put(ws, f"C{i}", c, CALC, f); put(ws, f"D{i}", d, CALC, wrap=True)
put(ws, "A19", "Baseline = Aug 10 – Oct 11, 2026 (63 days) from the audit. In this TEST workbook the period IS the baseline, so the two columns match.", wrap=False)
put(ws, "A20", "Yellow tabs (IN_…) are inputs: paste exports there. Grey cells are formulas. Decisions, pacing and funnel quality arrive in Part 7b.")
colour_status(ws, "B9")
ws.conditional_formatting.add("B14:B17", CellIsRule(operator="greaterThan", formula=["0"], fill=RED))

wb._sheets = [wb[n] for n in ["Summary", "Platform", "Campaigns", "Creative", "Settings", "IN_Plan", "IN_Meta_Campaign", "IN_Meta_Ads", "IN_Google", "IN_PostHog", "IN_AdMap"]]
for n in ["IN_Plan", "IN_Meta_Campaign", "IN_Meta_Ads", "IN_Google", "IN_PostHog", "IN_AdMap"]:
    wb[n].sheet_properties.tabColor = "F4C430"
wb.save(OUT); print("saved", OUT)
