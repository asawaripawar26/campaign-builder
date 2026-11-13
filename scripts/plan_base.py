"""Builds the campaign plan workbook (Part 2). Re-run to regenerate."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.comments import Comment
import datetime as dt

import os
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "build", "plan_base_only.xlsx")  # intermediate; the real plan is written by build_plan.py
os.makedirs(os.path.dirname(OUT), exist_ok=True)
F = "Arial"
HDR = PatternFill("solid", fgColor="1F2A44")
INP = PatternFill("solid", fgColor="FFF5CC")   # yellow = you edit
CALC = PatternFill("solid", fgColor="E7E9EE")  # grey = formula, don't edit
RED = PatternFill("solid", fgColor="F8C9C4")
ORG = PatternFill("solid", fgColor="FFD8A8")

wb = Workbook()
wb.remove(wb.active)


def sheet(name, headers, widths, rows, calc_cols=(), notes=None):
    ws = wb.create_sheet(name)
    ws.append(headers)
    for c in ws[1]:
        c.font = Font(name=F, bold=True, color="FFFFFF")
        c.fill = HDR
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for r in rows:
        ws.append(r)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for c in row:
            c.font = Font(name=F)
            c.alignment = Alignment(vertical="top", wrap_text=True)
            c.fill = CALC if c.column_letter in calc_cols else INP
    for i, w in enumerate(widths):
        ws.column_dimensions[ws.cell(1, i + 1).column_letter].width = w
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 32
    if notes:
        for col, txt in notes.items():
            ws[f"{col}1"].comment = Comment(txt, "Plan")
    return ws


def dv_list(ws, rng, items):
    dv = DataValidation(type="list", formula1='"' + ",".join(items) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(rng)


def flag_len(ws, rng, limit, fill=RED):
    first = rng.split(":")[0]
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f"LEN({first})>{limit}"], fill=fill))


# ---------------- How_to_use ----------------
ws = wb.create_sheet("How_to_use")
guide = [
    ("Campaign Plan — Dec 2026 rebuild", True),
    ("One row = one thing to build. The generator (Part 3) reads these tabs and writes the Google Ads Editor and Meta bulk-import files.", False),
    ("", False),
    ("Colour legend", True),
    ("Yellow cells = inputs you edit.", False),
    ("Grey cells = formulas. Don't type over them.", False),
    ("Red highlight = text over the platform's character limit (fix before export).", False),
    ("Orange highlight = row marked PLACEHOLDER (Y). Inferred copy/keywords — replace with real ones before launch.", False),
    ("", False),
    ("Tabs, in the order you fill them", True),
    ("1. Settings — brand, domain, month, UTM values, targets. Change once per month.", False),
    ("2. Campaigns — one row per campaign. 'Key' is a short unique ID that every other tab uses to link rows to a campaign. Campaign name and daily budget are calculated.", False),
    ("3. Budget — read-only check that built campaigns + reserve + existing PMax add up to the platform budgets.", False),
    ("4. Google_Keywords / Google_Negatives — one keyword per row.", False),
    ("5. Google_RSAs — one responsive search ad per ad group. Headlines ≤30 chars, descriptions ≤90, paths ≤15.", False),
    ("6. Meta_AdSets — targeting. Ad-set budgets fill in automatically when the campaign's Budget level is 'Ad set'.", False),
    ("7. Meta_Ads — one ad per row. Creative file name must match the file you upload to Meta's media library.", False),
    ("", False),
    ("Reusing next month", True),
    ("Duplicate the file, change Settings (month code, dates), adjust budgets in Campaigns, swap copy. Names, daily budgets and UTMs update themselves.", False),
    ("", False),
    ("Rules baked in", True),
    ("Everything is created PAUSED. utm_source stays 'facebook' / 'google' so PostHog source reporting stays comparable with history.", False),
    ("Meta also needs your Page ID, Instagram account ID and Pixel ID (Settings), and audience IDs written as ID:Name (Meta_AdSets). The easiest source is a CSV export of any existing campaign from Ads Manager.", False),
    ("Not built by the generator: Meta $11K conditional reserve, Google PMax US ($12K, existing campaign — budget change done manually). See Budget tab.", False),
]
for i, (t, b) in enumerate(guide, 1):
    c = ws.cell(i, 1, t)
    c.font = Font(name=F, bold=b, size=14 if i == 1 else 11)
    c.alignment = Alignment(wrap_text=True)
ws.column_dimensions["A"].width = 120

# ---------------- Settings ----------------
settings = [
    ("Brand name", "AgentStack", "PLACEHOLDER — inferred from ad name 'Static_AgentStack_Benefits'"),
    ("Landing domain", "https://www.example.com", "PLACEHOLDER — replace with the real site"),
    ("Month code", "202612", "Used in campaign names (YYYYMM)"),
    ("Days in month", 31, "Daily budget = monthly ÷ this"),
    ("Start date", dt.date(2026, 12, 1), ""),
    ("End date", dt.date(2026, 12, 31), ""),
    ("Default status", "Paused", "Nothing spends until you switch it on"),
    ("Currency", "USD", ""),
    ("Meta utm_source", "facebook", "Matches PostHog history — don't change"),
    ("Meta utm_medium", "paid_social", ""),
    ("Google utm_source", "google", "Matches PostHog history — don't change"),
    ("Google utm_medium", "cpc", ""),
    ("Target trial CPA – low", 250, "From brief"),
    ("Target trial CPA – high", 280, "From brief"),
    ("Meta platform budget", 65000, "From brief"),
    ("Google platform budget", 55000, "From brief"),
    ("Meta conditional reserve (not built)", 11000, "Released only when gates on Budget tab are met"),
    ("Google PMax US – existing (not built)", 12000, "Cut from ~$16K run-rate; change manually in Google Ads"),
    ("Facebook Page ID", "{{FB_PAGE_ID}}", "Copy from an Ads Manager export (Link Object ID column), e.g. o:1234567890"),
    ("Instagram account ID", "{{IG_ACCOUNT_ID}}", "Copy from an Ads Manager export (Instagram Account ID column)"),
    ("Meta Pixel ID", "{{META_PIXEL_ID}}", "Events Manager → your pixel → ID"),
]
ws = sheet("Settings", ["Setting", "Value", "Note"], [36, 30, 60], settings, calc_cols=("A", "C"))
for r in (6, 7):
    ws[f"B{r}"].number_format = "yyyy-mm-dd"
for r in (14, 15, 16, 17, 18, 19):
    ws[f"B{r}"].number_format = "$#,##0"
S = {k: f"Settings!$B${i}" for i, (k, _, _) in enumerate(settings, 2)}

# ---------------- Campaigns ----------------
camp_hdr = ["Platform", "Funnel", "Key", "Geo", "Campaign name", "Campaign type", "Objective",
            "Bid strategy", "Target CPA ($)", "Monthly budget ($)", "Daily budget ($)", "Budget level",
            "Conversion event", "Start date", "End date", "Locations", "Languages", "Status", "Notes"]
camps = [
    ("Google", "BRAND", "Brand", "US", "Search", "Leads", "Maximize conversions", None, 9000, "Campaign",
     "Trial signup", "United States", "English", "Capped: much brand demand exists anyway. Watch impression share."),
    ("Google", "NONBRAND", "NonBrand", "US", "Search", "Leads", "Target CPA", 290, 29000, "Campaign",
     "Trial signup", "United States", "English", "$29K is a ceiling, not a target. Start tCPA near the $292 history, step toward $260."),
    ("Google", "TEST", "CompTest", "US", "Search", "Leads", "Maximize conversions", None, 5000, "Campaign",
     "Trial signup", "United States", "English", "High-intent comparison/competitor terms only."),
    ("Meta", "PROS", "AdvPlus", "US", "Advantage+ sales", "Sales", "Highest volume", None, 40000, "Campaign",
     "StartTrial", "United States", "", "Geo assumed US — confirm. No cost cap until tracking is validated."),
    ("Meta", "TEST", "CreativeTest", "US", "Manual sales", "Sales", "Highest volume", None, 10000, "Ad set",
     "StartTrial", "United States", "", "Equal budget per concept so each gets a fair read (~$2.5K each)."),
    ("Meta", "RET", "RT14d", "US", "Manual sales", "Sales", "Highest volume", None, 4000, "Campaign",
     "StartTrial", "United States", "", "Small cap until incrementality is shown. 14d window is a test."),
]
rows = []
for i, c in enumerate(camps, 2):
    plat, fun, key, geo, typ, obj, bid, tcpa, bud, lvl, ev, loc, lang, note = c
    rows.append([plat, fun, key, geo,
                 f'=IF(A{i}="Google","GOOG","META")&"_"&B{i}&"_"&C{i}&"_"&D{i}&"_"&{S["Month code"]}',
                 typ, obj, bid, tcpa, bud, f"=ROUND(J{i}/{S['Days in month']},0)", lvl, ev,
                 f"={S['Start date']}", f"={S['End date']}", loc, lang, f"={S['Default status']}", note])
ws = sheet("Campaigns", camp_hdr, [9, 11, 13, 6, 34, 17, 9, 20, 10, 13, 11, 11, 13, 11, 11, 15, 10, 9, 50],
           rows, calc_cols=("E", "K", "N", "O", "R"),
           notes={"C": "Short unique ID. Other tabs link to the campaign through this.",
                  "E": "Auto: PLATFORM_FUNNEL_KEY_GEO_MONTH"})
for r in range(2, 8):
    ws[f"I{r}"].number_format = ws[f"J{r}"].number_format = ws[f"K{r}"].number_format = "$#,##0"
    ws[f"N{r}"].number_format = ws[f"O{r}"].number_format = "yyyy-mm-dd"
dv_list(ws, "A2:A200", ["Google", "Meta"])
dv_list(ws, "B2:B200", ["BRAND", "NONBRAND", "TEST", "PROS", "RET"])
dv_list(ws, "L2:L200", ["Campaign", "Ad set"])
dv_list(ws, "F2:F200", ["Search", "Advantage+ sales", "Manual sales"])

# ---------------- Budget ----------------
ws = wb.create_sheet("Budget")
b = [
    ["Line", "Meta ($)", "Google ($)", "Note"],
    ["Built by generator (Campaigns tab)", '=SUMIFS(Campaigns!J:J,Campaigns!A:A,"Meta")',
     '=SUMIFS(Campaigns!J:J,Campaigns!A:A,"Google")', "Sum of Monthly budget"],
    ["Conditional reserve (not built)", f"={S['Meta conditional reserve (not built)']}", 0, "Released only when all gates below are met"],
    ["PMax US – existing (not built)", 0, f"={S['Google PMax US – existing (not built)']}", "Reduced from ~$16K; manual change"],
    ["Total planned", "=SUM(B2:B4)", "=SUM(C2:C4)", ""],
    ["Platform budget (brief)", f"={S['Meta platform budget']}", f"={S['Google platform budget']}", ""],
    ["Difference (should be 0)", "=B5-B6", "=C5-C6", "Red if not zero"],
    [],
    ["Release gates for the reserve", "", "", ""],
    ["1. Tracking validated: Meta Pixel + Conversions API firing and de-duplicated on v3 signup pages; Google conversion tag verified.", "", "", ""],
    ["2. Funnel validated: signup-page conversion rate and trial activation not worse than before Sep 22.", "", "", ""],
    ["3. Performance: true (PostHog) cost per trial ≤ $280 for 7 consecutive days; trial quality tracked.", "", "", ""],
]
for r in b:
    ws.append(r)
for row in ws.iter_rows():
    for c in row:
        c.font = Font(name=F, bold=c.row in (1, 5, 9))
        if c.row in range(2, 8) and c.column in (2, 3):
            c.number_format = "$#,##0;($#,##0);-"
            c.fill = CALC
for c in ws[1]:
    c.font = Font(name=F, bold=True, color="FFFFFF"); c.fill = HDR
ws.conditional_formatting.add("B7:C7", FormulaRule(formula=["B7<>0"], fill=RED))
for col, w in zip("ABCD", (60, 14, 14, 45)):
    ws.column_dimensions[col].width = w

# ---------------- Google_Keywords ----------------
PH = "Y"
kw = []
def add(camp, ag, words, match):
    for w in words:
        kw.append([camp, ag, w, match, None, PH])
add("Brand", "Brand_Core", ["agentstack", "agentstack ai", "agentstack pricing", "agentstack free trial", "agentstack reviews"], "Exact")
add("Brand", "Brand_Core", ["agentstack"], "Phrase")
add("NonBrand", "AI_Agents_Agencies", ["ai agents for agencies", "ai agent platform for agencies", "ai automation for marketing agencies"], "Phrase")
add("NonBrand", "AI_Agents_Agencies", ["ai agents for marketing agencies", "agency ai automation software"], "Exact")
add("NonBrand", "Agentic_SEO", ["agentic seo", "ai seo agent", "ai seo automation"], "Phrase")
add("NonBrand", "Agentic_SEO", ["ai seo automation tool", "seo ai agent software"], "Exact")
add("NonBrand", "AI_Marketing_Automation", ["ai marketing agent", "ai marketing automation software", "ai workflow automation for marketing"], "Phrase")
add("NonBrand", "AI_Marketing_Automation", ["ai agent platform"], "Exact")
add("CompTest", "Competitor_Alternatives", ["{{COMPETITOR_1}} alternative", "{{COMPETITOR_2}} alternative", "{{COMPETITOR_1}} vs"], "Phrase")
add("CompTest", "Best_Platform", ["best ai agent platform", "best ai seo tools for agencies"], "Phrase")
add("CompTest", "Best_Platform", ["ai agent software comparison"], "Exact")
ws = sheet("Google_Keywords", ["Campaign key", "Ad group", "Keyword", "Match type", "Max CPC ($, optional)", "PLACEHOLDER"],
           [14, 26, 42, 11, 14, 13], kw, notes={"C": "{{...}} tokens must be replaced — the generator blocks export while they remain."})
dv_list(ws, "D2:D500", ["Exact", "Phrase", "Broad"])
dv_list(ws, "F2:F500", ["Y", "N"])
ws.conditional_formatting.add("A2:F500", FormulaRule(formula=['$F2="Y"'], fill=ORG))

# ---------------- Google_Negatives ----------------
neg = [["Brand", w, "Phrase", PH] for w in ["login", "sign in", "support", "careers", "jobs"]]
neg += [["NonBrand", w, "Phrase", PH] for w in ["jobs", "careers", "salary", "course", "certification", "tutorial",
                                                "github", "open source", "what is", "agentstack"]]
neg += [["CompTest", w, "Phrase", PH] for w in ["jobs", "careers", "course", "free download", "agentstack"]]
neg.append(["NonBrand", "{{IMPORT the 40 negatives added Oct 6 from the live account}}", "Phrase", PH])
ws = sheet("Google_Negatives", ["Campaign key", "Negative keyword", "Match type", "PLACEHOLDER"], [14, 60, 11, 13], neg,
           notes={"B": "'agentstack' is negative in non-brand/test so brand searches go to the Brand campaign only."})
dv_list(ws, "C2:C500", ["Exact", "Phrase", "Broad"])
ws.conditional_formatting.add("A2:D500", FormulaRule(formula=['$D2="Y"'], fill=ORG))

# ---------------- Google_RSAs ----------------
rsa = {
    ("Brand", "Brand_Core", "/", "Official", "Free-Trial"): (
        ["AgentStack Official Site", "Start Your Free Trial", "AI Agents for Agencies", "Automate Client SEO Work",
         "Launch Agents in Minutes", "Built for Marketing Teams", "Try AgentStack Free", "See AgentStack Pricing"],
        ["Build and run AI agents for SEO, content and reporting. Start your free trial today.",
         "Agencies use AgentStack to automate repetitive client work. See how it works.",
         "Set up your first agent in minutes and scale across every client account."]),
    ("NonBrand", "AI_Agents_Agencies", "/trial/agencies", "Agencies", "AI-Agents"): (
        ["AI Agents for Agencies", "Automate Client Deliverables", "Scale Without New Hires", "Start Your Free Trial",
         "Built for Agency Teams", "Agents for SEO and Content", "Free Up Billable Hours", "Try AgentStack Free"],
        ["Hand repetitive client work to AI agents and give your team time back for strategy.",
         "Audits, briefs and reports run on autopilot across every client. Start a free trial.",
         "Built for agencies that want to grow output without growing headcount."]),
    ("NonBrand", "Agentic_SEO", "/trial/seo", "Agentic-SEO", "Free-Trial"): (
        ["AI Agents That Do SEO Work", "Automate SEO Audits", "Content Briefs in Minutes", "Agentic SEO for Agencies",
         "Rank Tracking on Autopilot", "Start Your Free Trial", "Try AgentStack Free", "Built for SEO Teams"],
        ["Let AI agents run technical audits, keyword research and content briefs for you.",
         "Agentic SEO that works across every client site. Start your free trial today.",
         "Spend less time on SEO busywork and more time on strategy that moves rankings."]),
    ("NonBrand", "AI_Marketing_Automation", "/trial", "AI-Marketing", "Agents"): (
        ["AI Marketing Agents", "Automate Marketing Workflows", "Build Agents Without Code", "Start Your Free Trial",
         "From Brief to Report Faster", "Try AgentStack Free", "Connect Your Marketing Stack", "Launch Agents in Minutes"],
        ["Create AI agents that handle research, content and reporting across your stack.",
         "Automate the repetitive parts of marketing and keep your team on high-value work.",
         "Start a free trial and launch your first marketing agent today."]),
    ("CompTest", "Competitor_Alternatives", "/compare", "Compare", "Alternatives"): (
        ["A Smarter Alternative", "Compare AgentStack", "See How We Compare", "Built for Agency Workflows",
         "Start Your Free Trial", "Try AgentStack Free", "AI Agents for Agencies", "Switching Made Simple"],
        ["Looking for an alternative? See why agencies choose AgentStack for AI automation.",
         "Compare features side by side, then start a free trial and see the difference.",
         "Purpose-built AI agents for SEO, content and client reporting."]),
    ("CompTest", "Best_Platform", "/compare", "AI-Agents", "Compare"): (
        ["Compare AI Agent Platforms", "Built for Agency Workflows", "AI Agents for SEO & Content", "Start Your Free Trial",
         "Try AgentStack Free", "See How We Compare", "Launch Agents in Minutes", "AI Agents for Agencies"],
        ["Choosing an AI agent platform? Compare features built for agencies and SEO teams.",
         "See how AgentStack handles audits, briefs and reports, then start a free trial.",
         "Purpose-built AI agents for SEO, content and client reporting."]),
}
rsa_rows = []
for (camp, ag, path, p1, p2), (hs, ds) in rsa.items():
    for h in hs:
        assert len(h) <= 30, h
    for d in ds:
        assert len(d) <= 90, (len(d), d)
    assert len(p1) <= 15 and len(p2) <= 15
    rsa_rows.append([camp, ag, path, p1, p2] + hs + [None] * (15 - len(hs)) + ds + [None] * (4 - len(ds)) + [PH])
hdr = ["Campaign key", "Ad group", "Landing path", "Path 1", "Path 2"] + [f"Headline {i}" for i in range(1, 16)] + \
      [f"Description {i}" for i in range(1, 5)] + ["PLACEHOLDER"]
ws = sheet("Google_RSAs", hdr, [13, 24, 16, 13, 13] + [26] * 15 + [45] * 4 + [13], rsa_rows,
           notes={"C": "Path only. Generator adds Settings domain + UTMs.",
                  "F": "Max 30 characters each. 3 required, 8–15 recommended.",
                  "U": "Max 90 characters each. 2 required."})
flag_len(ws, "D2:E100", 15)
flag_len(ws, "F2:T100", 30)
flag_len(ws, "U2:X100", 90)
ws.conditional_formatting.add("Y2:Y100", FormulaRule(formula=['$Y2="Y"'], fill=ORG))

# ---------------- Meta_AdSets ----------------
EXCL = "Trial signups – 180d; Customers"
adsets = [
    ("AdvPlus", "ASC_Broad_US", 25, 65, "Advantage+ audience", EXCL, "Advantage+ placements"),
    ("CreativeTest", "TEST_C1_Carousel_Case_Study", 25, 65, "Broad", EXCL, "Advantage+ placements"),
    ("CreativeTest", "TEST_C2_Founder_NewHook", 25, 65, "Broad", EXCL, "Advantage+ placements"),
    ("CreativeTest", "TEST_C3_Agency_Results_Static", 25, 65, "Broad", EXCL, "Advantage+ placements"),
    ("CreativeTest", "TEST_C4_Problem_Solution_Video", 25, 65, "Broad", EXCL, "Advantage+ placements"),
    ("RT14d", "RT_Web14d_ExclTrials", 25, 65, "Website visitors – 14d", EXCL, "Advantage+ placements"),
]
rows = []
for i, (camp, name, amin, amax, inc, exc, plc) in enumerate(adsets, 2):
    lvl = f"INDEX(Campaigns!$L:$L,MATCH(A{i},Campaigns!$C:$C,0))"
    bud = (f'=IF({lvl}="Ad set",ROUND(INDEX(Campaigns!$J:$J,MATCH(A{i},Campaigns!$C:$C,0))'
           f'/{S["Days in month"]}/COUNTIF($A:$A,A{i}),0),"Campaign budget")')
    rows.append([camp, name, bud, "United States", amin, amax, "All", inc, exc, plc, "StartTrial", "Y"])
ws = sheet("Meta_AdSets", ["Campaign key", "Ad set name", "Ad set daily budget ($)", "Countries", "Age min", "Age max",
                           "Gender", "Include audiences", "Exclude audiences", "Placements", "Conversion event", "PLACEHOLDER"],
           [14, 32, 16, 15, 8, 8, 8, 26, 32, 22, 13, 13], rows, calc_cols=("C",),
           notes={"H": "Meta needs the audience ID: write ID:Name, e.g. 23851234567890:Website visitors – 14d. Separate several with ;",
                  "C": "Auto: splits the campaign's monthly budget equally across its ad sets when Budget level = Ad set."})
for r in range(2, 8):
    ws[f"C{r}"].number_format = "$#,##0"
ws.conditional_formatting.add("A2:L100", FormulaRule(formula=['$L2="Y"'], fill=ORG))

# ---------------- Meta_Ads ----------------
ads = [
    ("AdvPlus", "ASC_Broad_US", "UGC_Founder_v3", "Video", "UGC_Founder_v3.mp4",
     "I built AgentStack because my agency was drowning in repetitive SEO work. Now AI agents handle audits, briefs and reports.",
     "Start your free trial", "AI agents for agencies", "SIGN_UP", "/trial/agencies",
     "Refresh of best ad (v2, $200 CPA). Same concept, new hook/edit."),
    ("AdvPlus", "ASC_Broad_US", "UGC_Agency_Owner", "Video", "UGC_Agency_Owner.mp4",
     "We stopped hiring for busywork. Our AI agents now run SEO audits and client reports every week.",
     "Scale without new hires", "Try AgentStack free", "SIGN_UP", "/trial/agencies", "Existing ad, $218 pre-break CPA."),
    ("AdvPlus", "ASC_Broad_US", "Static_AgentStack_Benefits", "Image", "Static_AgentStack_Benefits.png",
     "Audits, content briefs and reporting — handled by AI agents built for agencies.",
     "AI agents for agencies", "Start free trial", "SIGN_UP", "/trial", "Existing ad, $243 pre-break CPA."),
    ("AdvPlus", "ASC_Broad_US", "Demo_Screen_Walkthrough", "Video", "Demo_Screen_Walkthrough.mp4",
     "See how an AI agent builds a full SEO audit in minutes. No setup headaches.",
     "Watch it work, then try it", "Free trial", "SIGN_UP", "/trial", "Existing ad, $247 pre-break CPA."),
    ("CreativeTest", "TEST_C1_Carousel_Case_Study", "Carousel_Case_Study", "Carousel", "Carousel_Case_Study_1.png|Carousel_Case_Study_2.png|Carousel_Case_Study_3.png",
     "How one agency cut reporting time with AI agents — swipe for the results.",
     "See the case study", "Start free trial", "LEARN_MORE", "/customers", "Re-test with clean tracking (only 1 clean week so far)."),
    ("CreativeTest", "TEST_C2_Founder_NewHook", "UGC_Founder_v3_HookB", "Video", "UGC_Founder_v3_HookB.mp4",
     "Your team didn't get into marketing to copy-paste reports. Let AI agents do it.",
     "Start your free trial", "AI agents for agencies", "SIGN_UP", "/trial/agencies", "New hook on the winning founder concept."),
    ("CreativeTest", "TEST_C3_Agency_Results_Static", "Static_Agency_Results", "Image", "Static_Agency_Results.png",
     "{{REAL CUSTOMER RESULT}} — what agencies get back when AI agents handle the busywork.",
     "Built for agencies", "Start free trial", "SIGN_UP", "/trial/agencies", "Needs a real, approved customer stat."),
    ("CreativeTest", "TEST_C4_Problem_Solution_Video", "Video_Problem_Solution", "Video", "Video_Problem_Solution.mp4",
     "Too many clients, not enough hours? AgentStack's AI agents take the repetitive work off your plate.",
     "Get hours back every week", "Start free trial", "SIGN_UP", "/trial", "New concept."),
    ("RT14d", "RT_Web14d_ExclTrials", "Retargeting_Testimonial", "Video", "Retargeting_Testimonial.mp4",
     "Still thinking it over? Hear why agencies switched to AgentStack.",
     "Start your free trial", "Set up in minutes", "SIGN_UP", "/trial", "Existing ad."),
    ("RT14d", "RT_Web14d_ExclTrials", "Demo_Screen_Walkthrough_RT", "Video", "Demo_Screen_Walkthrough.mp4",
     "You've seen the site — now see an AI agent build a full SEO audit in minutes.",
     "See it in action", "Free trial", "SIGN_UP", "/trial", "Reuses existing demo video."),
]
for a in ads:
    assert len(a[6]) <= 40 and len(a[7]) <= 30, a[2]
rows = [list(a[:10]) + ["Y", a[10]] for a in ads]
ws = sheet("Meta_Ads", ["Campaign key", "Ad set name", "Ad name", "Format", "Creative file(s)", "Primary text",
                        "Headline", "Description", "Call to action", "Landing path", "PLACEHOLDER", "Notes"],
           [13, 30, 28, 9, 34, 60, 26, 24, 13, 16, 13, 40], rows,
           notes={"E": "Must match file names in Meta's media library. Carousel: separate files with |",
                  "F": "Keep under 125 characters so it isn't cut off.",
                  "G": "Keep under 40 characters.", "H": "Keep under 30 characters."})
flag_len(ws, "F2:F200", 125, ORG)
flag_len(ws, "G2:G200", 40)
flag_len(ws, "H2:H200", 30)
dv_list(ws, "D2:D200", ["Image", "Video", "Carousel"])
dv_list(ws, "I2:I200", ["SIGN_UP", "LEARN_MORE", "GET_OFFER", "DOWNLOAD", "CONTACT_US"])

wb.save(OUT)
print("saved", OUT)
