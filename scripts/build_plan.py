"""Builds the plan workbook (plan/Campaign_Plan_Dec2026.xlsx) on top of plan_base.py. Re-run after changing the sample values."""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import plan_base as mp            # builds the base workbook in memory
from openpyxl.styles import Font, Alignment
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule

OUT = os.path.join(mp.ROOT, "plan", "Campaign_Plan_Dec2026.xlsx")
wb, F = mp.wb, mp.F
PH = "Y"


def style_row(ws, r, calc_cols=()):
    for c in ws[r]:
        c.font = Font(name=F)
        c.alignment = Alignment(vertical="top", wrap_text=True)
        c.fill = mp.CALC if c.column_letter in calc_cols else mp.INP


# ---------------- Settings: v2 rows ----------------
ws = wb["Settings"]
ws["B19"] = 0
ws["C19"] = "v2: PMax is now edited in place from the Campaigns tab, so nothing sits outside the build"
new_settings = [
    ("Target CPA for planning", "=AVERAGE(B14:B15)", "Midpoint of the target range; used for test evidence maths"),
    ("Max test concepts per month", 3, "Agreed default: judge concepts, not individual ads"),
    ("Min trials to declare a winner", 20, "Below this a concept is 'inconclusive', not 'failed'"),
    ("Early-cut window (days)", 7, "Early signals (CTR, cost per landing-page view, hook rate) may pause clear losers only"),
    ("Google tracking verified on", "", "Date someone confirmed Google conversions fire on the v3 signup flow. PMax export is BLOCKED while empty"),
    ("Meta tracking verified on", "", "Date someone confirmed Pixel + Conversions API fire and de-duplicate on v3. Meta export warns while empty"),
    ("PMax business name", "AgentStack", "PLACEHOLDER — max 25 characters"),
    ("Google billing days per month", 30.4, "Google caps monthly spend at daily budget × 30.4, so Google daily = monthly ÷ 30.4. Meta uses Days in month."),
]
start = ws.max_row + 1
for i, row in enumerate(new_settings):
    ws.append(list(row))
    style_row(ws, start + i, calc_cols=("A", "C"))
ws[f"B{start}"].number_format = "$#,##0"
ws[f"B{start}"].fill = mp.CALC
for r in (start + 4, start + 5):
    ws[f"B{r}"].number_format = "yyyy-mm-dd"
SET = {k: f"Settings!$B${start + i}" for i, (k, _, _) in enumerate(new_settings)}

# ---------------- Campaigns: PMax row + "edit in place" column ----------------
ws = wb["Campaigns"]
ws["T1"] = "Existing campaign name (edit in place)"
ws["T1"].font = Font(name=F, bold=True, color="FFFFFF"); ws["T1"].fill = mp.HDR
ws["T1"].alignment = Alignment(wrap_text=True, vertical="center")
ws["T1"].comment = Comment("Fill this to EDIT an existing campaign instead of creating one. The name must match Google/Meta exactly. "
                           "The generator then uses this name and never changes the live campaign's status.", "Plan")
ws.column_dimensions["T"].width = 30
ws.append(["Google", "PMAX", "PMax", "US", None, "Performance Max", "Leads", "Maximize conversions", None, 12000, None,
           "Campaign", "Trial signup", None, None, "United States", "English", None,
           "Rebuilt IN PLACE: keeps campaign history. New asset groups import Paused; switch them on and pause the old ones together. Brand exclusion is set manually.",
           "{{EXISTING_PMAX_NAME}}"])
r = ws.max_row
for col, f in (("K", f"=ROUND(J{r}/{mp.S['Days in month']},0)"), ("N", f"={mp.S['Start date']}"),
               ("O", f"={mp.S['End date']}"), ("R", f"={mp.S['Default status']}")):
    ws[f"{col}{r}"] = f
style_row(ws, r, calc_cols=("E", "K", "N", "O", "R"))
for rr in range(2, r + 1):
    ws[f"K{rr}"] = f'=ROUND(J{rr}/IF(A{rr}="Google",{SET["Google billing days per month"]},{mp.S["Days in month"]}),0)'
    ws[f"E{rr}"] =(f'=IF(T{rr}<>"",T{rr},IF(A{rr}="Google","GOOG","META")&"_"&B{rr}&"_"&C{rr}&"_"&D{rr}&"_"&{mp.S["Month code"]})')
    ws[f"T{rr}"].fill = mp.INP; ws[f"T{rr}"].font = Font(name=F)
for col in "IJK":
    ws[f"{col}{r}"].number_format = "$#,##0"
ws[f"N{r}"].number_format = ws[f"O{r}"].number_format = "yyyy-mm-dd"
mp.dv_list(ws, "F2:F200", ["Search", "Performance Max", "Advantage+ sales", "Manual sales"])
mp.dv_list(ws, "B2:B200", ["BRAND", "NONBRAND", "TEST", "PMAX", "PROS", "RET"])

# ---------------- Concepts (the hub) ----------------
STATUSES = ["Hypothesis", "In production", "Ready for testing", "Testing", "Validated", "Retired"]
concepts = [  # ..., Status, Evidence, Claim required, Claim ID
    ("FOUNDER_STORY", "Founder story (UGC)", "Agency owners", "Team drowning in repetitive SEO work", "Founder built AgentStack to get his agency's time back", "/trial/agencies", "SIGN_UP", "Both", "Video; Image (stills)", "Validated", "Best ad pre-break: $200 CPA (UGC_Founder_v2)", "N", ""),
    ("AGENCY_OWNER", "Agency owner testimonial (UGC)", "Agency owners / ops leads", "Scaling output without hiring", "Grow output without growing headcount", "/trial/agencies", "SIGN_UP", "Both", "Video", "Validated", "$218 CPA pre-break", "N", ""),
    ("BENEFITS_STATIC", "Benefits static", "Marketing & SEO leads", "Too much manual work across tools", "Audits, briefs and reports handled by AI agents", "/trial", "SIGN_UP", "Both", "Image", "Validated", "$243 CPA pre-break", "N", ""),
    ("DEMO_WALKTHROUGH", "Product demo walkthrough", "SEO practitioners", "Want to see it working before trying", "Watch an agent build a full SEO audit", "/trial/seo", "SIGN_UP", "Both", "Video; Image", "Validated", "$247 CPA pre-break", "N", ""),
    ("CASE_STUDY", "Agency case study", "Agency owners", "Need proof it works for agencies", "One agency cut reporting time with AI agents", "/customers", "LEARN_MORE", "Both", "Image", "Ready for testing", "Only 1 clean week so far ($380). Re-test", "N", ""),
    ("PROBLEM_SOLUTION", "Problem → solution video", "Agency owners", "Too many clients, not enough hours", "AI agents take the repetitive work off your plate", "/trial", "SIGN_UP", "Meta", "Video", "Ready for testing", "New concept", "N", ""),
    ("AGENCY_RESULTS", "Agency results static", "Agency owners", "Need numbers before committing", "Real, approved customer result (see Claims)", "/trial/agencies", "SIGN_UP", "Both", "Image", "In production", "Planned. Blocked until claim CL01 is approved; then it can take a test slot", "Y", "CL01"),
    ("BEFORE_AFTER_WORKFLOW", "Before / after workflow (fallback)", "Agency owners", "Manual workflows eat the week", "The same audit: manual vs. run by an AI agent", "/trial/agencies", "SIGN_UP", "Meta", "Image", "Ready for testing", "In the 3rd test slot (AGENCY_RESULTS blocked until CL01 approved). Shows the workflow; no quantified claims", "N", ""),
    ("TESTIMONIAL", "Retargeting testimonial", "Site visitors who didn't sign up", "Still deciding", "Hear why agencies switched", "/trial", "SIGN_UP", "Meta", "Video", "Testing", "$342 CPA pre-break; incrementality unproven", "N", ""),
    ("BRAND_LOGO", "Brand logos", "—", "—", "—", "/", "", "PMax", "Logo", "Validated", "Required by PMax: square logo (1:1), optional 4:1", "N", ""),
]
ws = mp.sheet("Concepts", ["Concept ID", "Concept name", "Audience / ICP", "Problem or intent", "Core message", "Landing path",
                           "Primary CTA", "Platforms", "Formats", "Status", "Evidence so far", "Claim required", "Claim ID", "PLACEHOLDER"],
              [22, 26, 24, 30, 36, 16, 12, 9, 11, 15, 34, 9, 9, 13], [list(c) + [PH] for c in concepts],
              notes={"A": "Short unique ID. Ads, files, variants and PMax asset groups all point here.",
                     "F": "Where this concept's traffic lands. Ads and asset groups inherit it unless they override it.",
                     "J": "Hypothesis → In production → Ready for testing → Testing → Validated → Retired. "
                          "Test ad sets accept Ready for testing / Testing. Always-on placements expect Validated. Retired never exports.",
                     "L": "Y = the copy makes a factual claim (result, stat, comparison). The concept can't export until its Claim ID is Approved on the Claims tab."})
mp.dv_list(ws, "J2:J200", STATUSES)
mp.dv_list(ws, "L2:L200", ["Y", "N"])

# ---------------- Claims ----------------
ws = mp.sheet("Claims", ["Claim ID", "Concept ID", "Claim type", "Claim text", "Evidence / source", "Approval status",
                         "Approved by", "Date approved", "PLACEHOLDER"],
              [9, 20, 16, 50, 40, 14, 18, 13, 13],
              [["CL01", "AGENCY_RESULTS", "Customer result", "{{REAL CUSTOMER RESULT}}", "{{SOURCE: case study / customer sign-off}}", "Pending", "", "", PH]],
              notes={"F": "Only 'Approved' with a source, an approver and a date unlocks the concept for export."})
mp.dv_list(ws, "F2:F200", ["Pending", "Approved", "Rejected"])
mp.dv_list(ws, "C2:C200", ["Customer result", "Statistic", "Comparison", "Award / rating", "Other"])
mp.dv_list(ws, "H2:H200", ["Meta", "PMax", "Both"])

# ---------------- Creative_Library ----------------
USE = {"4x5": "Meta feed; PMax portrait", "1x1": "Meta feed; PMax square", "9x16": "Meta Stories/Reels",
       "191x1": "PMax landscape", "16x9": "PMax video (landscape)", "4x1": "PMax landscape logo"}
LIB = [  # concept, version, kind, ratios
    ("FOUNDER_STORY", "v3", "Video", ["4x5", "9x16", "16x9"]),
    ("FOUNDER_STORY", "v3", "Image", ["191x1", "1x1"]),
    ("AGENCY_OWNER", "v1", "Video", ["4x5", "9x16", "16x9"]),
    ("BENEFITS_STATIC", "v1", "Image", ["1x1", "4x5", "9x16", "191x1"]),
    ("DEMO_WALKTHROUGH", "v1", "Video", ["4x5", "9x16", "16x9"]),
    ("DEMO_WALKTHROUGH", "v1", "Image", ["191x1", "1x1"]),
    ("CASE_STUDY", "v2", "Image", ["4x5", "9x16", "191x1", "1x1"]),
    ("PROBLEM_SOLUTION", "v1", "Video", ["4x5", "9x16"]),
    ("AGENCY_RESULTS", "v1", "Image", ["4x5", "9x16", "191x1", "1x1"]),
    ("TESTIMONIAL", "v1", "Video", ["4x5", "9x16"]),
    ("BEFORE_AFTER_WORKFLOW", "v1", "Image", ["4x5", "9x16"]),
    ("BRAND_LOGO", "v1", "Logo", ["1x1", "4x1"]),
]
lib_rows = []
for con, ver, kind, ratios in LIB:
    for ra in ratios:
        ext = "mp4" if kind == "Video" else "png"
        use = "PMax square logo" if (kind == "Logo" and ra == "1x1") else USE[ra]
        lib_rows.append([f"{con}_{ver}_{ra}_TEST.{ext}", con, ver, ra, kind, use, "TEST", 12 if kind == "Video" else None, PH])
ws = mp.sheet("Creative_Library", ["File name", "Concept ID", "Version", "Ratio", "Asset type", "Used for",
                                   "Approval status", "Video length (s)", "PLACEHOLDER"],
              [40, 20, 8, 7, 10, 26, 13, 12, 13], lib_rows,
              notes={"A": "Naming convention: CONCEPT_version_ratio.ext  (e.g. FOUNDER_STORY_v3_9x16.mp4). Test files end in _TEST.",
                     "D": "1x1, 4x5, 9x16, 16x9, 191x1 (=1.91:1), 4x1 (logo). Must match the ratio in the file name.",
                     "G": "Only 'Approved' files should go into a real launch. TEST files are generated stand-ins."})
mp.dv_list(ws, "D2:D500", ["1x1", "4x5", "9x16", "16x9", "191x1", "4x1"])
mp.dv_list(ws, "E2:E500", ["Image", "Video", "Logo"])
mp.dv_list(ws, "G2:G500", ["Approved", "TEST", "Retired"])
ws.conditional_formatting.add("A2:I500", FormulaRule(formula=['$G2="TEST"'], fill=mp.ORG))

# ---------------- Meta_AdSets / Meta_Ads (rebuilt around concepts) ----------------
for name in ("Meta_AdSets", "Meta_Ads"):
    del wb[name]
EXCL = "Trial signups – 180d; Customers"
adsets = [("AdvPlus", "ASC_Broad_US", "Advantage+ audience", ""),
          ("CreativeTest", "TEST_CASE_STUDY", "Broad", "CASE_STUDY"),
          ("CreativeTest", "TEST_PROBLEM_SOLUTION", "Broad", "PROBLEM_SOLUTION"),
          ("CreativeTest", "TEST_BEFORE_AFTER_WORKFLOW", "Broad", "BEFORE_AFTER_WORKFLOW"),
          ("RT14d", "RT_Web14d_ExclTrials", "Website visitors – 14d", "")]
rows = []
for i, (camp, name, inc, con) in enumerate(adsets, 2):
    lvl = f"INDEX(Campaigns!$L:$L,MATCH(A{i},Campaigns!$C:$C,0))"
    bud = (f'=IF({lvl}="Ad set",ROUND(INDEX(Campaigns!$J:$J,MATCH(A{i},Campaigns!$C:$C,0))'
           f'/{mp.S["Days in month"]}/COUNTIF($A:$A,A{i}),0),"Campaign budget")')
    rows.append([camp, name, bud, "United States", 25, 65, "All", inc, EXCL, "Advantage+ placements", "StartTrial", con, PH])
ws = mp.sheet("Meta_AdSets", ["Campaign key", "Ad set name", "Ad set daily budget ($)", "Countries", "Age min", "Age max",
                              "Gender", "Include audiences", "Exclude audiences", "Placements", "Conversion event",
                              "Test concept ID", "PLACEHOLDER"],
              [14, 28, 16, 15, 8, 8, 8, 26, 32, 22, 13, 20, 13], rows, calc_cols=("C",),
              notes={"H": "Meta needs the audience ID: write ID:Name, e.g. 23851234567890:Website visitors – 14d. Separate several with ;",
                     "L": "Creative-test ad sets only: one concept per ad set. The generator builds its ads from the Variants tab."})
for rr in range(2, 7):
    ws[f"C{rr}"].number_format = "$#,##0"

ads = [
    ("AdvPlus", "ASC_Broad_US", "FOUNDER_STORY_v3_HookA", "FOUNDER_STORY", "Video",
     "I built AgentStack because my agency was drowning in repetitive SEO work. Now AI agents handle audits, briefs and reports.",
     "Start your free trial", "AI agents for agencies", "", "Was UGC_Founder_v2 (best ad). Refreshed edit."),
    ("AdvPlus", "ASC_Broad_US", "FOUNDER_STORY_v3_HookB", "FOUNDER_STORY", "Video",
     "Your team didn't get into marketing to copy-paste reports. Let AI agents do it.",
     "Start your free trial", "AI agents for agencies", "", "New hook on the winning concept. Judged pooled with HookA."),
    ("AdvPlus", "ASC_Broad_US", "AGENCY_OWNER_v1", "AGENCY_OWNER", "Video",
     "We stopped hiring for busywork. Our AI agents now run SEO audits and client reports every week.",
     "Scale without new hires", "Try AgentStack free", "", "Was UGC_Agency_Owner."),
    ("AdvPlus", "ASC_Broad_US", "BENEFITS_STATIC_v1", "BENEFITS_STATIC", "Image",
     "Audits, content briefs and reporting — handled by AI agents built for agencies.",
     "AI agents for agencies", "Start free trial", "", "Was Static_AgentStack_Benefits."),
    ("AdvPlus", "ASC_Broad_US", "DEMO_WALKTHROUGH_v1", "DEMO_WALKTHROUGH", "Video",
     "See how an AI agent builds a full SEO audit in minutes. No setup headaches.",
     "Watch it work, then try it", "Free trial", "", "Was Demo_Screen_Walkthrough."),
    ("RT14d", "RT_Web14d_ExclTrials", "TESTIMONIAL_v1", "TESTIMONIAL", "Video",
     "Still thinking it over? Hear why agencies switched to AgentStack.",
     "Start your free trial", "Set up in minutes", "", "Was Retargeting_Testimonial."),
    ("RT14d", "RT_Web14d_ExclTrials", "DEMO_WALKTHROUGH_v1_RT", "DEMO_WALKTHROUGH", "Video",
     "You've seen the site — now see an AI agent build a full SEO audit in minutes.",
     "See it in action", "Free trial", "", "Reuses the demo concept for retargeting."),
]
for a in ads:
    assert len(a[5]) <= 125 and len(a[6]) <= 40 and len(a[7]) <= 30, a[2]
PREV = {"FOUNDER_STORY_v3_HookA": "UGC_Founder_v2", "FOUNDER_STORY_v3_HookB": "", "AGENCY_OWNER_v1": "UGC_Agency_Owner",
        "BENEFITS_STATIC_v1": "Static_AgentStack_Benefits", "DEMO_WALKTHROUGH_v1": "Demo_Screen_Walkthrough",
        "TESTIMONIAL_v1": "Retargeting_Testimonial", "DEMO_WALKTHROUGH_v1_RT": ""}
EXE = lambda name: next((p for p in name.split("_") if p.startswith("v") and p[1:].isdigit()), "")
HOOK = lambda name: "HB" if name.endswith("HookB") else "HA"
rows = [[a[0], a[1], a[2], a[3], EXE(a[2]), HOOK(a[2]), a[4], "", a[5], a[6], a[7], a[8], "",
         PREV[a[2]], "{{OLD_AD_ID}}" if PREV[a[2]] else "", PH, a[9]] for a in ads]
ws = mp.sheet("Meta_Ads", ["Campaign key", "Ad set name", "Ad name", "Concept ID", "Execution", "Hook code", "Format",
                           "Creative override (optional)", "Primary text", "Headline", "Description",
                           "Call to action (blank = concept's)", "Landing path override (optional)",
                           "Previous ad name", "Previous ad ID", "PLACEHOLDER", "Notes"],
              [13, 22, 26, 20, 9, 9, 8, 20, 58, 26, 22, 16, 18, 24, 16, 13, 36], rows,
              notes={"C": "Convention: CONCEPT_execution[_hook]. Reporting groups by concept → execution → hook → ad.",
                     "E": "Execution = the creative version (v3 = the third cut of this concept). Picks the matching files from Creative_Library.",
                     "F": "Hook code = the opening line variant. Kept separate so a weak new hook can't hide behind a winning one.",
                     "N": "The ad this one replaces. A new ad starts with no history in Meta; this column keeps the link for your own reporting.",
                     "O": "Copy from Ads Manager (Ad ID column)."})
mp.flag_len(ws, "I2:I200", 125, mp.ORG)
mp.flag_len(ws, "J2:J200", 40)
mp.flag_len(ws, "K2:K200", 30)

# ---------------- Variants (creative-test matrix) ----------------
V = {
    "CASE_STUDY": (["How one agency cut client reporting time with AI agents. See the numbers.",
                    "Reporting used to eat Fridays. Here's what changed when this agency handed it to AI agents."],
                   ["See the case study", "Real agency, real results"], "Start free trial"),
    "PROBLEM_SOLUTION": (["Too many clients, not enough hours? AI agents take the repetitive SEO work off your plate.",
                          "Your agency runs on busywork. AI agents can run it instead, so your team can do strategy."],
                         ["Get hours back every week", "Start your free trial"], "Start free trial"),
    "AGENCY_RESULTS": (["{{REAL CUSTOMER RESULT}}: what agencies get back when AI agents handle the busywork.",
                        "Agencies on AgentStack spend {{REAL STAT}} less time on audits. Try it free."],
                       ["Built for agencies", "Proof from real agencies"], "Start free trial"),
    "BEFORE_AFTER_WORKFLOW": (["Left: an SEO audit done step by step by hand. Right: the same audit, with an AI agent running the steps.",
                               "Same client audit, two workflows: your team doing every step, or an AI agent doing them while your team reviews."],
                              ["See the difference", "Let agents do the audit"], "Start free trial"),
}
rows = []
for con, (hooks, heads, desc) in V.items():
    for i, h in enumerate(hooks):
        assert len(h) <= 125; rows.append([con, "Hook", "H" + "AB"[i], h, PH])
    for i, h in enumerate(heads):
        assert len(h) <= 40; rows.append([con, "Headline", f"T{i + 1}", h, PH])
    rows.append([con, "Description", "D1", desc, PH])
ws = mp.sheet("Variants", ["Concept ID", "Element", "Code", "Text", "PLACEHOLDER"], [20, 12, 7, 90, 13], rows,
              notes={"B": "Hook = primary text, Headline, Description. The generator crosses every Hook with every Headline: 2 × 2 = 4 ads per concept.",
                     "C": "Short code used in the ad name, e.g. CASE_STUDY_HA_T2."})
mp.dv_list(ws, "B2:B500", ["Hook", "Headline", "Description"])

# ---------------- PMax_AssetGroups ----------------
AG = {
    "AG_Agencies": ("FOUNDER_STORY; AGENCY_OWNER", "/trial/agencies", "Agencies", "AI-Agents",
        ["AI Agents for Agencies", "Scale Without New Hires", "Automate Client Reporting", "Free Up Billable Hours",
         "Built for Agency Teams", "Audits on Autopilot", "Start Your Free Trial", "Try AgentStack Free",
         "More Clients, Same Team", "Agency Work, Automated", "AgentStack"],
        ["AI agents that handle audits, briefs and reports for every client account",
         "Grow your agency's output without growing headcount. Start a free trial."],
        ["Free trial for agencies.",
         "Hand repetitive client work to AI agents and give your team time back for strategy.",
         "Audits, briefs and reports run on autopilot across every client account.",
         "Built for agencies that want to grow output without growing headcount."]),
    "AG_Agentic_SEO": ("DEMO_WALKTHROUGH; BENEFITS_STATIC", "/trial/seo", "Agentic-SEO", "Free-Trial",
        ["Agentic SEO", "AI Agents That Do SEO Work", "Automate SEO Audits", "Content Briefs in Minutes",
         "Rank Tracking on Autopilot", "Built for SEO Teams", "Start Your Free Trial", "Try AgentStack Free",
         "AI SEO Agents", "Technical SEO, Automated", "Keyword Research in Minutes"],
        ["Let AI agents run technical audits, keyword research and content briefs for you",
         "Agentic SEO that works across every site you manage. Start your free trial today."],
        ["Try agentic SEO free.",
         "AI agents run audits, keyword research and briefs so your team can focus on strategy.",
         "Spend less time on SEO busywork and more time on work that moves rankings.",
         "Works across every client site you manage. Start your free trial today."]),
    "AG_Marketing_Automation": ("BENEFITS_STATIC; CASE_STUDY", "/trial", "AI-Marketing", "Agents",
        ["AI Marketing Agents", "Automate Marketing Workflows", "From Brief to Report Faster", "Connect Your Marketing Stack",
         "Launch Agents in Minutes", "Start Your Free Trial", "Try AgentStack Free", "Marketing Work, Automated",
         "Less Busywork, More Strategy", "AI Agents for Marketers", "Try It Free"],
        ["Create AI agents that handle research, content and reporting across your stack",
         "Automate the repetitive parts of marketing and keep your team on high-value work"],
        ["Launch your first agent today.",
         "AI agents handle research, content and reporting across your marketing stack.",
         "Automate the busywork and keep your team focused on strategy.",
         "See how agencies use AI agents to deliver more without hiring."]),
}
rows = []
for ag, (cons, path, p1, p2, H, LH, D) in AG.items():
    assert all(len(h) <= 30 for h in H) and any(len(h) <= 15 for h in H) and 3 <= len(H) <= 15, ag
    assert all(len(x) <= 90 for x in LH + D) and any(len(d) <= 60 for d in D), ag
    assert len(p1) <= 15 and len(p2) <= 15
    rows.append(["PMax", ag, "New (import Paused)", cons, path, p1, p2, "", "Automated"]
                + H + [None] * (15 - len(H)) + LH + [None] * (5 - len(LH)) + D + [None] * (5 - len(D)) + [PH])
rows.append(["PMax", "{{EXISTING_ASSET_GROUP_1}}", "Pause existing", "", "", "", "", "", ""] + [None] * 25 + [PH])
hdr = (["Campaign key", "Asset group", "Action", "Concept IDs (; separated)", "Final URL path", "Path 1", "Path 2",
        "Business name (blank = Settings)", "Call to action"] + [f"Headline {i}" for i in range(1, 16)]
       + [f"Long headline {i}" for i in range(1, 6)] + [f"Description {i}" for i in range(1, 6)] + ["PLACEHOLDER"])
ws = mp.sheet("PMax_AssetGroups", hdr, [10, 26, 18, 34, 16, 12, 12, 16, 12] + [24] * 15 + [44] * 5 + [44] * 5 + [13], rows,
              notes={"C": "New = create (imports Paused). Pause existing = name an asset group already in the campaign; the generator writes it as Paused.",
                     "D": "Images, videos and logos come from Creative_Library for these concepts (+ BRAND_LOGO). Google needs ≥1 landscape, ≥1 square image and ≥1 square logo.",
                     "J": "3–15 headlines, max 30 characters, at least one ≤15. Recommended 11+.",
                     "Y": "1–5 long headlines, max 90 characters.",
                     "AD": "2–5 descriptions, max 90 characters, at least one ≤60."})
mp.dv_list(ws, "C2:C100", ["New (import Paused)", "Pause existing"])
mp.flag_len(ws, "J2:X100", 30)
mp.flag_len(ws, "Y2:AH100", 90)
mp.flag_len(ws, "F2:G100", 15)

# ---------------- PMax_Signals ----------------
SIG = {
    "AG_Agencies": ["ai agents for agencies", "agency automation software", "client reporting automation",
                    "ai tools for marketing agencies", "white label seo automation"],
    "AG_Agentic_SEO": ["ai seo agent", "agentic seo", "seo audit automation", "content brief generator", "keyword research automation"],
    "AG_Marketing_Automation": ["ai marketing automation", "marketing workflow automation", "ai marketing agent",
                                "automated marketing reports", "ai content operations"],
}
rows = []
for ag, themes in SIG.items():
    rows += [["PMax", ag, "Search theme", t, PH] for t in themes]
    rows += [["PMax", ag, "Audience signal", a, PH] for a in
             ["Customer list: trial users", "Website visitors – 30d", "Custom segment: people searching competitor terms"]]
ws = mp.sheet("PMax_Signals", ["Campaign key", "Asset group", "Signal type", "Value", "PLACEHOLDER"], [10, 26, 16, 52, 13], rows,
              notes={"C": "Search themes and audience signals are HINTS to Google, not targeting limits. Audience names must match saved audiences in Google Ads.",
                     "D": "Brand terms don't belong here — the brand exclusion is set in Google Ads (manual step)."})
mp.dv_list(ws, "C2:C500", ["Search theme", "Audience signal"])

# ---------------- Budget: creative-test evidence check ----------------
ws = wb["Budget"]
ev = [
    ("Creative test evidence check", None, "Is this test big enough to name a winner this month?"),
    ("Test budget ($/month)", '=SUMIFS(Campaigns!J:J,Campaigns!C:C,"CreativeTest")', ""),
    ("Concepts in test", '=COUNTIFS(Meta_AdSets!A:A,"CreativeTest")', "One ad set per concept"),
    ("Max concepts allowed", f"={SET['Max test concepts per month']}", "Settings"),
    ("Planning CPA ($)", f"={SET['Target CPA for planning']}", "Settings"),
    ("Expected trials per concept / month", "=IFERROR(B15/B16/B18,0)", "Budget ÷ concepts ÷ CPA"),
    ("Trials needed to call a winner", f"={SET['Min trials to declare a winner']}", "Settings"),
    ("Weeks to reach that evidence", '=IFERROR(ROUNDUP(B20/B19*4.3,0),"-")', "Read results no earlier than this"),
    ("Verdict", '=IF(B16>B17,"Too many concepts – cut to the max",IF(B19>=B20,"Can name a winner this month","Directional in month 1; keep running to the weeks above"))', ""),
]
for i, (a, b, c) in enumerate(ev, 14):
    ws[f"A{i}"], ws[f"B{i}"], ws[f"D{i}"] = a, b, c
    for col in "ABD":
        ws[f"{col}{i}"].font = Font(name=F, bold=(i in (14, 22)))
    if i > 14:
        ws[f"B{i}"].fill = mp.CALC
ws["B15"].number_format = ws["B18"].number_format = "$#,##0"
ws["B19"].number_format = "0.0"

# ---------------- How_to_use: v2 section ----------------
ws = wb["How_to_use"]
v2 = [
    ("", False),
    ("Version 2 — concepts, creative library, PMax", True),
    ("Concepts is the hub: one row per creative idea with its audience, problem, message, landing page and CTA. Ads, files, test variants and PMax asset groups all point at a Concept ID.", False),
    ("Creative_Library lists every file, named CONCEPT_version_ratio.ext. The generator picks files by concept and ratio: 4x5 (else 1x1) for Meta Feed, 9x16 for Stories/Reels, 1.91:1 + 1x1 + 4x5 images and logos for PMax.", False),
    ("Changing a library row changes the NEXT export only. Ads already live don't change until you re-import.", False),
    ("Variants builds the creative test: every Hook × every Headline for each test concept (2 × 2 = 4 ads). Results are judged per concept, not per ad.", False),
    ("Budget tab → Creative test evidence check shows whether the test can name a winner this month (needs the Settings minimum, default 20 trials per concept).", False),
    ("PMax is rebuilt IN PLACE: fill 'Existing campaign name' on Campaigns so the import edits the live campaign instead of creating a new one. New asset groups import Paused; switch them on and pause old ones together.", False),
    ("Settings → 'Google tracking verified on' must have a date before the PMax export unlocks.", False),
    ("Manual steps: PMax brand exclusion (Google Ads), attaching PMax images/videos if Editor's import doesn't carry them, Meta Flexible format (optional).", False),
    ("Files ending _TEST are generated stand-ins marked 'TEST FILE – NOT APPROVED'. Never launch them.", False),
    ("Claims: any concept whose copy states a result or statistic needs an Approved claim (source, approver, date) on the Claims tab, or it can't export. BEFORE_AFTER_WORKFLOW is the no-claims fallback for a blocked test slot.", False),
    ("Budgets: Google daily = monthly ÷ 30.4 (Google's monthly cap); Meta daily = monthly ÷ days in month.", False),
    ("PMax deploys in two stages: Stage 1 adds the new asset groups Paused; after review approval, Stage 2 switches them on, pauses the old ones and sets the budget in one post.", False),
]
for t, b in v2:
    ws.append([t])
    c = ws.cell(ws.max_row, 1); c.font = Font(name=F, bold=b); c.alignment = Alignment(wrap_text=True)

# ---------------- order + save ----------------
order = ["How_to_use", "Settings", "Campaigns", "Budget", "Concepts", "Claims", "Creative_Library", "Google_Keywords",
         "Google_Negatives", "Google_RSAs", "PMax_AssetGroups", "PMax_Signals", "Meta_AdSets", "Meta_Ads", "Variants"]
wb._sheets = [wb[n] for n in order]
wb["Settings"]["A1"].parent.title  # sanity
wb.save(OUT)
print("saved", OUT)

# export the library spec for the test-asset generator
import json
json.dump([r[:5] for r in lib_rows], open(os.path.join(mp.ROOT, "build", "library_spec.json"), "w"))
