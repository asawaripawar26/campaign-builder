# How the campaign builder works

The builder turns a campaign plan (one Excel workbook) into **paused, import-ready** Google Search, Performance Max (PMax) and Meta campaigns. It needs no API access or ad-account login: Google Ads Editor and Meta Ads Manager both import campaigns from a CSV file, so the builder's job is to write those files correctly and refuse to write them when something is unsafe.

Everything runs in the browser. Plan and creative files are read on your computer and never uploaded anywhere.

## The pieces

| Piece | What it is | You touch it… |
|---|---|---|
| **Plan workbook** (`plan/Campaign_Plan_Dec2026.xlsx`) | One workbook per month. Yellow cells are inputs; grey cells are formulas | Every launch |
| **Generator** (`index.html`, also on GitHub Pages) | Reads the plan, checks it, maps creative files, estimates test evidence and writes the import files | Every launch |
| **Google Ads Editor export** (CSV of the live PMax campaign) | Proves the PMax import will edit the existing campaign, not create a duplicate | Every PMax change |
| **Creative files** (optional upload) | Checked in your browser for shape, size and type | When creatives change |
| **Import files** | `google_search_…`, `meta_bulk_import_…`, `pmax_stage1_…`, `pmax_stage2_…`, plus `ad_map_…` for reporting | Every launch |

## Monthly workflow

1. **Duplicate last month's plan.** On *Settings*, change Month code, Start date and End date.
2. **Budgets** go on *Campaigns* (monthly amounts). Check the *Budget* tab: the Difference row must be 0, and the creative-test evidence check shows whether the test can name a winner this month. Daily budgets are calculated for you:
   - **Google:** monthly ÷ 30.4, because Google caps a month's spend at the average daily budget × 30.4.
   - **Meta:** monthly ÷ days in the month.
3. **Concepts** is the hub: one row per creative idea, with its audience, problem, message, landing page, call to action, status and any claim it makes.
4. **Creative_Library** lists each file as `CONCEPT_version_ratio.ext`. Only *Approved* files belong in a real launch.
5. **Copy:**
   - Meta_Ads holds the always-on ads.
   - Variants holds the creative test (every hook × every headline).
   - PMax_AssetGroups holds the PMax text.
   - The Google tabs hold Search keywords, negatives and ads.
6. **Claims:** any copy that states a result or statistic needs an Approved claim, with a source, an approver and a date.
7. **Open the generator:**
   1. Upload the plan.
   2. Add the creative files.
   3. Upload a Google Ads Editor export of the PMax campaign.
   4. Fix every ERROR, and review what is HELD BACK.
8. **Import:**
   - **Search:** Editor → *Get recent changes* → *Import → From file* → review → *Post*.
   - **Meta:** upload creatives to the media library first, then Ads Manager → *Import ads in bulk*.
   - **PMax:** two stages (below).
9. **Download the ad map** for reporting (it links every Meta ad to its concept and hook).
10. **Turning campaigns on stays a human decision.** Everything new arrives Paused.

## What the generator does, step by step

1. **Read and recalculate.** Tabs are read by header name, not position, so adding a column doesn't break anything. Names and daily budgets are rebuilt from your inputs, so stale spreadsheet values can't leak into a launch.
2. **Gates** (rules that stop an export).
   - A concept can't export if it's Retired, or if its claim isn't approved.
   - Test ad sets accept only *Ready for testing* or *Testing* concepts. Always-on placements expect *Validated* (anything else gets a warning).
   - Status lifecycle: Hypothesis → In production → Ready for testing → Testing → Validated → Retired.
3. **Creative mapping.** Files are chosen by concept, version and shape:

   | Placement | Files used |
   |---|---|
   | Meta Feed | 4:5, or 1:1 if there's no 4:5 |
   | Meta Stories / Reels | 9:16 (listed; placement customization is a manual step) |
   | PMax images | 1.91:1 landscape, 1:1 square, 4:5 portrait |
   | PMax logos and video | logos 1:1 and 4:1; videos 16:9, 9:16 and 1:1 |

   A library change affects the **next export only**. Ads already live don't change until you re-import.
4. **Variant matrix.** Test ads are named `CONCEPT_version_hook_headline` (e.g. `CASE_STUDY_v2_HB_T1`). Reporting groups results by concept → version → hook → ad, so a weak new hook can't hide behind a winning one.
5. **Evidence estimate (before launch).** At $10K, a $265 cost per trial and 3 concepts, expect about 12.6 trials per concept, about 6 per hook and about 3 per ad in a month.
   - The page also shows the **cost-per-trial range**: even a concept that truly costs $265 could read anywhere from about $165 to $453 in month 1, just from chance.
   - "Weeks to the threshold" (about 7) is only the time to collect 20 trials. It is not proof.
   - A low-volume ad isn't a loser.
6. **File checker.** It reads files in your browser and flags:
   - files shaped wrongly for their listed ratio (it reads the real pixels, not the file name)
   - files below a platform's minimum size
   - PMax images over 5,120 KB, and PMax videos shorter than 10 seconds
   - Meta's limits (images 30 MB, videos 4 GB)
   - a file whose real type doesn't match its extension (e.g. a JPEG saved as .png)
   - identical files under different names (a likely mislabelled copy)
   - library files you haven't added, and added files that aren't in the library
7. **Tracking tags (UTMs).** `utm_source` stays `google` / `facebook` so results compare with PostHog history. `utm_content` carries the Meta ad name, so PostHog can count trials per ad and per concept, provided it stores UTMs on the trial event.
8. **Errors are scoped.**
   - **ERROR** blocks that platform's file (e.g. a missing Meta Page ID).
   - **HELD BACK** leaves out only the item named (a campaign, ad set, ad, keyword or file). The rest still exports.
   - In the creative test, if any ad in a concept is held back, that concept's whole ad set is held back, so the test stays fair.
   - **PMax is all-or-nothing.**
9. **Draft export.** A dry run for a test import: it allows TEST files and placeholder IDs, and drops audiences that have no ID. Never launch a draft.

## PMax: edit the live campaign in place, in two stages

The existing PMax campaign keeps its history, so it is edited rather than recreated.

**Checked against your Editor export before any PMax file unlocks:**
- the campaign name exists exactly once (near-misses that differ only in capital letters or spacing are flagged)
- every asset group marked "Pause existing" exists in that campaign
- no new asset group name collides with an existing one
- the Google tracking verification date is filled in

**The campaign row carries only the name and budget.** In Editor a blank field means "no change", so bidding, targeting, tracking and conversion settings are left alone.

| Stage | What the file does | Before moving on |
|---|---|---|
| **1. Add** | New asset groups (Paused) plus their search themes. Campaign: no changes | Attach images, logos and videos (the page lists which files go where). Post |
| **Checks** | — | Stage 1 preview updated the existing campaign, created no new one, and listed the new groups as new. New assets show *Approved*, brand exclusion applied, conversion goals unchanged. Ticking all five unlocks Stage 2 |
| **2. Switch** | Budget set to monthly ÷ 30.4, new groups Enabled, old groups Paused, all in one post | Only budget and asset-group statuses changed |

## One-time setup before the first real launch

- **Meta IDs:** Page ID, Instagram account ID and Pixel ID, plus audiences written as `ID:Name`.
- **PMax:** the exact existing campaign name, the asset groups to pause, and the Google tracking verification date.
- **Approved creative files** to replace the `_TEST` stand-ins.
- **Real copy, domain, keywords and competitor names**, and the negative keywords from the live account. The product name (AgentStack), keywords, copy and landing pages in the sample are placeholders inferred from the assessment data.
- **Column spellings to verify on the first real import:** Google `Campaign Negative Phrase`, `Maximize conversions` and the PMax columns; Meta `Outcome Sales`, the conversion event value, the date format and the audience separator.

## Not automated (by design)

- **PMax:** brand exclusion, audience signals, and attaching images and videos if Editor's import doesn't carry them.
- **Meta:** Flexible ad format, 9:16 placement customization, and carousel ads.
- **Budget:** the $11K Meta reserve, released only when its conditions are met (tracking verified, signup funnel stable, PostHog cost per trial ≤ $280 for 7 days).
- **Pausing, scaling or launching** anything. People decide.

## Extending it

- **New country:** add it to `LOC_IDS` (Google) and `COUNTRY` (Meta) in `src/generator_template.html`.
- **New call to action:** add it to `CTAS`. **New shape:** add it to `RATIO` and `MIN`.
- **New concept:** add a row on Concepts, add its files to the library, and point ads or variants at it.
- After changing the template or the sample plan, rebuild: `python3 scripts/build_generator.py`, then `npm test`.

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| PMax buttons greyed out | A PMax gate failed | Filter "What to fix" to PMax. Usual causes: no Editor export, a placeholder campaign name, an empty tracking date |
| Stage 2 greyed out | Checklist not ticked | Do the five checks, then tick them |
| "Would create a NEW campaign" | Name mismatch with the export | Copy the name exactly, including capital letters and spaces |
| An ad maps to "no file" | No library file for that concept, version and shape | Add the file, or set a creative override |
| A test concept is blocked | Status isn't ready, or its claim isn't approved | Finish it, or swap in a Ready-for-testing concept |
| Download didn't start | Browser blocked it | Use the text version shown on the page |
