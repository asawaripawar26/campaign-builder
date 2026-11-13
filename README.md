# Campaign Builder: from a plan workbook to import-ready Meta and Google campaigns

[![Tests](https://github.com/asawaripawar26/campaign-builder/actions/workflows/test.yml/badge.svg)](https://github.com/asawaripawar26/campaign-builder/actions/workflows/test.yml)

**Try it:** https://asawaripawar26.github.io/campaign-builder/ (or download this repo and open `index.html`)

You fill in one Excel workbook with next month's campaigns: budgets, audiences, keywords, ads, creative files. The generator checks it and writes the files that **Google Ads Editor** and **Meta Ads Manager** import in bulk. Everything arrives **paused**, so a person still decides what goes live.

No API keys or ad-account logins are needed. Both platforms already accept campaigns as CSV uploads, so the work is in writing those files correctly, and refusing to write them when something is unsafe.

```
plan workbook (.xlsx) ──► generator (index.html, runs in your browser) ──► google_search_202612.csv
  + creative files            • rebuilds names and daily budgets              meta_bulk_import_202612.csv
  + Google Ads Editor export  • checks claims, statuses, files, tracking       pmax_stage1 / pmax_stage2.csv
                              • holds back only what's broken                 ad_map_202612.csv (for reporting)
```

The sample plan is the December build from Task 1 of the assessment: 7 campaigns plus an $11K Meta reserve, $120K in total.

After launch, the [performance workbook](performance/) judges the results: it checks platform numbers against PostHog, shows how much evidence each result has, and recommends scale, reduce or pause.

## Next launch in 5 steps

1. **Copy the plan.** Download [`plan/Campaign_Plan_Dec2026.xlsx`](plan/Campaign_Plan_Dec2026.xlsx) (or last month's copy). On the *Settings* tab, change the month code and the start and end dates.
2. **Change what's different this month.** Budgets on *Campaigns*, new creative ideas on *Concepts*, files on *Creative_Library*, copy on the ad tabs. Yellow cells are yours to edit; grey cells calculate themselves.
3. **Open the generator** and upload the plan. Add the creative files, and for Performance Max, a Google Ads Editor export of the live campaign.
4. **Fix what it flags.** ERROR blocks a platform's file. HELD BACK drops just the named item. Tick *Draft export* if you want a dry run first.
5. **Download and import.** Google Search CSV → Ads Editor → *Import → From file*. Meta CSV → Ads Manager → *Import ads in bulk*. PMax in two stages, following the checklist on the page. Then switch on what you've reviewed.

The full guide, including every check the generator runs, is in [docs/HOW_IT_WORKS.md](docs/HOW_IT_WORKS.md).

## How it works, in short

- **The plan is the single source of truth.** One row per campaign, ad set, ad, keyword and creative idea. The generator reads tabs by their header names, so adding a column doesn't break it.
- **Budgets are recalculated, not copied.** Google daily = monthly ÷ 30.4 (Google caps a month at daily × 30.4). Meta daily = monthly ÷ days in the month.
- **Concepts are the hub.** Every ad points to a creative idea (concept). The concept supplies the landing page and call to action, and its status and claims decide whether it can export. An unapproved claim ("cut reporting time by 40%") can't reach an ad.
- **Performance Max is edited in place.** The existing campaign keeps its history. The generator proves from your Editor export that it is editing that campaign, not creating a second one, and splits the change into two stages with a checklist in between.
- **Tests are judged by concept, not by single ads.** The page estimates how many trials each concept, hook and ad will get, so nobody calls a winner on 3 trials.

## What's in this repo

| Path | What it is |
|---|---|
| `index.html` | The generator. One self-contained page with the sample plan built in |
| `plan/Campaign_Plan_Dec2026.xlsx` | The plan workbook: the template for every future month, filled in with the Task 1 build |
| `examples/` | What the generator produced from the sample (draft mode), so you can open the files without running anything |
| `src/generator_template.html` | The generator's source, before the sample plan is embedded |
| `scripts/` | `build_plan.py` builds the workbook, `recalc.py` fills in its formula results, `build_generator.py` builds `index.html`, `export_examples.js` refreshes `examples/`, `make_test_assets.py` makes the test images and videos |
| `tests/` | Automated browser tests, with a fictional plan, a fictional Editor export and stamped TEST creative files |
| `performance/` | The performance workbook: judges results after launch and recommends scale, reduce or pause ([how it works](performance/README.md)) |
| `docs/` | How it works: every check the generator runs, setup, troubleshooting |

## Proof it runs, without a live ad account

The assessment account isn't real and I can't sign in to a Meta or Google Ads account, so the import itself hasn't been tried. Instead:

- **Automated tests** run on every change ([GitHub Actions](https://github.com/asawaripawar26/campaign-builder/actions)). They open the generator in a real browser (headless Chromium), load plans, upload files, click the export buttons and check the downloaded CSVs. They cover:
  - the sample plan matches Task 1 (budgets, $109K committed, inside $120K)
  - daily budgets: Google ÷ 30.4, Meta ÷ days in the month
  - naming follows `PLATFORM_FUNNEL_KEY_GEO_MONTH`
  - strict export refuses TEST creative, missing tracking checks, and PMax without an account export
  - draft export downloads real CSVs with the right columns and numbers
  - the plan template round-trips (uploading the .xlsx gives the same result as the built-in sample)
  - Stage 2 of PMax stays locked until the checklist is done
  - a wrongly spelled campaign name, and an image in the wrong shape, are both caught
- **The tests can fail.** Changing the Google budget rule to ÷ 31 makes two tests fail.
- **Example output** is in `examples/`.

**Not yet verified:** a few column spellings only a real import confirms (listed in the guide). The first real launch should be a draft import reviewed in Editor's preview before anything is posted.

### Run the tests yourself

```bash
npm install
npx playwright install chromium
npm test                              # generator
python3 tests/test_performance.py     # performance workbook (needs openpyxl; LibreOffice for the full recalculation)
```

### Rebuild after changing the plan or the generator

```bash
pip install openpyxl pillow          # LibreOffice is also needed for recalc.py
python3 scripts/build_plan.py        # writes plan/Campaign_Plan_Dec2026.xlsx
python3 scripts/recalc.py plan/Campaign_Plan_Dec2026.xlsx
python3 scripts/build_generator.py   # writes index.html
npm test
```

## Placeholders

The product name (AgentStack), keyword themes, ad copy and landing-page paths are inferred from the ad names in the assessment data. They are marked `PLACEHOLDER` in the plan or written as `{{LIKE_THIS}}`. Strict export refuses to run until real values replace them.
