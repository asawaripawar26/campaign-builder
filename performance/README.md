# Performance workbook

The second half of the build: the generator launches campaigns, this workbook judges them. You paste in exports from Meta, Google and PostHog each week; it checks the numbers can be trusted, shows how much evidence each result has, and recommends what to do. It never pauses or scales anything itself. A person decides and records the decision.

`Performance_Workbook_AugOct2026.xlsx` is filled with the assessment's Aug 10 – Oct 11 data, so you can see it working.

## What it does

| Tab | What it shows |
|---|---|
| **Summary** | The period at a glance, baselines, tracking health, and the reallocation suggestion |
| **Decisions** | One recommendation per campaign and per creative concept, with the reason in plain words |
| **Platform** | Platform-reported trials vs PostHog, week by week. A run of "Diverging" weeks means tracking changed, not performance |
| **Campaigns** | Spend, cost per trial, its range, evidence stage, and the 14-day trend |
| **Creative** | Ads → concepts → hooks, so a weak new hook can't hide behind a winning one |
| **Funnel** | Sessions, activation, trial-to-paid (mature trials only), acquisition cost and payback, when that data is pasted in |
| **Pacing** | Planned vs actual spend to date, projected month, and the daily budget needed (Google ÷ 30.4, Meta ÷ days) |
| **Decision_Log** | What a person decided. It marks an OVERRIDE when the decision differs from the recommendation, and asks why |
| **Settings** | Targets and thresholds: $250–280 target, 20 trials, 3× minimum spend, 20% weekly step, and so on |

## The rules it applies

- **Cost-per-trial range:** a 90% interval from the trial count, because small numbers swing by chance.
- **Evidence stages:** not launched → below minimum spend (3× the planning cost per trial) → early signal (5+ trials) → sufficient evidence (20+ trials).
- **Scale candidate:** even the worst case is within $280, trial quality is OK, and there's proof it can absorb more. Step up at most 20% a week.
- **Reduce or refresh:** even the best case is above $280, and it's getting worse.
- **Pause candidate:** zero trials after 3× the planning cost per trial. Check tracking and conversion lag first.
- **Inconclusive:** everything else. A low-volume campaign isn't a loser.
- **Platform and PostHog numbers are never added together.**

## What it concluded on the Aug–Oct data, by itself

| Campaign or signal | Call |
|---|---|
| UK search | Pause candidate: $15,090 spent, 0 trials |
| Meta Broad prospecting, Retargeting | Reduce or refresh |
| PMax | Investigate: best case $287, above target |
| US non-brand search | Inconclusive: range $270–358 straddles $280 |
| FOUNDER_STORY concept | Winner: 132 trials at $200, range $173–233 |
| TESTIMONIAL concept | Retire or rework |
| Meta tracking | Diverging from PostHog for 4 weeks |

`tests/test_performance.py` checks these calls on every change. On GitHub it also throws away every stored result and recalculates all formulas from scratch in LibreOffice, so the test proves the formulas work.

## Using it for a live month

1. Copy the workbook and clear the yellow `IN_` tabs.
2. Paste the exports: `IN_Meta_Campaign` and `IN_Meta_Ads` (Ads Manager, by day), `IN_Google` (by day; add "Search lost IS (budget)" to test headroom), `IN_PostHog` (trials by `utm_source`; add `utm_campaign`, `utm_content`, Activated and Paid when you have them), `IN_Plan` (campaign names and planned budgets) and `IN_AdMap` (the generator's ad map).
3. Set the period and as-of date on *Settings*.
4. Read *Summary*, then *Decisions*. Record what you decide in *Decision_Log*.

To rebuild it from code: `python3 scripts/build_performance.py <data.xlsx>`, then `python3 scripts/recalc.py build/Performance_Workbook.xlsx`. The assessment's raw data file isn't included in the repo; the filled workbook carries the same data on its `IN_` tabs.
