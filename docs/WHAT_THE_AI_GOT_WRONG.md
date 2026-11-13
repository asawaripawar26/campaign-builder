# What the AI got wrong, and how I caught it

I built this with Claude, in a plan → discuss → approve → build loop: Claude proposed each part, I reviewed it, and nothing was built until I approved it. Most of my time went into reviewing, and that's where these came from.

## The headline example: every Google daily budget was wrong

**What the AI did:** it calculated every daily budget as monthly ÷ days in the month, for both platforms. That is right for Meta. It is wrong for Google, which caps a month's spend at the average daily budget × 30.4.

**Why it mattered:** with ÷ 31, Brand search got $290 a day instead of $296. Across the four Google campaigns, about **$1,100 of the $55K would never have been spent**, and nothing would have flagged it. PMax alone would have spent $11,765 of its $12K.

**How I caught it:** reviewing the PMax budget, I checked how $12K a month translates into the daily budget Google Ads actually uses: about $395 a day (÷ 30.4), not the ÷ 31 figure the plan showed. The same rule applied to every Google campaign.

**The fix, and how it stays fixed:**
- The plan has a "Google billing days per month" setting (30.4).
- Google daily = monthly ÷ 30.4; Meta daily = monthly ÷ days in the month.
- An automated test checks every daily budget. Putting the old formula back makes it fail ([tests/generator.test.js](../tests/generator.test.js)).

## Everything else I corrected or added

| # | What the AI got wrong or missed | How I caught it | What changed |
|---|---|---|---|
| 1 | **A factual error:** it said the UK search campaign generated one conversion | I checked the Google campaign data: **zero** conversions on $15,089.50 of spend | The plan states it correctly: $15,090 spent, 0 conversions, so UK search is paused |
| 2 | **(My own check)** Where the US spend went after the UK launch | I compared daily US spend before and after the Sep 15 UK launch: US spend fell about 31% while the UK campaign took roughly $559/day of the shared $1,800/day budget | The plan gives every Google campaign its own budget, so one campaign can't drain another |
| 3 | Treated two PMax campaigns as bidding against each other | I pointed out that campaigns in the same account don't bid against each other in the auction | That argument was dropped. The PMax change rests on its own numbers: $303–338 per trial, above target |
| 4 | Planned to rebuild PMax as a new campaign | I asked for it to be rebuilt **in place**, keeping the existing campaign and its history | PMax is edited in place. The generator proves from a Google Ads Editor export that it's editing, not creating |
| 5 | Said "replace a creative once and every ad updates" | I flagged that this overstated it | Now stated correctly: a creative library change affects the **next export only**. Live ads don't change until you re-import |
| 6 | **The "2× CPA" rule:** spend about $500–560 per ad before judging it | That buys only about two trials per ad, far too little evidence to call a winner or loser | Tests are judged by concept, not single ads. The generator shows expected trials per concept, hook and ad, and a cost-per-trial range; 20 trials before calling a winner |
| 7 | **The $12K PMax budget** came without conditions | The amount was defensible, but keeping that much behind a campaign with a historical $338 cost per trial needed stronger performance conditions | PMax was cut from about $16K to $12K, rebuilt in place, with tracking verification and a two-stage checklist gating it |
| 8 | No link between ads and the creative idea behind them | I asked for a **concept map** | Concepts became the hub: every ad, test variant and PMax asset group points to a concept, and reporting rolls up by concept → hook → ad |
| 9 | No explicit record of what each campaign is measured on | I asked for a **measurement field** | Every campaign now names its conversion event. Tracking must be verified (PMax is blocked and Meta warns until a verification date is entered), and funnel quality is reported after launch |
| 10 | The build stopped at launch | I asked for a **post-launch performance sheet** | A performance workbook that compares platform numbers with PostHog, shows cost-per-trial ranges, and recommends (never executes) scale, reduce or pause |

## Where the AI pushed back on me

When I asked to rebuild PMax in place, Claude challenged me on the technical consequences: an import could change the live campaign's status, and new asset groups had to arrive paused so nothing changed before they were reviewed. That turned my idea into a much safer implementation:

- The campaign row carries only the name and budget, so bidding, targeting and conversion settings stay untouched.
- Stage 1 adds the new asset groups, paused.
- A five-point checklist (including "the preview created no new campaign") must be ticked.
- Only then does Stage 2 switch the new groups on and pause the old ones, in one step.

## What this taught me about working with AI

- **Check claims against the raw data.** "One conversion" sounded plausible; the data said zero.
- **Check the arithmetic against how the platform actually works.** The ÷ 31 budget read sensibly; it was wrong only against how Google bills.
- **Ask how much evidence a rule really buys.** "2× CPA per ad" sounds rigorous until you count the trials.
- **Turn every catch into a check,** so the same mistake can't come back unnoticed.
