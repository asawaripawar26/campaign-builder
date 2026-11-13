# What the AI got wrong, and how I caught it

I built this with Claude, in a plan → discuss → approve → build loop: Claude proposed each part, I reviewed it, and nothing was built until I approved it. Most of my time went into reviewing, and that's where these came from.

> **DRAFT: the lines marked ⟨confirm⟩ need Asawari's own words before this is submitted.**

## The headline example: every Google daily budget was wrong

**What the AI did:** it calculated every daily budget as monthly ÷ days in the month, for both platforms. That is right for Meta. It is wrong for Google, which caps a month's spend at the average daily budget × 30.4.

**Why it mattered:** with ÷ 31, Brand search got $290 a day instead of $296. Across the four Google campaigns, about **$1,100 of the $55K would never have been spent**, without any error message. PMax alone would have spent $11,765 of its $12K.

**How I caught it:** ⟨confirm: what you noticed and how you checked it⟩

**The fix, and how it stays fixed:**
- The plan has a "Google billing days per month" setting (30.4).
- Google daily = monthly ÷ 30.4; Meta daily = monthly ÷ days in the month.
- An automated test checks every daily budget. Putting the old formula back makes it fail ([tests/generator.test.js](../tests/generator.test.js)).

## Everything else I corrected or added

| # | What the AI got wrong or missed | What I did | What changed |
|---|---|---|---|
| 1 | A factual error: ⟨confirm: what it was⟩ | ⟨confirm⟩ | ⟨confirm⟩ |
| 2 | Treated two PMax campaigns as bidding against each other | I pointed out that campaigns in the same account don't bid against each other in the auction | The reasoning for the PMax change was rewritten on the right basis ⟨confirm⟩ |
| 3 | Planned to rebuild PMax as a new campaign | I asked for it to be rebuilt **in place**, keeping the existing campaign and its history | PMax is now edited in place. The generator proves from a Google Ads Editor export that it's editing, not creating, and splits the change into two stages with a checklist |
| 4 | Said "replace a creative once and every ad updates" | I flagged that this overstated it | Now stated correctly: a change in the creative library affects the **next export only**. Live ads don't change until you re-import |
| 5 | Proposed a "2× CPA" rule for judging results ⟨confirm: what the rule was⟩ | ⟨confirm: why you challenged it⟩ | Results are now judged on evidence: a cost-per-trial range from the trial count, a minimum spend of 3× the planning cost per trial, and 20 trials before calling a winner |
| 6 | The $12K PMax budget ⟨confirm: what was wrong⟩ | ⟨confirm⟩ | PMax is $12K a month, $395 a day (÷ 30.4) |
| 7 | No link between ads and the creative idea behind them | I asked for a **concept map** | Concepts became the hub: every ad, test variant and PMax asset group points to a concept, and reporting rolls up by concept → hook → ad |
| 8 | No place to record how each result is measured | I asked for a **measurement field** | ⟨confirm: where it lives now⟩ |
| 9 | The build stopped at launch | I asked for a **post-launch performance sheet** | A performance workbook that compares platform numbers with PostHog, shows cost-per-trial ranges, and recommends (never executes) scale, reduce or pause |

## Where the AI pushed back on me

⟨confirm: the example you want to show⟩

## What this taught me about working with AI

- **Check the arithmetic, not just the logic.** The budget bug read perfectly sensibly. It was wrong only against how Google actually bills.
- **Ask "what does this overwrite?"** Rebuilding PMax would have thrown away its learning history. Editing in place keeps it.
- **Turn every catch into a test,** so the same mistake can't come back unnoticed.
