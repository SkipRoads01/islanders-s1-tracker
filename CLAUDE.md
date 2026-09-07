# CLAUDE.md — Islanders S1 Franchise Tracker

Operating manual for maintaining the hockey franchise stat workbook in Claude Code.
Same architecture as the baseball tracker at `~/Documents/GitHub/PS5 Baseball` — read
that project's `CLAUDE.md` if anything here is ambiguous about *mechanics*; this file
owns everything sport-specific and **wins on any conflict.**

Read this file fully before touching the workbook.

---

## 1. What this project is

A game-by-game statistical record of a **PS5 hockey franchise** (New York Islanders,
**Season 1**), played manually —
every game, no sims. The user pastes handwritten play-by-play notes; you parse them and
inject the results into a multi-sheet Excel workbook that serves as the permanent data
source, then regenerate and publish the site.

- This is a **video-game franchise, not the real NHL.** Never conflate the two. Engine
  artifacts are not anomalies to "correct."
- Preseason and Regular Season are **separate files** and must never be merged.

---

## 2. Files & layout

```
PS5 Hockey/
├── S1 NY Islanders.xlsx  # THE workbook — exact filename, set once and never renamed
├── CLAUDE.md             # this file
├── add_game.py           # injection engine — edit the GAME block, run it
├── scripts/              # reset_season.py and other helpers
└── site/                 # build_site.py, deploy.sh, chrome files, logos
```

- Commit the workbook to git after each game. Git history is the versioning.
- **Recalc is mandatory every game.** The workbook stores formulas, not cached values,
  until a headless LibreOffice pass bakes them:
  ```bash
  libreoffice --headless --calc --convert-to xlsx --outdir /tmp "<workbook>.xlsx"
  cp "/tmp/<workbook>.xlsx" "<workbook>.xlsx"
  ```
  Skipping it yields blank previews and a site built from empty cells.

---

## 3. Workbook architecture

**The sheet list is the user's call too — it is not settled.** The baseball workbook's
shape is the model, not a template to copy blindly: an `Overview`/`Summary` pair driven
by formulas, one hand-typed game log per unit, cumulative rollups that auto-sum, a
per-opponent set, and the supporting sheets (`Schedule`, `Teams`, `Roster Ref`,
`Recaps`, `Headlines`, `Box`, `Notes`). Build a sheet when he says what goes in it.

### Hand-typed vs. derived
Whatever the sheets end up being, the split holds: **game logs are typed, rollups are
formulas.** Never hand-edit a number on a sheet that recalculates.

### These recalc automatically — never hand-edit their numbers
`Overview`, `Summary`, `Team Skating`, `Team Goaltending`, `Skating vs Opp`,
`Goaltending vs Opp`. Driven by `SUMIFS`/`COUNTIFS`/`INDIRECT` over the game logs.

### `Opp Skating` is hybrid
Opponent cumulative skating totals, one row per opponent plus an `ALL OPP` row.
Counting columns are hand-entered; rate columns are formulas — leave them.

### `Opp Goaltending` is a plain game log
One row per opposing goalie per game. No formulas, no cumulative sheet, **no expansion
rule** — the site aggregates it directly. Just append rows. Mark `Catches` L/R;
**right is the default and `(L)` in the notes is the only marker.**

### The expansion rule (this is what bites)
The cumulative sheets auto-sum **only for names already in column A/B.**
- New skater/goalie → add a row to the relevant Team sheet (copy formulas from the row
  above, shift the row reference, rewrite the TEAM row's `SUM` ranges) **and** add the
  player's per-opponent row(s) in the vs-Opp sheet.
- First game vs a new opponent → per-opponent rows for each player, plus a new opponent
  row in `Opp Skating`.
- Known player vs known opponent → nothing to do; recalc picks it up.

### Row 2 is a blank styled template
Data starts at **row 3** on every log sheet. Row 2 is blank but styled — it's what
`add_game.py` copies formatting from. Don't tidy it away.

---

## 4. Exact schemas

**The user defines the columns. Do not invent them.** Every header on every input sheet
comes from him — either stated directly or evident from what his notes actually track.
Until he has given a sheet's columns, that sheet does not exist. Never add, rename,
reorder, or "improve" a header on your own initiative; if a stat shows up in the notes
that has no home, log what you can, **flag it, and ask** rather than opening a column
for it.

Once a schema is set, record it here verbatim — column order, exact spelling, and the
sanity checks that go with it — the way the baseball manual does. Confirm the live
header row before every write; `max_row` lies, so find the last data row by scanning
column A for the last numeric value.

| Sheet | Columns |
|---|---|
| `Games` | *awaiting the user* |
| `Skater Game Log` | *awaiting the user* |
| `Goalie Game Log` | *awaiting the user* |
| `Opp Goaltending` | *awaiting the user* |
| `Opp Skating` | *awaiting the user* |

---

## 5. Notation reference (user's shorthand)

**To be filled in from the first games.** The user will identify anything new as it
happens — when he does, add it here immediately so it becomes permanent.

Carried over from the baseball project because they're habits, not sport rules:
- Scores in parentheses are **leader-first**: `(3-2)` = leader has 3.
- Per-player summary lines at the end of notes are **gold** — verify the play-by-play
  against them. If a summary implies an event the play-by-play doesn't contain, **do not
  fabricate it**; log what the play-by-play supports and flag the discrepancy.
- Period labels can be mislabeled — verify by *which team's players are listed*.
- `(L)` marks a left-handed/left-catching player; unmarked means right. Never ask.

---

## 6. Per-game workflow

1. **Inspect the live schema** — confirm header rows; find the last data row by scanning
   column A for the last numeric value (`max_row` lies; trailing styled blanks inflate it).
2. **Append skater rows** to `Skater Game Log`, styling copied from the row above.
3. **Append goalie rows** to `Goalie Game Log`.
4. **Append opposing goalie rows** to `Opp Goaltending`.
5. **Insert the Games row ABOVE the totals row** (the first row whose column-A value is a
   formula) so the `INDIRECT` totals self-extend.
6. **Write the Streak formula explicitly** in the new Games row, referencing the row
   above — per whatever streak rule the user has set (§7).
7. **Expansion** if needed (§3).
8. **Update `Opp Skating`** — increment the opponent's counting cells, re-sum `ALL OPP`.
9. **Recalc + bake** (§2). Zero error cells before shipping.
10. **Rebuild and deploy the site**, not just the workbook.
11. **Summarize** with a tight box score.

---

## 7. Standing judgment calls

Empty until the user makes them. Hockey has several that baseball doesn't — how an
overtime or shootout loss affects a streak, whether empty-net and shootout goals count
where, how TOI is formatted — and **each one is his ruling, not yours.** When a call
arises: make the most defensible choice, apply it, and **flag it inline** for veto, then
record the resolution here so it never gets asked twice. His pattern is to accept unless
he corrects.

## 8. Communication style (strict)

- No filler, no hedging, no narration of work in progress, no self-correction theatre,
  no editorializing, no unnecessary questions.
- Execute silently; deliver results. Don't show the parsing.
- End each update with a concise **box-score-style summary**.
- Flag judgment calls inline.
- Correct errors immediately and silently when flagged; **do not apologize**.
- Don't conflate with the real NHL. Verify specifics; don't estimate when verification
  is possible. Know the current date.

---

## 9. The site & editorial guidelines

Mirrors the baseball tracker: `site/build_site.py` regenerates `site/index.html` from the
workbook, `site/deploy.sh` wraps it and pushes to a GitHub Pages repo. **Rebuild and
deploy after every game**, not just the workbook.

### Editorial voice — same rules as baseball, hockey subjects
- **`Inside the Numbers` is team-level season context** — records, paces, rate stats.
  **Take the data seriously.** Keep real-world comparisons where they're earned (a pace
  against an NHL record, a save percentage against the league's best). **Never write the
  "it's a video game, the engine skews this" caveat into the copy** — the numbers stand
  plainly on their own. No editorializing, no hype, no invented drama.
- **`Recent Headlines` is player-level current form** — point streaks, a goalie's recent
  run, a scoring drought, hot or cold stretches. Kept strictly distinct from
  `Inside the Numbers`; never state the same fact in both. ~6 rows per game, driven by
  the `Headlines` sheet, active/positive items first and past-tense items last (muted).
  **The figure carries the number; the sentence must add new information rather than
  restating it.** Derive streaks from the recaps, never by assumption — a game missed
  does not break a streak.
- **No instruction text anywhere** ("tap a tile for the rundown" and friends). Section
  subtitles carry data only.
- **Output must be pure ASCII.** Some mobile webviews decode as Latin-1 and mojibake
  `·`/`–`. Entity-escape emitted HTML; use `\25B2`-style escapes in CSS and `\u2013`
  in JS; assert the chrome files are ASCII at build time.
- Empty sections render an `.empty` placeholder so a 0-0-0 page looks deliberate.
- The Pages repo keeps its **own CLAUDE.md build spec** listing customizations that must
  survive every regeneration. Read it before changing the generator.

---

## 10. Current state

- **New York Islanders, Season 1.** Repo scaffolded; no workbook yet, nothing logged.
- **Blocked on the user for the sheet list and every column header** (§3, §4). Nothing
  gets built until he gives them — do not invent a schema to get moving.
- After that: build the empty workbook, port `add_game.py` and `build_site.py` from the
  baseball project, set up the Pages repo, then log game 1.
