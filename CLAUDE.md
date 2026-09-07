# CLAUDE.md — PS5 Hockey Franchise Tracker

Operating manual for maintaining the hockey franchise stat workbook in Claude Code.
Same architecture as the baseball tracker at `~/Documents/GitHub/PS5 Baseball` — read
that project's `CLAUDE.md` if anything here is ambiguous about *mechanics*; this file
owns everything sport-specific and **wins on any conflict.**

Read this file fully before touching the workbook.

---

## 1. What this project is

A game-by-game statistical record of a **PS5 hockey franchise**, played manually —
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
├── <TEAM> S<N>.xlsx      # THE workbook — exact filename, set once and never renamed
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

Sheets, in order: **Overview, Summary, Games, Skater Game Log, Goalie Game Log,
Team Skating, Team Goaltending, Skating vs Opp, Goaltending vs Opp, Opp Skating,
Opp Goaltending, Notes, Stars, Roster Ref, Schedule, Teams, Recaps, Headlines, Box.**

### You type into these sheets
| Sheet | One row per… |
|---|---|
| `Skater Game Log` | skater who dressed |
| `Goalie Game Log` | goalie who played |
| `Opp Goaltending` | opposing goalie who faced the team |
| `Games` | the game itself (line score, special teams, flags) |
| `Stars`, `Recaps`, `Box`, `Headlines` | three stars, play-by-play, opponent SOG/PIM, headlines |

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

**PROVISIONAL** until the first batch of real notes lands — confirm the live header row
before writing, and adjust these to whatever the user actually tracks.

**Games**:
```
G#, Opp, H, A, GF, GA, Res, End, Pts, SOG, SOGA, PIM, PPG, PPO, PKGA, PKO,
SHG, ENG, SO, CFBW, BLL, Streak
```
- `H`/`A`: exactly one is 1. Infer from which team is listed as home in the notes.
- `Res` ∈ { W, L, OTL, SOL }. `End` ∈ { REG, OT, SO }.
- `Pts` = 2 for W; 1 for OTL/SOL; 0 for L. **This is the key departure from baseball —
  the record is three-column (W-L-OTL) and the standings run on points.**
- `PPO`/`PPG` = power plays drawn / converted. `PKO`/`PKGA` = times shorthanded / goals
  allowed on them. `PP%` and `PK%` are derived on the Team sheets, never typed.
- `SO` = shutout thrown (0/1). `ENG` = empty-net goals scored.
- `Streak` is a **formula — never typed** (see §6).

**Skater Game Log**:
```
G#, Opp, Player, Pos, TOI, G, A, PM, PIM, SOG, Hits, BLK, FOW, FOL, PPG, PPA, SHG, GWG
```
- Points (`P` = G+A) and `FO%` are derived on `Team Skating` — don't store them.
- `PM` is plus/minus and **can be negative**; it's the one counting column that can be.
- `TOI` in `MM:SS`, stored as text or an Excel time — pick one on the first game and
  never mix.
- Sanity: `PPG + SHG <= G`; `FOW + FOL` only for centers.

**Goalie Game Log**:
```
G#, Opp, Player, Dec, TOI, SA, SV, GA, SO, EN
```
- `Dec` ∈ { "", W, L, OTL, SOL }. At most one decision per game.
- Sanity: `SA = SV + GA`; `SV%` is derived; sum of goalies' `GA` == `GA` on the Games
  row **minus** empty-net goals against, so track `EN` explicitly.

**Opp Goaltending**:
```
G#, Opp, Goalie, Catches, Dec, TOI, SA, SV, GA
```
Derived from the team's shots and goals against him; requires the notes to mark
**opponent goalie changes** (starter by name, then any change and when).

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
   above. Hockey's version must treat OTL/SOL as streak-breaking-but-not-a-loss-streak
   unless the user says otherwise — **ask once, then honor it forever.**
7. **Expansion** if needed (§3).
8. **Update `Opp Skating`** — increment the opponent's counting cells, re-sum `ALL OPP`.
9. **Recalc + bake** (§2). Zero error cells before shipping.
10. **Rebuild and deploy the site**, not just the workbook.
11. **Summarize** with a tight box score.

---

## 7. Standing judgment calls

- **Empty-net goals** count in `GF`/`GA` but are unearned against the goalie — keep `EN`
  separate so `SV%` and `GAA` stay honest.
- **Shootout goals** do not count in a skater's `G` total; the shootout winner gets the
  team a `W` and the game's lone extra goal in `GF`. Log SO attempts in the notes only.
- **TOI** format is fixed on game 1 and never changes.
- When a judgment call arises, **make the call, log it, and flag it inline** for veto.
  The user's pattern is to accept unless he corrects.

---

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

## 9. The site

Mirrors the baseball tracker: `site/build_site.py` regenerates `site/index.html` from the
workbook, `site/deploy.sh` wraps it and pushes to a GitHub Pages repo. **Rebuild and
deploy after every game.**

Non-negotiables carried over from that project:
- **Output must be pure ASCII.** Some mobile webviews decode as Latin-1 and mojibake
  `·`/`–`. Entity-escape emitted HTML; use `\25B2`-style escapes in CSS and `–` in
  JS; assert the chrome files are ASCII at build time.
- **No instruction text anywhere** ("tap a tile for the rundown" and friends). Section
  subtitles carry data only.
- Empty sections render an `.empty` placeholder so a 0-0-0 page looks deliberate.
- `Headlines` sheet drives a Recent Headlines section — player-level *current form*,
  kept distinct from the team-level `Inside the Numbers`.
- The Pages repo keeps its **own CLAUDE.md build spec** listing customizations that must
  survive every regeneration. Read it before changing the generator.

---

## 10. Current state

- **Repo scaffolded; no workbook yet.** Nothing has been played or logged.
- **Team and season not yet set** — they determine the workbook filename, the `Teams`
  sheet, the schedule length, and the site title.
- Next steps: name the team/season, build the empty workbook, port `add_game.py` and
  `build_site.py` from the baseball project, then log game 1.
