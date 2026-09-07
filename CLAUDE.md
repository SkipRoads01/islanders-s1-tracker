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
islanders-s1-tracker/          # data repo AND the GitHub Pages repo (one repo, site at the root)
├── S1 NY Islanders.xlsx       # THE workbook — exact filename, set once and never renamed
├── index.html                 # generated — never hand-edit
├── version.txt                # generated build id; must match <meta name="build">
├── sw.js                      # offline shell: network-first page cache, version.txt never intercepted
├── manifest.webmanifest       # add-to-home-screen; icon-192/512.png + apple-touch-icon.png from the crest
├── CLAUDE.md                  # this file
├── add_game.py                # injection engine — edit the GAME block, run it (not yet ported)
├── scripts/seed_workbook.py   # one-time seed from the 09/29/2026 franchise screens; record only
└── site/
    ├── build_site.py          # workbook -> index.html + version.txt
    ├── deploy.sh              # build, ASCII + build-id checks, commit, push
    ├── chrome.css             # the Royals CSS ported to Islanders tokens (--isles, --orange)
    ├── extra.css              # hockey-only additions (line cards, team abbr tiles)
    ├── site.js                # sortable tables, tabs, version check
    └── logos/                 # crest.svg (masthead), nhl.svg + east.svg (hero marks),
        │                      # wordmark.png = ghosted greyscale backdrop once it exists
        └── teams/XXX.svg      # all 32 club logos, embedded as .lg-XXX classes; schedule rows
                               # carry the opponent's at 72px, hero and divisions smaller
```

- Commit the workbook to git after each game. Git history is the versioning.
- `site/deploy.sh` is the publish path. It rebuilds, refuses non-ASCII output or a build-id
  mismatch, commits, and pushes the current branch.
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
| `Games` | `G` `Date` `Opp` `H/A` `Result` `GF` `GA` `P1 F` `P1 A` `P2 F` `P2 A` `P3 F` `P3 A` `OT F` `OT A` `SO F` `SO A` `Shots F` `Shots A` `Hits F` `Hits A` `TOA F` `TOA A` `Pass% F` `Pass% A` `FOW F` `FOW A` `PIM F` `PIM A` `PP F` `PP A` `PPM F` `PPM A` `SHG F` `SHG A` `Streak` `Record` `Summary` |
| `Scoring` | `G` `Opp` `Period` `Team` `Scorer` `A1` `A2` `Type` `Score` |
| `Goalie Game Log` | `G` `Opp` `Goalie` `Dec` `SA` `SV` `GA` `SV%` `TOI` `SO` `Start/Relief` |
| `Opp Goaltending` | `G` `Opp` `Goalie` `Catches` `Dec` `SA` `SV` `GA` `SV%` `TOI` |
| `Skater Game Log` | `G` `Opp` `Player` `Pos` `G` `A` `Pts` `+/-` `SOG` `PIM` `Hits` `Blk` `PPG` `PPA` `SHG` `GWG` `ENG` `FOW` `FOL` `TOI` |
| `Recaps` | `G` `Opp` `Period` `Text` |
| `Inside` | `G` `Bullet` |
| `Headlines` | `G` `Fig` `Kicker` `Text` `Tone` |
| `Opp Skating` | *awaiting the user* |

The `Games` team-stat columns are **the game's own end-of-period screen**, in its order and
its wording: Total Shots, Hits, Time on Attack, Passing, Faceoffs Won, Penalty Minutes,
Powerplays, Powerplay Minutes, Shorthanded Goals. `F` is NYI, `A` is the opponent. Period
goals/shots/hits come from the four cumulative screens differenced into per-period lines.
| `Roster Ref` | `Player` `Pos` `Group` `Status` `#` `OVR` `POT` `POT Cert` `Age` `Ht` `Wt` `Shoots` `Type` `Ext` `Clause` `FSC` `26-27` `27-28` `28-29` `29-30` `30-31` `31-32` `32-33` `33-34` `Then` |
| `Lines` | `Unit` `Slot` `Player` `Pos` `OVR` `Chem` |
| `Owner Goals` | `Tier` `Goal` `Reward` `Eval Date` `Status` |
| `Budget` | `Line` `Allocated` `Spent` `Remaining` |
| `Cap` | `Season` `Salary Cap` `Main Roster` `System` `Contracts` |
| `Front Office` | `Key` `Value` |
| `Transactions` | `Date` `Type` `Partner` `Direction` `Out` `In` `Result` |
| `Front Office` tab order | strip, **General Manager** (Transactions, Extension Eligible), Owner, Owner Goals, Operations Budget, Cap Outlook, League Cap Rules |
| `Schedule` | `G` `Date` `H/A` `Opp` `Time (ET)` |
| `Teams` | `Team` `Abbr` `Conference` `Division` |
| `Notes` | `Date` `Note` |

`Transactions` covers everything the front office is offered or does, not just trades: `Type`
is `Trade offer`, `Signing`, `Waivers` or the like, and `Result` is `Accepted` or `Declined`.
On a `Waivers` row `Direction` is `Claim` and `Partner` is `-`, because a waiver claim has no
counterparty. A field the screen did not capture is a `-`, never a guess.

`Roster Ref` conventions, taken from the game's List All Contracts screen: `Group` is
`Main Roster` or `In the System`; `Status` is `Dressed` / `Scratched` for the main roster;
year columns hold the salary in $M as a number, or the tag the game shows in that year
(`RFA`, `UFA`, `UNSIGNED`); `Then` repeats the tag that follows the last paid year.
`Lines` units: `F1`-`F4`, `D1`-`D3`, `PP1`-`PP2`, `PK1`-`PK3`, `G` (slot 1/2 as the lineup
screen orders them), `SCR`. `Chem` is the unit's line-chemistry OVR impact, repeated on
each row of the unit.

---

## 5. Notation reference (user's shorthand)

### How a game gets logged (the cheap path)

G1 proved the split: **the screens carry the numbers, the notes carry the story.** Every
figure in the `Games` row came from photographs, not from anything the user had to write.

| Tier | What | Fills |
|---|---|---|
| 1 | The team-stats screen at each intermission **and** at the final - 4 photos | every `Games` team column. Per-period lines come from differencing the cumulative screens, so no per-period screen is needed |
| 2 | The postgame player-stats screen - 1 photo | `Skater Game Log`, and every skater's season card |
| 3 | Whatever the user feels like writing | `Scoring`, `Recaps`, `Headlines`, the game summary |

Tier 3 can be as thin as the scoring line and still produce a full page, because tiers 1 and
2 are already complete. **A missing tier 3 never blocks a publish.** What must never happen is
inventing tier 2 out of tier 3 - if the box score is not photographed, the skater log stays
empty and the site says so.

**The user will identify new shorthand as it happens** - when he does, add it here immediately
so it becomes permanent.

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

- **Record format is `W-L-OTL`** because the game itself shows `0-0-0` on its matchup
  screen. Applied to the hero, footer and every `vs.` record.
- **84-game season.** The 2026-27 CBA expands the regular season from 82 to 84; both added
  games are intra-division, so every Metropolitan rival is played 4 times and the preseason
  is capped at 4 games. Used for pace math and the schedule count.
- **Own division first** in `vs. Divisions`: Metropolitan, Atlantic, Central, Pacific.
- **Schedule** is the NHL.com official 2026-27 release: **84 games, 42 home, 42 away**,
  opening **09/30/2026 at TOR**. An earlier load from a third-party calendar PDF was short
  one game (it omitted the September opener) and was replaced. Preseason is not played.
  **Verify any schedule source against the club's own release before loading it.**
- **A table must fit its column; horizontal scrolling is a defect.** The chrome's default
  `tbody td { white-space: nowrap }` is what pushes the last column off-screen, so any table
  carrying prose or long strings needs `table-layout: fixed`, per-column widths, and
  `white-space: normal; overflow-wrap: anywhere`. Verify at 390px as well as desktop, and
  remember a `nowrap` chip inside a cell sets the floor for that column. Drop the least
  important columns on a phone rather than letting the table scroll.
- **Every table is sortable; totals rows are locked.** A totals row goes in `<tfoot>` with
  class `tot` and never takes part in a sort. `build_site.py`'s `table()` does this via its
  `tfoot=` argument; never emit a totals row into `<tbody>`.
- **Overview strip tiles** (proposed, veto any): Goals For, Goals Ag., Goal Diff, 1-Goal,
  OT/SO, Shutouts, Comebacks (+share of wins), PP, PK, Most Goals, Most Goals Allowed, BLL.
  `Pace` joins from game 10, per the Royals spec.
- **An OTL is not a loss.** The user's own wording. It never starts or extends a loss streak;
  it gets its own streak token, `OT1`, `OT2`. A shootout loss takes `SO1`. The standings award
  a point, so the record stays `W-L-OTL`.
- **BLL counts an OTL.** A lead held and not converted to a win is a blown lead whichever
  column the game lands in. G1 counts: NYI led 1-0 and lost in overtime.
- **`1-Goal` includes games decided in OT or the shootout.** A 2-1 overtime result is a
  one-goal game.
- **`PP` reads goals/opportunities, `PK` reads killed/faced**, each with its percentage in the
  tile's `<small>` slot.
- **Player names are clickable wherever they appear** - roster, lines, scoring, goalie cards,
  and inside the editorial copy. The card carries the photo, bio, contract, season line, and
  any `Headlines` row that names the player. Photos live at `site/logos/players/<key>.png`
  where `<key>` is the name lowercased with punctuation stripped (`bhorvat`, `isorokin`); the
  crest stands in until a photo exists.
- Still open, ask when they first arise: shootout goals in the skater log, empty-net goals
  against, TOI format, Three Stars vs. one Player of the Game.

Everything else is empty until the user makes the call. Hockey has several that baseball doesn't — how an
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
- **A goal always carries the scorer's season total in parentheses**, hockey box-score style:
  `Horvat (1)`. It is the running count through that game, not the season-to-date figure, so
  G1's goal reads `(1)` forever. Applied in the `Scoring` table and in recap prose. Opponent
  goals get no parenthetical: all we can count is what they have scored against NYI, which is
  not their season total.
- **A contrasting clause takes `but`, never `and`.** When the second half of a sentence cuts
  against the first, the conjunction has to carry that turn: "Romanov, Pelech and Barzal each
  took one, **but** Toronto came away with nothing on the power play." `And` flattens the two
  halves into a list and throws away the point of the sentence. Read every compound sentence
  and ask whether the clauses agree or contrast before choosing the word.
- **Passing percentage is never a story.** The user plays every shift; the CPU side does not
  misfire passes the way a human stick does. A gap of 15-20 points in the computer's favour is
  the normal state of this game, not a finding about either team. Never write it as a
  comparison, never build a bullet around it, and never call it a weakness. This is the same
  principle as the baseball manual's rule that scarce walks are not a story. The figure still
  gets **logged** in `Games` and **shown** in the game card's Team Stats table, where it is a
  raw number rather than a claim. The same caution applies to any stat the engine drives
  rather than the player: if the gap is structural, it is context, not news.
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

- **New York Islanders, Season 1, preseason.** In-game date 09/29/2026. Record 0-0-0.
- Workbook seeded with the front-office and roster sheets in §4; **no game-log sheets yet**
  because their columns are still the user's to give (§4). Do not invent them.
- **G1 logged**: 09/30/2026 at TOR, 1-2 OTL. Record 0-0-1.
- Site built from this repo: tabs Overview, Roster, Lines, Front Office, Schedule, Goalies,
  vs. Divisions. Game-driven sections render `.empty` placeholders. Works offline once loaded
  (service worker, verified with the network cut). GitHub Pages serves `main` at the repo
  root: https://skiproads01.github.io/islanders-s1-tracker/ once the branch is merged and
  Pages is switched on.
- Schedule loaded (84 games, see §7). Hero shows the next game, the NHL shield
  and the Eastern Conference mark.
- **Extension Eligible** lists every player whose deal ends after the current season, sorted
  interested-first then by OVR, with `Roster Ref`'s `Ext` column as the `Interest` value
  (`Yes` / `No` / `-` when the screen did not say). Unsigned prospects are excluded: they need
  signing, not extending. It replaced the old Wants Extension table on the Roster tab.
- Pending from the user: goalie contracts (Sorokin, Varlamov); the wordmark image file
  (`site/logos/wordmark.png`; the crest stands in until then); player photos
  (`site/logos/players/`); and **the skater box score for G1** - the play-by-play supports
  only Horvat's goal and the three NYI minors, so `Skater Game Log` is deliberately empty and
  Skating Leaders / Team Skating render placeholders rather than partial totals.
- `add_game.py` is the injection engine: edit its `GAME` block, run it, then `site/deploy.sh`.
