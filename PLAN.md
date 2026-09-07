# Islanders S1 Tracker — site plan (nothing built yet)

Derived from the live Royals S12 site (https://skiproads01.github.io/royals-s12-tracker/)
and the build spec in that Pages repo's `CLAUDE.md`. Every column and tile below is a
**proposal for sign-off**, not a schema. Per `CLAUDE.md` §4, nothing gets built until the
user confirms or corrects each list.

## 1. What the Royals site is

Header: title, `Through G5 · Sep 5, 2026`, record `3–2`, streak pill `W1`, `Next G6 vs MIN`,
season-series line. Then 11 tabs, all visible at once:

| Tab | What it shows |
|---|---|
| Overview | stat strip (13 tiles), `Inside the Numbers`, `Recent Headlines`, `Game Log` (box score + play-by-play per game), `Batting Leaders`, `Team Batting`, `Pitching Leaders`, `Team Pitching` |
| Roster | Player, POS, League, Age, Overall, Potential, 40-Man, PS Elig, MLB Serv, Salary, Contract |
| Schedule | month-grouped list, home/away accent, result, line-up change notes |
| Starters | one card per starting pitcher: line per start, season stats flip |
| Relievers | same for the bullpen |
| vs. Divisions | record vs every club, grouped by division |
| vs. Ranked | record vs ranked opponents |
| Opp Pitching | opposing pitchers faced (G, IP, H, R, ER, BB, K, HR, ERA, AVG) and `Arms Faced` |
| Splits | batting and pitching split tiles with per-player rundowns |
| POG | Player of the Game, one per win, per-player list |
| Errata | shift results, bases loaded, count-based hits, steals, errors, challenges |

Site rules that carry over unchanged: pure-ASCII output, no instruction text, the
`Inside the Numbers` / `Recent Headlines` split and voice rules (§8 of the Royals spec),
sortable tables with totals rows pinned, `version.txt` + `<meta name="build">` bump on
every publish, light/dark palette tokens.

## 2. Baseball → hockey mapping

| Royals section | Islanders equivalent | Source sheet |
|---|---|---|
| Overview strip | see §3 | `Games` |
| Game Log (R/H/E box + PBP) | period-by-period box (`1 · 2 · 3 · OT · SO · F`, plus `SOG`), PBP by period | `Games`, `Recaps` |
| Batting Leaders / Team Batting | Skating Leaders / Team Skating | `Skater Game Log` |
| Pitching Leaders / Team Pitching | Goaltending Leaders / Team Goaltending | `Goalie Game Log` |
| Starters / Relievers | Goalies (one card per goalie, line per start, season flip). Hockey has no bullpen; relief appearances are rows on the same card | `Goalie Game Log` |
| Roster | Roster (columns are the user's call; NHL 26 / Age / Overall / Potential / Contract are the obvious analogues) | `Roster Ref` |
| Schedule | Schedule, month-grouped, 82 games, line-up-change notes become lines/pairs/goalie-start notes | `Schedule` |
| vs. Divisions | vs. Divisions — Metropolitan / Atlantic / Central / Pacific as the game presents them | `Teams`, `Games` |
| vs. Ranked | vs. Ranked, same mechanic | `Teams`, `Games` |
| Opp Pitching / Arms Faced | Opp Goaltending / Goalies Faced | `Opp Goaltending` |
| Batting / Pitching Splits | Skating / Goaltending Splits (candidates: PP, PK, EN, SO, 1st-goal, home/road) | game logs |
| POG | Three Stars (one set per game, all games, not just wins) or POG one-per-win — **user's call** | `Games` |
| Errata | hockey oddities the notes actually track: penalties by type, faceoff results, hits/blocks, shootout attempts, challenges. Populate only from what shows up in the notes | TBD |
| Challenges | coach's challenges (offside / goalie interference / missed stoppage), same accept–deny table | TBD |

Team tag on the page: `NYI` where the Royals page says `KC`; `the Islanders` / `the Isles`
in prose.

## 3. Overview strip — proposed tiles

Mirrors the Royals order and its rules (Pace conditional from G10, Comebacks with `<small>`
share, BLL kept at 0).

| Tile | Meaning |
|---|---|
| Goals For | total goals scored |
| Goals Ag. | total goals allowed |
| Goal Diff | `+N`, green when positive |
| 1-Goal | record in one-goal games (regulation only, or all — **ruling needed**) |
| OT/SO | record in games past regulation |
| Shutouts | shutouts by NYI goalies |
| Comebacks | wins after trailing, with share of total wins |
| PP | power play `G/Opp (pct)` |
| PK | penalty kill `killed/faced (pct)` |
| Most Goals | most goals in a game |
| Most Goals Allowed | most allowed in a game |
| Pace | projected 82-game `W-L-OTL` from G10 on |
| BLL | Blown Lead Losses, same definition (led at any point, lost) |

## 4. Workbook — proposed sheet list

Same shape as the baseball workbook. Sheets marked * are formula-driven and never
hand-edited.

`Overview`*, `Summary`*, `Games`, `Skater Game Log`, `Goalie Game Log`,
`Team Skating`*, `Team Goaltending`*, `Skating vs Opp`*, `Goaltending vs Opp`*,
`Opp Skating` (hybrid), `Opp Goaltending` (plain log), `Schedule`, `Teams`,
`Roster Ref`, `Recaps`, `Headlines`, `Box`, `Notes`.

### Candidate columns (awaiting the user — do not build from these)

- `Games`: `G`, `Date`, `Opp`, `H/A`, `Result` (W / L / OTL / SOL), `GF`, `GA`, `1st`,
  `2nd`, `3rd`, `OT`, `SO`, `SOG For`, `SOG Ag`, `PP` (`g/opp`), `PK` (`k/faced`),
  `PIM`, `Streak` (formula), `Record` (formula), `Stars` or `POG`, `Notes`.
- `Skater Game Log`: `G`, `Opp`, `Player`, `Pos`, `Goals`, `A`, `Pts`, `+/-`, `SOG`,
  `PIM`, `Hits`, `Blk`, `PPG`, `PPA`, `SHG`, `GWG`, `ENG`, `SO Goal`, `FOW`, `FOL`,
  `TOI`.
- `Goalie Game Log`: `G`, `Opp`, `Goalie`, `Catches`, `Dec` (W / L / OTL / SOL / ND),
  `SA`, `SV`, `GA`, `SV%` (formula), `TOI`, `SO` (shutout flag), `EN` (goals against
  while pulled — logged separately, not charged), `Start/Relief`.
- `Opp Goaltending`: `G`, `Opp`, `Goalie`, `Catches` (L/R, right default), `Dec`, `SA`,
  `SV`, `GA`, `TOI`, `SO Att/Saves`.
- `Opp Skating`: per-opponent counting totals (`GP`, `G`, `A`, `SOG`, `PIM`, `Hits`,
  `PP g/opp`), rate columns as formulas, `ALL OPP` row.

## 5. Rulings needed before anything is built (CLAUDE.md §7)

1. **Streak after an OTL / SOL** — does it break a win streak, extend a loss streak, or
   sit outside both (NHL standings treat OTL as a point, not a loss)?
2. **Record format** — `W-L-OTL` (NHL) or `W-L` only.
3. **Shootout goals** — count as player goals in the log or only as the deciding result?
   NHL counts them as neither goal nor assist; the winning goal is credited to the team.
4. **Empty-net goals against** — charged to the goalie or logged as team GA only.
5. **TOI format** — `mm:ss` text, or seconds as a number with a display formula.
6. **Three Stars vs. one Player of the Game** — and whether it is awarded in losses.
7. **1-Goal record** — regulation-only, or all games decided by one goal including OT/SO.
8. **Season length and league alignment** — 82 games and the division layout as the game
   presents them; confirm before `vs. Divisions` is laid out.
9. **Repo layout** — the baseball project keeps data in `PS5 Baseball` and the site in a
   separate `royals-s12-tracker` Pages repo. This repo is named like the Pages repo. One
   repo for both, or a separate Pages repo?

## 6. Order of work once §4 and §5 are answered

1. Build the empty workbook `S1 NY Islanders.xlsx` with the confirmed sheets and headers.
2. Port `add_game.py` from the baseball project and rewrite the GAME block for hockey.
3. Port `build_site.py` / `deploy.sh`, swap the chrome (logo, `--royal` → Islanders
   blue/orange tokens, `NYI` tag, 82-game paces), keep every rule from the Royals spec
   that is not baseball-specific.
4. Set up the Pages repo with its own `CLAUDE.md` build spec.
5. Log game 1.
