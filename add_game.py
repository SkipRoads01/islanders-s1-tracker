#!/usr/bin/env python3
"""Injection engine. Edit the GAME block, run it, then site/deploy.sh.

Creates the game-log sheets on first use. Column sets come from the game's own
end-of-period team-stats screen, not from invention: Total Shots, Hits, Time on
Attack, Passing, Faceoffs Won, Penalty Minutes, Powerplays, Powerplay Minutes,
Shorthanded Goals.

Two passes are supported, because the notes usually arrive before the photographs:
  topup=False  log the game from the play-by-play. Any screen-only figure left None
               stays empty in the workbook and prints as a dash on the site.
  topup=True   the screens arrived: write the team-stat columns into the Games row
               that is already there and append the goalie / skater rows. A None
               never overwrites a value already on the sheet.
"""
import sys
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
WB = ROOT / "S1 NY Islanders.xlsx"

# ===================================================================== GAME
# Tier 1 (the four team-stats screens) and tier 2 (the postgame player screen) can be
# None when the screens have not been photographed yet; the row is logged from the
# play-by-play alone and the site prints a dash rather than a number. When the screens
# arrive, fill the same block in, set topup=True, and run again: the Games row is
# topped up in place and the goalie / skater rows are appended.
GAME = dict(
    g=1, date="09/30/2026", opp="TOR", ha="A",
    result=None,                       # W / L / OTL / SOL
    topup=False,
    # period lines: (NYI, OPP) goals and shots and hits, in order 1,2,3,OT,SO
    goals=[],
    shots=None,
    hits=None,
    # final team stats, exactly as the game's screen shows them
    toa=None, passing=None, fow=None,
    pim=None, pp=None, ppm=None, shg=None,
    goalies=[],
    opp_goalies=[],
    skaters=[],
    # scoring: the site adds the scorer's running season total in parens, so do not type it here.
    # Score is NYI-first here; recap prose is leader-first, per the user's own shorthand.
    scoring=[],
    recap=[],
    inside=[],
    headlines=[],
    summary="",
    notes="",
)
# ===========================================================================

HEAD_FILL = PatternFill("solid", fgColor="00539B")
HEAD_FONT = Font(bold=True, color="FFFFFF")
BORDER = Border(bottom=Side(style="thin", color="D7DFEB"))

SHEETS = {
 "Games": ["G", "Date", "Opp", "H/A", "Result", "GF", "GA",
           "P1 F", "P1 A", "P2 F", "P2 A", "P3 F", "P3 A", "OT F", "OT A", "SO F", "SO A",
           "Shots F", "Shots A", "Hits F", "Hits A", "TOA F", "TOA A", "Pass% F", "Pass% A",
           "FOW F", "FOW A", "PIM F", "PIM A", "PP F", "PP A", "PPM F", "PPM A", "SHG F", "SHG A",
           "Streak", "Record", "Summary"],
 "Scoring": ["G", "Opp", "Period", "Team", "Scorer", "A1", "A2", "Type", "Score"],
 "Goalie Game Log": ["G", "Opp", "Goalie", "Dec", "SA", "SV", "GA", "SV%", "TOI", "SO", "Start/Relief"],
 "Opp Goaltending": ["G", "Opp", "Goalie", "Catches", "Dec", "SA", "SV", "GA", "SV%", "TOI"],
 "Skater Game Log": ["G", "Opp", "Player", "Pos", "G", "A", "Pts", "+/-", "SOG", "PIM", "Hits",
                     "Blk", "PPG", "PPA", "SHG", "GWG", "ENG", "FOW", "FOL", "TOI"],
 "Recaps": ["G", "Opp", "Period", "Text"],
 "Inside": ["G", "Bullet"],
 "Headlines": ["G", "Fig", "Kicker", "Text", "Tone"],
}

def ensure(wb, name):
    if name in wb.sheetnames:
        return wb[name]
    ws = wb.create_sheet(name)
    hdr = SHEETS[name]
    ws.append(hdr)
    for c in range(1, len(hdr) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = HEAD_FILL; cell.font = HEAD_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=2, column=c).border = BORDER
        ws.column_dimensions[get_column_letter(c)].width = max(len(str(hdr[c - 1])) + 2, 7)
    ws.freeze_panes = "A3"
    return ws

def last_row(ws):
    """max_row lies; find the last row with a value in column A."""
    r = ws.max_row
    while r >= 3 and ws.cell(row=r, column=1).value in (None, ""):
        r -= 1
    return r

def append(ws, values):
    ws.cell(row=last_row(ws) + 1, column=1)  # anchor
    r = last_row(ws) + 1
    for c, v in enumerate(values, start=1):
        ws.cell(row=r, column=c, value=v)
    return r

def streak_and_record(prev_record, result):
    """W-L-OTL. An OT or shootout loss is its own streak, never a loss streak,
    because the standings award a point for it."""
    w, l, o = prev_record
    if result == "W": w += 1
    elif result == "L": l += 1
    else: o += 1
    return "%d-%d-%d" % (w, l, o), (w, l, o)

def pair(v):
    """A team-stat pair the screens have not supplied yet stays empty."""
    return (None, None) if v is None else v

def totals(periods):
    """Per-period (NYI, OPP) lines summed; None when the screens are pending."""
    if periods is None:
        return (None, None)
    return (sum(a for a, b in periods), sum(b for a, b in periods))

def team_stat_cells(g):
    """Columns R..AI of the Games row, in sheet order."""
    sf, sa = totals(g["shots"]); hf, ha = totals(g["hits"])
    out = [sf, sa, hf, ha]
    for key in ("toa", "passing", "fow", "pim", "pp", "ppm", "shg"):
        out += list(pair(g[key]))
    return out

def skater_row(g, s):
    gl, a = s.get("g", 0), s.get("a", 0)
    return [g["g"], g["opp"], s["player"], s.get("pos"), gl, a, gl + a, s.get("pm"), s.get("sog"),
            s.get("pim"), s.get("hits"), s.get("blk"), s.get("ppg"), s.get("ppa"), s.get("shg"),
            s.get("gwg"), s.get("eng"), s.get("fow"), s.get("fol"), s.get("toi")]

def goalie_row(g, gl):
    return [g["g"], g["opp"], gl["goalie"], gl["dec"], gl["sa"], gl["sv"], gl["ga"],
            round(gl["sv"] / gl["sa"], 3) if gl["sa"] else None, gl["toi"],
            1 if gl["ga"] == 0 else 0, gl["start"]]

def opp_goalie_row(g, gl):
    return [g["g"], g["opp"], gl["goalie"], gl["catches"], gl["dec"], gl["sa"], gl["sv"], gl["ga"],
            round(gl["sv"] / gl["sa"], 3) if gl["sa"] else None, gl["toi"]]

def logged_games(ws, gnum):
    return [r for r in range(3, last_row(ws) + 1) if ws.cell(row=r, column=1).value == gnum]

def topup(wb, g):
    """The screens arrived after the game was logged from the play-by-play: write the
    team-stat columns into the existing Games row and append whatever box-score rows
    are not there yet. Nothing already on the sheet is overwritten with a blank."""
    games = wb["Games"]
    rows = logged_games(games, g["g"])
    if not rows:
        sys.exit("G%d is not logged yet; run with topup=False first" % g["g"])
    r = rows[0]
    wrote = 0
    for i, v in enumerate(team_stat_cells(g), start=18):   # column R
        if v is not None:
            games.cell(row=r, column=i, value=v); wrote += 1
    added = []
    for name, rowsrc, items in (("Goalie Game Log", goalie_row, g.get("goalies") or []),
                                ("Opp Goaltending", opp_goalie_row, g.get("opp_goalies") or []),
                                ("Skater Game Log", skater_row, g.get("skaters") or [])):
        ws = wb[name]
        if logged_games(ws, g["g"]):
            continue
        for it in items:
            append(ws, rowsrc(g, it))
        if items:
            added.append("%d %s" % (len(items), name))
    wb.save(WB)
    print("topped up G%d: %d team-stat cells%s" % (g["g"], wrote, (", " + ", ".join(added)) if added else ""))

def main():
    wb = load_workbook(WB)
    g = GAME
    for name in SHEETS:
        ensure(wb, name)

    if g.get("topup"):
        return topup(wb, g)

    games = wb["Games"]
    prev = (0, 0, 0)
    lr = last_row(games)
    if lr >= 3:
        rec = games.cell(row=lr, column=37).value  # Record
        if rec:
            prev = tuple(int(x) for x in str(rec).split("-"))
    if logged_games(games, g["g"]):
        sys.exit("G%d already logged" % g["g"])

    per = g["goals"] + [(0, 0)] * (5 - len(g["goals"]))
    gf = sum(a for a, b in g["goals"]); ga = sum(b for a, b in g["goals"])
    streak_map = {"W": "W", "L": "L", "OTL": "OT", "SOL": "SO"}
    record, newrec = streak_and_record(prev, g["result"])
    streak = streak_map[g["result"]] + "1"

    append(games, [g["g"], g["date"], g["opp"], g["ha"], g["result"], gf, ga]
           + [v for p in per[:5] for v in p]
           + team_stat_cells(g)
           + [streak, record, g["summary"]])

    for s in g["scoring"]:
        append(wb["Scoring"], [g["g"], g["opp"], s["per"], s["team"], s["scorer"], s["a1"], s["a2"], s["typ"], s["score"]])
    for gl in g.get("goalies") or []:
        append(wb["Goalie Game Log"], goalie_row(g, gl))
    for gl in g.get("opp_goalies") or []:
        append(wb["Opp Goaltending"], opp_goalie_row(g, gl))
    for sk in g.get("skaters") or []:
        append(wb["Skater Game Log"], skater_row(g, sk))
    for p, text in g["recap"]:
        append(wb["Recaps"], [g["g"], g["opp"], p, text])
    for b in g.get("inside") or []:
        append(wb["Inside"], [g["g"], b])
    for fig, kicker, text, tone in g.get("headlines") or []:
        append(wb["Headlines"], [g["g"], fig, kicker, text, tone])
    if g.get("notes"):
        append(wb["Notes"], [g["date"], "G%d %s %s: %s" % (g["g"], "at" if g["ha"] == "A" else "vs", g["opp"], g["notes"])])

    wb.save(WB)
    print("logged G%d %s %s  %d-%d %s  record %s" % (
        g["g"], "at" if g["ha"] == "A" else "vs", g["opp"], gf, ga, g["result"], record))

if __name__ == "__main__":
    main()
