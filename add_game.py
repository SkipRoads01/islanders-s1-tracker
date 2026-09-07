#!/usr/bin/env python3
"""Injection engine. Edit the GAME block, run it, then site/deploy.sh.

Creates the game-log sheets on first use. Column sets come from the game's own
end-of-period team-stats screen, not from invention: Total Shots, Hits, Time on
Attack, Passing, Faceoffs Won, Penalty Minutes, Powerplays, Powerplay Minutes,
Shorthanded Goals.
"""
import sys
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
WB = ROOT / "S1 NY Islanders.xlsx"

# ===================================================================== GAME
GAME = dict(
    g=1, date="09/30/2026", opp="TOR", ha="A",
    result="OTL",                      # W / L / OTL / SOL
    # period lines: (NYI, OPP) goals and shots and hits, in order 1,2,3,OT,SO
    goals=[(1, 0), (0, 0), (0, 1), (0, 1)],
    shots=[(9, 11), (8, 14), (13, 11), (2, 7)],
    hits=[(7, 13), (11, 10), (5, 13), (5, 1)],
    # final team stats, exactly as the game's screen shows them
    toa=("07:06", "05:15"), passing=(71.6, 88.1), fow=(21, 16),
    pim=("06:00", "04:00"), pp=("0/2", "0/3"), ppm=("02:48", "04:48"), shg=(0, 0),
    goalies=[dict(goalie="I. Sorokin", dec="OTL", sa=43, sv=41, ga=2, toi="65:00", start="Start")],
    opp_goalies=[dict(goalie="M. Stolarz", catches="L", dec="W", sa=32, sv=31, ga=1, toi="65:00")],
    # scoring: the site adds the scorer's running season total in parens, so do not type it here
    scoring=[
        dict(per="1",  team="NYI", scorer="B. Horvat",  a1=None, a2=None, typ="EV", score="1-0"),
        dict(per="3",  team="TOR", scorer="B. Duhaime", a1=None, a2=None, typ="EV", score="1-1"),
        dict(per="OT", team="TOR", scorer="J. Roslovic", a1=None, a2=None, typ="GWG", score="1-2"),
    ],
    recap=[
        ("1",  "Horvat picks up a loose puck over the blue line and dekes Stolarz at the right post (1-0). "
               "Ekman-Larsson slashing. Nylander is hauled down on a breakaway by Holmstrom and gets a penalty shot. "
               "Sorokin turns him away at the right post. Romanov slashing with two minutes left. "
               "Pelech slashes Paul with 7 seconds left."),
        ("2",  "Nylander for slashing, 4-on-4 for 1:14. Barzal for slashing."),
        ("3",  "Duhaime gets a late break in the slot and has hands in close (1-1) with 3 minutes left."),
        ("OT", "Roslovic rebounds himself at the right post for the winner (2-1)."),
    ],
    summary="Gut punch OTL on the road to start the year. Isles look strong all game but wither when it counts.",
    notes="Skater box score pending; Horvat's goal and the three NYI minors are the only per-player figures the "
          "play-by-play supports.",
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

def main():
    wb = load_workbook(WB)
    g = GAME
    for name in SHEETS:
        ensure(wb, name)

    games = wb["Games"]
    prev = (0, 0, 0)
    lr = last_row(games)
    if lr >= 3:
        rec = games.cell(row=lr, column=37).value  # Record
        if rec:
            prev = tuple(int(x) for x in str(rec).split("-"))
    if any(games.cell(row=r, column=1).value == g["g"] for r in range(3, lr + 1)):
        sys.exit("G%d already logged" % g["g"])

    per = g["goals"] + [(0, 0)] * (5 - len(g["goals"]))
    sh = g["shots"] + [(0, 0)] * (5 - len(g["shots"]))
    hi = g["hits"] + [(0, 0)] * (5 - len(g["hits"]))
    gf = sum(a for a, b in g["goals"]); ga = sum(b for a, b in g["goals"])
    streak_map = {"W": "W", "L": "L", "OTL": "OT", "SOL": "SO"}
    record, newrec = streak_and_record(prev, g["result"])
    streak = streak_map[g["result"]] + "1"

    append(games, [g["g"], g["date"], g["opp"], g["ha"], g["result"], gf, ga]
           + [v for p in per[:5] for v in p]
           + [sum(a for a, b in sh), sum(b for a, b in sh),
              sum(a for a, b in hi), sum(b for a, b in hi),
              g["toa"][0], g["toa"][1], g["passing"][0], g["passing"][1],
              g["fow"][0], g["fow"][1], g["pim"][0], g["pim"][1],
              g["pp"][0], g["pp"][1], g["ppm"][0], g["ppm"][1], g["shg"][0], g["shg"][1],
              streak, record, g["summary"]])

    for s in g["scoring"]:
        append(wb["Scoring"], [g["g"], g["opp"], s["per"], s["team"], s["scorer"], s["a1"], s["a2"], s["typ"], s["score"]])
    for gl in g["goalies"]:
        append(wb["Goalie Game Log"], [g["g"], g["opp"], gl["goalie"], gl["dec"], gl["sa"], gl["sv"], gl["ga"],
                                       round(gl["sv"] / gl["sa"], 3) if gl["sa"] else None,
                                       gl["toi"], 1 if gl["ga"] == 0 else 0, gl["start"]])
    for gl in g["opp_goalies"]:
        append(wb["Opp Goaltending"], [g["g"], g["opp"], gl["goalie"], gl["catches"], gl["dec"], gl["sa"], gl["sv"],
                                       gl["ga"], round(gl["sv"] / gl["sa"], 3) if gl["sa"] else None, gl["toi"]])
    for p, text in g["recap"]:
        append(wb["Recaps"], [g["g"], g["opp"], p, text])
    if g.get("notes"):
        append(wb["Notes"], [g["date"], "G%d %s %s: %s" % (g["g"], "at" if g["ha"] == "A" else "vs", g["opp"], g["notes"])])

    wb.save(WB)
    print("logged G%d %s %s  %d-%d %s  record %s" % (
        g["g"], "at" if g["ha"] == "A" else "vs", g["opp"], gf, ga, g["result"], record))

if __name__ == "__main__":
    main()
