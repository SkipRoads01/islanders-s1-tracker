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
    g=2, date="10/03/2026", opp="NJD", ha="H",
    result="W",                        # W / L / OTL / SOL
    topup=True,
    # period lines: (NYI, OPP) goals and shots and hits, in order 1,2,3,OT,SO
    goals=[(3, 0), (0, 1), (3, 0)],
    shots=[(12, 19), (4, 14), (19, 9)],
    hits=[(12, 14), (9, 6), (19, 14)],
    # final team stats, exactly as the game's screen shows them
    toa=("07:31", "05:44"), passing=(74.4, 89.2), fow=(14, 32),
    pim=("06:00", "10:00"), pp=("1/5", "0/3"), ppm=("07:05", "05:32"), shg=(1, 0),
    goalies=[],
    opp_goalies=[dict(goalie="Rittich", catches="R", dec="L", sa=35, sv=29, ga=6, toi="60:00")],
    skaters=[],
    # scoring: the site adds the scorer's running season total in parens, so do not type it here
    scoring=[
        dict(per="1", team="NYI", scorer="K. Palmieri", a1="B. Horvat",   a2="M. Schaefer", typ="GWG", score="1-0"),
        dict(per="1", team="NYI", scorer="A. Duclair",  a1="B. Schenn",   a2="A. Pelech",   typ="EV",  score="2-0"),
        dict(per="1", team="NYI", scorer="M. Schaefer", a1="C. Cizikas",  a2="M. Kessel",   typ="SHG", score="3-0"),
        dict(per="2", team="NJD", scorer="J. Bratt",    a1=None,          a2=None,          typ="EV",  score="3-1"),
        dict(per="3", team="NYI", scorer="B. Schenn",   a1="M. Coronato", a2="M. Schaefer", typ="PPG", score="4-1"),
        dict(per="3", team="NYI", scorer="C. Cizikas",  a1=None,          a2=None,          typ="EV",  score="5-1"),
        dict(per="3", team="NYI", scorer="M. Maccelli", a1=None,          a2=None,          typ="EV",  score="6-1"),
    ],
    recap=[
        ("1", "Horvat breaks in on the left post and goes cross slot to Palmieri (1), who one-times it from "
              "inside the hash in the left circle (1-0). Duclair (1) skates in on a delayed penalty and dekes "
              "to a backhand at the right post (2-0). Gritsyuk slashing. Noesen slashing, 5-on-3. Schenn hooking. "
              "Palmieri tripping, 5-on-3 the other way. Pageau hooking, 5-on-3 for 6 seconds. Schaefer (1) goes "
              "down broadway on the break and has hands on Rittich shorthanded (3-0)."),
        ("2", "Bratt rebounds himself at the right post after a flurry of shots to start the period (3-1)."),
        ("3", "Mantha slashing. Bjugstad holding, 5-on-3. Coronato feeds Schenn (1), who finishes from above "
              "the hash in the right circle on the power play (4-1). Hischier interference. Cizikas (1) is "
              "credited after losing the puck at the side of the net and watching Bratt knock it into his own "
              "net (5-1). Maccelli (1) breaks away with two defensemen in tow and has hands down the slot on "
              "Rittich (6-1)."),
    ],
    inside=[
        "NYI is <b>1-0-1</b> through two games, worth <b>3</b> of a possible 4 points.",
        "<b>Six</b> different Islanders scored in G2 vs. NJD; the team had one goal in G1 at TOR.",
        "NYI took the 1st and the 3rd <b>3-0</b> in G2 vs. NJD and lost the 2nd <b>1-0</b>.",
        "Three straight NYI minors in the 1st of G2 vs. NJD put New Jersey two men up twice, and the only goal "
        "of that stretch went the other way.",
        "NYI won G2 vs. NJD by five while being outshot <b>42-35</b> and winning <b>14 of 46 faceoffs (30%)</b>.",
        "The Isles have killed all <b>6</b> power plays they have faced this season and are <b>1 for 7</b> on their "
        "own, the goal coming in the 3rd of G2 vs. NJD.",
        # topup passes ignore the editorial lists; these two were appended to the Inside sheet
        # directly when the screens arrived.
    ],
    headlines=[
        (3, "Points", "<b>Schaefer</b> set up the opening goal, fed Schenn on the power play and scored "
                      "shorthanded himself in G2 vs. NJD, a three-point night from a 19-year-old defenseman.", "active"),
        (2, "Points", "<b>Schenn</b> found Duclair at the right post in the 1st and finished Coronato's pass "
                      "from above the right circle in the 3rd of G2 vs. NJD.", "active"),
        (1, "Goals", "<b>Palmieri</b> one-timed Horvat's cross-slot feed from the left circle to open the "
                     "scoring in G2 vs. NJD, the first of six Islanders goals.", "active"),
        (2, "Games", "<b>Horvat</b> has a point in each of the first two games: the goal in G1 at TOR and the "
                     "primary assist on Palmieri's opener in G2 vs. NJD.", "active"),
        (1, "Goals", "<b>Maccelli</b> outran two defensemen down the slot and beat Rittich for the last of the "
                     "six in G2 vs. NJD.", "active"),
        (1, "Goals", "<b>Cizikas</b> was credited in the 3rd of G2 vs. NJD after he lost the puck at the side "
                     "of the net and Bratt knocked it in himself.", "active"),
    ],
    summary="Six goals from six different scorers in the home opener, three of them in the first period. "
            "New Jersey's only reply came off its own rebound.",
    notes="Team stats and both box scores held for the screens; the play-by-play carries the scoring, the "
          "penalties and nothing else.",
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
