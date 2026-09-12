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
    g=3, date="10/06/2026", opp="NYR", ha="A",
    result="L",                        # W / L / OTL / SOL
    topup=False,
    # period lines: (NYI, OPP) goals and shots and hits, in order 1,2,3,OT,SO
    # 1st and 2nd intermission screens plus the final, differenced into per-period lines
    goals=[(1, 2), (0, 0), (1, 2)],
    shots=[(12, 8), (10, 12), (25, 11)],
    hits=[(14, 10), (18, 14), (12, 13)],
    # final team stats, exactly as the game's screen shows them
    toa=("11:53", "03:12"), passing=(78.9, 91.7), fow=(13, 27),
    pim=("08:00", "16:00"), pp=("1/8", "0/4"), ppm=("08:09", "03:07"), shg=(0, 3),
    goalies=[dict(goalie="I. Sorokin", dec="L", sa=31, sv=27, ga=4, toi="60:00", start="Start")],
    # the Rangers goalie is not named in the notes and there is no box score screen
    opp_goalies=[],
    skaters=[],
    # scoring: the site adds the scorer's running season total in parens, so do not type it here.
    # Score is NYI-first here; recap prose is leader-first, per the user's own shorthand.
    scoring=[
        dict(per="1", team="NYR", scorer="M. Rempe",      a1=None,        a2=None, typ="SHG", score="0-1"),
        dict(per="1", team="NYR", scorer="J. Veleno",     a1=None,        a2=None, typ="SHG", score="0-2"),
        dict(per="1", team="NYI", scorer="A. Romanov",    a1="M. Maccelli", a2=None, typ="EV",  score="1-2"),
        dict(per="3", team="NYI", scorer="E. Heineman",   a1="M. Barzal", a2=None, typ="PPG", score="2-2"),
        dict(per="3", team="NYR", scorer="P. Dorofeyev",  a1=None,        a2=None, typ="SHG", score="2-3"),
        dict(per="3", team="NYR", scorer="A. Lafreniere", a1=None,        a2=None, typ="EV",  score="2-4"),
    ],
    recap=[
        ("1", "Bjorkstrand charging. Rempe strips Barzal along the wall and beats Sorokin blocker side "
              "shorthanded (1-0). Veleno takes it off the boards after the Isles lose the puck in their own "
              "end and makes it two shorthanded goals in the period (2-0). Romanov (1) one-times Maccelli's "
              "pass from above the left dot out of the high slot (2-1). Dorofeyev elbowing. Sorokin leaves "
              "the ice behind his own net and takes a delay of game, 4-on-4. Cuylle slashing. Cizikas "
              "boarding, 4-on-4 again. Dorofeyev interference with 30 seconds left."),
        ("2", "Heineman slashing, more 4-on-4 to start the period. Palmieri hooking. Neither side scored."),
        ("3", "Rempe charging at 12:06. Miller charging behind him, 5-on-3 for a minute, and Schneider "
              "penalized on top of it. Barzal finds Heineman (1) at the left post on the power play (2-2). "
              "Dorofeyev answers in alone on Sorokin with New York still a man down (3-2). Seven and a half "
              "minutes left. Lafreniere breaks in alone and has hands in close (4-2)."),
    ],
    inside=[
        # the story lead already prints the record, so a bullet restating it is dead copy
        "The Isles have killed all <b>10</b> power plays they have faced this season, but New York scored "
        "<b>three</b> shorthanded goals in G3 at NYR.",
        "NYI has out-shot its opponent once in three games, <b>47-31</b> in G3 at NYR, and lost by two.",
        "Eight power plays and <b>8:09</b> with the extra man in G3 at NYR produced one goal. NYI is "
        "<b>2 for 15</b> on the season, both goals in the 3rd period.",
        "NYI has won <b>48 of 123 faceoffs (39%)</b> this season and <b>13 of 40</b> in G3 at NYR.",
        "<b>Five</b> of the nine goals NYI has scored this season have come in the 1st period.",
    ],
    headlines=[
        (27, "Saves", "<b>Sorokin</b> stopped 27 of 31 in G3 at NYR and gave up four goals in a night after "
                      "allowing three across his first two starts combined.", "active"),
        (1, "Goals", "<b>Romanov</b> one-timed Maccelli's feed out of the high slot in the 1st of G3 at NYR, "
                     "the defenseman's first of the season and the only Islanders goal until the 3rd.", "active"),
        (1, "Goals", "<b>Heineman</b> finished Barzal's pass at the left post to pull G3 at NYR level at 2-2, "
                     "the only goal NYI got out of eight power plays.", "active"),
        (1, "Assists", "<b>Barzal</b> picked up his first point of the season on the tying goal in the 3rd of "
                       "G3 at NYR.", "active"),
        (2, "Games", "<b>Maccelli</b> has a point in each of the last two: the sixth goal in G2 vs. NJD and the "
                     "feed to Romanov in G3 at NYR.", "active"),
        (2, "Games", "<b>Horvat</b>'s point streak ended at two in G3 at NYR.", "past"),
    ],
    summary="Out-shot New York 47-31 and drew eight power plays, but gave up three shorthanded goals and "
            "lost a game they had pulled level at 2-2 in the 3rd.",
    notes="Team stats from the 1st and 2nd intermission screens and the final; per-period shots and hits "
          "differenced from the cumulative ones. Both box scores held for the screens. The Rangers goalie is "
          "not named anywhere, so Opp Goaltending has no row. New York's third shorthanded goal is assigned to "
          "Dorofeyev: the screens put SHG A at 2 after the 2nd and 3 at the final, NYR scored twice in the 3rd "
          "and took no power play in it, and Dorofeyev's came straight off the Heineman power-play goal with "
          "penalties still being served. Lafreniere's is logged even strength.",
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
