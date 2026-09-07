#!/usr/bin/env python3
"""Regenerate index.html (repo root) and version.txt from `S1 NY Islanders.xlsx`.

Rules carried over from the Royals build spec: pure-ASCII output (entities in HTML,
\\uXXXX in JS, \\25B2-style in CSS), no instruction text, empty sections render a
deliberate `.empty` placeholder, every publish bumps the build id in version.txt and
in <meta name="build">.
"""
import base64
import datetime as dt
import html
import io
import re
import sys
from collections import OrderedDict, defaultdict
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
SITE = ROOT / "site"
WB = ROOT / "S1 NY Islanders.xlsx"
OUT = ROOT / "index.html"
VERSION = ROOT / "version.txt"

TEAM = "New York Islanders"
TAG = "NYI"
SEASON = 1
GAMES_IN_SEASON = 82   # assumption until the schedule is on the page

# ------------------------------------------------------------------ helpers
def esc(s):
    return html.escape("" if s is None else str(s), quote=True)

def money(v):
    """0.975 -> $0.975M ; strings pass through (RFA/UFA/UNSIGNED)."""
    if v is None or v == "":
        return ""
    if isinstance(v, (int, float)):
        return "$%.3fM" % v
    return str(v)

def rows(ws):
    """Data rows (row 3 on) as dicts keyed by the header row; skips fully blank rows."""
    hdr = [c.value for c in ws[1]]
    out = []
    for r in ws.iter_rows(min_row=3, values_only=True):
        if all(v is None or str(v).strip() == "" for v in r):
            continue
        out.append(OrderedDict(zip(hdr, r)))
    return out

def data_uri(path, mime):
    return "data:%s;base64,%s" % (mime, base64.b64encode(path.read_bytes()).decode("ascii"))

def ghost_image():
    """Wordmark, greyscaled at build time, as the page backdrop; the crest until the
    wordmark file exists."""
    for name in ("wordmark.png", "wordmark.jpg", "wordmark.webp"):
        p = SITE / "logos" / name
        if p.exists():
            from PIL import Image
            im = Image.open(p).convert("RGBA")
            g = im.convert("L")
            out = Image.merge("RGBA", (g, g, g, im.getchannel("A")))
            buf = io.BytesIO(); out.save(buf, "PNG", optimize=True)
            return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode("ascii"), "wordmark"
    return data_uri(SITE / "logos" / "crest.svg", "image/svg+xml"), "crest"

def assert_ascii(label, text):
    bad = [(i, ch) for i, ch in enumerate(text) if ord(ch) > 127]
    if bad:
        i, ch = bad[0]
        sys.exit("%s: non-ASCII %r at %d (%s)" % (label, ch, i, text[max(0, i - 30):i + 30]))

# ------------------------------------------------------------------ data
wb = load_workbook(WB, data_only=True)
roster = rows(wb["Roster Ref"])
lines = rows(wb["Lines"])
goals = rows(wb["Owner Goals"])
budget = rows(wb["Budget"])
cap = rows(wb["Cap"])
front = {r["Key"]: r["Value"] for r in rows(wb["Front Office"])}
trades = rows(wb["Trades"])
teams = rows(wb["Teams"])
schedule = rows(wb["Schedule"])
team_name = {t["Abbr"]: t["Team"] for t in teams}

# game logs do not exist yet (columns are the user's call); everything below is 0-0-0
games = []
W = L = OTL = 0
GP = len(games)

as_of = str(front.get("As Of", ""))
as_of_dt = dt.datetime.strptime(as_of, "%m/%d/%Y")
as_of_txt = as_of_dt.strftime("%b %-d, %Y")

def salary_through(p):
    """Last paid season and the tag after it, from the year columns."""
    seasons = ["26-27", "27-28", "28-29", "29-30", "30-31", "31-32", "32-33", "33-34"]
    last = None
    for s in seasons:
        v = p.get(s)
        if isinstance(v, (int, float)) or v == "UNSIGNED":
            last = s
    return last, p.get("Then")

# ------------------------------------------------------------------ pieces
def empty(text):
    return '<p class="empty">%s</p>' % esc(text)

def sec(title, body, meta=None, cls=""):
    m = '<div class="meta">%s</div>' % meta if meta else ""
    return '<section%s><div class="sec-head"><h2>%s</h2>%s</div>%s</section>' % (
        ' class="%s"' % cls if cls else "", esc(title), m, body)

def tile(k, v, good=False, small=None):
    sm = " <small>%s</small>" % esc(small) if small is not None else ""
    return '<div class="stat"><div class="k">%s</div><div class="v num%s">%s%s</div></div>' % (
        esc(k), " good" if good else "", esc(v), sm)

def table(headers, body_rows, cls="num", tfoot=None):
    th = "".join("<th>%s</th>" % esc(h) for h in headers)
    tb = "".join("<tr%s>%s</tr>" % (attrs, "".join(cells)) for attrs, cells in body_rows)
    tf = ('<tfoot><tr class="tot">%s</tr></tfoot>' % "".join(tfoot)) if tfoot else ""
    return ('<div class="tbl-wrap"><div class="tbl-scroll"><table class="%s"><thead><tr>%s</tr></thead>'
            '<tbody>%s</tbody>%s</table></div></div>') % (cls, th, tb, tf)

def td(v, cls=None):
    return "<td%s>%s</td>" % (' class="%s"' % cls if cls else "", v if v is not None else "")

# ---- club logos: .lg-XXX classes, embedded once, used by schedule / hero / divisions
def logo_css():
    out = []
    for f in sorted((SITE / "logos" / "teams").glob("*.svg")):
        out.append(".lg-%s{background-image:url(%s)}" % (f.stem, data_uri(f, "image/svg+xml")))
    return "\n".join(out)

def tlogo(abbr, extra=""):
    return '<span class="tlogo lg-%s%s" aria-hidden="true"></span>' % (esc(abbr), (" " + extra) if extra else "")

# ---- masthead + hero
crest_uri = data_uri(SITE / "logos" / "crest.svg", "image/svg+xml")
ghost_uri, ghost_kind = ghost_image()
record_txt = "%d-%d-%d" % (W, L, OTL)
record_html = ('<div class="record num"><b>%d</b><span class="dash">&ndash;</span><b>%d</b>'
               '<span class="dash">&ndash;</span><b>%d</b></div>') % (W, L, OTL)
phase_pill = ('<div class="streak-pill" style="background:var(--surface-2);color:var(--ink-soft);'
              'border-color:var(--line)"><span class="dot" style="background:var(--orange)"></span>Preseason</div>')
nhl_uri = data_uri(SITE / "logos" / "nhl.svg", "image/svg+xml")
east_uri = data_uri(SITE / "logos" / "east.svg", "image/svg+xml")
league_marks = ('<div class="leagues"><img class="lm lm-nhl" src="%s" alt="NHL"><img class="lm lm-east" src="%s" alt="Eastern Conference"></div>'
                % (nhl_uri, east_uri))

def game_date(g):
    return dt.datetime.strptime(str(g["Date"]), "%m/%d/%Y").date()

next_game = schedule[GP] if GP < len(schedule) else None
if next_game:
    d = game_date(next_game)
    nextgame = ('<div class="nextgame"><span class="ng-k">Next</span><span class="ng-g">G%d</span>'
                '<span class="ng-opp"><span class="loc">%s</span> %s</span>%s'
                '<span class="ng-series"><b>%s</b><small>%s ET</small></span></div>') % (
        int(next_game["G"]), "vs" if next_game["H/A"] == "H" else "@", esc(next_game["Opp"]), tlogo(next_game["Opp"]),
        esc(d.strftime("%a %b %-d")), esc(next_game["Time (ET)"]))
else:
    nextgame = ""
hero = '<div class="hero"><div class="hero-top">%s<div class="hero-right">%s%s</div></div>%s</div>' % (
    record_html, phase_pill, league_marks, nextgame)

# ---- overview
strip = "".join([
    tile("Goals For", 0), tile("Goals Ag.", 0), tile("Goal Diff", 0),
    tile("1-Goal", "0-0"), tile("OT/SO", "0-0"), tile("Shutouts", 0),
    tile("Comebacks", 0, small="0.0%"), tile("PP", "0/0"), tile("PK", "0/0"),
    tile("Most Goals", 0), tile("Most Goals Allowed", 0), tile("BLL", 0),
])
overview = '<div class="strip">%s</div>' % strip
overview += sec("Inside the Numbers", '<div class="story-card"><div class="story-lead"><div class="story-big num">%s</div></div>%s</div>'
                % (record_txt.replace("-", "&ndash;"), empty("No games played")), cls="story")
overview += sec("Recent Headlines", empty("No games played"))
overview += sec("Game Log", empty("0 of %d played" % (len(schedule) or GAMES_IN_SEASON)), meta="0 played &middot; %d season" % (len(schedule) or GAMES_IN_SEASON))
overview += sec("Skating Leaders", empty("None yet"))
overview += sec("Team Skating", empty("No games played"))
overview += sec("Goaltending Leaders", empty("None yet"))
overview += sec("Team Goaltending", empty("No games played"))

# ---- roster
def roster_table(players, goalie=False):
    hdr = ["Player", "Pos", "#", "OVR", "POT", "Age", "Ht", "Wt", "Shoots", "Type", "Clause", "Ext", "Salary", "Through", "Then"]
    body = []
    for p in players:
        thr, then = salary_through(p)
        pot = p["POT"] or ""
        if p.get("POT Cert"):
            pot += ' <span class="pos">%s</span>' % esc(p["POT Cert"])
        name = esc(p["Player"])
        if p.get("Status") == "Scratched":
            name += ' <span class="pos">SCR</span>'
        cl = p.get("Clause") or "-"
        body.append(("", [
            td(name), td(esc(p["Pos"])), td(esc(p.get("#") or "")), td(esc(p["OVR"])), td(pot),
            td(esc(p.get("Age") or "")), td(esc(p.get("Ht") or "")), td(esc(p.get("Wt") or "")),
            td(esc(p.get("Shoots") or "")), td(esc(p.get("Type") or "")), td(esc(cl)),
            td(esc(p.get("Ext") or "-")), td(esc(money(p.get("26-27")))), td(esc(thr or "")), td(esc(then or "")),
        ]))
    return table(hdr, body, cls="num roster")

main_sk = [p for p in roster if p["Group"] == "Main Roster" and p["Pos"] != "G"]
main_g = [p for p in roster if p["Group"] == "Main Roster" and p["Pos"] == "G"]
system = [p for p in roster if p["Group"] == "In the System"]
dressed = [p for p in main_sk if p.get("Status") == "Dressed"]
scratched = [p for p in main_sk if p.get("Status") == "Scratched"]
roster_html = sec("Main Roster", roster_table(main_sk),
                  meta="%d skaters &middot; %d dressed &middot; %d scratched" % (len(main_sk), len(dressed), len(scratched)))
roster_html += sec("Goalies", roster_table(main_g, goalie=True), meta="contracts pending")
roster_html += sec("In the System", roster_table(system), meta="%d skaters" % len(system))
roster_html += sec("Wants Extension", table(["Player", "Pos", "OVR", "Age", "Salary", "Then"],
    [("", [td(esc(p["Player"])), td(esc(p["Pos"])), td(esc(p["OVR"])), td(esc(p["Age"])),
           td(esc(money(p.get("26-27")))), td(esc(p.get("Then") or ""))])
     for p in roster if p.get("Ext") == "Yes"]))

# ---- lines
by_unit = defaultdict(list)
for r in lines:
    by_unit[r["Unit"]].append(r)

def chem_html(v):
    if v is None or v == "":
        return ""
    v = int(v)
    if v > 0:
        return '<span class="chem up">+%d OVR</span>' % v
    if v < 0:
        return '<span class="chem down">%d OVR</span>' % v
    return '<span class="chem">0</span>'

def unit_card(label, unit, order=None):
    ps = by_unit.get(unit, [])
    if order:
        ps = sorted(ps, key=lambda r: order.index(r["Slot"]) if r["Slot"] in order else 99)
    if not ps:
        return ""
    chem = ps[0].get("Chem")
    cells = "".join('<div class="lp"><span class="slot">%s</span><span class="nm">%s</span><span class="ov num">%s</span></div>'
                    % (esc(r["Slot"]), esc(r["Player"]), esc(r["OVR"])) for r in ps)
    return '<div class="line-card"><div class="line-head"><span class="line-k">%s</span>%s</div><div class="lps n%d">%s</div></div>' % (
        esc(label), chem_html(chem), len(ps), cells)

lines_html = sec("Even Strength", "".join([
    unit_card("Line 1", "F1", ["LW", "C", "RW"]), unit_card("Line 2", "F2", ["LW", "C", "RW"]),
    unit_card("Line 3", "F3", ["LW", "C", "RW"]), unit_card("Line 4", "F4", ["LW", "C", "RW"]),
    unit_card("Pair 1", "D1", ["LD", "RD"]), unit_card("Pair 2", "D2", ["LD", "RD"]), unit_card("Pair 3", "D3", ["LD", "RD"]),
]), meta="as of %s" % esc(as_of_txt))
lines_html += sec("Power Play", unit_card("PP1", "PP1", ["LW", "C", "RW", "LD", "RD"]) + unit_card("PP2", "PP2", ["LW", "C", "RW", "LD", "RD"]))
lines_html += sec("Penalty Kill", unit_card("PK1", "PK1", ["LW", "C", "LD", "RD"]) + unit_card("PK2", "PK2", ["LW", "C", "LD", "RD"]) + unit_card("PK3", "PK3", ["LW", "C", "LD", "RD"]))
lines_html += sec("Goalies", unit_card("Goalies", "G", ["1", "2"]))
lines_html += sec("Scratched", unit_card("Scratched", "SCR"))

# ---- front office
def num(v, fmt="%.3f"):
    return (fmt % v) if isinstance(v, (int, float)) else esc(v)

fo_strip = "".join([
    tile("Cap Space", "$%sM" % num(front.get("Cap Space ($M)"))),
    tile("Cap Hit", "$%sM" % num(front.get("Team Cap Hit ($M)"))),
    tile("Salary Cap", "$%sM" % num(front.get("Salary Cap ($M)"))),
    tile("Playoff Cap", "$%sM" % num(front.get("Playoff Cap ($M)"))),
    tile("Contracts", front.get("Contracts")),
    tile("Funds", "$%sM" % num(front.get("Funds Remaining ($M)"))),
    tile("Salary Target", "$%sM" % num(front.get("Salary Target ($M)"))),
    tile("Retained", "$%sM" % num(front.get("Retained Salary ($M)"), "%.0f")),
])
goal_rows = []
for g in goals:
    goal_rows.append(("", [td('<span class="pos" style="margin-left:0">%s</span>' % esc(g["Tier"])), td(esc(g["Goal"])),
                           td("$%s" % format(int(g["Reward"]), ","), "num"), td(esc(g["Eval Date"]), "num"), td(esc(g["Status"]))]))
owner = ('<div class="story-card"><div class="story-lead"><div class="story-big">%s</div><div class="story-lead-txt">%s</div></div>'
         '<ul class="story-list"><li>Owner happiness: <b>%s</b></li><li>State of the team: <b>%s</b></li></ul></div>') % (
    esc(front.get("State of the Team")), esc(front.get("Owner Message")), esc(front.get("Owner Happiness")), esc(front.get("State of the Team")))
fo = '<div class="strip">%s</div>' % fo_strip
fo += sec("Owner", owner, meta="as of %s" % esc(as_of_txt))
fo += sec("Owner Goals", table(["Tier", "Goal", "Reward", "Evaluated", "Status"], goal_rows, cls=""))
fo += sec("Operations Budget", table(["Line", "Allocated", "Spent", "Remaining"],
    [("", [td(esc(b["Line"])), td("$%.3fM" % b["Allocated"]), td("$%.3fM" % b["Spent"]), td("$%.3fM" % b["Remaining"])]) for b in budget],
    tfoot=[td("Total"), td("$%.3fM" % sum(b["Allocated"] for b in budget)), td("$%.3fM" % sum(b["Spent"] for b in budget)),
           td("$%.3fM" % sum(b["Remaining"] for b in budget))]),
    meta="salary target $%sM" % num(front.get("Salary Target ($M)")))
fo += sec("Cap Outlook", table(["Season", "Salary Cap", "Main Roster", "System", "Contracts"],
    [("", [td(esc(c["Season"])), td(("$%.3fM" % c["Salary Cap"]) if c["Salary Cap"] else "-"),
           td("$%.3fM" % c["Main Roster"]), td("$%.3fM" % c["System"]),
           td(esc(c["Contracts"]) if c["Contracts"] is not None else "-")]) for c in cap]),
    meta="skater salaries &middot; goalie deals pending")
fo += sec("Trade Offers", table(["Date", "Team", "NYI sends", "NYI gets", "Result"],
    [("", [td(esc(t["Date"])), td(esc(t["Partner"])), td(esc(t["Out"])), td(esc(t["In"])), td(esc(t["Result"]))]) for t in trades], cls=""),
    meta="%d received" % len(trades))
fo += sec("League Cap Rules", table(["Rule", "Value"],
    [("", [td(esc(k)), td("$%.3fM" % front[k])]) for k in ("Salary Cap ($M)", "Salary Cap Floor ($M)", "Max Player Salary ($M)",
                                                            "Min Player Salary ($M)", "Max Rookie Salary ($M)")], cls=""))

# ---- schedule / goalies / divisions
def schedule_panel():
    if not schedule:
        return sec("Schedule", empty("Regular season schedule not logged yet"), meta="0 of %d played" % GAMES_IN_SEASON)
    out = ""; month = None
    for g in schedule:
        d = game_date(g)
        mlabel = d.strftime("%B %Y") if d.month == 10 or d.month == 1 else d.strftime("%B")
        if mlabel != month:
            if month is not None:
                out += "</div>"
            out += '<div class="sched-month">%s</div><div class="sched">' % esc(mlabel)
            month = mlabel
        home = g["H/A"] == "H"
        gn = int(g["G"])
        cls = "home" if home else "away"
        is_next = next_game is not None and gn == int(next_game["G"])
        out += ('<div class="sgame upcoming %s%s"><span class="g">G%d</span>%s'
                '<div class="mu"><span class="opp"><span class="loc">%s</span> %s</span><span class="club">%s</span></div>'
                '<div class="out"><span class="sdate">%s</span><span class="stime">%s</span></div></div>') % (
            cls, " next" if is_next else "", gn, tlogo(g["Opp"], "big"), "vs" if home else "@", esc(g["Opp"]),
            esc(team_name.get(g["Opp"], "")), esc(d.strftime("%a %b %-d")), esc(g["Time (ET)"]))
    out += "</div>"
    n_home = sum(1 for g in schedule if g["H/A"] == "H")
    return sec("Schedule", out, meta="%d of %d played &middot; %d home &middot; %d away" % (GP, len(schedule), n_home, len(schedule) - n_home))

schedule_html = schedule_panel()
goalie_cards = "".join(
    '<div class="starter"><div class="starter-head"><div><span class="starter-name">%s</span><span class="starter-meta">%s OVR%s</span></div>'
    '<div class="starter-tot"><span>0 GP</span></div></div>%s</div>' % (
        esc(g["Player"]), esc(g["OVR"]), (" &middot; #%s" % esc(g["#"])) if g.get("#") else "", empty("No starts yet"))
    for g in sorted(main_g, key=lambda g: -int(g["OVR"])))
goalies_html = sec("Goalies", goalie_cards, meta="%d on the roster" % len(main_g))

DIV_ORDER = ["Metropolitan", "Atlantic", "Central", "Pacific"]   # own division first
divs = OrderedDict()
for t in sorted(teams, key=lambda t: (DIV_ORDER.index(t["Division"]), t["Team"])):
    divs.setdefault((t["Conference"], t["Division"]), []).append(t)
div_html = ""
for (conf, div), ts in divs.items():
    trs = "".join('<div class="teamrow%s">%s<span class="teamname">%s</span><span class="teamrec muted">%s</span></div>' % (
        " you" if t["Abbr"] == TAG else "", tlogo(t["Abbr"]), esc(t["Team"]), "n/a" if t["Abbr"] == TAG else "0&ndash;0&ndash;0")
        for t in ts)
    div_html += '<div class="div-head"><span class="div-name">%s</span><span class="div-rec">0&ndash;0&ndash;0</span></div><div class="teams">%s</div>' % (esc(div), trs)
divisions_html = sec("vs. Divisions", div_html, meta="record vs each club")

# ------------------------------------------------------------------ page
TABS = [("p-overview", "Overview", overview), ("p-roster", "Roster", roster_html), ("p-lines", "Lines", lines_html),
        ("p-front", "Front Office", fo), ("p-schedule", "Schedule", schedule_html), ("p-goalies", "Goalies", goalies_html),
        ("p-divisions", "vs. Divisions", divisions_html)]
tabs = "".join('<button class="tab%s" role="tab" aria-selected="%s" data-panel="%s">%s</button>' % (
    " is-active" if i == 0 else "", "true" if i == 0 else "false", pid, esc(label)) for i, (pid, label, _) in enumerate(TABS))
panels = "".join('<div class="panel%s" id="%s" role="tabpanel">%s</div>' % (" is-active" if i == 0 else "", pid, body)
                 for i, (pid, _, body) in enumerate(TABS))

build = dt.datetime.now(dt.timezone.utc).strftime("%Y%m%d-%H%M%S")
chrome_css = (SITE / "chrome.css").read_text()
extra_css = (SITE / "extra.css").read_text()
js = (SITE / "site.js").read_text()
for label, txt in (("chrome.css", chrome_css), ("extra.css", extra_css), ("site.js", js)):
    assert_ascii(label, txt)

page = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="color-scheme" content="light dark">
<meta name="build" content="%(build)s">
<meta name="theme-color" content="#00539b">
<link rel="manifest" href="manifest.webmanifest">
<link rel="apple-touch-icon" href="apple-touch-icon.png">
<link rel="icon" type="image/png" sizes="192x192" href="icon-192.png">
<title>Islanders S%(season)d Tracker</title>
<style>
%(chrome)s
%(extra)s
%(logos)s
</style>
</head>
<body>
<div class="bg-crest bg-%(ghost_kind)s" aria-hidden="true"><img src="%(ghost)s" alt=""></div>

<header class="mast">
  <div class="mast-in">
    <button type="button" class="mast-home" aria-label="Back to Overview">
      <div class="crest"><img src="%(crest)s" alt="New York Islanders crest"></div>
    <div>
      <div class="mast-title">Islanders <span>S%(season)d</span> Tracker</div>
    </div>
    </button>
    <div class="mast-right">
      <div class="updated">Preseason<br>%(as_of)s</div>
    </div>
  </div>
</header>

<div class="wrap">
  %(hero)s
  <nav class="tabs" role="tablist" aria-label="Views">%(tabs)s</nav>
  %(panels)s
  <footer>
    %(team)s &middot; Season %(season)d &middot; <b>%(record)s</b> &middot; Preseason
  </footer>
</div>

<script>
%(js)s
</script>
</body>
</html>
""" % dict(build=build, season=SEASON, chrome=chrome_css, extra=extra_css, ghost=ghost_uri, ghost_kind=ghost_kind,
           crest=crest_uri, as_of=esc(as_of_txt), hero=hero, logos=logo_css(), tabs=tabs, panels=panels, team=TEAM, record=record_txt, js=js)

assert_ascii("index.html", page)
OUT.write_text(page)
VERSION.write_text(build + "\n")
print("wrote %s (%d bytes) build %s" % (OUT.name, len(page), build))
