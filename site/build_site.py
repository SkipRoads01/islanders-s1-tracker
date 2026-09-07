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
GAMES_IN_SEASON = 84   # 2026-27 CBA expands the regular season from 82 to 84

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
trades = rows(wb["Transactions"]) if "Transactions" in wb.sheetnames else rows(wb["Trades"])
teams = rows(wb["Teams"])
schedule = rows(wb["Schedule"])
team_name = {t["Abbr"]: t["Team"] for t in teams}

def sheet_rows(name):
    return rows(wb[name]) if name in wb.sheetnames else []

games = sheet_rows("Games")
scoring = sheet_rows("Scoring")
goalie_log = sheet_rows("Goalie Game Log")
opp_goalies = sheet_rows("Opp Goaltending")
skater_log = sheet_rows("Skater Game Log")
recaps = sheet_rows("Recaps")
inside = sheet_rows("Inside")
headlines = sheet_rows("Headlines")

GP = len(games)
W = sum(1 for g in games if g["Result"] == "W")
L = sum(1 for g in games if g["Result"] == "L")
OTL = sum(1 for g in games if g["Result"] in ("OTL", "SOL"))
PTS = 2 * W + OTL

def n(v):
    return 0 if v is None else v

GF = sum(n(g["GF"]) for g in games)
GA = sum(n(g["GA"]) for g in games)
SHOTS_F = sum(n(g["Shots F"]) for g in games)
SHOTS_A = sum(n(g["Shots A"]) for g in games)

def pp_pair(v):
    """'0/2' -> (0, 2)"""
    if not v:
        return (0, 0)
    a, b = str(v).split("/")
    return (int(a), int(b))

PP_G = sum(pp_pair(g["PP F"])[0] for g in games)
PP_OPP = sum(pp_pair(g["PP F"])[1] for g in games)
PK_GA = sum(pp_pair(g["PP A"])[0] for g in games)
PK_OPP = sum(pp_pair(g["PP A"])[1] for g in games)

def one_goal(g):
    return abs(n(g["GF"]) - n(g["GA"])) == 1

def rec_of(sel):
    return "%d-%d-%d" % (sum(1 for g in sel if g["Result"] == "W"),
                         sum(1 for g in sel if g["Result"] == "L"),
                         sum(1 for g in sel if g["Result"] in ("OTL", "SOL")))

def went_to_extra(g):
    return n(g["OT F"]) + n(g["OT A"]) + n(g["SO F"]) + n(g["SO A"]) > 0 or g["Result"] in ("OTL", "SOL")

def period_pairs(g, kind):
    keys = [("P1 F", "P1 A"), ("P2 F", "P2 A"), ("P3 F", "P3 A"), ("OT F", "OT A"), ("SO F", "SO A")]
    return [(n(g[a]), n(g[b])) for a, b in keys]

def led_at_any_point(g):
    """Running period-by-period score; did NYI ever hold a lead?"""
    f = a = 0
    for pf, pa in period_pairs(g, "goals"):
        f += pf; a += pa
        if f > a:
            return True
    return False

SHUTOUTS = sum(1 for g in games if n(g["GA"]) == 0)
COMEBACKS = sum(1 for g in games if g["Result"] == "W" and not led_at_any_point(g) and n(g["GA"]) > 0)
BLL = sum(1 for g in games if g["Result"] != "W" and led_at_any_point(g))
MOST_GF = max([n(g["GF"]) for g in games], default=0)
MOST_GA = max([n(g["GA"]) for g in games], default=0)

as_of = str(front.get("As Of", ""))
as_of_dt = dt.datetime.strptime(as_of, "%m/%d/%Y")
as_of_txt = as_of_dt.strftime("%b %-d, %Y")
if games:
    as_of_txt = dt.datetime.strptime(str(games[-1]["Date"]), "%m/%d/%Y").strftime("%b %-d, %Y")
    through_txt = "Through G%d" % int(games[-1]["G"])
else:
    through_txt = "Preseason"

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

# ---- players: key, photo, clickable name, season card
import unicodedata

def pkey(name):
    t = unicodedata.normalize("NFKD", str(name)).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]", "", t.lower())

PLAYERS = {pkey(p["Player"]): p for p in roster}
SURNAME = {}
for k, p in PLAYERS.items():
    sn = str(p["Player"]).split(".")[-1].strip()
    SURNAME.setdefault(pkey(sn), k)

def photo_uri(key):
    for ext, mime in (("png", "image/png"), ("jpg", "image/jpeg"), ("jpeg", "image/jpeg"), ("webp", "image/webp")):
        f = SITE / "logos" / "players" / ("%s.%s" % (key, ext))
        if f.exists():
            return data_uri(f, mime)
    return None

def pname(name, cls="pname"):
    """A roster name becomes a button that opens that player's card."""
    k = pkey(name)
    if k not in PLAYERS:
        k = SURNAME.get(k, "")
    if not k:
        return esc(name)
    return '<button type="button" class="%s" data-p="%s">%s</button>' % (cls, k, esc(name))

def linkify(html_text):
    """Turn <b>Surname</b> in editorial copy into a player button."""
    def sub(m):
        inner = m.group(1)
        k = SURNAME.get(pkey(inner))
        if not k:
            return m.group(0)
        return '<button type="button" class="pname strong" data-p="%s">%s</button>' % (k, inner)
    return re.sub(r"<b>([A-Za-z\'\-]+)</b>", sub, html_text)

# ---- club logos: .lg-XXX classes, embedded once, used by schedule / hero / divisions
def logo_css():
    out = []
    for f in sorted((SITE / "logos" / "teams").glob("*.svg")):
        out.append(".lg-%s{background-image:url(%s)}" % (f.stem, data_uri(f, "image/svg+xml")))
    return "\n".join(out)

def tlogo(abbr, extra=""):
    return '<span class="tlogo lg-%s%s" aria-hidden="true"></span>' % (esc(abbr), (" " + extra) if extra else "")

# ---- goalie season lines
def goalie_season(name):
    rs = [r for r in goalie_log if pkey(r["Goalie"]) == pkey(name)]
    if not rs:
        return None
    sa = sum(n(r["SA"]) for r in rs); sv = sum(n(r["SV"]) for r in rs); ga = sum(n(r["GA"]) for r in rs)
    return dict(gp=len(rs), sa=sa, sv=sv, ga=ga,
                svpct=("%.3f" % (sv / sa)).lstrip("0") if sa else "-",
                w=sum(1 for r in rs if r["Dec"] == "W"), l=sum(1 for r in rs if r["Dec"] == "L"),
                otl=sum(1 for r in rs if r["Dec"] in ("OTL", "SOL")),
                so=sum(n(r["SO"]) for r in rs), starts=sum(1 for r in rs if r["Start/Relief"] == "Start"))

def skater_season(name):
    rs = [r for r in skater_log if pkey(r["Player"]) == pkey(name)]
    if not rs:
        return None
    tot = lambda k: sum(n(r[k]) for r in rs)
    return dict(gp=len(rs), g=tot("G"), a=tot("A"), pts=tot("Pts"), pm=tot("+/-"),
                sog=tot("SOG"), pim=tot("PIM"), hits=tot("Hits"), blk=tot("Blk"))

def player_notes(key):
    """Headline rows and recap lines that name this player."""
    p = PLAYERS[key]
    sn = str(p["Player"]).split(".")[-1].strip()
    out = []
    for h in headlines:
        if re.search(r"\b%s\b" % re.escape(sn), str(h["Text"])):
            out.append((h["Fig"], h["Kicker"], str(h["Text"])))
    return out

def player_card(key):
    p = PLAYERS[key]
    ph = photo_uri(key)
    photo = ('<div class="pm-photo"><img src="%s" alt=""></div>' % ph) if ph else \
            ('<div class="pm-photo empty-photo"><span class="tlogo lg-NYI" aria-hidden="true"></span></div>')
    meta = " &middot; ".join(esc(str(v)) for v in [p["Pos"], ("#%s" % p["#"]) if p["#"] else None,
                                                   ("%s OVR" % p["OVR"]) if p["OVR"] else None,
                                                   ("Age %s" % p["Age"]) if p["Age"] else None,
                                                   "In the System" if p.get("Group") == "In the System" else None] if v)
    bits = []
    for lab, k in (("Height", "Ht"), ("Weight", "Wt"), ("Shoots", "Shoots"), ("Type", "Type"),
                   ("Potential", "POT"), ("Clause", "Clause"), ("Status", "Status")):
        # Group is dropped: "Main Roster" on a main-roster card says nothing. In the System
        # is real information, so it rides in the meta line under the name instead.
        if p.get(k):
            bits.append('<div class="pstat"><span class="k">%s</span><span class="v">%s</span></div>' % (esc(lab), esc(p[k])))
    thr, then = salary_through(p)
    if p.get("26-27"):
        bits.append('<div class="pstat"><span class="k">Salary</span><span class="v">%s</span></div>' % esc(money(p["26-27"])))
    if thr:
        bits.append('<div class="pstat"><span class="k">Through</span><span class="v">%s</span></div>' % esc(thr))
    if then:
        bits.append('<div class="pstat"><span class="k">Then</span><span class="v">%s</span></div>' % esc(then))
    body = '<div class="pstats">%s</div>' % "".join(bits)

    gs = goalie_season(p["Player"])
    ss = skater_season(p["Player"])
    season = ""
    if gs:
        cells = [("Record", "%d-%d-%d" % (gs["w"], gs["l"], gs["otl"])), ("GP", gs["gp"]), ("SV%", gs["svpct"]),
                 ("SA", gs["sa"]), ("SV", gs["sv"]), ("GA", gs["ga"]), ("SO", gs["so"])]
        season = ('<div class="eyebrow" style="margin:14px 0 8px 2px;">Season</div><div class="pstats">%s</div>'
                  % "".join('<div class="pstat"><span class="k">%s</span><span class="v">%s</span></div>' % (esc(k), esc(v)) for k, v in cells))
    elif ss:
        cells = [("GP", ss["gp"]), ("G", ss["g"]), ("A", ss["a"]), ("PTS", ss["pts"]), ("+/-", ss["pm"]),
                 ("SOG", ss["sog"]), ("PIM", ss["pim"]), ("Hits", ss["hits"])]
        season = ('<div class="eyebrow" style="margin:14px 0 8px 2px;">Season</div><div class="pstats">%s</div>'
                  % "".join('<div class="pstat"><span class="k">%s</span><span class="v">%s</span></div>' % (esc(k), esc(v)) for k, v in cells))

    notes = player_notes(key)
    nhtml = ""
    if notes:
        rows_ = "".join('<div class="hl"><div class="hl-fig num">%s</div><div><div class="hl-k">%s</div>'
                        '<div class="hl-txt">%s</div></div></div>' % (esc(f), esc(k), t) for f, k, t in notes)
        nhtml = ('<div class="eyebrow" style="margin:14px 0 8px 2px;">Recent</div><div class="hl-card">%s</div>' % rows_)
    return ('<div class="pm-head">%s<div><div class="pm-name">%s</div><div class="pm-meta">%s</div></div></div>%s%s%s'
            % (photo, esc(p["Player"]), meta, body, season, nhtml))

def player_modal():
    bodies = "".join('<template id="pm-%s">%s</template>' % (k, player_card(k)) for k in PLAYERS)
    return ('<div class="pmodal" id="pmodal" hidden><div class="pmodal-back" data-close></div>'
            '<div class="pmodal-box" role="dialog" aria-modal="true"><button type="button" class="pmodal-x" data-close '
            'aria-label="Close">&times;</button><div id="pmodal-slot"></div></div></div>%s') % bodies

# ---- masthead + hero
crest_uri = data_uri(SITE / "logos" / "crest.svg", "image/svg+xml")
ghost_uri, ghost_kind = ghost_image()
record_txt = "%d-%d-%d" % (W, L, OTL)
record_html = ('<div class="record num"><b>%d</b><span class="dash">&ndash;</span><b>%d</b>'
               '<span class="dash">&ndash;</span><b>%d</b></div>') % (W, L, OTL)
if games:
    st = str(games[-1]["Streak"])
    tone = {"W": ("var(--win-bg)", "var(--win)"), "L": ("var(--loss-bg)", "var(--loss)")}.get(
        st[0] if st[:2] != "OT" and st[:2] != "SO" else "X",
        ("var(--surface-2)", "var(--ink-soft)"))
    dot = "var(--win)" if st.startswith("W") else "var(--loss)" if st.startswith("L") else "var(--orange)"
    phase_pill = ('<div class="streak-pill" style="background:%s;color:%s;border-color:color-mix(in srgb,%s 30%%,transparent)">'
                  '<span class="dot" style="background:%s"></span>%s</div>') % (tone[0], tone[1], tone[1], dot, esc(st))
else:
    phase_pill = ('<div class="streak-pill" style="background:var(--surface-2);color:var(--ink-soft);'
                  'border-color:var(--line)"><span class="dot" style="background:var(--orange)"></span>Preseason</div>')
nhl_uri = data_uri(SITE / "logos" / "nhl.svg", "image/svg+xml")
east_uri = data_uri(SITE / "logos" / "east.svg", "image/svg+xml")
league_marks = ('<div class="leagues"><img class="lm lm-east" src="%s" alt="Eastern Conference">'
                '<img class="lm lm-nhl" src="%s" alt="NHL"></div>' % (east_uri, nhl_uri))

def game_date(g):
    return dt.datetime.strptime(str(g["Date"]), "%m/%d/%Y").date()

played_nums = {int(x["G"]) for x in games}
next_game = next((g for g in schedule if int(g["G"]) not in played_nums), None)
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
def pct(a, b, dec=1):
    return ("%." + str(dec) + "f%%") % (100.0 * a / b) if b else "-"

def trunc_pct(a, b):
    if not b:
        return "0.0%"
    v = int(100000.0 * a / b) / 1000.0
    return "%.1f%%" % v

diff = GF - GA
strip_tiles = [
    tile("Goals For", GF), tile("Goals Ag.", GA),
    tile("Goal Diff", ("+%d" % diff) if diff > 0 else str(diff), good=diff > 0),
    tile("1-Goal", rec_of([g for g in games if one_goal(g)])),
    tile("OT/SO", rec_of([g for g in games if went_to_extra(g)])),
    tile("Shutouts", SHUTOUTS),
    tile("Comebacks", COMEBACKS, small=trunc_pct(COMEBACKS, W)),
    tile("PP", "%d/%d" % (PP_G, PP_OPP), small=pct(PP_G, PP_OPP)),
    tile("PK", "%d/%d" % (PK_OPP - PK_GA, PK_OPP), small=pct(PK_OPP - PK_GA, PK_OPP)),
    tile("Most Goals", MOST_GF), tile("Most Goals Allowed", MOST_GA),
]
if GP >= 10:   # a projection off a handful of games is noise, not a record
    w84 = int(round(84.0 * (2 * W + OTL) / (2.0 * GP) * 1))
    pw = int(round(84.0 * W / GP)); po = int(round(84.0 * OTL / GP))
    strip_tiles.append(tile("Pace", "%d-%d-%d" % (pw, 84 - pw - po, po)))
strip_tiles.append(tile("BLL", BLL))
strip = "".join(strip_tiles)
overview = '<div class="strip">%s</div>' % strip
ins = [x for x in inside if not games or x["G"] == games[-1]["G"]] or inside
if ins:
    lis = "".join("<li>%s</li>" % linkify(str(x["Bullet"])) for x in ins)
    story_body = '<ul class="story-list">%s</ul>' % lis
else:
    story_body = empty("No games played")
overview += sec("Inside the Numbers",
                '<div class="story-card"><div class="story-lead"><div class="story-big num">%s</div></div>%s</div>'
                % (record_txt.replace("-", "&ndash;"), story_body), cls="story")

hls = [x for x in headlines if not games or x["G"] == games[-1]["G"]] or headlines
if hls:
    hrows = "".join('<div class="hl"><div class="hl-fig num%s">%s</div><div><div class="hl-k">%s</div>'
                    '<div class="hl-txt">%s</div></div></div>' % (
                        " past" if str(x["Tone"]).lower() == "past" else "", esc(x["Fig"]), esc(x["Kicker"]),
                        linkify(str(x["Text"]))) for x in hls)
    overview += sec("Recent Headlines", '<div class="hl-card">%s</div>' % hrows)
else:
    overview += sec("Recent Headlines", empty("No games played"))
def box_table(g):
    labels, fcells, acells = [], [], []
    for lab, kf, ka in [("1", "P1 F", "P1 A"), ("2", "P2 F", "P2 A"), ("3", "P3 F", "P3 A")]:
        labels.append(lab); fcells.append(n(g[kf])); acells.append(n(g[ka]))
    if n(g["OT F"]) + n(g["OT A"]) or g["Result"] in ("OTL", "SOL"):
        labels.append("OT"); fcells.append(n(g["OT F"])); acells.append(n(g["OT A"]))
    if n(g["SO F"]) + n(g["SO A"]):
        labels.append("SO"); fcells.append(n(g["SO F"])); acells.append(n(g["SO A"]))
    labels += ["F", "SOG"]; fcells += [n(g["GF"]), n(g["Shots F"])]; acells += [n(g["GA"]), n(g["Shots A"])]
    head = "<th></th>" + "".join("<th>%s</th>" % esc(x) for x in labels)
    def row(abbr, cells, win):
        tds = "".join('<td%s>%s</td>' % (' class="fin"' if i >= len(cells) - 2 else "", c) for i, c in enumerate(cells))
        return '<tr%s><td class="tm">%s</td>%s</tr>' % (' class="wnr"' if win else "", esc(abbr), tds)
    nyi_win = g["Result"] == "W"
    return ('<div class="tbl-wrap"><div class="tbl-scroll"><table class="num box"><thead><tr>%s</tr></thead><tbody>%s%s</tbody></table></div></div>'
            % (head, row(TAG, fcells, nyi_win), row(g["Opp"], acells, not nyi_win)))

TEAMSTATS = [("Shots", "Shots F", "Shots A", None), ("Hits", "Hits F", "Hits A", None),
             ("Time on Attack", "TOA F", "TOA A", None), ("Passing", "Pass% F", "Pass% A", "%"),
             ("Faceoffs Won", "FOW F", "FOW A", None), ("Penalty Minutes", "PIM F", "PIM A", None),
             ("Power Plays", "PP F", "PP A", None), ("Power Play Minutes", "PPM F", "PPM A", None),
             ("Shorthanded Goals", "SHG F", "SHG A", None)]

def teamstat_table(g):
    body = []
    for label, kf, ka, suf in TEAMSTATS:
        vf, va = g[kf], g[ka]
        sf = ("%s%s" % (vf, suf)) if suf else str(vf)
        sa = ("%s%s" % (va, suf)) if suf else str(va)
        body.append(("", [td(esc(sf), "tsv"), td(esc(label), "tsk"), td(esc(sa), "tsv")]))
    return table([TAG, "", g["Opp"]], body, cls="num tstat")

def scoring_table(g):
    rs = [x for x in scoring if x["G"] == g["G"]]
    if not rs:
        return ""
    body = []
    for x in rs:
        who = pname(x["Scorer"]) if x["Team"] == TAG else esc(x["Scorer"])
        if x["Type"] and x["Type"] != "EV":
            who += ' <span class="pos">%s</span>' % esc(x["Type"])
        helpers = ", ".join(esc(v) for v in (x["A1"], x["A2"]) if v) or "unassisted"
        body.append(("", [td(esc(x["Period"])), td('<span class="tabbr sm">%s</span>' % esc(x["Team"])),
                          td(who), td(helpers), td(esc(x["Score"]), "num")]))
    return sec("Scoring", table(["Per", "", "Goal", "Assists", "Score"], body, cls=""))

def recap_html(g):
    rs = [x for x in recaps if x["G"] == g["G"]]
    if not rs:
        return ""
    ps = "".join('<p><span class="inn">%s</span> %s</p>' % (esc(x["Period"]), esc(x["Text"])) for x in rs)
    return '<div class="recap">%s</div>' % ps

def game_card(g, open_first=False):
    home = g["H/A"] == "H"
    res = g["Result"]
    rcls = "w" if res == "W" else "l"
    score = "%d&ndash;%d" % (n(g["GF"]), n(g["GA"]))
    tail = "" if res in ("W", "L") else " " + res
    head = ('<summary class="game"><span class="g">G%d</span><span class="sres %s">%s</span>'
            '%s<span class="mu"><span class="loc">%s</span> %s</span>'
            '<span class="gscore num">%s%s</span><span class="chev">&rsaquo;</span></summary>') % (
        int(g["G"]), rcls, esc(res if res in ("W", "L") else res), tlogo(g["Opp"], "mid"),
        "vs" if home else "@", esc(g["Opp"]), score, esc(tail))
    body = box_table(g) + scoring_table(g) + sec("Team Stats", teamstat_table(g)) + sec("Recap", recap_html(g))
    if g["Summary"]:
        body += '<p class="gsum">%s</p>' % esc(g["Summary"])
    return '<details class="gamed %s"%s>%s<div class="gbody">%s</div></details>' % (
        "home" if home else "away", " open" if open_first else "", head, body)

if games:
    cards = "".join(game_card(g, open_first=(i == 0)) for i, g in enumerate(reversed(games)))
    overview += sec("Game Log", '<div class="games">%s</div>' % cards,
                    meta="%d played &middot; %d season" % (GP, len(schedule) or GAMES_IN_SEASON))
else:
    overview += sec("Game Log", empty("0 of %d played" % (len(schedule) or GAMES_IN_SEASON)),
                    meta="0 played &middot; %d season" % (len(schedule) or GAMES_IN_SEASON))
overview += sec("Skating Leaders", empty("None yet"), meta="skater box score pending")
overview += sec("Team Skating", empty("No skater box score logged"), meta="skater box score pending")

if goalie_log:
    body = []
    for gname in sorted({r["Goalie"] for r in goalie_log}, key=lambda x: -goalie_season(x)["sa"]):
        gs = goalie_season(gname)
        body.append(("", [td(pname(gname)), td("%d-%d-%d" % (gs["w"], gs["l"], gs["otl"])), td(gs["gp"]),
                          td(gs["svpct"]), td(gs["sa"]), td(gs["sv"]), td(gs["ga"]), td(gs["so"])]))
    tot_sa = sum(n(r["SA"]) for r in goalie_log); tot_sv = sum(n(r["SV"]) for r in goalie_log)
    tot_ga = sum(n(r["GA"]) for r in goalie_log)
    foot = [td("Team"), td("%d-%d-%d" % (W, L, OTL)), td(GP),
            td(("%.3f" % (tot_sv / tot_sa)).lstrip("0") if tot_sa else "-"), td(tot_sa), td(tot_sv), td(tot_ga),
            td(sum(n(r["SO"]) for r in goalie_log))]
    overview += sec("Team Goaltending", table(["Goalie", "Record", "GP", "SV%", "SA", "SV", "GA", "SO"], body, tfoot=foot))
else:
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
        name = pname(p["Player"])
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
                    % (esc(r["Slot"]), pname(r["Player"]), esc(r["OVR"])) for r in ps)
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

def pieces(v):
    """'A; B' -> two stacked lines so a long return never widens the table."""
    v = str(v or "-")
    return "<br>".join(esc(x.strip()) for x in v.split(";") if x.strip()) or "-"

def txn_row(t):
    res = str(t["Result"])
    partner = str(t["Partner"] or "-")
    tm = tlogo(partner, "xs") if partner in team_name else '<span class="dash">-</span>'
    return ("", [td(esc(t["Date"])), td(esc(t.get("Type") or "-")), td(tm),
                 td(pieces(t["Out"])), td(pieces(t["In"])),
                 td('<span class="txres %s">%s</span>' % ("ok" if res == "Accepted" else "no", esc(res)))])

declined = sum(1 for t in trades if str(t["Result"]) == "Declined")
gm = sec("Transactions", table(["Date", "Type", "Team", "NYI sends", "NYI gets", "Result"],
    [txn_row(t) for t in trades], cls="txn"),
    meta="%d logged &middot; %d declined" % (len(trades), declined))

# players whose deal runs out after this season - the ones an extension can be offered to
CURRENT_SEASON = "26-27"

def ext_rows():
    out = []
    for p in roster:
        thr, then = salary_through(p)
        if thr != CURRENT_SEASON or p.get(CURRENT_SEASON) == "UNSIGNED":
            continue
        interest = str(p.get("Ext") or "-")
        out.append((p, then, interest))
    order = {"Yes": 0, "-": 1, "No": 2}
    return sorted(out, key=lambda x: (order.get(x[2], 1), -int(x[0]["OVR"] or 0)))

erows = ext_rows()
if erows:
    body = []
    for p, then, interest in erows:
        cls = {"Yes": "ok", "No": "no"}.get(interest, "unk")
        label = {"Yes": "Yes", "No": "No"}.get(interest, "-")
        body.append(("", [td(pname(p["Player"])), td(esc(p["Pos"])), td(esc(p["OVR"])), td(esc(p["Age"])),
                          td(esc(money(p.get(CURRENT_SEASON)))), td(esc(then or "-")),
                          td('<span class="txres %s">%s</span>' % (cls, esc(label))),
                          td(esc("Main" if p["Group"] == "Main Roster" else "System"))]))
    yes = sum(1 for _, _, i in erows if i == "Yes")
    gm += sec("Extension Eligible",
              table(["Player", "Pos", "OVR", "Age", "Salary", "Then", "Interest", "Roster"], body, cls="ext"),
              meta="%d expiring &middot; %d interested" % (len(erows), yes))

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
fo += sec("General Manager", gm, cls="gmsec")
fo += sec("Owner", owner, meta="as of %s" % esc(as_of_txt))
fo += sec("Owner Goals", table(["Tier", "Goal", "Reward", "Evaluated", "Status"], goal_rows, cls="goals"))
fo += sec("Operations Budget", table(["Line", "Allocated", "Spent", "Remaining"],
    [("", [td(esc(b["Line"])), td("$%.3fM" % b["Allocated"]), td("$%.3fM" % b["Spent"]), td("$%.3fM" % b["Remaining"])]) for b in budget],
    tfoot=[td("Total"), td("$%.3fM" % sum(b["Allocated"] for b in budget)), td("$%.3fM" % sum(b["Spent"] for b in budget)),
           td("$%.3fM" % sum(b["Remaining"] for b in budget))], cls="fin"),
    meta="salary target $%sM" % num(front.get("Salary Target ($M)")))
fo += sec("Cap Outlook", table(["Season", "Salary Cap", "Main Roster", "System", "Contracts"],
    [("", [td(esc(c["Season"])), td(("$%.3fM" % c["Salary Cap"]) if c["Salary Cap"] else "-"),
           td("$%.3fM" % c["Main Roster"]), td("$%.3fM" % c["System"]),
           td(esc(c["Contracts"]) if c["Contracts"] is not None else "-")]) for c in cap], cls="fin"),
    meta="skater salaries &middot; goalie deals pending")
fo += sec("League Cap Rules", table(["Rule", "Value"],
    [("", [td(esc(k)), td("$%.3fM" % front[k])]) for k in ("Salary Cap ($M)", "Salary Cap Floor ($M)", "Max Player Salary ($M)",
                                                            "Min Player Salary ($M)", "Max Rookie Salary ($M)")], cls=""))

# ---- schedule / goalies / divisions
def schedule_panel():
    if not schedule:
        return sec("Schedule", empty("Regular season schedule not logged yet"), meta="0 of %d played" % GAMES_IN_SEASON)
    out = ""; month = None
    by_num = {int(x["G"]): x for x in games}
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
def goalie_card(p):
    gs = goalie_season(p["Player"])
    starts = [r for r in goalie_log if pkey(r["Goalie"]) == pkey(p["Player"])]
    ph = photo_uri(pkey(p["Player"]))
    pic = ('<div class="pm-photo sm"><img src="%s" alt=""></div>' % ph) if ph else \
          ('<div class="pm-photo sm empty-photo"><span class="tlogo lg-NYI" aria-hidden="true"></span></div>')
    tot = ('<div class="starter-tot"><span>%d GP</span><span>%d SV</span><span>%s SV%%</span></div>'
           % (gs["gp"], gs["sv"], gs["svpct"])) if gs else '<div class="starter-tot"><span>0 GP</span></div>'
    head = ('<div class="starter-head">%s<div>%s<span class="starter-meta">%s OVR%s</span></div>%s</div>'
            % (pic, '<span class="starter-name">%s</span>' % pname(p["Player"], "pname starter-link"),
               esc(p["OVR"]), (" &middot; #%s" % esc(p["#"])) if p.get("#") else "", tot))
    if not starts:
        return '<div class="starter">%s%s</div>' % (head, empty("No starts yet"))
    rows_ = ""
    for r in starts:
        g = next((x for x in games if x["G"] == r["G"]), None)
        loc = ("vs" if g and g["H/A"] == "H" else "@")
        res = g["Result"] if g else r["Dec"]
        rcls = "w" if res == "W" else "l"
        rows_ += ('<div class="start"><span class="g">G%d <span class="pdec">(%s)</span></span><div class="start-body">'
                  '<div class="start-top"><span class="mu"><span class="loc">%s</span> %s</span>'
                  '<span class="gout %s">%s</span></div><div class="line">%s SV on %s SA &middot; %s GA &middot; %s TOI</div>'
                  '</div></div>') % (int(r["G"]), esc(r["Dec"]), loc, esc(r["Opp"]), rcls, esc(res),
                                     esc(r["SV"]), esc(r["SA"]), esc(r["GA"]), esc(r["TOI"]))
    return '<div class="starter">%s<div class="starts">%s</div></div>' % (head, rows_)

goalies_html = sec("Goalies", "".join(goalie_card(p) for p in sorted(main_g, key=lambda g: -int(g["OVR"]))),
                   meta="%d on the roster" % len(main_g))
if opp_goalies:
    body = [("", [td(esc(r["Goalie"]) + ((' <span class="pos">(L)</span>') if r["Catches"] == "L" else "")),
                  td('<span class="tabbr sm">%s</span>' % esc(r["Opp"])), td("G%d" % int(r["G"])), td(esc(r["Dec"])),
                  td(esc(r["SA"])), td(esc(r["SV"])), td(esc(r["GA"])),
                  td(("%.3f" % r["SV%"]).lstrip("0") if r["SV%"] else "-")]) for r in opp_goalies]
    goalies_html += sec("Goalies Faced", table(["Goalie", "Team", "G", "Dec", "SA", "SV", "GA", "SV%"], body),
                        meta="%d faced" % len({pkey(r["Goalie"]) for r in opp_goalies}))

DIV_ORDER = ["Metropolitan", "Atlantic", "Central", "Pacific"]   # own division first
divs = OrderedDict()
for t in sorted(teams, key=lambda t: (DIV_ORDER.index(t["Division"]), t["Team"])):
    divs.setdefault((t["Conference"], t["Division"]), []).append(t)
div_html = ""
for (conf, div), ts in divs.items():
    def club_rec(abbr):
        sel = [g for g in games if g["Opp"] == abbr]
        return rec_of(sel).replace("-", "&ndash;"), bool(sel)
    cells = []
    for t in ts:
        if t["Abbr"] == TAG:
            cells.append('<div class="teamrow you">%s<span class="teamname">%s</span><span class="teamrec muted">n/a</span></div>'
                         % (tlogo(t["Abbr"]), esc(t["Team"])))
            continue
        r, seen = club_rec(t["Abbr"])
        cells.append('<div class="teamrow">%s<span class="teamname">%s</span><span class="teamrec%s">%s</span></div>'
                     % (tlogo(t["Abbr"]), esc(t["Team"]), "" if seen else " muted", r))
    trs = "".join(cells)
    drec = rec_of([g for g in games if g["Opp"] in {t["Abbr"] for t in ts}]).replace("-", "&ndash;")
    div_html += '<div class="div-head"><span class="div-name">%s</span><span class="div-rec">%s</span></div><div class="teams">%s</div>' % (esc(div), drec, trs)
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
      <div class="updated">%(through)s<br>%(as_of)s</div>
    </div>
  </div>
</header>

<div class="wrap">
  %(hero)s
  <nav class="tabs" role="tablist" aria-label="Views">%(tabs)s</nav>
  %(panels)s
  %(pmodal)s
  <footer>
    %(team)s &middot; Season %(season)d &middot; <b>%(record)s</b> &middot; %(through)s
  </footer>
</div>

<script>
%(js)s
</script>
</body>
</html>
""" % dict(build=build, season=SEASON, chrome=chrome_css, extra=extra_css, ghost=ghost_uri, ghost_kind=ghost_kind,
           crest=crest_uri, as_of=esc(as_of_txt), through=esc(through_txt), hero=hero, logos=logo_css(), pmodal=player_modal(), tabs=tabs, panels=panels, team=TEAM, record=record_txt, js=js)

assert_ascii("index.html", page)
OUT.write_text(page)
VERSION.write_text(build + "\n")
print("wrote %s (%d bytes) build %s" % (OUT.name, len(page), build))
