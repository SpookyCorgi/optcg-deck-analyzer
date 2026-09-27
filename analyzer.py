"""OPTCG deck analyzer.

Pulls tournament decklists from Limitless (sim events through its API, in-person
Regionals / Treasure Cups / Championships from its website) and card stats from
OPTCG API into a local SQLite database, then writes a per-deck breakdown to an
Excel workbook, grouped by source and set.
"""
import argparse
import html
import re
import sqlite3
import sys
import time
from collections import Counter, defaultdict
from datetime import datetime, timedelta, timezone
from pathlib import Path

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).parent
DB_PATH = ROOT / "data" / "optcg.db"
LIMITLESS = "https://play.limitlesstcg.com/api"
LIMITLESS_SITE = "https://onepiece.limitlesstcg.com"  # in-person events; no API, so we read the pages
USER_AGENT = "optcg-deck-analyzer (+https://github.com/chunisama/optcg-deck-analyzer)"
IN_PERSON_TYPES = ("Regional", "Treasure Cup", "Championship")  # official events, by name prefix
OPTCGAPI = "https://optcgapi.com/api"
CARD_CACHE_MAX_AGE = timedelta(days=7)
REQUEST_GAP = 1.0  # seconds between per-tournament requests, to stay under Limitless's rate limits
SET_MIN_DECKS = 3  # a set has "started" at the first tournament where this many decks play its cards
MAX_COST = 10

# Timing tags that end an [On Play] / [Main] section of card text.
_TAG = r"\[(?:On Play|When Attacking|Trigger|Main|Activate: ?Main|Counter|On K\.O\.|End of Your Turn|On Block|Your Turn|Opponent's Turn|On Your Opponent's Attack)\]"
_ON_PLAY = re.compile(r"\[On Play\](.*?)(?=" + _TAG + r"|$)", re.S)
_MAIN = re.compile(r"\[Main\](.*?)(?=" + _TAG + r"|$)", re.S)
_LOOK = re.compile(r"(?:look at|search) .*?cards? from the top of your deck|search your deck", re.I)
_TAKE = re.compile(r"reveal up to|add (?:it|them|up to \d+[^.]*?) to your hand", re.I)
# OPTCG API decorates names: "Nami (062)", "Sabo (004) (Alternate Art)", "Dracule Mihawk - OP14-020".
_NAME_SUFFIX = re.compile(r"(?:\s*\([^()]*\)|\s*-\s*[A-Z0-9]+-\d+)+\s*$")
_FORMAT_TAG = re.compile(r"\[([A-Z0-9]+)\]")
# Tournament formats that are not standard constructed; kept apart from the set's normal events.
SIDE_FORMATS = {"EGB"}

SCHEMA = """
CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT);
CREATE TABLE IF NOT EXISTS card (
    id TEXT PRIMARY KEY, name TEXT, type TEXT, color TEXT,
    cost INTEGER, power INTEGER, counter INTEGER NOT NULL DEFAULT 0, text TEXT,
    on_play_searcher INTEGER NOT NULL, event_searcher INTEGER NOT NULL, counter_event INTEGER NOT NULL
);
-- source: 'sim' (Limitless API) or 'in-person' (Limitless website). set_tag: Limitless's
-- official set for in-person events; sim events get their set from set_era instead.
CREATE TABLE IF NOT EXISTS tournament (
    id TEXT PRIMARY KEY, name TEXT, date TEXT, format TEXT, players INTEGER,
    complete INTEGER NOT NULL DEFAULT 0, source TEXT NOT NULL DEFAULT 'sim', set_tag TEXT
);
CREATE TABLE IF NOT EXISTS entry (
    id INTEGER PRIMARY KEY, tournament_id TEXT NOT NULL REFERENCES tournament(id),
    player TEXT, placing INTEGER, wins INTEGER, losses INTEGER, ties INTEGER,
    leader_id TEXT, leader_name TEXT
);
CREATE TABLE IF NOT EXISTS deck_card (
    entry_id INTEGER NOT NULL REFERENCES entry(id), card_id TEXT NOT NULL, count INTEGER NOT NULL,
    PRIMARY KEY (entry_id, card_id)
);
-- Set eras detected from the decklists (see update_set_eras). end_date is NULL for the current set.
CREATE TABLE IF NOT EXISTS set_era (
    code TEXT PRIMARY KEY, start_date TEXT NOT NULL, end_date TEXT, partial INTEGER NOT NULL
);
CREATE INDEX IF NOT EXISTS entry_tournament ON entry(tournament_id);
CREATE INDEX IF NOT EXISTS entry_leader ON entry(leader_id);

"""

# One row per deck with the headline numbers, for ad-hoc and trend queries.
# Same definitions as analyze() below; cards missing from `card` are not counted.
# Bump VIEW_VERSION whenever this changes so existing databases pick it up.
VIEW_VERSION = "3"
VIEW = """
DROP VIEW IF EXISTS deck_summary;
CREATE VIEW deck_summary AS
SELECT e.id AS entry_id, t.source,
       CASE WHEN t.source = 'in-person' THEN t.set_tag ELSE
       (SELECT code FROM set_era s WHERE t.date >= s.start_date AND (s.end_date IS NULL OR t.date < s.end_date)) END AS set_code,
       t.date, t.format, t.name AS tournament, e.player, e.placing,
       e.wins, e.losses, e.leader_id, e.leader_name,
       SUM(CASE WHEN c.type = 'Character' THEN dc.count ELSE 0 END) AS characters,
       SUM(CASE WHEN c.type = 'Event' THEN dc.count ELSE 0 END) AS events,
       SUM(CASE WHEN c.type = 'Stage' THEN dc.count ELSE 0 END) AS stages,
       SUM(CASE WHEN c.counter = 0 THEN dc.count ELSE 0 END) AS bricks,
       SUM(CASE WHEN c.counter = 0 AND c.counter_event THEN dc.count ELSE 0 END) AS bricks_counter_events,
       SUM(CASE WHEN c.counter > 0 THEN dc.count ELSE 0 END) AS with_counter,
       SUM(CASE WHEN c.counter = 1000 THEN dc.count ELSE 0 END) AS counter_1k,
       SUM(CASE WHEN c.counter = 2000 THEN dc.count ELSE 0 END) AS counter_2k,
       SUM(c.counter * dc.count) AS counter_total,
       SUM(CASE WHEN c.on_play_searcher THEN dc.count ELSE 0 END) AS searchers,
       SUM(CASE WHEN c.event_searcher THEN dc.count ELSE 0 END) AS event_searchers,
       ROUND(1.0 * SUM(MIN(c.cost, 10) * dc.count) / SUM(CASE WHEN c.cost IS NOT NULL THEN dc.count END), 2) AS avg_cost
FROM entry e
JOIN tournament t ON t.id = e.tournament_id
JOIN deck_card dc ON dc.entry_id = e.id
JOIN card c ON c.id = dc.card_id AND c.type != 'Leader'
GROUP BY e.id;
"""


def connect():
    DB_PATH.parent.mkdir(exist_ok=True)
    db = sqlite3.connect(DB_PATH, timeout=30)  # wait for another running sync instead of failing
    db.row_factory = sqlite3.Row
    cols = {r["name"] for r in db.execute("PRAGMA table_info(tournament)")}
    if cols and "source" not in cols:  # database from before in-person support
        db.execute("ALTER TABLE tournament ADD COLUMN source TEXT NOT NULL DEFAULT 'sim'")
        db.execute("ALTER TABLE tournament ADD COLUMN set_tag TEXT")
    db.executescript(SCHEMA)
    row = db.execute("SELECT value FROM meta WHERE key = 'view_version'").fetchone()
    if not row or row["value"] != VIEW_VERSION:
        db.executescript(VIEW)
        with db:
            db.execute("INSERT OR REPLACE INTO meta VALUES ('view_version', ?)", (VIEW_VERSION,))
    return db


def _get(url, json=True, **params):
    for attempt in range(6):
        r = requests.get(url, params=params, timeout=30, headers={"User-Agent": USER_AGENT})
        if r.status_code == 429:
            wait = int(r.headers.get("Retry-After") or 0) or 15 * (attempt + 1)
            print(f"  Rate limited by {url.split('/')[2]}, waiting {wait}s...", file=sys.stderr)
            time.sleep(wait)
            continue
        r.raise_for_status()
        return r.json() if json else r.text
    sys.exit("Still rate limited after several retries. Wait a few minutes and run again; progress so far is saved.")


def _date(s):
    return datetime.fromisoformat(s.replace("Z", "+00:00"))


# ---------------------------------------------------------------- cards

def _num(v):
    try:
        return int(v)
    except (TypeError, ValueError):
        return None


def _is_searcher(text, pattern):
    return any(_LOOK.search(seg) and _TAKE.search(seg) for seg in pattern.findall(text or ""))


def load_cards(db, refresh=False):
    """Return {card_id: card dict}, re-downloading from OPTCG API when stale."""
    row = db.execute("SELECT value FROM meta WHERE key = 'cards_updated'").fetchone()
    stale = not row or datetime.now() - datetime.fromisoformat(row["value"]) > CARD_CACHE_MAX_AGE
    if stale or refresh:
        print("Downloading card database from OPTCG API...", file=sys.stderr)
        raw = []
        for endpoint in ("allSetCards", "allSTCards", "allPromos"):
            raw += _get(f"{OPTCGAPI}/{endpoint}/")
        seen = {}
        for c in raw:
            if c["card_set_id"] in seen:  # parallel/alt-art reprint, same stats
                continue
            text = c.get("card_text") or ""
            seen[c["card_set_id"]] = (
                c["card_set_id"], _NAME_SUFFIX.sub("", c["card_name"]), c["card_type"], c.get("card_color"),
                _num(c.get("card_cost")), _num(c.get("card_power")), _num(c.get("counter_amount")) or 0, text,
                c["card_type"] == "Character" and _is_searcher(text, _ON_PLAY),
                c["card_type"] == "Event" and _is_searcher(text, _MAIN),
                c["card_type"] == "Event" and "[Counter]" in text,
            )
        with db:
            db.executemany("INSERT OR REPLACE INTO card VALUES (?,?,?,?,?,?,?,?,?,?,?)", seen.values())
            db.execute("INSERT OR REPLACE INTO meta VALUES ('cards_updated', ?)", (datetime.now().isoformat(),))
    return {r["id"]: dict(r) for r in db.execute("SELECT * FROM card")}


# ---------------------------------------------------------------- sync

def fetch_tournaments(event, fmt, cutoff, min_players):
    found, page = [], 1
    while True:
        batch = _get(f"{LIMITLESS}/tournaments", game="OP", limit=100, page=page)
        if not batch:
            break
        for t in batch:
            name = t["name"].lower()
            if event and not any(e.lower() in name for e in event):
                continue
            if fmt and fmt.lower() not in name:
                continue
            if (t.get("players") or 0) < min_players:
                continue
            if _date(t["date"]) >= cutoff:
                found.append(t)
        if _date(batch[-1]["date"]) < cutoff:
            break
        page += 1
    return found


def _card_id(c):
    return f"{c['set']}-{c['number']}"


def store_tournament(db, t, standings):
    # Events under a day old may still be running, and lists are occasionally posted late:
    # leave those incomplete so the next sync fetches them again.
    age = datetime.now(timezone.utc) - _date(t["date"])
    has_lists = any(p.get("decklist") for p in standings)
    complete = age > timedelta(days=1) and (has_lists or age > timedelta(days=30))
    m = _FORMAT_TAG.search(t["name"])
    with db:
        db.execute("DELETE FROM deck_card WHERE entry_id IN (SELECT id FROM entry WHERE tournament_id = ?)", (t["id"],))
        db.execute("DELETE FROM entry WHERE tournament_id = ?", (t["id"],))
        db.execute("INSERT OR REPLACE INTO tournament VALUES (?,?,?,?,?,?,'sim',NULL)",
                   (t["id"], t["name"], t["date"], m.group(1) if m else None, t.get("players"), complete))
        for p in standings:
            deck = p.get("decklist")
            if not deck or not deck.get("leader"):
                continue
            r = p.get("record") or {}
            cur = db.execute(
                "INSERT INTO entry (tournament_id, player, placing, wins, losses, ties, leader_id, leader_name) "
                "VALUES (?,?,?,?,?,?,?,?)",
                (t["id"], p["name"], p.get("placing"), r.get("wins", 0), r.get("losses", 0), r.get("ties", 0),
                 _card_id(deck["leader"]), deck["leader"]["name"]))
            counts = Counter()
            for group, items in deck.items():
                if group != "leader":
                    for c in items:
                        counts[_card_id(c)] += c["count"]
            db.executemany("INSERT INTO deck_card VALUES (?,?,?)",
                           [(cur.lastrowid, cid, n) for cid, n in counts.items()])


def sync_cutoff(db, args):
    """--days if given; otherwise a week before the newest matching event we have; otherwise all history."""
    if args.days:
        return datetime.now(timezone.utc) - timedelta(days=args.days)
    where = " OR ".join("name LIKE ?" for _ in args.event)
    row = db.execute(f"SELECT MAX(date) AS d FROM tournament WHERE source = 'sim' AND ({where})", [f"%{e}%" for e in args.event]).fetchone()
    if row["d"]:
        return _date(row["d"]) - timedelta(days=7)
    print("First sync: downloading full history, this takes a few minutes...", file=sys.stderr)
    return datetime(2000, 1, 1, tzinfo=timezone.utc)


def sync(db, args):
    if args.source in ("all", "sim"):
        sync_sim(db, args)
    if args.source in ("all", "in-person"):
        sync_in_person(db, args)


def sync_sim(db, args):
    tournaments = fetch_tournaments(args.event, args.format, sync_cutoff(db, args), args.min_players)
    done = {r["id"] for r in db.execute("SELECT id FROM tournament WHERE complete = 1")}
    # Oldest first: if a long sync is interrupted, the next run resumes from the newest stored event.
    new = sorted((t for t in tournaments if t["id"] not in done), key=lambda t: t["date"])
    for i, t in enumerate(new, 1):
        print(f"  [{i}/{len(new)}] {t['date'][:10]} {t['name']}", file=sys.stderr)
        store_tournament(db, t, _get(f"{LIMITLESS}/tournaments/{t['id']}/standings"))
        if i < len(new):
            time.sleep(REQUEST_GAP)
    print(f"Sim: {len(new)} new/updated, {len(tournaments) - len(new)} already in database.", file=sys.stderr)
    update_set_eras(db)


# ---------------------------------------------------------------- in-person (website)

_ROW_ATTR = re.compile(r'data-(date|name|format|players)="([^"]*)"')
_ROW_LINK = re.compile(r'href="/tournaments/(\d+)"')
_DECK_CARD = re.compile(r'class="decklist-card" data-count="(\d+)" data-id="([^"]+)"')
_TOGGLE = re.compile(r'class="decklist-toggle"[^>]*>(.*?)</div>', re.S)


def _text(fragment):
    return html.unescape(re.sub(r"<[^>]+>", " ", fragment)).split()


def fetch_in_person_events(min_players):
    page = _get(f"{LIMITLESS_SITE}/tournaments", json=False, time="all", show=500)
    events = []
    # Parse each table row on its own so a malformed row can't borrow fields from its neighbour.
    for row in page.split("<tr ")[1:]:
        row = row.split("</tr>")[0]
        attrs, link = dict(_ROW_ATTR.findall(row)), _ROW_LINK.search(row)
        if not link or not {"date", "name", "players"} <= attrs.keys():
            continue
        date, name, fmt, tid = attrs["date"], html.unescape(attrs["name"]), attrs.get("format"), link.group(1)
        players = attrs["players"] if attrs["players"].isdigit() else "0"
        if name.startswith(IN_PERSON_TYPES) and int(players) >= min_players:
            events.append({"id": f"ip-{tid}", "num": tid, "name": name, "date": f"{date}T00:00:00.000Z",
                           "set_tag": fmt or None, "players": int(players or 0)})
    return events


def parse_in_person_decklists(page, label=""):
    """Yield (placing, player, leader_id, {card_id: count}) from a /tournaments/N/decklists page."""
    for block in page.split('<div class="tournament-decklist">')[1:]:
        toggle = _TOGGLE.search(block)
        words = _text(toggle.group(1)) if toggle else []
        m = re.match(r"\D*(\d+)", words[0]) if words else None
        placing = int(m.group(1)) if m else None
        player = " ".join(words[1:]) or None
        cards = [(cid, int(n)) for n, cid in _DECK_CARD.findall(block)]
        if not cards or placing is None:
            # The page layout may have changed; say so rather than silently losing data.
            print(f"  Warning: couldn't read a decklist in {label} ({' '.join(words) or 'no heading'}); skipped. "
                  "If this keeps happening, the Limitless page layout has probably changed.", file=sys.stderr)
            continue
        leader_id = cards[0][0]  # the Leader column is always first
        counts = Counter()
        for cid, n in cards[1:]:
            counts[cid] += n
        yield placing, player, leader_id, counts


def sync_in_person(db, args):
    events = [e for e in fetch_in_person_events(args.min_players)
              if not args.days or _date(e["date"]) >= datetime.now(timezone.utc) - timedelta(days=args.days)]
    done = {r["id"] for r in db.execute("SELECT id FROM tournament WHERE source = 'in-person' AND complete = 1")}
    new = sorted((e for e in events if e["id"] not in done), key=lambda e: e["date"])
    if len(new) > 20:
        print(f"Downloading {len(new)} in-person events from the Limitless website (about 1 per second)...", file=sys.stderr)
    for i, e in enumerate(new, 1):
        print(f"  [{i}/{len(new)}] {e['date'][:10]} {e['name']} ({e['set_tag']})", file=sys.stderr)
        decks = list(parse_in_person_decklists(
            _get(f"{LIMITLESS_SITE}/tournaments/{e['num']}/decklists", json=False), e["name"]))
        # Lists are sometimes added a while after the event; retry recent ones that have none yet.
        complete = bool(decks) or datetime.now(timezone.utc) - _date(e["date"]) > timedelta(days=30)
        with db:
            db.execute("DELETE FROM deck_card WHERE entry_id IN (SELECT id FROM entry WHERE tournament_id = ?)", (e["id"],))
            db.execute("DELETE FROM entry WHERE tournament_id = ?", (e["id"],))
            db.execute("INSERT OR REPLACE INTO tournament VALUES (?,?,?,NULL,?,?,'in-person',?)",
                       (e["id"], e["name"], e["date"], e["players"], complete, e["set_tag"]))
            for placing, player, leader_id, counts in decks:
                cur = db.execute(
                    "INSERT INTO entry (tournament_id, player, placing, leader_id, leader_name) VALUES (?,?,?,?,?)",
                    (e["id"], player, placing, leader_id, None))
                db.executemany("INSERT INTO deck_card VALUES (?,?,?)", [(cur.lastrowid, cid, n) for cid, n in counts.items()])
        if i < len(new):
            time.sleep(REQUEST_GAP)
    print(f"In-person: {len(new)} new/updated, {len(events) - len(new)} already in database.", file=sys.stderr)


# ---------------------------------------------------------------- sets

def update_set_eras(db):
    """Work out when each main set (OPxx) started on the sim, from the first tournament where it was played.

    A set starts at the first standard tournament where at least SET_MIN_DECKS decks play one
    of its cards, and ends when the next set starts. Sets already being played at the oldest
    tournament in the database started earlier than our data, so only the newest of those is
    kept, marked partial.
    """
    first = db.execute(f"""
        SELECT set_code, MIN(date) AS start_date FROM (
            SELECT substr(dc.card_id, 1, 4) AS set_code, t.id, t.date
            FROM deck_card dc JOIN entry e ON e.id = dc.entry_id JOIN tournament t ON t.id = e.tournament_id
            WHERE t.source = 'sim' AND dc.card_id GLOB 'OP[0-9][0-9]-*' AND COALESCE(t.format, '') NOT IN ({','.join('?' * len(SIDE_FORMATS))})
            GROUP BY set_code, t.id HAVING COUNT(DISTINCT e.id) >= {SET_MIN_DECKS})
        GROUP BY set_code ORDER BY start_date, set_code""", list(SIDE_FORMATS)).fetchall()
    if not first:
        return
    oldest = db.execute(
        f"SELECT MIN(date) AS d FROM tournament WHERE source = 'sim' AND COALESCE(format, '') NOT IN "
        f"({','.join('?' * len(SIDE_FORMATS))})", list(SIDE_FORMATS)).fetchone()["d"]
    predating = [r for r in first if r["start_date"] <= oldest]
    eras = ([(predating[-1]["set_code"], oldest, True)] if predating else []) + \
           [(r["set_code"], r["start_date"], False) for r in first if r["start_date"] > oldest]
    with db:
        db.execute("DELETE FROM set_era")
        for i, (code, start, partial) in enumerate(eras):
            end = eras[i + 1][1] if i + 1 < len(eras) else None
            db.execute("INSERT INTO set_era VALUES (?,?,?,?)", (code, start, end, partial))


def load_set_eras(db):
    return [dict(r) for r in db.execute("SELECT * FROM set_era ORDER BY start_date")]


def set_label(eras, date, fmt):
    code = next((e["code"] for e in eras if date >= e["start_date"] and (not e["end_date"] or date < e["end_date"])), "?")
    return f"{code} · {fmt}" if fmt in SIDE_FORMATS else code


# ---------------------------------------------------------------- analysis

def load_entries(db, args):
    eras = load_set_eras(db)
    where, params = ["t.players >= ?"], [args.min_players]
    if args.source != "all":
        where.append("t.source = ?")
        params.append(args.source)
    if args.days:
        where.append("t.date >= ?")
        params.append((datetime.now(timezone.utc) - timedelta(days=args.days)).strftime("%Y-%m-%d"))
    if args.event:
        # --event picks sim series; in-person events are already limited to IN_PERSON_TYPES
        where.append("(t.source = 'in-person' OR " + " OR ".join("t.name LIKE ?" for _ in args.event) + ")")
        params += [f"%{e}%" for e in args.event]
    if args.format:
        where.append("t.name LIKE ?")
        params.append(f"%{args.format}%")
    if args.top:
        where.append("e.placing <= ?")
        params.append(args.top)
    if args.leader:
        where.append("(e.leader_name LIKE ? OR e.leader_id LIKE ?)")
        params += [f"%{args.leader}%"] * 2
    rows = db.execute(
        "SELECT e.*, t.name AS tournament, t.date AS iso_date, t.format, t.source, t.set_tag FROM entry e "
        "JOIN tournament t ON t.id = e.tournament_id WHERE " + " AND ".join(where) +
        " ORDER BY t.date DESC, e.placing", params).fetchall()

    wanted_set = None
    if args.set and args.set.lower() != "all":
        wanted_set = eras[-1]["code"] if args.set.lower() == "current" and eras else args.set.upper()
    entries, ids = [], []
    for r in rows:
        label = (r["set_tag"] or "?") if r["source"] == "in-person" else set_label(eras, r["iso_date"], r["format"])
        if wanted_set and label != wanted_set and not label.startswith(wanted_set + " "):
            continue
        ids.append(r["id"])
        entries.append({
            "id": r["id"], "tournament_id": r["tournament_id"], "set": label,
            "source": "In-person" if r["source"] == "in-person" else "Sim",
            "leader_id": r["leader_id"], "leader_name": r["leader_name"],
            "player": r["player"], "placing": r["placing"],
            "record": f"{r['wins']}-{r['losses']}-{r['ties']}" if r["wins"] is not None else "",
            "tournament": r["tournament"], "date": r["iso_date"][:10],
        })
    counts = defaultdict(Counter)
    for chunk in range(0, len(ids), 500):
        part = ids[chunk:chunk + 500]
        for dc in db.execute(f"SELECT * FROM deck_card WHERE entry_id IN ({','.join('?' * len(part))})", part):
            counts[dc["entry_id"]][dc["card_id"]] = dc["count"]
    names = dict(db.execute("SELECT id, name FROM card"))
    for e in entries:
        e["counts"] = counts[e["id"]]
        e["leader_name"] = f"{e['leader_name'] or names.get(e['leader_id'], '?')} ({e['leader_id']})"
    return entries, eras


def analyze(counts, cards):
    s = Counter()
    curve = Counter()
    unknown = []
    for cid, n in counts.items():
        c = cards.get(cid)
        if not c:
            unknown.append(cid)
            continue
        if c["type"] == "Leader":
            continue
        s["total"] += n
        s[c["type"].lower() + "s"] += n
        if c["counter"] > 0:
            s["with_counter"] += n
            s[f"counter_{c['counter'] // 1000}k"] += n
            s["counter_total"] += c["counter"] * n
        else:
            s["bricks"] += n
            if c["counter_event"]:
                s["bricks_counter_events"] += n
        if c["on_play_searcher"]:
            s["searchers"] += n
        if c["event_searcher"]:
            s["event_searchers"] += n
        if c["cost"] is not None:
            curve[min(c["cost"], MAX_COST)] += n
    s["avg_cost"] = round(sum(k * v for k, v in curve.items()) / max(sum(curve.values()), 1), 2)
    return s, curve, unknown


# ---------------------------------------------------------------- excel

DECK_COLS = [
    ("Characters", "characters"), ("Events", "events"), ("Stages", "stages"),
    ("Bricks (no counter)", "bricks"), ("  of which [Counter] events", "bricks_counter_events"),
    ("Cards w/ counter", "with_counter"), ("+1k counters", "counter_1k"), ("+2k counters", "counter_2k"),
    ("Total counter value", "counter_total"),
    ("On Play searchers", "searchers"), ("Event searchers", "event_searchers"),
    ("Avg cost", "avg_cost"),
]
COST_COLS = [f"{i}{'+' if i == MAX_COST else ''} cost" for i in range(MAX_COST + 1)]
TYPE_ORDER = {"Character": 0, "Event": 1, "Stage": 2}
HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(bold=True, color="FFFFFF")


def _sheet(wb, title, header, rows, widths=None, freeze="B2"):
    ws = wb.create_sheet(title)
    ws.append(header)
    for cell in ws[1]:
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for r in rows:
        ws.append(r)
    ws.freeze_panes = freeze
    ws.auto_filter.ref = ws.dimensions
    for i, h in enumerate(header, 1):
        ws.column_dimensions[get_column_letter(i)].width = (widths or {}).get(h, max(10, min(len(str(h)) + 2, 28)))
    return ws


def _deck_row(s, curve):
    return [s.get(k, 0) for _, k in DECK_COLS] + [curve.get(i, 0) for i in range(MAX_COST + 1)]


def _set_sort(label):
    """Sort 'OP9' before 'OP10' and 'OP14' before 'OP14.5'."""
    m = re.match(r"[A-Z]+(\d+)(?:\.(\d+))?", label or "")
    return (int(m.group(1)), int(m.group(2) or 0), label) if m else (999, 0, label or "")


def _yes(flag):
    return "Yes" if flag else ""


def write_workbook(entries, cards, out, eras=(), selection=""):
    wb = Workbook()
    wb.remove(wb.active)
    for e in entries:
        e["stats"], e["curve"], e["unknown"] = analyze(e["counts"], cards)

    # Sets: which sources, sets and date ranges the report covers
    if selection:
        per_set = defaultdict(lambda: [set(), 0, [], ""])
        for e in entries:
            g = per_set[(e["source"], e["set"])]
            g[0].add(e["tournament_id"])
            g[1] += 1
            g[2].append(e["date"])
        era_by_code = {era["code"]: era for era in eras}
        rows = []
        for (source, label), (tids, n, dates, _) in sorted(per_set.items(), key=lambda kv: (kv[0][0] != "Sim", _set_sort(kv[0][1]))):
            era = era_by_code.get(label.split(" ")[0]) if source == "Sim" else None
            if era:
                start, end = era["start_date"][:10], "(current)" if not era["end_date"] else \
                    (_date(era["end_date"]) - timedelta(days=1)).strftime("%Y-%m-%d")
                note = "Set started before our sim data; start date is our oldest tournament" if era["partial"] else ""
            else:
                start, end, note = min(dates), max(dates), "Official set tag from Limitless" if source == "In-person" else ""
            rows.append([source, label, start, end, len(tids), n, note])
        ws = _sheet(wb, "Sets", ["Source", "Set", "From", "To", "Tournaments", "Decks in report", "Note"], rows,
                    {"Note": 60}, freeze="A2")
        ws.append([])
        ws.append([f"Selection: {selection}"])
        ws.append(["Sim: a set starts at the first ChinoizeCup where 3+ decks play its cards, and ends when the next set starts. "
                   "EGB events are a separate format and listed on their own."])
        ws.append(["In-person: Regionals, Treasure Cups and Championship Finals from the Limitless website, using its "
                   "official set tag. Only the top ~16 decklists are published; 'From'/'To' are the first and last event dates."])

    # Decks: one row per decklist
    header = ["Source", "Set", "Leader", "Player", "Placing", "Record", "Tournament", "Date"] + [c for c, _ in DECK_COLS] + COST_COLS + ["Unknown cards"]
    rows = [[e.get("source", ""), e.get("set", ""), e["leader_name"], e["player"], e["placing"], e["record"], e["tournament"], e["date"]]
            + _deck_row(e["stats"], e["curve"]) + [", ".join(e["unknown"])] for e in entries]
    _sheet(wb, "Decks", header, rows, {"Leader": 26, "Tournament": 34, "Player": 16}, freeze="D2")

    # Decklists: every card in every deck
    header = ["Source", "Set", "Date", "Tournament", "Placing", "Player", "Leader", "Card ID", "Card name", "Copies",
              "Type", "Cost", "Counter", "Power", "Color", "Brick (no counter)", "Searcher", "[Counter] event"]
    rows = []
    for e in entries:
        def order(item):
            c = cards.get(item[0], {})
            return TYPE_ORDER.get(c.get("type"), 9), c.get("cost") if c.get("cost") is not None else 99, item[0]
        for cid, n in sorted(e["counts"].items(), key=order):
            c = cards.get(cid)
            if c and c["type"] == "Leader":
                continue
            if not c:
                rows.append([e.get("source", ""), e.get("set", ""), e["date"], e["tournament"], e["placing"], e["player"],
                             e["leader_name"], cid, "(unknown card)", n] + [""] * 8)
                continue
            rows.append([e.get("source", ""), e.get("set", ""), e["date"], e["tournament"], e["placing"], e["player"],
                         e["leader_name"], cid, c["name"], n, c["type"], c["cost"], c["counter"], c["power"], c["color"],
                         _yes(c["counter"] == 0), _yes(c["on_play_searcher"] or c["event_searcher"]), _yes(c["counter_event"])])
    _sheet(wb, "Decklists", header, rows, {"Tournament": 34, "Leader": 26, "Player": 16, "Card name": 30}, freeze="I2")

    # Leaders: averages per set + leader
    by_leader = defaultdict(list)
    for e in entries:
        by_leader[(e.get("source", ""), e.get("set", ""), e["leader_name"])].append(e)
    groups = sorted(by_leader.items(), key=lambda kv: (kv[0][0] != "Sim", _set_sort(kv[0][1]), -len(kv[1])))
    header = ["Source", "Set", "Leader", "Decks", "Best placing"] + [c for c, _ in DECK_COLS] + COST_COLS
    rows = []
    for (source, set_code, leader), es in groups:
        n = len(es)
        def avg(vals):
            return round(sum(vals) / n, 1)
        placings = [e["placing"] for e in es if e["placing"]]
        rows.append([source, set_code, leader, n, min(placings) if placings else None]
                    + [avg(e["stats"].get(k, 0) for e in es) for _, k in DECK_COLS]
                    + [avg(e["curve"].get(i, 0) for e in es) for i in range(MAX_COST + 1)])
    _sheet(wb, "Leaders", header, rows, {"Leader": 26}, freeze="D2")

    # Card usage per set + leader: core vs flex
    header = ["Source", "Set", "Leader", "Card ID", "Card name", "Type", "Cost", "Counter", "Searcher", "% of decks", "Avg copies (when played)"]
    rows = []
    for (source, set_code, leader), es in groups:
        usage = defaultdict(list)
        for e in es:
            for cid, n in e["counts"].items():
                usage[cid].append(n)
        for cid, ns in sorted(usage.items(), key=lambda kv: (-len(kv[1]), -sum(kv[1]))):
            c = cards.get(cid, {})
            if c.get("type") == "Leader":
                continue
            rows.append([source, set_code, leader, cid, c.get("name", "(unknown card)"), c.get("type"), c.get("cost"), c.get("counter"),
                         _yes(c.get("on_play_searcher") or c.get("event_searcher")),
                         round(100 * len(ns) / len(es)), round(sum(ns) / len(ns), 2)])
    _sheet(wb, "Card Usage", header, rows, {"Leader": 26, "Card name": 30}, freeze="D2")

    # Definitions so the numbers are unambiguous
    defs = [
        ("Source", "Sim = ChinoizeCup on OPTCGSim (Limitless API, every player's list). In-person = official Regionals, Treasure Cups and Championship Finals (Limitless website, top ~16 lists only)."),
        ("Set", "Sim: which set was newest on the sim when the event was played; EGB events are labelled e.g. 'OP17 · EGB'. In-person: Limitless's official set tag, e.g. OP16 or OP14.5."),
        ("Record", "Win-loss-tie. Only available for sim events."),
        ("Bricks (no counter)", "Non-leader cards with no counter value (0 or blank). Includes all events and stages."),
        ("  of which [Counter] events", "Bricks that are events with a [Counter] effect, i.e. still usable on defense."),
        ("Cards w/ counter", "Non-leader cards with a +1000 or +2000 counter value."),
        ("On Play searchers", "Characters whose [On Play] looks at/searches the deck AND reveals or adds a card to hand. Pure 'look and reorder' effects are excluded."),
        ("Event searchers", "Events whose [Main] effect searches the deck the same way. Counted separately."),
        ("N cost", "Number of non-leader cards at that DON!! cost. The last column groups 10+."),
        ("Unknown cards", "Card IDs not found in OPTCG API (usually brand-new cards). They are left out of all counts."),
    ]
    _sheet(wb, "Definitions", ["Column", "Meaning"], defs, {"Column": 28, "Meaning": 110}, freeze="A2")

    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)


# ---------------------------------------------------------------- CLI

def parse_sim_deck(text):
    """Parse an OPTCGSim export: one card per line, like '4xOP12-015'."""
    counts = Counter()
    for line in text.splitlines():
        m = re.match(r"\s*(\d+)\s*x\s*([A-Z0-9]+-\d+)", line, re.I)
        if m:
            counts[m.group(2).upper()] += int(m.group(1))
    return counts


def cmd_sync(args):
    db = connect()
    load_cards(db, args.refresh_cards)
    sync(db, args)


def cmd_analyze(args):
    db = connect()
    cards = load_cards(db, args.refresh_cards)
    if not args.no_sync:
        sync(db, args)
    entries, eras = load_entries(db, args)
    if not entries:
        sys.exit("No decklists matched your filters. Try a different --set, --event or --top.")
    parts = [f"source {args.source}", f"sim events matching {' / '.join(args.event)}", f"set {args.set}",
             f"top {args.top}" if args.top else "all players"]
    parts += [f"last {args.days} days"] if args.days else []
    parts += [f"leader {args.leader}"] if args.leader else []
    parts += [f"name contains {args.format}"] if args.format else []
    out = Path(args.out or f"reports/optcg_{datetime.now():%Y-%m-%d}.xlsx")
    write_workbook(entries, cards, out, eras, ", ".join(parts))
    print(f"Wrote {len(entries)} decks from {len({e['tournament_id'] for e in entries})} tournaments "
          f"across {len({e['set'] for e in entries})} set(s) to {out}")


def cmd_deck(args):
    db = connect()
    cards = load_cards(db, args.refresh_cards)
    counts = parse_sim_deck(Path(args.file).read_text(encoding="utf-8", errors="replace"))
    leader = next((cid for cid in counts if cards.get(cid, {}).get("type") == "Leader"), None)
    if leader:
        counts.pop(leader)
    name = f"{cards[leader]['name']} ({leader})" if leader else "Unknown leader"
    entry = {"source": "", "set": "", "leader_name": name, "player": "(you)", "placing": None, "record": "",
             "tournament": Path(args.file).name, "date": f"{datetime.now():%Y-%m-%d}", "counts": counts}
    out = Path(args.out or f"reports/{Path(args.file).stem}.xlsx")
    write_workbook([entry], cards, out)
    s, curve = entry["stats"], entry["curve"]
    print(f"{name}: {s['total']} cards | bricks {s['bricks']} | counters {s['with_counter']} "
          f"(+1k {s['counter_1k']}, +2k {s['counter_2k']}) | searchers {s['searchers']} | events {s['events']}")
    print("Cost curve: " + "  ".join(f"{i}:{curve.get(i, 0)}" for i in range(MAX_COST + 1)))
    if entry["unknown"]:
        print(f"Unknown cards (not counted): {', '.join(entry['unknown'])}")
    print(f"Wrote {out}")


def cmd_sql(args):
    connect().close()  # make sure the database and view exist
    db = sqlite3.connect(f"{DB_PATH.as_uri()}?mode=ro", uri=True, timeout=30)  # read-only: queries can't change data
    try:
        cur = db.execute(args.query)
        rows = cur.fetchmany(args.limit)
    except sqlite3.Error as e:
        sys.exit(f"SQL error: {e}")
    if not cur.description:
        return
    cols = [d[0] for d in cur.description]
    table = [cols] + [["" if v is None else str(v) for v in r] for r in rows]
    widths = [min(max(len(row[i]) for row in table), 40) for i in range(len(cols))]
    for i, row in enumerate(table):
        print("  ".join(v[:40].ljust(w) for v, w in zip(row, widths)))
        if i == 0:
            print("  ".join("-" * w for w in widths))


def _filters(p):
    p.add_argument("--event", action="append", help="tournament name contains this (repeatable, default: chinoize)")
    p.add_argument("--format", help="only tournaments whose name contains this, e.g. EGB")
    p.add_argument("--days", type=int, help="only the last N days (default: no limit)")
    p.add_argument("--min-players", type=int, default=16, help="skip smaller tournaments (default 16)")
    p.add_argument("--source", choices=["all", "sim", "in-person"], default="all",
                   help="sim = ChinoizeCup via the Limitless API; in-person = Regionals/Treasure Cups/Championships "
                        "from the Limitless website (default: all)")


def main():
    # Player and tournament names can be non-ASCII; don't crash on older Windows consoles.
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(errors="replace")
    ap = argparse.ArgumentParser(prog="optcg", description="Analyze One Piece TCG tournament decks into an Excel report.")
    ap.add_argument("--refresh-cards", action="store_true", help="re-download the card database (auto-refreshes weekly)")
    sub = ap.add_subparsers(dest="cmd", required=True)

    a = sub.add_parser("analyze", help="sync, then write an Excel report from the database")
    _filters(a)
    a.add_argument("--set", default="all", help="all (default), current, or a set code like OP16")
    a.add_argument("--top", type=int, default=8, help="only decks that placed this or better (default 8; 0 = everyone)")
    a.add_argument("--leader", help="only include this leader (name or card ID, partial match)")
    a.add_argument("--out", help="output .xlsx path (default reports/optcg_<date>.xlsx)")
    a.add_argument("--no-sync", action="store_true", help="use only what's already in the database (works offline)")
    a.set_defaults(func=cmd_analyze)

    s = sub.add_parser("sync", help="download new tournaments into the database without writing a report")
    _filters(s)
    s.set_defaults(func=cmd_sync)

    d = sub.add_parser("deck", help="analyze one decklist exported from OPTCGSim")
    d.add_argument("file", help="text file with lines like 4xOP12-015")
    d.add_argument("--out", help="output .xlsx path (default reports/<file>.xlsx)")
    d.set_defaults(func=cmd_deck)

    q = sub.add_parser("sql", help="run a SQL query against the database, e.g. on the deck_summary view")
    q.add_argument("query")
    q.add_argument("--limit", type=int, default=50, help="max rows to print (default 50)")
    q.set_defaults(func=cmd_sql)

    args = ap.parse_args()
    if args.cmd in ("analyze", "sync") and not args.event:
        args.event = ["chinoize"]
    args.func(args)


if __name__ == "__main__":
    main()
