# OPTCG Deck Analyzer

Breaks down top-performing One Piece TCG decks into an Excel report: card types,
bricks, counters, searchers and DON!! cost curve. Every deck it downloads is kept
in a local SQLite database, so you can also track how the meta changes over time.

Data comes from two free public APIs, with no scraping and no API key:
- **Decklists:** [Limitless](https://play.limitlesstcg.com/api), which hosts ChinoizeCup and other online sim events
- **Card stats:** [OPTCG API](https://optcgapi.com)

## Setup (once)

Needs Python 3.10 or newer.

**macOS / Linux / WSL**
```bash
git clone https://github.com/chunisama/optcg-deck-analyzer.git
cd optcg-deck-analyzer
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

**Windows (PowerShell or cmd)**
```bat
git clone https://github.com/chunisama/optcg-deck-analyzer.git
cd optcg-deck-analyzer
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
```

## Usage

Run the commands below from the project folder. On macOS/Linux/WSL, type `./optcg`.
On Windows, type `optcg` (in PowerShell: `.\optcg`).

```bash
# Every ChinoizeCup from the last 30 days -> reports/optcg_<date>.xlsx
./optcg analyze

# Top-8 OP17 lists from the last 2 weeks
./optcg analyze --format OP17 --days 14 --top 8

# One leader, custom output file
./optcg analyze --leader "Luffy & Ace" --out reports/luffy_ace.xlsx

# Other events (repeatable)
./optcg analyze --event chinoize --event "treasure cup"

# Your own deck, exported from OPTCGSim (lines like 4xOP12-015)
./optcg deck examples/luffy_ace.txt

# Just update the database, no report (e.g. weekly)
./optcg sync --days 60
```

The first run takes 10–30 seconds to download card stats and tournaments. Later runs
only download new tournaments.

### Options for `analyze` and `sync`

| Option | Default | Meaning |
|---|---|---|
| `--event` | `chinoize` | Tournament name contains this text (repeatable) |
| `--format` | any | Tournament name contains this, e.g. `OP17`, `EGB` |
| `--days` | 30 | How far back to look |
| `--min-players` | 16 | Skip smaller events |
| `--top` | all | *(analyze only)* Only decks that placed this or better |
| `--leader` | all | *(analyze only)* Leader name or card ID, partial match |
| `--out` | `reports/optcg_<date>.xlsx` | *(analyze only)* Output file |
| `--no-sync` | off | *(analyze only)* Report from the database without downloading (works offline) |

`--refresh-cards` re-downloads card stats. It goes before the command
(`./optcg --refresh-cards analyze`), and happens automatically once a week.

## The report

| Sheet | Contents |
|---|---|
| **Decks** | One row per decklist: placing, record, character/event/stage counts, bricks, counter cards (+1k/+2k), searchers, and cards at each cost 0–10+ |
| **Leaders** | The same columns averaged per leader |
| **Card Usage** | Per leader: every card played, % of decks running it, average copies (core vs. flex) |
| **Definitions** | Exactly how each column is counted |

Key definitions:
- **Brick:** any non-leader card with no counter value. The `of which [Counter] events` column shows how many bricks are still defensive events.
- **On Play searcher:** a character whose `[On Play]` looks at the top of the deck and reveals or adds a card to hand. Look-and-reorder effects don't count.
- **Event searcher:** the same, for an event's `[Main]` effect. Counted separately.

Cards missing from OPTCG API (usually brand-new) appear in the `Unknown cards` column and are left out of the counts.

## The database

Everything is stored in `data/optcg.db` (SQLite). Query it with `./optcg sql "..."`,
or open it in [DB Browser for SQLite](https://sqlitebrowser.org/).

| Table / view | Contents |
|---|---|
| `tournament` | id, name, date, format (`OP17`, `EGB`…), players |
| `entry` | One row per player's deck: placing, wins/losses, leader |
| `deck_card` | Cards in each deck: entry, card ID, count |
| `card` | Card stats: type, cost, counter, power, color, text, searcher flags |
| `deck_summary` | **View:** one row per deck with bricks, counters, searchers, events, etc. Start here. |

Example queries:

```bash
# How has Mihawk's build changed between formats?
./optcg sql "SELECT format, COUNT(*) decks, AVG(bricks), AVG(with_counter), AVG(searchers)
             FROM deck_summary WHERE leader_name LIKE '%Mihawk%' GROUP BY format"

# Most-played leaders among top-8 finishes
./optcg sql "SELECT leader_name, leader_id, COUNT(*) top8s FROM deck_summary
             WHERE placing <= 8 GROUP BY leader_id ORDER BY top8s DESC"

# Which +2k counters does Nico Robin play most?
./optcg sql "SELECT c.name, c.id, SUM(dc.count) copies FROM deck_card dc
             JOIN entry e ON e.id = dc.entry_id JOIN card c ON c.id = dc.card_id
             WHERE e.leader_name LIKE '%Robin%' AND c.counter = 2000
             GROUP BY c.id ORDER BY copies DESC"
```

The database is local to each person and isn't committed. Delete `data/optcg.db` to start fresh.

## How it works

1. **Card stats:** downloads all cards from OPTCG API into `card`, and tags searchers and [Counter] events by reading card text.
2. **Sync:** finds matching tournaments on Limitless and stores each new one's decklists in `tournament`, `entry` and `deck_card`. Events less than a day old are re-downloaded next time in case they were still running.
3. **Analyze:** loads the decks matching your filters from the database and counts everything per deck.
4. **Report:** writes the Excel workbook to `reports/`.

All the code is in `analyzer.py`. The only dependencies are `requests` and `openpyxl`.

## License

MIT
