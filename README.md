# OPTCG Deck Analyzer

Breaks down top-performing One Piece TCG decks into an Excel report: card types,
bricks, counters, searchers and DON!! cost curve. Every deck it downloads is kept
in a local SQLite database, so you can also track how the meta changes over time.

Data sources (all free, no API key or account):
- **Sim decklists:** the [Limitless API](https://docs.limitlesstcg.com/developer.html), for ChinoizeCup events played on OPTCGSim. Includes every player's list and record.
- **In-person decklists:** the [Limitless One Piece website](https://onepiece.limitlesstcg.com/tournaments), for official Regionals, Treasure Cups and Championship Finals. It has no API, so the tool reads the pages: one request per second, and each finished event is fetched only once. Only the top ~16 lists per event are published.
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
# Top-8 decks from every ChinoizeCup and official in-person event, grouped by set
./optcg analyze

# Only sim or only in-person events
./optcg analyze --source sim
./optcg analyze --source in-person --set OP16

# Only the current set (or a specific one, e.g. OP16)
./optcg analyze --set current
./optcg analyze --set OP16

# Every player instead of just the top 8
./optcg analyze --set current --top 0

# One leader, custom output file
./optcg analyze --leader "Luffy & Ace" --out reports/luffy_ace.xlsx

# Other events (repeatable)
./optcg analyze --event chinoize --event "treasure cup"

# Your own deck, exported from OPTCGSim (lines like 4xOP12-015)
./optcg deck examples/luffy_ace.txt

# Just update the database, no report
./optcg sync
```

**The first run downloads the full history:** 180+ ChinoizeCups since July 2025 and 220+
in-person events since 2023. It takes about 10–15 minutes, because the tool paces its
requests to Limitless. The tool waits and retries
automatically, and if you stop it, the next run carries on where it left off. After that,
each run only downloads new tournaments.

### Which decklists end up in the report

The tool doesn't pick or judge decks. It picks **tournaments**, then takes their published decklists:

1. **Sim:** tournaments whose name contains `--event` (default `chinoize`). Every player's registered list is available.
   **In-person:** every Regional, Treasure Cup and Championship Finals listed on Limitless. Only the top ~16 lists are published.
2. Both need at least `--min-players` (16) players.
3. From those, players who placed `--top` 8 or better (1st–8th; sim ties can add a few extra).
4. Optionally narrowed to one `--source`, `--set`, `--leader` or recent `--days`.

The **Sets** sheet in the report shows exactly which sets, date ranges and how many tournaments were included.

**How sets are dated:**
- **Sim:** a set starts at the first ChinoizeCup where 3+ decks play its cards and ends when the next set starts, e.g. OP17 from 2026-08-17. Tournaments tagged `[EGB]` (Extra Grand Battle) are a separate format and labelled on their own (`OP17 · EGB`).
- **In-person:** Limitless's official set tag, e.g. `OP16` or `OP14.5`. Paper releases run behind the sim, so the two sources are always shown side by side, never averaged together.

### Options for `analyze` and `sync`

| Option | Default | Meaning |
|---|---|---|
| `--event` | `chinoize` | Sim tournament name contains this text (repeatable) |
| `--format` | any | Tournament name contains this, e.g. `EGB` |
| `--days` | no limit | Only the last N days |
| `--min-players` | 16 | Skip smaller events |
| `--source` | `all` | `sim`, `in-person` or `all` |
| `--set` | `all` | *(analyze only)* `all`, `current`, or a set code like `OP16` |
| `--top` | 8 | *(analyze only)* Only decks that placed this or better; `0` = every player |
| `--leader` | all | *(analyze only)* Leader name or card ID, partial match |
| `--out` | `reports/optcg_<date>.xlsx` | *(analyze only)* Output file |
| `--no-sync` | off | *(analyze only)* Report from the database without downloading (works offline) |

`--refresh-cards` re-downloads card stats. It goes before the command
(`./optcg --refresh-cards analyze`), and happens automatically once a week.

## The report

| Sheet | Contents |
|---|---|
| **Sets** | Which sets the report covers: date range, tournaments and decks per set |
| **Decks** | One row per decklist: set, placing, record, character/event/stage counts, bricks, counter cards (+1k/+2k), searchers, and cards at each cost 0–10+ |
| **Decklists** | Every card in every deck: card ID, name, copies, type, cost, counter, power, color, and brick/searcher flags. Filter by player or tournament to see a full list. |
| **Leaders** | The same columns as Decks, averaged per set and leader |
| **Card Usage** | Per set and leader: every card played, % of decks running it, average copies (core vs. flex) |
| **Definitions** | Exactly how each column is counted |

Key definitions:
- **Brick:** any non-leader card with no counter value. The `of which [Counter] events` column shows how many bricks are still defensive events.
- **On Play searcher:** a character whose `[On Play]` looks at the top of the deck and reveals or adds a card to hand. Look-and-reorder effects don't count.
- **Event searcher:** the same, for an event's `[Main]` effect. Counted separately.

Cards missing from OPTCG API (usually brand-new) appear in the `Unknown cards` column and are left out of the counts.

## The database

Everything is stored in `data/optcg.db` (SQLite). Query it with `./optcg sql "..."`
(read-only, so a query can't change your data), or open it in [DB Browser for SQLite](https://sqlitebrowser.org/).

| Table / view | Contents |
|---|---|
| `tournament` | id, name, date, players, `source` (`sim` / `in-person`), format tag (`EGB`…), official `set_tag` for in-person |
| `entry` | One row per player's deck: placing, wins/losses (sim only), leader |
| `deck_card` | Cards in each deck: entry, card ID, count |
| `card` | Card stats: type, cost, counter, power, color, text, searcher flags |
| `set_era` | Each set's detected start and end date |
| `deck_summary` | **View:** one row per deck with its source, set and bricks, counters, searchers, events, etc. Start here. |

Example queries:

```bash
# How has Mihawk's build changed from set to set (top 8, sim vs in-person)?
./optcg sql "SELECT source, set_code, COUNT(*) decks, AVG(bricks), AVG(with_counter), AVG(searchers)
             FROM deck_summary WHERE leader_id = 'OP14-020' AND placing <= 8 GROUP BY source, set_code"

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
2. **Sync:** finds matching sim tournaments through the Limitless API, and in-person events from the Limitless website's tournament list and decklist pages. Each new event's decklists go into `tournament`, `entry` and `deck_card`. Recent events are re-checked on later runs in case they were still running or their lists weren't posted yet.
3. **Sets:** works out each sim set's date range from when its cards first appear in decklists (`set_era`). In-person events use their official tag.
4. **Analyze:** loads the decks matching your filters from the database, labels each with its set, and counts everything per deck.
5. **Report:** writes the Excel workbook to `reports/`.

All the code is in `analyzer.py`. The only dependencies are `requests` and `openpyxl`.

If Limitless changes its website layout, the in-person part may stop finding decklists
until `analyzer.py` is updated. The sim part uses the API and isn't affected.

## License

MIT
