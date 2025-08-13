# The Big Tebowski – League History

Beautiful Streamlit app to explore your fantasy football league history from a single Excel workbook.

- Data source: `fantasy_football.xlsm` bundled next to `app.py`
- Sheets used: `championship_games`, `teams_owners`, `reg_season_tables`, `draft`, `gamelog` (and `records` if present)
- Modern light UI, Plotly visuals, and sidebar filters for Year/Team/Owner

## Features

- Overview cards and visuals
	- Most Titles (Owner) – gold card
	- Current Champion and Current Toilet Bowl Champion
	- New: Current Elo Rating #1 – gold card with owner and Elo
	- Latest Season snapshot: standings table and last games as match cards
	- Finals appearances/results by owner (badges for Grand Final / Toilet Bowl)
- Elo Ratings (Rating tab)
	- Owner-only Elo (fixed Initial=1000, K=32)
	- Current Elo leaderboard (top 12), full-width modern table
	- Movement indicators since last week: rank ▲/▼ and Elo +/− deltas
	- Timeline chart: “All owners” or “Pick owners…” selector
	- Records: Peak/Lowest Elo (with Year/Week), Most weeks ranked #1, Longest #1 streak (start→end), Biggest win/loss Elo change
	- Match rating details table (per game: start/end Elo and delta)
- Championships & Toilet Bowl tab
	- Season winners, Grand Final/Toilet Bowl match cards, summary tables
- Regular Season tab
	- Modern standings table with seed badges and owner info
- Draft (Round 1) tab
	- Round 1 view from the draft sheet (when available)
- Head-to-Head & Game Log tab
	- Interactive H2H and game log views (when available)
- Teams & Owners tab
	- Teams/owners mappings and details

All tabs respect the sidebar filters (Year/Team/Owner). The Elo page computes ratings from the filtered `gamelog` rows.

## Data file

- Default: `fantasy_football.xlsm` located next to `app.py` is used automatically.
- Supported sheets:
	- `championship_games` – finals and playoffs history
	- `teams_owners` – team and owner mappings over time
	- `reg_season_tables` – season standings
	- `draft` – draft data (Round 1 emphasized)
	- `gamelog` – every game with scores and owners
	- `records` (optional) – additional records to display
- Column normalization: the app tolerates some header variants, e.g. owner columns like `HomeOwner`/`AwayOwner` (and legacy aliases such as `Owner.1`, `Home Owner`, `Away Owner`); score columns inferred if needed (e.g. legacy `Team A Points` / `Team B Points`).

## Requirements

- Python 3.9+ recommended
- See `requirements.txt`:
	- streamlit, pandas, plotly, openpyxl, statsmodels

## Setup and Run (Windows PowerShell)

Option A – quick run:

```powershell
cd "C:\Users\rtayl\OneDrive\Rob Documents\FF"
python -m pip install -r requirements.txt
streamlit run app.py
```

Option B – with a virtual environment (recommended):

```powershell
cd "C:\Users\rtayl\OneDrive\Rob Documents\FF"
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
streamlit run app.py
```

The app uses the bundled `fantasy_football.xlsm` by default. Place your workbook beside `app.py` and keep the same filename, or update the code to point to a different path.

## Tips

- Filters: Use the sidebar to focus on specific years, teams, or owners; the Elo tab recalculates from those filtered games.
- Elo movement arrows: Rank shows ▲/▼ change since the previous week; Elo shows +/− points moved since the previous week.
- Match cards: Finals (Grand Final/Toilet Bowl) display with thematic styling; regular-season “Last Games” are shown if finals aren’t available.

## Troubleshooting

- Excel in use: If OneDrive/Excel has the workbook locked, close it and retry.
- Missing columns: Ensure owner and score columns exist in `gamelog`. The app attempts common fallbacks, but truly missing data can prevent Elo computation.
- No ratings computed: Verify `HomeScore`/`AwayScore` are numeric and not empty for the filtered range.
- Slow load: Large workbooks can take a moment on first run; subsequent loads are faster.

## Directory

```
FF/
	app.py
	fantasy_football.xlsm
	requirements.txt
	README.md
	schema_mappings.json
	.streamlit/  # local Streamlit config (optional)
```

Enjoy exploring your league’s history.
