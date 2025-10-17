# The Big Tebowski – League History Dashboard

The Big Tebowski is a Streamlit analytics suite that turns a single Excel workbook into a full league-history portal. It reads box scores, season tables, draft results, and ownership changes, then generates interactive dashboards, streak trackers, rating systems, and rivalry storylines.

This README documents the overall architecture, per-tab visuals, calculation details, and operational tips so you can maintain, extend, or reuse the project with confidence.

---

## Quick Start

- Data source: `fantasy_football.xlsm` stored alongside `app.py` (default path configurable in code).
- Python: 3.9 or newer.
- Install dependencies: `python -m pip install -r requirements.txt`.
- Launch: `streamlit run app.py` from the repository root.
- Filters: use the sidebar to scope the analysis by season, team, or owner. Every tab responds to the filters and recomputes its metrics accordingly.

## High-Level Architecture

- **Streamlit UI (`app.py`)** orchestrates tab rendering and manages user interaction. Each major feature lives in a dedicated render function (overview, finals, regular season, Elo, draft, streaks, head-to-head, etc.).
- **Data access layer** wraps Excel reads with caching via `@st.cache_data`. The loader normalizes header names, handles alternate sheet labels, and gracefully skips missing sheets.
- **Schema normalization** (see `schema_mappings.json`) maps the workbook’s varying column names into a consistent internal schema. When the schema evolves, bump `SCHEMA_VERSION` so cached data refreshes automatically.
- **Processing helpers** perform game-level transformations: winner computation, Elo calculation, streak detection, power index blending, and matchup summaries. They sit close to the render functions for clarity.
- **Plotly visualization** handles interactive line charts, bar charts, heatmaps, and scatter plots. Streamlit handles tables, metrics, and story-card layout.

## Data Expectations

The workbook can be exported straight from the league’s tracking spreadsheet. Sheet names are matched loosely, so variants like `Draft (Round 1)` or `Game Log` still load. Required columns per sheet:

- `gamelog`: Year, Week, Home/Away teams and owners, scores. Winner/loser columns are recomputed if absent. Legacy columns (`Team A Points`, `Owner.1`, etc.) are automatically remapped.
- `reg_season_tables`: Year, TeamName, Owner, Wins, Losses, optional Ties, Points For/Against. Standings, power index, and story cards rely on these values.
- `championship_games`: Finals matchups with owners, scores, and match type.
- `draft`: Round 1 picks with owner and player info.
- `teams_owners`: Team name to owner history. Used to list rosters in the Teams tab and fill missing owner info.
- `records` (optional): Additional text records to surface in the Records tab.

If a sheet is missing, the associated tab displays an informational message but the app continues running.

## Core Tabs and Analytics

- **Overview**
- Summaries highlight total titles, reigning champions, toilet bowl holders, and current Elo leader.
- Dynamic hero cards display recent championship outcomes; tables show finals appearances by owner.
- Latest-season snapshot blends standings with marquee match cards when finals data is absent.

- **Regular Season Progression**
- Weekly standings table with seed indicators, plus inline trends (win%, points-for, margin).
- Cumulative win%, scoring streaks, and “Season progression” Plotly chart that toggles owners via checkboxes.
- Weekly Power Index computed per owner: 40% win percentage, 35% Elo, 25% average scoring margin (weights configurable in `compute_power_index`). The tab exposes:
    - Line chart of power rank across the season.
    - Snapshot slider to inspect any given week’s top 12.
    - Peak table listing each owner’s highest power rank (with season/week context).

- **Consistency & Streaks**
- Re-usable `_compute_condition_streaks` helper tags each game for conditions such as scoring 100+, winning by 30+, or beating league average. The streak tab surfaces:
    - Active streaks table with descriptive context (“6 straight above league average”).
    - All-time streaks leaderboard with longest run and time period.
- Additional streaks analyze consecutive weeks as points leader, most weeks above .500, and more.

- **Elo Ratings**
- Custom Elo implementation seeded at 1000, K-factor 32, with no home-field adjustments.
- Entire calculation re-runs over the filtered data, guaranteeing consistency with the current view.
- Metrics include current standings, deltas week-to-week, all-time peak/valley, longest #1 reign, and biggest jump/drop per game.
- The timeline chart compares selected owners while the match-level table exposes every Elo change for auditing.

- **Championships & Toilet Bowl**
- Cards for each season summarizing champion and runner-up, with scorelines and owner names.
- Grand Final and Toilet Bowl match cards highlight narratives (winner badges, point differential, italicized context).
- Aggregated table lists each owner’s performance in finals across seasons.

- **Draft (Round 1)**
- Displays the Round 1 sheet as a table or filtered view.
- When `ActualVsExpected` columns are available, the app breaks out top over/under performers and correlates draft slot with win percentage.

- **Head-to-Head & Game Log**
- Owner-versus-owner selector populates custom series summary with win counts, margin, and notable streaks.
- Global table shows all head-to-head records (wins, losses, ties, PF/PA, average margin).
- Heatmaps visualize win rate and games played across every owner pairing.
- Current streak tracker states “5 wins, last meeting 2023 W12” for each rivalry.
- **Story cards** use the filtered results plus season win percentages to surface:
    - Biggest upset (largest negative gap between winner and loser season win%).
    - High-scoring heartbreak (most points scored in a loss).
    - Revenge secured (largest-margin rematch flip after a prior loss).
- The bottom of the tab shows the filtered game log with safe column ordering.

- **Teams & Owners**
- Lists franchise history, owner transitions, and aggregate win counts from the season tables.
- Provides hooks for future expansion (mascots, logos, etc.).

## Calculation Details

- **Elo**: For each game, compute expected scores from current ratings, apply `K=32`, and update both owners. Weeks act as the ordering dimension; when multiple games share a week the original Excel order decides processing. Elo standings show delta columns derived from the previous chronological entry.
- **Season progression**: Cumulative records derive from sorted regular-season data. The app handles ties, missing weeks, and partial seasons by forward-filling values where possible.
- **Streaks**: The streak helper groups filtered gamelog rows by owner, assigns boolean flags per condition, and converts runs of `True` into contiguous segments using `shift()` and `cumsum()` logic. Active streaks are the segments that touch the final week in view.
- **Power Index**: Normalizes win%, Elo, and scoring margin on a per-week basis. Win% uses the cumulative record, Elo is imported from the weekly simulation, and margin averages points for/against through the week.
- **Head-to-head stories**: Season-level win% lookup merges regular-season tables by owner-year. Upset detection compares the two win percentages; revenge wins inspect chronological meetings to find a flip in victor.

## Running the App

### Recommended workflow (Windows PowerShell)

```powershell
cd "C:\Users\rtayl\OneDrive\Rob Documents\TheBigTebowski"
python -m pip install -r requirements.txt
streamlit run app.py
```

For a virtual environment:

```powershell
cd "C:\Users\rtayl\OneDrive\Rob Documents\TheBigTebowski"
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
streamlit run app.py
```

The Streamlit launcher prints a local URL (usually http://localhost:8501) for your browser.

### Customizing the data file

- To swap datasets, place a different workbook next to `app.py` and name it `fantasy_football.xlsm`, or edit `DEFAULT_FILE_PATH` in `app.py` to point elsewhere.
- If your columns diverge significantly, update `schema_mappings.json` or extend `normalize_sheet` with additional fallbacks.

## Development Notes

- **Caching**: Streamlit’s `@st.cache_data` caches sheet loads keyed by `SCHEMA_VERSION` and the file hash. Bump the version after schema changes to avoid stale results.
- **Notebook-friendly**: There are no global singletons; render functions accept DataFrames so you can reuse them in tests or Jupyter for debugging.
- **Styling**: Custom CSS applies a light theme, responsive cards, and scrollable tables. Mobile-specific tweaks ensure usability on phones.
- **Error handling**: Functions swallow missing optional sheets gracefully and show `st.info` messages rather than raising exceptions so the dashboard keeps running.
- **Performance**: Heavy calculations (Elo, streak scanning) use vectorized pandas operations. For extremely large leagues, consider precomputing Elo or limiting the filters with default ranges.

## Troubleshooting

- **Excel open in another program**: Close the workbook or clear the OneDrive lock before restarting the app.
- **Missing data**: If Elo or power index sections are blank, confirm `HomeScore`, `AwayScore`, and win/loss columns are numeric in the source sheet.
- **Unexpected owners**: Use the Teams tab to verify owner normalization. Adjust the spreadsheet or update `teams_owners` to merge aliases.
- **Slow first load**: Cache warms up after the initial run. Streamlit restarts will re-read the Excel file only if the signature changed.
- **Streamlit errors on launch**: Confirm dependencies were installed in the environment you used to run `streamlit`.

## Repository Layout

```
TheBigTebowski/
    app.py
    fantasy_football.xlsm
    requirements.txt
    schema_mappings.json
    README.md
    __pycache__/
```

Feel free to fork the project, adjust analytics weightings, or plug in additional sheets (waivers, trades, awards). Contributions that improve story cards or add advanced visuals are welcome.

Happy scouting!
