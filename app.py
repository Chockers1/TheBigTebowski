r"""
The Big Tebowski – League History

Streamlit app to explore a fantasy football league's history from a single Excel .xlsm.

Data source (local file on Windows):
    C:\Users\rtayl\OneDrive\Rob Documents\FF\fantasy_football.xlsm

Requirements fulfilled:
- Streamlit + pandas + plotly
- Modern dark UI + responsive layout
- Tabs: Overview, Championships & Toilet Bowl, Regular Season, Draft (Round 1), Head-to-Head & Game Log, Teams & Owners
- Filters: Year (multi), Team/Owner selectors
- Caching for Excel reads; robust file/missing-sheet handling
- Visuals computed only from provided sheets

Run:
    streamlit run app.py
"""

from __future__ import annotations

import os
import urllib.request
import zipfile
from typing import Dict, List, Optional, Tuple
import hashlib

import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
 
    ## 
# --------------------- Config & Constants ---------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_FILE_PATH = os.path.join(BASE_DIR, "fantasy_football.xlsm")
# Increment to invalidate cached reads when schema normalization changes
SCHEMA_VERSION = 6

REQUIRED_SHEETS = {
    "championship_games": [
        "Year",
        "Week",
        "MatchType",
        "WinnerTeam",
        "WinnerOwner",
        "RunnerUpTeam",
        "RunnerUpOwner",
        "WinnerScore",
        "RunnerUpScore",
    ],
    # Year is optional here per the user's sheet; mapping will include Year if present
    "teams_owners": ["TeamName", "Owner"],
    "reg_season_tables": [
        "Year",
        "TeamName",
        "Owner",
        "Wins",
        "Losses",
        "PointsFor",
        "PointsAgainst",
        # Seed optional
    ],
    # 'Team' (NFL team) is optional per requirements
    "draft": ["Year", "Pick", "Owner", "Player", "Position"],
    # 'records' is optional/unstructured; will be handled if present
    "gamelog": [
        "Year",
        "Week",
        "HomeTeam",
        "AwayTeam",
        "HomeOwner",
        "AwayOwner",
        "HomeScore",
        "AwayScore",
        # WinnerTeam optional (we will compute if missing)
        # Notes optional
    ],
}

# Friendly captions for charts by title
CHART_CAPTIONS: Dict[str, str] = {
    # Removed: championships over time charts
    "Average Points per Game by Year": "League-wide average points per game by season.",
    "Average Points per Game by Year (Owner)": "Average points per game by owner for each season.",
    "Titles by Owner": "Total Grand Final wins by owner across seasons.",
    "Titles by Team": "Total Grand Final wins by team across seasons.",
    "Points For (Total)": "Total points scored by each team across the selected seasons.",
    "Wins vs Points For": "Relationship between total points scored and total wins; includes a trendline.",
    "All-time Wins by Owner": "Cumulative wins per owner; color shows win percentage.",
    "Efficiency vs Volume (Games vs Win %)": "How win percentage varies with games played; bubble size is wins.",
    "Net Points per Game (PPG F − PPG A)": "Average points margin per game (points for minus points against).",
    "Cumulative Win% by Year (Owner)": "Cumulative win percentage for each owner over time.",
    "Round 1 Picks by Position": "Which positions are most picked in Round 1 across selected seasons.",
    
    "Round 1 Pick vs Regular-season Win%": "Relationship between draft position and regular-season win percentage.",
    
    "Average Win% by Draft Pick (all seasons)": "Average win percentage for each draft pick across all seasons; labels show sample size.",
    "No. 1 Overall Round 1 Picks by Position": "Which positions have been selected with the #1 overall pick in Round 1 across seasons.",
}


def _strip_strings(df: pd.DataFrame) -> pd.DataFrame:
    """Trim leading/trailing whitespace for all string (object) columns."""
    if df is None or df.empty:
        return df
    for c in df.columns:
        if pd.api.types.is_object_dtype(df[c]):
            df[c] = df[c].apply(lambda v: v.strip() if isinstance(v, str) else v)
    return df


def _style_light_ui() -> None:
    """Inject custom CSS for a clean, modern light theme and responsive look."""
    st.markdown(
        """
        <style>
            :root {
                --bg: #ffffff;
                --panel: #f5f7fb;
                --text: #1f2937;
                --muted: #6b7280;
                --accent: #2563eb;
                --border: rgba(0,0,0,0.08);
            }
            .stApp { background: var(--bg); color: var(--text); }
            .block-container { padding-top: 1.2rem; max-width: 1400px; }
            header[data-testid="stHeader"] { background: rgba(255,255,255,0.6); backdrop-filter: blur(8px); }
            h1, h2, h3 { color: var(--text); letter-spacing: 0.2px; }
            .small-muted { color: var(--muted); font-size: 0.95rem; }
            .stTabs [data-baseweb="tab-list"] { gap: 8px; }
            .stTabs [data-baseweb="tab"] {
                background: var(--panel);
                border-radius: 10px; padding: 10px 16px; color: var(--text);
                border: 1px solid var(--border);
            }
            .stTabs [aria-selected="true"] {
                background: #fff; border: 1px solid rgba(0,0,0,0.12);
                box-shadow: 0 1px 3px rgba(0,0,0,0.06);
            }
            .metric-card {
                background: #fff; padding: 14px 18px; border-radius: 12px;
                border: 1px solid var(--border); box-shadow: 0 1px 3px rgba(0,0,0,0.04);
            }
            /* Records grid & cards */
            .record-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 12px; }
            .record-card { background:#fff; border:1px solid var(--border); border-radius:12px; padding: 12px 14px; box-shadow: 0 1px 3px rgba(0,0,0,0.04); }
            .record-title { font-size:.9rem; color: var(--muted); display:flex; align-items:center; gap:.5rem; }
            .record-stat { font-size:1.6rem; font-weight:700; margin:.2rem 0 .1rem 0; color: var(--text); }
            .record-sub { font-size:.9rem; color: var(--muted); }
            .pill { display:inline-block; background: var(--panel); color: var(--text); border:1px solid var(--border); border-radius: 999px; padding: 2px 8px; font-size:.75rem; }
            /* Winner indicators */
            .winner-badge { background: linear-gradient(135deg, #10b981, #059669); color: white; font-weight: 600; padding: 4px 10px; border-radius: 20px; font-size: .8rem; }
            .loser-badge { background: linear-gradient(135deg, #ef4444, #dc2626); color: white; font-weight: 600; padding: 4px 10px; border-radius: 20px; font-size: .8rem; }
            /* Season winners grid */
            .season-winners { display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 16px; margin: 20px 0; }
            .champion-card { 
                background: #ffffff;
                border: 1px solid var(--border);
                border-left: 5px solid #f59e0b;
                padding: 16px; 
                border-radius: 12px; 
                text-align: left; 
                box-shadow: 0 2px 8px rgba(0,0,0,0.06);
                transition: transform 0.2s ease, box-shadow 0.2s ease;
            }
            .champion-card:hover {
                transform: translateY(-3px);
                box-shadow: 0 4px 12px rgba(0,0,0,0.1);
            }
            .champion-year { font-size: 0.9rem; font-weight: 600; color: #6b7280; margin-bottom: 8px; }
            .champion-team { font-size: 1.2rem; font-weight: 700; color: #1f2937; margin-bottom: 4px; }
            .champion-owner { font-size: 1rem; color: #4b5563; }
            .champion-matchup { font-size: 0.85rem; color: #6b7280; font-style: italic; margin-bottom: 4px; }
            .loser-card {
                background: #ffffff;
                border: 1px solid var(--border);
                border-left: 5px solid #a16207; /* Brown accent */
                padding: 16px; 
                border-radius: 12px; 
                text-align: left; 
                box-shadow: 0 2px 8px rgba(0,0,0,0.06);
            }
            /* Overview metric mini-cards */
            .overview-cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 14px; margin: 10px 0 18px; }
            .overview-card { background:#fff; border:1px solid var(--border); border-radius:14px; padding: 12px 14px; box-shadow: 0 1px 4px rgba(0,0,0,0.05); display:flex; flex-direction:column; gap:6px; min-height:84px; }
            .overview-card.titles { border-left:4px solid #f59e0b; }
            .overview-card.champ { border-left:4px solid #10b981; }
            .overview-card.toilet { border-left:4px solid #a16207; }
            .overview-card .head { display:flex; align-items:center; gap:8px; color: var(--muted); font-size:.85rem; }
            .overview-card .emoji { width:26px; height:26px; display:grid; place-items:center; border-radius:999px; font-size:1rem; }
            .overview-card.titles .emoji { background:#fff7ed; }
            .overview-card.champ .emoji { background:#ecfdf5; }
            .overview-card.toilet .emoji { background:#fdf6e7; }
            .overview-card .value { font-size:1.05rem; font-weight:700; color: var(--text); line-height:1.2; }
            .overview-card .sub { font-size:.82rem; color: var(--muted); }

            /* Modern table styling for standings */
            .standings-wrap { background:#fff; border:1px solid var(--border); border-radius:14px; box-shadow:0 1px 4px rgba(0,0,0,0.05); overflow:hidden; }
            .standings-title { padding:12px 14px; font-weight:700; border-bottom:1px solid var(--border); display:flex; align-items:center; gap:8px; }
            .table-modern { width:100%; border-collapse:separate; border-spacing:0; }
            .table-modern thead th { position:sticky; top:0; background:#f9fafb; color:#374151; font-weight:600; font-size:.9rem; padding:10px 12px; border-bottom:1px solid var(--border); text-align:left; }
            .table-modern tbody td { padding:10px 12px; border-bottom:1px solid #f1f5f9; font-size:.92rem; color:#111827; }
            .table-modern tbody tr:hover { background:#f9fafb; }
            .seed-badge { display:inline-grid; place-items:center; min-width:28px; height:28px; padding:0 8px; border-radius:999px; font-weight:700; font-size:.85rem; border:1px solid var(--border); background:#f3f4f6; color:#111827; }
            .seed-1 { background:linear-gradient(135deg,#fef3c7,#fde68a); border-color:#f59e0b; }
            .seed-2 { background:linear-gradient(135deg,#e5e7eb,#d1d5db); border-color:#9ca3af; }
            .seed-3 { background:linear-gradient(135deg,#fce7f3,#fbcfe8); border-color:#db2777; }
            .wl { font-variant-numeric: tabular-nums; color:#374151; }
            .pct { font-variant-numeric: tabular-nums; color:#111827; font-weight:700; }
            .owner-sub { color:#6b7280; font-size:.85rem; }

            /* Finals match cards */
            .match-grid { display:grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap:14px; }
            .match-card { background:#fff; border:1px solid var(--border); border-radius:14px; box-shadow:0 1px 4px rgba(0,0,0,0.05); overflow:hidden; }
            .match-card.gfinal { border-left:5px solid #f59e0b; }
            .match-card.toilet { border-left:5px solid #a16207; }
            .match-header { padding:12px 14px; display:flex; align-items:center; gap:8px; border-bottom:1px solid var(--border); color:#374151; font-weight:600; }
            .match-header .emoji { font-size:1.1rem; }
            .match-body { padding:10px 12px; display:flex; flex-direction:column; gap:8px; }
            .team-row { display:flex; align-items:center; justify-content:space-between; gap:10px; }
            .team-info { display:flex; flex-direction:column; }
            .team-name { font-weight:700; color:#111827; }
            .team-owner { color:#6b7280; font-size:.85rem; }
            .team-score { font-weight:800; font-variant-numeric: tabular-nums; }
            .win .team-name { color:#065f46; }
            .win .team-score { color:#065f46; }
            .loss .team-name { color:#991b1b; }
            .loss .team-score { color:#991b1b; }
            .tie .team-name, .tie .team-score { color:#6b7280; }

            /* Finals summary badges */
            .type-badge { display:inline-block; padding:2px 8px; border-radius:999px; font-weight:600; font-size:.75rem; border:1px solid var(--border); background:#f3f4f6; color:#374151; }
            .type-badge.gf { background:#fff7ed; border-color:#f59e0b; color:#92400e; }
            .type-badge.tb { background:#fdf6e7; border-color:#a16207; color:#7c2d12; }
        </style>
        """,
        unsafe_allow_html=True,
    )


# --------------------- Data Loading & Caching ---------------------


@st.cache_data(show_spinner=False)
def load_sheet(
    path: str,
    sheet_name: str,
    schema_version: int = SCHEMA_VERSION,
    file_sig: Optional[str] = None,
) -> Optional[pd.DataFrame]:
    """Load a sheet from the Excel file with caching.

    Returns None if file/sheet not found.
    """
    try:
        if not os.path.exists(path):
            return None
        # First try exact sheet name
        try:
            df = pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")
            try:
                df.attrs["__requested_sheet__"] = sheet_name
                df.attrs["__source_sheet__"] = sheet_name
            except Exception:
                pass
        except ValueError:
            # Fallback: try to find a sheet with a similar/normalized name
            try:
                xls = pd.ExcelFile(path, engine="openpyxl")
                target = _norm(sheet_name)
                # Build candidates by normalized equality, then contains
                normalized = {s: _norm(s) for s in xls.sheet_names}
                # Exact normalized match first
                exact = [s for s, n in normalized.items() if n == target]
                chosen = exact[0] if exact else None
                if chosen is None:
                    # Any that contains the target token (e.g., "draft" in "draft (round 1)")
                    contains = [s for s, n in normalized.items() if target in n or n in target]
                    chosen = contains[0] if contains else None
                if chosen is None:
                    return None
                df = pd.read_excel(path, sheet_name=chosen, engine="openpyxl")
                try:
                    df.attrs["__requested_sheet__"] = sheet_name
                    df.attrs["__source_sheet__"] = chosen
                except Exception:
                    pass
            except Exception:
                return None
        df.columns = [str(c).strip() for c in df.columns]
        # Normalize to the app's internal schema; include schema_version in the cache key
        # Touch schema_version and file_sig so Streamlit includes them in the cache key.
        _ = (schema_version, file_sig)
        return normalize_sheet(sheet_name, df)
    except Exception:
        return None


def _file_signature(path: str) -> Optional[str]:
    """Compute a stable signature for the data file to bust Streamlit cache when it changes."""
    try:
        if not os.path.exists(path):
            return None
        with open(path, "rb") as f:
            data = f.read()
        return hashlib.md5(data).hexdigest()
    except Exception:
        return None


def ensure_required_columns(df: pd.DataFrame, required: List[str]) -> Tuple[pd.DataFrame, List[str]]:
    """Ensure that the DataFrame contains the required columns.

    Returns a tuple of (possibly-trimmed DataFrame, missing_columns).
    The function does not raise; caller can decide how to display warnings.
    """
    df = df.copy()
    cols = list(df.columns)
    missing = [c for c in required if c not in cols]
    ordered = [c for c in required if c in cols] + [c for c in cols if c not in required]
    return df[ordered], missing


def _norm(s: str) -> str:
    """Normalize a column name: lowercase, strip, collapse spaces/punctuation."""
    return (
    str(s).replace("\u00A0", " ")  # convert NBSP to space
        .strip()
        .lower()
        .replace(".", " ")
        .replace("_", " ")
        .replace("-", " ")
    )


def normalize_sheet(sheet: str, df: Optional[pd.DataFrame]) -> Optional[pd.DataFrame]:
    """Normalize a sheet's columns to the app's internal schema based on user-provided headers.

    Mappings based on user attachments:
    - gamelog:  Season, Week, Team A, Owner, Team A Points, Team B, Owner, Team B Points
    - championship_games: Season, Week, Home Team, Owner, Points, Away Team, Owner, Points
    - teams_owners: Home Team, Owner
    - reg_season_tables: Rank, Team, Owner, P, W, L, T, Pct, For, Against, PPG F, PPG A, Season
    - draft: Season, Pick, Owner, Player, Position
    """
    if df is None or df.empty:
        return df

    cols = list(df.columns)
    norm_cols = [_norm(c) for c in cols]

    def get_col(targets: List[str], multiple: bool = False) -> Optional[List[str] | str]:
        """Find the first matching column(s) by normalized name. If multiple, return list of matches in order."""
        matches = []
        for i, nc in enumerate(norm_cols):
            for t in targets:
                if nc == _norm(t):
                    matches.append(cols[i])
                    break
        if not matches:
            return None
        return matches if multiple else matches[0]

    sheet_nc = _norm(sheet)

    # Clone to avoid mutating original
    out = df.copy()

    if sheet_nc == _norm("gamelog"):
        # Explicitly handle duplicates by order (Owner and Points appear twice)
        # Determine Owner columns, prefer explicit A/B owner names
        owner_a_col = get_col(["a owner", "team a owner", "owner a", "home owner"]) or None
        owner_b_col = get_col(["b owner", "team b owner", "owner b", "away owner"]) or None
        # Fallback: determine by order of any columns that contain 'owner'
        owner_indices = [i for i, nc in enumerate(norm_cols) if "owner" in nc]
        points_a = get_col(["team a points"]) or get_col(["teama points"]) or get_col(["points a"])  # fallback
        points_b = get_col(["team b points"]) or get_col(["teamb points"]) or get_col(["points b"])  # fallback
        mapping = {}
        # Base mapping
        mapping[get_col(["season"]) or "Season"] = "Year"
        if get_col(["week"]):
            mapping[get_col(["week"])] = "Week"
        if get_col(["team a"]) :
            mapping[get_col(["team a"]) ] = "HomeTeam"
        if get_col(["team b"]) :
            mapping[get_col(["team b"]) ] = "AwayTeam"
        # Owners: prefer explicit A/B owner columns; else by order
        if owner_a_col:
            mapping[owner_a_col] = "HomeOwner"
        if owner_b_col:
            mapping[owner_b_col] = "AwayOwner"
        if not owner_a_col or not owner_b_col:
            # Fill missing from ordered owner-like columns
            if len(owner_indices) >= 1 and "HomeOwner" not in mapping:
                mapping[cols[owner_indices[0]]] = "HomeOwner"
            if len(owner_indices) >= 2 and "AwayOwner" not in mapping:
                mapping[cols[owner_indices[1]]] = "AwayOwner"
        # Points by explicit names
        if points_a:
            mapping[points_a] = "HomeScore"
        if points_b:
            mapping[points_b] = "AwayScore"
        out = out.rename(columns=mapping)
        return out

    if sheet_nc == _norm("championship_games"):
        # Map columns to Home/Away then compute Winner/RunnerUp columns
        # Handle duplicate headers like 'Owner'/'Owner.1' and 'Points'/'Points.1'
        owner_indices = [i for i, nc in enumerate(norm_cols) if nc.startswith("owner")]
        points_indices = [i for i, nc in enumerate(norm_cols) if nc.startswith("points")]
        # Build new column names list to handle duplicate headers robustly
        new_cols = out.columns.tolist()
        # Year
        season_col = get_col(["season"])
        if season_col:
            new_cols[cols.index(season_col)] = "Year"
        elif "Season" in cols:
            new_cols[cols.index("Season")] = "Year"
        # Week
        wk_col = get_col(["week"])
        if wk_col:
            new_cols[cols.index(wk_col)] = "Week"
        # Teams
        ht_col = get_col(["home team"])
        if ht_col:
            new_cols[cols.index(ht_col)] = "HomeTeam"
        at_col = get_col(["away team"])
        if at_col:
            new_cols[cols.index(at_col)] = "AwayTeam"
    # Owners/Points: prefer explicit 'Home Points'/'Away Points' if present; otherwise choose nearest-right columns
        try:
            home_team_idx = cols.index(ht_col) if ht_col else None
        except ValueError:
            home_team_idx = None
        try:
            away_team_idx = cols.index(at_col) if at_col else None
        except ValueError:
            away_team_idx = None

        def _nearest_right(indices, anchor):
            if anchor is None:
                return None
            right = [i for i in indices if i > anchor]
            return min(right) if right else None

        # Home Owner/Score
        ho_i = _nearest_right(owner_indices, home_team_idx)
        hs_i = _nearest_right(points_indices, home_team_idx)
        if ho_i is not None:
            new_cols[ho_i] = "HomeOwner"
        # Explicit Home Points if present
        hp_col = get_col(["home points", "home pts"])  # explicit
        if hp_col:
            try:
                new_cols[cols.index(hp_col)] = "HomeScore"
            except ValueError:
                pass
        elif hs_i is not None and "HomeScore" not in new_cols:
            new_cols[hs_i] = "HomeScore"
        # Away Owner/Score
        ao_i = _nearest_right(owner_indices, away_team_idx)
        as_i = _nearest_right(points_indices, away_team_idx)
        if ao_i is not None:
            new_cols[ao_i] = "AwayOwner"
        # Explicit Away Points if present
        ap_col = get_col(["away points", "away pts"])  # explicit
        if ap_col:
            try:
                new_cols[cols.index(ap_col)] = "AwayScore"
            except ValueError:
                pass
        elif as_i is not None and "AwayScore" not in new_cols:
            new_cols[as_i] = "AwayScore"
        out.columns = new_cols

        # Compute winners/runner-ups (vectorized)
        if {"HomeScore", "AwayScore"}.issubset(out.columns):
            out["HomeScore"] = pd.to_numeric(out["HomeScore"], errors="coerce")
            out["AwayScore"] = pd.to_numeric(out["AwayScore"], errors="coerce")
            home_win = out["HomeScore"] > out["AwayScore"]
            away_win = out["AwayScore"] > out["HomeScore"]
            # Winner/RunnerUp Team
            out["WinnerTeam"] = np.where(home_win, out.get("HomeTeam"), np.where(away_win, out.get("AwayTeam"), pd.NA))
            out["RunnerUpTeam"] = np.where(home_win, out.get("AwayTeam"), np.where(away_win, out.get("HomeTeam"), pd.NA))
            # Winner/RunnerUp Owner
            if "HomeOwner" in out.columns and "AwayOwner" in out.columns:
                out["WinnerOwner"] = np.where(home_win, out.get("HomeOwner"), np.where(away_win, out.get("AwayOwner"), pd.NA))
                out["RunnerUpOwner"] = np.where(home_win, out.get("AwayOwner"), np.where(away_win, out.get("HomeOwner"), pd.NA))
            # Winner/RunnerUp Score
            out["WinnerScore"] = np.where(home_win, out.get("HomeScore"), np.where(away_win, out.get("AwayScore"), pd.NA))
            out["RunnerUpScore"] = np.where(home_win, out.get("AwayScore"), np.where(away_win, out.get("HomeScore"), pd.NA))
        # Derive MatchType from Week text if needed (handles sheets where Week is 'Grand Final'/'Toilet Bowl')
        if "MatchType" not in out.columns:
            out["MatchType"] = None
        if "Week" in out.columns:
            wk = out["Week"].astype(str).str.lower()
            mask_gf = wk.str.contains("grand final|grand", na=False)
            mask_tb = wk.str.contains("toilet bowl|toilet", na=False)
            out.loc[mask_gf, "MatchType"] = "Grand Final"
            out.loc[mask_tb, "MatchType"] = "Toilet Bowl"
            # Default any remaining finals rows to generic 'Final' if they look like finals
            out["MatchType"] = out["MatchType"].fillna("Final")
        out = _strip_strings(out)
        return out

    if sheet_nc == _norm("teams_owners"):
        mapping = {}
        if get_col(["home team"]):
            mapping[get_col(["home team"])] = "TeamName"
        if get_col(["team"]):
            mapping[get_col(["team"])] = "TeamName"
        if get_col(["owner"]):
            mapping[get_col(["owner"])] = "Owner"
        if get_col(["season"]):
            mapping[get_col(["season"])] = "Year"
        out = out.rename(columns=mapping)
        out = _strip_strings(out)
        return out

    if sheet_nc == _norm("reg_season_tables"):
        mapping = {}
        if get_col(["season"]):
            mapping[get_col(["season"]) ] = "Year"
        if get_col(["team"]):
            mapping[get_col(["team"]) ] = "TeamName"
        if get_col(["owner"]):
            mapping[get_col(["owner"]) ] = "Owner"
        if get_col(["w"]):
            mapping[get_col(["w"]) ] = "Wins"
        if get_col(["l"]):
            mapping[get_col(["l"]) ] = "Losses"
        if get_col(["for"]):
            mapping[get_col(["for"]) ] = "PointsFor"
        if get_col(["against"]):
            mapping[get_col(["against"]) ] = "PointsAgainst"
        if get_col(["rank"]):
            mapping[get_col(["rank"]) ] = "Seed"
        out = out.rename(columns=mapping)
        out = _strip_strings(out)
        return out

    if sheet_nc == _norm("draft"):
        mapping = {}
        if get_col(["season"]):
            mapping[get_col(["season"]) ] = "Year"
        for k in ["pick", "owner", "player", "position", "team"]:
            c = get_col([k])
            if c:
                mapping[c] = k.capitalize() if k != "team" else "Team"
        out = out.rename(columns=mapping)
        out = _strip_strings(out)
        return out

    return out


def compute_winner_team(df_gl: pd.DataFrame) -> pd.DataFrame:
    """Compute WinnerTeam column if missing in gamelog."""
    df = df_gl.copy()
    if "WinnerTeam" not in df.columns:
        def _winner(row):
            try:
                if pd.isna(row.get("HomeScore")) or pd.isna(row.get("AwayScore")):
                    return None
                if row["HomeScore"] > row["AwayScore"]:
                    return row.get("HomeTeam")
                if row["AwayScore"] > row["HomeScore"]:
                    return row.get("AwayTeam")
                return "Tie"
            except Exception:
                return None

        df["WinnerTeam"] = df.apply(_winner, axis=1)
    return df


def render_records(
    df_gl: Optional[pd.DataFrame],
    df_reg: Optional[pd.DataFrame],
    selected_years: Optional[List[int]],
    selected_teams: Optional[List[str]],
    selected_owners: Optional[List[str]],
):
    """Display top-10 records for seasons (from regular season) and games (from gamelog)."""
    st.subheader("Season Records (Regular Season)")
    if df_reg is None or df_reg.empty:
        st.info("No regular season data available.")
    else:
        reg = df_reg.copy()
        # Ensure numeric
        for c in ["Year", "Wins", "Losses", "PointsFor", "PointsAgainst"]:
            if c in reg.columns:
                reg[c] = pd.to_numeric(reg[c], errors="coerce")
        # Apply filters
        reg = apply_year_team_owner_filters(reg, years=selected_years, teams=selected_teams, owners=selected_owners)
        if reg is None or reg.empty:
            st.info("No rows after filters.")
        else:
            # Compute win percentage
            games = reg.get("Wins", 0).fillna(0) + reg.get("Losses", 0).fillna(0)
            if "T" in reg.columns:
                try:
                    games = games + pd.to_numeric(reg["T"], errors="coerce").fillna(0)
                except Exception:
                    pass
            reg = reg.assign(GP=games)
            with np.errstate(divide='ignore', invalid='ignore'):
                reg["WinPct"] = np.where(reg["GP"] > 0, reg["Wins"] / reg["GP"], np.nan)

            # Highest win %
            cols = [c for c in ["Year", "TeamName", "Owner", "Wins", "Losses", "GP", "WinPct"] if c in reg.columns]
            top_winpct = reg[cols].dropna(subset=["WinPct"]).sort_values(["WinPct", "Wins"], ascending=[False, False]).head(10)
            # Format
            if not top_winpct.empty:
                df_disp = top_winpct.copy()
                df_disp["Win %"] = (df_disp["WinPct"] * 100).round(1)
                df_disp = df_disp.rename(columns={"TeamName": "Team"})
                st.markdown("**Top 10 Highest Win % (by Season)**")
                st.dataframe(df_disp[[c for c in ["Year", "Team", "Owner", "Wins", "Losses", "GP", "Win %"] if c in df_disp.columns]], use_container_width=True)
            else:
                st.info("No data for Highest Win %.")

            # Most points in a season
            if "PointsFor" in reg.columns:
                top_pf = reg.sort_values(["PointsFor"], ascending=False).head(10)
                df_disp = top_pf.rename(columns={"TeamName": "Team", "PointsFor": "Points"})
                st.markdown("**Top 10 Most Points in a Season**")
                st.dataframe(df_disp[[c for c in ["Year", "Team", "Owner", "Points"] if c in df_disp.columns]], use_container_width=True)
            else:
                st.info("No PointsFor column found.")

            # Most points against in a season
            if "PointsAgainst" in reg.columns:
                top_pa = reg.sort_values(["PointsAgainst"], ascending=False).head(10)
                df_disp = top_pa.rename(columns={"TeamName": "Team", "PointsAgainst": "Points Against"})
                st.markdown("**Top 10 Most Points Against in a Season**")
                st.dataframe(df_disp[[c for c in ["Year", "Team", "Owner", "Points Against"] if c in df_disp.columns]], use_container_width=True)
            else:
                st.info("No PointsAgainst column found.")

            # Best offense/defense by PPG and best differential (per season)
            if {"PointsFor", "PointsAgainst", "GP"}.issubset(reg.columns) and (reg["GP"] > 0).any():
                reg_pp = reg.copy()
                with np.errstate(divide='ignore', invalid='ignore'):
                    reg_pp["PPG F"] = np.where(reg_pp["GP"] > 0, reg_pp["PointsFor"] / reg_pp["GP"], np.nan)
                    reg_pp["PPG A"] = np.where(reg_pp["GP"] > 0, reg_pp["PointsAgainst"] / reg_pp["GP"], np.nan)
                    reg_pp["PPG Diff"] = np.where(reg_pp["GP"] > 0, (reg_pp["PointsFor"] - reg_pp["PointsAgainst"]) / reg_pp["GP"], np.nan)

                # Best offense (highest PPG F)
                best_off = reg_pp.sort_values(["PPG F"], ascending=False).head(10)
                if not best_off.empty:
                    df_off = best_off.rename(columns={"TeamName": "Team"}).copy()
                    df_off["PPG F"] = df_off["PPG F"].round(2)
                    st.markdown("**Top 10 Best Offense (PPG For) — Season**")
                    st.dataframe(df_off[[c for c in ["Year", "Team", "Owner", "GP", "PPG F"] if c in df_off.columns]], use_container_width=True)

                # Best defense (lowest PPG A)
                best_def = reg_pp.sort_values(["PPG A"], ascending=True).head(10)
                if not best_def.empty:
                    df_def = best_def.rename(columns={"TeamName": "Team"}).copy()
                    df_def["PPG A"] = df_def["PPG A"].round(2)
                    st.markdown("**Top 10 Best Defense (Lowest PPG Against) — Season**")
                    st.dataframe(df_def[[c for c in ["Year", "Team", "Owner", "GP", "PPG A"] if c in df_def.columns]], use_container_width=True)

                # Best differential (PPG F − PPG A)
                best_diff = reg_pp.sort_values(["PPG Diff"], ascending=False).head(10)
                if not best_diff.empty:
                    df_diff = best_diff.rename(columns={"TeamName": "Team"}).copy()
                    df_diff["PPG Diff"] = df_diff["PPG Diff"].round(2)
                    st.markdown("**Top 10 Best Points Differential per Game — Season**")
                    st.dataframe(df_diff[[c for c in ["Year", "Team", "Owner", "GP", "PPG Diff"] if c in df_diff.columns]], use_container_width=True)

    st.subheader("Game Records")
    if df_gl is None or df_gl.empty:
        st.info("No game log data available.")
    else:
        gl = df_gl.copy()
        # Normalize numeric
        for c in ["Year", "Week", "HomeScore", "AwayScore"]:
            if c in gl.columns:
                gl[c] = pd.to_numeric(gl[c], errors="coerce")
        # Apply filters
        gl = apply_year_team_owner_filters(gl, years=selected_years, teams=selected_teams, owners=selected_owners)
        if gl is None or gl.empty:
            st.info("No rows after filters.")
            return
        # Combined and margin (absolute for ranking)
        gl = gl.assign(
            Combined=(gl.get("HomeScore").fillna(0) + gl.get("AwayScore").fillna(0)),
            Margin=(gl.get("HomeScore") - gl.get("AwayScore")),
        )
        gl["AbsMargin"] = gl["Margin"].abs()

        # Derive max/min scoring team details for clarity
        home = gl["HomeScore"].astype(float)
        away = gl["AwayScore"].astype(float)
        mask_home_max = home >= away
        mask_home_min = home <= away
        gl["MaxTeamPoints"] = np.where(mask_home_max, home, away)
        gl["MaxTeam"] = np.where(mask_home_max, gl.get("HomeTeam"), gl.get("AwayTeam"))
        gl["MaxOwner"] = np.where(mask_home_max, gl.get("HomeOwner"), gl.get("AwayOwner"))
        gl["MinTeamPoints"] = np.where(mask_home_min, home, away)
        gl["MinTeam"] = np.where(mask_home_min, gl.get("HomeTeam"), gl.get("AwayTeam"))
        gl["MinOwner"] = np.where(mask_home_min, gl.get("HomeOwner"), gl.get("AwayOwner"))

        def _fmt_team_owner(name, owner):
            name = "" if pd.isna(name) else str(name)
            owner = "" if pd.isna(owner) else str(owner)
            if name and owner:
                return f"{name} ({owner})"
            return name or owner

        gl["MaxTeamDisp"] = [
            _fmt_team_owner(n, o) for n, o in zip(gl.get("MaxTeam"), gl.get("MaxOwner"))
        ]
        gl["MinTeamDisp"] = [
            _fmt_team_owner(n, o) for n, o in zip(gl.get("MinTeam"), gl.get("MinOwner"))
        ]
        # Build a compact display row for each game
        def _mk_row(r):
            try:
                home = f"{r.get('HomeTeam','')} ({r.get('HomeOwner','')})"
                away = f"{r.get('AwayTeam','')} ({r.get('AwayOwner','')})"
                score = f"{int(r['HomeScore'])}-{int(r['AwayScore'])}" if pd.notna(r["HomeScore"]) and pd.notna(r["AwayScore"]) else "-"
                return pd.Series({
                    "Year": r.get("Year"),
                    "Week": r.get("Week"),
                    "Home": home,
                    "Away": away,
                    "Score": score,
                    "Combined": r.get("Combined"),
                    "Margin": r.get("AbsMargin"),
                })
            except Exception:
                return pd.Series()

        base_cols = gl.apply(_mk_row, axis=1)

        # Top 10 most points by a single team (with clear team/owner shown)
        top_single = pd.concat([base_cols, gl[["MaxTeamDisp", "MaxTeamPoints"]]], axis=1)
        top_single = top_single.dropna(subset=["MaxTeamPoints"]).sort_values("MaxTeamPoints", ascending=False).head(10)
        if not top_single.empty:
            st.markdown("**Top 10 Most Points by a Team (Game)**")
            ts = top_single.rename(columns={"MaxTeamDisp": "Team (Owner)", "MaxTeamPoints": "Points"})
            st.dataframe(ts[[c for c in ["Team (Owner)", "Points", "Year", "Week", "Home", "Away", "Score"] if c in ts.columns]], use_container_width=True)

        # Top 10 least points by a single team (include combined points)
        low_single = pd.concat([base_cols, gl[["MinTeamDisp", "MinTeamPoints"]]], axis=1)
        low_single = low_single.dropna(subset=["MinTeamPoints"]).sort_values("MinTeamPoints", ascending=True).head(10)
        if not low_single.empty:
            st.markdown("**Top 10 Least Points by a Team (Game)**")
            ls = low_single.rename(columns={"MinTeamDisp": "Team (Owner)", "MinTeamPoints": "Points"})
            st.dataframe(ls[[c for c in ["Team (Owner)", "Points", "Combined", "Year", "Week", "Home", "Away", "Score"] if c in ls.columns]], use_container_width=True)

        # Most combined points (use base_cols which already contains Combined)
        most_comb = base_cols.dropna(subset=["Combined"]).sort_values("Combined", ascending=False).head(10)
        if not most_comb.empty:
            st.markdown("**Top 10 Most Combined Points (Game)**")
            st.dataframe(most_comb[[c for c in ["Year", "Week", "Home", "Away", "Score", "Combined"] if c in most_comb.columns]], use_container_width=True)

        # Least combined points
        least_comb = base_cols.dropna(subset=["Combined"]).sort_values("Combined", ascending=True).head(10)
        if not least_comb.empty:
            st.markdown("**Top 10 Least Combined Points (Game)**")
            st.dataframe(least_comb[[c for c in ["Year", "Week", "Home", "Away", "Score", "Combined"] if c in least_comb.columns]], use_container_width=True)

        # Biggest win margin (exclude ties where Margin == 0)
        biggest = base_cols.dropna(subset=["Margin"]).query("Margin > 0").sort_values("Margin", ascending=False).head(10)
        if not biggest.empty:
            st.markdown("**Top 10 Biggest Win Margins (Game)**")
            st.dataframe(biggest[[c for c in ["Year", "Week", "Home", "Away", "Score", "Margin"] if c in biggest.columns]], use_container_width=True)

        # Narrowest win margin (> 0)
        narrow = base_cols.dropna(subset=["Margin"]).query("Margin > 0").sort_values("Margin", ascending=True).head(10)
        if not narrow.empty:
            st.markdown("**Top 10 Narrowest Win Margins (Game)**")
            st.dataframe(narrow[[c for c in ["Year", "Week", "Home", "Away", "Score", "Margin"] if c in narrow.columns]], use_container_width=True)

        # Highest losing score (team still lost)
        gl["LosingPoints"] = np.where(gl["HomeScore"] > gl["AwayScore"], gl["AwayScore"], np.where(gl["AwayScore"] > gl["HomeScore"], gl["HomeScore"], np.nan))
        gl["LosingTeamDisp"] = np.where(
            gl["HomeScore"] > gl["AwayScore"], gl["AwayTeam"].astype(str) + " (" + gl["AwayOwner"].astype(str) + ")",
            np.where(gl["AwayScore"] > gl["HomeScore"], gl["HomeTeam"].astype(str) + " (" + gl["HomeOwner"].astype(str) + ")", np.nan),
        )
        losing_tbl = pd.concat([base_cols[[c for c in ["Year", "Week", "Home", "Away", "Score"] if c in base_cols.columns]], gl[["LosingTeamDisp", "LosingPoints"]]], axis=1)
        losing_tbl = losing_tbl.dropna(subset=["LosingPoints"]).sort_values("LosingPoints", ascending=False).head(10)
        if not losing_tbl.empty:
            st.markdown("**Top 10 Highest Scoring Losing Teams (Game)**")
            lt = losing_tbl.rename(columns={"LosingTeamDisp": "Team (Owner)", "LosingPoints": "Points"})
            st.dataframe(lt[[c for c in ["Team (Owner)", "Points", "Year", "Week", "Home", "Away", "Score"] if c in lt.columns]], use_container_width=True)

        # Lowest winning score (team still won)
        gl["WinningPoints"] = np.where(gl["HomeScore"] > gl["AwayScore"], gl["HomeScore"], np.where(gl["AwayScore"] > gl["HomeScore"], gl["AwayScore"], np.nan))
        gl["WinningTeamDisp"] = np.where(
            gl["HomeScore"] > gl["AwayScore"], gl["HomeTeam"].astype(str) + " (" + gl["HomeOwner"].astype(str) + ")",
            np.where(gl["AwayScore"] > gl["HomeScore"], gl["AwayTeam"].astype(str) + " (" + gl["AwayOwner"].astype(str) + ")", np.nan),
        )
        winning_tbl = pd.concat([base_cols[[c for c in ["Year", "Week", "Home", "Away", "Score"] if c in base_cols.columns]], gl[["WinningTeamDisp", "WinningPoints"]]], axis=1)
        winning_tbl = winning_tbl.dropna(subset=["WinningPoints"]).sort_values("WinningPoints", ascending=True).head(10)
        if not winning_tbl.empty:
            st.markdown("**Top 10 Lowest Winning Scores (Game)**")
            wt = winning_tbl.rename(columns={"WinningTeamDisp": "Team (Owner)", "WinningPoints": "Points"})
            st.dataframe(wt[[c for c in ["Team (Owner)", "Points", "Year", "Week", "Home", "Away", "Score"] if c in wt.columns]], use_container_width=True)

    # Longest win streak by owner (overall)
        try:
            # Two-row per game: owner + result
            a = pd.DataFrame({
                "Owner": gl.get("HomeOwner"),
                "Year": gl.get("Year"),
                "Week": gl.get("Week"),
                "Win": gl.get("HomeScore") > gl.get("AwayScore"),
            })
            b = pd.DataFrame({
                "Owner": gl.get("AwayOwner"),
                "Year": gl.get("Year"),
                "Week": gl.get("Week"),
                "Win": gl.get("AwayScore") > gl.get("HomeScore"),
            })
            long = pd.concat([a, b], ignore_index=True)
            long["Year"] = pd.to_numeric(long["Year"], errors="coerce")
            long["Week"] = pd.to_numeric(long["Week"], errors="coerce")
            long = long.dropna(subset=["Owner", "Year", "Week"])  # keep valid rows only
            long = long.sort_values(["Owner", "Year", "Week"])  # chronological

            # Compute run-lengths of consecutive wins per owner
            def _streaks(g: pd.DataFrame) -> pd.DataFrame:
                s = g["Win"].astype(bool)
                # Identify segments where Win value changes
                grp = (s != s.shift()).cumsum()
                out = g.copy()
                out["seg"] = grp
                out["is_win_seg"] = s
                return out

            seg = long.groupby("Owner", group_keys=False).apply(_streaks)
            win_segs = seg[seg["is_win_seg"]]
            agg = win_segs.groupby(["Owner", "seg"]).agg(
                Streak=("Win", "size"),
                StartYear=("Year", "first"),
                StartWeek=("Week", "first"),
                EndYear=("Year", "last"),
                EndWeek=("Week", "last"),
            ).reset_index(drop=False)
            top_streaks = agg.sort_values(["Streak"], ascending=False).head(10)
            if not top_streaks.empty:
                st.markdown("**Top 10 Longest Win Streaks (Owner)**")
                st.dataframe(top_streaks[["Owner", "Streak", "StartYear", "StartWeek", "EndYear", "EndWeek"]], use_container_width=True)
        except Exception:
            pass

        # Longest losing streak by owner (overall)
        try:
            # Re-use the long dataframe if available; else recompute quickly
            if 'long' not in locals():
                a = pd.DataFrame({
                    "Owner": gl.get("HomeOwner"),
                    "Year": gl.get("Year"),
                    "Week": gl.get("Week"),
                    "Win": gl.get("HomeScore") > gl.get("AwayScore"),
                })
                b = pd.DataFrame({
                    "Owner": gl.get("AwayOwner"),
                    "Year": gl.get("Year"),
                    "Week": gl.get("Week"),
                    "Win": gl.get("AwayScore") > gl.get("HomeScore"),
                })
                long = pd.concat([a, b], ignore_index=True)
                long["Year"] = pd.to_numeric(long["Year"], errors="coerce")
                long["Week"] = pd.to_numeric(long["Week"], errors="coerce")
                long = long.dropna(subset=["Owner", "Year", "Week"]).sort_values(["Owner", "Year", "Week"])  # chronological

            # Invert wins to mark losing segments
            s = (~long["Win"].astype(bool)).rename("Lose")
            grp = (s != s.shift()).cumsum()
            seg2 = long.copy()
            seg2["seg"] = grp
            seg2["is_lose_seg"] = s
            lose_segs = seg2[seg2["is_lose_seg"]]
            agg2 = lose_segs.groupby(["Owner", "seg"]).agg(
                Streak=("Lose", "size"),
                StartYear=("Year", "first"),
                StartWeek=("Week", "first"),
                EndYear=("Year", "last"),
                EndWeek=("Week", "last"),
            ).reset_index(drop=False)
            top_losing = agg2.sort_values(["Streak"], ascending=False).head(10)
            if not top_losing.empty:
                st.markdown("**Top 10 Longest Losing Streaks (Owner)**")
                st.dataframe(top_losing[["Owner", "Streak", "StartYear", "StartWeek", "EndYear", "EndWeek"]], use_container_width=True)
        except Exception:
            pass


def apply_year_team_owner_filters(
    df: pd.DataFrame,
    years: Optional[List[int]] = None,
    teams: Optional[List[str]] = None,
    owners: Optional[List[str]] = None,
) -> pd.DataFrame:
    """Filter a DataFrame by Year, TeamName/HomeTeam/AwayTeam, and Owner columns if present."""
    if df is None or df.empty:
        return df
    out = df.copy()
    if years:
        if "Year" in out.columns:
            out = out[out["Year"].isin(years)]
    if teams:
        team_cols = [c for c in ["TeamName", "WinnerTeam", "RunnerUpTeam", "HomeTeam", "AwayTeam"] if c in out.columns]
        if team_cols:
            out = out[out[team_cols].apply(lambda r: any(str(v) in set(teams) for v in r.values), axis=1)]
    if owners:
        owner_cols = [
            c
            for c in [
                "Owner",
                "WinnerOwner",
                "RunnerUpOwner",
                "HomeOwner",
                "AwayOwner",
                # Legacy/alternate owner columns seen in raw sheets
                "A Owner",
                "B Owner",
                "Owner A",
                "Owner B",
                "Team A Owner",
                "Team B Owner",
                "Away Owner",
                "Home Owner",
                "Owner.1",
            ]
            if c in out.columns
        ]
        if owner_cols:
            out = out[out[owner_cols].apply(lambda r: any(str(v) in set(owners) for v in r.values), axis=1)]
    return out


def first_round_draft(
    df_draft: pd.DataFrame,
    df_teams_owners: Optional[pd.DataFrame],
    df_reg: Optional[pd.DataFrame] = None,
) -> pd.DataFrame:
    """Return only Round 1 picks.

    Logic:
    - If a 'Round' column exists, filter Round == 1.
    - Else, infer first round as Pick <= number_of_teams_in_year (if mapping available),
      otherwise take the first N picks where N is the minimum pick count across years.
    """
    if df_draft is None or df_draft.empty:
        return df_draft

    df = df_draft.copy()
    if "Round" in df.columns:
        try:
            df_round = df[df["Round"].astype(str).str.strip().isin(["1", "1.0"])].copy()
            return df_round
        except Exception:
            pass

    # Prefer teams per Year from teams_owners; else from reg_season_tables
    teams_per_year = None
    if df_teams_owners is not None and not df_teams_owners.empty and {"Year", "TeamName"}.issubset(df_teams_owners.columns):
        to = df_teams_owners.copy()
        to["Year"] = pd.to_numeric(to["Year"], errors="coerce")
        teams_per_year = to.dropna(subset=["Year"]).groupby("Year")["TeamName"].nunique().to_dict()
    elif df_reg is not None and not df_reg.empty and {"Year", "TeamName"}.issubset(df_reg.columns):
        rg = df_reg.copy()
        rg["Year"] = pd.to_numeric(rg["Year"], errors="coerce")
        teams_per_year = rg.dropna(subset=["Year"]).groupby("Year")["TeamName"].nunique().to_dict()

    if teams_per_year:
        df = df.copy()
        df["_Ynum"] = pd.to_numeric(df["Year"], errors="coerce")
        df["_teams_in_year"] = df["_Ynum"].map(teams_per_year)
        out = df[df["Pick"] <= df["_teams_in_year"].fillna(df["Pick"])].drop(columns=["_Ynum", "_teams_in_year"], errors="ignore")
        return out

    # Fallback: infer per-year first-round size from max Pick seen in that Year
    n_by_year = df.groupby("Year")["Pick"].max(min_count=1)
    try:
        n_by_year = pd.to_numeric(n_by_year, errors="coerce")
    except Exception:
        pass
    n_by_year = n_by_year.dropna()
    if not n_by_year.empty:
        df = df.copy()
        df["_Y"] = df["Year"]
        df["_max_pick_year"] = df["_Y"].map(n_by_year.to_dict())
        return df[df["Pick"] <= df["_max_pick_year"].fillna(df["Pick"])].drop(columns=["_Y", "_max_pick_year"], errors="ignore")
    return df


def get_years(df_list: List[pd.DataFrame]) -> List[int]:
    """Collect sorted unique years across provided DataFrames."""
    years = set()
    for df in df_list:
        if df is not None and not df.empty and "Year" in df.columns:
            years.update([int(y) for y in pd.unique(df["Year"].dropna())])
    return sorted(years)


def compute_overview_metrics(
    df_ch: Optional[pd.DataFrame], df_gl: Optional[pd.DataFrame]
) -> Dict[str, Optional[str]]:
    """Compute high-level overview metrics based on championships and gamelog."""
    metrics = {
        "seasons": None,
        "most_titles_owner": None,
        "most_titles_team": None,
        "highest_score": None,
        "largest_margin": None,
        "current_champion": None,
        "current_toilet_champion": None,
    # Extras for compact owner-only display
    "most_titles_owner_only": None,
    "current_champion_owner_only": None,
    "current_toilet_owner_only": None,
    "current_champion_year": None,
    "current_toilet_year": None,
    }
    try:
        # Seasons from championship sheet if available
        if df_ch is not None and not df_ch.empty:
            if "Year" in df_ch.columns:
                seasons = df_ch["Year"].dropna().unique()
                metrics["seasons"] = f"{len(seasons)}"
            # Compute team titles from championship winners (kept as-is)
            gf = df_ch.copy()
            if "MatchType" in df_ch.columns:
                mt = df_ch["MatchType"].astype(str)
                gf_mask = mt.str.contains("grand|gf", case=False, na=False)
                if not gf_mask.any():
                    gf_mask = mt.str.contains("final", case=False, na=False) & ~mt.str.contains("semi|toilet|tb", case=False, na=False)
                gf = df_ch[gf_mask]
            if gf is None or gf.empty:
                tmp = df_ch.copy()
                tmp["_Y"] = pd.to_numeric(tmp.get("Year"), errors="coerce")
                tmp["_W"] = pd.to_numeric(tmp.get("Week"), errors="coerce")
                if tmp["_Y"].notna().any() and tmp["_W"].notna().any():
                    idx = tmp.dropna(subset=["_Y", "_W"]).sort_values(["_Y", "_W"]).groupby("_Y")["_W"].idxmax()
                    gf = tmp.loc[idx]
            if gf is not None and not gf.empty and "WinnerTeam" in gf.columns:
                by_team = gf.groupby("WinnerTeam")["Year"].count().sort_values(ascending=False)
                if not by_team.empty:
                    metrics["most_titles_team"] = f"{by_team.index[0]} ({by_team.iloc[0]})"

            # Most titles (Owner) from championships Grand Final winners (preferred over gamelog)
            if gf is not None and not gf.empty:
                gfo = gf.copy()
                # Ensure WinnerOwner exists; compute from scores if necessary
                if "WinnerOwner" not in gfo.columns and {"HomeOwner", "AwayOwner", "HomeScore", "AwayScore"}.issubset(gfo.columns):
                    gfo["HomeScore"] = pd.to_numeric(gfo["HomeScore"], errors="coerce")
                    gfo["AwayScore"] = pd.to_numeric(gfo["AwayScore"], errors="coerce")
                    home_win = gfo["HomeScore"] > gfo["AwayScore"]
                    gfo["WinnerOwner"] = np.where(home_win, gfo.get("HomeOwner"), gfo.get("AwayOwner"))
                if "WinnerOwner" in gfo.columns:
                    by_owner_gf = gfo.groupby("WinnerOwner")["Year"].count().sort_values(ascending=False)
                    if not by_owner_gf.empty:
                        top_owner = by_owner_gf.index[0]
                        top_count = int(by_owner_gf.iloc[0])
                        metrics["most_titles_owner"] = f"{top_owner} ({top_count})"
                        metrics["most_titles_owner_only"] = f"{top_owner} ({top_count})"

            # Current Champion and Toilet Bowl Champion from latest season in championships
            latest_y = None
            if "Year" in df_ch.columns:
                years = pd.to_numeric(df_ch["Year"], errors="coerce").dropna()
                if not years.empty:
                    latest_y = int(years.max())
            if latest_y is not None:
                ch_latest = df_ch[pd.to_numeric(df_ch["Year"], errors="coerce") == latest_y].copy()
                # Helper to extract winner text
                def _winner_text(rows: pd.DataFrame) -> Optional[str]:
                    if rows is None or rows.empty:
                        return None
                    r = rows.iloc[0]
                    owner = r.get("WinnerOwner")
                    team = r.get("WinnerTeam")
                    if pd.isna(owner) and {"HomeOwner", "AwayOwner", "HomeScore", "AwayScore"}.issubset(rows.columns):
                        try:
                            rows = rows.copy()
                            rows["HomeScore"] = pd.to_numeric(rows["HomeScore"], errors="coerce")
                            rows["AwayScore"] = pd.to_numeric(rows["AwayScore"], errors="coerce")
                            rw = rows.iloc[0]
                            if rw["HomeScore"] > rw["AwayScore"]:
                                owner = rw.get("HomeOwner")
                                team = rw.get("HomeTeam")
                            elif rw["AwayScore"] > rw["HomeScore"]:
                                owner = rw.get("AwayOwner")
                                team = rw.get("AwayTeam")
                        except Exception:
                            pass
                    if pd.notna(owner) and pd.notna(team):
                        return f"{team} ({owner}), {latest_y}"
                    if pd.notna(owner):
                        return f"{owner}, {latest_y}"
                    if pd.notna(team):
                        return f"{team}, {latest_y}"
                    return None

                # Grand Final
                gf_rows = None
                if "MatchType" in ch_latest.columns:
                    gf_rows = ch_latest[ch_latest["MatchType"].astype(str).str.contains("grand|gf", case=False, na=False)]
                if (gf_rows is None or gf_rows.empty) and "Week" in ch_latest.columns:
                    gf_rows = ch_latest[ch_latest["Week"].astype(str).str.contains("grand", case=False, na=False)]
                if gf_rows is not None and not gf_rows.empty:
                    metrics["current_champion"] = _winner_text(gf_rows)
                    # Extract owner/year for compact view
                    try:
                        r = gf_rows.iloc[0]
                        metrics["current_champion_year"] = int(pd.to_numeric(r.get("Year"), errors="coerce")) if pd.notna(r.get("Year")) else None
                        metrics["current_champion_owner_only"] = r.get("WinnerOwner") or r.get("HomeOwner") or r.get("AwayOwner")
                    except Exception:
                        pass

                # Toilet Bowl
                tb_rows = None
                if "MatchType" in ch_latest.columns:
                    tb_rows = ch_latest[ch_latest["MatchType"].astype(str).str.contains("toilet", case=False, na=False)]
                if (tb_rows is None or tb_rows.empty) and "Week" in ch_latest.columns:
                    tb_rows = ch_latest[ch_latest["Week"].astype(str).str.contains("toilet", case=False, na=False)]
                if tb_rows is not None and not tb_rows.empty:
                    # Define the Toilet Bowl "champion" as the loser
                    try:
                        r = tb_rows.iloc[0]
                        y = int(pd.to_numeric(r.get("Year"), errors="coerce")) if pd.notna(r.get("Year")) else None
                        hs = pd.to_numeric(r.get("HomeScore"), errors="coerce")
                        as_ = pd.to_numeric(r.get("AwayScore"), errors="coerce")
                        home_owner = r.get("HomeOwner")
                        away_owner = r.get("AwayOwner") or r.get("Owner.1")
                        home_team = r.get("HomeTeam")
                        away_team = r.get("AwayTeam")
                        loser_owner = None
                        winner_owner = None
                        loser_team = None
                        winner_team = None
                        if pd.notna(hs) and pd.notna(as_):
                            if hs < as_:
                                loser_owner, winner_owner = home_owner, away_owner
                                loser_team, winner_team = home_team, away_team
                                loser_score, winner_score = hs, as_
                            elif as_ < hs:
                                loser_owner, winner_owner = away_owner, home_owner
                                loser_team, winner_team = away_team, home_team
                                loser_score, winner_score = as_, hs
                            else:
                                # Tie edge case: treat as no champion
                                loser_owner = None
                        # Populate owner-only and year for overview cards
                        if loser_owner is not None:
                            metrics["current_toilet_owner_only"] = loser_owner
                            metrics["current_toilet_year"] = y
                        else:
                            # Fallback to prior behavior (winner), if any
                            metrics["current_toilet_owner_only"] = away_owner or home_owner
                            metrics["current_toilet_year"] = y
                        # Also keep a verbose text if needed elsewhere (loser format)
                        if loser_team is not None and winner_team is not None and pd.notna(loser_score) and pd.notna(winner_score):
                            metrics["current_toilet_champion"] = f"{loser_team} ({int(loser_score)}) lost to {winner_team} ({int(winner_score)})"
                        elif metrics.get("current_toilet_champion") is None:
                            metrics["current_toilet_champion"] = None
                    except Exception:
                        pass

        # Highest score and largest margin from game log
        if df_gl is not None and not df_gl.empty:
            needed = {"Year", "Week", "HomeTeam", "AwayTeam", "HomeScore", "AwayScore"}
            if needed.issubset(df_gl.columns):
                hs = pd.concat([
                    df_gl[["Year", "Week", "HomeTeam", "HomeScore"]].rename(columns={"HomeTeam": "Team", "HomeScore": "Score"}),
                    df_gl[["Year", "Week", "AwayTeam", "AwayScore"]].rename(columns={"AwayTeam": "Team", "AwayScore": "Score"}),
                ], ignore_index=True)
                hs = hs.dropna(subset=["Score"]).sort_values("Score", ascending=False)
                if not hs.empty:
                    top = hs.iloc[0]
                    metrics["highest_score"] = f"{int(top['Score'])} ({top['Team']}, {int(top['Year'])})"
                margins = df_gl.copy()
                margins["Margin"] = (margins["HomeScore"] - margins["AwayScore"]).abs()
                margins = margins.dropna(subset=["Margin"]).sort_values("Margin", ascending=False)
                if not margins.empty:
                    m = margins.iloc[0]
                    metrics["largest_margin"] = f"{int(m['Margin'])} (Week {m['Week']}, {int(m['Year'])})"

                # Most Titles (Owner) fallback from gamelog if not computed from championships
                if metrics.get("most_titles_owner") is None and {"HomeOwner", "AwayOwner"}.issubset(df_gl.columns):
                    wk = df_gl["Week"].astype(str).str.lower()
                    finals = df_gl[wk.str.contains("grand final", na=False)].copy()
                    if not finals.empty:
                        finals["HomeScore"] = pd.to_numeric(finals["HomeScore"], errors="coerce")
                        finals["AwayScore"] = pd.to_numeric(finals["AwayScore"], errors="coerce")
                        home_win = finals["HomeScore"] > finals["AwayScore"]
                        finals["WinnerOwnerGL"] = np.where(home_win, finals.get("HomeOwner"), finals.get("AwayOwner"))
                        by_owner_gl = finals.groupby("WinnerOwnerGL")["Year"].count().sort_values(ascending=False)
                        if not by_owner_gl.empty:
                            metrics["most_titles_owner"] = f"{by_owner_gl.index[0]} ({by_owner_gl.iloc[0]})"
    except Exception:
        pass
    return metrics


def safe_chart(fig: go.Figure, use_container_width: bool = True, caption: Optional[str] = None):
    """Render a Plotly chart with a light template and error safety."""
    try:
        fig.update_layout(template="plotly_white", margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig, use_container_width=use_container_width, theme="streamlit")
        # Auto-caption based on figure title unless an explicit caption is provided
        if caption is None:
            title_text = None
            try:
                title_text = fig.layout.title.text if fig.layout and fig.layout.title else None
            except Exception:
                title_text = None
            if title_text:
                # Direct mapping
                cap = CHART_CAPTIONS.get(str(title_text))
                # Pattern-based fallbacks
                if cap is None and isinstance(title_text, str):
                    if title_text.startswith("Scores Over Time"):
                        cap = "Head-to-head scores by year for the selected matchup."
                if cap:
                    st.caption(cap)
        else:
            st.caption(caption)
    except Exception as e:
        st.info(f"Chart could not be rendered: {e}")


# --------------------- UI Builders ---------------------


def owner_finals_summary(df_ch: Optional[pd.DataFrame], df_gl: Optional[pd.DataFrame]) -> Optional[pd.DataFrame]:
    """Return a summary DataFrame of finals appearances and W/L per owner for Grand Final and Toilet Bowl.

    Prefers championship_games; falls back to gamelog where Week contains 'Grand Final' or 'Toilet Bowl'.
    Output: Owner, Week, Appearances, Wins, Losses
    """
    def _build(df: pd.DataFrame, source: str) -> Optional[pd.DataFrame]:
        if df is None or df.empty:
            return None
        if not {"HomeOwner", "AwayOwner"}.issubset(df.columns):
            return None
        d = df.copy()
        # Derive finals label
        if source == "ch":
            if "MatchType" in d.columns:
                week_label = d["MatchType"].astype(str)
            else:
                wk = d.get("Week").astype(str) if "Week" in d.columns else pd.Series(["" for _ in range(len(d))])
                week_label = pd.Series(np.where(wk.str.contains("grand", case=False, na=False), "Grand Final",
                                      np.where(wk.str.contains("toilet", case=False, na=False), "Toilet Bowl", wk)))
        else:
            wk = d.get("Week").astype(str) if "Week" in d.columns else pd.Series(["" for _ in range(len(d))])
            week_label = pd.Series(np.where(wk.str.contains("grand", case=False, na=False), "Grand Final",
                                  np.where(wk.str.contains("toilet", case=False, na=False), "Toilet Bowl", wk)))
        d["_WeekLabel"] = week_label
        finals = d[d["_WeekLabel"].astype(str).str.contains("grand|toilet", case=False, na=False)].copy()
        if finals.empty:
            return None
        # WinnerOwner
        if "WinnerOwner" not in finals.columns and {"HomeScore", "AwayScore"}.issubset(finals.columns):
            finals["HomeScore"] = pd.to_numeric(finals["HomeScore"], errors="coerce")
            finals["AwayScore"] = pd.to_numeric(finals["AwayScore"], errors="coerce")
            home_win = finals["HomeScore"] > finals["AwayScore"]
            finals["WinnerOwner"] = np.where(home_win, finals.get("HomeOwner"), finals.get("AwayOwner"))
        rows = []
        for _, r in finals.iterrows():
            week = r.get("_WeekLabel")
            ho = r.get("HomeOwner")
            ao = r.get("AwayOwner")
            w = r.get("WinnerOwner")
            if pd.notna(ho):
                rows.append({"Owner": ho, "Week": week, "Win": 1 if w == ho else 0, "Loss": 1 if (pd.notna(w) and w != ho) else 0})
            if pd.notna(ao):
                rows.append({"Owner": ao, "Week": week, "Win": 1 if w == ao else 0, "Loss": 1 if (pd.notna(w) and w != ao) else 0})
        res = pd.DataFrame(rows)
        if res.empty:
            return None
        out = res.groupby(["Owner", "Week"]).agg(Appearances=("Win", "count"), Wins=("Win", "sum"), Losses=("Loss", "sum")).reset_index()
        order_map = {"Grand Final": 0, "Toilet Bowl": 1}
        out["_ord"] = out["Week"].map(order_map).fillna(99)
        return out.sort_values(["_ord", "Owner"]).drop(columns=["_ord"], errors="ignore")

    # Prefer championship sheet, then fall back to gamelog
    summary = _build(df_ch, "ch") if df_ch is not None else None
    if summary is None:
        summary = _build(df_gl, "gl") if df_gl is not None else None
    return summary


def sidebar_filters(
    df_list: List[pd.DataFrame],
    teams_owners: Optional[pd.DataFrame],
) -> Tuple[List[int], List[str], List[str], str]:
    """Build sidebar filters and return (years, teams, owners, file_path)."""
    st.sidebar.markdown("## Filters")

    # Use bundled Excel in the repo when deployed; no manual path needed
    st.sidebar.caption("Data source: bundled Excel file in the app repo.")
    file_path = DEFAULT_FILE_PATH

    years = get_years(df_list)
    # Use a range slider for years; map the range back to the available years list
    if years:
        y_min, y_max = int(min(years)), int(max(years))
        if y_min == y_max:
            chosen = st.sidebar.slider("Year", min_value=y_min, max_value=y_max, value=y_min)
            selected_years = [chosen]
        else:
            yr_lo, yr_hi = st.sidebar.slider("Year range", min_value=y_min, max_value=y_max, value=(y_min, y_max))
            selected_years = [y for y in years if yr_lo <= int(y) <= yr_hi]
    else:
        selected_years = []

    team_options: List[str] = []
    owner_options: List[str] = []
    if teams_owners is not None and not teams_owners.empty:
        if "TeamName" in teams_owners.columns:
            team_options = sorted(teams_owners["TeamName"].dropna().astype(str).unique())
        if "Owner" in teams_owners.columns:
            owner_options = sorted(teams_owners["Owner"].dropna().astype(str).unique())

    selected_teams = st.sidebar.multiselect("Team(s)", options=team_options)
    selected_owners = st.sidebar.multiselect("Owner(s)", options=owner_options)

    return selected_years, selected_teams, selected_owners, file_path


def render_overview(df_ch, df_gl, df_reg, df_to, selected_years, selected_teams, selected_owners):
    """Overview tab with metrics and a couple of league-wide charts."""
    #t.subheader("League Overview")

    df_ch_f = apply_year_team_owner_filters(df_ch, selected_years, selected_teams, selected_owners) if df_ch is not None else None
    df_gl_f = apply_year_team_owner_filters(df_gl, selected_years, selected_teams, selected_owners) if df_gl is not None else None
    df_reg_f = apply_year_team_owner_filters(df_reg, selected_years, selected_teams, selected_owners) if df_reg is not None else None

    # Helper: compute current Elo #1 (Owner-only)
    def _current_elo_leader(gl: Optional[pd.DataFrame]) -> Tuple[Optional[str], Optional[float]]:
        if gl is None or gl.empty:
            return None, None
        d = gl.copy()
        # Ensure score columns
        if not {"HomeScore", "AwayScore"}.issubset(d.columns):
            cols = list(d.columns)
            norm_map = {c: _norm(c) for c in cols}
            def _find(name: str) -> Optional[str]:
                target = _norm(name)
                for c, n in norm_map.items():
                    if n == target:
                        return c
                return None
            a_pts = _find("team a points") or _find("points a")
            b_pts = _find("team b points") or _find("points b")
            if a_pts and b_pts:
                d["HomeScore"] = pd.to_numeric(d[a_pts], errors="coerce")
                d["AwayScore"] = pd.to_numeric(d[b_pts], errors="coerce")
        # Ensure numeric
        for sc in ["HomeScore", "AwayScore"]:
            if sc in d.columns:
                d[sc] = pd.to_numeric(d[sc], errors="coerce")
        d = d.dropna(subset=["HomeScore", "AwayScore"]) if {"HomeScore", "AwayScore"}.issubset(d.columns) else pd.DataFrame()
        if d is None or d.empty:
            return None, None
        # Owner columns
        ent_a_col = "HomeOwner" if "HomeOwner" in d.columns else next((c for c in ["Home Owner", "A Owner", "Owner A", "Team A Owner"] if c in d.columns), None)
        ent_b_col = "AwayOwner" if "AwayOwner" in d.columns else next((c for c in ["Away Owner", "B Owner", "Owner B", "Team B Owner", "Owner.1"] if c in d.columns), None)
        if ent_a_col is None or ent_b_col is None:
            return None, None
        # Sort by Year/Week
        d["_Y"] = pd.to_numeric(d.get("Year"), errors="coerce")
        d["_W"] = pd.to_numeric(d.get("Week"), errors="coerce")
        d = d.sort_values(["_Y", "_W"]).reset_index(drop=True)
        # Elo sim
        ratings: Dict[str, float] = {}
        initial_elo = 1000.0
        k = 32.0
        for _, r in d.iterrows():
            a = str(r.get(ent_a_col)) if pd.notna(r.get(ent_a_col)) else None
            b = str(r.get(ent_b_col)) if pd.notna(r.get(ent_b_col)) else None
            if not a or not b:
                continue
            hs = r.get("HomeScore"); as_ = r.get("AwayScore")
            if pd.isna(hs) or pd.isna(as_):
                continue
            ra = ratings.get(a, initial_elo)
            rb = ratings.get(b, initial_elo)
            if hs > as_:
                res_a = 1.0
            elif as_ > hs:
                res_a = 0.0
            else:
                res_a = 0.5
            exp_a = 1.0 / (1.0 + 10 ** ((rb - ra) / 400.0))
            exp_b = 1.0 - exp_a
            ratings[a] = round(ra + k * (res_a - exp_a), 2)
            ratings[b] = round(rb + k * ((1.0 - res_a) - exp_b), 2)
        if not ratings:
            return None, None
        leader, val = max(ratings.items(), key=lambda kv: kv[1])
        return leader, val

    # Overview metrics
    metrics = compute_overview_metrics(df_ch_f, df_gl_f)
    elo_leader_owner, elo_leader_val = _current_elo_leader(df_gl_f)
    st.markdown('<div class="overview-cards">', unsafe_allow_html=True)
    st.markdown(
        f"""
        <div class="overview-card titles">
            <div class="head"><span class="emoji">👑</span><span>Most Titles (Owner)</span></div>
            <div class="value">{metrics.get('most_titles_owner_only') or metrics.get('most_titles_owner') or '–'}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        f"""
        <div class="overview-card champ">
            <div class="head"><span class="emoji">🏆</span><span>Current Champion</span></div>
            <div class="value">{(metrics.get('current_champion_owner_only') or '–')}</div>
            <div class="sub">{metrics.get('current_champion_year') or ''}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        f"""
        <div class="overview-card toilet">
            <div class="head"><span class="emoji">💩</span><span>Current Toilet Bowl Champion</span></div>
            <div class="value">{(metrics.get('current_toilet_owner_only') or '–')}</div>
            <div class="sub">{metrics.get('current_toilet_year') or ''}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    # New: Current Elo Rating #1 (Owner)
    st.markdown(
        f"""
        <div class="overview-card titles">
            <div class="head"><span class="emoji">🥇</span><span>Current Elo Rating #1</span></div>
            <div class="value">{elo_leader_owner or '–'}</div>
            <div class="sub">{'' if elo_leader_val is None else f'Elo {int(round(elo_leader_val))}'}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown('</div>', unsafe_allow_html=True)

    # Compute latest season based on available filtered data
    def _coerce_int(series: pd.Series) -> pd.Series:
        return pd.to_numeric(series, errors="coerce").astype('Int64') if series is not None else series

    def latest_year(reg: Optional[pd.DataFrame], ch: Optional[pd.DataFrame], gl: Optional[pd.DataFrame]) -> Optional[int]:
        candidates = []
        for df in [reg, ch, gl]:
            if df is not None and not df.empty and "Year" in df.columns:
                years = pd.to_numeric(df["Year"], errors="coerce").dropna()
                if not years.empty:
                    candidates.append(int(years.max()))
        return max(candidates) if candidates else None

    def last_week_for_year(ch: Optional[pd.DataFrame], gl: Optional[pd.DataFrame], year: int) -> Optional[int]:
        week_vals = []
        if ch is not None and not ch.empty and {"Year", "Week"}.issubset(ch.columns):
            w = pd.to_numeric(ch.loc[pd.to_numeric(ch["Year"], errors="coerce") == year, "Week"], errors="coerce").dropna()
            if not w.empty:
                week_vals.append(int(w.max()))
        if gl is not None and not gl.empty and {"Year", "Week"}.issubset(gl.columns):
            w = pd.to_numeric(gl.loc[pd.to_numeric(gl["Year"], errors="coerce") == year, "Week"], errors="coerce").dropna()
            if not w.empty:
                week_vals.append(int(w.max()))
        return max(week_vals) if week_vals else None

    ly = latest_year(df_reg_f, df_ch_f, df_gl_f)
    if ly is not None:
        st.markdown(f"### Latest Season: {ly}")
        c_left, c_right = st.columns(2)

        with c_left:
            st.markdown("#### Regular Season Standings")
            if df_reg_f is not None and not df_reg_f.empty and "Year" in df_reg_f.columns:
                reg_latest = df_reg_f[pd.to_numeric(df_reg_f["Year"], errors="coerce") == ly].copy()
                if not reg_latest.empty:
                    # Numeric conversions
                    for c in ["Seed", "Wins", "Losses", "T", "Pct", "PointsFor", "PointsAgainst"]:
                        if c in reg_latest.columns:
                            reg_latest[c] = pd.to_numeric(reg_latest[c], errors="coerce")
                    # Compute Win % if not present
                    if "Pct" not in reg_latest.columns and {"Wins", "Losses"}.issubset(reg_latest.columns):
                        ties = reg_latest["T"].fillna(0) if "T" in reg_latest.columns else 0
                        games = reg_latest["Wins"].fillna(0) + reg_latest["Losses"].fillna(0) + ties
                        reg_latest["Pct"] = np.where(games > 0, (reg_latest["Wins"].fillna(0) + 0.5 * (ties if isinstance(ties, pd.Series) else 0)) / games, np.nan)
                    # Sort: Seed then Wins
                    if "Seed" in reg_latest.columns:
                        reg_latest = reg_latest.sort_values(["Seed", "Wins" if "Wins" in reg_latest.columns else "TeamName"], ascending=[True, False])
                    elif "Wins" in reg_latest.columns:
                        reg_latest = reg_latest.sort_values(["Wins", "PointsFor" if "PointsFor" in reg_latest.columns else "TeamName"], ascending=[False, False])

                    # Build custom HTML table
                    cols = {
                        "Seed": "Seed",
                        "TeamName": "Team",
                        "Owner": "Owner",
                        "Wins": "W",
                        "Losses": "L",
                        "T": "T",
                        "Pct": "Pct",
                    }
                    available = [src for src in cols.keys() if src in reg_latest.columns]
                    html = [
                        "<div class='standings-wrap'>",
                        "<div class='standings-title'>📊 Standings</div>",
                        "<table class='table-modern'>",
                        "<thead><tr>",
                    ]
                    for k in available:
                        title = cols[k]
                        align = "right" if title in ["W", "L", "T", "Pct"] else "left"
                        html.append(f"<th style='text-align:{align};'>{title}</th>")
                    html.append("</tr></thead><tbody>")

                    def seed_badge(val):
                        if pd.isna(val):
                            return ""
                        v = int(val)
                        cls = "seed-badge"
                        if v == 1:
                            cls += " seed-1"
                        elif v == 2:
                            cls += " seed-2"
                        elif v == 3:
                            cls += " seed-3"
                        return f"<span class='{cls}'>{v}</span>"

                    for _, r in reg_latest.iterrows():
                        html.append("<tr>")
                        for k in available:
                            if k == "Seed":
                                cell = seed_badge(r.get(k))
                                html.append(f"<td>{cell}</td>")
                            elif k == "TeamName":
                                team = str(r.get(k) or "")
                                owner = str(r.get("Owner") or "")
                                cell = (
                                    f"<div class='team-info'><div class='team-name'>{team}</div>"
                                    f"<div class='owner-sub'>{owner}</div></div>"
                                )
                                html.append(f"<td>{cell}</td>")
                            elif k in ["Wins", "Losses", "T"]:
                                val = r.get(k)
                                html.append(
                                    f"<td class='wl' style='text-align:right;'>{'' if pd.isna(val) else int(val)}</td>"
                                )
                            elif k == "Pct":
                                val = r.get(k)
                                pct = f"{val:.3f}" if pd.notna(val) else ""
                                html.append(f"<td class='pct' style='text-align:right;'>{pct}</td>")
                            else:
                                html.append(f"<td>{r.get(k) if pd.notna(r.get(k)) else ''}</td>")
                        html.append("</tr>")
                    html.append("</tbody></table></div>")
                    st.markdown("".join(html), unsafe_allow_html=True)
                else:
                    st.info("No regular season standings found for the latest season.")
            else:
                st.info("Regular season table not available.")

        with c_right:
            st.markdown("#### Last Games")
            def render_match_cards(df_in: pd.DataFrame, from_ch: bool = True) -> bool:
                if df_in is None or df_in.empty:
                    return False
                d = df_in.copy()
                # Try to label GF/TB
                if from_ch and "MatchType" in d.columns:
                    d["_kind"] = d["MatchType"].astype(str).str.lower()
                else:
                    d["_kind"] = d["Week"].astype(str).str.lower() if "Week" in d.columns else ""
                finals = d[d["_kind"].str.contains("grand|toilet", na=False)].copy() if "_kind" in d.columns else pd.DataFrame()
                if finals.empty:
                    return False
                # Build cards
                st.markdown("<div class='match-grid'>", unsafe_allow_html=True)
                for _, r in finals.iterrows():
                    kind = r.get("MatchType") or r.get("Week") or "Final"
                    k = str(kind).lower()
                    cls = "gfinal" if "grand" in k else ("toilet" if "toilet" in k else "")
                    emoji = "🏆" if cls == "gfinal" else ("💩" if cls == "toilet" else "🎯")
                    hs = pd.to_numeric(r.get("HomeScore"), errors="coerce")
                    as_ = pd.to_numeric(r.get("AwayScore"), errors="coerce")
                    # winner/loser classification
                    if pd.notna(hs) and pd.notna(as_):
                        home_win = hs > as_
                        away_win = as_ > hs
                    else:
                        home_win = away_win = False
                    def row_html(team, owner, score, is_win, is_tie=False):
                        state = "win" if is_win else ("tie" if is_tie else "loss")
                        score_txt = "" if pd.isna(score) else (str(int(score)) if float(score).is_integer() else f"{score:.0f}")
                        return f"<div class='team-row {state}'><div class='team-info'><div class='team-name'>{team or ''}</div><div class='team-owner'>{owner or ''}</div></div><div class='team-score'>{score_txt}</div></div>"
                    is_tie = pd.notna(hs) and pd.notna(as_) and hs == as_
                    home_html = row_html(r.get("HomeTeam"), r.get("HomeOwner"), hs, home_win, is_tie)
                    away_html = row_html(r.get("AwayTeam"), r.get("AwayOwner") or r.get("Owner.1"), as_, away_win, is_tie)
                    card = f"""
                    <div class='match-card {cls}'>
                        <div class='match-header'><span class='emoji'>{emoji}</span><span>{kind}</span></div>
                        <div class='match-body'>
                            {home_html}
                            {away_html}
                        </div>
                    </div>
                    """
                    st.markdown(card, unsafe_allow_html=True)
                st.markdown("</div>", unsafe_allow_html=True)
                return True

            shown = False
            # 1) Prefer from championships
            if df_ch_f is not None and not df_ch_f.empty and "Year" in df_ch_f.columns:
                ch_season = df_ch_f[pd.to_numeric(df_ch_f["Year"], errors="coerce") == ly].copy()
                shown = render_match_cards(ch_season, from_ch=True)
            # 2) Fall back to gamelog
            if not shown and df_gl_f is not None and not df_gl_f.empty and "Year" in df_gl_f.columns:
                gl_season = df_gl_f[pd.to_numeric(df_gl_f["Year"], errors="coerce") == ly].copy()
                shown = render_match_cards(gl_season, from_ch=False)
            # 3) Final fallback: show last week as simple dataframe
            if not shown and df_gl_f is not None and not df_gl_f.empty and {"Year", "Week"}.issubset(df_gl_f.columns):
                mask_year = pd.to_numeric(df_gl_f["Year"], errors="coerce") == ly
                gl_season = df_gl_f.loc[mask_year].copy()
                if not gl_season.empty:
                    lw = pd.to_numeric(gl_season["Week"], errors="coerce").max()
                    if pd.notna(lw):
                        gl_last = gl_season[pd.to_numeric(gl_season["Week"], errors="coerce") == lw].copy()
                        cols = [c for c in [
                            "Year", "Week", "HomeTeam", "HomeOwner", "HomeScore", "AwayTeam", "AwayOwner", "AwayScore"
                        ] if c in gl_last.columns]
                        st.dataframe(gl_last[cols], use_container_width=True)
                        shown = True
            if not shown:
                st.info("No games found for the latest season.")

    # Finals appearances and W/L per owner (Grand Final and Toilet Bowl) from championships sheet
    st.markdown("#### Finals appearances and results by owner")
    finals_summary = owner_finals_summary(df_ch_f, df_gl_f)
    if finals_summary is not None and not finals_summary.empty:
        # Build a modern HTML table (no fixed height) with badges for types
        fs = finals_summary.copy()
        fs_cols = [c for c in ["Owner", "Week", "Appearances", "Wins", "Losses"] if c in fs.columns]
        # Map week label to badge class
        def week_badge(w: str) -> str:
            w_str = str(w or "")
            cls = "type-badge gf" if "grand" in w_str.lower() else ("type-badge tb" if "toilet" in w_str.lower() else "type-badge")
            return f"<span class='{cls}'>{w_str}</span>"

        html = [
            "<div class='standings-wrap'>",
            "<div class='standings-title'>🏅 Finals Summary</div>",
            "<table class='table-modern'>",
            "<thead><tr>",
        ]
        for c in fs_cols:
            align = "right" if c in ["Appearances", "Wins", "Losses"] else "left"
            html.append(f"<th style='text-align:{align};'>{c}</th>")
        html.append("</tr></thead><tbody>")
        for _, r in fs.iterrows():
            html.append("<tr>")
            for c in fs_cols:
                if c == "Week":
                    cell = week_badge(r.get(c))
                    html.append(f"<td>{cell}</td>")
                elif c in ["Appearances", "Wins", "Losses"]:
                    val = r.get(c)
                    html.append(f"<td class='wl' style='text-align:right;'>{'' if pd.isna(val) else int(val)}</td>")
                else:
                    html.append(f"<td>{'' if pd.isna(r.get(c)) else r.get(c)}</td>")
            html.append("</tr>")
        html.append("</tbody></table></div>")
        st.markdown("".join(html), unsafe_allow_html=True)
    else:
        st.info("No finals data found in championships sheet to build owner summary.")

    if df_ch_f is not None and not df_ch_f.empty and "WinnerOwner" in df_ch_f.columns:
        gf = df_ch_f
        if "MatchType" in df_ch_f.columns:
            mt = df_ch_f["MatchType"].astype(str)
            gf_mask = mt.str.contains("grand|gf", case=False, na=False)
            if not gf_mask.any():
                gf_mask = mt.str.contains("final", case=False, na=False) & ~mt.str.contains("semi|toilet|tb", case=False, na=False)
            gf = df_ch_f[gf_mask]
            if gf is None or gf.empty and {"Year", "Week"}.issubset(df_ch_f.columns):
                tmp = df_ch_f.copy()
                tmp["_Y"] = pd.to_numeric(tmp.get("Year"), errors="coerce")
                tmp["_W"] = pd.to_numeric(tmp.get("Week"), errors="coerce")
                if tmp["_Y"].notna().any() and tmp["_W"].notna().any():
                    idx = tmp.dropna(subset=["_Y", "_W"]).sort_values(["_Y", "_W"]).groupby("_Y")["_W"].idxmax()
                    gf = tmp.loc[idx].copy()
            # Removed per request: Championships by Owner Over Time chart

    if df_gl_f is not None and not df_gl_f.empty:
        gl = df_gl_f.copy()
        # If the normalized score columns are missing, try to infer from legacy names
        if not {"HomeScore", "AwayScore"}.issubset(gl.columns):
            cols = list(gl.columns)
            norm_map = {c: _norm(c) for c in cols}
            def _find(name: str) -> Optional[str]:
                target = _norm(name)
                for c, n in norm_map.items():
                    if n == target:
                        return c
                return None
            a_pts = _find("team a points") or _find("points a")
            b_pts = _find("team b points") or _find("points b")
            if a_pts and b_pts:
                gl["HomeScore"] = pd.to_numeric(gl[a_pts], errors="coerce")
                gl["AwayScore"] = pd.to_numeric(gl[b_pts], errors="coerce")
        # Owner-split average PPG by year with owners as legend
        if {"HomeScore", "AwayScore", "Year"}.issubset(gl.columns):
            gl["HomeScore"] = pd.to_numeric(gl["HomeScore"], errors="coerce")
            gl["AwayScore"] = pd.to_numeric(gl["AwayScore"], errors="coerce")
            # Build long form with Owner and Score per game side
            owners_home = gl[["Year", "HomeOwner", "HomeScore"]].rename(columns={"HomeOwner": "Owner", "HomeScore": "Score"}) if "HomeOwner" in gl.columns else pd.DataFrame(columns=["Year","Owner","Score"])
            owners_away = gl[["Year", "AwayOwner", "AwayScore"]].rename(columns={"AwayOwner": "Owner", "AwayScore": "Score"}) if "AwayOwner" in gl.columns else pd.DataFrame(columns=["Year","Owner","Score"])
            long = pd.concat([owners_home, owners_away], ignore_index=True)
            long = long.dropna(subset=["Owner", "Score"]).copy()
            if not long.empty:
                long["Year"] = pd.to_numeric(long["Year"], errors="coerce")
                long = long.dropna(subset=["Year"]) 
                by_owner_year = long.groupby(["Owner", "Year"])['Score'].mean().reset_index(name="AvgPoints")
                if not by_owner_year.empty:
                    fig2 = px.line(by_owner_year, x="Year", y="AvgPoints", color="Owner", markers=True, title="Average Points per Game by Year (Owner)")
                    safe_chart(fig2)


def _ensure_data_file(local_path: str) -> str:
    """Ensure the Excel data file exists locally; if missing, try to download.

    Order of sources:
    1) st.secrets["DATA_URL"] (if defined)
    2) env var DATA_URL
    3) GitHub raw URL for the repo file
    Returns the local path (existing or downloaded). May raise on download failure.
    """
    def _looks_like_lfs_pointer(p: str) -> bool:
        try:
            if not os.path.exists(p) or os.path.getsize(p) > 2048:
                return False
            with open(p, 'rb') as f:
                head = f.read(256)
            return b"git-lfs" in head
        except Exception:
            return False

    def _download_to(p: str, url: str) -> bool:
        try:
            urllib.request.urlretrieve(url, p)
            return True
        except Exception as e:
            st.error(f"Failed to fetch data from {url}. {e}")
            return False

    # If file exists, verify it's a valid Excel (zip container) and not an LFS pointer
    if os.path.exists(local_path):
        if _looks_like_lfs_pointer(local_path) or not zipfile.is_zipfile(local_path):
            # Try to replace with a real download
            url = None
            try:
                url = st.secrets.get("DATA_URL")
            except Exception:
                url = None
            url = url or os.environ.get("DATA_URL") or "https://raw.githubusercontent.com/Chockers1/TheBigTebowski/main/fantasy_football.xlsm"
            _download_to(local_path, url)
        return local_path
    # 1) Streamlit secrets
    url = None
    try:
        url = st.secrets.get("DATA_URL")  # type: ignore[attr-defined]
    except Exception:
        url = None
    # 2) Environment variable
    if not url:
        url = os.environ.get("DATA_URL")
    # 3) Fallback to GitHub raw URL (public)
    if not url:
        url = "https://raw.githubusercontent.com/Chockers1/TheBigTebowski/main/fantasy_football.xlsm"
    # Download
    _download_to(local_path, url)
    return local_path


def render_championships(df_ch, selected_years, selected_teams, selected_owners):
    st.subheader("Championships & Toilet Bowl")
    if df_ch is None or df_ch.empty:
        st.info("championship_games sheet not found or empty.")
        return

    df = apply_year_team_owner_filters(df_ch, selected_years, selected_teams, selected_owners)
    
    # Season Winners section at the top
    st.markdown("### 🏆 Season Winners")
    if not df.empty:
        # Use the same processed dataframe that's shown in the table
        # Look for Grand Final in Week column and determine winner by comparing scores
        gf_matches = df[df["Week"].astype(str).str.contains("Grand Final", case=False, na=False)].copy()
        
        if not gf_matches.empty:
            # Try different possible score column names
            home_score_col = None
            away_score_col = None
            
            # Check for the actual column names in the processed dataframe
            for col in ["HomeScore", "Home Pts", "Home Points", "HomePoints"]:
                if col in gf_matches.columns:
                    home_score_col = col
                    break
            
            for col in ["AwayScore", "Away Pts", "Away Points", "AwayPoints", "Points.1"]:
                if col in gf_matches.columns:
                    away_score_col = col
                    break
            
            if home_score_col and away_score_col:
                gf_matches[home_score_col] = pd.to_numeric(gf_matches[home_score_col], errors="coerce")
                gf_matches[away_score_col] = pd.to_numeric(gf_matches[away_score_col], errors="coerce")
                
                # Determine winner by comparing Home Pts vs Away Pts
                def get_winner_info(row):
                    home_score = row.get(home_score_col, 0)
                    away_score = row.get(away_score_col, 0)
                    
                    if pd.notna(home_score) and pd.notna(away_score):
                        if home_score > away_score:
                            return {
                                "team": row.get("HomeTeam", ""),
                                "owner": row.get("HomeOwner", ""),
                                "score": int(home_score),
                                "loser_team": row.get("AwayTeam", "") or row.get("Away Team", ""),
                                "loser_score": int(away_score)
                            }
                        else:
                            # Handle potential duplicate header names
                            away_owner = row.get("AwayOwner", "") or row.get("Owner.1", "") or row.get("Away Owner", "")
                            away_team = row.get("AwayTeam", "") or row.get("Away Team", "")
                            return {
                                "team": away_team,
                                "owner": away_owner,
                                "score": int(away_score),
                                "loser_team": row.get("HomeTeam", ""),
                                "loser_score": int(home_score)
                            }
                    return None
                
                # Sort by year descending
                if "Year" in gf_matches.columns:
                    gf_matches["_YearNum"] = pd.to_numeric(gf_matches["Year"], errors="coerce")
                    gf_matches = gf_matches.dropna(subset=["_YearNum"]).sort_values("_YearNum", ascending=False)
                
                st.markdown('<div class="season-winners">', unsafe_allow_html=True)
                for _, row in gf_matches.iterrows():
                    year = int(row["_YearNum"]) if pd.notna(row.get("_YearNum")) else row.get("Year", "")
                    winner_info = get_winner_info(row)
                    
                    if winner_info:
                        matchup_text = f"beat {winner_info['loser_team']} ({winner_info['loser_score']})"
                        st.markdown(
                            f"""
                            <div class="champion-card">
                                <div class="champion-year">🏆 {year}</div>
                                <div class="champion-team">{winner_info['team']} ({winner_info['score']})</div>
                                <div class="champion-matchup">{matchup_text}</div>
                                <div class="champion-owner">{winner_info['owner']}</div>
                            </div>
                            """,
                            unsafe_allow_html=True
                        )
                st.markdown('</div>', unsafe_allow_html=True)
            else:
                st.info(f"Score columns not found. Available columns: {list(gf_matches.columns)}")
        else:
            st.info("No Grand Final matches found in Week column.")
    
    st.markdown("---")

    # Toilet Bowl Losers section
    st.markdown("### 💩 Toilet Bowl Losers")
    if not df.empty:
        tb_matches = df[df["Week"].astype(str).str.contains("Toilet Bowl", case=False, na=False)].copy()

        if not tb_matches.empty:
            home_score_col = None
            away_score_col = None
            
            for col in ["HomeScore", "Home Pts", "Home Points", "HomePoints"]:
                if col in tb_matches.columns:
                    home_score_col = col
                    break
            
            for col in ["AwayScore", "Away Pts", "Away Points", "AwayPoints", "Points.1"]:
                if col in tb_matches.columns:
                    away_score_col = col
                    break

            if home_score_col and away_score_col:
                tb_matches[home_score_col] = pd.to_numeric(tb_matches[home_score_col], errors="coerce")
                tb_matches[away_score_col] = pd.to_numeric(tb_matches[away_score_col], errors="coerce")

                def get_loser_info(row):
                    home_score = row.get(home_score_col, 0)
                    away_score = row.get(away_score_col, 0)

                    if pd.notna(home_score) and pd.notna(away_score):
                        if home_score < away_score:
                            return {
                                "team": row.get("HomeTeam", ""),
                                "owner": row.get("HomeOwner", ""),
                                "score": int(home_score),
                                "winner_team": row.get("AwayTeam", "") or row.get("Away Team", ""),
                                "winner_score": int(away_score),
                                "winner_owner": row.get("AwayOwner", "") or row.get("Owner.1", "") or row.get("Away Owner", "")
                            }
                        else:
                            return {
                                "team": row.get("AwayTeam", "") or row.get("Away Team", ""),
                                "owner": row.get("AwayOwner", "") or row.get("Owner.1", "") or row.get("Away Owner", ""),
                                "score": int(away_score),
                                "winner_team": row.get("HomeTeam", ""),
                                "winner_score": int(home_score),
                                "winner_owner": row.get("HomeOwner", "")
                            }
                    return None

                if "Year" in tb_matches.columns:
                    tb_matches["_YearNum"] = pd.to_numeric(tb_matches["Year"], errors="coerce")
                    tb_matches = tb_matches.dropna(subset=["_YearNum"]).sort_values("_YearNum", ascending=False)

                st.markdown('<div class="season-winners">', unsafe_allow_html=True)
                for _, row in tb_matches.iterrows():
                    year = int(row["_YearNum"]) if pd.notna(row.get("_YearNum")) else row.get("Year", "")
                    loser_info = get_loser_info(row)

                    if loser_info:
                        matchup_text = f"Lost to {loser_info['winner_team']} ({loser_info['winner_score']})"
                        st.markdown(
                            f"""
                            <div class="loser-card">
                                <div class="champion-year">💩 {year}</div>
                                <div class="champion-team">{loser_info['team']} ({loser_info['score']})</div>
                                <div class="champion-matchup">{matchup_text}</div>
                                <div class="champion-owner">{loser_info['owner']}</div>
                            </div>
                            """,
                            unsafe_allow_html=True
                        )
                st.markdown('</div>', unsafe_allow_html=True)
            else:
                st.info(f"Score columns not found for Toilet Bowl. Available columns: {list(tb_matches.columns)}")
        else:
            st.info("No Toilet Bowl matches found in Week column.")

    # With explicit Home/Away Points present, just compute winner/runner-up and keep going
    dfx = df.copy()
    if {"HomeScore", "AwayScore"}.issubset(dfx.columns):
        for c in ["HomeScore", "AwayScore"]:
            dfx[c] = pd.to_numeric(dfx[c], errors="coerce")
        home_win = dfx["HomeScore"] > dfx["AwayScore"]
        away_win = dfx["AwayScore"] > dfx["HomeScore"]
        if "HomeTeam" in dfx.columns and "AwayTeam" in dfx.columns:
            dfx["WinnerTeam"] = np.where(home_win, dfx.get("HomeTeam"), np.where(away_win, dfx.get("AwayTeam"), pd.NA))
            dfx["RunnerUpTeam"] = np.where(home_win, dfx.get("AwayTeam"), np.where(away_win, dfx.get("HomeTeam"), pd.NA))
        if "HomeOwner" in dfx.columns and "AwayOwner" in dfx.columns:
            dfx["WinnerOwner"] = np.where(home_win, dfx.get("HomeOwner"), np.where(away_win, dfx.get("AwayOwner"), pd.NA))
            dfx["RunnerUpOwner"] = np.where(home_win, dfx.get("AwayOwner"), np.where(away_win, dfx.get("HomeOwner"), pd.NA))
        dfx["WinnerScore"] = np.where(home_win, dfx.get("HomeScore"), np.where(away_win, dfx.get("AwayScore"), pd.NA))
        dfx["RunnerUpScore"] = np.where(home_win, dfx.get("AwayScore"), np.where(away_win, dfx.get("HomeScore"), pd.NA))
    # Derive MatchType from Week if missing
    if "MatchType" not in dfx.columns and "Week" in dfx.columns:
        wk = dfx["Week"].astype(str).str.lower()
        dfx["MatchType"] = None
        dfx.loc[wk.str.contains("grand final|grand", na=False), "MatchType"] = "Grand Final"
        dfx.loc[wk.str.contains("toilet bowl|toilet", na=False), "MatchType"] = "Toilet Bowl"
        dfx["MatchType"] = dfx["MatchType"].fillna("Final")
    df = dfx

    # If raw Home/Away scores are missing or have nulls, derive/fill from winner info (fallback)
    need_scores = ("HomeScore" not in df.columns) or ("AwayScore" not in df.columns)
    winners_available = {
        "HomeTeam",
        "AwayTeam",
        "WinnerTeam",
        "RunnerUpTeam",
        "WinnerScore",
        "RunnerUpScore",
    }.issubset(df.columns)
    if need_scores and winners_available:
        dtemp = df.copy()
        # Home score: winner score if home team won, else runner-up score
        if "HomeScore" not in dtemp.columns:
            dtemp["HomeScore"] = np.where(
                dtemp["HomeTeam"] == dtemp["WinnerTeam"], dtemp["WinnerScore"],
                np.where(dtemp["HomeTeam"] == dtemp["RunnerUpTeam"], dtemp["RunnerUpScore"], pd.NA)
            )
        if "AwayScore" not in dtemp.columns:
            dtemp["AwayScore"] = np.where(
                dtemp["AwayTeam"] == dtemp["WinnerTeam"], dtemp["WinnerScore"],
                np.where(dtemp["AwayTeam"] == dtemp["RunnerUpTeam"], dtemp["RunnerUpScore"], pd.NA)
            )
        df = dtemp
    elif ("HomeScore" in df.columns and "AwayScore" in df.columns) and winners_available:
        # Fill any missing values for Home/Away scores using winner/runner-up
        dtemp = df.copy()
        if dtemp["HomeScore"].isna().any():
            dtemp.loc[dtemp["HomeScore"].isna(), "HomeScore"] = np.where(
                dtemp.loc[dtemp["HomeScore"].isna(), "HomeTeam"] == dtemp.loc[dtemp["HomeScore"].isna(), "WinnerTeam"],
                dtemp.loc[dtemp["HomeScore"].isna(), "WinnerScore"],
                dtemp.loc[dtemp["HomeScore"].isna(), "RunnerUpScore"],
            )
        if dtemp["AwayScore"].isna().any():
            dtemp.loc[dtemp["AwayScore"].isna(), "AwayScore"] = np.where(
                dtemp.loc[dtemp["AwayScore"].isna(), "AwayTeam"] == dtemp.loc[dtemp["AwayScore"].isna(), "WinnerTeam"],
                dtemp.loc[dtemp["AwayScore"].isna(), "WinnerScore"],
                dtemp.loc[dtemp["AwayScore"].isna(), "RunnerUpScore"],
            )
        df = dtemp

    # Order columns if available; no warning needed as we compute fallbacks
    df, _ = ensure_required_columns(df, REQUIRED_SHEETS["championship_games"])

    # Display with user-friendly names only at render-time
    df_display = df.copy()
    if "Year" in df_display.columns:
        df_display["_YearNum"] = pd.to_numeric(df_display["Year"], errors="coerce")
    # Build rename map, handling duplicates gracefully
    rename_map = {
        "Year": "Year",
        "Week": "Week",
        "MatchType": "Type",
        "HomeTeam": "Home Team",
        "HomeOwner": "Home Owner",
        "HomeScore": "Home Pts",
        "AwayTeam": "Away Team",
        "WinnerTeam": "Winner Team",
        "WinnerOwner": "Winner Owner",
        "RunnerUpTeam": "Runner-up Team",
        "RunnerUpOwner": "Runner-up Owner",
        "WinnerScore": "Winner Pts",
        "RunnerUpScore": "Runner-up Pts",
    }
    # Choose best available Away owner/points source for display
    if "AwayOwner" in df_display.columns:
        rename_map["AwayOwner"] = "Away Owner"
    elif "Owner.1" in df_display.columns:
        rename_map["Owner.1"] = "Away Owner"
    if "AwayScore" in df_display.columns:
        rename_map["AwayScore"] = "Away Pts"
    elif "Points.1" in df_display.columns:
        rename_map["Points.1"] = "Away Pts"

    # Apply rename for display
    rename_map = {k: v for k, v in rename_map.items() if k in df_display.columns}
    df_display = df_display.rename(columns=rename_map)

    # Add winner indicators to make it clear who won
    if "Winner Team" in df_display.columns and "Home Team" in df_display.columns and "Away Team" in df_display.columns:
        def add_winner_badge(row):
            winner = row.get("Winner Team")
            home = row.get("Home Team")
            away = row.get("Away Team")
            # Determine if this row is a Toilet Bowl (prefer Type, fallback to Week text)
            type_val = str(row.get("Type")) if row.get("Type") is not None else ""
            week_val = str(row.get("Week")) if row.get("Week") is not None else ""
            is_toilet = ("toilet" in type_val.lower()) or ("toilet" in week_val.lower())
            loser_emoji = "💩" if is_toilet else "💔"

            if pd.notna(winner):
                if winner == home:
                    row["Home Team"] = f'🏆 {home}'
                    row["Away Team"] = f'{loser_emoji} {away}' if pd.notna(away) else away
                elif winner == away:
                    row["Away Team"] = f'🏆 {away}'
                    row["Home Team"] = f'{loser_emoji} {home}' if pd.notna(home) else home
            return row

        df_display = df_display.apply(add_winner_badge, axis=1)

    # Order by readable names
    core_order_disp = [
        "Year", "Week", "Type",
        "Home Team", "Home Owner", "Home Pts",
        "Away Team", "Away Owner", "Away Pts",
        "Winner Team", "Winner Owner", "Runner-up Team", "Runner-up Owner", "Winner Pts", "Runner-up Pts",
    ]
    ordered_cols = [c for c in core_order_disp if c in df_display.columns]
    remaining = [c for c in df_display.columns if c not in ordered_cols + ["_YearNum"]]
    if ordered_cols:
        df_display = df_display[ordered_cols + remaining + (["_YearNum"] if "_YearNum" in df_display.columns else [])]
    if "_YearNum" in df_display.columns:
        df_display = df_display.sort_values("_YearNum", ascending=False).drop(columns=["_YearNum"], errors="ignore")
    st.dataframe(df_display, use_container_width=True)

    # Quick highlights as metric cards (Grand Final and Toilet Bowl)
    def _highlights(final_phrase: str):
        if "MatchType" in df.columns:
            mt_mask = df["MatchType"].astype(str).str.lower().str.contains(final_phrase, na=False)
        else:
            mt_mask = pd.Series(False, index=df.index)
        if "Week" in df.columns:
            wk_mask = df["Week"].astype(str).str.lower().str.contains(final_phrase, na=False)
        else:
            wk_mask = pd.Series(False, index=df.index)
        sub = df[mt_mask | wk_mask].copy()
        if sub.empty or not {"HomeScore", "AwayScore"}.issubset(sub.columns):
            return None
        # Ensure numeric
        sub["HomeScore"] = pd.to_numeric(sub["HomeScore"], errors="coerce")
        sub["AwayScore"] = pd.to_numeric(sub["AwayScore"], errors="coerce")
        # Highest score (report owner not team)
        sub["_MaxSideIsHome"] = sub["HomeScore"] >= sub["AwayScore"]
        sub["_MaxScore"] = sub[["HomeScore", "AwayScore"]].max(axis=1)
        idx_max = sub["_MaxScore"].idxmax() if not sub["_MaxScore"].isna().all() else None
        high = None
        if idx_max is not None and pd.notna(idx_max):
            r = sub.loc[idx_max]
            if bool(r["_MaxSideIsHome"]):
                high = {
                    "score": int(r["HomeScore"]) if pd.notna(r["HomeScore"]) else None,
                    "owner": r.get("HomeOwner"),
                    "year": r.get("Year"),
                }
            else:
                high = {
                    "score": int(r["AwayScore"]) if pd.notna(r["AwayScore"]) else None,
                    "owner": r.get("AwayOwner") or r.get("Owner.1"),
                    "year": r.get("Year"),
                }
        # Biggest margin (report owners not teams)
        sub["_Margin"] = (sub["HomeScore"] - sub["AwayScore"]).abs()
        idx_m = sub["_Margin"].idxmax() if not sub["_Margin"].isna().all() else None
        margin = None
        if idx_m is not None and pd.notna(idx_m):
            r = sub.loc[idx_m]
            if r["HomeScore"] >= r["AwayScore"]:
                margin = {
                    "margin": int(r["HomeScore"] - r["AwayScore"]) if pd.notna(r["HomeScore"]) and pd.notna(r["AwayScore"]) else None,
                    "winner_owner": r.get("HomeOwner"),
                    "loser_owner": r.get("AwayOwner") or r.get("Owner.1"),
                    "year": r.get("Year"),
                }
            else:
                margin = {
                    "margin": int(r["AwayScore"] - r["HomeScore"]) if pd.notna(r["HomeScore"]) and pd.notna(r["AwayScore"]) else None,
                    "winner_owner": r.get("AwayOwner") or r.get("Owner.1"),
                    "loser_owner": r.get("HomeOwner"),
                    "year": r.get("Year"),
                }
        return high, margin

    c_h1, c_h2 = st.columns(2)
    with c_h1:
        hi = _highlights("grand final")
        if hi:
            high, margin = hi
            st.markdown("#### Grand Final highlights")
            if high and high.get("score") is not None:
                st.metric("Highest Team Score", high["score"], f"{high.get('owner')}, {high.get('year')}")
            if margin and margin.get("margin") is not None:
                st.metric("Biggest Winning Margin", margin["margin"], f"{margin.get('winner_owner')} over {margin.get('loser_owner')}, {margin.get('year')}")
    with c_h2:
        hi = _highlights("toilet bowl")
        if hi:
            high, margin = hi
            st.markdown("#### Toilet Bowl highlights")
            if high and high.get("score") is not None:
                st.metric("Highest Team Score", high["score"], f"{high.get('owner')}, {high.get('year')}")
            if margin and margin.get("margin") is not None:
                st.metric("Biggest Winning Margin", margin["margin"], f"{margin.get('winner_owner')} over {margin.get('loser_owner')}, {margin.get('year')}")

    if "MatchType" in df.columns:
        gf = df[df["MatchType"].astype(str).str.contains("grand", case=False, na=False)]
        c1, c2 = st.columns(2)
        with c1:
            if not gf.empty and "WinnerOwner" in gf.columns:
                by_owner = gf.groupby("WinnerOwner")["Year"].count().reset_index(name="Titles")
                fig = px.bar(by_owner, x="WinnerOwner", y="Titles", title="Titles by Owner (Grand Final)")
                fig.update_xaxes(title="Owner")
                safe_chart(fig)
            else:
                st.info("No Grand Final data available for Titles by Owner.")
        with c2:
            # Toilet Bowl losers by owner (the 'champion' is the loser)
            tb = df[df["MatchType"].astype(str).str.contains("toilet", case=False, na=False)].copy()
            if not tb.empty:
                # Prefer RunnerUpOwner if present
                if "RunnerUpOwner" in tb.columns:
                    by_loser = (
                        tb.groupby("RunnerUpOwner").size().reset_index(name="Losses")
                    )
                    fig_tb = px.bar(by_loser, x="RunnerUpOwner", y="Losses", title="Toilet Bowl Losers by Owner")
                    fig_tb.update_xaxes(title="Owner")
                    safe_chart(fig_tb)
                elif {"HomeOwner", "AwayOwner", "HomeScore", "AwayScore"}.issubset(tb.columns):
                    tb["HomeScore"] = pd.to_numeric(tb["HomeScore"], errors="coerce")
                    tb["AwayScore"] = pd.to_numeric(tb["AwayScore"], errors="coerce")
                    loser_is_home = tb["HomeScore"] < tb["AwayScore"]
                    away_owner_col = "AwayOwner" if "AwayOwner" in tb.columns else ("Owner.1" if "Owner.1" in tb.columns else None)
                    if away_owner_col is not None:
                        tb["LoserOwner"] = np.where(loser_is_home, tb.get("HomeOwner"), tb.get(away_owner_col))
                        by_loser = tb.groupby("LoserOwner").size().reset_index(name="Losses")
                        fig_tb = px.bar(by_loser, x="LoserOwner", y="Losses", title="Toilet Bowl Losers by Owner")
                        fig_tb.update_xaxes(title="Owner")
                        safe_chart(fig_tb)
                    else:
                        st.info("Could not find Away owner column to compute Toilet Bowl losers.")
            else:
                st.info("No Toilet Bowl data available.")
        # Toilet Bowl chart removed per user request

    # Finals Score Distribution chart removed per user request

    # Records section: Grand Final and Toilet Bowl
    def _pick(col_opts):
        for c in col_opts:
            if c in df.columns:
                return c
        return None

    def _records_for(final_type: str):
        # Build mask from MatchType and Week text to be resilient
        mt_mask = (
            df.get("MatchType", "").astype(str).str.lower().str.contains(final_type, na=False)
            if "MatchType" in df.columns else pd.Series(False, index=df.index)
        )
        if final_type.startswith("grand"):
            phrase = "grand final"
        else:
            phrase = "toilet bowl"
        wk_mask = (
            df["Week"].astype(str).str.lower().str.contains(phrase, na=False)
            if "Week" in df.columns else pd.Series(False, index=df.index)
        )
        sub = df[mt_mask | wk_mask].copy()
        if sub.empty:
            return None
        # Column picks robust to duplicate headers
        ht, ho, hs = "HomeTeam", "HomeOwner", "HomeScore"
        at = "AwayTeam"
        ao = _pick(["AwayOwner", "Owner.1"]) or "AwayOwner"
        ascore = _pick(["AwayScore", "Points.1"]) or "AwayScore"
        # Ensure numeric
        for c in [hs, ascore]:
            if c in sub.columns:
                sub[c] = pd.to_numeric(sub[c], errors="coerce")
        # Long form for single-team records
        cols_ok = all(c in sub.columns for c in [ht, ho, hs, at, ao, ascore])
        if not cols_ok:
            return None
        home_long = sub[["Year", ht, ho, hs, at, ao, ascore]].copy()
        home_long.columns = ["Year", "Team", "Owner", "Score", "Opponent", "OppOwner", "OppScore"]
        away_long = sub[["Year", at, ao, ascore, ht, ho, hs]].copy()
        away_long.columns = ["Year", "Team", "Owner", "Score", "Opponent", "OppOwner", "OppScore"]
        long_df = pd.concat([home_long, away_long], ignore_index=True)
        long_df["Score"] = pd.to_numeric(long_df["Score"], errors="coerce")
        long_df["OppScore"] = pd.to_numeric(long_df["OppScore"], errors="coerce")

        # Highest team score
        high_idx = long_df["Score"].idxmax()
        rec_high = long_df.loc[[high_idx]].assign(Record="Highest Team Score")[
            ["Record", "Year", "Team", "Owner", "Score", "Opponent", "OppScore"]
        ] if pd.notna(high_idx) else None

        # Biggest win (by margin)
        sub["Margin"] = (sub[hs] - sub[ascore]).abs()
        m_idx = sub["Margin"].idxmax()
        if pd.notna(m_idx):
            row = sub.loc[m_idx]
            # Determine winner context
            if row[hs] > row[ascore]:
                w_team, w_owner, w_score = row[ht], row[ho], row[hs]
                l_team, l_owner, l_score = row[at], row[ao], row[ascore]
            else:
                w_team, w_owner, w_score = row[at], row[ao], row[ascore]
                l_team, l_owner, l_score = row[ht], row[ho], row[hs]
            rec_margin = pd.DataFrame([
                {
                    "Record": "Biggest Winning Margin",
                    "Year": row["Year"],
                    "Team": w_team,
                    "Owner": w_owner,
                    "Score": w_score,
                    "Opponent": l_team,
                    "OppOwner": l_owner,
                    "OppScore": l_score,
                    "Margin": abs(w_score - l_score),
                }
            ])
        else:
            rec_margin = None

        # Highest combined score
        sub["Total"] = pd.to_numeric(sub[hs], errors="coerce") + pd.to_numeric(sub[ascore], errors="coerce")
        t_idx = sub["Total"].idxmax()
        rec_total = None
        if pd.notna(t_idx):
            row = sub.loc[t_idx]
            rec_total = pd.DataFrame([
                {
                    "Record": "Highest Combined Score",
                    "Year": row["Year"],
                    "HomeTeam": row[ht],
                    "AwayTeam": row[at],
                    "Total": row["Total"],
                }
            ])

        # Lowest winning score
        if "WinnerScore" in sub.columns:
            lw_idx = sub["WinnerScore"].idxmin()
            rec_lowwin = None
            if pd.notna(lw_idx):
                row = sub.loc[lw_idx]
                rec_lowwin = pd.DataFrame([
                    {
                        "Record": "Lowest Winning Score",
                        "Year": row["Year"],
                        "WinnerTeam": row.get("WinnerTeam"),
                        "WinnerOwner": row.get("WinnerOwner"),
                        "WinnerScore": row.get("WinnerScore"),
                    }
                ])
        else:
            rec_lowwin = None

        # Most wins by owner (top 5)
        most_wins = None
        if "WinnerOwner" in sub.columns:
            most_wins = (
                sub.groupby("WinnerOwner")["Year"].count().reset_index(name="Wins").sort_values(["Wins", "WinnerOwner"], ascending=[False, True]).head(5)
            )

        # Most appearances by owner (top 5)
        most_apps = None
        if all(c in sub.columns for c in ["WinnerOwner", "RunnerUpOwner"]):
            owners = pd.concat([sub[["WinnerOwner"]].rename(columns={"WinnerOwner": "Owner"}), sub[["RunnerUpOwner"]].rename(columns={"RunnerUpOwner": "Owner"})])
            most_apps = owners.groupby("Owner").size().reset_index(name="Appearances").sort_values(["Appearances", "Owner"], ascending=[False, True]).head(5)

        return {
            "high": rec_high,
            "margin": rec_margin,
            "total": rec_total,
            "lowwin": rec_lowwin,
            "wins": most_wins,
            "apps": most_apps,
        }

    st.markdown("### Records")
    c1, c2 = st.columns(2)

    def _render_record_cards(title: str, rec: Optional[dict]):
        st.markdown(f"#### {title}")
        if not rec:
            st.info(f"No {title.lower()} records available for the selected filters.")
            return
        # Pull simple values from dataframes for cards
        high_txt = None
        if isinstance(rec.get("high"), pd.DataFrame) and not rec["high"].empty:
            r = rec["high"].iloc[0]
            high_txt = {
                "stat": int(r.get("Score")) if pd.notna(r.get("Score")) else None,
                "sub": f"{r.get('Team')} ({r.get('Owner')}), {int(r.get('Year')) if pd.notna(r.get('Year')) else ''}",
            }
        margin_txt = None
        if isinstance(rec.get("margin"), pd.DataFrame) and not rec["margin"].empty:
            r = rec["margin"].iloc[0]
            margin_txt = {
                "stat": int(r.get("Margin")) if pd.notna(r.get("Margin")) else None,
                "sub": f"{r.get('Team')} over {r.get('Opponent')}, {int(r.get('Year')) if pd.notna(r.get('Year')) else ''}",
            }
        total_txt = None
        if isinstance(rec.get("total"), pd.DataFrame) and not rec["total"].empty:
            r = rec["total"].iloc[0]
            total_txt = {
                "stat": int(r.get("Total")) if pd.notna(r.get("Total")) else None,
                "sub": f"{r.get('HomeTeam')} vs {r.get('AwayTeam')}, {int(r.get('Year')) if pd.notna(r.get('Year')) else ''}",
            }
        lowwin_txt = None
        if isinstance(rec.get("lowwin"), pd.DataFrame) and not rec["lowwin"].empty:
            r = rec["lowwin"].iloc[0]
            lowwin_txt = {
                "stat": int(r.get("WinnerScore")) if pd.notna(r.get("WinnerScore")) else None,
                "sub": f"{r.get('WinnerTeam')} ({r.get('WinnerOwner')}), {int(r.get('Year')) if pd.notna(r.get('Year')) else ''}",
            }

        st.markdown('<div class="record-grid">', unsafe_allow_html=True)
        for label, data in [
            ("Highest Team Score", high_txt),
            ("Biggest Winning Margin", margin_txt),
            ("Highest Combined Score", total_txt),
            ("Lowest Winning Score", lowwin_txt),
        ]:
            if data and data.get("stat") is not None:
                st.markdown(
                    f"""
                    <div class='record-card'>
                        <div class='record-title'><span class='pill'>{label}</span></div>
                        <div class='record-stat'>{data['stat']}</div>
                        <div class='record-sub'>{data['sub']}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
        st.markdown('</div>', unsafe_allow_html=True)

    # Removed per request: bottom row charts for top wins/appearances by Owner

    # Detailed tables removed per user request

    with c1:
        gf_rec = _records_for("grand")
        _render_record_cards("Grand Final", gf_rec)
    with c2:
        tb_rec = _records_for("toilet")
        _render_record_cards("Toilet Bowl", tb_rec)


def render_regular_season(df_reg, selected_years, selected_teams, selected_owners):
    st.subheader("Regular Season")
    if df_reg is None or df_reg.empty:
        st.info("reg_season_tables sheet not found or empty.")
        return

    df = apply_year_team_owner_filters(df_reg, selected_years, selected_teams, selected_owners)
    df, missing = ensure_required_columns(df, REQUIRED_SHEETS["reg_season_tables"])
    if missing:
        st.warning(f"Missing columns in reg_season_tables: {missing}")

    st.dataframe(df, use_container_width=True)

    # New: Yearly trend charts (Best Seed and Win%)
    try:
        cys1, cys2 = st.columns(2)
        # Prepare a clean copy with numeric conversions
        dyy = df.copy()
        if "Year" in dyy.columns:
            dyy["Year"] = pd.to_numeric(dyy["Year"], errors="coerce")
        for c in ["Wins", "Losses", "Seed", "Pct"]:
            if c in dyy.columns:
                dyy[c] = pd.to_numeric(dyy[c], errors="coerce")
        ties_col = None
        if "T" in dyy.columns:
            ties_col = "T"
        elif "Ties" in dyy.columns:
            ties_col = "Ties"
        if ties_col:
            dyy["Ties"] = pd.to_numeric(dyy[ties_col], errors="coerce")
        else:
            dyy["Ties"] = 0

        # Left: Best Seed by Year per Owner (lower is better)
        with cys1:
            if {"Owner", "Year", "Seed"}.issubset(dyy.columns):
                best_seed = (
                    dyy.dropna(subset=["Year", "Owner", "Seed"])
                    .groupby(["Owner", "Year"]) ["Seed"].min()
                    .reset_index()
                    .sort_values(["Owner", "Year"])
                )
                if not best_seed.empty:
                    fig_seed = px.line(best_seed, x="Year", y="Seed", color="Owner", markers=True,
                                       title="Best Seed by Year (Owner)")
                    fig_seed.update_yaxes(autorange="reversed", title_text="Best Seed")
                    safe_chart(fig_seed)
                else:
                    st.info("No seed data available for yearly trend.")
            else:
                st.info("Seed column not available to plot Best Seed by Year.")

        # Right: Win Percentage by Year per Owner
        with cys2:
            if {"Owner", "Year"}.issubset(dyy.columns) and ( {"Wins", "Losses"}.issubset(dyy.columns) or "Pct" in dyy.columns ):
                per_year = dyy.dropna(subset=["Year"]).copy()
                if {"Wins", "Losses"}.issubset(per_year.columns):
                    per_year["Wins"] = pd.to_numeric(per_year["Wins"], errors="coerce").fillna(0)
                    per_year["Losses"] = pd.to_numeric(per_year["Losses"], errors="coerce").fillna(0)
                    per_year["Ties"] = pd.to_numeric(per_year["Ties"], errors="coerce").fillna(0)
                    agg = per_year.groupby(["Owner", "Year"]).agg(Wins=("Wins", "sum"), Losses=("Losses", "sum"), Ties=("Ties", "sum")).reset_index()
                    agg["Games"] = agg["Wins"] + agg["Losses"] + agg["Ties"]
                    agg["WinPct"] = np.where(agg["Games"] > 0, (agg["Wins"] + 0.5 * agg["Ties"]) / agg["Games"], np.nan)
                else:
                    # Use existing Pct if Wins/Losses missing; average across entries per owner-year
                    agg = per_year.groupby(["Owner", "Year"]) ["Pct"].mean().reset_index().rename(columns={"Pct": "WinPct"})
                if not agg.empty:
                    fig_wp = px.line(agg, x="Year", y="WinPct", color="Owner", markers=True,
                                     title="Win Percentage by Year (Owner)", labels={"WinPct": "Win %"})
                    safe_chart(fig_wp)
                else:
                    st.info("No win percentage data available for yearly trend.")
            else:
                st.info("Not enough data to plot Win Percentage by Year.")
    except Exception:
        pass

    if not df.empty:
        # All-time regular-season record grouped by Owner
        def _owner_all_time(dfin: pd.DataFrame) -> Optional[pd.DataFrame]:
            if dfin is None or dfin.empty or "Owner" not in dfin.columns:
                return None
            d = dfin.copy()
            for c in ["Wins", "Losses", "PointsFor", "PointsAgainst", "Seed"]:
                if c in d.columns:
                    d[c] = pd.to_numeric(d[c], errors="coerce")
            # Optional ties column can be 'T' or 'Ties'
            ties_col = None
            if "T" in d.columns:
                ties_col = "T"
            elif "Ties" in d.columns:
                ties_col = "Ties"
            if ties_col:
                d[ties_col] = pd.to_numeric(d[ties_col], errors="coerce").fillna(0)
            agg_dict = {
                "Wins": ("Wins", "sum") if "Wins" in d.columns else ("Owner", "size"),
                "Losses": ("Losses", "sum") if "Losses" in d.columns else ("Owner", "size"),
            }
            if "PointsFor" in d.columns:
                agg_dict["PointsFor"] = ("PointsFor", "sum")
            if "PointsAgainst" in d.columns:
                agg_dict["PointsAgainst"] = ("PointsAgainst", "sum")
            if ties_col:
                agg_dict["Ties"] = (ties_col, "sum")
            if "Year" in d.columns:
                seasons_series = d.groupby("Owner")["Year"].nunique()
            else:
                seasons_series = d.groupby("Owner").size()
            grouped = d.groupby("Owner").agg(**agg_dict).reset_index()
            grouped = grouped.merge(seasons_series.rename("Seasons").reset_index(), on="Owner", how="left")
            # Compute derived metrics
            grouped["Wins"] = pd.to_numeric(grouped.get("Wins"), errors="coerce").fillna(0).astype(int)
            grouped["Losses"] = pd.to_numeric(grouped.get("Losses"), errors="coerce").fillna(0).astype(int)
            if "Ties" in grouped.columns:
                grouped["Ties"] = pd.to_numeric(grouped["Ties"], errors="coerce").fillna(0).astype(int)
            else:
                grouped["Ties"] = 0
            grouped["Games"] = grouped["Wins"] + grouped["Losses"] + grouped["Ties"]
            grouped["WinPct"] = np.where(grouped["Games"] > 0, (grouped["Wins"] + 0.5 * grouped["Ties"]) / grouped["Games"], np.nan)
            grouped["AvgWinsPerSeason"] = np.where(grouped["Seasons"] > 0, grouped["Wins"] / grouped["Seasons"], np.nan)
            # Points per game (for and against)
            if "PointsFor" in grouped.columns:
                grouped["PPG F"] = np.where(grouped["Games"] > 0, grouped["PointsFor"] / grouped["Games"], np.nan)
            if "PointsAgainst" in grouped.columns:
                grouped["PPG A"] = np.where(grouped["Games"] > 0, grouped["PointsAgainst"] / grouped["Games"], np.nan)
            # Best Seed is minimum Seed across seasons when available
            if "Seed" in d.columns:
                best_seed = d.groupby("Owner")["Seed"].min().rename("BestSeed").reset_index()
                grouped = grouped.merge(best_seed, on="Owner", how="left")
            # Order columns
            cols = [
                "Owner", "Seasons", "Games", "Wins", "Losses", "Ties", "WinPct", "AvgWinsPerSeason",
            ]
            if "PointsFor" in grouped.columns:
                cols.append("PointsFor")
            if "PointsAgainst" in grouped.columns:
                cols.append("PointsAgainst")
            if "PPG F" in grouped.columns:
                cols.append("PPG F")
            if "PPG A" in grouped.columns:
                cols.append("PPG A")
            if "BestSeed" in grouped.columns:
                cols.append("BestSeed")
            grouped = grouped[cols].sort_values(["Wins", "WinPct"], ascending=[False, False])
            # Formatting
            grouped["WinPct"] = grouped["WinPct"].round(3)
            grouped["AvgWinsPerSeason"] = grouped["AvgWinsPerSeason"].round(2)
            if "PPG F" in grouped.columns:
                grouped["PPG F"] = grouped["PPG F"].round(2)
            if "PPG A" in grouped.columns:
                grouped["PPG A"] = grouped["PPG A"].round(2)
            return grouped

        st.markdown("### All-time regular-season record by Owner")
        owner_all = _owner_all_time(df)
        if owner_all is not None and not owner_all.empty:
            st.dataframe(owner_all, use_container_width=True)
            # Owner charts
            cch1, cch2 = st.columns(2)
            with cch1:
                fig_wins = px.bar(owner_all, x="Owner", y="Wins", color="WinPct", title="All-time Wins by Owner",
                                   labels={"WinPct": "Win %"})
                safe_chart(fig_wins)
            with cch2:
                if "PPG F" in owner_all.columns:
                    fig_eff = px.scatter(owner_all, x="Games", y="WinPct", size="Wins", color="PPG F",
                                         hover_name="Owner", title="Efficiency vs Volume (Games vs Win %)",
                                         labels={"WinPct": "Win %", "PPG F": "PPG F"})
                else:
                    fig_eff = px.scatter(owner_all, x="Games", y="WinPct", size="Wins",
                                         hover_name="Owner", title="Efficiency vs Volume (Games vs Win %)",
                                         labels={"WinPct": "Win %"})
                safe_chart(fig_eff)

            if {"PPG F", "PPG A"}.issubset(owner_all.columns):
                owner_all = owner_all.copy()
                owner_all["Net PPG"] = owner_all["PPG F"] - owner_all["PPG A"]
                fig_net = px.bar(owner_all, x="Owner", y="Net PPG", title="Net Points per Game (PPG F − PPG A)")
                safe_chart(fig_net)

            # Cumulative Win% by Year per Owner
            try:
                df_year = df.copy()
                needed = {"Owner", "Year", "Wins", "Losses"}
                if needed.issubset(df_year.columns):
                    df_year["Year"] = pd.to_numeric(df_year["Year"], errors="coerce")
                    df_year["Wins"] = pd.to_numeric(df_year["Wins"], errors="coerce").fillna(0)
                    df_year["Losses"] = pd.to_numeric(df_year["Losses"], errors="coerce").fillna(0)
                    ties_col = None
                    if "T" in df_year.columns:
                        ties_col = "T"
                    elif "Ties" in df_year.columns:
                        ties_col = "Ties"
                    if ties_col:
                        df_year["Ties"] = pd.to_numeric(df_year[ties_col], errors="coerce").fillna(0)
                    else:
                        df_year["Ties"] = 0
                    by_year = df_year.dropna(subset=["Year"]).groupby(["Owner", "Year"]).agg(
                        Wins=("Wins", "sum"), Losses=("Losses", "sum"), Ties=("Ties", "sum")
                    ).reset_index()
                    by_year = by_year.sort_values(["Owner", "Year"]) 
                    by_year["CumWins"] = by_year.groupby("Owner")["Wins"].cumsum()
                    by_year["CumLosses"] = by_year.groupby("Owner")["Losses"].cumsum()
                    by_year["CumTies"] = by_year.groupby("Owner")["Ties"].cumsum()
                    by_year["CumGames"] = by_year["CumWins"] + by_year["CumLosses"] + by_year["CumTies"]
                    by_year["CumWinPct"] = np.where(
                        by_year["CumGames"] > 0,
                        (by_year["CumWins"] + 0.5 * by_year["CumTies"]) / by_year["CumGames"],
                        np.nan,
                    )
                    fig_cum = px.line(by_year, x="Year", y="CumWinPct", color="Owner",
                                      title="Cumulative Win% by Year (Owner)", labels={"CumWinPct": "Win %"})
                    safe_chart(fig_cum)
            except Exception:
                pass
        else:
            st.info("No owner-based season records could be computed.")

    # Removed per user request: Points For (Total) and Wins vs Points For charts


def render_draft(df_draft, df_teams_owners, df_reg, selected_years, selected_teams, selected_owners, file_path: Optional[str] = None):
    st.subheader("Draft (Round 1)")
    if df_draft is None or df_draft.empty:
        # Try on-the-fly fallback: probe the workbook for any sheet that looks like 'draft'
        fallback_loaded = False
        if file_path and os.path.exists(file_path):
            try:
                xls = pd.ExcelFile(file_path, engine="openpyxl")
                target = _norm("draft")
                candidates = [s for s in xls.sheet_names if target in _norm(s) or _norm(s) in target]
                if candidates:
                    chosen = candidates[0]
                    raw = pd.read_excel(file_path, sheet_name=chosen, engine="openpyxl")
                    df_draft = normalize_sheet("draft", raw) or raw
                    st.info(f"Loaded draft data from sheet '{chosen}'.")
                    fallback_loaded = True
            except Exception:
                fallback_loaded = False
        if not fallback_loaded:
            resolved = None
            try:
                resolved = df_draft.attrs.get("__source_sheet__") if df_draft is not None else None
            except Exception:
                resolved = None
            if resolved:
                st.info(f"Draft sheet '{resolved}' appears empty or missing required columns.")
            else:
                st.info("draft sheet not found or empty.")
            return

    # Debug: show which sheet and columns we have
    with st.expander("Debug: Draft data", expanded=False):
        try:
            src = df_draft.attrs.get("__source_sheet__") if hasattr(df_draft, "attrs") else None
        except Exception:
            src = None
        if file_path and os.path.exists(file_path):
            try:
                xls = pd.ExcelFile(file_path, engine="openpyxl")
                st.write("Workbook sheets:", xls.sheet_names)
            except Exception as e:
                st.write("Could not list workbook sheets:", str(e))
        st.write("Source sheet:", src or "(unknown)")
        st.write("Rows x Cols:", (len(df_draft), len(df_draft.columns)))
        st.write("Columns:", list(df_draft.columns))
        st.dataframe(df_draft.head(10))

    df_round1 = first_round_draft(df_draft, df_teams_owners, df_reg)
    df_round1 = apply_year_team_owner_filters(df_round1, selected_years, selected_teams, selected_owners)
    df_round1, missing = ensure_required_columns(df_round1, REQUIRED_SHEETS["draft"])
    if missing:
        st.warning(f"Missing columns in draft: {missing}")

    # Sort Round 1 table: Year newest->oldest, Pick lowest->highest
    if not df_round1.empty:
        df_round1 = df_round1.copy()
        if "Year" in df_round1.columns:
            df_round1["_YearNum"] = pd.to_numeric(df_round1["Year"], errors="coerce")
        if "Pick" in df_round1.columns:
            df_round1["_PickNum"] = pd.to_numeric(df_round1["Pick"], errors="coerce")
        sort_cols = [c for c in ["_YearNum", "_PickNum"] if c in df_round1.columns]
        if sort_cols:
            df_round1 = df_round1.sort_values(sort_cols, ascending=[False, True][:len(sort_cols)])
        df_display = df_round1.drop(columns=["_YearNum", "_PickNum"], errors="ignore")
    else:
        df_display = df_round1

    # Warn if essential fields have missing values
    if not df_display.empty and {"Year", "Owner", "Pick"}.issubset(df_display.columns):
        missing_rows = df_display[["Year", "Owner", "Pick"]].isna().any(axis=1).sum()
        if missing_rows:
            st.info(f"Note: {missing_rows} Round 1 rows have missing Year/Owner/Pick and may affect averages and joins.")

    st.dataframe(df_display, use_container_width=True)
    with st.expander("Debug: Round 1 snapshot", expanded=False):
        st.write("Round1 rows:", len(df_round1))
        st.write("Columns:", list(df_round1.columns))
        st.dataframe(df_round1.head(10))

    if not df_round1.empty and "Position" in df_round1.columns:
        # Ensure consistent colors per position across both charts
        try:
            positions = [p for p in pd.unique(df_round1["Position"].dropna().astype(str))]
        except Exception:
            positions = []
        palette = px.colors.qualitative.Set2 if hasattr(px.colors.qualitative, "Set2") else px.colors.qualitative.Plotly
        pos_color_map = {p: palette[i % len(palette)] for i, p in enumerate(sorted(positions))}

        c1, c2 = st.columns(2)
        with c1:
            fig = px.histogram(
                df_round1,
                x="Position",
                color="Position",
                title="Round 1 Picks by Position",
                category_orders={"Position": sorted(positions)},
                color_discrete_map=pos_color_map,
            )
            safe_chart(fig)
        with c2:
            # #1 overall position distribution across seasons
            try:
                d1_overall = df_round1.copy()
                d1_overall["Pick"] = pd.to_numeric(d1_overall["Pick"], errors="coerce")
                d1_overall = d1_overall[d1_overall["Pick"] == 1]
                if not d1_overall.empty and "Position" in d1_overall.columns:
                    pos_counts = (
                        d1_overall.groupby("Position").size().reset_index(name="#1 Overall Count")
                    )
                    fig_no1 = px.bar(
                        pos_counts,
                        x="Position",
                        y="#1 Overall Count",
                        title="No. 1 Overall Round 1 Picks by Position",
                        category_orders={"Position": sorted(positions)},
                        color="Position",
                        color_discrete_map=pos_color_map,
                    )
                    safe_chart(fig_no1)
            except Exception:
                pass

    # Average Round 1 draft position by Owner
    st.markdown("### Average Round 1 draft position by Owner")
    if not df_round1.empty and {"Owner", "_PickNum"}.issubset(df_round1.columns):
        avg_pick = (
            df_round1.dropna(subset=["Owner", "_PickNum"]) 
            .groupby("Owner")["_PickNum"].agg(AvgPick="mean", MinPick="min", MaxPick="max", Round1Picks="count")
            .reset_index()
        )
        avg_pick["AvgPick"] = avg_pick["AvgPick"].round(2)
        avg_pick = avg_pick.sort_values(["AvgPick", "Round1Picks"], ascending=[True, False])
        st.dataframe(avg_pick, use_container_width=True)
    else:
        st.info("Not enough data to compute average Round 1 draft position by owner.")

    # Draft position vs Regular Season performance (join on Year + Owner)
    st.markdown("### Draft position vs regular-season Win%")
    if df_reg is not None and not df_reg.empty and "Owner" in df_round1.columns:
        reg = df_reg.copy()
        # Aggregate to Owner-Year with best Seed (min) and totals
        for c in ["Year", "Seed", "Wins", "Losses", "PointsFor", "PointsAgainst"]:
            if c in reg.columns:
                reg[c] = pd.to_numeric(reg[c], errors="coerce")
        # Handle optional ties column
        ties_col = None
        if "T" in reg.columns:
            ties_col = "T"
        elif "Ties" in reg.columns:
            ties_col = "Ties"
        if ties_col:
            reg["Ties"] = pd.to_numeric(reg[ties_col], errors="coerce").fillna(0)
        else:
            reg["Ties"] = 0
        reg_grp = reg.groupby(["Owner", "Year"]).agg(
            Seed=("Seed", "min") if "Seed" in reg.columns else ("Wins", "size"),
            Wins=("Wins", "sum") if "Wins" in reg.columns else ("Owner", "size"),
            Losses=("Losses", "sum") if "Losses" in reg.columns else ("Owner", "size"),
            Ties=("Ties", "sum"),
            PointsFor=("PointsFor", "sum") if "PointsFor" in reg.columns else ("Owner", "size"),
            PointsAgainst=("PointsAgainst", "sum") if "PointsAgainst" in reg.columns else ("Owner", "size"),
        ).reset_index()

        # Compute Win%
        if {"Wins", "Losses"}.issubset(reg_grp.columns):
            reg_grp["Wins"] = pd.to_numeric(reg_grp["Wins"], errors="coerce").fillna(0)
            reg_grp["Losses"] = pd.to_numeric(reg_grp["Losses"], errors="coerce").fillna(0)
            reg_grp["Ties"] = pd.to_numeric(reg_grp["Ties"], errors="coerce").fillna(0)
            reg_grp["Games"] = reg_grp["Wins"] + reg_grp["Losses"] + reg_grp["Ties"]
            reg_grp["WinPct"] = np.where(reg_grp["Games"] > 0, (reg_grp["Wins"] + 0.5 * reg_grp["Ties"]) / reg_grp["Games"], np.nan)

        # Round1 pick per Owner-Year (earliest pick if multiple)
        d1 = df_round1.copy()
        d1["Pick"] = pd.to_numeric(d1["Pick"], errors="coerce")
        d1["Year"] = pd.to_numeric(d1["Year"], errors="coerce")
        d1_grp = d1.dropna(subset=["Year", "Owner", "Pick"]).sort_values("Pick").groupby(["Owner", "Year"]).first().reset_index()

    dv = pd.merge(d1_grp, reg_grp, on=["Owner", "Year"], how="inner")
    dv = dv.dropna(subset=["Pick"]) 
    if "WinPct" in dv.columns:
            # Scatter: Pick vs Win%
            fig_sc = px.scatter(
                dv, x="Pick", y="WinPct", hover_data=["Owner", "Year", "Wins" if "Wins" in dv.columns else None],
                trendline="ols", title="Round 1 Pick vs Regular-season Win%",
                labels={"WinPct": "Win %"}
            )
            safe_chart(fig_sc)

            # (Removed) Average Win% by Draft Position (quartiles) per user request

            # (Removed) Rank correlation chart per user request

            # Overall Win% per Draft Pick (across all years)
            try:
                pick_stats = dv.groupby("Pick").agg(
                    AvgWinPct=("WinPct", "mean"),
                    MedianWinPct=("WinPct", "median"),
                    Samples=("WinPct", "count"),
                ).reset_index()
                fig_pick = px.bar(
                    pick_stats.sort_values("Pick"), x="Pick", y="AvgWinPct",
                    title="Average Win% by Draft Pick (all seasons)", labels={"AvgWinPct": "Win %"},
                    text="Samples"
                )
                fig_pick.update_traces(textposition="outside", cliponaxis=False)
                safe_chart(fig_pick)
            except Exception:
                pass


def render_head_to_head(df_gl, selected_years, selected_teams, selected_owners):
    st.subheader("Head-to-Head & Game Log")
    if df_gl is None or df_gl.empty:
        st.info("gamelog sheet not found or empty.")
        return

    df = compute_winner_team(df_gl)
    df = apply_year_team_owner_filters(df, selected_years, selected_teams, selected_owners)

    # Runtime fallback: alias legacy owner columns to normalized names if needed
    if "HomeOwner" not in df.columns:
        for cand in ["Home Owner", "A Owner", "Owner A", "Team A Owner"]:
            if cand in df.columns:
                df = df.copy()
                df["HomeOwner"] = df[cand]
                break
    if "AwayOwner" not in df.columns:
        for cand in ["Away Owner", "B Owner", "Owner B", "Team B Owner", "Owner.1"]:
            if cand in df.columns:
                df = df.copy()
                df["AwayOwner"] = df[cand]
                break

    # Ensure scores exist for charts
    if not {"HomeScore", "AwayScore"}.issubset(df.columns):
        cols = list(df.columns)
        norm_map = {c: _norm(c) for c in cols}
        def _find(name: str) -> Optional[str]:
            target = _norm(name)
            for c, n in norm_map.items():
                if n == target:
                    return c
            return None
        a_pts = _find("team a points") or _find("points a")
        b_pts = _find("team b points") or _find("points b")
        if a_pts and b_pts:
            df = df.copy()
            df["HomeScore"] = pd.to_numeric(df[a_pts], errors="coerce")
            df["AwayScore"] = pd.to_numeric(df[b_pts], errors="coerce")

    st.markdown("### Head-to-Head (Owners)")
    with st.expander("Debug: Game Log columns (Head-to-Head)", expanded=False):
        st.write("Columns:", list(df.columns))
        st.dataframe(df.head(10))
    subset_df = None
    if {"HomeOwner", "AwayOwner"}.issubset(df.columns):
        options = sorted(pd.unique(pd.concat([df["HomeOwner"], df["AwayOwner"]], ignore_index=True).dropna().astype(str)))
        a = st.selectbox("Side A (Owner)", options=options, index=0 if options else None)
        b = st.selectbox("Side B (Owner)", options=[o for o in options if o != a] if options else [], index=0 if len(options) > 1 else None)
        if a and b:
            subset = df[((df["HomeOwner"] == a) & (df["AwayOwner"] == b)) | ((df["HomeOwner"] == b) & (df["AwayOwner"] == a))].copy()
            show_head_to_head_summary(subset, label_a=a, label_b=b, by="owner")
            subset_df = subset
    else:
        st.info("Owner columns not found in gamelog.")

    st.markdown("### Game Log")
    table_df = subset_df if subset_df is not None else df
    # Sort safely by Year/Week if present
    if {"Year", "Week"}.issubset(table_df.columns):
        tmp = table_df.copy()
        tmp["_Y"] = pd.to_numeric(tmp["Year"], errors="coerce")
        tmp["_W"] = pd.to_numeric(tmp["Week"], errors="coerce")
        tmp = tmp.sort_values(["_Y", "_W", "Year", "Week"]).drop(columns=["_Y", "_W"], errors="ignore")
        st.dataframe(tmp.reset_index(drop=True), use_container_width=True)
    else:
        st.dataframe(table_df.reset_index(drop=True), use_container_width=True)

    # All owners head-to-head records (within current filters)
    st.markdown("### All owners head-to-head records")
    needed_cols = {"HomeOwner", "AwayOwner", "HomeScore", "AwayScore"}
    if needed_cols.issubset(df.columns):
        d = df.copy()
        # Ensure numeric scores
        d["HomeScore"] = pd.to_numeric(d["HomeScore"], errors="coerce")
        d["AwayScore"] = pd.to_numeric(d["AwayScore"], errors="coerce")
        d = d.dropna(subset=["HomeOwner", "AwayOwner", "HomeScore", "AwayScore"]).copy()
        if not d.empty:
            # Long form: one row per owner perspective per game
            home = d[["HomeOwner", "AwayOwner", "HomeScore", "AwayScore"]].copy()
            home.columns = ["Owner", "Opponent", "PF", "PA"]
            home["Win"] = (home["PF"] > home["PA"]).astype(int)
            home["Loss"] = (home["PF"] < home["PA"]).astype(int)
            home["Tie"] = (home["PF"] == home["PA"]).astype(int)

            away = d[["AwayOwner", "HomeOwner", "AwayScore", "HomeScore"]].copy()
            away.columns = ["Owner", "Opponent", "PF", "PA"]
            away["Win"] = (away["PF"] > away["PA"]).astype(int)
            away["Loss"] = (away["PF"] < away["PA"]).astype(int)
            away["Tie"] = (away["PF"] == away["PA"]).astype(int)

            long_df = pd.concat([home, away], ignore_index=True)
            # Remove self-matches just in case
            long_df = long_df[long_df["Owner"].astype(str) != long_df["Opponent"].astype(str)]
            # Aggregate by Owner vs Opponent
            agg = long_df.groupby(["Owner", "Opponent"]).agg(
                Games=("Win", "count"),
                Wins=("Win", "sum"),
                Losses=("Loss", "sum"),
                Ties=("Tie", "sum"),
                PF=("PF", "sum"),
                PA=("PA", "sum"),
            ).reset_index()
            # Derived metrics
            agg["WinPct"] = np.where(
                agg["Games"] > 0,
                (agg["Wins"] + 0.5 * agg["Ties"]) / agg["Games"],
                np.nan,
            )
            agg["AvgPF"] = np.where(agg["Games"] > 0, agg["PF"] / agg["Games"], np.nan)
            agg["AvgPA"] = np.where(agg["Games"] > 0, agg["PA"] / agg["Games"], np.nan)
            agg["PtsDiff"] = agg["PF"] - agg["PA"]
            # Formatting
            agg["WinPct"] = agg["WinPct"].round(3)
            agg["AvgPF"] = agg["AvgPF"].round(1)
            agg["AvgPA"] = agg["AvgPA"].round(1)
            # Order columns
            cols = [
                "Owner", "Opponent", "Games", "Wins", "Losses", "Ties", "WinPct",
                "PF", "PA", "AvgPF", "AvgPA", "PtsDiff",
            ]
            # Sort by Owner then Opponent (and games desc within owner)
            agg = agg.sort_values(["Owner", "Games", "Opponent"], ascending=[True, False, True])
            st.dataframe(agg[cols], use_container_width=True)
        else:
            st.info("No complete owner matchup rows found for the current filters.")
    else:
        missing = needed_cols.difference(df.columns)
        st.info(f"Cannot compute owner matchup table; missing columns: {sorted(missing)}")


def show_head_to_head_summary(df: pd.DataFrame, label_a: str, label_b: str, by: str):
    """Render a compact head-to-head summary and charts."""
    if df is None or df.empty:
        st.info("No games found for this matchup.")
        return

    df = df.copy()
    if by == "team":
        df["A_Score"] = df.apply(lambda r: r["HomeScore"] if r["HomeTeam"] == label_a else r["AwayScore"], axis=1)
        df["B_Score"] = df.apply(lambda r: r["AwayScore"] if r["HomeTeam"] == label_a else r["HomeScore"], axis=1)
        df["A_Won"] = df["WinnerTeam"] == label_a
    else:
        df["A_Score"] = df.apply(lambda r: r["HomeScore"] if r["HomeOwner"] == label_a else r["AwayScore"], axis=1)
        df["B_Score"] = df.apply(lambda r: r["AwayScore"] if r["HomeOwner"] == label_a else r["HomeScore"], axis=1)
        df["A_Won"] = df.apply(lambda r: (r["HomeOwner"] == label_a and r["HomeScore"] > r["AwayScore"]) or (r["AwayOwner"] == label_a and r["AwayScore"] > r["HomeScore"]), axis=1)

    wins = int(df["A_Won"].sum())
    losses = int((~df["A_Won"]).sum())
    pf = float(df["A_Score"].sum())
    pa = float(df["B_Score"].sum())

    c1, c2, c3, c4 = st.columns(4)
    c1.metric(f"{label_a} Wins", wins)
    c2.metric(f"{label_b} Wins", losses)
    c3.metric(f"{label_a} Points For", f"{int(pf)}")
    c4.metric(f"{label_b} Points For", f"{int(pa)}")

    # Chart granularity selector
    granularity = st.selectbox(
        "Chart granularity",
        ["Per game", "Per year: average", "Per year: median", "Per year: max"],
        index=0,
    )

    def _plot_per_game(data: pd.DataFrame):
        fig = go.Figure()
        try:
            x_vals = pd.to_numeric(data["Year"], errors="coerce").astype("Int64")
        except Exception:
            x_vals = data["Year"]
        fig.add_trace(go.Scatter(x=x_vals, y=data["A_Score"], mode="lines+markers", name=label_a))
        fig.add_trace(go.Scatter(x=x_vals, y=data["B_Score"], mode="lines+markers", name=label_b))
        fig.update_layout(title=f"Scores Over Time: {label_a} vs {label_b}")
        safe_chart(fig)

    def _plot_per_year(data: pd.DataFrame, func_name: str, title_suffix: str):
        d = data.copy()
        d["YearNum"] = pd.to_numeric(d["Year"], errors="coerce")
        d = d.dropna(subset=["YearNum"])  # remove rows without a numeric year
        if d.empty:
            _plot_per_game(data)
            return
        func = {"average": "mean", "median": "median", "max": "max"}[func_name]
        agg = d.groupby("YearNum").agg(
            A_Score=("A_Score", func),
            B_Score=("B_Score", func),
            Games=("A_Score", "count"),
        ).reset_index().sort_values("YearNum")
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=agg["YearNum"].astype(int), y=agg["A_Score"], mode="lines+markers", name=label_a,
            customdata=np.stack([agg["Games"]], axis=-1),
            hovertemplate="Year %{x}<br>Score %{y:.0f}<br>Games %{customdata[0]}<extra></extra>",
        ))
        fig.add_trace(go.Scatter(
            x=agg["YearNum"].astype(int), y=agg["B_Score"], mode="lines+markers", name=label_b,
            customdata=np.stack([agg["Games"]], axis=-1),
            hovertemplate="Year %{x}<br>Score %{y:.0f}<br>Games %{customdata[0]}<extra></extra>",
        ))
        fig.update_layout(title=f"Scores Over Time: {label_a} vs {label_b} ({title_suffix})")
        safe_chart(fig)

    if granularity == "Per game":
        _plot_per_game(df)
    elif granularity == "Per year: average":
        _plot_per_year(df, "average", "avg per year")
    elif granularity == "Per year: median":
        _plot_per_year(df, "median", "median per year")
    else:
        _plot_per_year(df, "max", "max per year")


def render_teams_owners(df_to, selected_years, selected_teams, selected_owners):
    st.subheader("Teams & Owners")
    if df_to is None or df_to.empty:
        st.info("teams_owners sheet not found or empty.")
        return

    df = apply_year_team_owner_filters(df_to, selected_years, selected_teams, selected_owners)
    df, missing = ensure_required_columns(df, REQUIRED_SHEETS["teams_owners"])
    if missing:
        st.warning(f"Missing columns in teams_owners: {missing}")

    sort_cols = [c for c in ["Year", "TeamName"] if c in df.columns]
    st.dataframe(df.sort_values(sort_cols) if sort_cols else df, use_container_width=True)

    if not df.empty:
        if "Year" in df.columns:
            by_owner = df.groupby("Owner")["Year"].nunique().reset_index(name="Seasons")
            title = "Seasons by Owner"
            y_col = "Seasons"
        else:
            # Fallback: count unique teams per owner when Season/Year not provided
            if "TeamName" in df.columns:
                by_owner = df.groupby("Owner")["TeamName"].nunique().reset_index(name="Teams")
                y_col = "Teams"
                title = "Teams by Owner"
            else:
                by_owner = df.groupby("Owner").size().reset_index(name="Entries")
                y_col = "Entries"
                title = "Entries by Owner"
        fig = px.bar(by_owner.sort_values(y_col, ascending=False), x="Owner", y=y_col, title=title)
        safe_chart(fig)


# --------------------- Main App ---------------------


def render_rating(df_gl, selected_years, selected_teams, selected_owners):
    """Elo-style rating tab computed from all games in gamelog.

    Features:
    - Choose to rate by Team or Owner.
    - Controls for K-factor and initial rating.
    - Leaderboard (current Elo), timeline for selected entries, and match details.
    """
    #st.subheader("Rating (Elo) – Owners")
    if df_gl is None or df_gl.empty:
        st.info("gamelog sheet not found or empty.")
        return

    # Filter per sidebar selections (years/teams/owners)
    df = apply_year_team_owner_filters(df_gl, selected_years, selected_teams, selected_owners)

    # Ensure required columns
    needed_cols = {"Year", "Week", "HomeTeam", "AwayTeam", "HomeScore", "AwayScore"}
    # Try to infer scores if missing
    if not {"HomeScore", "AwayScore"}.issubset(df.columns):
        cols = list(df.columns)
        norm_map = {c: _norm(c) for c in cols}
        def _find(name: str) -> Optional[str]:
            target = _norm(name)
            for c, n in norm_map.items():
                if n == target:
                    return c
            return None
        a_pts = _find("team a points") or _find("points a")
        b_pts = _find("team b points") or _find("points b")
        if a_pts and b_pts:
            df = df.copy()
            df["HomeScore"] = pd.to_numeric(df[a_pts], errors="coerce")
            df["AwayScore"] = pd.to_numeric(df[b_pts], errors="coerce")
    # Basic guards
    for sc in ["HomeScore", "AwayScore"]:
        if sc in df.columns:
            df[sc] = pd.to_numeric(df[sc], errors="coerce")

    # Fixed configuration (no controls per request)
    initial_elo = 1000
    k_factor = 32
    #t.caption("Elo config: Initial = 1000, K = 32")

    # Build entity columns (Owner-only)
    ent_a_col = "HomeOwner" if "HomeOwner" in df.columns else next((c for c in ["Home Owner", "A Owner", "Owner A", "Team A Owner"] if c in df.columns), None)
    ent_b_col = "AwayOwner" if "AwayOwner" in df.columns else next((c for c in ["Away Owner", "B Owner", "Owner B", "Team B Owner", "Owner.1"] if c in df.columns), None)
    if ent_a_col is None or ent_b_col is None or not {ent_a_col, ent_b_col}.issubset(df.columns):
        st.info("Not enough columns to rate by owner.")
        return

    # Sort chronologically by Year then Week (numeric when possible)
    d = df.copy()
    d["_Y"] = pd.to_numeric(d.get("Year"), errors="coerce")
    d["_W"] = pd.to_numeric(d.get("Week"), errors="coerce")
    d = d.dropna(subset=["HomeScore", "AwayScore"]).copy()
    d = d.sort_values(["_Y", "_W", "Year", "Week"]).reset_index(drop=True)
    if d.empty:
        st.info("No completed games found to compute ratings.")
        return

    # Elo helpers
    def expected_score(elo_a: float, elo_b: float) -> float:
        return 1.0 / (1.0 + 10 ** ((elo_b - elo_a) / 400.0))

    def update_pair(elo_a: float, elo_b: float, result_a: float, k: float) -> Tuple[float, float]:
        exp_a = expected_score(elo_a, elo_b)
        exp_b = 1.0 - exp_a
        new_a = elo_a + k * (result_a - exp_a)
        new_b = elo_b + k * ((1.0 - result_a) - exp_b)
        return round(float(new_a), 2), round(float(new_b), 2)

    # Iterate matches
    ratings: Dict[str, float] = {}
    tally: Dict[str, Dict[str, int]] = {}
    rows = []
    # Track per-entity history for timeline
    history: Dict[str, List[Tuple[int, float]]] = {}

    for idx, r in d.iterrows():
        a = str(r.get(ent_a_col)) if pd.notna(r.get(ent_a_col)) else None
        b = str(r.get(ent_b_col)) if pd.notna(r.get(ent_b_col)) else None
        if not a or not b:
            continue
        hs = r.get("HomeScore")
        as_ = r.get("AwayScore")
        if pd.isna(hs) or pd.isna(as_):
            continue
        # Initialize
        ra = ratings.get(a, float(initial_elo))
        rb = ratings.get(b, float(initial_elo))
        # Result from A's perspective (Home side of chosen mode)
        if hs > as_:
            res_a = 1.0
            winner = a
        elif as_ > hs:
            res_a = 0.0
            winner = b
        else:
            res_a = 0.5
            winner = None
        # Update
        new_a, new_b = update_pair(ra, rb, res_a, float(k_factor))

        # Tally
        if a not in tally:
            tally[a] = {"Games": 0, "Wins": 0, "Losses": 0, "Ties": 0}
        if b not in tally:
            tally[b] = {"Games": 0, "Wins": 0, "Losses": 0, "Ties": 0}
        tally[a]["Games"] += 1
        tally[b]["Games"] += 1
        if res_a == 1.0:
            tally[a]["Wins"] += 1
            tally[b]["Losses"] += 1
        elif res_a == 0.0:
            tally[b]["Wins"] += 1
            tally[a]["Losses"] += 1
        else:
            tally[a]["Ties"] += 1
            tally[b]["Ties"] += 1

        rows.append({
            "Year": r.get("Year"),
            "Week": r.get("Week"),
            "A": a,
            "B": b,
            "A_elo_start": round(float(ra), 2),
            "B_elo_start": round(float(rb), 2),
            "A_elo_end": new_a,
            "B_elo_end": new_b,
            "A_change": round(new_a - ra, 2),
            "B_change": round(new_b - rb, 2),
            "Winner": winner,
        })

        ratings[a] = new_a
        ratings[b] = new_b

        # History points use a running index for x-axis
        for name, elo in [(a, new_a), (b, new_b)]:
            if name not in history:
                history[name] = []
            history[name].append((idx, float(elo)))

        # Helper to capture end-of-previous-week ratings for movement arrows
        def _prev_week_snapshot(df_in: pd.DataFrame) -> Dict[str, float]:
            ratings_sim: Dict[str, float] = {}
            last_key = None
            snapshots: List[Tuple[Tuple[Optional[float], Optional[float]], Dict[str, float]]] = []
            for _, r in df_in.iterrows():
                key = (r.get("_Y"), r.get("_W"))
                if last_key is None:
                    last_key = key
                elif key != last_key:
                    snapshots.append((last_key, dict(ratings_sim)))
                    last_key = key
                a = str(r.get(ent_a_col)) if pd.notna(r.get(ent_a_col)) else None
                b = str(r.get(ent_b_col)) if pd.notna(r.get(ent_b_col)) else None
                if not a or not b:
                    continue
                hs = r.get("HomeScore"); as_ = r.get("AwayScore")
                if pd.isna(hs) or pd.isna(as_):
                    continue
                ra = ratings_sim.get(a, float(initial_elo))
                rb = ratings_sim.get(b, float(initial_elo))
                # result from home perspective
                if hs > as_:
                    res_a = 1.0
                elif as_ > hs:
                    res_a = 0.0
                else:
                    res_a = 0.5
                exp_a = 1.0 / (1.0 + 10 ** ((rb - ra) / 400.0))
                exp_b = 1.0 - exp_a
                ratings_sim[a] = round(ra + float(k_factor) * (res_a - exp_a), 2)
                ratings_sim[b] = round(rb + float(k_factor) * ((1.0 - res_a) - exp_b), 2)
            if last_key is not None:
                snapshots.append((last_key, dict(ratings_sim)))
            if len(snapshots) >= 2:
                return snapshots[-2][1]
            return {}

    # Leaderboard
    lb = pd.DataFrame([
        {"Owner": k, "Elo": v, **tally.get(k, {"Games": 0, "Wins": 0, "Losses": 0, "Ties": 0})}
        for k, v in ratings.items()
    ])
    if lb.empty:
        st.info("No ratings could be computed from the selected data.")
        return
    lb = lb.sort_values(["Elo", "Wins", "Games"], ascending=[False, False, False])

    # Compute win percentage for display
    lb = lb.copy()
    lb["WinPct"] = np.where(lb["Games"] > 0, (lb["Wins"] + 0.5 * lb["Ties"]) / lb["Games"], np.nan)
    lb["WinPct"] = lb["WinPct"].round(3)

    st.markdown("### Current Elo leaderboard")
    top_n = min(12, len(lb))
    top12 = lb.head(top_n).reset_index(drop=True)
    top12["Rank"] = np.arange(1, len(top12) + 1)
    # Movement deltas vs previous week
    prev_ratings = _prev_week_snapshot(d)
    if prev_ratings:
        prev_sorted = sorted(prev_ratings.items(), key=lambda kv: kv[1], reverse=True)
        prev_rank_map = {name: i + 1 for i, (name, _) in enumerate(prev_sorted)}
    else:
        prev_rank_map = {}
    elo_deltas = []
    rank_deltas = []
    for _, rr in top12.iterrows():
        name = rr["Owner"]
        curr_elo = float(rr["Elo"]) if pd.notna(rr["Elo"]) else np.nan
        prev_elo = prev_ratings.get(name, np.nan)
        elo_deltas.append(curr_elo - prev_elo if pd.notna(curr_elo) and pd.notna(prev_elo) else np.nan)
        prev_rank = prev_rank_map.get(name, np.nan)
        curr_rank = int(rr["Rank"]) if pd.notna(rr["Rank"]) else np.nan
        rank_deltas.append(prev_rank - curr_rank if pd.notna(prev_rank) and pd.notna(curr_rank) else np.nan)
    top12["EloDelta"] = elo_deltas
    top12["RankDelta"] = rank_deltas
    # Build modern table HTML (full width)
    html = [
        "<div class='standings-wrap'>",
        "<div class='standings-title'>🏅 Elo Leaderboard</div>",
        "<table class='table-modern'>",
        "<thead><tr>",
        "<th>Rank</th><th>Owner</th><th style='text-align:right;'>Elo</th><th style='text-align:right;'>Games</th><th style='text-align:right;'>W</th><th style='text-align:right;'>L</th><th style='text-align:right;'>T</th><th style='text-align:right;'>Win%</th>",
        "</tr></thead><tbody>",
    ]
    def _rank_arrow(val: float) -> str:
        if pd.isna(val):
            return ""
        if val > 0:
            return f" <span style='color:#2e7d32;font-size:0.85em;'>▲+{int(val)}</span>"
        if val < 0:
            return f" <span style='color:#c62828;font-size:0.85em;'>▼{int(abs(val))}</span>"
        return " <span style='color:#6b7280;font-size:0.85em;'>•</span>"

    def _elo_arrow(val: float) -> str:
        if pd.isna(val):
            return ""
        if val > 0:
            return f" <span style='color:#2e7d32;font-size:0.85em;'>+{int(round(val))}</span>"
        if val < 0:
            return f" <span style='color:#c62828;font-size:0.85em;'>-{int(round(abs(val)))}</span>"
        return " <span style='color:#6b7280;font-size:0.85em;'>0</span>"

    for _, r in top12.iterrows():
        html.append("<tr>")
        rank_cell = f"{int(r['Rank'])}" + _rank_arrow(r.get("RankDelta"))
        html.append(f"<td>{rank_cell}</td>")
        html.append(f"<td>{r['Owner']}</td>")
        elo_cell = f"{int(round(r['Elo']))}" + _elo_arrow(r.get("EloDelta"))
        html.append(f"<td class='wl' style='text-align:right;'>{elo_cell}</td>")
        html.append(f"<td class='wl' style='text-align:right;'>{int(r['Games'])}</td>")
        html.append(f"<td class='wl' style='text-align:right;'>{int(r['Wins'])}</td>")
        html.append(f"<td class='wl' style='text-align:right;'>{int(r['Losses'])}</td>")
        html.append(f"<td class='wl' style='text-align:right;'>{int(r['Ties'])}</td>")
        wp = ("" if pd.isna(r["WinPct"]) else f"{r['WinPct']:.3f}")
        html.append(f"<td class='pct' style='text-align:right;'>{wp}</td>")
        html.append("</tr>")
    html.append("</tbody></table></div>")
    st.markdown("".join(html), unsafe_allow_html=True)

    # Timeline chart for selected entries
    st.markdown("### Rating timeline")
    options_all = list(lb["Owner"].values)
    mode_pick = st.selectbox("Owners to plot", ["All owners", "Pick owners..."] , index=0)
    if mode_pick == "All owners":
        picks = options_all
    else:
        with st.expander("Choose owners", expanded=True):
            picks = st.multiselect("", options=options_all, default=options_all[:5])
    if picks:
        fig = go.Figure()
        for name in picks:
            series = history.get(name, [])
            if not series:
                continue
            xs = [i for i, _ in series]
            ys = [e for _, e in series]
            fig.add_trace(go.Scatter(x=xs, y=ys, mode="lines", name=name))
        fig.update_layout(title="Elo over games (Owner)", xaxis_title="Game index (within selection)", yaxis_title="Elo")
        safe_chart(fig)
    else:
        st.info("Pick at least one to plot the timeline.")

    # Records section (after timeline)
    try:
        st.markdown("### Records")

        # Build details DataFrame if not already
        details = pd.DataFrame(rows)
        if not details.empty:
            # Compute biggest Elo gain/loss (single match)
            max_gain_val = None
            max_gain_owner = None
            max_gain_yw = None
            min_loss_val = None
            min_loss_owner = None
            min_loss_yw = None
            if "A_change" in details.columns and "B_change" in details.columns:
                # Max gain among A and B
                if details["A_change"].notna().any():
                    idx_a = details["A_change"].idxmax()
                    max_a = details.loc[idx_a, "A_change"]
                    owner_a = details.loc[idx_a, "A"]
                    yw_a = (details.loc[idx_a, "Year"], details.loc[idx_a, "Week"]) if {"Year", "Week"}.issubset(details.columns) else (None, None)
                else:
                    max_a = -np.inf
                    owner_a = None
                    yw_a = (None, None)
                if details["B_change"].notna().any():
                    idx_b = details["B_change"].idxmax()
                    max_b = details.loc[idx_b, "B_change"]
                    owner_b = details.loc[idx_b, "B"]
                    yw_b = (details.loc[idx_b, "Year"], details.loc[idx_b, "Week"]) if {"Year", "Week"}.issubset(details.columns) else (None, None)
                else:
                    max_b = -np.inf
                    owner_b = None
                    yw_b = (None, None)
                if max_a >= max_b:
                    max_gain_val, max_gain_owner, max_gain_yw = max_a, owner_a, yw_a
                else:
                    max_gain_val, max_gain_owner, max_gain_yw = max_b, owner_b, yw_b

                # Most negative change (largest drop)
                if details["A_change"].notna().any():
                    idx_a_min = details["A_change"].idxmin()
                    min_a = details.loc[idx_a_min, "A_change"]
                    min_a_owner = details.loc[idx_a_min, "A"]
                    min_a_yw = (details.loc[idx_a_min, "Year"], details.loc[idx_a_min, "Week"]) if {"Year", "Week"}.issubset(details.columns) else (None, None)
                else:
                    min_a = np.inf
                    min_a_owner = None
                    min_a_yw = (None, None)
                if details["B_change"].notna().any():
                    idx_b_min = details["B_change"].idxmin()
                    min_b = details.loc[idx_b_min, "B_change"]
                    min_b_owner = details.loc[idx_b_min, "B"]
                    min_b_yw = (details.loc[idx_b_min, "Year"], details.loc[idx_b_min, "Week"]) if {"Year", "Week"}.issubset(details.columns) else (None, None)
                else:
                    min_b = np.inf
                    min_b_owner = None
                    min_b_yw = (None, None)
                if min_a <= min_b:
                    min_loss_val, min_loss_owner, min_loss_yw = min_a, min_a_owner, min_a_yw
                else:
                    min_loss_val, min_loss_owner, min_loss_yw = min_b, min_b_owner, min_b_yw

        # Peak/lowest Elo and #1 rankings require tracking; compute from history/ratings snapshots
        # Re-simulate quick pass to gather peaks, troughs, and weekly #1 counts
        peaks: Dict[str, float] = {}
        troughs: Dict[str, float] = {}
        weekly_num1: Dict[str, int] = {}
        best_streak_owner = None
        best_streak_len = 0
        best_streak_start_key = None  # (Year, Week)
        best_streak_end_key = None    # (Year, Week)
        curr_streak_owner = None
        curr_streak_len = 0
        curr_streak_start_key = None  # (Year, Week)
        curr_streak_end_key = None    # (Year, Week)
        ratings_sim: Dict[str, float] = {}
        last_key = None
        # Iterate same order as before
        for idx, r in d.iterrows():
            key = (r.get("_Y"), r.get("_W"))
            if last_key is None:
                last_key = key
            elif key != last_key:
                # finalize last week leader
                if ratings_sim:
                    leader = max(ratings_sim.items(), key=lambda kv: kv[1])[0]
                    weekly_num1[leader] = weekly_num1.get(leader, 0) + 1
                    if curr_streak_owner == leader:
                        curr_streak_len += 1
                        curr_streak_end_key = last_key
                    else:
                        if curr_streak_owner is not None:
                            if curr_streak_len > best_streak_len:
                                best_streak_len = curr_streak_len
                                best_streak_owner = curr_streak_owner
                                best_streak_start_key = curr_streak_start_key
                                best_streak_end_key = curr_streak_end_key
                        curr_streak_owner = leader
                        curr_streak_len = 1
                        curr_streak_start_key = last_key
                        curr_streak_end_key = last_key
                last_key = key
            # process match
            a = str(r.get(ent_a_col)) if pd.notna(r.get(ent_a_col)) else None
            b = str(r.get(ent_b_col)) if pd.notna(r.get(ent_b_col)) else None
            if not a or not b:
                continue
            hs = r.get("HomeScore"); as_ = r.get("AwayScore")
            if pd.isna(hs) or pd.isna(as_):
                continue
            ra = ratings_sim.get(a, float(initial_elo))
            rb = ratings_sim.get(b, float(initial_elo))
            # result from home perspective
            if hs > as_:
                res_a = 1.0
            elif as_ > hs:
                res_a = 0.0
            else:
                res_a = 0.5
            # update
            exp_a = 1.0 / (1.0 + 10 ** ((rb - ra) / 400.0))
            exp_b = 1.0 - exp_a
            new_a = ra + float(k_factor) * (res_a - exp_a)
            new_b = rb + float(k_factor) * ((1.0 - res_a) - exp_b)
            ratings_sim[a] = round(float(new_a), 2)
            ratings_sim[b] = round(float(new_b), 2)
            # peaks/troughs
            peaks[a] = max(peaks.get(a, float(initial_elo)), ratings_sim[a])
            peaks[b] = max(peaks.get(b, float(initial_elo)), ratings_sim[b])
            troughs[a] = min(troughs.get(a, float(initial_elo)), ratings_sim[a]) if a in troughs else ratings_sim[a]
            troughs[b] = min(troughs.get(b, float(initial_elo)), ratings_sim[b]) if b in troughs else ratings_sim[b]
        # finalize last week's leader
        if ratings_sim:
            leader = max(ratings_sim.items(), key=lambda kv: kv[1])[0]
            weekly_num1[leader] = weekly_num1.get(leader, 0) + 1
            if curr_streak_owner == leader:
                curr_streak_len += 1
                curr_streak_end_key = last_key
            else:
                if curr_streak_owner is not None:
                    if curr_streak_len > best_streak_len:
                        best_streak_len = curr_streak_len
                        best_streak_owner = curr_streak_owner
                        best_streak_start_key = curr_streak_start_key
                        best_streak_end_key = curr_streak_end_key
                curr_streak_owner = leader
                curr_streak_len = 1
                curr_streak_start_key = last_key
                curr_streak_end_key = last_key
        # finalize best streak
        if curr_streak_len > best_streak_len:
            best_streak_len = curr_streak_len
            best_streak_owner = curr_streak_owner
            best_streak_start_key = curr_streak_start_key
            best_streak_end_key = curr_streak_end_key
        # Build a long-form DataFrame to locate Year/Week for per-owner Elo snapshots
        peak_owner = None
        peak_val = None
        peak_yw = (None, None)
        low_owner = None
        low_val = None
        low_yw = (None, None)
        most_weeks_owner = max(weekly_num1.items(), key=lambda kv: kv[1])[0] if weekly_num1 else None
        most_weeks_val = weekly_num1.get(most_weeks_owner) if most_weeks_owner else None
        try:
            long_rows = []
            for _, rr in details.iterrows():
                y = rr.get("Year")
                w = rr.get("Week")
                if pd.notna(rr.get("A")) and pd.notna(rr.get("A_elo_end")):
                    long_rows.append({"Owner": rr.get("A"), "Elo": float(rr.get("A_elo_end")), "Year": y, "Week": w})
                if pd.notna(rr.get("B")) and pd.notna(rr.get("B_elo_end")):
                    long_rows.append({"Owner": rr.get("B"), "Elo": float(rr.get("B_elo_end")), "Year": y, "Week": w})
            L = pd.DataFrame(long_rows)
            if not L.empty:
                idx_peak = L["Elo"].idxmax()
                r_peak = L.loc[idx_peak]
                peak_owner = str(r_peak["Owner"]) if pd.notna(r_peak.get("Owner")) else None
                peak_val = float(r_peak["Elo"])
                peak_yw = (pd.to_numeric(r_peak.get("Year"), errors="coerce"), pd.to_numeric(r_peak.get("Week"), errors="coerce"))
                idx_low = L["Elo"].idxmin()
                r_low = L.loc[idx_low]
                low_owner = str(r_low["Owner"]) if pd.notna(r_low.get("Owner")) else None
                low_val = float(r_low["Elo"])
                low_yw = (pd.to_numeric(r_low.get("Year"), errors="coerce"), pd.to_numeric(r_low.get("Week"), errors="coerce"))
        except Exception:
            # Fallback to dict-based values if long-form fails
            peak_owner = max(peaks.items(), key=lambda kv: kv[1])[0] if peaks else None
            peak_val = (peaks.get(peak_owner) if peak_owner else None)
            low_owner = min(troughs.items(), key=lambda kv: kv[1])[0] if troughs else None
            low_val = (troughs.get(low_owner) if low_owner else None)

        # Render record cards
        st.markdown('<div class="record-grid">', unsafe_allow_html=True)
        # Peak Elo rating
        if peak_owner is not None and peak_val is not None:
            py, pw = peak_yw if 'peak_yw' in locals() else (None, None)
            y_txt = f", {int(pd.to_numeric(py, errors='coerce'))}" if pd.notna(py) else ""
            w_txt = f" (Wk {int(pd.to_numeric(pw, errors='coerce'))})" if pd.notna(pw) else ""
            st.markdown(
                f"""
                <div class='record-card'>
                    <div class='record-title'><span class='pill'>Peak Elo rating</span></div>
                    <div class='record-stat'>{int(round(peak_val))}</div>
                    <div class='record-sub'>{peak_owner}{w_txt}{y_txt}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        # Lowest Elo rating
        if low_owner is not None and low_val is not None:
            ly, lw = low_yw if 'low_yw' in locals() else (None, None)
            y_txt = f", {int(pd.to_numeric(ly, errors='coerce'))}" if pd.notna(ly) else ""
            w_txt = f" (Wk {int(pd.to_numeric(lw, errors='coerce'))})" if pd.notna(lw) else ""
            st.markdown(
                f"""
                <div class='record-card'>
                    <div class='record-title'><span class='pill'>Lowest Elo rating</span></div>
                    <div class='record-stat'>{int(round(low_val))}</div>
                    <div class='record-sub'>{low_owner}{w_txt}{y_txt}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        # Most weeks ranked #1
        if most_weeks_owner is not None and most_weeks_val is not None:
            st.markdown(
                f"""
                <div class='record-card'>
                    <div class='record-title'><span class='pill'>Most weeks ranked #1</span></div>
                    <div class='record-stat'>{int(most_weeks_val)}</div>
                    <div class='record-sub'>{most_weeks_owner}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        # Longest #1 streak
        if best_streak_owner is not None and best_streak_len:
            # Format start/end range if available
            def _fmt_key(k):
                if not k or (pd.isna(k[0]) and pd.isna(k[1])):
                    return ""
                y = k[0]
                w = k[1]
                y_txt = f", {int(pd.to_numeric(y, errors='coerce'))}" if pd.notna(y) else ""
                w_txt = f" (Wk {int(pd.to_numeric(w, errors='coerce'))})" if pd.notna(w) else ""
                return f"{w_txt}{y_txt}"
            range_txt = ""
            if 'best_streak_start_key' in locals() and 'best_streak_end_key' in locals():
                start_txt = _fmt_key(best_streak_start_key)
                end_txt = _fmt_key(best_streak_end_key)
                if start_txt or end_txt:
                    range_txt = f" — {start_txt} to {end_txt}"
            st.markdown(
                f"""
                <div class='record-card'>
                    <div class='record-title'><span class='pill'>Longest #1 streak</span></div>
                    <div class='record-stat'>{int(best_streak_len)}</div>
                    <div class='record-sub'>{best_streak_owner}{range_txt}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        # Biggest Elo gain (single match)
        if max_gain_owner is not None and pd.notna(max_gain_val):
            y_txt = f", {int(pd.to_numeric(max_gain_yw[0], errors='coerce'))}" if max_gain_yw and max_gain_yw[0] is not None else ""
            w_txt = f" (Wk {int(pd.to_numeric(max_gain_yw[1], errors='coerce'))})" if max_gain_yw and max_gain_yw[1] is not None else ""
            st.markdown(
                f"""
                <div class='record-card'>
                    <div class='record-title'><span class='pill'>Biggest win Elo change</span></div>
                    <div class='record-stat'>+{abs(float(max_gain_val)):.2f}</div>
                    <div class='record-sub'>{max_gain_owner}{w_txt}{y_txt}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        # Biggest Elo loss (single match)
        if min_loss_owner is not None and pd.notna(min_loss_val):
            y_txt = f", {int(pd.to_numeric(min_loss_yw[0], errors='coerce'))}" if min_loss_yw and min_loss_yw[0] is not None else ""
            w_txt = f" (Wk {int(pd.to_numeric(min_loss_yw[1], errors='coerce'))})" if min_loss_yw and min_loss_yw[1] is not None else ""
            st.markdown(
                f"""
                <div class='record-card'>
                    <div class='record-title'><span class='pill'>Biggest loss Elo change</span></div>
                    <div class='record-stat'>-{abs(float(min_loss_val)):.2f}</div>
                    <div class='record-sub'>{min_loss_owner}{w_txt}{y_txt}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        st.markdown('</div>', unsafe_allow_html=True)
    except Exception:
        pass

    # Match details (compact)
    st.markdown("### Match rating details")
    details = pd.DataFrame(rows)
    # Safer ordering
    if {"Year", "Week"}.issubset(details.columns):
        details["_Y"] = pd.to_numeric(details["Year"], errors="coerce")
        details["_W"] = pd.to_numeric(details["Week"], errors="coerce")
        details = details.sort_values(["_Y", "_W"]).drop(columns=["_Y", "_W"], errors="ignore")
    st.dataframe(details, use_container_width=True)


def main():
    st.set_page_config(
        page_title="The Big Tebowski – League History",
        page_icon="🏈",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    _style_light_ui()

    st.markdown("""
    # The Big Tebowski – League History
    <span class="small-muted">Explore championships, seasons, drafts, head-to-heads, and more from your league's history.</span>
    """, unsafe_allow_html=True)

    # Ensure the data file exists in deployments
    _ensure_data_file(DEFAULT_FILE_PATH)
    try:
        size = os.path.getsize(DEFAULT_FILE_PATH) if os.path.exists(DEFAULT_FILE_PATH) else 0
        valid_zip = zipfile.is_zipfile(DEFAULT_FILE_PATH) if size else False
        st.caption(f"Data file check: size={size} bytes, zip={valid_zip}")
        # List workbook sheets to verify content in Cloud
        if size and valid_zip:
            try:
                xls = pd.ExcelFile(DEFAULT_FILE_PATH, engine="openpyxl")
                st.caption("Workbook sheets: " + ", ".join(xls.sheet_names))
            except Exception as e:
                st.caption(f"Workbook open error: {e}")
    except Exception:
        pass

    # Compute a file signature to bust cache when the Excel changes
    _sig = _file_signature(DEFAULT_FILE_PATH)
    df_ch = load_sheet(DEFAULT_FILE_PATH, "championship_games", file_sig=_sig)
    df_to = load_sheet(DEFAULT_FILE_PATH, "teams_owners", file_sig=_sig)
    df_reg = load_sheet(DEFAULT_FILE_PATH, "reg_season_tables", file_sig=_sig)
    df_draft = load_sheet(DEFAULT_FILE_PATH, "draft", file_sig=_sig)
    df_gl = load_sheet(DEFAULT_FILE_PATH, "gamelog", file_sig=_sig)
    df_records = load_sheet(DEFAULT_FILE_PATH, "records", file_sig=_sig)
    # Small status to confirm data presence in deployments
    try:
        def _n(df):
            return 0 if df is None or getattr(df, 'empty', True) else len(df)
        st.caption(
            f"Data file: '{os.path.basename(DEFAULT_FILE_PATH)}' — sheets loaded: "
            f"ch={_n(df_ch)}, gl={_n(df_gl)}, reg={_n(df_reg)}, draft={_n(df_draft)}, to={_n(df_to)}"
        )
    except Exception:
        pass

    selected_years, selected_teams, selected_owners, file_path = sidebar_filters(
        [df_ch, df_to, df_reg, df_draft, df_gl], teams_owners=df_to
    )

    # Always use bundled path; already loaded above.

    if not file_path or not os.path.exists(file_path):
        exp = os.path.join(BASE_DIR, "fantasy_football.xlsm")
        st.error(
            f"Excel file not found in app directory. Expected: '{exp}'. "
            "We attempted to download the file; if this persists, set DATA_URL in Streamlit secrets or environment."
        )
        st.stop()

    tabs = st.tabs([
        "Overview",
        "Championships & Toilet Bowl",
        "Regular Season",
    "Rating",
    "Records",
        "Draft (Round 1)",
        "Head-to-Head & Game Log",
        "Teams & Owners",
    ])

    with tabs[0]:
        render_overview(df_ch, df_gl, df_reg, df_to, selected_years, selected_teams, selected_owners)
    with tabs[1]:
        render_championships(df_ch, selected_years, selected_teams, selected_owners)
    with tabs[2]:
        render_regular_season(df_reg, selected_years, selected_teams, selected_owners)
    with tabs[3]:
        render_rating(df_gl, selected_years, selected_teams, selected_owners)
    with tabs[4]:
        render_records(df_gl, df_reg, selected_years, selected_teams, selected_owners)
    with tabs[5]:
        render_draft(df_draft, df_to, df_reg, selected_years, selected_teams, selected_owners, file_path)
    with tabs[6]:
        render_head_to_head(df_gl, selected_years, selected_teams, selected_owners)
    with tabs[7]:
        render_teams_owners(df_to, selected_years, selected_teams, selected_owners)

    if df_records is not None and not df_records.empty:
        st.markdown("---")
        st.markdown("### Records (from sheet)")
        st.dataframe(df_records, use_container_width=True)


if __name__ == "__main__":
    main()
