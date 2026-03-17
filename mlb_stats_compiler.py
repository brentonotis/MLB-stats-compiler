"""
MLB Stats Compiler — 10-Year Historical Stats for 2026 Roster Players
=====================================================================
Pulls batting and pitching stats (2016–2025) for every player on a
2026 MLB 40-man roster and writes them to a formatted Excel workbook.

Data sources (all free, no API key required):
  • FanGraphs via pybaseball  — batting & pitching season stats
  • MLB Stats API             — 2026 40-man rosters

Requirements:
  pip install pybaseball openpyxl requests

Usage:
  python mlb_stats_compiler.py
"""

import sys
import json
import time
import requests
import warnings
import pandas as pd
from pathlib import Path

warnings.filterwarnings("ignore")

# ── Attempt pybaseball import ────────────────────────────────────────
try:
    from pybaseball import batting_stats, pitching_stats
    from pybaseball import cache
    cache.enable()
except ImportError:
    print("ERROR: pybaseball is not installed.")
    print("Run:  pip install pybaseball")
    sys.exit(1)

OUTPUT_FILE = "MLB_Player_Stats_2016_2025.xlsx"
START_YEAR = 2016
END_YEAR = 2025
ROSTER_SEASON = 2026

# ── Column mappings (FanGraphs column → friendly name) ───────────────
BATTING_COLS_MAP = {
    "Name": "Player",
    "Team": "Team",
    "Season": "Season",
    "G": "G",
    "PA": "PA",
    "AB": "AB",
    "R": "R",
    "2B": "2B",
    "3B": "3B",
    "HR": "HR",
    "RBI": "RBI",
    "SB": "SB",
    "AVG": "AVG",
    "H": "H",
    "BB": "BB",
    "SO": "SO",
    "OBP": "OBP",
    "SLG": "SLG",
    "OPS": "OPS",
}

PITCHING_COLS_MAP = {
    "Name": "Player",
    "Team": "Team",
    "Season": "Season",
    "W": "W",
    "L": "L",
    "SV": "SV",
    "G": "G",
    "GS": "GS",
    "IP": "IP",
    "SO": "K",
    "QS": "QS",
    "ERA": "ERA",
    "WHIP": "WHIP",
    "K/9": "K/9",
    "BB/9": "BB/9",
    "HR/9": "HR/9",
    "H": "H",
    "BB": "BB",
    "ER": "ER",
    "FIP": "FIP",
}

# Final columns the user requested (plus a few extras for context)
BATTING_FINAL = ["Player", "Team", "Season", "G", "PA", "AB", "H",
                 "R", "2B", "3B", "HR", "RBI", "SB", "AVG",
                 "OBP", "SLG", "OPS"]

PITCHING_FINAL = ["Player", "Team", "Season", "W", "L", "SV",
                  "G", "GS", "IP", "K", "QS", "ERA", "WHIP",
                  "K/9", "BB/9", "FIP"]


def get_2026_rosters():
    """Fetch 40-man rosters for all 30 MLB teams from the Stats API."""
    print("Fetching 2026 40-man rosters from MLB Stats API...")
    teams_url = f"https://statsapi.mlb.com/api/v1/teams?sportId=1&season={ROSTER_SEASON}"
    resp = requests.get(teams_url, timeout=30)
    resp.raise_for_status()
    teams = resp.json().get("teams", [])
    print(f"  Found {len(teams)} teams")

    roster_players = []
    for team in teams:
        tid = team["id"]
        tname = team.get("abbreviation", team["name"])
        url = f"https://statsapi.mlb.com/api/v1/teams/{tid}/roster?rosterType=40Man&season={ROSTER_SEASON}"
        try:
            r = requests.get(url, timeout=20)
            r.raise_for_status()
            for p in r.json().get("roster", []):
                person = p.get("person", {})
                roster_players.append({
                    "mlb_id": person.get("id"),
                    "name": person.get("fullName"),
                    "team": tname,
                    "position": p.get("position", {}).get("abbreviation", ""),
                })
        except Exception as e:
            print(f"  ⚠ Could not fetch roster for {tname}: {e}")
        time.sleep(0.3)          # be polite to the API

    print(f"  Total 40-man roster players: {len(roster_players)}")
    return pd.DataFrame(roster_players)


def pull_batting(start=START_YEAR, end=END_YEAR):
    """Pull season-level batting stats from FanGraphs via pybaseball."""
    print(f"Pulling batting stats {start}–{end} from FanGraphs (this may take a minute)...")
    df = batting_stats(start, end, qual=1, ind=1)    # qual=1 → min 1 PA; ind=1 → individual seasons
    print(f"  Raw batting rows: {len(df)}")
    return df


def pull_pitching(start=START_YEAR, end=END_YEAR):
    """Pull season-level pitching stats from FanGraphs via pybaseball."""
    print(f"Pulling pitching stats {start}–{end} from FanGraphs (this may take a minute)...")
    df = pitching_stats(start, end, qual=1, ind=1)
    print(f"  Raw pitching rows: {len(df)}")
    return df


def normalize_name(name: str) -> str:
    """Lowercase, strip accents & suffixes for fuzzy matching."""
    import unicodedata
    n = unicodedata.normalize("NFD", name)
    n = "".join(c for c in n if unicodedata.category(c) != "Mn")
    for suffix in [" jr.", " sr.", " ii", " iii", " iv"]:
        n = n.replace(suffix, "")
    return n.strip().lower()


def filter_to_roster(stats_df, roster_df):
    """Keep only rows whose player name matches a 2026 roster name."""
    roster_names = set(roster_df["name"].apply(normalize_name))
    mask = stats_df["Player"].apply(normalize_name).isin(roster_names)
    filtered = stats_df[mask].copy()
    return filtered


def write_excel(bat_df, pit_df, path):
    """Write a formatted .xlsx with Batting and Pitching tabs."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
    from openpyxl.utils import get_column_letter

    wb = Workbook()

    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    data_font = Font(name="Arial", size=10)
    data_align_c = Alignment(horizontal="center")
    data_align_l = Alignment(horizontal="left")
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    stripe_fill = PatternFill("solid", fgColor="F2F7FB")

    def write_sheet(ws, df, sheet_name, pct_cols=None, dec_cols=None):
        ws.title = sheet_name
        pct_cols = pct_cols or []
        dec_cols = dec_cols or []
        cols = list(df.columns)

        # Header row
        for c_idx, col in enumerate(cols, 1):
            cell = ws.cell(row=1, column=c_idx, value=col)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = thin_border

        # Data rows
        for r_idx, row in enumerate(df.itertuples(index=False), 2):
            for c_idx, (col, val) in enumerate(zip(cols, row), 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.font = data_font
                cell.border = thin_border
                if col in ("Player",):
                    cell.alignment = data_align_l
                else:
                    cell.alignment = data_align_c
                if col in pct_cols:
                    cell.number_format = "0.000"
                elif col in dec_cols:
                    cell.number_format = "0.00"
                elif isinstance(val, (int,)):
                    cell.number_format = "#,##0"
                # Alternate row shading
                if r_idx % 2 == 0:
                    cell.fill = stripe_fill

        # Auto-fit column widths
        for c_idx, col in enumerate(cols, 1):
            max_len = max(len(str(col)), *(len(str(ws.cell(row=r, column=c_idx).value or ""))
                                           for r in range(2, min(ws.max_row + 1, 50))))
            ws.column_dimensions[get_column_letter(c_idx)].width = min(max_len + 3, 18)

        # Freeze header row
        ws.freeze_panes = "A2"
        # Auto-filter
        ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{ws.max_row}"

    # ── Batting tab ──────────────────────────────────────────────────
    ws_bat = wb.active
    write_sheet(ws_bat, bat_df, "Batting",
                pct_cols=["AVG", "OBP", "SLG", "OPS"])

    # ── Pitching tab ─────────────────────────────────────────────────
    ws_pit = wb.create_sheet()
    write_sheet(ws_pit, pit_df, "Pitching",
                dec_cols=["ERA", "WHIP", "K/9", "BB/9", "FIP", "IP"])

    wb.save(path)
    print(f"\n✅  Saved to {path}")
    print(f"    Batting tab:  {len(bat_df):,} player-seasons")
    print(f"    Pitching tab: {len(pit_df):,} player-seasons")


# ═════════════════════════════════════════════════════════════════════
#  MAIN
# ═════════════════════════════════════════════════════════════════════
def main():
    print("=" * 60)
    print("  MLB Stats Compiler — 10-Year Stats for 2026 Rosters")
    print("=" * 60)

    # 1. Rosters
    roster = get_2026_rosters()
    if roster.empty:
        print("ERROR: Could not retrieve any roster data.")
        sys.exit(1)

    # 2. Batting stats
    bat_raw = pull_batting()
    bat_cols_available = [c for c in BATTING_COLS_MAP if c in bat_raw.columns]
    bat = bat_raw[bat_cols_available].rename(columns=BATTING_COLS_MAP).copy()
    bat_final_avail = [c for c in BATTING_FINAL if c in bat.columns]
    bat = bat[bat_final_avail]

    # 3. Pitching stats
    pit_raw = pull_pitching()
    pit_cols_available = [c for c in PITCHING_COLS_MAP if c in pit_raw.columns]
    pit = pit_raw[pit_cols_available].rename(columns=PITCHING_COLS_MAP).copy()
    pit_final_avail = [c for c in PITCHING_FINAL if c in pit.columns]
    pit = pit[pit_final_avail]

    # 4. Filter to 2026 roster players
    print("\nFiltering to 2026 roster players...")
    bat_filtered = filter_to_roster(bat, roster)
    pit_filtered = filter_to_roster(pit, roster)
    print(f"  Batting:  {len(bat)} → {len(bat_filtered)} player-seasons")
    print(f"  Pitching: {len(pit)} → {len(pit_filtered)} player-seasons")

    # 5. Sort
    bat_filtered = bat_filtered.sort_values(["Player", "Season"]).reset_index(drop=True)
    pit_filtered = pit_filtered.sort_values(["Player", "Season"]).reset_index(drop=True)

    # 6. Write Excel
    write_excel(bat_filtered, pit_filtered, OUTPUT_FILE)

    print("\nDone! Open the file in Excel or Google Sheets.")


if __name__ == "__main__":
    main()
