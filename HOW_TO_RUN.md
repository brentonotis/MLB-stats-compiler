# MLB Stats Compiler — Setup & Usage

## What This Does
Pulls **10 years of batting and pitching stats (2016–2025)** for every player currently on an MLB 2026 40-man roster and writes them to a clean, formatted Excel workbook with two tabs: **Batting** and **Pitching**.

All data sources are **100% free** — no API keys, no subscriptions.

## Stats Included

**Batting tab:** Player, Team, Season, G, PA, AB, H, R, 2B, 3B, HR, RBI, SB, AVG, OBP, SLG, OPS

**Pitching tab:** Player, Team, Season, W, L, SV, G, GS, IP, K, QS, ERA, WHIP, K/9, BB/9, FIP

## Setup (one-time, ~2 minutes)

### 1. Install Python
If you don't already have Python 3.8+, download it from [python.org](https://www.python.org/downloads/). During installation on Windows, **check "Add Python to PATH"**.

### 2. Install dependencies
Open Terminal (Mac) or Command Prompt (Windows) and run:

```
pip install pybaseball openpyxl requests
```

## Running the Script (~3–5 minutes)

### 3. Run it
Navigate to the folder where you saved `mlb_stats_compiler.py` and run:

```
python mlb_stats_compiler.py
```

The script will:
1. Fetch all 30 teams' 2026 40-man rosters from the MLB Stats API
2. Pull batting stats for all MLB players, 2016–2025, from FanGraphs
3. Pull pitching stats for all MLB players, 2016–2025, from FanGraphs
4. Filter to only include players on 2026 rosters
5. Write a formatted Excel file: `MLB_Player_Stats_2016_2025.xlsx`

## Output
The Excel file will appear in the same folder as the script. It has:
- **Batting** tab — one row per player per season, sorted alphabetically
- **Pitching** tab — one row per player per season, sorted alphabetically
- Frozen header rows, auto-filters, and alternating row shading

## Troubleshooting
- **"ModuleNotFoundError: No module named 'pybaseball'"** → Run `pip install pybaseball` again
- **Script seems slow** → FanGraphs data pulls can take 1–3 minutes each; this is normal
- **Missing players** → Name matching is fuzzy but some edge cases (accented names, traded players) may slip through. The script handles most accent variations automatically.
- **Want more/fewer years?** → Edit `START_YEAR` and `END_YEAR` at the top of the script
