# WSR Financials Export

A small standalone CLI that exports **Wall $treet Raider** balance sheets and cash flow
statements to Excel, read live from the game's local bridge. It modifies no game files.

## Requirements
- Python 3 with `requests` (`pip install requests`).
- For a single multi-sheet **`.xlsx`** workbook: `pip install openpyxl`. Without it, the tool
  writes one **CSV** per entity instead (still opens in Excel).
- The game must be **running** — it publishes its bridge port to `runtime.json` while open.

## Usage
```
python3 wsr_financials.py --out-dir ./statements
```
By default it exports **you (the player) plus every company you control**, into a dated file
`wsr_financials_<YYYY>-<MM>.xlsx` — a new file each game-month, overwritten if you re-run in the
same month, so a monthly history accumulates.

### Options
| flag | effect |
|---|---|
| `--symbols RELI,IMD` | just these tickers (instead of your empire) |
| `--all` | you + every company in the game (slow — one read each) |
| `--controlled-only` | only the companies you control (no personal sheet) |
| `--player-only` | only your personal statement |
| `--overwrite` | one rolling `wsr_financials.xlsx` instead of dated files |
| `--settle 0.3` | per-company read delay (seconds) |
| `--runtime PATH` | path to the game's `runtime.json` (auto-found by default) |
| `--port N` | bridge REST port directly (bypass `runtime.json`) |

## Output
Each workbook has a **Summary** sheet plus **one sheet per entity**, each containing a Balance
Sheet (assets / liabilities / equity, with an `Assets - Liabilities - Equity` check row that
reconciles to ~0) and a Cash Flow statement. For companies, total liabilities include reserves
and deposits, not just interest-bearing debt (which is shown separately). All figures are in
**$ millions**.

If the game's bridge isn't found automatically, run with `--runtime` pointing at the game's
`runtime.json`, or `--port` with the REST port the engine is listening on.
