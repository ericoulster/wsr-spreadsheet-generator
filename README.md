# WSR Spreadsheet Generator

A small standalone command-line tool that exports **Wall $treet Raider** balance sheets and cash
flow statements to Excel, read live from the game's local bridge while you play. It modifies no game
files (read-only).

By default it exports **you (the player) plus every company you control**, into a folder called
`WSR_Statements` next to the script. Company balance sheets reconcile exactly (an
`Assets - Liabilities - Equity` check row that lands at 0).

---

## Quick start

The game must be running with a save loaded.

**With [uv](https://docs.astral.sh/uv/) (simplest - nothing to install first).** The script declares
its own dependencies inline (PEP 723), so uv fetches them into an isolated environment and runs it -
no venv, no `pip install`:

```
uv run wsr_financials.py
```

**With pip.** Python 3 plus `requests` and `openpyxl`:

```
pip install requests openpyxl
python wsr_financials.py
```

On Windows you can instead double-click **`run.bat`**, which installs those two packages and runs the
script. Without `openpyxl` it still works - it writes one **CSV** per entity instead of a single
`.xlsx`.

Open the workbook it writes into the **`WSR_Statements`** folder. By default each game-month gets its
own dated file, so a monthly history builds up as you re-run it.

### Auto-export every game-month

Run it in **watch mode**: it stays open and writes a new file each time the game's month advances.

```
uv run wsr_financials.py --watch
```

Stop it with Ctrl-C.

---

## Options

| flag | effect |
|---|---|
| *(no flags)* | you + every company you control, dated file, into `WSR_Statements/` |
| `--out-dir FOLDER` | write somewhere else (created if missing) |
| `--symbols RELI,IMD` | just these tickers (instead of your empire) |
| `--all` | you + every company in the game (slow - one read each) |
| `--controlled-only` | only the companies you control (no personal sheet) |
| `--player-only` | only your personal statement |
| `--watch` | stay open and auto-export each time the game month advances |
| `--interval 30` | how often `--watch` checks, in seconds (default 30) |
| `--overwrite` | one rolling `wsr_financials.xlsx` instead of dated files |
| `--settle 0.3` | per-company read delay (seconds) |
| `--no-pause` | don't wait for a keypress when finished |
| `--runtime PATH` | path to the game's `runtime.json` (auto-found by default) |
| `--port N` | bridge REST port directly (bypass `runtime.json`) |

---

## Output

Each workbook has a **Summary** sheet plus **one sheet per entity**, each containing a Balance
Sheet (assets / liabilities / equity, with an `Assets - Liabilities - Equity` check row that
reconciles to ~0) and a Cash Flow statement. For companies, total liabilities include reserves and
deposits, not just interest-bearing debt (which is shown separately). All figures are in
**$ millions**.

## Troubleshooting

- **"Make sure Wall Street Raider is running"** - the tool talks to the game over a local port the
  game only opens while it's running. Start the game, load a save, then run the tool.
- **Bridge not found** - if auto-detection fails, pass `--runtime` pointing at the game's
  `runtime.json`, or `--port` with the REST port the engine is listening on.

## License

[MIT](LICENSE) - use, copy, modify, and redistribute freely (including in your own mods or tools),
as long as you keep the copyright and license notice (attribution).

## Disclaimer

A fan-made, unofficial utility. Not affiliated with or endorsed by the makers of Wall $treet
Raider. It only reads the running game and writes spreadsheets; it changes nothing in the game.
