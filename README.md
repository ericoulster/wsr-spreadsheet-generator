# WSR Spreadsheet Generator

A small standalone tool that exports **Wall $treet Raider** balance sheets and cash flow
statements to Excel, read live from the game's local bridge while you play. It modifies no game
files (read-only).

By default it exports **you (the player) plus every company you control**, into a folder called
`WSR_Statements` next to the program. Company balance sheets reconcile exactly (an
`Assets - Liabilities - Equity` check row that lands at 0).

---

## Quick start - no setup (recommended for most people)

1. Get **`wsr_financials.exe`** - a single file, no Python and no install. It isn't checked into
   the repo; someone builds it once (see [Build the .exe](#build-the-exe-yourself-to-share-with-non-developers))
   and shares it, or you download it from the repo's GitHub **Actions** artifacts / a Release.
2. Start **Wall $treet Raider** and load your game.
3. **Double-click `wsr_financials.exe`.** A small window opens, writes your statements, and says
   where they went.
4. Open the `.xlsx` it created in the **`WSR_Statements`** folder next to the .exe.

That's it. Run it again whenever you want a fresh snapshot - by default each game-month gets its
own file, so a monthly history builds up on its own.

### Want it to export automatically every game-month?

Run it in **watch mode**: it stays open and writes a new file each time the game's month advances.
Make a shortcut to the .exe and add ` --watch` to the Target, or from a command prompt:

```
wsr_financials.exe --watch
```

Leave it running alongside the game; stop it with Ctrl-C or by closing the window.

---

## Run from source

**With [uv](https://docs.astral.sh/uv/) (simplest - nothing to install first).** The script
declares its own dependencies inline (PEP 723), so uv fetches them into an isolated environment and
runs it - no venv, no `pip install`:

```
uv run wsr_financials.py
uv run wsr_financials.py --watch
uv run wsr_financials.py --symbols RELI,IMD
```

**With pip.** Python 3 plus `requests` and `openpyxl`:

```
pip install requests openpyxl
python wsr_financials.py
```

(On Windows you can instead double-click **`run.bat`**, which installs those two packages and runs
it.) Without `openpyxl` it still works - it writes one **CSV** per entity instead of a single
`.xlsx`.

### Options

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

## Build the .exe yourself (to share with non-developers)

The `.exe` is what makes this usable by people who don't have Python. Two ways to produce it:

- **On a Windows machine with Python:** double-click **`build.bat`** (or run it from a prompt). It
  installs PyInstaller and produces **`dist\wsr_financials.exe`**. Share that one file.
- **Without a Windows machine:** run the **build-exe** GitHub Actions workflow (Actions tab ->
  Run workflow, or push a `v*` tag). It builds the `.exe` on a Windows runner; download it from the
  run's Artifacts. See `.github/workflows/build-exe.yml`.

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
