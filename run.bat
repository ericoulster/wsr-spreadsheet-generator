@echo off
REM ---------------------------------------------------------------------------
REM Double-click to export your Wall Street Raider statements.
REM Needs Python installed and the game running.
REM ---------------------------------------------------------------------------
python -m pip install --quiet requests openpyxl
python "%~dp0wsr_financials.py" %*
