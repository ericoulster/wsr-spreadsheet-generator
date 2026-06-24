@echo off
REM ---------------------------------------------------------------------------
REM Double-click to export your Wall Street Raider statements (FROM SOURCE).
REM Needs Python installed and the game running. For a no-Python option, build
REM the .exe with build.bat instead, or get a prebuilt wsr_financials.exe.
REM ---------------------------------------------------------------------------
python -m pip install --quiet requests openpyxl
python "%~dp0wsr_financials.py" %*
