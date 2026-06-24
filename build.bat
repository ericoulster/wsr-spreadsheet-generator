@echo off
REM ---------------------------------------------------------------------------
REM Build a standalone wsr_financials.exe that needs NO Python to RUN.
REM You only need Python on THIS machine to build it. Result: dist\wsr_financials.exe
REM Double-click this file on Windows, or run it from a command prompt.
REM ---------------------------------------------------------------------------
echo Installing build dependencies (pyinstaller, requests, openpyxl)...
python -m pip install --upgrade pyinstaller requests openpyxl || goto :err
echo.
echo Building wsr_financials.exe ...
python -m PyInstaller --onefile --name wsr_financials wsr_financials.py || goto :err
echo.
echo Done. Your single-file program is:  dist\wsr_financials.exe
echo Share that one .exe - users just double-click it while the game is running.
pause
exit /b 0
:err
echo.
echo Build failed. Make sure Python is installed and on your PATH (python.org).
pause
exit /b 1
