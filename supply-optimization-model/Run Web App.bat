@echo off
setlocal
cd /d "%~dp0"
title Propane Model Web App

rem ===================================================================
rem  Run Web App.bat - double-click to open the propane model web app.
rem
rem  Starts the app on this PC and opens it in your browser at
rem  http://127.0.0.1:8050 as soon as it is ready. If the app is already
rem  running, it just opens the page again.
rem
rem  Keep this window open while you use the app - it also shows what the
rem  app is doing, which helps when debugging. Close it to stop the app.
rem
rem  Local only: nothing else on the network can reach it, and there is
rem  no login yet. See webapp\app.py before sharing it with anyone.
rem ===================================================================

set "PY="
for %%C in (py python python3) do (
    if not defined PY (
        %%C -c "import sys" >nul 2>&1 && set "PY=%%C"
    )
)
if not defined PY (
    echo.
    echo  ERROR: Python was not found. Install it from https://www.python.org/downloads/
    echo  then run:   py -m pip install flask==3.1.3 openpyxl scipy numpy
    echo.
    pause
    exit /b 1
)

%PY% -c "import flask, openpyxl, numpy, scipy" >nul 2>&1
if errorlevel 1 (
    echo  Installing the packages the web app needs...
    %PY% -m pip install "flask==3.1.3" openpyxl scipy numpy
    %PY% -c "import flask, openpyxl, numpy, scipy" >nul 2>&1
    if errorlevel 1 (
        echo  Install did not succeed. Run this by hand and check the errors:
        echo      %PY% -m pip install flask==3.1.3 openpyxl scipy numpy
        pause
        exit /b 1
    )
)

rem ---- already running? open the page (restarting it first if it is an ----
rem ---- out-of-date copy left over from before an update)              ----
%PY% webapp\open_browser.py --check >nul 2>&1
if not errorlevel 1 (
    %PY% webapp\open_browser.py
    timeout /t 5 >nul
    exit /b 0
)

echo.
echo  ============================================================
echo    Propane Model Web App
echo    http://127.0.0.1:8050
echo.
echo    Your browser opens by itself once the app is ready.
echo    Keep this window open while you work. Close it to stop.
echo  ============================================================
echo.
rem  The helper waits for the app to answer, then opens the browser (with
rem  fallbacks if the normal way is blocked). It runs in its own minimised
rem  window so its messages don't get tangled with the app's log below.
start "Opening browser" /min %PY% webapp\open_browser.py --wait
%PY% webapp\app.py
echo.
echo  The web app has stopped.
pause
