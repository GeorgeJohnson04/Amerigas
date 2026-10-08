@echo off
setlocal
cd /d "%~dp0"

rem ===================================================================
rem  update_tracker.bat - daily refresh of the hedge product tracker
rem
rem  1. Put the new ArgusFC_YYYYMMDD.csv file(s) in the LPG folder named
rem     in hedge_products.json (default: data\LPG), or refresh the
rem     price parquet with build_price_index.py.
rem  2. Double-click this file.
rem
rem  Any Argus file newer than the parquet is added to the parquet
rem  (a .bak copy is kept) and "Hedge Product Tracker.xlsx" is rebuilt.
rem  Close the workbook in Excel first, otherwise a time-stamped copy
rem  is written instead.
rem
rem  Pass NOOPEN to skip opening the workbook at the end (for scheduled runs):
rem      update_tracker.bat NOOPEN
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
    echo  then run:   py -m pip install pandas pyarrow openpyxl numpy
    echo.
    pause
    exit /b 1
)

%PY% -c "import pandas, pyarrow, openpyxl, numpy" >nul 2>&1
if errorlevel 1 (
    echo  Installing missing packages...
    %PY% -m pip install pandas pyarrow openpyxl numpy
    %PY% -c "import pandas, pyarrow, openpyxl, numpy" >nul 2>&1
    if errorlevel 1 (
        echo  Package install failed. Run by hand:  %PY% -m pip install pandas pyarrow openpyxl numpy
        pause
        exit /b 1
    )
)

%PY% hedge_tracker.py
set "RC=%ERRORLEVEL%"
echo.
if not "%RC%"=="0" (
    echo  Update FAILED with exit code %RC%. See the messages above.
    if /i not "%~1"=="NOOPEN" pause
    exit /b %RC%
)

if /i "%~1"=="NOOPEN" exit /b 0
if exist "Hedge Product Tracker.xlsx" start "" "Hedge Product Tracker.xlsx"
pause
exit /b 0
