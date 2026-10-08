@echo off
setlocal
cd /d "%~dp0"

rem ===================================================================
rem  Run Model.bat - launcher for model.py
rem
rem  Double-click it for a menu, or pass the region(s) straight in:
rem
rem      "Run Model.bat" Rail
rem      "Run Model.bat" Rail,Rockies,PNW
rem      "Run Model.bat" "New England"
rem      "Run Model.bat" ALL
rem
rem  Separate several regions with COMMAS, not spaces - some region
rem  names contain a space ("New England", "North East").
rem
rem  To give a manager their own one-click file, make a Windows shortcut
rem  to this .bat and add their region(s) to the end of the Target box.
rem
rem  Manual entry: if overrides.txt sits next to this file it is applied
rem  automatically. Use the OVERRIDES menu choice to create or edit it and
rem  CHECK to validate it without solving.
rem
rem  Exit codes mirror model.py:
rem      0  every area solved cleanly
rem      1  no workbooks found / bad setup / named override file missing
rem      2  bad command line (unknown region or group)
rem      3  at least one area could not be solved at all
rem      4  solved, but some areas needed unmet-demand / surplus
rem         allowances because their inputs are over-committed
rem      5  the override file could not be parsed, or one of its lines
rem         matched nothing - run CHECK to see exactly which line
rem ===================================================================

set "RC=0"
set "DECKNAME=overrides.txt"
set "DECK=%~dp0%DECKNAME%"

rem ---- locate a usable Python -------------------------------------------
rem  "py" (the launcher) is tried first because a plain "python" on
rem  Windows is often the Microsoft Store stub, which is not a real
rem  interpreter and exits non-zero.
set "PY="
for %%C in (py python python3) do (
    if not defined PY (
        %%C -c "import sys" >nul 2>&1 && set "PY=%%C"
    )
)

if not defined PY (
    echo.
    echo  ERROR: Python was not found on this machine.
    echo.
    echo  Install it from https://www.python.org/downloads/  ^(tick
    echo  "Add python.exe to PATH" during setup^), then install the
    echo  packages the model needs:
    echo.
    echo      py -m pip install openpyxl scipy numpy
    echo.
    pause
    exit /b 1
)

rem ---- confirm the third-party packages are importable -------------------
rem  Without this the user gets a bare traceback on their first run.
%PY% -c "import openpyxl, numpy, scipy" >nul 2>&1
if errorlevel 1 (
    echo.
    echo  Python was found ^(%PY%^) but the packages the model needs are
    echo  missing. Installing them now...
    echo.
    %PY% -m pip install openpyxl scipy numpy
    echo.
    %PY% -c "import openpyxl, numpy, scipy" >nul 2>&1
    if errorlevel 1 (
        echo  Install did not succeed. Run this by hand and check the errors:
        echo      %PY% -m pip install openpyxl scipy numpy
        echo.
        pause
        exit /b 1
    )
    echo  Packages installed.
    echo.
)

rem ---- arguments supplied? skip the menu --------------------------------
rem  Strip any quotes the caller wrapped around a multi-word region so we
rem  do not end up passing --region ""New England"" to Python.
if not "%~1"=="" (
    set "CHOICE=%*"
    call set CHOICE=%%CHOICE:"=%%
    goto :dispatch
)

:menu
cls
echo ============================================================
echo   Propane Supply Model
echo ============================================================
echo.
%PY% model.py --list-regions
echo.
echo  Type the region^(s^) you want, separated by COMMAS. A saved group
echo  name works here too. Spelling is forgiving - "north east" finds
echo  Northeast.
echo.
echo    ALL        run every region
echo    STRICT     run every region without the automatic relaxations
echo    GROUPS     show the saved manager groups
echo    LIST       show which workbook each area would use ^(no solve^)
echo    OVERRIDES  create or edit the manual-entry file
echo    CHECK      validate the manual-entry file without solving
echo    Q          quit
echo.
if exist "%DECK%" (
    echo    Manual entries ACTIVE: %DECKNAME% will be applied.
) else (
    echo    No %DECKNAME% - running straight from the workbooks.
)
echo.
set "CHOICE="
set /p "CHOICE=Run which regions? "
if not defined CHOICE goto :menu
call set CHOICE=%%CHOICE:"=%%

:dispatch
if /i "%CHOICE%"=="Q"      goto :done
if /i "%CHOICE%"=="QUIT"   goto :done
if /i "%CHOICE%"=="EXIT"   goto :done

rem  The two information-only choices jump to labels rather than running
rem  inside an if(...) block. Inside a block %ERRORLEVEL% would be expanded
rem  when the block is parsed - before the command has actually run - so the
rem  captured exit code would be the previous one.
if /i "%CHOICE%"=="GROUPS"    goto :do_groups
if /i "%CHOICE%"=="LIST"      goto :do_list
if /i "%CHOICE%"=="OVERRIDES" goto :do_overrides
if /i "%CHOICE%"=="CHECK"     goto :do_check
goto :do_run

:do_groups
echo.
%PY% model.py --list-groups
set "RC=%ERRORLEVEL%"
echo.
rem  Loop back to the menu for a human, but exit straight away when called
rem  with an argument so a scripted run cannot hang on a pause.
if "%~1"=="" (
    pause
    goto :menu
)
goto :done

:do_list
echo.
%PY% model.py --list
set "RC=%ERRORLEVEL%"
echo.
if "%~1"=="" (
    pause
    goto :menu
)
goto :done

:do_overrides
rem  Create the deck from a template the first time, so the user starts from a
rem  file that already lists every real location spelled correctly, then open
rem  it and wait for them to finish editing.
echo.
if not exist "%DECK%" (
    echo  No %DECKNAME% yet - building one from the workbooks.
    echo  This reads every area, so give it a moment.
    echo.
    %PY% model.py --template-overrides -o "%DECK%"
    if errorlevel 1 (
        echo.
        echo  Could not build the template. See the messages above.
        echo.
        if "%~1"=="" pause
        goto :menu
    )
    echo.
    echo  Created %DECKNAME%. Everything in it is commented out, so it changes
    echo  nothing until you uncomment a line.
) else (
    echo  Opening the existing %DECKNAME%.
)
echo.
echo  Notepad will open. Save and close it when you are done.
echo.
start /wait notepad "%DECK%"
echo  Validating what you saved...
echo.
%PY% model.py --check-overrides "%DECK%"
set "RC=%ERRORLEVEL%"
echo.
if "%RC%"=="0" (
    echo  The deck is valid. Pick a region to run it.
) else (
    echo  The deck has a problem - see above. Choose OVERRIDES again to fix it.
)
echo.
if "%~1"=="" (
    pause
    goto :menu
)
goto :done

:do_check
echo.
if not exist "%DECK%" (
    echo  There is no %DECKNAME% to check yet.
    echo  Choose OVERRIDES to create one.
    echo.
    set "RC=1"
    if "%~1"=="" (
        pause
        goto :menu
    )
    goto :done
)
%PY% model.py --check-overrides "%DECK%"
set "RC=%ERRORLEVEL%"
if not defined RC set "RC=0"
echo.
if "%RC%"=="0" (
    echo  The deck is valid.
) else (
    echo  The deck has a problem - see the message above for the line number.
)
echo.
if "%~1"=="" (
    pause
    goto :menu
)
goto :done

:do_run

rem  Pass the deck explicitly when it exists. model.py would pick up
rem  overrides.txt on its own, but naming it here means the scenario can be
rem  echoed before anything runs - nobody should discover mid-read that their
rem  numbers came from a hand-edited file.
set "DECKARG="
set "SCEN="
if not exist "%DECK%" goto :deck_done
set DECKARG=--overrides "%DECK%"
rem  Pull the scenario name out of the [scenario] block for the banner. Best
rem  effort only - if the deck does not name one, the banner just omits it.
for /f "usebackq tokens=1,* delims==" %%A in (`findstr /b /i /c:"name" "%DECK%"`) do (
    set "SCEN=%%B"
)
:deck_done

echo.
if defined SCEN (
    echo Using manual entries from %DECKNAME%  ^(scenario:%SCEN%^)
    echo.
) else if defined DECKARG (
    echo Using manual entries from %DECKNAME%
    echo.
)
if /i "%CHOICE%"=="ALL" (
    echo Running ALL regions...
    echo.
    %PY% model.py %DECKARG%
) else if /i "%CHOICE%"=="STRICT" (
    echo Running ALL regions, strict ^(no relaxations^)...
    echo.
    %PY% model.py --no-auto-relax %DECKARG%
) else (
    echo Running: %CHOICE%
    echo.
    %PY% model.py --region "%CHOICE%" %DECKARG%
)
set "RC=%ERRORLEVEL%"

echo.
echo ------------------------------------------------------------
if "%RC%"=="0" (
    echo  Done - every area solved cleanly.
    echo  Workbooks are in the Output folder next to this file.
) else if "%RC%"=="4" (
    echo  Done - workbooks were written, BUT some areas only solved
    echo  after allowing unmet demand or surplus volume. That means
    echo  their contracted supply does not cover demand as entered.
    echo.
    echo  Open the workbook, check the Status column on Summary, then
    echo  read the Diagnostics sheet for the exact months/locations.
) else if "%RC%"=="3" (
    echo  At least one area could not be solved at all.
    echo  See the Diagnostics sheet in the workbook for the reason.
) else if "%RC%"=="5" (
    echo  Nothing ran - %DECKNAME% could not be used. Either it would not
    echo  parse, or one of its lines matched nothing in the workbooks.
    echo.
    echo  Choose CHECK from the menu to see the exact line and the nearest
    echo  correct spellings.
) else if "%RC%"=="2" (
    echo  Nothing ran - the region or group name was not recognised.
    echo  Pick from the list above, or use GROUPS to see saved sets.
) else (
    echo  Run failed with exit code %RC%. See the messages above.
)
echo ------------------------------------------------------------
echo.

if "%~1"=="" (
    pause
    goto :menu
)

:done
rem  Belt and braces: never let an empty RC turn into a bare "exit /b", which
rem  would leak whatever errorlevel happened to be lying around.
if not defined RC set "RC=0"
if "%~1"=="" exit /b 0
exit /b %RC%
