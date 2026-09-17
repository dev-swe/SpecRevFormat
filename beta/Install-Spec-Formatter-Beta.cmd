@echo off
REM One-click Spec Formatter (Beta) installer.
REM Registers THIS folder as a per-user Office trusted add-in catalog, so the beta
REM add-in shows up under Insert > Add-ins > Shared Folder. Pure cmd + reg.exe:
REM no admin rights and no PowerShell needed. Run it from the shared/network folder
REM that holds manifest.xml (so every tester's Word can reach the same location).
setlocal EnableDelayedExpansion

set "CAT=%~dp0"
if "!CAT:~-1!"=="\" set "CAT=!CAT:~0,-1!"

if not exist "%CAT%\manifest.xml" (
  echo ERROR: manifest.xml was not found next to this script.
  echo Make sure you kept manifest.xml and this .cmd together.
  echo.
  pause
  exit /b 1
)

set "GUID={9C1B5E2A-3D47-4F8B-A6C1-7E2D9F0A4B58}"
set "KEY=HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%GUID%"

reg add "%KEY%" /v Id    /t REG_SZ    /d "%GUID%" /f >nul
reg add "%KEY%" /v Url   /t REG_SZ    /d "%CAT%"  /f >nul
reg add "%KEY%" /v Flags /t REG_DWORD /d 1        /f >nul

echo.
echo   Spec Formatter (Beta) catalog registered:
echo     %CAT%
echo.
echo   Finish in Word:
echo     1. Fully close Word (every window).
echo     2. Reopen Word and open any document.
echo     3. Insert tab  ^>  Add-ins  ^>  My Add-ins  ^>  SHARED FOLDER tab
echo          ^>  Spec Formatter (Beta)  ^>  Add.
echo     4. The "Spec Formatter" button appears on the Home ribbon.
echo.
echo   (If the SHARED FOLDER tab is empty: File ^> Options ^> Trust Center ^>
echo    Trust Center Settings ^> Trusted Add-in Catalogs, tick "Show in Menu"
echo    for the entry above, click OK, and restart Word.)
echo.
pause
