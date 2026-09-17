@echo off
REM Removes the Spec Formatter (Beta) trusted-catalog registration for this user.
set "GUID={9C1B5E2A-3D47-4F8B-A6C1-7E2D9F0A4B58}"
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\%GUID%" /f >nul 2>&1
echo.
echo   Removed the Spec Formatter (Beta) catalog registration.
echo   Close Word to finish. To also remove the pane entry, open Word:
echo     Insert ^> Add-ins ^> My Add-ins ^> (right-click the add-in) Remove.
echo.
pause
