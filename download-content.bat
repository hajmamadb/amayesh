@echo off
rem ============================================================
rem  Download all content from amar.org.ir + iranstatis.com
rem  Output folder:  D:\amayesh\<target folder shown in console>
rem  Double-click to run. Extra args are passed to the script,
rem  e.g.:  download-content.bat -Site amar
rem ============================================================
title Download Statistics Sites Content
cd /d "%~dp0"

where pwsh >nul 2>nul
if %errorlevel%==0 (
    pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0download-content.ps1" %*
) else (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0download-content.ps1" %*
)

echo.
echo ============================================================
echo  Finished.
echo  Output folder: D:\amayesh + subfolder (see console log)
echo  Summary file :  Khalase-Koli.txt inside the output folder
echo  Full log     :  gozarisht\log.txt
echo ============================================================
pause
