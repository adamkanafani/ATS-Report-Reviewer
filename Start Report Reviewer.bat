@echo off
title ATS Report Reviewer - Launcher
echo Starting ATS Report Reviewer...
echo.
echo A new window will open running the app - leave THAT window open while you
echo work (you can minimize it). Your browser will open automatically in a
echo few seconds. This window will close on its own.
echo.
if not exist ".next" (
  echo It looks like setup hasn't been run yet on this computer.
  echo Please double-click "Setup - Run This First.bat" first.
  echo.
  pause
  exit /b 1
)
start "ATS Report Reviewer - Server (leave this open)" cmd /k "npm run start"
timeout /t 5 /nobreak >nul
start "" http://localhost:3200
exit
