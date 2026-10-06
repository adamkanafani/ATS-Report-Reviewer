@echo off
title ATS Report Reviewer - Update
echo ============================================
echo   ATS Report Reviewer - Update
echo ============================================
echo.
echo Downloading the latest version...
call git pull
if errorlevel 1 (
  echo.
  echo Could not update automatically. If this folder wasn't set up with git
  echo (e.g. it was copied as a ZIP), download a fresh copy instead and ask
  echo whoever set this up for help.
  echo.
  pause
  exit /b 1
)
echo.
echo Installing any new dependencies...
call npm install
echo.
echo Rebuilding the app...
call npm run build
echo.
echo ============================================
echo   Update complete! Your API key setting was not affected.
echo ============================================
pause
