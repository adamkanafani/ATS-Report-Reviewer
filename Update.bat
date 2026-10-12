@echo off
title ATS Report Reviewer - Update
echo ============================================
echo   ATS Report Reviewer - Update
echo ============================================
echo.
echo IMPORTANT: close the "ATS Report Reviewer - Server" window first if it is open.
echo The update can't rebuild the app while it is running.
echo.
pause
echo.
echo Downloading the latest version...
call git pull
if errorlevel 1 (
  echo.
  echo Could not download the update. If this folder wasn't set up with git
  echo ^(e.g. it was copied as a ZIP^), download a fresh copy instead and ask
  echo whoever set this up for help.
  echo.
  pause
  exit /b 1
)
echo.
echo Installing any new dependencies...
call npm install
if errorlevel 1 (
  echo.
  echo Installing dependencies failed. Scroll up to see the error, or send a
  echo screenshot of this window to whoever set this up.
  pause
  exit /b 1
)
echo.
echo Rebuilding the app...
call npm run build
if errorlevel 1 (
  echo.
  echo Rebuilding the app failed - the OLD version is still what will run.
  echo Make sure the server window is closed, then run Update.bat again.
  echo If it still fails, send a screenshot of this window to whoever set this up.
  pause
  exit /b 1
)
echo.
echo ============================================
echo   Update complete! Now running version:
git log -1 --format="  %%h  (%%cd)" --date=format:"%%b %%d, %%Y"
echo   Start the app again with "Start Report Reviewer.bat".
echo   Your API key setting was not affected.
echo ============================================
pause
