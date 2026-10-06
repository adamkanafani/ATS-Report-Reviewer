@echo off
title ATS Report Reviewer - Setup
echo ============================================
echo   ATS Report Reviewer - First-Time Setup
echo ============================================
echo.
echo Checking for Node.js...
where node >nul 2>nul
if errorlevel 1 (
  echo.
  echo Node.js is not installed on this computer.
  echo Please install it first from https://nodejs.org ^(choose the "LTS" version,
  echo click Next through the installer with the default options^), then run this
  echo file again.
  echo.
  pause
  exit /b 1
)
echo Node.js found.
echo.
if not exist ".env.local" (
  echo The spelling, grammar, photo-check and Overall Assessment reviews use Claude AI
  echo and need an Anthropic API key. Paste it below and press Enter, or just press
  echo Enter to skip ^(you can add it later in a file named .env.local^).
  echo.
  set /p APIKEY="Anthropic API key: "
)
if not exist ".env.local" (
  if defined APIKEY (
    > ".env.local" echo ANTHROPIC_API_KEY=%APIKEY%
    echo API key saved.
  ) else (
    echo Skipped -- the AI checks will be turned off until a key is added.
  )
  echo.
)
echo Installing app dependencies - this can take a few minutes the first time...
call npm install
if errorlevel 1 (
  echo.
  echo Something went wrong installing dependencies. Scroll up to see the error,
  echo or send a screenshot of this window to whoever set this up.
  pause
  exit /b 1
)
echo.
echo Building the app for first use...
call npm run build
if errorlevel 1 (
  echo.
  echo Something went wrong building the app. Scroll up to see the error,
  echo or send a screenshot of this window to whoever set this up.
  pause
  exit /b 1
)
echo.
echo ============================================
echo   Setup complete!
echo   From now on, just double-click "Start Report Reviewer.bat" to run the app.
echo ============================================
pause
