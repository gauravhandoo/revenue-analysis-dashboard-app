@echo off
setlocal
cd /d "%~dp0"

set "CACHE_DIR=%LOCALAPPDATA%\RevenueAnalysisDashboard"
if not exist "%CACHE_DIR%" mkdir "%CACHE_DIR%"

set "VENV_DIR=%CACHE_DIR%\venv"
set "STAMP_FILE=%CACHE_DIR%\requirements.sha256"
set "REQ_FILE=%~dp0requirements.txt"

if not exist "%VENV_DIR%\Scripts\python.exe" (
  if exist "%~dp0.venv\Scripts\python.exe" (
    set "VENV_DIR=%~dp0.venv"
  ) else (
    where py >nul 2>nul
    if not errorlevel 1 (
      py -3 -m venv "%VENV_DIR%"
    ) else (
      where python >nul 2>nul
      if not errorlevel 1 (
        python -m venv "%VENV_DIR%"
      ) else (
        echo Python runtime not found. Install Python or share a prebuilt .venv in the app folder.
        exit /b 1
      )
    )
  )
)

if not exist "%VENV_DIR%\Scripts\activate.bat" (
  echo Python environment is missing activation script at "%VENV_DIR%\Scripts\activate.bat"
  exit /b 1
)

call "%VENV_DIR%\Scripts\activate.bat"

for /f %%H in ('powershell -NoProfile -Command "(Get-FileHash -Algorithm SHA256 '%REQ_FILE%').Hash"') do set "CURRENT_HASH=%%H"
set "CACHED_HASH="
if exist "%STAMP_FILE%" set /p CACHED_HASH=<"%STAMP_FILE%"

if not "%CURRENT_HASH%"=="%CACHED_HASH%" (
  python -m pip install --upgrade pip
  python -m pip install -r "%REQ_FILE%"
  >"%STAMP_FILE%" echo %CURRENT_HASH%
)

python -m streamlit run app.py
