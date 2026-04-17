@echo off
echo ========================================
echo  Audio Switcher - Build Script
echo ========================================
echo.

:: Check Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Please install Python 3.10+ from https://python.org
    pause
    exit /b 1
)

echo [1/3] Installing pip and dependencies...

:: Try to use pip — if missing, download get-pip.py and install with --user (no admin needed)
python -m pip --version >nul 2>&1
if errorlevel 1 (
    echo pip not found - downloading it now...
    powershell -NoProfile -ExecutionPolicy Bypass -Command "Invoke-WebRequest -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile 'get-pip.py'"
    if errorlevel 1 (
        echo ERROR: Could not download pip. Please check your internet connection.
        pause
        exit /b 1
    )
    python get-pip.py --user
    del get-pip.py
)

:: Install packages to user folder (no admin rights needed)
python -m pip install pycaw pystray Pillow comtypes pyinstaller --user --quiet
if errorlevel 1 (
    echo ERROR: Failed to install dependencies.
    pause
    exit /b 1
)

:: Clean up old build artifacts so we don't hit permission errors
if exist dist\AudioSwitcher.exe (
    taskkill /f /im AudioSwitcher.exe >nul 2>&1
    timeout /t 1 /nobreak >nul
    del /f dist\AudioSwitcher.exe >nul 2>&1
)
if exist AudioSwitcher.spec del /f AudioSwitcher.spec >nul 2>&1

echo [2/3] Building AudioSwitcher.exe ...
python -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name "AudioSwitcher" ^
    --hidden-import=comtypes ^
    --hidden-import=comtypes.client ^
    --hidden-import=pystray._win32 ^
    audio_switcher.py

if errorlevel 1 (
    echo ERROR: Build failed. See output above for details.
    pause
    exit /b 1
)

echo [3/3] Done!
echo.
echo Your app is at: dist\AudioSwitcher.exe
echo.
echo TIP: To run on startup, right-click AudioSwitcher.exe ^> Create shortcut,
echo      then move the shortcut to:
echo      %%APPDATA%%\Microsoft\Windows\Start Menu\Programs\Startup
echo.
pause
