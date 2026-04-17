@echo off
echo ========================================
echo  Audio Switcher - Install Startup Entry
echo ========================================
echo.

set "EXE=%~dp0dist\AudioSwitcher.exe"
set "STARTUP=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "SHORTCUT=%STARTUP%\AudioSwitcher.lnk"

if not exist "%EXE%" (
    echo ERROR: dist\AudioSwitcher.exe not found.
    echo        Run build.bat first to compile the application.
    pause
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -Command "$ws=New-Object -ComObject WScript.Shell; $s=$ws.CreateShortcut('%SHORTCUT%'); $s.TargetPath='%EXE%'; $s.WorkingDirectory='%~dp0dist'; $s.Description='AudioSwitcher tray app'; $s.Save()"

if errorlevel 1 (
    echo ERROR: Failed to create startup shortcut.
    pause
    exit /b 1
)

echo Done! AudioSwitcher will now run on startup.
echo Shortcut created at:
echo   %SHORTCUT%
echo.
pause
