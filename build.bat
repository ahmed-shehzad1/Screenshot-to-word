@echo off
setlocal
cd /d "%~dp0"

echo ==========================================
echo   Screenshot to Word - Windows Build
echo ==========================================
echo.

REM Use the installed Python 3.13 interpreter.
where py >nul 2>nul
if errorlevel 1 (
    echo ERROR: Python launcher was not found.
    echo Please make sure the Python launcher is installed.
    goto :error
)

py -3.13 --version
if errorlevel 1 (
    echo ERROR: Python 3.13 was not found.
    goto :error
)

echo.
echo Installing/verifying application and build dependencies...
py -3.13 -m pip install -e . pyinstaller
if errorlevel 1 goto :error

echo.
echo Building Screenshot to Word...
py -3.13 -m PyInstaller --noconfirm --clean --windowed --onedir ^
    --name ScreenshotToWord --paths src ^
    --icon "src\screenshot_to_word\assets\app_icon.ico" ^
    --add-data "src\screenshot_to_word\assets;screenshot_to_word\assets" ^
    --collect-all PySide6 ^
    "src\screenshot_to_word\__main__.py"
if errorlevel 1 goto :error

echo.
echo Looking for Inno Setup compiler...

set "ISCC="

REM Inno Setup 7 - per-user installation (the location found on this PC).
if exist "%LOCALAPPDATA%\Programs\Inno Setup 7\ISCC.exe" (
    set "ISCC=%LOCALAPPDATA%\Programs\Inno Setup 7\ISCC.exe"
)

REM Inno Setup 7 - machine-wide installation locations.
if not defined ISCC if exist "%ProgramFiles%\Inno Setup 7\ISCC.exe" (
    set "ISCC=%ProgramFiles%\Inno Setup 7\ISCC.exe"
)
if not defined ISCC if exist "%ProgramFiles(x86)%\Inno Setup 7\ISCC.exe" (
    set "ISCC=%ProgramFiles(x86)%\Inno Setup 7\ISCC.exe"
)

REM Also support Inno Setup 6 installations.
if not defined ISCC if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" (
    set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
)
if not defined ISCC if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" (
    set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
)

if not defined ISCC (
    echo ERROR: Inno Setup compiler was not found.
    echo The application files were built in dist\ScreenshotToWord.
    echo Install Inno Setup 6 or 7, then run this file again.
    goto :error
)

echo Found Inno Setup compiler:
echo "%ISCC%"
echo.
echo Creating Windows installer...
"%ISCC%" "installer\ScreenshotToWord.iss"
if errorlevel 1 goto :error

echo.
echo ==========================================
echo   BUILD COMPLETED SUCCESSFULLY
echo ==========================================
echo Check the release folder for the installer.
echo.
pause
exit /b 0

:error
echo.
echo ==========================================
echo   BUILD FAILED
echo ==========================================
echo Read the error messages above and share them for troubleshooting.
echo.
pause
exit /b 1
