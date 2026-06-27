@echo off
setlocal enabledelayedexpansion

:: ============================================================
:: Word <-> Markdown Converter (Drag & Drop)
:: Drop .docx or .md files onto this bat icon to convert
:: ============================================================

cd /d "%~dp0"

:: ---- Find Python ----
set PYTHON=
for %%p in (python py) do (
    %%p --version >nul 2>&1
    if not errorlevel 1 (
        set "PYTHON=%%p"
        goto :found
    )
)
for /d %%d in (
    "%LOCALAPPDATA%\Programs\Python\Python3*"
    "%PROGRAMFILES%\Python3*"
    "C:\Python3*"
) do (
    if exist "%%d\python.exe" (
        set "PYTHON=%%d\python.exe"
        goto :found
    )
)
echo [ERROR] Python not found!
echo Please install Python 3.8+ from https://www.python.org/downloads/
echo Make sure to check "Add Python to PATH" during installation.
pause
exit /b 1

:found

:: ---- Check dependencies ----
%PYTHON% -c "import docx" >nul 2>&1
if errorlevel 1 (
    echo Installing dependencies...
    %PYTHON% -m pip install python-docx mammoth -q
    if errorlevel 1 (
        echo [ERROR] Install failed. Please run manually:
        echo    %PYTHON% -m pip install python-docx mammoth
        pause
        exit /b 1
    )
)

:: ---- No file dropped: show help ----
if "%~1"=="" (
    echo.
    echo ==============================================
    echo   Word ^<--^> Markdown Converter
    echo ==============================================
    echo.
    echo Usage:
    echo   1. Drop .docx file -^> generates .md
    echo   2. Drop .md file   -^> generates .docx
    echo   3. Drop multiple files at once
    echo.
    echo Advanced:
    echo   python convert_docx.py --watch FOLDER_PATH
    echo ==============================================
    pause
    exit /b 0
)

:: ---- Process dropped files ----
echo.
echo ==============================================
echo   Converting...
echo ==============================================
echo.

set count=0
set fail=0

:loop
if "%~1"=="" goto done
    set "file=%~1"
    echo [%count%] %~nx1
    %PYTHON% "%~dp0convert_docx.py" "!file!"
    if errorlevel 1 (
        echo   [FAILED]
        set /a fail+=1
    )
    set /a count+=1
    shift
    goto loop

:done
echo.
set /a ok=%count% - %fail%
echo ==============================================
echo   Done: %ok% OK / %count% total
echo ==============================================
echo.
echo Press any key to close...
pause >nul
