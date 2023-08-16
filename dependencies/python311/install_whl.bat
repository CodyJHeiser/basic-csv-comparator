@echo off
setlocal enabledelayedexpansion

REM Set the directory containing the .whl files. Remove trailing backslash for safety.
set WHEEL_DIR=%~dp0
set WHEEL_DIR=%WHEEL_DIR:~0,-1%

echo Installing wheel files from: %WHEEL_DIR%

for %%f in (%WHEEL_DIR%\*.whl) do (
    echo Installing: %%f
    echo %PIP_EXECUTABLE% install --no-index --find-links="%WHEEL_DIR%" "%%f"
    pip install --no-index --find-links="%WHEEL_DIR%" "%%f"
)

echo All wheel files installed.
pause
