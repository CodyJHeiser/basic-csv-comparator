@echo off
setlocal enabledelayedexpansion

REM Set the directory containing the .whl files. You can change this to any path you want.
set WHEEL_DIR=%~dp0

echo Installing wheel files from: %WHEEL_DIR%

for %%f in (%WHEEL_DIR%*.whl) do (
    echo Installing: %%f
    pip install --no-index --find-links="%WHEEL_DIR%" "%%f"
)

echo All wheel files installed.
pause
