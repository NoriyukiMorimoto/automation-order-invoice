@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "WORKBOOK=%SCRIPT_DIR%..\é©ìÆâª_íçï∂èë_êøãÅèë.xlsm"
set "SOURCE=%SCRIPT_DIR%..\â^ópíÜVBA"

"%WINDIR%\SysWOW64\cscript.exe" //nologo "%SCRIPT_DIR%sync_open_workbook_vba.vbs" "%WORKBOOK%" "%SOURCE%"
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if not "%EXIT_CODE%"=="0" (
    echo VBA update failed. Exit code: %EXIT_CODE%
) else (
    echo VBA update completed. The workbook remains open.
)

pause
exit /b %EXIT_CODE%
