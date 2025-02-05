@ECHO OFF
SETLOCAL
cls
@TITLE -- AMS_UpdateUsersTool --

:: Check for PowerShell executable
IF NOT EXIST "%SYSTEMROOT%\system32\windowspowershell\v1.0\powershell.exe" (
    COLOR 0C
    ECHO - "powershell.exe" not found!
    ECHO - This script requires PowerShell. Please install PowerShell, then re-run this script.
    COLOR
    pause
    EXIT
)

:: Check for minimum PowerShell version
ECHO - Checking for PowerShell 5.0 (minimum)...
FOR /F "delims=" %%A IN ("%SYSTEMROOT%\system32\windowspowershell\v1.0\powershell.exe" -Command "$PSVersionTable.PSVersion.Major") DO SET PSVersion=%%A
IF %PSVersion% LSS 5 (
    COLOR 0C
    ECHO - This script requires a minimum PowerShell version of 5.0!
    ECHO - Please install PowerShell v5.0 or higher, then re-run this script.
    COLOR
    pause
    EXIT
)
ECHO - PowerShell version %PSVersion% detected. OK.

:: Ensure elevated permissions
NET SESSION >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    ECHO - This script requires administrative privileges. Relaunching with elevated permissions...
    PowerShell -Command "Start-Process -Verb RunAs -FilePath '%~f0'"
    EXIT
)

:: Launch the PowerShell script
ECHO - Starting AMS_UpdateUsersTool...
PowerShell -NoExit -ExecutionPolicy Bypass -File "%~dp0AMS_UpdateUsersTool.ps1"

GOTO END

:END
ENDLOCAL