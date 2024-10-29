@echo off
setlocal enabledelayedexpansion

REM Get the current user's AppData\Roaming folder path
set "appDataPath=%APPDATA%"

REM Specify the relative path to the source files based on the current script location
set "sourcePath=%~dp0"

REM Specify the destination base folder in the recipient's AppData\Roaming folder
set "destinationBaseFolder=%appDataPath%\Autodesk\Revit\Addins"

REM Get a list of installed Revit versions
for /D %%V in ("%destinationBaseFolder%\*") do (
    set "destinationFolder=%%V"

    REM Create the destination folder if it doesn't exist
    if not exist "!destinationFolder!" mkdir "!destinationFolder!"

    REM Step 1: Copy all files and folders except .addin files to the destination folder
    echo Copying files from %sourcePath% to "!destinationFolder!" excluding .addin files...

    robocopy "%sourcePath%\" "!destinationFolder!" /E /XF *.addin

    REM Step 2: Unblock all copied files
    echo Unblocking all copied files in "!destinationFolder!"...
    pushd "!destinationFolder!"
    powershell -Command "Get-ChildItem -Recurse | Unblock-File" >nul 2>&1
    popd

    REM Step 3: Copy all .addin files from the source to the destination folder
    echo Copying .addin files from %sourcePath% to "!destinationFolder!"...
    robocopy "%sourcePath%\" "!destinationFolder!" *.addin /E

    REM Step 4: Unblock all .addin files in the destination folder
    echo Unblocking all .addin files in "!destinationFolder!"...
    pushd "!destinationFolder!"
    powershell -Command "Get-ChildItem -Recurse -Filter *.addin | Unblock-File" >nul 2>&1
    popd

    REM Display a success message for each version
    echo Add-in files have been successfully processed for %%~nxV.
)

echo Process completed.
pause
