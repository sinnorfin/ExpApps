@echo off
setlocal enabledelayedexpansion
REM Get the current user's AppData\Roaming folder path
set "appDataPath=%APPDATA%"

REM Specify the path to the add-in files and resource folder
set "sourcePath=X:\Manuals\Revit Addins\ExpAppsInstall"
set "resourceFolderName=ExpApps.bundle"

REM Specify the destination base folder in the recipient's AppData\Roaming folder
set "destinationBaseFolder=%appDataPath%\Autodesk\Revit\Addins"

REM Get a list of installed Revit versions
for /D %%V in ("%destinationBaseFolder%\*") do (
    set "destinationFolder=%%V"

    REM Create the destination folder if it doesn't exist
    if not exist "%%V" mkdir "%%V"

    REM Step 1: Copy the resource folder and its contents to the destination folder
   
    robocopy "X:\Manuals\Revit Addins\ExpAppsInstall\ExpApps.bundle" "%%V\%resourceFolderName%" /E /NP

    REM Step 2:
    set "dllCount=0"
    pushd "%%V\%resourceFolderName%"
    for /R %%F in (*.dll) do (
        REM Unblock the .dll file if it is blocked
        echo "Unblocking file: %%F"
        powershell -Command "Unblock-File -Path '%%~fF'" >nul 2>&1
        set /a "dllCount+=1"
    )
    popd
    echo Number of DLL files found and unblocked: !dllCount!

    REM Step 3: Copy the .addin file to the destination folder
    for %%A in ("%sourcePath%\*.addin") do (
    	  set "addinFilePath=%%~fA"
   	  set "destinationAddinPath=%%V\%%~nxA"
    	  robocopy "%sourcePath%" "%%V" "%%~nxA" /NP >nul 2>&1
    )

    REM Step 4: Unblock the .addin file if needed
    powershell -Command "Unblock-File -Path '%%V\*.addin'" >nul 2>&1

    REM Display a success message for each version
    echo Add-in files have been successfully copied to %%~nxV.
)