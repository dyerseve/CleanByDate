:: //***************************************************************************
:: //
:: // File:      CleanByDate.cmd
:: //
:: // Additional files required:  None.  Script creates required elevate.cmd and 
:: //                             elevate.vbs in %Temp% when run.
:: //
:: // Purpose:   clears folder of files older than X, only works on files in the folder and only works on local paths
:: //            
:: //
:: // Usage:     CleanByDate.cmd
:: //
:: // Version:   0.3
:: //
:: // History:
:: // 0.3   18.05.02 Removed password prompt, not needed, runs as SYSTEM
:: // 0.2   17.03.03 Fixed: Folder spaces in script path don't work, commented debug lines, added script path variable display
:: // 0.1   17.02.23 First Revision
:: //
:: // ***** End Header *****
:: //***************************************************************************

@echo off
setlocal enabledelayedexpansion

set CmdDir=%~dp0
set CmdDir=%CmdDir:~0,-1%


:: ////////////////////////////////////////////////////////////////////////////
:: Check whether running elevated
:: ////////////////////////////////////////////////////////////////////////////
call :CREATE_ELEVATE_SCRIPTS

:: Check for Mandatory Label\High Mandatory Level
whoami /groups | find "S-1-16-12288" > nul
if "%errorlevel%"=="0" (
    echo Running as elevated user.  Continuing script.
) else (
    echo Not running as elevated user.
    echo Relaunching Elevated: "%~dpnx0" %*

    if exist "%Temp%\elevate.cmd" (
        set ELEVATE_COMMAND="%Temp%\elevate.cmd"
    ) else (
        set ELEVATE_COMMAND=elevate.cmd
    )

    set CARET=^^
    !ELEVATE_COMMAND! cmd /k cd /d "%~dp0" !CARET!^& call "%~dpnx0" %*
    goto :EOF
)

if exist %ELEVATE_CMD% del %ELEVATE_CMD%
if exist %ELEVATE_VBS% del %ELEVATE_VBS%


:: ////////////////////////////////////////////////////////////////////////////
:: Main script code starts here
:: ////////////////////////////////////////////////////////////////////////////
@echo off
SETLOCAL ENABLEEXTENSIONS

ECHO Wscript.Echo Msgbox("CleanByDate Script v0.2 (17.03.03) Not updated for XP, W7 or later.")>%TEMP%\~input.vbs
cscript //nologo %TEMP%\~input.vbs
DEL %TEMP%\~input.vbs

::Input Destination
ECHO Wscript.Echo Inputbox("Enter Destination for the created script without trailing slash, no spaces please(Default: C:\Windows):")>%TEMP%\~input.vbs
FOR /f "delims=/" %%G IN ('cscript //nologo %TEMP%\~input.vbs') DO set dest=%%G
DEL %TEMP%\~input.vbs

IF "%dest%" == "" set dest=C:\Windows
::Remove Trailing Slash
IF %dest:~-1%==\ SET dest=%dest:~0,-1%

::ECHO ON
if not exist "%dest%" mkdir "%dest%"
::echo DEBUG: mkdir
::pause
::ECHO OFF

set /a count=0

:menuLOOP
cls

::This is here to avoid password botching
SETLOCAL ENABLEDELAYEDEXPANSION
echo Script Path %dest%
echo Last Folder %folder%
echo Last Mask !mask!
echo Last Days to keep %days%
echo Folder Count is %count%

::debug pause
::pause

for /f "tokens=1,2,* delims=_ " %%A in ('"findstr /b /c:":menu_" "%~f0""') do echo.  %%B  %%C
set choice=
echo.&set /p choice=Selection(Q to quit): ||GOTO:EOF
echo.&call:menu_%choice%
GOTO:menuLOOP

:menu_A   Add Folder
@ECHO OFF
set folder=
set mask=
set days=

::Input Destination
ECHO Wscript.Echo Inputbox("Enter Folder to clean by date")>%TEMP%\~input.vbs
FOR /f "delims=/" %%G IN ('cscript //nologo %TEMP%\~input.vbs') DO set folder=%%G
DEL %TEMP%\~input.vbs

::Input Mask
ECHO Wscript.Echo Inputbox("Enter File Mask (Default *.*)")>%TEMP%\~input.vbs
FOR /f "delims=/" %%G IN ('cscript //nologo %TEMP%\~input.vbs') DO set mask=%%G
DEL %TEMP%\~input.vbs

::Input Days
ECHO Wscript.Echo Inputbox("Enter Number of Days to keep (Default 14)")>%TEMP%\~input.vbs
FOR /f "delims=/" %%G IN ('cscript //nologo %TEMP%\~input.vbs') DO set days=%%G
DEL %TEMP%\~input.vbs

IF "%folder%" == "" goto ERROR
IF "%mask%" == "" set mask=*.*
IF "%days%" == "" set days=14
::Remove Trailing Slash
IF %folder:~-1%==\ SET dest=%folder:~0,-1%
::pause
IF %count% NEQ 0 goto :next  
echo forfiles /p "%folder%" /m !mask! /d -%days% /c "cmd /c echo @PATH" > %dest%\echotest.cmd
echo forfiles /p "%folder%" /m !mask! /d -%days% /c "cmd /c del @PATH" > %dest%\cleanfiles.cmd
goto addcomplete
:next
echo forfiles /p "%folder%" /m !mask! /d -%days% /c "cmd /c echo @PATH" >> %dest%\echotest.cmd
echo forfiles /p "%folder%" /m !mask! /d -%days% /c "cmd /c del @PATH" >> %dest%\cleanfiles.cmd
:addcomplete

set /A Count+=1
@ECHO OFF
::pause
GOTO:menuLOOP

:menu_T   Do a WhatIf Run
ECHO ON
call "%dest%\echotest.cmd"
@ECHO OFF
pause
GOTO:menuLOOP

:menu_S   Schedule and Run Task
ECHO ON
schtasks /delete /tn cleanfiles /F
schtasks /create /sc daily /tn cleanfiles /tr "%dest%\cleanfiles.cmd" /RU "System"
schtasks /run /tn cleanfiles

pause
@ECHO OFF
GOTO:menuLOOP

:menu_Q   Quit
EXIT
GOTO:EOF

:ERROR
ECHO Error field was blank, returning to menu
PING 1.1.1.1 -n 1 -w 5000 >NUL
@ECHO OFF
GOTO:menuLOOP

:: ////////////////////////////////////////////////////////////////////////////
:: End of main script code here
:: ////////////////////////////////////////////////////////////////////////////
goto :EOF


:: ////////////////////////////////////////////////////////////////////////////
:: Subroutines
:: ////////////////////////////////////////////////////////////////////////////

:CREATE_ELEVATE_SCRIPTS

    set ELEVATE_CMD="%Temp%\elevate.cmd"

    echo @setlocal>%ELEVATE_CMD%
    echo @echo off>>%ELEVATE_CMD%
    echo. >>%ELEVATE_CMD%
    echo :: Pass raw command line agruments and first argument to Elevate.vbs>>%ELEVATE_CMD%
    echo :: through environment variables.>>%ELEVATE_CMD%
    echo set ELEVATE_CMDLINE=%%*>>%ELEVATE_CMD%
    echo set ELEVATE_APP=%%1>>%ELEVATE_CMD%
    echo. >>%ELEVATE_CMD%
    echo start wscript //nologo "%%~dpn0.vbs" %%*>>%ELEVATE_CMD%


    set ELEVATE_VBS="%Temp%\elevate.vbs"

    echo Set objShell ^= CreateObject^("Shell.Application"^)>%ELEVATE_VBS% 
    echo Set objWshShell ^= WScript.CreateObject^("WScript.Shell"^)>>%ELEVATE_VBS%
    echo Set objWshProcessEnv ^= objWshShell.Environment^("PROCESS"^)>>%ELEVATE_VBS%
    echo. >>%ELEVATE_VBS%
    echo ' Get raw command line agruments and first argument from Elevate.cmd passed>>%ELEVATE_VBS%
    echo ' in through environment variables.>>%ELEVATE_VBS%
    echo strCommandLine ^= objWshProcessEnv^("ELEVATE_CMDLINE"^)>>%ELEVATE_VBS%
    echo strApplication ^= objWshProcessEnv^("ELEVATE_APP"^)>>%ELEVATE_VBS%
    echo strArguments ^= Right^(strCommandLine, ^(Len^(strCommandLine^) - Len^(strApplication^)^)^)>>%ELEVATE_VBS%
    echo. >>%ELEVATE_VBS%
    echo If ^(WScript.Arguments.Count ^>^= 1^) Then>>%ELEVATE_VBS%
    echo     strFlag ^= WScript.Arguments^(0^)>>%ELEVATE_VBS%
    echo     If ^(strFlag ^= "") OR (strFlag="help") OR (strFlag="/h") OR (strFlag="\h") OR (strFlag="-h"^) _>>%ELEVATE_VBS%
    echo         OR ^(strFlag ^= "\?") OR (strFlag = "/?") OR (strFlag = "-?") OR (strFlag="h"^) _>>%ELEVATE_VBS%
    echo         OR ^(strFlag ^= "?"^) Then>>%ELEVATE_VBS%
    echo         DisplayUsage>>%ELEVATE_VBS%
    echo         WScript.Quit>>%ELEVATE_VBS%
    echo     Else>>%ELEVATE_VBS%
    echo         objShell.ShellExecute strApplication, strArguments, "", "runas">>%ELEVATE_VBS%
    echo     End If>>%ELEVATE_VBS%
    echo Else>>%ELEVATE_VBS%
    echo     DisplayUsage>>%ELEVATE_VBS%
    echo     WScript.Quit>>%ELEVATE_VBS%
    echo End If>>%ELEVATE_VBS%
    echo. >>%ELEVATE_VBS%
    echo. >>%ELEVATE_VBS%
    echo Sub DisplayUsage>>%ELEVATE_VBS%
    echo. >>%ELEVATE_VBS%
    echo     WScript.Echo "Elevate - Elevation Command Line Tool for Windows Vista" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "Purpose:" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "--------" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "To launch applications that prompt for elevation (i.e. Run as Administrator)" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "from the command line, a script, or the Run box." ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "Usage:   " ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "    elevate application <arguments>" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "Sample usage:" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "    elevate notepad ""C:\Windows\win.ini""" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "    elevate cmd /k cd ""C:\Program Files""" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "    elevate powershell -NoExit -Command Set-Location 'C:\Windows'" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "Usage with scripts: When using the elevate command with scripts such as" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "Windows Script Host or Windows PowerShell scripts, you should specify" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "the script host executable (i.e., wscript, cscript, powershell) as the " ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "application." ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "Sample usage with scripts:" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "    elevate wscript ""C:\windows\system32\slmgr.vbs"" –dli" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "    elevate powershell -NoExit -Command & 'C:\Temp\Test.ps1'" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "The elevate command consists of the following files:" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "    elevate.cmd" ^& vbCrLf ^& _>>%ELEVATE_VBS%
    echo                  "    elevate.vbs" ^& vbCrLf>>%ELEVATE_VBS%
    echo. >>%ELEVATE_VBS%
    echo End Sub>>%ELEVATE_VBS%

goto :EOF




