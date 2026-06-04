@echo off
setlocal
 
@REM file to run
@REM set UI=%1
set UI="C:\Projects\MORPC\DTA\Resources\ui"
 
@REM Outer iteration 
@REM set OUTER=%2
set OUTER=1
 
@REM Below 
@REM Scenario Name
@REM set SCEN_NAME=%3
set SCEN_NAME=Base
 
@REM had to change run_tsm 
@REM set REGION=%4
set REGION=MOR

@REM This isn't in the WSP batch file but is needed
set INNER=1
 
 
set TSM60_FOLDER=C:\Program Files\TransModeler 8.0
 
echo "%TSM60_FOLDER%/tsm.exe" -a %UI% -ai run_dta -q
"%TSM60_FOLDER%/tsm.exe" -a %UI% -ai run_dta -q
pause