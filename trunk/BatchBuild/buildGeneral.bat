@echo off

rem %1 - Delphi version - 4 or 6
rem %2 - Project name
rem %3 - Map file - optional (Y or N)
rem %4 - Folder - optional if different to Project name
rem %4 - Header File name - optional if different to build.h
rem %5 - relative folder to get to root (required if more than one folder below) e.g. ..\

set DELPHI=%1
set PROJECT=%2
set MAP=%3
set FOLDER=%4
set HEADERFILE=%5
set RELATIVEFOLDER=%6

if %DELPHI%%FOLDER%==%DELPHI% (
set FOLDER=%2
)

if %DELPHI4%1==1 (
set DELPHI4=c:\dev\Delphi4
)

if %DELPHI5%1==1 (
set DELPHI5=C:\Program Files\Borland\Delphi5
)

if %DELPHI6%1==1 (
set DELPHI6=G:\Programs\Borland\Delphi6
)

pushd ..
cd %FOLDER%
if exist %PROJECT%.cfg del %PROJECT%.cfg
if exist %PROJECT%-Build.cfg copy %PROJECT%-Build.cfg %PROJECT%.cfg

if %DELPHI%==6 (

%DELPHI6%\bin\dcc32 %PROJECT% -E%RELATIVEFOLDER%..\exe

)

if %DELPHI%==5 ( 

"%DELPHI5%\bin\dcc32" %PROJECT% -E%RELATIVEFOLDER%..\exe

)

if %DELPHI%==4 (

%DELPHI4%\bin\dcc32 %PROJECT% -E%RELATIVEFOLDER%..\exe

)

IF errorlevel=1 GOTO fail

GOTO end
:fail
echo Build failed
pause
:end