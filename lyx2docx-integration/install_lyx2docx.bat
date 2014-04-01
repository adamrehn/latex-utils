@echo off

@rem Attempt to determine the location of the LyX executable
@for /f "tokens=*" %%i in ('where lyx') do @set LYXLOC=%%i
if "%LYXLOC%" == "" (
	echo Error: LyX not found in the system PATH!
	pause
	exit
)

@rem Determine the location of Python bundled with LyX
setlocal enabledelayedexpansion
set "string=%LYXLOC%"
set find=bin\LyX.exe
set replace=Python\python.exe
call set PYTHONLOC=%%string:!find!=!replace!%%

@rem Execute the install script using the Python bundled with LyX
"%PYTHONLOC%" install_lyx2docx.py
pause