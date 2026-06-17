@echo off
setlocal

set "PYTHON=D:\miniconda\conda_envs\camelot_clean2\python.exe"
set "SCRIPT=%~dp0camelotToExcel_FitToDaicho2.0.py"

if not exist "%PYTHON%" (
    echo Python was not found:
    echo %PYTHON%
    pause
    exit /b 1
)

if not exist "%SCRIPT%" (
    echo Application code was not found:
    echo %SCRIPT%
    pause
    exit /b 1
)

pushd "%~dp0"
"%PYTHON%" "%SCRIPT%"
set "EXIT_CODE=%ERRORLEVEL%"
popd

if not "%EXIT_CODE%"=="0" (
    echo.
    echo Application exited with code %EXIT_CODE%.
    pause
)

exit /b %EXIT_CODE%
