@echo off
rem Builds the Windows app in app\dist\ and, if Inno Setup is installed,
rem an installer: app\dist\TAM Case Review Setup.exe
rem Needs Python 3.9+ from python.org ("py" launcher). Double-click this file or run it in a terminal.
setlocal
cd /d "%~dp0"
set NAME=TAM Case Review

if not exist .venv\Scripts\python.exe (
  py -3 -m venv .venv || goto :fail
)
.venv\Scripts\python -m pip install -q --upgrade pip || goto :fail
.venv\Scripts\python -m pip install -q -r requirements-build.txt || goto :fail

set PROFILE_ARG=
if exist ..\profile.json (
  set PROFILE_ARG=--add-data "..\profile.json;."
  echo Bundling profile.json
) else (
  echo No profile.json: building with neutral defaults ^(see profiles\example.json^)
)

if exist build rmdir /s /q build
if exist "dist\%NAME%" rmdir /s /q "dist\%NAME%"
.venv\Scripts\python -m PyInstaller --noconfirm --windowed --clean ^
  --name "%NAME%" ^
  --icon assets\icon.ico ^
  --paths .. ^
  --add-data "ui.html;." ^
  %PROFILE_ARG% ^
  tam_app.py || goto :fail

set ISCC=
where iscc >nul 2>nul && set ISCC=iscc
if not defined ISCC if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set ISCC="%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not defined ISCC if exist "%LocalAppData%\Programs\Inno Setup 6\ISCC.exe" set ISCC="%LocalAppData%\Programs\Inno Setup 6\ISCC.exe"
if defined ISCC (
  %ISCC% /Q installer.iss || goto :fail
  echo.
  echo Built: app\dist\%NAME% Setup.exe   ^<- send this one to coworkers
) else (
  powershell -NoProfile -Command "Compress-Archive -Force -Path 'dist\%NAME%' -DestinationPath 'dist\%NAME% (Windows).zip'"
  echo.
  echo Built: app\dist\%NAME%\%NAME%.exe
  echo        app\dist\%NAME% ^(Windows^).zip
  echo Install Inno Setup ^(jrsoftware.org/isinfo.php^) and run this again to get a one-click installer.
)
exit /b 0

:fail
echo.
echo Build failed - see the messages above.
exit /b 1
