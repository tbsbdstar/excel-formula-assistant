@echo off
setlocal
cd /d "%~dp0"
if exist "%~dp0python\python.exe" goto run

rem ---- python\ is missing next to this file: tell the user why (messages are base64 UTF-8 Chinese) ----
if exist "%~dp0build_toolkit.py" goto source
call :say "5om+5LiN5Yiw5bel5YW36Ieq5bim55qEIFB5dGhvbu+8iHB5dGhvbiDmlofku7blpLnvvInjgIIK5aSa5Y2K5piv5Zyo5Y6L57yp5YyF6YeM55u05o6l5Y+M5Ye75LqGIHJ1bi5iYXTvvIzmiJbogIXlj6rmi7fotJ3kuoYgcnVuLmJhdCDov5nkuIDkuKrmlofku7bjgIIKCuivt+i/meagt+WBmu+8muWFs+mXreacrOeql+WPoyDihpIg5Y+z6ZSu5Y6L57yp5YyFIOKGkuOAjOWFqOmDqOino+WOi+e8qeOAjeKGkiDmiZPlvIDop6Pljovlh7rmnaXnmoTjgIxFeGNlbOW3peS9nOa1geW3peWFt+OAjeaWh+S7tuWkuSDihpIg5Y+M5Ye76YeM6Z2i55qEIHJ1bi5iYXTjgII="
goto fail

:source
rem Source folder of the project: start the packaged toolkit in dist\ if it has been built.
for /d %%D in ("%~dp0dist\Excel*") do if exist "%%D\python\python.exe" if exist "%%D\run.bat" (
  call :say "5o+Q56S677ya6L+Z6YeM5piv5rqQ56CB55uu5b2V77yM5rKh5pyJ6Ieq5bimIFB5dGhvbu+8jOW3suS4uuS9oOWQr+WKqOaJk+WMheWlveeahOW3peWFt++8mgrvvIjphY3nva7mlofku7bmlL7lroPnmoQgY29uZmlncyDmlofku7blpLnvvIzopoHlpITnkIbnmoTooajmoLzmlL7lroPnmoQgaW5wdXQg5paH5Lu25aS577yJ"
  echo   %%D
  echo.
  call "%%D\run.bat" %*
  exit /b
)
call :say "6L+Z6YeM5piv5rqQ56CB55uu5b2V77yM5rKh5pyJ6Ieq5bimIFB5dGhvbu+8jOS5n+i/mOayoeacieaJk+WMheWlveeahOW3peWFt+OAggror7flnKjov5nkuKrnm67lvZXov5DooYzvvJpweXRob24gYnVpbGRfdG9vbGtpdC5weQrnhLblkI7kvb/nlKggZGlzdCDmlofku7blpLnph4znlJ/miJDnmoTjgIxFeGNlbOW3peS9nOa1geW3peWFt+OAje+8iOaIluWQjOebruW9leeahCB6aXAg5Y6L57yp5YyF77yJ44CC"
goto fail

:fail
echo.
pause
exit /b 1

:say
powershell -NoProfile -ExecutionPolicy Bypass -Command "[Console]::OutputEncoding=[Text.Encoding]::UTF8; Write-Host ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('%~1'))) -ForegroundColor Yellow" 2>nul
if errorlevel 1 echo [ERROR] python\python.exe not found. Unzip the whole package first, then run run.bat in the unzipped folder.
exit /b 0

:run
chcp 65001 >nul
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
set "PYTHONDONTWRITEBYTECODE=1"
set "PYTHONNOUSERSITE=1"
set "PATH=%~dp0python;%~dp0libs;%PATH%"
if not exist "%~dp0input" mkdir "%~dp0input"
if not exist "%~dp0output" mkdir "%~dp0output"
"%~dp0python\python.exe" "%~dp0process.py" %*
echo.
pause
