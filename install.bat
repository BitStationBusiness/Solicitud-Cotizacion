@echo off
setlocal EnableExtensions EnableDelayedExpansion
title BitStation Multimedia Downloader - Installer/Launcher
cd /d "%~dp0"

REM ===== Config =====
set "APPNAME=Solicitud Cotizacion Proveedores"
set "ICON=%CD%\ico\request.ico"
set "PYEXE=%CD%\venv\Scripts\python.exe"
set "PYWEXE=%CD%\venv\Scripts\pythonw.exe"
set "MAIN=%CD%\main.py"

echo.
echo [%APPNAME%] Preparando entorno...

REM ===== Crear venv si no existe =====
if not exist "%PYEXE%" (
  echo Creando entorno virtual...
  where py >nul 2>nul && ( py -3 -m venv venv ) || ( python -m venv venv )
  if errorlevel 1 (
    echo No se pudo crear el entorno virtual.
    pause
    exit /b 1
  )
)

REM ===== Actualizar pip y dependencias =====
echo Actualizando pip y requirements...
"%PYEXE%" -m pip install --upgrade pip
if exist requirements.txt (
  "%PYEXE%" -m pip install -r requirements.txt
)

REM ===== Crear accesos directos (Escritorio y Menú Inicio) =====
echo Creando accesos directos...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$shell=New-Object -ComObject WScript.Shell; " ^
  "$desktop=[Environment]::GetFolderPath('Desktop'); " ^
  "$lnk=$shell.CreateShortcut( (Join-Path $desktop '%APPNAME%.lnk') ); " ^
  "$lnk.TargetPath='%PYWEXE%'; $lnk.Arguments='\"%MAIN%\"'; " ^
  "$lnk.WorkingDirectory='%CD%'; " ^
  "$lnk.IconLocation='%ICON%'; " ^
  "$lnk.Description='%APPNAME%'; $lnk.Save(); " ^
  "$prog=[Environment]::GetFolderPath('Programs'); $dir=Join-Path $prog 'BitStation'; New-Item -ItemType Directory -Force -Path $dir ^| Out-Null; " ^
  "$lnk2=$shell.CreateShortcut( (Join-Path $dir '%APPNAME%.lnk') ); " ^
  "$lnk2.TargetPath='%PYWEXE%'; $lnk2.Arguments='\"%MAIN%\"'; " ^
  "$lnk2.WorkingDirectory='%CD%'; $lnk2.IconLocation='%ICON%'; " ^
  "$lnk2.Description='%APPNAME%'; $lnk2.Save(); "

REM ===== Modo opcional: build EXE con PyInstaller =====
if /I "%~1"=="--build-exe" (
  echo.
  echo Compilando EXE con PyInstaller...
  "%PYEXE%" -m pip install --upgrade pyinstaller
  "%PYEXE%" -m PyInstaller --noconfirm --clean --onefile --windowed --icon "%ICON%" --name "BitStation" "%MAIN%"
  if exist "%CD%\dist\BitStation.exe" (
    echo Creando acceso directo al EXE compilado...
    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
      "$shell=New-Object -ComObject WScript.Shell; " ^
      "$desktop=[Environment]::GetFolderPath('Desktop'); " ^
      "$lnk=$shell.CreateShortcut( (Join-Path $desktop '%APPNAME% (EXE).lnk') ); " ^
      "$lnk.TargetPath=(Join-Path '%CD%' 'dist\\BitStation.exe'); " ^
      "$lnk.WorkingDirectory='%CD%'; $lnk.IconLocation='%ICON%'; $lnk.Description='%APPNAME%'; $lnk.Save(); "
  ) else (
    echo No se encontro el EXE en dist\. Continuo sin accesos directos del EXE.
  )
)

echo.
echo Iniciando la aplicacion...
REM >>> Lanzar sin consola y cerrar esta ventana inmediatamente <<<
if exist "%PYWEXE%" (
  start "" "%PYWEXE%" "%MAIN%"
) else (
  REM fallback si pythonw no existe
  start "" "%PYEXE%" "%MAIN%"
)
exit /b 0
