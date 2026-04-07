@echo off
setlocal
title Export Classement (Preview)

set "SCRIPT_DIR=%~dp0"
set "PS1=%SCRIPT_DIR%routines\export_classement_hebdo.ps1"
set "XLSX=%SCRIPT_DIR%Poker Stanley Hiver 2026.xlsx"
set "PSH=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "PUBLIC_URL=https://sim14ch-habs.github.io/Poker-Classement-preview/"
set "PUBLISH_REPO_DIR=%SCRIPT_DIR%Poker-Classement-preview"

if not "%~1"=="" set "XLSX=%~1"

echo.
echo [INFO] Export preview...
echo [INFO] Script   : "%PS1%"
echo [INFO] Classeur : "%XLSX%"
echo [INFO] URL preview : "%PUBLIC_URL%"
echo [INFO] Repo preview: "%PUBLISH_REPO_DIR%"
echo.

if not exist "%PS1%" (
  echo [ERROR] Script introuvable: "%PS1%"
  pause
  exit /b 1
)

if not exist "%XLSX%" (
  echo [ERROR] Classeur introuvable: "%XLSX%"
  pause
  exit /b 1
)

if not exist "%PUBLISH_REPO_DIR%" (
  echo [ERROR] Dossier repo preview introuvable: "%PUBLISH_REPO_DIR%"
  pause
  exit /b 1
)

"%PSH%" -NoProfile -NonInteractive -ExecutionPolicy RemoteSigned -File "%PS1%" -WorkbookPath "%XLSX%" -PublicUrl "%PUBLIC_URL%" -PublishRepoDir "%PUBLISH_REPO_DIR%"
set "RC=%ERRORLEVEL%"

echo.
if not "%RC%"=="0" (
  echo [ERROR] Export preview echoue. code=%RC%.
  pause
  exit /b %RC%
)

echo [OK] Export preview termine.
pause
exit /b 0
