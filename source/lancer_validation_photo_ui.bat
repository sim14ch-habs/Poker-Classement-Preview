@echo off
setlocal
title Validation Poker - OCR

set "SCRIPT_DIR=%~dp0"
set "PS1=%SCRIPT_DIR%routines\validation_photo_classement_ui.ps1"
set "XLSX=%SCRIPT_DIR%Poker Stanley Hiver 2026.xlsx"
set "PSH=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "EMAIL_TO=sim_621@hotmail.com"
set "PUBLIC_URL=https://sim14ch-habs.github.io/Poker-Classement/"
set "PUBLISH_REPO_DIR=%SCRIPT_DIR%Poker-Classement"
set "PREVIEW_PUBLIC_URL=https://sim14ch-habs.github.io/Poker-Classement-Preview/"
set "PREVIEW_PUBLISH_REPO_DIR=%SCRIPT_DIR%Poker-Classement-preview"

if not "%~1"=="" set "XLSX=%~1"

echo.
echo [INFO] Demarrage de la validation photo OCR...
echo [INFO] Script   : "%PS1%"
echo [INFO] Classeur : "%XLSX%"
if not "%PUBLIC_URL%"=="" echo [INFO] URL publique : "%PUBLIC_URL%"
echo [INFO] Repo web : "%PUBLISH_REPO_DIR%"
echo [INFO] Sync preview : "%PREVIEW_PUBLISH_REPO_DIR%"
echo.

if not exist "%PS1%" (
  echo [ERROR] Script introuvable: "%PS1%"
  pause
  exit /b 1
)

if not exist "%XLSX%" (
  echo [ERROR] Classeur introuvable: "%XLSX%"
  echo Astuce: glisse-depose un .xlsx sur ce .bat.
  pause
  exit /b 1
)

if not defined OPENAI_API_KEY (
  for /f "tokens=2,*" %%A in ('reg query "HKCU\Environment" /v OPENAI_API_KEY 2^>nul ^| find /i "OPENAI_API_KEY"') do set "OPENAI_API_KEY=%%B"
)

if not defined OPENAI_API_KEY (
  for /f "tokens=2,*" %%A in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v OPENAI_API_KEY 2^>nul ^| find /i "OPENAI_API_KEY"') do set "OPENAI_API_KEY=%%B"
)

if not defined OPENAI_API_KEY (
  echo [ERROR] OPENAI_API_KEY introuvable. La version conservee est OCR seulement.
  echo Lance setup une fois, puis relance ce .bat.
  pause
  exit /b 1
)

echo [INFO] Ouverture de l'interface OCR (attendre 5-20 secondes)...
"%PSH%" -NoProfile -NonInteractive -ExecutionPolicy RemoteSigned -File "%PS1%" -WorkbookPath "%XLSX%" -EmailTo "%EMAIL_TO%" -PublicUrl "%PUBLIC_URL%" -PublishRepoDir "%PUBLISH_REPO_DIR%" -MirrorPreviewPublicUrl "%PREVIEW_PUBLIC_URL%" -MirrorPreviewPublishRepoDir "%PREVIEW_PUBLISH_REPO_DIR%"

set "RC=%ERRORLEVEL%"
echo.

if "%RC%"=="2" (
  echo [INFO] Une autre fenetre de validation est deja ouverte.
  pause
  exit /b 0
)

if not "%RC%"=="0" (
  echo [ERROR] La validation a echoue. code=%RC%.
  echo Si la fenetre n'apparait pas: ferme Excel puis relance ce .bat.
  pause
  exit /b %RC%
)

echo [OK] Termine.
pause
exit /b 0


