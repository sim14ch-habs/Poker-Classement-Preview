@echo off
setlocal
title Validation Poker - OCR (Preview)

set "SCRIPT_DIR=%~dp0"
set "PS1=%SCRIPT_DIR%routines\validation_photo_classement_ui.ps1"
set "XLSX=%SCRIPT_DIR%Poker Stanley Hiver 2026.xlsx"
set "PSH=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "EMAIL_TO=sim_621@hotmail.com"
set "PUBLIC_URL=https://sim14ch-habs.github.io/Poker-Classement-Preview/"
set "PUBLISH_REPO_DIR=%SCRIPT_DIR%Poker-Classement-preview"

if not "%~1"=="" set "XLSX=%~1"

echo.
echo [INFO] Demarrage de la validation photo OCR (PREVIEW)...
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
  echo Astuce: glisse-depose un .xlsx sur ce .bat.
  pause
  exit /b 1
)

if not exist "%PUBLISH_REPO_DIR%" (
  echo [ERROR] Dossier repo preview introuvable: "%PUBLISH_REPO_DIR%"
  echo Cree ce dossier local ou re-clone le repo preview.
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
  echo [ERROR] OPENAI_API_KEY introuvable.
  pause
  exit /b 1
)

echo [INFO] Ouverture de l'interface OCR (attendre 5-20 secondes)...
"%PSH%" -NoProfile -NonInteractive -ExecutionPolicy RemoteSigned -File "%PS1%" -WorkbookPath "%XLSX%" -EmailTo "%EMAIL_TO%" -PublicUrl "%PUBLIC_URL%" -PublishRepoDir "%PUBLISH_REPO_DIR%"

set "RC=%ERRORLEVEL%"
echo.

if "%RC%"=="2" (
  echo [INFO] Une autre fenetre de validation est deja ouverte.
  pause
  exit /b 0
)

if not "%RC%"=="0" (
  echo [ERROR] La validation preview a echoue. code=%RC%.
  pause
  exit /b %RC%
)

echo [OK] Preview terminee.
pause
exit /b 0
