@echo off
chcp 65001 >nul
title DEJT + JT Juris - Sexta-feira

cd /d "%~dp0"

REM Usa o Python do venv se existir, senão usa o do PATH
if exist "..\jurisgen\backend\venv\Scripts\python.exe" (
    set PYTHON="..\jurisgen\backend\venv\Scripts\python.exe"
) else (
    set PYTHON=python
)

echo ============================================================
echo  DEJT + JT Juris + WhatsApp - Execucao automatica
echo  %date% %time%
echo ============================================================

%PYTHON% executar_sexta.py

if %errorlevel% neq 0 (
    echo.
    echo [ERRO] Execucao falhou com codigo %errorlevel%
    pause
    exit /b %errorlevel%
)

echo.
echo [OK] Concluido com sucesso!
timeout /t 5
