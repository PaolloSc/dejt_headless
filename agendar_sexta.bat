@echo off
chcp 65001 >nul
title Agendar tarefa - DEJT Sexta-feira

echo Criando tarefa agendada para toda sexta-feira as 08:00...
echo.

schtasks /create /tn "DEJT_JT_Sexta" /tr "\"%~dp0executar_sexta.bat\"" /sc weekly /d FRI /st 08:00 /rl HIGHEST /f

if %errorlevel% equ 0 (
    echo.
    echo [OK] Tarefa "DEJT_JT_Sexta" criada com sucesso!
    echo     Executa toda sexta-feira as 08:00
    echo.
    echo Para alterar horario:
    echo     schtasks /change /tn "DEJT_JT_Sexta" /st 09:00
    echo.
    echo Para remover:
    echo     schtasks /delete /tn "DEJT_JT_Sexta" /f
) else (
    echo.
    echo [ERRO] Falha ao criar tarefa. Execute como Administrador.
)

echo.
pause
