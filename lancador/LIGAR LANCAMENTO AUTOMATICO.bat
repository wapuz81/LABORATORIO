@echo off
rem Cria a tarefa agendada que confere a pasta ENTRADA a cada 5 minutos (so neste PC, so com voce logado)
schtasks /Create /F /SC MINUTE /MO 5 /TN "Lancador Ensaios Campo" /TR "wscript.exe \"%~dp0Vigiar-Entrada.vbs\""
echo.
echo Lancamento automatico LIGADO neste computador. Ligue em UM computador so.
pause
