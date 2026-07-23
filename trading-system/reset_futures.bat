@echo off
REM Limpa a conta de FUTUROS testnet: cancela ordens e fecha posições.
REM Use quando o bot acumular posições/ordens entre reinícios.
title Reset da conta de Futuros
cd /d %~dp0
call .venv\Scripts\activate
python reset.py config/config.futures.yaml
echo.
echo Pronto. Agora pode ligar o bot com start_futures.bat
pause
