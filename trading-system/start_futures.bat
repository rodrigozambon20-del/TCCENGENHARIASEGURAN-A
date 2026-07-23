@echo off
REM Liga o bot de FUTUROS (long+short, alavancagem 3x) em testnet.
REM Roda em paralelo ao spot — pode ter as duas janelas abertas ao mesmo tempo.
REM Requer chaves de FUTUROS testnet em .env.futures (veja README).
title Bot de Trading - FUTUROS Testnet (long+short)
cd /d %~dp0
call .venv\Scripts\activate
REM Carrega as chaves de futuros (arquivo separado do spot)
if exist .env.futures (
    for /f "usebackq tokens=1,* delims==" %%a in (".env.futures") do set %%a=%%b
)
python main.py config/config.futures.yaml
echo.
echo O bot de futuros foi encerrado. Pressione qualquer tecla para fechar.
pause >nul
