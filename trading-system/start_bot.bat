@echo off
REM Liga o bot de trading. Funciona de qualquer lugar: sempre entra na
REM pasta onde este arquivo esta salvo (ex: D:\bot).
title Bot de Trading - Binance Testnet
cd /d %~dp0
call .venv\Scripts\activate
python main.py
echo.
echo O bot foi encerrado. Pressione qualquer tecla para fechar.
pause >nul
