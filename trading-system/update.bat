@echo off
REM Atualiza o bot para a versao mais recente do GitHub com UM clique.
REM Preserva o seu .env (as chaves nunca sao tocadas).
title Atualizador do Bot
cd /d %~dp0
echo Baixando a versao mais recente...
curl -L -s -o update.zip https://github.com/rodrigozambon20-del/TCCENGENHARIASEGURAN-A/archive/refs/heads/claude/multi-agent-trading-system-8bwrtb.zip
if errorlevel 1 goto erro
echo Extraindo arquivos...
tar -xf update.zip --strip-components=2 "TCCENGENHARIASEGURAN-A-claude-multi-agent-trading-system-8bwrtb/trading-system"
if errorlevel 1 goto erro
del update.zip
echo Instalando dependencias novas (se houver)...
call .venv\Scripts\activate
python -m pip install -q -r requirements.txt
echo.
echo ============================================
echo  Atualizacao concluida! Chaves preservadas.
echo  Pode ligar o bot com start_bot.bat
echo ============================================
pause
exit /b 0
:erro
echo.
echo Algo falhou no download/extracao. Verifique sua internet e tente de novo.
pause
exit /b 1
