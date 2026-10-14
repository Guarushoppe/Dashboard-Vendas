@echo off
REM ==============================================================
REM  OSTEL - DASHBOARD VENDAS
REM  1) ATUALIZAR AGORA
REM  Le as planilhas Comercial CTe e Comercial Cotacoes e gera o
REM  dados.js (mesmo que nada tenha mudado).
REM ==============================================================
title Ostel - Atualizar painel agora
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-ChildItem -LiteralPath '%~dp0' -Recurse -File | Unblock-File" >nul 2>nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0atualizar.ps1" -Forcar
set "CODIGO=%ERRORLEVEL%"
echo.
if "%CODIGO%"=="0" (
  echo  Pronto! O painel aberto no navegador se atualiza sozinho em ate 30 segundos.
  timeout /t 6 >nul
  exit /b 0
)
if "%CODIGO%"=="2" echo  ATENCAO: planilha nao encontrada. Confira os caminhos no config.ps1.
if "%CODIGO%"=="3" echo  ATENCAO: planilha invalida. Veja a mensagem acima e o LEIAME.txt.
if "%CODIGO%"=="1" echo  ATENCAO: falha temporaria (planilha em uso?). Tente de novo em instantes.
echo  O detalhe completo esta em logs\atualizacao.log
echo.
pause
exit /b %CODIGO%
