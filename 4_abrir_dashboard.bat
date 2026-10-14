@echo off
REM ==============================================================
REM  OSTEL - DASHBOARD VENDAS
REM  4) ABRIR O PAINEL
REM  Atualiza os dados (se alguma planilha mudou) e abre o index.html
REM  no navegador padrao.
REM ==============================================================
title Ostel - Abrir painel
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0atualizar.ps1"
start "" "%~dp0index.html"
exit /b 0
