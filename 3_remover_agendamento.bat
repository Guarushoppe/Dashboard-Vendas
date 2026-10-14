@echo off
REM ==============================================================
REM  OSTEL - DASHBOARD VENDAS
REM  3) REMOVER AGENDAMENTO
REM  Apaga a tarefa criada pelo 2_agendar_atualizacao.bat
REM ==============================================================
title Ostel - Remover agendamento
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0atualizar.ps1" -RemoverAgendamento
echo.
pause
