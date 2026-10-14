@echo off
REM ==============================================================
REM  OSTEL - DASHBOARD VENDAS
REM  5) PUBLICAR NO SITE (opcional)
REM  Envia o painel para o site por FTP/FTPS (dados em config.ps1).
REM  Na primeira vez pede a senha e guarda criptografada.
REM ==============================================================
title Ostel - Publicar no site
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0publicar.ps1"
echo.
pause
