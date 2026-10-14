@echo off
REM ==============================================================
REM  OSTEL - DASHBOARD VENDAS
REM  2) AGENDAR ATUALIZACAO AUTOMATICA
REM  Cria uma tarefa no Agendador de Tarefas do Windows que confere
REM  as planilhas a cada X minutos (IntervaloMinutos no config.ps1,
REM  padrao 5) e ao entrar no Windows. Nao precisa ser administrador.
REM  Quando uma planilha muda, o Gerar_Versao_Para_Enviar.bat roda sozinho.
REM ==============================================================
title Ostel - Agendar atualizacao automatica
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-ChildItem -LiteralPath '%~dp0' -Recurse -File | Unblock-File" >nul 2>nul
echo.
echo  Criando a tarefa agendada...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0atualizar.ps1" -Agendar
if errorlevel 1 (
  echo.
  echo  Nao foi possivel criar a tarefa. Veja o LEIAME.txt, secao "Solucao de problemas".
) else (
  echo.
  echo  Tarefa criada! Para conferir: menu Iniciar ^> "Agendador de Tarefas".
  echo  Para mudar o intervalo, edite IntervaloMinutos no config.ps1 e rode este arquivo de novo.
)
echo.
pause
