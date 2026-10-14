@echo off
rem =====================================================================
rem  OSTEL - DASHBOARD VENDAS
rem  GERAR VERSAO PARA ENVIAR
rem  Cria UM UNICO arquivo HTML com tudo dentro (dados, graficos, mapa,
rem  logo) que abre em qualquer computador ou celular, mesmo sem internet:
rem      enviar\Vendas.html   (SEMPRE o mesmo nome - o anterior e trocado)
rem  Este arquivo e iniciado AUTOMATICAMENTE pelo atualizar.ps1 sempre
rem  que uma planilha e atualizada (com o parametro /silencioso).
rem  Dois cliques aqui: confere as planilhas e gera na hora.
rem =====================================================================
setlocal
cd /d "%~dp0"
if /i "%~1"=="/silencioso" goto :silencioso

title Ostel - Gerar versao para enviar
echo.
echo  1/2  Conferindo se as planilhas mudaram...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0atualizar.ps1" -SemEnvio
echo.
echo  2/2  Gerando a versao para envio...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0gerar_versao.ps1" -AbrirPasta
if errorlevel 1 (
  echo.
  echo  [ERRO] Nao foi possivel gerar. Veja logs\atualizacao.log
) else (
  echo.
  echo  [OK] Pronto! Envie o arquivo da pasta "enviar" por e-mail, WhatsApp ou Teams.
)
echo.
pause
endlocal
exit /b

:silencioso
powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0gerar_versao.ps1" -Silencioso
exit /b %errorlevel%
