<#
  =====================================================================
   OSTEL TRANSPORTES - DASHBOARD VENDAS
   gerar_versao.ps1  |  VERSÃO PARA ENVIO (arquivo único)
  ---------------------------------------------------------------------
   Junta o painel inteiro (página, visual, lógica, gráficos, mapa, logo
   e DADOS) em UM só arquivo .html, para mandar por e-mail, WhatsApp ou
   Teams. Abre em qualquer computador ou celular, sem internet e sem
   instalar nada, com filtros, tabela e botões funcionando. Mostra os
   dados do momento em que foi gerado.

   O nome é SEMPRE O MESMO (NomeArquivoEnvio no config.ps1, padrão
   "Vendas.html"): o arquivo anterior é substituído.

   Quem chama: Gerar_Versao_Para_Enviar.bat (dois cliques) e o
   atualizar.ps1, sozinho, sempre que uma planilha muda.
   Uso: gerar_versao.ps1 [-Silencioso] [-AbrirPasta]
  =====================================================================
#>
param([switch]$Silencioso, [switch]$AbrirPasta)
$ErrorActionPreference = 'Stop'

$Pasta     = Split-Path -Parent $MyInvocation.MyCommand.Path
$PastaLogs = Join-Path $Pasta 'logs'
$ArqLog    = Join-Path $PastaLogs 'atualizacao.log'
$Utf8      = New-Object Text.UTF8Encoding($false)
if (-not (Test-Path -LiteralPath $PastaLogs)) { New-Item -ItemType Directory -Path $PastaLogs | Out-Null }

$Config = @{}
$ArqConfig = Join-Path $Pasta 'config.ps1'
if (Test-Path -LiteralPath $ArqConfig) { . $ArqConfig }
if (-not $Config.NomeArquivoEnvio) { $Config.NomeArquivoEnvio = 'Vendas.html' }

function Log([string]$msg, [string]$nivel = 'INFO') {
    $linha = '[{0}] [{1}] {2}' -f (Get-Date -Format 'dd/MM/yyyy HH:mm:ss'), $nivel, ('Envio: ' + $msg)
    try { [IO.File]::AppendAllText($ArqLog, $linha + [Environment]::NewLine, (New-Object Text.UTF8Encoding($true))) } catch { }
    if (-not $Silencioso) {
        $cor = switch ($nivel) { 'ERRO' { 'Red' } 'AVISO' { 'Yellow' } 'OK' { 'Green' } default { 'Gray' } }
        Write-Host $linha -ForegroundColor $cor
    }
}
function Ler([string]$rel) { return [IO.File]::ReadAllText((Join-Path $Pasta $rel), [Text.Encoding]::UTF8) }
function Existe([string]$rel) { return [IO.File]::Exists((Join-Path $Pasta $rel)) }
# dentro de <script>, o texto "</script" encerraria o bloco antes da hora
function ScriptInline([string]$js) { return "<script>`r`n" + $js.Replace('</script', '<\/script').Replace('</SCRIPT', '<\/SCRIPT') + "`r`n</script>" }

# só uma geração por vez (o atualizar.ps1 e o clique manual podem coincidir)
$nomeTrava = 'Local\OstelTorreSACVendasEnvio_' + ($Pasta.ToLowerInvariant() -replace '[^a-z0-9]', '_')
$trava = New-Object Threading.Mutex -ArgumentList $false, $nomeTrava
$temTrava = $false
try { $temTrava = $trava.WaitOne(120000) } catch [Threading.AbandonedMutexException] { $temTrava = $true }

$codigo = 0
try {
    foreach ($f in @('index.html', 'style.css', 'app.js', 'dados.js')) {
        if (-not (Existe $f)) { throw "Falta o arquivo $f na pasta do painel. Execute 1_atualizar_agora.bat primeiro." }
    }
    $html = Ler 'index.html'

    # 1) visual
    $html = $html.Replace('<link rel="stylesheet" href="style.css">', "<style>`r`n" + (Ler 'style.css') + "`r`n</style>")
    # 2) logo e ícone viram texto (base64) dentro do próprio arquivo
    if (Existe 'assets\img\logo.png') {
        $html = $html.Replace('src="assets/img/logo.png"', 'src="data:image/png;base64,' + [Convert]::ToBase64String([IO.File]::ReadAllBytes((Join-Path $Pasta 'assets\img\logo.png'))) + '"')
    }
    if (Existe 'assets\img\favicon.svg') {
        $html = $html.Replace('href="assets/img/favicon.svg"', 'href="data:image/svg+xml;base64,' + [Convert]::ToBase64String([IO.File]::ReadAllBytes((Join-Path $Pasta 'assets\img\favicon.svg'))) + '"')
    }
    # 3) biblioteca de gráficos: sai o endereço da internet, entra a cópia local
    $padraoChart = '(?s)<script src="https://cdn\.jsdelivr\.net/npm/chart\.js[^"]*"></script>\s*<script>.*?</script>'
    if (Existe 'assets\chart.umd.min.js') {
        $bloco = ScriptInline (Ler 'assets\chart.umd.min.js')
        $html = [regex]::Replace($html, $padraoChart, [Text.RegularExpressions.MatchEvaluator] { param($m) $bloco })
    } else { Log 'assets\chart.umd.min.js não encontrado: o arquivo de envio vai precisar de internet para os gráficos.' 'AVISO' }
    # 4) mapa
    if (Existe 'assets\mapa-brasil.js') { $html = $html.Replace('<script src="assets/mapa-brasil.js"></script>', (ScriptInline (Ler 'assets\mapa-brasil.js'))) }
    # 5) dados + lógica (OSTEL_EMBUTIDO avisa o app.js para usar os dados que estão no próprio arquivo)
    $marca = '<script src="app.js"></script>'
    if (-not $html.Contains($marca)) { throw 'Não achei a chamada do app.js no index.html.' }
    $html = $html.Replace($marca, "<script>window.OSTEL_EMBUTIDO = true;</script>`r`n" + (ScriptInline (Ler 'dados.js')) + "`r`n" + (ScriptInline (Ler 'app.js')))

    # 6) grava com nome fixo (temporário + troca)
    $pastaSaida = ([string]$Config.PastaArquivoEnvio).Trim().Trim('"')
    if (-not $pastaSaida) { $pastaSaida = Join-Path $Pasta 'enviar' }
    if (-not (Test-Path -LiteralPath $pastaSaida)) { New-Item -ItemType Directory -Path $pastaSaida -Force | Out-Null }
    $nome = ([string]$Config.NomeArquivoEnvio).Trim(); if ($nome -notmatch '\.html?$') { $nome += '.html' }
    $destino = Join-Path $pastaSaida $nome
    $tmp = $destino + '.tmp'
    [IO.File]::WriteAllText($tmp, $html, $Utf8)
    try { [IO.File]::Copy($tmp, $destino, $true) }
    catch { throw "Não foi possível substituir $nome (está aberto em outro programa?). Feche-o e gere de novo." }
    finally { try { [IO.File]::Delete($tmp) } catch { } }
    $kb = [Math]::Round((New-Object IO.FileInfo($destino)).Length / 1KB)
    Log "Arquivo para envio gerado: $destino ($kb KB)" 'OK'
    if (-not $Silencioso) { Write-Host '  Atenção: o arquivo tem nomes de clientes e valores. Envie só para quem pode ver.' -ForegroundColor Yellow }
    if ($AbrirPasta -and $env:OS -eq 'Windows_NT') { Start-Process explorer.exe $pastaSaida }
}
catch {
    Log ('Arquivo para envio NÃO gerado: ' + $_.Exception.Message) 'ERRO'
    $codigo = 1
}
finally { if ($temTrava) { try { $trava.ReleaseMutex() } catch { } } }
exit $codigo
