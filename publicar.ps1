<#
  =====================================================================
   OSTEL TRANSPORTES - DASHBOARD VENDAS
   publicar.ps1  |  PUBLICAÇÃO NO SITE VIA FTP / FTPS  (OPCIONAL)
  ---------------------------------------------------------------------
   Envia index.html, style.css, app.js, dados.js, dados_status.js e a
   pasta assets para o servidor do site.
   Host, usuário e pasta ficam no config.ps1 (FtpHost, FtpUsuario...).

   A SENHA NUNCA FICA NO CÓDIGO: na primeira execução ela é pedida numa
   janela e guardada criptografada (Export-Clixml / DPAPI do Windows) em
   credenciais\ftp.xml. Só o MESMO usuário, no MESMO computador,
   consegue abrir esse arquivo.

   Uso:
     publicar.ps1                  -> publica (pede a senha na 1ª vez)
     publicar.ps1 -RedefinirSenha  -> pede a senha de novo
     publicar.ps1 -NaoInterativo   -> nunca pede senha (uso automático)
  =====================================================================
#>
param([switch]$RedefinirSenha, [switch]$NaoInterativo, [switch]$Silencioso)
$ErrorActionPreference = 'Stop'

$Pasta     = Split-Path -Parent $MyInvocation.MyCommand.Path
$PastaLogs = Join-Path $Pasta 'logs'
$ArqLog    = Join-Path $PastaLogs 'publicacao.log'
$ArqCred   = Join-Path (Join-Path $Pasta 'credenciais') 'ftp.xml'
if (-not (Test-Path -LiteralPath $PastaLogs)) { New-Item -ItemType Directory -Path $PastaLogs | Out-Null }

$Config = @{}
. (Join-Path $Pasta 'config.ps1')
foreach ($k in @{ FtpPorta = 21; FtpPasta = '/public_html/vendas'; FtpUsarSSL = $true; FtpPassivo = $true }.GetEnumerator()) {
    if ($null -eq $Config[$k.Key]) { $Config[$k.Key] = $k.Value }
}

function Log([string]$msg, [string]$nivel = 'INFO') {
    $linha = '[{0}] [{1}] {2}' -f (Get-Date -Format 'dd/MM/yyyy HH:mm:ss'), $nivel, $msg
    try { [IO.File]::AppendAllText($ArqLog, $linha + [Environment]::NewLine, (New-Object Text.UTF8Encoding($true))) } catch { }
    if (-not $Silencioso) {
        $cor = switch ($nivel) { 'ERRO' { 'Red' } 'AVISO' { 'Yellow' } 'OK' { 'Green' } default { 'Gray' } }
        Write-Host $linha -ForegroundColor $cor
    }
}

# ---- credencial criptografada --------------------------------------------
function Obter-Credencial {
    if ($RedefinirSenha -and (Test-Path -LiteralPath $ArqCred)) { Remove-Item -LiteralPath $ArqCred -Force }
    if (Test-Path -LiteralPath $ArqCred) {
        try { return Import-Clixml -LiteralPath $ArqCred }
        catch { throw 'Não foi possível abrir a senha salva (foi criada por outro usuário ou computador?). Rode: publicar.ps1 -RedefinirSenha' }
    }
    if ($NaoInterativo) { throw 'Senha do FTP ainda não cadastrada. Execute o publicar.ps1 uma vez manualmente.' }
    $cred = Get-Credential -UserName ([string]$Config.FtpUsuario) -Message ('Senha do FTP de ' + $Config.FtpHost + '. Ela será guardada criptografada neste computador.')
    if (-not $cred) { throw 'Publicação cancelada: nenhuma senha informada.' }
    $p = Split-Path -Parent $ArqCred
    if (-not (Test-Path -LiteralPath $p)) { New-Item -ItemType Directory -Path $p -Force | Out-Null }
    $cred | Export-Clixml -LiteralPath $ArqCred            # DPAPI: só este usuário/PC consegue ler
    Log "Senha do FTP salva (criptografada) em $ArqCred" 'OK'
    return $cred
}

# ---- operações FTP ---------------------------------------------------------
function Requisicao([string]$remoto, [string]$metodo) {
    $req = [Net.FtpWebRequest]::Create(('ftp://{0}:{1}{2}' -f $Config.FtpHost, $Config.FtpPorta, $remoto))
    $req.Method = $metodo
    $req.Credentials = $script:Cred.GetNetworkCredential()
    $req.EnableSsl = [bool]$Config.FtpUsarSSL
    $req.UsePassive = [bool]$Config.FtpPassivo
    $req.UseBinary = $true; $req.KeepAlive = $false; $req.Timeout = 60000
    return $req
}
function Criar-Pasta([string]$caminho) {
    $acum = ''
    foreach ($parte in ($caminho.Trim('/') -split '/')) {
        if (-not $parte) { continue }
        $acum += '/' + $parte
        try { (Requisicao ($acum + '/') ([Net.WebRequestMethods+Ftp]::MakeDirectory)).GetResponse().Close() } catch { }   # "já existe" é ignorado
    }
}
function Enviar([string]$local, [string]$remoto) {
    $compart = [IO.FileShare]([int][IO.FileShare]::ReadWrite -bor [int][IO.FileShare]::Delete)
    $fs = [IO.File]::Open($local, [IO.FileMode]::Open, [IO.FileAccess]::Read, $compart)
    try { $ms = New-Object IO.MemoryStream; $fs.CopyTo($ms); [byte[]]$bytes = $ms.ToArray() } finally { $fs.Dispose() }
    $req = Requisicao $remoto ([Net.WebRequestMethods+Ftp]::UploadFile)
    $req.ContentLength = $bytes.Length
    $s = $req.GetRequestStream(); try { $s.Write($bytes, 0, $bytes.Length) } finally { $s.Close() }
    $req.GetResponse().Close()
}
function Renomear([string]$de, [string]$para) {
    try { (Requisicao $para ([Net.WebRequestMethods+Ftp]::DeleteFile)).GetResponse().Close() } catch { }
    $req = Requisicao $de ([Net.WebRequestMethods+Ftp]::Rename); $req.RenameTo = [IO.Path]::GetFileName($para)
    $req.GetResponse().Close()
}

$codigo = 0
try {
    if (-not ([string]$Config.FtpHost).Trim()) { throw 'Preencha FtpHost, FtpUsuario e FtpPasta no config.ps1 antes de publicar.' }
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $script:Cred = Obter-Credencial
    $base = '/' + ([string]$Config.FtpPasta).Trim('/')
    Log ('Publicando em ftp' + $(if ($Config.FtpUsarSSL) { 's' } else { '' }) + '://' + $Config.FtpHost + $base + ' ...')
    Criar-Pasta $base; Criar-Pasta ($base + '/assets'); Criar-Pasta ($base + '/assets/img')
    $arquivos = @('index.html', 'style.css', 'app.js', 'dados_status.js', 'assets\chart.umd.min.js', 'assets\mapa-brasil.js', 'assets\img\logo.png', 'assets\img\favicon.svg')
    foreach ($rel in $arquivos) {
        $local = Join-Path $Pasta $rel
        if (-not [IO.File]::Exists($local)) { Log "  (não encontrado, pulado) $rel" 'AVISO'; continue }
        Enviar $local ($base + '/' + ($rel -replace '\\', '/'))
        Log "  enviado: $rel"
    }
    # dados.js por último e por nome temporário: o site nunca serve um arquivo pela metade
    $dados = Join-Path $Pasta 'dados.js'
    if ([IO.File]::Exists($dados)) { Enviar $dados ($base + '/dados.js.tmp'); Renomear ($base + '/dados.js.tmp') ($base + '/dados.js'); Log '  enviado: dados.js' }
    Log 'Publicação concluída.' 'OK'
}
catch { $codigo = 1; Log ('ERRO na publicação: ' + $_.Exception.Message) 'ERRO' }
exit $codigo
