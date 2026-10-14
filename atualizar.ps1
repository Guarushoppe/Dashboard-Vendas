<#
  =====================================================================
   OSTEL TRANSPORTES - DASHBOARD VENDAS - ATUALIZADOR
  =====================================================================
   O que este script faz:
     1. Lê as duas planilhas (Comercial CTe e Comercial Cotações) SEM
        precisar do Excel - inclusive com elas abertas. As planilhas são
        só lidas, nunca alteradas;
     2. Aceita também CSV/TXT: detecta sozinho o separador (; , tab),
        a codificação (UTF-8 ou Windows-1252/ANSI do Excel) e o formato
        brasileiro (datas dd/mm/aaaa, 1.234,56);
     3. Lê as metas do mês (dados\metas.csv);
     4. Grava o dados.js (primeiro num arquivo temporário e depois troca,
        para o painel nunca ler um arquivo pela metade);
     5. Grava o dados_status.js a cada execução (o painel usa para avisar
        quando a atualização automática parou);
     6. Quando alguma planilha mudou, inicia sozinho o
        Gerar_Versao_Para_Enviar.bat (arquivo único para envio);
     7. Registra tudo em logs\atualizacao.log.

   Uso:
     atualizar.ps1                      -> atualiza (pula se nada mudou)
     atualizar.ps1 -Forcar              -> atualiza mesmo sem mudanças
     atualizar.ps1 -Agendar             -> cria a tarefa no Agendador do Windows
     atualizar.ps1 -RemoverAgendamento  -> remove a tarefa agendada
     atualizar.ps1 -Silencioso          -> sem mensagens na tela (usado pela tarefa)
     atualizar.ps1 -SemEnvio            -> não inicia o Gerar_Versao_Para_Enviar.bat

   Códigos de saída: 0 = ok | 1 = falha temporária (arquivo em uso) |
                     2 = planilha não encontrada | 3 = planilha inválida
   Compatível com Windows PowerShell 5.1 (Windows 10/11) e PowerShell 7.
  =====================================================================
#>
param(
    [switch]$Forcar,
    [switch]$Agendar,
    [switch]$RemoverAgendamento,
    [switch]$Silencioso,
    [switch]$SemEnvio
)
$ErrorActionPreference = 'Stop'

# =====================================================================
#  BLOCO DE CONFIGURAÇÃO DAS FONTES (ajuste aqui se o layout mudar)
#  -------------------------------------------------------------------
#  Para achar a LINHA DE TÍTULOS de cada planilha, o atualizador procura,
#  nas 30 primeiras linhas, a que contém TODAS estas colunas (o nome é
#  comparado sem acento, sem maiúsculas e sem espaços duplos).
#  O significado de cada coluna (vendedora, frete, placa...) é definido
#  no começo do app.js (bloco MAPA), pelo nome ou pela letra da coluna.
# =====================================================================
$Fontes = [ordered]@{
    cte = @{
        Nome        = 'Comercial CTe'
        Obrigatorias= @('Data de Emissao', 'Valor do Frete')
        ColunaData  = 'Data de Emissao'          # linhas sem esta data são ignoradas (com aviso)
    }
    cotacoes = @{
        Nome        = 'Comercial Cotações'
        Obrigatorias= @('COTACAO', 'DATA HORA INCLUSAO')
        ColunaData  = 'DATA HORA INCLUSAO'
    }
}

# ------------------------------------------------------------------ caminhos
$Pasta       = Split-Path -Parent $MyInvocation.MyCommand.Path
$ScriptPath  = $MyInvocation.MyCommand.Path
$PastaDados  = Join-Path $Pasta 'dados'
$PastaLogs   = Join-Path $Pasta 'logs'
$PastaAssets = Join-Path $Pasta 'assets'
$Saida       = Join-Path $Pasta 'dados.js'
$SaidaStatus = Join-Path $Pasta 'dados_status.js'
$ArqEstado   = Join-Path $PastaLogs '.ultimo_estado.txt'
$ArqLog      = Join-Path $PastaLogs 'atualizacao.log'
$Inv         = [Globalization.CultureInfo]::InvariantCulture
$Utf8        = New-Object Text.UTF8Encoding($false)
$Utf8Bom     = New-Object Text.UTF8Encoding($true)

foreach ($p in @($PastaDados, $PastaLogs, $PastaAssets)) {
    if (-not (Test-Path -LiteralPath $p)) { New-Item -ItemType Directory -Path $p | Out-Null }
}

# ------------------------------------------------------------------ configuração (config.ps1)
$Config = @{}
$ArqConfig = Join-Path $Pasta 'config.ps1'
if (Test-Path -LiteralPath $ArqConfig) { . $ArqConfig }
$Padroes = @{
    ArquivoCTe = ''; AbaCTe = ''; ArquivoCotacoes = ''; AbaCotacoes = ''
    ProcurarCTe = 'cte'; ProcurarCotacoes = 'cota'; ArquivoMetas = 'dados\metas.csv'; ArquivoCorrecoes = 'dados\correcoes_vendedora.csv'; ArquivoOcorrencias = 'dados\ocorrencias_ssw.csv'; ArquivoFeriados = 'dados\Feriados.csv'; ArquivoPremios = 'dados\premios.csv'; ArquivoComissao = 'dados\comissao.csv'; Delimitador = ''
    PalavrasData = @('data', 'previsao', 'validade', 'entrega programada')
    PalavrasNumero = @('frete', 'valor', 'peso', 'cubagem', 'qtd', 'quantidade', 'proposta', 'desc', 'distancia')
    PalavrasTexto = @('cnpj', 'cpf', 'serie', 'numero', 'cotacao', 'ctrc', 'nota fiscal', 'placa', 'codigo', 'manifesto', 'romaneio', 'cep', 'telefone', 'contato', 'tabela')
    IntervaloMinutos = 5; NomeTarefa = 'Ostel - Dashboard Vendas'
    GerarEnvioAoMudar = $true; NomeArquivoEnvio = 'Vendas.html'; PastaArquivoEnvio = ''
    PublicarAutomaticamente = $false; FtpHost = ''; LogMaxMB = 5
    UrlChartJs = 'https://cdn.jsdelivr.net/npm/chart.js@4.5.1/dist/chart.umd.min.js'
}
foreach ($k in $Padroes.Keys) { if (-not $Config.ContainsKey($k) -or $null -eq $Config[$k]) { $Config[$k] = $Padroes[$k] } }

# ------------------------------------------------------------------ log (tela + logs\atualizacao.log)
function Log([string]$msg, [string]$nivel = 'INFO') {
    $linha = '[{0}] [{1}] {2}' -f (Get-Date -Format 'dd/MM/yyyy HH:mm:ss'), $nivel, $msg
    try {
        if ([IO.File]::Exists($ArqLog) -and (New-Object IO.FileInfo($ArqLog)).Length -gt ([double]$Config.LogMaxMB * 1MB)) {
            [IO.File]::Copy($ArqLog, $ArqLog + '.1', $true); [IO.File]::Delete($ArqLog)
        }
        [IO.File]::AppendAllText($ArqLog, $linha + [Environment]::NewLine, $Utf8Bom)
    } catch { }
    if (-not $Silencioso) {
        $cor = switch ($nivel) { 'ERRO' { 'Red' } 'AVISO' { 'Yellow' } 'OK' { 'Green' } default { 'Gray' } }
        Write-Host $linha -ForegroundColor $cor
    }
}

# =====================================================================
#  LEITOR (C#, compilado na hora pelo próprio Windows). Lê .xlsx/.xlsm
#  direto do arquivo (um .xlsx é um ZIP com XML dentro), lê CSV e
#  converte tudo para JSON. Escrito em C# 5 (compilador do Windows).
# =====================================================================
$CodigoLeitor = @'
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

public static class OstelVendasLeitor
{
    const string NsRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
    public static int UltimaQtd = 0;
    public static int UltimaSemData = 0;
    public static string UltimaAba = "";

    // ------------------------------------------------------------ XLSX
    // Devolve as linhas (cada uma um object[]: string, double ou null).
    public static List<object[]> LerXlsx(string caminho, string aba)
    {
        using (FileStream fs = new FileStream(caminho, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
        using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read))
        {
            List<string> nomes = new List<string>(); List<string> ids = new List<string>();
            using (XmlReader x = Abrir(zip, "xl/workbook.xml", true))
            {
                while (x.Read())
                    if (x.NodeType == XmlNodeType.Element && x.LocalName == "sheet") { nomes.Add(x.GetAttribute("name")); ids.Add(x.GetAttribute("id", NsRel)); }
            }
            if (nomes.Count == 0) throw new FormatException("A planilha não tem nenhuma aba.");
            int idx = 0;
            if (!string.IsNullOrEmpty(aba))
            {
                idx = -1;
                string a = Simples(aba);
                for (int i = 0; i < nomes.Count; i++) if (Simples(nomes[i]) == a) { idx = i; break; }
                if (idx < 0) throw new FormatException("Aba '" + aba + "' não encontrada. Abas da planilha: " + string.Join(", ", nomes.ToArray()));
            }
            UltimaAba = nomes[idx];
            string alvo = null;
            using (XmlReader x = Abrir(zip, "xl/_rels/workbook.xml.rels", true))
            {
                while (x.Read())
                    if (x.NodeType == XmlNodeType.Element && x.LocalName == "Relationship" && x.GetAttribute("Id") == ids[idx]) alvo = x.GetAttribute("Target");
            }
            if (alvo == null) throw new FormatException("Planilha fora do padrão: aba sem arquivo.");
            string caminhoAba = alvo.StartsWith("/") ? alvo.TrimStart('/') : "xl/" + alvo;

            // textos compartilhados
            List<string> textos = new List<string>();
            using (XmlReader x = Abrir(zip, "xl/sharedStrings.xml", false))
            {
                if (x != null)
                {
                    StringBuilder sb = new StringBuilder(); bool dentro = false;
                    x.Read();
                    while (!x.EOF)
                    {
                        if (x.NodeType == XmlNodeType.Element)
                        {
                            if (x.LocalName == "si") { sb.Length = 0; dentro = true; if (x.IsEmptyElement) { textos.Add(""); dentro = false; } }
                            else if (x.LocalName == "rPh") { x.Skip(); continue; }
                            else if (x.LocalName == "t" && dentro) { sb.Append(x.ReadElementContentAsString()); continue; }
                        }
                        else if (x.NodeType == XmlNodeType.EndElement && x.LocalName == "si") { textos.Add(Desescapar(sb.ToString())); dentro = false; }
                        x.Read();
                    }
                }
            }

            // células
            List<object[]> linhas = new List<object[]>();
            using (XmlReader x = Abrir(zip, caminhoAba, true))
            {
                List<object> atual = null; int proxCol = 0; bool temValor = false;
                x.Read();
                while (!x.EOF)
                {
                    if (x.NodeType == XmlNodeType.Element && x.LocalName == "row")
                    {
                        atual = new List<object>(); proxCol = 0; temValor = false;
                        if (x.IsEmptyElement) { atual = null; }
                    }
                    else if (x.NodeType == XmlNodeType.Element && x.LocalName == "c" && atual != null)
                    {
                        string refCel = x.GetAttribute("r"); string tipo = x.GetAttribute("t");
                        int col = refCel != null ? ColunaDe(refCel) : proxCol;
                        proxCol = col + 1;
                        string bruto = null, embutido = null;
                        if (x.IsEmptyElement) { x.Read(); }
                        else
                        {
                            int prof = x.Depth; x.Read();
                            while (!x.EOF && !(x.NodeType == XmlNodeType.EndElement && x.Depth == prof))
                            {
                                if (x.NodeType == XmlNodeType.Element && x.LocalName == "v") { bruto = x.ReadElementContentAsString(); continue; }
                                if (x.NodeType == XmlNodeType.Element && x.LocalName == "is") { embutido = TextoEmbutido(x); continue; }
                                if (x.NodeType == XmlNodeType.Element && x.LocalName == "f") { x.Skip(); continue; }
                                x.Read();
                            }
                            x.Read();
                        }
                        object v = Converter(tipo, bruto, embutido, textos);
                        if (v != null && col >= 0)
                        {
                            while (atual.Count <= col) atual.Add(null);
                            atual[col] = v; temValor = true;
                        }
                        continue;
                    }
                    else if (x.NodeType == XmlNodeType.EndElement && x.LocalName == "row")
                    {
                        if (temValor && atual != null) linhas.Add(atual.ToArray());
                        atual = null;
                    }
                    x.Read();
                }
            }
            return linhas;
        }
    }

    static object Converter(string tipo, string bruto, string embutido, List<string> textos)
    {
        if (tipo == "inlineStr") return Limpo(embutido);
        if (bruto == null) return Limpo(embutido);
        if (tipo == "s") { int i; return (int.TryParse(bruto, NumberStyles.Integer, Inv, out i) && i >= 0 && i < textos.Count) ? Limpo(textos[i]) : null; }
        if (tipo == "str") return Limpo(Desescapar(bruto));
        if (tipo == "b") return bruto == "1" ? "Sim" : "Não";
        if (tipo == "e") return null;
        if (tipo == "d") { DateTime dt; if (DateTime.TryParse(bruto, Inv, DateTimeStyles.RoundtripKind, out dt)) return dt.ToOADate(); return Limpo(bruto); }
        double d; if (double.TryParse(bruto, NumberStyles.Float, Inv, out d)) return d;
        return Limpo(bruto);
    }

    // texto sem espaços nas pontas (inclui o espaço "duro" do SSW, \u00A0); vazio -> null
    static string Limpo(string s)
    {
        if (s == null) return null;
        s = s.Replace('\u00A0', ' ').Trim();
        return s.Length == 0 ? null : s;
    }

    // ------------------------------------------------------------ CSV
    public static List<object[]> LerCsv(string texto, char sep)
    {
        List<object[]> linhas = new List<object[]>();
        List<object> atual = new List<object>();
        StringBuilder campo = new StringBuilder();
        bool aspas = false, temValor = false; int i = 0, n = texto.Length;
        while (i < n)
        {
            char c = texto[i];
            if (aspas)
            {
                if (c == '"') { if (i + 1 < n && texto[i + 1] == '"') { campo.Append('"'); i += 2; continue; } aspas = false; i++; continue; }
                campo.Append(c); i++; continue;
            }
            if (c == '"' && campo.Length == 0) { aspas = true; i++; continue; }
            if (c == sep) { string v = Limpo(campo.ToString()); atual.Add(v); if (v != null) temValor = true; campo.Length = 0; i++; continue; }
            if (c == '\r' || c == '\n')
            {
                string v = Limpo(campo.ToString()); atual.Add(v); if (v != null) temValor = true; campo.Length = 0;
                if (temValor) linhas.Add(atual.ToArray());
                atual = new List<object>(); temValor = false;
                if (c == '\r' && i + 1 < n && texto[i + 1] == '\n') i++;
                i++; continue;
            }
            campo.Append(c); i++;
        }
        { string v = Limpo(campo.ToString()); atual.Add(v); if (v != null) temValor = true; if (temValor) linhas.Add(atual.ToArray()); }
        return linhas;
    }

    // ------------------------------------------------------------ JSON
    // tipos: uma letra por coluna escolhida: d = data | t = texto | a = automático (número ou texto)
    // deCsv = true: textos com cara de número brasileiro ("1.234,56") viram número nas colunas "a"
    public static string LinhasJson(List<object[]> linhas, int linhaCab, int[] cols, string tipos, int colData, bool deCsv)
    {
        StringBuilder sb = new StringBuilder(); int qtd = 0; UltimaSemData = 0;
        for (int r = linhaCab + 1; r < linhas.Count; r++)
        {
            object[] l = linhas[r]; int cheias = 0;
            for (int k = 0; k < cols.Length; k++) if (cols[k] < l.Length && l[cols[k]] != null) cheias++;
            if (cheias < 2) continue;                                  // linha vazia ou "lixo" de rodapé
            if (colData >= 0)
            {
                object dv = colData < l.Length ? l[colData] : null;
                if (DataIso(dv) == null) { UltimaSemData++; continue; }
            }
            if (qtd > 0) sb.Append(",\r\n");
            sb.Append('[');
            for (int k = 0; k < cols.Length; k++)
            {
                if (k > 0) sb.Append(',');
                object v = cols[k] < l.Length ? l[cols[k]] : null;
                sb.Append(Valor(v, tipos[k], deCsv));
            }
            sb.Append(']'); qtd++;
        }
        UltimaQtd = qtd;
        return sb.ToString();
    }

    static string Valor(object v, char tipo, bool deCsv)
    {
        if (v == null) return "null";
        if (tipo == 'd') { string iso = DataIso(v); return iso == null ? Str(Convert.ToString(v, Inv)) : "\"" + iso + "\""; }
        if (v is double)
        {
            double d = (double)v;
            if (double.IsNaN(d) || double.IsInfinity(d)) return "null";
            if (tipo == 't') return Str(d == Math.Floor(d) && Math.Abs(d) < 1e15 ? d.ToString("0", Inv) : d.ToString("0.##########", Inv));
            return Math.Round(d, 6).ToString("0.######", Inv);
        }
        string s = (string)v;
        if (tipo == 'a' && deCsv) { double n; if (NumeroBR(s, out n)) return Math.Round(n, 6).ToString("0.######", Inv); }
        return Str(s);
    }

    static readonly string[] FormatosData = new string[] {
        "dd/MM/yyyy", "d/M/yyyy", "dd/MM/yy", "d/M/yy", "dd/MM/yyyy HH:mm", "dd/MM/yyyy HH:mm:ss", "d/M/yyyy HH:mm", "d/M/yyyy HH:mm:ss",
        "dd/MM/yyyy H:mm", "yyyy-MM-dd", "yyyy-MM-dd HH:mm:ss", "yyyy-MM-ddTHH:mm:ss", "yyyy-MM-dd HH:mm", "dd-MM-yyyy", "dd.MM.yyyy" };

    // data -> "aaaa-mm-dd" (ou "aaaa-mm-ddThh:mm" quando tem hora); inválida -> null
    public static string DataIso(object v)
    {
        if (v == null) return null;
        DateTime dt;
        if (v is double)
        {
            double d = (double)v;
            if (d < 20000 || d > 80000) return null;
            d = Math.Round(d * 1440.0) / 1440.0;                     // arredonda para o minuto
            dt = DateTime.FromOADate(d);
        }
        else
        {
            string s = ((string)v).Trim();
            if (!DateTime.TryParseExact(s, FormatosData, Inv, DateTimeStyles.None, out dt))
            {
                double n;
                if (Regex.IsMatch(s, "^\\d{5}([.,]\\d+)?$") && double.TryParse(s.Replace(',', '.'), NumberStyles.Float, Inv, out n)) return DataIso(n);
                return null;
            }
        }
        if (dt.Year < 2000 || dt.Year > 2100) return null;
        return dt.TimeOfDay.TotalMinutes >= 1 ? dt.ToString("yyyy-MM-dd'T'HH:mm", Inv) : dt.ToString("yyyy-MM-dd", Inv);
    }

    // "1.234,56" | "R$ 2.150,01" | "-12,5" | "1234" -> número. Códigos com zero à esquerda ("0123") não.
    public static bool NumeroBR(string s, out double n)
    {
        n = 0;
        if (s == null) return false;
        string t = s.Replace("R$", "").Replace(" ", "").Replace("\u00A0", "").Trim();
        if (t.Length == 0 || t.Length > 20) return false;
        if (t.Length > 1 && t[0] == '0' && char.IsDigit(t[1])) return false;
        if (Regex.IsMatch(t, "^-?\\d{1,3}(\\.\\d{3})+(,\\d+)?$") || Regex.IsMatch(t, "^-?\\d+(,\\d+)?$"))
            return double.TryParse(t.Replace(".", "").Replace(',', '.'), NumberStyles.Float, Inv, out n);
        if (Regex.IsMatch(t, "^-?\\d+\\.\\d+$")) return double.TryParse(t, NumberStyles.Float, Inv, out n);
        return false;
    }

    public static string Str(string s)
    {
        if (s == null) return "null";
        StringBuilder sb = new StringBuilder(s.Length + 2); sb.Append('"');
        foreach (char c in s)
        {
            switch (c)
            {
                case '"': sb.Append("\\\""); break;
                case '\\': sb.Append("\\\\"); break;
                case '\n': sb.Append("\\n"); break;
                case '\r': break;
                case '\t': sb.Append(' '); break;
                case '<': sb.Append("\\u003c"); break;                // evita "</script>" dentro do HTML de envio
                case '\u2028': sb.Append("\\u2028"); break;
                case '\u2029': sb.Append("\\u2029"); break;
                default: if (c < ' ') sb.Append(' '); else sb.Append(c); break;
            }
        }
        sb.Append('"'); return sb.ToString();
    }

    // ------------------------------------------------------------ auxiliares
    static string Simples(string s)
    {
        StringBuilder sb = new StringBuilder();
        foreach (char ch in (s ?? "").Normalize(NormalizationForm.FormD))
        {
            if (CharUnicodeInfo.GetUnicodeCategory(ch) == UnicodeCategory.NonSpacingMark) continue;
            if (char.IsLetterOrDigit(ch)) sb.Append(char.ToLowerInvariant(ch));
        }
        return sb.ToString();
    }
    static XmlReader Abrir(ZipArchive zip, string nome, bool obrigatorio)
    {
        ZipArchiveEntry e = zip.GetEntry(nome);
        if (e == null) foreach (ZipArchiveEntry o in zip.Entries) if (string.Equals(o.FullName.Replace('\\', '/'), nome, StringComparison.OrdinalIgnoreCase)) { e = o; break; }
        if (e == null) { if (obrigatorio) throw new FormatException("Planilha fora do padrão (falta " + nome + ")."); return null; }
        XmlReaderSettings cfg = new XmlReaderSettings(); cfg.IgnoreComments = true; cfg.IgnoreProcessingInstructions = true; cfg.CloseInput = true;
        return XmlReader.Create(e.Open(), cfg);
    }
    public static int ColunaDe(string referencia)
    {
        int col = 0;
        foreach (char ch in referencia)
        {
            if (ch >= 'A' && ch <= 'Z') col = col * 26 + (ch - 'A' + 1);
            else if (ch >= 'a' && ch <= 'z') col = col * 26 + (ch - 'a' + 1);
            else break;
        }
        return col - 1;
    }
    public static string LetraDe(int i)
    {
        string s = ""; i++;
        while (i > 0) { int r = (i - 1) % 26; s = (char)('A' + r) + s; i = (i - 1) / 26; }
        return s;
    }
    static string TextoEmbutido(XmlReader x)
    {
        StringBuilder sb = new StringBuilder();
        if (x.IsEmptyElement) { x.Read(); return ""; }
        int prof = x.Depth; x.Read();
        while (!x.EOF && !(x.NodeType == XmlNodeType.EndElement && x.Depth == prof))
        {
            if (x.NodeType == XmlNodeType.Element && x.LocalName == "rPh") { x.Skip(); continue; }
            if (x.NodeType == XmlNodeType.Element && x.LocalName == "t") { sb.Append(x.ReadElementContentAsString()); continue; }
            x.Read();
        }
        x.Read();
        return Desescapar(sb.ToString());
    }
    static string Desescapar(string s)
    {
        if (s.IndexOf("_x", StringComparison.Ordinal) < 0) return s;
        return Regex.Replace(s, "_x([0-9A-Fa-f]{4})_", delegate(Match m) {
            return ((char)int.Parse(m.Groups[1].Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture)).ToString();
        });
    }
}
'@

function Carregar-Leitor {
    if ('OstelVendasLeitor' -as [type]) { return }
    if ($PSVersionTable.PSEdition -eq 'Core') {
        Add-Type -TypeDefinition $CodigoLeitor -Language CSharp
    } else {
        Add-Type -AssemblyName System.IO.Compression
        Add-Type -TypeDefinition $CodigoLeitor -Language CSharp -ReferencedAssemblies @('System.IO.Compression', 'System.Xml')
    }
}

# ==================================================================
#  UTILITÁRIOS
# ==================================================================
# Normaliza textos de cabeçalho: minúsculas, sem acentos e sem espaços duplos
function Norm([object]$s) {
    if ($null -eq $s) { return '' }
    $t = ([string]$s).Replace([string][char]0xA0, ' ').Trim().ToLowerInvariant() -replace '\s+', ' '
    $d = $t.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object Text.StringBuilder
    foreach ($c in $d.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($c) -ne [Globalization.UnicodeCategory]::NonSpacingMark) { [void]$sb.Append($c) }
    }
    return $sb.ToString()
}
# Mensagem de erro sem o "Exception calling ... with N argument(s)" que o PowerShell acrescenta
function Msg($erro) {
    $e = $erro; if ($e -is [Management.Automation.ErrorRecord]) { $e = $e.Exception }
    while ($e.InnerException) { $e = $e.InnerException }
    return $e.Message
}
function JsonStr([object]$v) { if ($null -eq $v) { return 'null' }; return [OstelVendasLeitor]::Str([string]$v) }

# Pasta temporária (formato longo: evita o erro do "C:\Users\ANDR~1" em usuários com acento)
function Get-PastaTemp {
    foreach ($p in @($(if ($env:LOCALAPPDATA) { [IO.Path]::Combine($env:LOCALAPPDATA, 'Temp') }), [IO.Path]::GetTempPath())) {
        if ($p -and [IO.Directory]::Exists($p)) { return $p }
    }
    return $PastaLogs
}
function Apagar-Arquivo([string]$c) { try { if ($c -and [IO.File]::Exists($c)) { [IO.File]::Delete($c) } } catch { } }

# Lê os bytes do arquivo mesmo que esteja aberto no Excel (até 3 tentativas)
function Ler-Bytes([string]$arq) {
    $ultimo = $null
    for ($i = 1; $i -le 3; $i++) {
        try {
            $compart = [IO.FileShare]([int][IO.FileShare]::ReadWrite -bor [int][IO.FileShare]::Delete)
            $fs = [IO.File]::Open($arq, [IO.FileMode]::Open, [IO.FileAccess]::Read, $compart)
            try { $ms = New-Object IO.MemoryStream; $fs.CopyTo($ms); return $ms.ToArray() } finally { $fs.Dispose() }
        } catch { $ultimo = $_; Start-Sleep -Seconds 2 }
    }
    throw $ultimo
}

# Planilha bloqueada pelo Excel: pede ao próprio Excel uma cópia (não salva nem fecha a planilha aberta)
function Copiar-DoExcelAberto([string]$origem, [string]$destino) {
    $xl = $null
    try { $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { return $false }
    try {
        $nome = [IO.Path]::GetFileName($origem)
        foreach ($wb in $xl.Workbooks) { if ($wb.Name -ieq $nome) { Apagar-Arquivo $destino; $wb.SaveCopyAs($destino); return [IO.File]::Exists($destino) } }
    } catch { } finally { if ($xl) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch { } } }
    return $false
}

# .xlsb / .xls: converte com o Excel instalado para um .xlsx temporário
function Converter-ViaExcel([string]$origem, [string]$destinoXlsx) {
    $tmpIn = [IO.Path]::Combine((Get-PastaTemp), 'ostel_in_' + [guid]::NewGuid().ToString('N') + [IO.Path]::GetExtension($origem))
    [IO.File]::WriteAllBytes($tmpIn, (Ler-Bytes $origem))
    $xl = $null
    try {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false; $xl.DisplayAlerts = $false
        $wb = $xl.Workbooks.Open($tmpIn, 0, $true)
        $wb.SaveAs($destinoXlsx, 51); $wb.Close($false)
    } catch {
        throw [FormatException]('Arquivos .xlsb/.xls precisam do Excel instalado. Salve a planilha como .xlsx ou CSV. Detalhe: ' + $_.Exception.Message)
    } finally {
        if ($xl) { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
        Apagar-Arquivo $tmpIn
    }
}

# Texto de um CSV: detecta a codificação (BOM, UTF-8 válido ou Windows-1252/ANSI)
function Texto-Csv([byte[]]$bytes) {
    if ($bytes.Length -eq 0) { throw [FormatException]'O arquivo CSV está vazio.' }
    $enc = $null; $ini = 0
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { $enc = $Utf8; $ini = 3 }
    elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) { $enc = [Text.Encoding]::Unicode; $ini = 2 }
    elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) { $enc = [Text.Encoding]::BigEndianUnicode; $ini = 2 }
    else {
        try { [void](New-Object Text.UTF8Encoding($false, $true)).GetString($bytes); $enc = $Utf8 }
        catch { $enc = [Text.Encoding]::GetEncoding(1252) }
    }
    return @{ Texto = $enc.GetString($bytes, $ini, $bytes.Length - $ini); Codificacao = $enc.WebName }
}
# Separador: o caractere que mais aparece nas primeiras linhas com conteúdo (; tab , |)
function Separador-Csv([string]$texto) {
    $d = [string]$Config.Delimitador
    if ($d) { return [char]$d[0] }
    $amostra = @($texto -split "`r?`n" | Where-Object { $_.Trim() -ne '' } | Select-Object -First 5)
    $melhor = ';'; $max = -1
    foreach ($c in @(';', "`t", ',', '|')) {
        $qtd = 0; foreach ($l in $amostra) { $qtd += ($l.Split([char]$c)).Count - 1 }
        if ($qtd -gt $max) { $max = $qtd; $melhor = $c }
    }
    return [char]$melhor
}

# ==================================================================
#  LOCALIZAÇÃO DAS PLANILHAS
# ==================================================================
function Get-Origem([string]$cfg, [string]$procurar, [string]$nome) {
    $cfg = ([string]$cfg).Trim().Trim('"')
    if ($cfg) {
        $p = $cfg; if (-not [IO.Path]::IsPathRooted($p)) { $p = Join-Path $Pasta $p }
        if (Test-Path -LiteralPath $p -PathType Leaf) { return (Get-Item -LiteralPath $p) }
        # mesmo nome na pasta ACIMA do painel (ex.: OneDrive em outro usuário do Windows)
        $acima = Join-Path (Split-Path -Parent $Pasta) ([IO.Path]::GetFileName($p))
        if (Test-Path -LiteralPath $acima -PathType Leaf) { Log "Planilha '$nome' não está em $p; usando $acima" 'AVISO'; return (Get-Item -LiteralPath $acima) }
    }
    $arq = Get-ChildItem -LiteralPath $PastaDados -File |
           Where-Object { $_.Extension -match '^\.(csv|txt|xlsx|xlsm|xlsb|xls)$' -and $_.Name -notlike '~$*' -and (Norm $_.Name).Contains((Norm $procurar)) } |
           Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($arq) { return $arq }
    $onde = if ($cfg) { $cfg } else { 'dados\' }
    throw [IO.FileNotFoundException]("Planilha '$nome' não encontrada ($onde).")
}

# ==================================================================
#  LEITURA DE UMA FONTE -> objeto JSON (texto)
# ==================================================================
function Ler-Fonte([string]$chave, [IO.FileInfo]$origem, [string]$aba) {
    $F = $Fontes[$chave]
    $ext = $origem.Extension.ToLower(); $deCsv = $false; $info = ''
    if ($ext -eq '.xlsx' -or $ext -eq '.xlsm') {
        try { $linhas = [OstelVendasLeitor]::LerXlsx($origem.FullName, $aba) }
        catch [IO.InvalidDataException] { throw [FormatException]'A planilha não pôde ser aberta (arquivo corrompido ou não é um .xlsx válido).' }
        catch [FormatException] { throw [FormatException](Msg $_) }
        catch [IO.IOException] {
            # arquivo travado (Excel/OneDrive sincronizando): tenta pela cópia do Excel aberto
            $erroOriginal = $_.Exception
            $tmp = [IO.Path]::Combine((Get-PastaTemp), 'ostel_' + [guid]::NewGuid().ToString('N') + $ext)
            try {
                if (-not (Copiar-DoExcelAberto $origem.FullName $tmp)) { throw $erroOriginal }
                $linhas = [OstelVendasLeitor]::LerXlsx($tmp, $aba)
            } finally { Apagar-Arquivo $tmp }
        }
        $info = "planilha, aba '" + [OstelVendasLeitor]::UltimaAba + "'"
    } elseif ($ext -eq '.xlsb' -or $ext -eq '.xls') {
        $tmp = [IO.Path]::Combine((Get-PastaTemp), 'ostel_' + [guid]::NewGuid().ToString('N') + '.xlsx')
        try { Converter-ViaExcel $origem.FullName $tmp; $linhas = [OstelVendasLeitor]::LerXlsx($tmp, $aba) } finally { Apagar-Arquivo $tmp }
        $info = "planilha $ext convertida pelo Excel, aba '" + [OstelVendasLeitor]::UltimaAba + "'"
    } else {
        $t = Texto-Csv (Ler-Bytes $origem.FullName)
        $sep = Separador-Csv $t.Texto
        $linhas = [OstelVendasLeitor]::LerCsv($t.Texto, $sep)
        $deCsv = $true
        $info = "CSV, separador '" + $(if ($sep -eq "`t") { 'TAB' } else { [string]$sep }) + "', codificação " + $t.Codificacao
    }
    if ($linhas.Count -eq 0) { throw [FormatException]("A planilha '" + $origem.Name + "' está vazia.") }

    # ---- linha de títulos: a primeira (até a 30ª) que tem todas as colunas obrigatórias
    $obrig = @($F.Obrigatorias | ForEach-Object { Norm $_ })
    $cab = -1
    for ($r = 0; $r -lt [Math]::Min(30, $linhas.Count); $r++) {
        $nomes = @($linhas[$r] | ForEach-Object { Norm $_ })
        $ok = $true; foreach ($o in $obrig) { if ($nomes -notcontains $o) { $ok = $false; break } }
        if ($ok) { $cab = $r; break }
    }
    if ($cab -lt 0) {
        throw [FormatException]("Não encontrei a linha de títulos em '" + $origem.Name + "'. Ela precisa ter as colunas: " + ($F.Obrigatorias -join ', ') + '.')
    }

    # ---- colunas: mantém todas as que têm nome (mesma ordem da planilha); define o tipo de cada uma
    $hdr = $linhas[$cab]
    $cols = New-Object Collections.Generic.List[int]; $nomesCols = New-Object Collections.Generic.List[string]
    $letras = New-Object Collections.Generic.List[string]; $tipos = New-Object Text.StringBuilder
    $palData = @($Config.PalavrasData | ForEach-Object { Norm $_ }); $palTexto = @($Config.PalavrasTexto | ForEach-Object { Norm $_ })
    $palNum = @($Config.PalavrasNumero | ForEach-Object { Norm $_ })
    $colData = -1; $vistos = @{}
    for ($c = 0; $c -lt $hdr.Length; $c++) {
        $h = $hdr[$c]
        if ($null -eq $h) { continue }
        $hs = ([string]$h).Trim()
        if ($hs -match '^[\d.,\s]*$') { continue }                     # sem nome ou só números (ex.: coluna A "1")
        $n = Norm $hs
        if ($vistos.ContainsKey($n)) { $hs = $hs + ' (' + [OstelVendasLeitor]::LetraDe($c) + ')' }
        $vistos[$n] = $true
        $tipo = 'a'
        foreach ($p in $palTexto) { if ($n.Contains($p)) { $tipo = 't'; break } }
        foreach ($p in $palNum) { if ($n.Contains($p)) { $tipo = 'a'; break } }        # valor vence código ("FRETE CTRC")
        foreach ($p in $palData) { if ($n.StartsWith($p) -or $n.Contains(' ' + $p)) { $tipo = 'd'; break } }
        if ($n -eq (Norm $F.ColunaData)) { $colData = $c; $tipo = 'd' }
        $cols.Add($c); $nomesCols.Add($hs); $letras.Add([OstelVendasLeitor]::LetraDe($c)); [void]$tipos.Append($tipo)
    }
    $json = [OstelVendasLeitor]::LinhasJson($linhas, $cab, $cols.ToArray(), $tipos.ToString(), $colData, $deCsv)
    $qtd = [OstelVendasLeitor]::UltimaQtd; $semData = [OstelVendasLeitor]::UltimaSemData
    if ($semData -gt 0) { $script:Avisos.Add($F.Nome + ": $semData linha(s) ignorada(s) por não terem '" + $F.ColunaData + "' válida.") }
    if ($qtd -eq 0) { throw [FormatException]("A planilha '" + $origem.Name + "' foi lida, mas nenhuma linha tem '" + $F.ColunaData + "' válida.") }
    Log ("{0}: {1} linhas, {2} colunas ({3})" -f $F.Nome, $qtd, $cols.Count, $info) 'OK'

    return "{`r`n    ok: true,`r`n    arquivo: " + (JsonStr $origem.Name) +
           ",`r`n    arquivoModificado: " + (JsonStr $origem.LastWriteTime.ToString('yyyy-MM-ddTHH:mm:ss')) +
           ",`r`n    leitura: " + (JsonStr $info) + ",`r`n    qtd: " + $qtd +
           ",`r`n    colunas: [" + (($nomesCols | ForEach-Object { JsonStr $_ }) -join ',') + "]" +
           ",`r`n    letras: [" + (($letras | ForEach-Object { '"' + $_ + '"' }) -join ',') + "]" +
           ",`r`n    tipos: " + (JsonStr $tipos.ToString()) +
           ",`r`n    linhas: [`r`n" + $json + "`r`n    ]`r`n  }"
}

function Fonte-Erro([string]$arquivo, [string]$erro, [string]$dica) {
    return "{ ok: false, arquivo: " + (JsonStr $arquivo) + ", erro: " + (JsonStr $erro) + ", dica: " + (JsonStr $dica) + " }"
}

# ==================================================================
#  METAS DO MÊS (dados\metas.csv:  Mês ; Vendedora ; Meta)
# ==================================================================
# ==================================================================
#  CORREÇÕES DE VENDEDORA (dados\correcoes_vendedora.csv:  CT-e ; Vendedora correta)
#  Para os CT-e listados, o painel usa a vendedora escrita aqui no lugar
#  da que veio da planilha (coluna V ou BC). A planilha não é alterada.
# ==================================================================
function Caminho-Correcoes {
    $p = ([string]$Config.ArquivoCorrecoes).Trim().Trim('"')
    if (-not $p) { $p = 'dados\correcoes_vendedora.csv' }
    if (-not [IO.Path]::IsPathRooted($p)) { $p = Join-Path $Pasta $p }
    return $p
}
function Ler-Correcoes {
    $arq = Caminho-Correcoes
    if (-not [IO.File]::Exists($arq)) {
        try { [IO.File]::WriteAllText($arq, "CT-e;Vendedora correta`r`n", $Utf8Bom); Log "Modelo de correções de vendedora criado em $arq." 'INFO' } catch { }
        return '[]'
    }
    $t = Texto-Csv (Ler-Bytes $arq)
    $linhas = [OstelVendasLeitor]::LerCsv($t.Texto, (Separador-Csv $t.Texto))
    $saida = New-Object Collections.Generic.List[string]
    foreach ($l in $linhas) {
        if ($l.Length -lt 2) { continue }
        $num = [string]$l[0]; $vend = [string]$l[1]
        if (-not $num -or -not $vend) { continue }
        if ((Norm $num) -eq 'ct-e' -or (Norm $num) -eq 'cte') { continue }      # títulos
        $saida.Add('[' + (JsonStr $num) + ',' + (JsonStr $vend) + ']')
    }
    Log ("Correções de vendedora: {0} CT-e em {1}" -f $saida.Count, (Split-Path -Leaf $arq)) 'OK'
    return '[' + ($saida -join ',') + ']'
}

# ==================================================================
#  PRÊMIOS DAS VENDEDORAS (dados\premios.csv:
#     Faixa ; Faturamento mínimo ; Prêmio ; Exige bater a meta (S/N))
#  Cada faixa vale quando o faturamento do mês da vendedora chega ao mínimo.
#  "Exige bater a meta" = S: a faixa só vale se a vendedora bater a própria meta.
# ==================================================================
function Caminho-Premios {
    $p = ([string]$Config.ArquivoPremios).Trim().Trim('"')
    if (-not $p) { $p = 'dados\premios.csv' }
    if (-not [IO.Path]::IsPathRooted($p)) { $p = Join-Path $Pasta $p }
    return $p
}
function Ler-Premios {
    $arq = Caminho-Premios
    if (-not [IO.File]::Exists($arq)) { return '[]' }
    $t = Texto-Csv (Ler-Bytes $arq)
    $linhas = [OstelVendasLeitor]::LerCsv($t.Texto, (Separador-Csv $t.Texto))
    $saida = New-Object Collections.Generic.List[string]; $ruins = 0
    foreach ($l in $linhas) {
        if ($l.Length -lt 3 -or $null -eq $l[1] -or $null -eq $l[2]) { continue }
        $fx = ([string]$l[0]).Trim() -replace '[^0-9]', ''
        if (-not $fx) { continue }                                              # títulos
        $fat = 0.0; $pre = 0.0
        $okF = if ($l[1] -is [double]) { $fat = [double]$l[1]; $true } else { [OstelVendasLeitor]::NumeroBR(((([string]$l[1]) -replace 'R\$', '').Trim()), [ref]$fat) }
        $okP = if ($l[2] -is [double]) { $pre = [double]$l[2]; $true } else { [OstelVendasLeitor]::NumeroBR(((([string]$l[2]) -replace 'R\$', '').Trim()), [ref]$pre) }
        if (-not $okF -or -not $okP) { $ruins++; continue }
        $exige = $false; if ($l.Length -ge 4 -and $null -ne $l[3]) { $exige = (Norm $l[3]) -like 's*' }
        $saida.Add('[' + [int]$fx + ',' + $fat.ToString('0.##', $Inv) + ',' + $pre.ToString('0.##', $Inv) + ',' + $(if ($exige) { 'true' } else { 'false' }) + ']')
    }
    if ($ruins) { $script:Avisos.Add("premios.csv: $ruins linha(s) com valor inválido foram ignoradas.") }
    Log ("Prêmios: {0} faixa(s) em {1}" -f $saida.Count, (Split-Path -Leaf $arq)) 'OK'
    return '[' + ($saida -join ',') + ']'
}

# ==================================================================
#  COMISSÃO DAS VENDEDORAS (dados\comissao.csv:  Vendedora ; Percentual)
#  "Todas" vale para quem não tiver linha própria. Ex.: Todas;1  (= 1%)
# ==================================================================
function Caminho-Comissao {
    $p = ([string]$Config.ArquivoComissao).Trim().Trim('"')
    if (-not $p) { $p = 'dados\comissao.csv' }
    if (-not [IO.Path]::IsPathRooted($p)) { $p = Join-Path $Pasta $p }
    return $p
}
function Ler-Comissao {
    $arq = Caminho-Comissao
    if (-not [IO.File]::Exists($arq)) { return '[]' }
    $t = Texto-Csv (Ler-Bytes $arq)
    $linhas = [OstelVendasLeitor]::LerCsv($t.Texto, (Separador-Csv $t.Texto))
    $saida = New-Object Collections.Generic.List[string]
    foreach ($l in $linhas) {
        if ($l.Length -lt 2 -or $null -eq $l[0] -or $null -eq $l[1]) { continue }
        $nome = ([string]$l[0]).Trim(); if (-not $nome -or (Norm $nome) -eq 'vendedora') { continue }
        $pc = 0.0
        $ok = if ($l[1] -is [double]) { $pc = [double]$l[1]; $true } else { [OstelVendasLeitor]::NumeroBR(((([string]$l[1]) -replace '%', '').Trim()), [ref]$pc) }
        if (-not $ok) { $script:Avisos.Add("comissao.csv: percentual inválido para $nome."); continue }
        $saida.Add('[' + (JsonStr $nome) + ',' + $pc.ToString('0.####', $Inv) + ']')
    }
    Log ("Comissão: {0} linha(s) em {1}" -f $saida.Count, (Split-Path -Leaf $arq)) 'OK'
    return '[' + ($saida -join ',') + ']'
}

# ==================================================================
#  FERIADOS (dados\Feriados.csv: Data ; Dia da semana ; Feriado ; Abrangência ; Tipo ; Folga (S/N))
#  Dias com Folga = S deixam de ser dia útil na meta (ritmo, quanto falta
#  por dia útil, projeção). Folga em branco = S. Folga = N continua dia útil.
# ==================================================================
function Caminho-Feriados {
    $p = ([string]$Config.ArquivoFeriados).Trim().Trim('"')
    if (-not $p) { $p = 'dados\Feriados.csv' }
    if (-not [IO.Path]::IsPathRooted($p)) { $p = Join-Path $Pasta $p }
    return $p
}
function Ler-Feriados {
    $arq = Caminho-Feriados
    if (-not [IO.File]::Exists($arq)) {
        try { [IO.File]::WriteAllText($arq, "Data;Dia da semana;Feriado;Abrangência;Tipo;Folga (S/N)`r`n", $Utf8Bom); Log "Modelo de feriados criado em $arq." 'INFO' } catch { }
        return '[]'
    }
    $t = Texto-Csv (Ler-Bytes $arq)
    $linhas = [OstelVendasLeitor]::LerCsv($t.Texto, (Separador-Csv $t.Texto))
    $iNome = 2; $iFolga = -1
    $saida = New-Object Collections.Generic.List[string]; $ignorados = 0
    foreach ($l in $linhas) {
        if ($l.Length -lt 1 -or $null -eq $l[0]) { continue }
        $iso = [OstelVendasLeitor]::DataIso($l[0])
        if (-not $iso) {
            # linha de títulos: descobre as colunas pelo nome
            for ($j = 0; $j -lt $l.Length; $j++) { $h = Norm $l[$j]; if ($h -like 'folga*') { $iFolga = $j } elseif ($h -eq 'feriado' -or $h -like 'descri*') { $iNome = $j } }
            continue
        }
        $iso = $iso.Substring(0, 10)
        $folga = ''; if ($iFolga -ge 0 -and $iFolga -lt $l.Length -and $null -ne $l[$iFolga]) { $folga = (Norm $l[$iFolga]) }
        if ($folga -like 'n*') { $ignorados++; continue }
        $nome = ''; if ($iNome -lt $l.Length -and $null -ne $l[$iNome]) { $nome = ([string]$l[$iNome]).Trim() }
        $saida.Add('[' + (JsonStr $iso) + ',' + (JsonStr $nome) + ']')
    }
    Log ("Feriados: {0} dias de folga em {1} ({2} marcados para trabalhar)" -f $saida.Count, (Split-Path -Leaf $arq), $ignorados) 'OK'
    return '[' + ($saida -join ',') + ']'
}

# ==================================================================
#  TABELA DE OCORRÊNCIAS DO SSW (dados\ocorrencias_ssw.csv ou .xlsx: COD ; DESCRICAO)
#  Troca o número da coluna "Codigo da Ultima Ocorrencia" pelo texto.
# ==================================================================
function Caminho-Ocorrencias {
    $p = ([string]$Config.ArquivoOcorrencias).Trim().Trim('"')
    if (-not $p) { $p = 'dados\ocorrencias_ssw.csv' }
    if (-not [IO.Path]::IsPathRooted($p)) { $p = Join-Path $Pasta $p }
    return $p
}
function Ler-Ocorrencias {
    $arq = Caminho-Ocorrencias
    if (-not [IO.File]::Exists($arq)) { return '{}' }
    if ($arq -match '\.xls[xm]$') { $linhas = [OstelVendasLeitor]::LerXlsx($arq, '') }
    else { $t = Texto-Csv (Ler-Bytes $arq); $linhas = [OstelVendasLeitor]::LerCsv($t.Texto, (Separador-Csv $t.Texto)) }
    $saida = New-Object Collections.Generic.List[string]
    foreach ($l in $linhas) {
        if ($l.Length -lt 2 -or $null -eq $l[0] -or $null -eq $l[1]) { continue }
        $cod = $l[0]; if ($cod -is [double]) { $cod = ([double]$cod).ToString('0', $Inv) }
        $cod = ([string]$cod).Trim(); if ($cod -notmatch '^\d+$') { continue }      # títulos
        $saida.Add((JsonStr $cod) + ':' + (JsonStr ([string]$l[1])))
    }
    Log ("Ocorrências do SSW: {0} códigos em {1}" -f $saida.Count, (Split-Path -Leaf $arq)) 'OK'
    return '{' + ($saida -join ',') + '}'
}

function Caminho-Metas {
    $p = ([string]$Config.ArquivoMetas).Trim().Trim('"')
    if (-not $p) { $p = 'dados\metas.csv' }
    if (-not [IO.Path]::IsPathRooted($p)) { $p = Join-Path $Pasta $p }
    return $p
}
function Ler-Metas {
    $arq = Caminho-Metas
    if (-not [IO.File]::Exists($arq)) {
        # cria o modelo para a empresa preencher (abre no Excel)
        $mes = (Get-Date).ToString('MM/yyyy')
        $modelo = "Mês;Vendedora;Meta`r`n$mes;TOTAL;`r`n"
        try { [IO.File]::WriteAllText($arq, $modelo, $Utf8Bom); Log "Modelo de metas criado em $arq (preencha os valores)." 'AVISO' } catch { }
        return '[]'
    }
    $t = Texto-Csv (Ler-Bytes $arq)
    $sep = Separador-Csv $t.Texto
    $linhas = [OstelVendasLeitor]::LerCsv($t.Texto, $sep)
    $saida = New-Object Collections.Generic.List[string]; $ruins = 0
    $mesesAbrev = @{ 'jan'=1; 'fev'=2; 'mar'=3; 'abr'=4; 'mai'=5; 'jun'=6; 'jul'=7; 'ago'=8; 'set'=9; 'out'=10; 'nov'=11; 'dez'=12 }
    for ($i = 0; $i -lt $linhas.Count; $i++) {
        $l = $linhas[$i]
        if ($l.Length -lt 3) { continue }
        $m = [string]$l[0]; $v = [string]$l[1]; $valor = [string]$l[2]
        if ((Norm $m) -eq 'mes' -or (Norm $m) -eq 'mês') { continue }      # títulos
        if (-not $v) { continue }                                        # linha sem vendedora
        # meta em branco: a vendedora entra na LISTA dos cards do mês, sem meta
        if (-not $valor) { $valor = '' }
        $ano = 0; $mm = 0; $mn = Norm $m
        if ($mn -match '^(\d{1,2})[/\-.](\d{4})$') { $mm = [int]$Matches[1]; $ano = [int]$Matches[2] }
        elseif ($mn -match '^(\d{4})[/\-.](\d{1,2})') { $ano = [int]$Matches[1]; $mm = [int]$Matches[2] }
        elseif ($mn -match '^\d{1,2}/(\d{1,2})/(\d{4})') { $mm = [int]$Matches[1]; $ano = [int]$Matches[2] }
        elseif ($mn -match '^([a-z]{3})[a-z]*[/\-. ]+(\d{2,4})$' -and $mesesAbrev.ContainsKey($Matches[1])) { $mm = $mesesAbrev[$Matches[1]]; $ano = [int]$Matches[2]; if ($ano -lt 100) { $ano += 2000 } }
        $num = 0.0
        if ($mm -lt 1 -or $mm -gt 12 -or $ano -lt 2000 -or ($valor -and -not [OstelVendasLeitor]::NumeroBR($valor, [ref]$num))) { $ruins++; continue }
        $metaJs = if ($valor) { $num.ToString('0.##', $Inv) } else { 'null' }
        $saida.Add('[' + (JsonStr ('{0:0000}-{1:00}' -f $ano, $mm)) + ',' + (JsonStr $v) + ',' + $metaJs + ']')
    }
    if ($ruins -gt 0) { $script:Avisos.Add("metas.csv: $ruins linha(s) com mês ou valor inválido foram ignoradas (use mês 10/2026 e valor 150000,00).") }
    Log ("Metas: {0} linha(s) lida(s) de {1}" -f $saida.Count, (Split-Path -Leaf $arq)) 'OK'
    return '[' + ($saida -join ',') + ']'
}

# ==================================================================
#  GRAVAÇÃO (troca atômica: o navegador nunca lê um arquivo pela metade)
# ==================================================================
function Gravar-Atomico([string]$destino, [string]$conteudo) {
    $tmp = $destino + '.tmp'
    [IO.File]::WriteAllText($tmp, $conteudo, $Utf8)
    if ([IO.File]::Exists($destino)) {
        try { [IO.File]::Replace($tmp, $destino, $null) }
        catch { [IO.File]::Copy($tmp, $destino, $true); Apagar-Arquivo $tmp }
    } else { [IO.File]::Move($tmp, $destino) }
}
function Gravar-Status([bool]$ok, [string]$msg) {
    $js = "/* Gerado por atualizar.ps1 - NAO EDITE. Mostra quando o atualizador rodou pela ultima vez. */`r`n" +
          "window.OSTEL_STATUS = { verificadoEm: " + (JsonStr (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss')) + ", ok: " + $(if ($ok) { 'true' } else { 'false' }) +
          ", mensagem: " + (JsonStr $msg) + ", intervaloMinutos: " + [int]$Config.IntervaloMinutos + " };`r`n"
    try { Gravar-Atomico $SaidaStatus $js } catch { }
}

# ------------------------------------------------------------------ garante o Chart.js local (modo offline)
function Garantir-ChartJs {
    $local = Join-Path $PastaAssets 'chart.umd.min.js'
    if ((Test-Path -LiteralPath $local) -and (Get-Item -LiteralPath $local).Length -gt 50000) { return }
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $Config.UrlChartJs -OutFile $local -UseBasicParsing -TimeoutSec 30
        Log 'Cópia local do Chart.js baixada em assets\chart.umd.min.js' 'OK'
    } catch { Log ('Não foi possível baixar o Chart.js (' + $_.Exception.Message + '). O painel usará a internet. Veja o LEIAME.txt.') 'AVISO' }
}

# ==================================================================
#  AGENDAMENTO (Agendador de Tarefas do Windows)
# ==================================================================
function Registrar-Agendamento {
    $min = [Math]::Max(1, [int]$Config.IntervaloMinutos)
    $nome = [string]$Config.NomeTarefa
    $ps = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $argumentos = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`" -Silencioso"
    $usuario = "$env:USERDOMAIN\$env:USERNAME"
    try {
        $acao   = New-ScheduledTaskAction -Execute $ps -Argument $argumentos -WorkingDirectory $Pasta
        $diario = New-ScheduledTaskTrigger -Daily -At '00:00'
        $diario.Repetition = (New-ScheduledTaskTrigger -Once -At '00:00' -RepetitionInterval (New-TimeSpan -Minutes $min) -RepetitionDuration (New-TimeSpan -Hours 23 -Minutes 59)).Repetition
        $logon  = New-ScheduledTaskTrigger -AtLogOn -User $usuario
        $conf   = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 10)
        $princ  = New-ScheduledTaskPrincipal -UserId $usuario -LogonType Interactive -RunLevel Limited
        $desc   = 'Atualiza o painel Dashboard Vendas da Ostel (' + $Pasta + ')'
        try { Register-ScheduledTask -TaskName $nome -Action $acao -Trigger @($diario, $logon) -Settings $conf -Principal $princ -Description $desc -Force | Out-Null }
        catch { Register-ScheduledTask -TaskName $nome -Action $acao -Trigger $diario -Settings $conf -Principal $princ -Description $desc -Force | Out-Null }
        Start-ScheduledTask -TaskName $nome
    } catch {
        # alternativa: schtasks.exe (Windows sem o módulo ScheduledTasks)
        Log ('Register-ScheduledTask indisponível, usando schtasks.exe (' + $_.Exception.Message + ')') 'AVISO'
        $tr = "`"$ps`" $argumentos"
        & schtasks.exe /Create /F /SC MINUTE /MO $min /TN $nome /TR $tr | Out-Null
        if ($LASTEXITCODE -ne 0) { throw 'Falha ao criar a tarefa agendada com schtasks.exe.' }
        & schtasks.exe /Run /TN $nome | Out-Null
    }
    Log ("Tarefa agendada '{0}' criada: atualiza a cada {1} minuto(s) e ao entrar no Windows." -f $nome, $min) 'OK'
    # nome antigo da tarefa (antes de 14/10/2026): remove para não rodar duas vezes
    $antigo = 'Ostel - Torre de Controle SAC - Vendas'
    if ($antigo -ne $nome) {
        try { if (Get-ScheduledTask -TaskName $antigo -ErrorAction SilentlyContinue) { Unregister-ScheduledTask -TaskName $antigo -Confirm:$false; Log "Tarefa antiga '$antigo' removida." 'OK' } }
        catch { & schtasks.exe /Delete /F /TN $antigo 2>$null | Out-Null }
    }
}
function Remover-Agendamento {
    $nome = [string]$Config.NomeTarefa; $ok = $false
    foreach ($antigo in @('Ostel - Torre de Controle SAC - Vendas')) { if ($antigo -ne $nome) { try { if (Get-ScheduledTask -TaskName $antigo -ErrorAction SilentlyContinue) { Unregister-ScheduledTask -TaskName $antigo -Confirm:$false; $ok = $true } } catch { & schtasks.exe /Delete /F /TN $antigo 2>$null | Out-Null } } }
    try { if (Get-ScheduledTask -TaskName $nome -ErrorAction SilentlyContinue) { Unregister-ScheduledTask -TaskName $nome -Confirm:$false; $ok = $true } }
    catch { & schtasks.exe /Delete /F /TN $nome 2>$null | Out-Null; $ok = ($LASTEXITCODE -eq 0) }
    if (-not $ok) { & schtasks.exe /Delete /F /TN $nome 2>$null | Out-Null; $ok = ($LASTEXITCODE -eq 0) }
    if ($ok) { Log "Tarefa agendada '$nome' removida." 'OK' } else { Log "Nenhuma tarefa com o nome '$nome' foi encontrada." 'AVISO' }
}

# ==================================================================
#  EXECUÇÃO
# ==================================================================
if (-not $Silencioso) {
    try { $Host.UI.RawUI.WindowTitle = 'Ostel - Dashboard Vendas' } catch { }
    Write-Host ''
    Write-Host '  OSTEL TRANSPORTES  |  Dashboard Vendas' -ForegroundColor Cyan
    Write-Host '  ---------------------------------------------------' -ForegroundColor DarkCyan
}
if ($Agendar)            { try { Registrar-Agendamento; exit 0 } catch { Log ('Falha ao agendar: ' + $_.Exception.Message) 'ERRO'; exit 1 } }
if ($RemoverAgendamento) { try { Remover-Agendamento;   exit 0 } catch { Log ('Falha ao remover: ' + $_.Exception.Message) 'ERRO'; exit 1 } }

# só uma atualização por vez (Agendador + clique manual ao mesmo tempo)
$nomeTrava = 'Local\OstelTorreSACVendas_' + ($Pasta.ToLowerInvariant() -replace '[^a-z0-9]', '_')
$trava = New-Object Threading.Mutex -ArgumentList $false, $nomeTrava
$temTrava = $false
try { $temTrava = $trava.WaitOne(60000) } catch [Threading.AbandonedMutexException] { $temTrava = $true }
if (-not $temTrava) { Log 'Outra atualização ainda está rodando; esta foi cancelada.' 'AVISO'; exit 1 }

$codigo = 0; $mudou = $false
try {
    Carregar-Leitor
    Garantir-ChartJs
    $script:Avisos = New-Object Collections.Generic.List[string]

    # ---- localiza as planilhas (uma pode faltar: o painel mostra a outra e avisa)
    $origens = @{}; $faltas = @{}
    foreach ($par in @(@('cte', $Config.ArquivoCTe, $Config.ProcurarCTe), @('cotacoes', $Config.ArquivoCotacoes, $Config.ProcurarCotacoes))) {
        try { $origens[$par[0]] = Get-Origem $par[1] $par[2] $Fontes[$par[0]].Nome }
        catch [IO.FileNotFoundException] { $faltas[$par[0]] = $_.Exception.Message; Log $_.Exception.Message 'ERRO' }
    }
    $arqMetas = Caminho-Metas

    # ---- assinatura (caminho + data + tamanho) das planilhas, das metas e do dados.js
    $partes = New-Object Collections.Generic.List[string]
    foreach ($k in @('cte', 'cotacoes')) {
        if ($origens.ContainsKey($k)) { $o = $origens[$k]; $partes.Add($k + '=' + $o.FullName + '|' + $o.LastWriteTimeUtc.Ticks + '|' + $o.Length) } else { $partes.Add($k + '=ausente') }
    }
    if ([IO.File]::Exists($arqMetas)) { $fm = New-Object IO.FileInfo($arqMetas); $partes.Add('metas=' + $fm.LastWriteTimeUtc.Ticks + '|' + $fm.Length) }
    $arqOco = Caminho-Ocorrencias
    if ([IO.File]::Exists($arqOco)) { $fo = New-Object IO.FileInfo($arqOco); $partes.Add('ocorrencias=' + $fo.LastWriteTimeUtc.Ticks + '|' + $fo.Length) }
    $arqCom = Caminho-Comissao
    if ([IO.File]::Exists($arqCom)) { $fco = New-Object IO.FileInfo($arqCom); $partes.Add('comissao=' + $fco.LastWriteTimeUtc.Ticks + '|' + $fco.Length) }
    $arqPre = Caminho-Premios
    if ([IO.File]::Exists($arqPre)) { $fpr = New-Object IO.FileInfo($arqPre); $partes.Add('premios=' + $fpr.LastWriteTimeUtc.Ticks + '|' + $fpr.Length) }
    $arqFer = Caminho-Feriados
    if ([IO.File]::Exists($arqFer)) { $ff = New-Object IO.FileInfo($arqFer); $partes.Add('feriados=' + $ff.LastWriteTimeUtc.Ticks + '|' + $ff.Length) }
    $arqCorr = Caminho-Correcoes
    if ([IO.File]::Exists($arqCorr)) { $fc = New-Object IO.FileInfo($arqCorr); $partes.Add('correcoes=' + $fc.LastWriteTimeUtc.Ticks + '|' + $fc.Length) }
    # versão do próprio atualizador: um atualizar.ps1 novo refaz o dados.js na primeira execução
    $fa = New-Object IO.FileInfo($ScriptPath); $partes.Add('script=' + $fa.LastWriteTimeUtc.Ticks + '|' + $fa.Length)
    # painel novo (app.js, index.html, style.css): refaz também o arquivo para envio, que leva o painel dentro
    foreach ($arqPainel in @('app.js', 'index.html', 'style.css')) {
        $fp = Join-Path $Pasta $arqPainel
        if ([IO.File]::Exists($fp)) { $fi = New-Object IO.FileInfo($fp); $partes.Add($arqPainel + '=' + $fi.LastWriteTimeUtc.Ticks + '|' + $fi.Length) }
    }
    if ([IO.File]::Exists($Saida)) { $fs = New-Object IO.FileInfo($Saida); $partes.Add('saida=' + $fs.LastWriteTimeUtc.Ticks + '|' + $fs.Length) }
    $estado = $partes -join ';'
    $anterior = ''; if ([IO.File]::Exists($ArqEstado)) { $anterior = ([IO.File]::ReadAllText($ArqEstado)).Trim() }

    if (-not $Forcar -and $anterior -eq $estado) {
        if (-not $Silencioso) { Log 'Sem alterações nas planilhas. Nada a fazer (use 1_atualizar_agora.bat para refazer).' }
        if ($faltas.Count -eq 2) { $codigo = 2 }
    } else {
        if ($faltas.Count -eq 2) {
            $msg = 'As duas planilhas (Comercial CTe e Comercial Cotações) não foram encontradas.'
            $js = "/* Gerado por atualizar.ps1 em " + (Get-Date -Format 'dd/MM/yyyy HH:mm:ss') + " - NAO EDITE */`r`nwindow.OSTEL_DATA = {`r`n  ok: false,`r`n  geradoEm: " +
                  (JsonStr (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss')) + ",`r`n  erro: " + (JsonStr $msg) + ",`r`n  dica: " +
                  (JsonStr 'Confira os caminhos ArquivoCTe e ArquivoCotacoes no config.ps1 (ou coloque os arquivos na pasta dados do painel) e execute 1_atualizar_agora.bat.') + "`r`n};`r`n"
            Gravar-Atomico $Saida $js
            $codigo = 2
        } else {
            $blocos = @{}
            foreach ($k in @('cte', 'cotacoes')) {
                if (-not $origens.ContainsKey($k)) {
                    $blocos[$k] = Fonte-Erro '' $faltas[$k] 'Confira o caminho no config.ps1.'
                    $script:Avisos.Add($faltas[$k]); continue
                }
                $aba = if ($k -eq 'cte') { [string]$Config.AbaCTe } else { [string]$Config.AbaCotacoes }
                try { $blocos[$k] = Ler-Fonte $k $origens[$k] $aba }
                catch [FormatException] {
                    $m = Msg $_; Log $m 'ERRO'; $script:Avisos.Add($m)
                    $blocos[$k] = Fonte-Erro $origens[$k].Name $m 'Confira se é a planilha certa, com a linha de títulos. Depois execute 1_atualizar_agora.bat.'
                }
            }
            $ocorrencias = '{}'
            try { $ocorrencias = Ler-Ocorrencias } catch { $script:Avisos.Add('ocorrencias_ssw não pôde ser lido: ' + (Msg $_)) }
            $correcoes = '[]'
            try { $correcoes = Ler-Correcoes } catch { $script:Avisos.Add('correcoes_vendedora.csv não pôde ser lido: ' + (Msg $_)) }
            $comissao = '[]'
            try { $comissao = Ler-Comissao } catch { $script:Avisos.Add('comissao.csv não pôde ser lido: ' + (Msg $_)) }
            $premios = '[]'
            try { $premios = Ler-Premios } catch { $script:Avisos.Add('premios.csv não pôde ser lido: ' + (Msg $_)) }
            $feriados = '[]'
            try { $feriados = Ler-Feriados } catch { $script:Avisos.Add('Feriados.csv não pôde ser lido: ' + (Msg $_)) }
            $metas = '[]'
            try { $metas = Ler-Metas } catch { $script:Avisos.Add('metas.csv não pôde ser lido: ' + (Msg $_)); Log ('metas.csv: ' + (Msg $_)) 'AVISO' }
            foreach ($a in $script:Avisos) { Log $a 'AVISO' }
            $js = "/* Gerado por atualizar.ps1 em " + (Get-Date -Format 'dd/MM/yyyy HH:mm:ss') + " - NAO EDITE */`r`n" +
                  "window.OSTEL_DATA = {`r`n  ok: true,`r`n  geradoEm: " + (JsonStr (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss')) + ",`r`n" +
                  "  avisos: [" + (($script:Avisos | ForEach-Object { JsonStr $_ }) -join ',') + "],`r`n" +
                  "  metas: " + $metas + ",`r`n" +
                  "  metasArquivo: " + (JsonStr (Split-Path -Leaf $arqMetas)) + ",`r`n" +
                  "  correcoes: " + $correcoes + ",`r`n" +
                  "  ocorrencias: " + $ocorrencias + ",`r`n" +
                  "  feriados: " + $feriados + ",`r`n" +
                  "  premios: " + $premios + ",`r`n" +
                  "  comissao: " + $comissao + ",`r`n" +
                  "  cte: " + $blocos['cte'] + ",`r`n" +
                  "  cotacoes: " + $blocos['cotacoes'] + "`r`n};`r`n"
            Gravar-Atomico $Saida $js
            Log ('dados.js atualizado (' + [Math]::Round((New-Object IO.FileInfo($Saida)).Length / 1KB) + ' KB).') 'OK'
            $mudou = $true
        }
        # guarda a assinatura já com o dados.js novo
        $fs = New-Object IO.FileInfo($Saida)
        $estado = (($partes | Where-Object { $_ -notlike 'saida=*' }) + ('saida=' + $fs.LastWriteTimeUtc.Ticks + '|' + $fs.Length)) -join ';'
        [IO.File]::WriteAllText($ArqEstado, $estado, $Utf8)
    }
}
catch {
    # erro temporário (ex.: arquivo bloqueado): mantém o último dados.js válido
    Log ('Falha temporária, o painel continua com os últimos dados: ' + (Msg $_)) 'ERRO'
    $codigo = 1
}
finally {
    Gravar-Status ($codigo -eq 0) $(if ($codigo -eq 0) { 'ok' } elseif ($codigo -eq 1) { 'falha temporária ao ler as planilhas' } else { 'planilha ausente ou inválida' })
    if ($temTrava) { try { $trava.ReleaseMutex() } catch { } }
}

# ---- planilha mudou: inicia o Gerar_Versao_Para_Enviar.bat sozinho (sem janela)
if ($mudou -and $Config.GerarEnvioAoMudar -and -not $SemEnvio) {
    $bat = Join-Path $Pasta 'Gerar_Versao_Para_Enviar.bat'
    if ([IO.File]::Exists($bat)) {
        try {
            if ($env:OS -eq 'Windows_NT') { Start-Process -FilePath $bat -ArgumentList '/silencioso' -WindowStyle Hidden -WorkingDirectory $Pasta }
            else { & pwsh -NoProfile -File (Join-Path $Pasta 'gerar_versao.ps1') -Silencioso }
            Log 'Gerar_Versao_Para_Enviar.bat iniciado (planilha atualizada).' 'OK'
        } catch { Log ('Não foi possível iniciar o Gerar_Versao_Para_Enviar.bat: ' + $_.Exception.Message) 'AVISO' }
    }
}
# ---- publicação automática no site (opcional)
if ($mudou -and $Config.PublicarAutomaticamente -and ([string]$Config.FtpHost).Trim()) {
    try { & (Join-Path $Pasta 'publicar.ps1') -NaoInterativo -Silencioso:$Silencioso } catch { Log ('Publicação automática falhou: ' + $_.Exception.Message) 'AVISO' }
}
exit $codigo
