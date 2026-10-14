<#
  =====================================================================
   OSTEL TRANSPORTES - DASHBOARD VENDAS - CONFIGURAÇÃO
  =====================================================================
   Este arquivo é lido pelos scripts atualizar.ps1, gerar_versao.ps1
   e publicar.ps1. Altere somente o que estiver entre aspas ou os números.
   Linhas que começam com # são comentários e são ignoradas.
   Edite com o Bloco de Notas e salve normalmente.
   IMPORTANTE: nunca coloque senhas neste arquivo.
  =====================================================================
#>

$Config = @{

    # -----------------------------------------------------------------
    # 1) PLANILHAS DE ORIGEM (só são LIDAS, nunca alteradas)
    # -----------------------------------------------------------------
    # Planilha de VENDAS (CT-e emitidos). Aceita .xlsx, .xlsm, .csv ou .txt
    # (.xlsb/.xls precisam do Excel instalado).
    ArquivoCTe = 'C:\Users\André\OneDrive\Dashboard Ostel\Comercial CTe.xlsx'
    # Aba lida. Vazio ('') = a primeira aba da planilha.
    AbaCTe = ''

    # Planilha de COTAÇÕES (todas as cotações do dia a dia)
    ArquivoCotacoes = 'C:\Users\André\OneDrive\Dashboard Ostel\Comercial Cotações.xlsx'
    AbaCotacoes = ''

    # Se o caminho acima não existir neste computador (outro usuário do
    # Windows, outra pasta do OneDrive...), o atualizador procura:
    #   1) um arquivo com o MESMO NOME na pasta ACIMA do painel;
    #   2) o arquivo mais recente da pasta "dados\" do painel cujo nome
    #      contenha o texto abaixo (ex.: "cte_outubro.csv", "cotacoes.csv").
    ProcurarCTe       = 'cte'
    ProcurarCotacoes  = 'cota'

    # Metas do mês (geral e por vendedora): arquivo CSV que a empresa
    # preenche todo mês. Abra no Excel, preencha e salve como CSV.
    ArquivoMetas = 'dados\metas.csv'

    # Correções de vendedora: CT-e emitido com a vendedora errada. Uma linha por
    # CT-e (CT-e ; Vendedora correta), ex.: SPO070331-1;Jussara
    ArquivoCorrecoes = 'dados\correcoes_vendedora.csv'

    # Tabela de ocorrências do SSW (COD ; DESCRICAO): troca o número da coluna
    # "Codigo da Ultima Ocorrencia" pelo texto. Aceita .csv ou .xlsx.
    ArquivoOcorrencias = 'dados\ocorrencias_ssw.csv'

    # Feriados (Data ; Dia da semana ; Feriado ; Abrangência ; Tipo ; Folga (S/N)).
    # Folga = S: o dia não conta como dia útil na meta. Folga = N: conta.
    ArquivoFeriados = 'dados\Feriados.csv'

    # Prêmios das vendedoras por faixa de faturamento do mês
    # (Faixa ; Faturamento mínimo ; Prêmio ; Exige bater a meta (S/N)).
    ArquivoPremios = 'dados\premios.csv'

    # Comissão das vendedoras sobre o faturamento do mês (Vendedora ; Percentual).
    # "Todas;1" = 1% para todas; uma linha com o nome muda só para ela.
    ArquivoComissao = 'dados\comissao.csv'

    # Separador do CSV: '' = detectar sozinho | ';' | ',' | "`t" (tab)
    Delimitador = ''

    # -----------------------------------------------------------------
    # 2) TIPOS DAS COLUNAS (pelo NOME do cabeçalho, sem acento/maiúscula)
    # -----------------------------------------------------------------
    # Colunas cujo nome CONTÉM um destes textos viram DATA:
    PalavrasData   = @('data', 'previsao', 'validade', 'entrega programada')
    # ...e estas ficam sempre como TEXTO (códigos, documentos, placas):
    PalavrasTexto  = @('cnpj', 'cpf', 'serie', 'numero', 'cotacao', 'ctrc', 'nota fiscal', 'placa',
                       'codigo', 'manifesto', 'romaneio', 'cep', 'telefone', 'contato', 'tabela')
    # ...mas se o nome tiver um destes textos, é VALOR (número), ex.: "FRETE CTRC":
    PalavrasNumero = @('frete', 'valor', 'peso', 'cubagem', 'qtd', 'quantidade', 'proposta', 'desc', 'distancia')
    # Colunas sem nome ou com nome só de números (ex.: a coluna A "1") são ignoradas.

    # -----------------------------------------------------------------
    # 3) ATUALIZAÇÃO AUTOMÁTICA (Agendador de Tarefas do Windows)
    # -----------------------------------------------------------------
    # De quantos em quantos minutos as planilhas são conferidas (mínimo 1)
    IntervaloMinutos = 5
    # Nome da tarefa que aparece no Agendador de Tarefas
    NomeTarefa = 'Ostel - Dashboard Vendas'

    # -----------------------------------------------------------------
    # 4) VERSÃO PARA ENVIO (um só arquivo .html, abre sem internet)
    # -----------------------------------------------------------------
    # Quando uma planilha muda, o atualizador inicia sozinho o
    # Gerar_Versao_Para_Enviar.bat. $false = não iniciar.
    GerarEnvioAoMudar = $true
    # Nome do arquivo (SEMPRE o mesmo: o anterior é substituído)
    NomeArquivoEnvio = 'Vendas.html'
    # Pasta onde ele é gravado. Vazio ('') = pasta "enviar" do painel.
    PastaArquivoEnvio = ''

    # -----------------------------------------------------------------
    # 5) PUBLICAÇÃO NO SITE VIA FTP/FTPS (opcional - publicar.ps1)
    #    A senha é pedida na 1ª publicação e guardada CRIPTOGRAFADA
    #    (só o seu usuário do Windows, neste computador, consegue ler).
    # -----------------------------------------------------------------
    FtpHost         = ''                          # ex.: 'ftp.ostel.com.br'
    FtpPorta        = 21
    FtpUsuario      = ''                          # ex.: 'usuario@ostel.com.br'
    FtpPasta        = '/public_html/vendas'       # pasta no servidor
    FtpUsarSSL      = $true                       # FTPS (recomendado)
    FtpPassivo      = $true
    # $true = publica no site sozinho depois de cada atualização dos dados
    PublicarAutomaticamente = $false

    # -----------------------------------------------------------------
    # 6) MANUTENÇÃO
    # -----------------------------------------------------------------
    # Tamanho máximo do logs\atualizacao.log (MB). Passou: vira .log.1
    LogMaxMB = 5
    # Endereço para baixar o Chart.js caso assets\chart.umd.min.js falte
    UrlChartJs = 'https://cdn.jsdelivr.net/npm/chart.js@4.5.1/dist/chart.umd.min.js'
}
