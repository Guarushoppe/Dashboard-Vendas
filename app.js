/* =====================================================================
   OSTEL TRANSPORTES - DASHBOARD VENDAS
   app.js  |  lógica do painel (JavaScript puro + Chart.js)
   ---------------------------------------------------------------------
   Fluxo:
     1. carrega o dados.js (gerado pelo atualizar.ps1 a partir das
        planilhas Comercial CTe e Comercial Cotações) e o dados_status.js;
     2. normaliza as linhas (descobre a vendedora, a origem da carga,
        o que é cancelado...);
     3. aplica os filtros e desenha meta, indicadores, cards das
        vendedoras, gráficos e tabela;
     4. relê os dados a cada CONFIG.recarregarSegundos segundos e, só se
        algo mudou, redesenha tudo mantendo os filtros escolhidos.
   Se o dados.js faltar ou vier com erro, mostra uma mensagem amigável.

   REGRA ÚNICA DE CÁLCULO (igual nos indicadores, gráficos, tabela e Excel):
     - documento CANCELADO (CT-e com Tipo de Baixa "CANCELADO" ou Data do
       Cancelamento preenchida; cotação com Situação "CANCELAD...") entra
       só nas QUANTIDADES, nunca nas somas de valores;
     - Faturamento = soma do "Valor do Frete" dos CT-e não cancelados;
     - Coletado = coluna E (Placa de Coleta) com placa; Armazém = coluna E
       com "ARMAZEM";
     - Valor cotado (frete atual) = soma da "PROPOSTA ATUAL" das cotações não canceladas;
       frete inicial = soma da "PROPOSTA INICIAL"; variação = (atual − inicial) ÷ inicial;
     - Taxa de fechamento = CT-e não cancelados ÷ cotações não canceladas.
   ===================================================================== */
(function () {
  'use strict';

  /* ===================================================================
     AJUSTES SIMPLES DO PAINEL (pode mudar à vontade)
     =================================================================== */
  var CONFIG = {
    arquivoDados: 'dados.js',          // gerado pelo atualizar.ps1
    arquivoStatus: 'dados_status.js',  // gravado a cada execução do atualizar.ps1
    recarregarSegundos: 30,            // de quanto em quanto tempo o painel relê os dados
    alertaDesatualizadoHoras: 2,       // alerta vermelho se as planilhas não forem conferidas há mais de X horas
    // período ao abrir (e para onde "Limpar filtros" volta): 'mes' = mês atual | 'ano' = ano atual | 'tudo'
    periodoInicial: 'mes',
    // ROLAGEM AUTOMÁTICA (painel numa TV): depois deste tempo sem ninguém mexer, a página
    // desce sozinha, devagar, rola a tabela, volta ao topo e recomeça. 0 = desligada.
    rolagemAutomaticaSegundos: 45,
    rolagemVelocidade: 45,             // pixels por segundo
    linhasPorPagina: 25,
    // VENDEDORA: nome usado quando não há vendedora na coluna V nem na coluna BC
    vendedoraPadrao: 'Ostel',
    // apelidos -> nome que aparece no painel ('' = não é vendedora). Escreva em minúsculas, sem acento.
    apelidos: { jussarab: 'Jussara', ritac: 'Rita', monica: 'Mônica', sistema: 'Sistema', webservice: '' },
    // outros nomes de vendedora que podem aparecer na Observação (coluna BC)
    vendedorasExtras: [],
    // palavras que NUNCA são nome de vendedora (quando o nome vem depois de "TRATATIVAS", "VENDEDORA"...)
    naoSaoNomes: ['DE', 'DO', 'DA', 'POR', 'A', 'O', 'E', 'COM', 'COTACAO', 'CLIENTE', 'FEITA', 'FEITAS', 'NF', 'NOTA', 'CTRC', 'CT', 'SEM', 'PARA'],
    // CT-e AGUARDANDO EMBARQUE: não cancelado, sem "Primeiro Manifesto" e com a "Localizacao Atual" SEM
    // nenhum destes textos (ou seja: ainda não saiu da origem). Escreva em maiúsculas.
    jaEmbarcou: ['ENTREGUE', 'TRANSITO', 'EM ENTREGA'],
    // COMPLEMENTO DE CARGA: CT-e sem saída (aguardando embarque) com Valor do Frete ACIMA deste valor
    complementoAcimaDe: 1500,
    // ENTREGAS EM ATRASO: CT-e sem "Data da Entrega Realizada" cujo prazo (Entrega Programada, ou
    // Previsão de Entrega) já passou. Códigos da última ocorrência que NÃO contam como atraso:
    // 93 = CTRC emitido p/ efeito frete/imposto, 94 = CT-e baixado/substituído, 23 = devolvida ao remetente
    atrasoIgnorarOcorr: ['93', '94', '23'],
    // faixas de dias de atraso (até N dias); a última faixa é "mais de" o último número
    atrasoFaixas: [2, 5, 10, 30],
    // v51 ENTREGAS NO PRAZO (ainda não realizadas): faixas de dias que faltam (0 = vence hoje)
    prazoFaixas: [0, 1, 3, 7],
    // CANCELADOS: CT-e e cotações canceladas ficam FORA do painel inteiro (não aparecem em nada)
    ocultarCancelados: true,
    versao: '57'                        // número da versão do painel (aparece no rodapé)
  };

  /* ===================================================================
     MAPEAMENTO DAS COLUNAS (ajuste aqui se o layout da planilha mudar)
     Cada campo aceita o NOME do cabeçalho e/ou a LETRA da coluna.
     O nome é procurado primeiro (sem acento e sem maiúsculas);
     se não achar, vale a letra (ex.: 'V' = coluna V da planilha).
     =================================================================== */
  var MAPA = {
    cte: {
      numero:       ['Serie/Numero CTRC', 'B'],
      tipoDoc:      ['Tipo do Documento', 'C'],
      emissao:      ['Data de Emissao', 'D'],
      placaColeta:  ['Placa de Coleta', 'E'],          // coluna E: placa = coletado | ARMAZEM = entregue no armazém
      cnpjRemetente:['CNPJ Remetente', 'F'],
      remetente:    ['Cliente Remetente', 'G'],
      cnpjPagador:  ['CNPJ Pagador', 'J'],
      cnpjDest:     ['CNPJ Destinatario', 'L'],          // v57: duplicidade
      cubagem:      ['Cubagem em m3', 'R'],
      pagador:      ['Cliente Pagador', 'K'],
      destinatario: ['Cliente Destinatario', 'M'],
      cidade:       ['Cidade de Entrega', 'N'],
      uf:           ['UF de Entrega', 'O'],
      nf:           ['Numero da Nota Fiscal', 'P'],
      peso:         ['Peso Real em Kg', 'Q'],
      baixa:        ['Tipo de Baixa', 'T'],            // status do CT-e
      vlMerc:       ['Valor da Mercadoria', 'U'],
      vendedoraV:   ['Tipo de Calculo', 'V'],          // coluna V: "COTACAO - renata"
      frete:        ['Valor do Frete', 'W'],
      freteSI:      ['Valor do Frete sem ICMS', 'X'],
      manifesto:    ['Primeiro Manifesto', 'Z'],
      ocorrencia:   ['Complemento da Ultima Ocorrencia', 'AR'],   // texto da última ocorrência
      codOcorr:     ['Codigo da Ultima Ocorrencia', 'AM'],        // número -> texto pela tabela dados\ocorrencias_ssw.csv
      localizacao:  ['Localizacao Atual', 'AY'],
      previsao:     ['Previsao de Entrega', 'AU'],        // v50: entregas em atraso
      programada:   ['Entrega Programada', 'AV'],
      entregue:     ['Data da Entrega Realizada', 'AW'],
      dtCancel:     ['Data do Cancelamento', 'AZ'],
      motivoCancel: ['Motivo do Cancelamento', 'BA'],
      obs:          ['Observacao', 'BC']               // coluna BC: "RENATA 1080528", "VENDEDORA MONICA - ..."
    },
    cot: {
      numero:       ['COTACAO', 'A'],
      usuario:      ['USUARIO INCLUSAO', 'C'],         // quem fez a cotação
      vendedor:     ['VENDEDOR', 'G'],
      cnpjPagador:  ['CNPJ PAGADOR', 'D'],
      pagador:      ['NOME PAGADOR', 'E'],
      origem:       ['ORIGEM', 'H'],
      destino:      ['DESTINO', 'K'],                  // "CIDADE/UF"
      vlNF:         ['VALOR NF', 'Q'],
      cubagem:      ['CUBAGEM', 'U'],                  // v57: duplicidade
      peso:         ['PESO', 'T'],
      propostaIni:  ['PROPOSTA INICIAL', 'X'],
      proposta:     ['PROPOSTA ATUAL', 'Y'],
      data:         ['DATA HORA INCLUSAO', 'AD'],
      situacao:     ['SITUACAO', 'AF'],
      ctrc:         ['CTRC', 'AG'],
      freteCtrc:    ['FRETE CTRC', 'AI']
    },
    // colunas que entram na linha TOTAL da tabela (só documentos não cancelados) e as que são dinheiro
    somarCte:    ['Peso Real em Kg', 'Cubagem em m3', 'Quantidade de Volumes', 'Valor da Mercadoria', 'Valor do Frete', 'Valor do Frete sem ICMS', 'Peso Calculado em Kg'],
    dinheiroCte: ['Valor da Mercadoria', 'Valor do Frete', 'Valor do Frete sem ICMS'],
    somarCot:    ['VALOR NF', 'QTD VOLUMES', 'QTD PARES', 'PESO', 'CUBAGEM', 'PESO CALCULO', 'FRETE NTC', 'PROPOSTA INICIAL', 'PROPOSTA ATUAL', 'FRETE CTRC'],
    dinheiroCot: ['VALOR NF', 'FRETE NTC', 'PROPOSTA INICIAL', 'PROPOSTA ATUAL', 'FRETE CTRC'],
    // colunas que NÃO aparecem no Detalhamento (tabela, CSV e ficha). Nome exato do título da planilha.
    ocultarCte:  ['Conferente de Coleta', 'Complemento', 'Localizacao Atual', 'Data do Cancelamento', 'Motivo do Cancelamento', 'Codigo dos Correios',
                  'Latitude da Ultima Ocorrencia', 'Longitude da Ultima Ocorrencia',
                  // v23
                  'Tipo do Documento', 'CNPJ Remetente', 'Cidade do Remetente', 'UF do Remetente', 'CNPJ Destinatario', 'Cliente Destinatario',
                  'Tipo de Baixa', 'Valor do Frete sem ICMS', 'Conferente de Carga', 'Conferente de Descarga',
                  'Data de inclusao da Ultima Ocorrencia', 'Data da Ultima Ocorrencia', 'Usuario da Ultima Ocorrencia', 'Unidade da Ultima Ocorrencia'],
    ocultarCot:  ['DESC NTC', 'RC', 'DESC INICIAL']   // v55: colunas que não aparecem no Detalhamento das cotações
  };

  /* ===================================================================
     FUNÇÕES AUXILIARES (formatação, DOM, etc.)
     =================================================================== */
  var $  = function (s, r) { return (r || document).querySelector(s); };
  var $$ = function (s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); };

  var MESES   = ['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez'];
  var MESES_L = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];

  var fBRL = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' });
  var fInt = new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 0 });
  var fDec = new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 1 });
  var fNum = new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 3 });

  function brl(v)  { return fBRL.format(v || 0); }                 // R$ 1.234,56 (dinheiro sempre completo)
  function brlc(v) {                                               // só nas escalas dos eixos: R$ 1,2 mi | R$ 350 mil
    v = v || 0; var a = Math.abs(v);
    if (a >= 1e6) return 'R$ ' + fDec.format(v / 1e6) + ' mi';
    if (a >= 1e3) return 'R$ ' + fDec.format(v / 1e3) + ' mil';
    return 'R$ ' + fInt.format(v);
  }
  function nInt(v) { return fInt.format(v || 0); }
  function pct(v)  { return isFinite(v) ? fDec.format(v * 100) + '%' : '—'; }
  function esc(s)  { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function fDate(s){ if (!s) return '—'; var p = String(s).slice(0, 10).split('-'); return p[2] + '/' + p[1] + '/' + p[0]; }
  function fDataHora(s) { if (!s) return '—'; s = String(s); return fDate(s) + (s.length > 10 ? ' ' + s.slice(11, 16) : ''); }
  function css(v)  { return getComputedStyle(document.documentElement).getPropertyValue(v).trim(); }
  function sum(a, f) { var s = 0; for (var i = 0; i < a.length; i++) { var v = f(a[i]); if (v) s += v; } return s; }
  function titulo(s) { return String(s).toLowerCase().replace(/(^|[\s\/-])\S/g, function (t) { return t.toUpperCase(); }); }
  function cortar(s, n) { s = String(s); return s.length > n ? s.slice(0, n - 1) + '…' : s; }
  function debounce(fn, ms) { var t; return function () { var a = arguments; clearTimeout(t); t = setTimeout(function () { fn.apply(null, a); }, ms); }; }
  function guardar(k, v) { try { if (v === undefined) return localStorage.getItem(k); localStorage.setItem(k, v); } catch (e) { return null; } }
  function fmtStamp(s) { if (!s) return '—'; var d = new Date(s); return isNaN(d) ? s : d.toLocaleString('pt-BR', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }); }
  function semAcento(t) {
    t = String(t == null ? '' : t).toLowerCase();
    try { t = t.normalize('NFD').replace(/[̀-ͯ]/g, ''); } catch (e) { /* navegador antigo */ }
    return t.replace(/\s+/g, ' ').trim();
  }
  function chave(t) { return semAcento(t).toUpperCase(); }
  function alfa(cor, a) {
    var m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(String(cor).trim());
    if (!m) return cor;
    return 'rgba(' + parseInt(m[1], 16) + ',' + parseInt(m[2], 16) + ',' + parseInt(m[3], 16) + ',' + a + ')';
  }
  function iso(d) { return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }
  function dataDe(s) { var p = s.split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); }
  function nomeMes(ym) { return MESES[+ym.slice(5, 7) - 1] + '/' + ym.slice(0, 4); }
  function nomeMesLongo(ym) { return MESES_L[+ym.slice(5, 7) - 1] + ' de ' + ym.slice(0, 4); }
  function hojeIso() { return iso(new Date()); }

  /* ícones usados nos indicadores, na meta e na rota do cabeçalho */
  var IC = {
    truck:   '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"><path d="M2 6h12v10H2zM14 9h4.5L22 12.5V16h-8z"/><circle cx="6.5" cy="17.5" r="2" fill="currentColor"/><circle cx="17.5" cy="17.5" r="2" fill="currentColor"/></svg>',
    money:   '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="6" width="20" height="12" rx="2"/><circle cx="12" cy="12" r="3"/></svg>',
    receipt: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"><path d="M5 3h14v18l-3-2-2 2-2-2-2 2-2-2-3 2z"/><path d="M9 8h6M9 12h6"/></svg>',
    doc:     '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"><path d="M14 3H6a1 1 0 0 0-1 1v16a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V8z"/><path d="M14 3v5h5M9 13h6M9 17h6"/></svg>',
    check:   '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M8 12.5l2.5 2.5L16 9.5"/></svg>',
    box:     '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"><path d="M3 21V9l9-6 9 6v12"/><path d="M7 21v-8h10v8M7 17h10"/></svg>',
    ticket:  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"><path d="M3 8a2 2 0 0 0 0 4 2 2 0 0 1 0 4v2h18v-2a2 2 0 0 1 0-4 2 2 0 0 0 0-4V6H3z"/></svg>',
    tag:     '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"><path d="M3 12V4h8l10 10-8 8z"/><circle cx="7.5" cy="8.5" r="1.5" fill="currentColor"/></svg>',
    alvo:    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="5"/><circle cx="12" cy="12" r="1.5" fill="currentColor"/></svg>',
    cal:     '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4"/></svg>',
    pin:     '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 21s7-6.2 7-11.5A7 7 0 0 0 5 9.5C5 14.8 12 21 12 21z"/><circle cx="12" cy="9.5" r="2.5"/></svg>',
    building:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 21V5l8-3v19M12 8h8v13M2 21h20M7 9h2M7 13h2M15 12h2M15 16h2"/></svg>',
    alert:   '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M12 3l10 18H2z"/><path d="M12 10v5M12 18v.01"/></svg>',
    file:    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"><path d="M14 3H6a1 1 0 0 0-1 1v16a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V8z"/><path d="M14 3v5h5M9 13h6M9 17h6"/></svg>'
  };

  /* ===================================================================
     TEMA CLARO / ESCURO (botão lua/sol; a escolha fica salva)
     =================================================================== */
  function temaAtual() {
    var t = document.documentElement.getAttribute('data-theme');
    if (t) return t;
    return matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
  }
  $('#btnTheme').addEventListener('click', function () {
    var novo = temaAtual() === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', novo);
    guardar('ostel-theme', novo);
    redesenharGraficos(true);
  });
  try {
    matchMedia('(prefers-color-scheme: dark)').addEventListener('change', function () {
      if (!document.documentElement.getAttribute('data-theme')) redesenharGraficos(true);
    });
  } catch (e) { /* navegador antigo */ }

  /* ===================================================================
     ESTADO DA APLICAÇÃO
     =================================================================== */
  var RAW = null;             // objeto vindo do dados.js
  var STATUS = null;          // objeto vindo do dados_status.js
  var CTE = [], COT = [];     // linhas normalizadas
  var FC = [], FQ = [];       // linhas após os filtros (CT-e e cotações)
  var METAS = {};             // METAS['2026-10']['RENATA'] = valor ; ['TOTAL'] = meta geral
  var NOMES = {};             // chave da vendedora -> nome para mostrar
  var CARIMBO = null;         // identifica a versão do dados.js já carregada
  var GRAFICOS = {};          // instâncias do Chart.js, por id do canvas
  var IDX = { cte: {}, cot: {} };   // posição de cada campo do MAPA nas colunas do dados.js

  // filtros ativos (vazio = sem filtro)
  var F = { de: '', ate: '', ano: '', mes: '', ym: '', dia: '', status: '', vendedora: '', uf: '', cliente: '', origem: '', situacao: '', motivo: '', ocorr: '', codoc: '', compl: '', atraso: '', q: '' };
  var F_ROTULO = { de: 'De', ate: 'Até', ano: 'Ano', mes: 'Mês', ym: 'Mês/ano', dia: 'Dia', status: 'Status do CT-e', vendedora: 'Vendedora', uf: 'UF',
                   cliente: 'Cliente', origem: 'Origem da carga', situacao: 'Situação da cotação', motivo: 'Motivo do cancelamento', ocorr: 'Ocorrência (aguardando embarque)', codoc: 'Última ocorrência do CT-e', compl: 'Complemento de carga', atraso: 'Entregas', q: 'Busca' };
  // filtros que só valem para uma das planilhas
  var SO_CTE = { status: 1, origem: 1, motivo: 1, ocorr: 1, codoc: 1, compl: 1, atraso: 1 }, SO_COT = { situacao: 1 };
  var ORIGENS = { COLETA: 'Coletado (placa)', ARMAZEM: 'Entregue no armazém', SEM: 'Sem informação' };

  // escolhas de visualização (botões segmentados, tabela)
  var UI = { evo: 'p', evoMed: 'fat', uf: 'fat', vend: 'fat', cli: 'fat', baixa: 'ocor',
             aba: 'cte', ordem: -1, dir: -1, pagina: 1, porPagina: CONFIG.linhasPorPagina, busca: '' };

  /* ===================================================================
     PREPARAÇÃO DOS DADOS
     =================================================================== */
  // acha a posição de cada campo do MAPA: primeiro pelo nome do cabeçalho, depois pela letra
  function resolverColunas(fonte, mapa) {
    var idx = {}, nomes = (fonte.colunas || []).map(semAcento), letras = fonte.letras || [];
    Object.keys(mapa).forEach(function (campo) {
      var alvos = mapa[campo], achou = -1;
      alvos.forEach(function (a) { if (achou < 0 && !/^[A-Z]{1,3}$/.test(a)) achou = nomes.indexOf(semAcento(a)); });
      if (achou < 0) alvos.forEach(function (a) { if (achou < 0 && /^[A-Z]{1,3}$/.test(a)) achou = letras.indexOf(a); });
      idx[campo] = achou;
    });
    return idx;
  }
  function val(r, i) { return i >= 0 && i < r.length ? r[i] : null; }
  function num(v) { if (typeof v === 'number') return v; if (v == null || v === '') return 0; var n = parseFloat(String(v).replace(/\./g, '').replace(',', '.')); return isFinite(n) ? n : 0; }
  function txt(v) { return v == null ? '' : String(v).trim(); }

  // nome que aparece no painel para um apelido/nome cru (ex.: "jussarab" -> "Jussara")
  function nomeVendedora(cru) {
    var k = semAcento(cru);
    if (!k) return '';
    if (Object.prototype.hasOwnProperty.call(CONFIG.apelidos, k)) return CONFIG.apelidos[k];
    return titulo(k);
  }
  function registrarNome(nome) { var k = chave(nome); if (k && !NOMES[k]) NOMES[k] = nome; return k; }

  /* VENDEDORA DO CT-e
     1) coluna V (Tipo de Calculo): "COTACAO - renata" -> Renata
     2) senão, coluna BC (Observação): o PRIMEIRO nome que aparecer no texto, seja de uma
        vendedora conhecida (quem faz cotações, nomes da coluna V, vendedorasExtras) ou
        o nome escrito depois de "VENDEDORA", "TRATATIVAS (FEITA POR)", "TRATADO (POR)";
     3) senão, a vendedora padrão (Ostel). */
  var CONHECIDAS = [];
  function montarConhecidas(rowsCte, rowsCot) {
    var set = {};
    var add = function (nome) { var k = chave(nome); if (k && k.length >= 3) set[k] = nome; };
    rowsCot.forEach(function (r) { add(nomeVendedora(val(r, IDX.cot.usuario))); });
    rowsCte.forEach(function (r) { var m = /COTA[CÇ][AÃ]O\s*-\s*([A-Za-zÀ-ÿ]+)/i.exec(txt(val(r, IDX.cte.vendedoraV))); if (m) add(nomeVendedora(m[1])); });
    CONFIG.vendedorasExtras.forEach(add);
    CONHECIDAS = Object.keys(set).map(function (k) { return { k: k, nome: set[k], re: new RegExp('(^|[^A-Z])' + k + '(?![A-Z])') }; });
  }
  var PADROES_NOME = [/VENDEDORA\s*:?\s*([A-Z]{3,})/, /TRATATIVAS?\s+(?:FEITAS?\s+)?(?:POR\s+)?([A-Z]{3,})/, /TRATAD[OA]\s+(?:POR\s+)?([A-Z]{3,})/];
  function vendedoraDoCte(r) {
    // TAXA DE COLETA: se a Observação (coluna BC) tiver "TAXA DE COLETA", o frete é da Ostel (vence as colunas V e BC)
    if (/TAXA\s+DE\s+COLETA/.test(chave(val(r, IDX.cte.obs)))) return { nome: CONFIG.vendedoraPadrao, fonte: 'taxa de coleta na Observação (coluna BC) → Ostel' };
    var v = txt(val(r, IDX.cte.vendedoraV)), m = /COTA[CÇ][AÃ]O\s*-\s*([A-Za-zÀ-ÿ]+)/i.exec(v);
    if (m) { var n1 = nomeVendedora(m[1]); if (n1) return { nome: n1, fonte: 'coluna V (Tipo de Calculo)' }; }
    var obs = chave(val(r, IDX.cte.obs));
    if (obs) {
      var melhor = null;
      CONHECIDAS.forEach(function (c) { var x = c.re.exec(obs); if (x) { var pos = x.index + x[1].length; if (!melhor || pos < melhor.pos) melhor = { pos: pos, nome: c.nome }; } });
      PADROES_NOME.forEach(function (re) {
        var x = re.exec(obs); if (!x || CONFIG.naoSaoNomes.indexOf(x[1]) >= 0) return;
        var pos = x.index + x[0].length - x[1].length;
        if (!melhor || pos < melhor.pos) melhor = { pos: pos, nome: nomeVendedora(x[1]) };
      });
      if (melhor && melhor.nome) return { nome: melhor.nome, fonte: 'coluna BC (Observação)' };
    }
    return { nome: CONFIG.vendedoraPadrao, fonte: 'padrão (sem vendedora nas colunas V e BC)' };
  }

  // CORREÇÕES DE VENDEDORA: número do CT-e (só letras e números, maiúsculas) -> vendedora correta
  var CORRECOES = {}, CORR_USADAS = {}, CORR_LISTA = [];
  function chaveDoc(n) { return chave(n).replace(/[^A-Z0-9]/g, ''); }
  function lerCorrecoes(lista) {
    CORRECOES = {}; CORR_USADAS = {}; CORR_LISTA = [];
    (lista || []).forEach(function (c) {
      if (!Array.isArray(c) || !c[0] || !c[1]) return;
      var nome = nomeVendedora(c[1]) || String(c[1]).trim();
      CORRECOES[chaveDoc(c[0])] = nome; CORR_LISTA.push(String(c[0]).trim());
    });
  }
  // OCORRÊNCIAS DO SSW: código -> "62 - SAIDA DA UNIDADE" (tabela dados\ocorrencias_ssw.csv)
  var OCORR = {};
  function textoOcorr(cod) { cod = String(cod == null ? '' : cod).trim(); if (!cod) return ''; return OCORR[cod] ? cod + ' - ' + OCORR[cod] : cod; }
  function dataDoValor(v) { var s = txt(v); return /^\d{4}-\d{2}-\d{2}/.test(s) ? s : ''; }

  var HOJE_ISO = hojeIso();
  var CANC_FORA = { cte: 0, cot: 0 };   // quantos cancelados ficaram fora do painel (só para informação)
  function normalizarCte(fonte) {
    CANC_FORA.cte = 0;
    var I = IDX.cte, out = [];
    fonte.linhas.forEach(function (r) {
      if (!Array.isArray(r)) return;
      var dtc = dataDoValor(val(r, I.emissao)); if (!dtc) return;
      var o = { t: 'cte', v: r, dt: dtc.slice(0, 10) };
      o.ano = o.dt.slice(0, 4); o.mes = String(+o.dt.slice(5, 7)); o.ym = o.dt.slice(0, 7);
      o.frete = num(val(r, I.frete)); o.freteSI = num(val(r, I.freteSI)); o.icms = o.frete - o.freteSI;
      o.vlMerc = num(val(r, I.vlMerc)); o.peso = num(val(r, I.peso));
      o.baixa = txt(val(r, I.baixa)).toUpperCase() || 'NÃO INFORMADO';
      o.cancelado = /CANCEL/.test(o.baixa) || !!txt(val(r, I.dtCancel));
      o.soma = !o.cancelado;                                // REGRA ÚNICA: cancelado só conta na quantidade
      var loc = chave(val(r, I.localizacao));
      o.aguardando = !o.cancelado && !txt(val(r, I.manifesto)) && !CONFIG.jaEmbarcou.some(function (x) { return loc.indexOf(x) >= 0; });
      o.ocorr = o.aguardando ? ocorrenciaDe(val(r, I.ocorrencia)) : '';
      o.complemento = o.aguardando && o.frete > CONFIG.complementoAcimaDe;   // possível complemento de carga
      o.motivo = o.cancelado ? (txt(val(r, I.motivoCancel)).toUpperCase() || 'SEM MOTIVO INFORMADO') : '';
      var placa = chave(val(r, I.placaColeta));
      o.origem = !placa ? 'SEM' : (/ARMAZ/.test(placa) ? 'ARMAZEM' : 'COLETA');
      o.placa = placa;
      o.uf = txt(val(r, I.uf)).toUpperCase() || '—';
      o.cidade = txt(val(r, I.cidade)).toUpperCase() || '—';
      o.cliente = txt(val(r, I.pagador)) || txt(val(r, I.remetente)) || '—';
      o.numero = txt(val(r, I.numero));
      o.tipoDoc = txt(val(r, I.tipoDoc));
      o.codOcorr = txt(val(r, I.codOcorr)); o.codOcorrTxt = textoOcorr(o.codOcorr); o.codoc = o.codOcorr || '—';
      // v50 ENTREGA EM ATRASO: sem entrega realizada e prazo (programada ou previsão) antes de hoje
      o.prazo = dataDoValor(val(r, I.programada)).slice(0, 10) || dataDoValor(val(r, I.previsao)).slice(0, 10);
      o.entregueEm = dataDoValor(val(r, I.entregue)).slice(0, 10);
      o.diasAtraso = 0; o.diasFalta = -1;
      if (!o.cancelado && !o.entregueEm && o.prazo && CONFIG.atrasoIgnorarOcorr.indexOf(o.codOcorr) < 0) {
        var dd = Math.round((Date.parse(o.prazo) - Date.parse(HOJE_ISO)) / 864e5);
        if (dd < 0) o.diasAtraso = -dd; else o.diasFalta = dd;   // v51: no prazo = faltam N dias (0 = vence hoje)
      }
      o.atraso = o.diasAtraso > 0;
      o.noPrazo = o.diasFalta >= 0;
      o.nf = txt(val(r, I.nf)); o.cnpjRem = txt(val(r, I.cnpjRemetente));
      // v57 INDÍCIO DE DUPLICIDADE: NF + CNPJ remetente + CNPJ destinatário + cidade de entrega + valor da NF + peso + cubagem
      o.chaveInd = [o.nf, o.cnpjRem, txt(val(r, I.cnpjDest)), chave(val(r, I.cidade)) + '/' + chave(val(r, I.uf)),
                    num(val(r, I.vlMerc)).toFixed(2), num(val(r, I.peso)).toFixed(3), num(val(r, I.cubagem)).toFixed(4)].join('|');
      var vd = vendedoraDoCte(r);
      // CORREÇÃO MANUAL (dados\correcoes_vendedora.csv) vence a regra das colunas V / BC
      var corr = CORRECOES[chaveDoc(o.numero)];
      if (corr) { vd = { nome: corr, fonte: 'corrigida no arquivo correcoes_vendedora.csv (era ' + vd.nome + ')' }; o.vendCorrigida = true; CORR_USADAS[chaveDoc(o.numero)] = 1; }
      o.vendNome = vd.nome; o.vendFonte = vd.fonte; o.vendedora = registrarNome(vd.nome);
      o._q = r.map(function (x) { return x == null ? '' : String(x); }).join(' ').toLowerCase() + ' ' + vd.nome.toLowerCase() + ' ' + o.codOcorrTxt.toLowerCase();
      if (o.cancelado && CONFIG.ocultarCancelados) { CANC_FORA.cte++; return; }   // cancelado: fora do painel
      out.push(o);
    });
    return out;
  }
  /* DUPLICIDADE NO DETALHAMENTO (letras VERMELHAS na tabela, para conferir):
     - CT-e: o mesmo número de CT-e em mais de uma linha (duplicada) ou, entre os NÃO cancelados,
       a mesma NF + CNPJ remetente + CNPJ destinatário + cidade de entrega + valor da NF + peso +
       cubagem (indício, qualquer data). CT-e Complementar, Reentrega e Devolução não contam.
     - Cotação: o mesmo número de cotação em mais de uma linha (duplicada) ou, entre as NÃO
       canceladas, o mesmo CNPJ pagador + cidade destino + valor da NF + peso + cubagem (indício).
     Só marca: os valores continuam somados como estão na planilha. */
  function marcarDuplicidades(lista, cte) {
    var porNum = {}, porInd = {}, add = function (m, k, o) { (m[k] = m[k] || []).push(o); };
    lista.forEach(function (o) {
      o.dup = false; o.ind = false; o.dupMotivo = ''; o.dupGrupo = '';
      if (o.numero) add(porNum, o.numero, o);
      if (o.soma) {
        if (cte) { if (o.nf && !/COMPLEMENT|REENTREGA|DEVOLU/i.test(o.tipoDoc)) add(porInd, o.chaveInd, o); }
        else add(porInd, o.chaveInd, o);
      }
    });
    Object.keys(porNum).forEach(function (k) {
      var g = porNum[k]; if (g.length < 2) return;
      g.forEach(function (o) { o.dup = true; o.dupGrupo = 'N' + k; o.dupMotivo = (cte ? 'o mesmo CT-e ' : 'a mesma cotação ') + k + ' aparece em ' + g.length + ' linhas'; });
    });
    Object.keys(porInd).forEach(function (k) {
      var g = porInd[k]; if (g.length < 2) return;
      g.forEach(function (o) {
        if (o.dup) return;
        o.ind = true; o.dupGrupo = 'I' + k;
        var outros = g.map(function (x) { return x.numero + ' de ' + fDate(x.dt); }).join(', ');
        o.dupMotivo = cte ? 'mesma NF ' + o.nf + ', remetente, destinatário, cidade, valor da NF, peso e cubagem em ' + g.length + ' CT-e (' + outros + ')'
                          : 'mesmo pagador, cidade destino, valor da NF, peso e cubagem em ' + g.length + ' cotações (' + outros + ')';
      });
    });
  }
  // texto da última ocorrência agrupável: sem datas, horas e números ("FRETE PAGO 09/10, CARGA LIBERADA." ->
  // "FRETE PAGO CARGA LIBERADA"); só a autorização do CT-e/RPS = "SEM OCORRÊNCIA (SÓ EMITIDO)"
  function ocorrenciaDe(v) {
    var t = chave(v);
    if (!t) return 'SEM OCORRENCIA';
    if (/^(CT-E|RPS|SUBCONTRATO) (AUTORIZADO|EMITIDA|EMITIDO)/.test(t)) return 'SEM OCORRENCIA (SO EMITIDO)';
    t = t.replace(/\d{1,2}\/\d{1,2}(\/\d{2,4})?/g, ' ').replace(/\d{1,2}:\d{2}H?/g, ' ').replace(/\b\d+\b/g, ' ')
         .replace(/[.,;:]+/g, ' ').replace(/\s+/g, ' ').trim();
    return t || 'SEM OCORRENCIA';
  }
  function nomeCodOc(k) { return k === '—' || !k ? 'Sem ocorrência informada' : textoOcorr(k); }
  function desconto(g) { var d = g.valorIni - g.valor; return { r: d, p: g.valorIni ? d / g.valorIni : 0 }; }
  function nomeOcorr(k) { return k === 'SEM OCORRENCIA (SO EMITIDO)' ? 'Sem ocorrência (só emitido)' : k === 'SEM OCORRENCIA' ? 'Sem ocorrência' : titulo(k); }
  function situacaoDe(s) {
    var k = semAcento(s);
    if (!k) return 'Não informada';
    if (/^cancel/.test(k)) return 'Cancelada';
    if (/^ctrc/.test(k)) return 'CT-e emitido';
    if (/^contrat/.test(k)) return 'Contratada';
    if (/^cotad/.test(k)) return 'Cotada';
    return titulo(k);
  }
  function normalizarCot(fonte) {
    CANC_FORA.cot = 0;
    var I = IDX.cot, out = [];
    fonte.linhas.forEach(function (r) {
      if (!Array.isArray(r)) return;
      var d = dataDoValor(val(r, I.data)); if (!d) return;
      var o = { t: 'cot', v: r, dt: d.slice(0, 10), hora: d.length > 10 ? d.slice(11, 16) : '' };
      o.ano = o.dt.slice(0, 4); o.mes = String(+o.dt.slice(5, 7)); o.ym = o.dt.slice(0, 7);
      var sit = txt(val(r, I.situacao));
      o.situacao = situacaoDe(sit);
      o.cancelada = o.situacao === 'Cancelada';
      o.soma = !o.cancelada;
      o.valor = num(val(r, I.proposta)); o.valorIni = num(val(r, I.propostaIni));
      var dest = txt(val(r, I.destino)), p = dest.lastIndexOf('/');
      o.uf = p >= 0 ? dest.slice(p + 1).trim().toUpperCase() : '—';
      o.cidade = p >= 0 ? dest.slice(0, p).trim().toUpperCase() : dest.toUpperCase();
      o.cliente = txt(val(r, I.pagador)) || '—';
      o.numero = txt(val(r, I.numero));
      // v57 INDÍCIO DE DUPLICIDADE: CNPJ pagador + cidade destino + valor da NF + peso + cubagem (qualquer data)
      o.chaveInd = [txt(val(r, I.cnpjPagador)), chave(val(r, I.destino)), num(val(r, I.vlNF)).toFixed(2), num(val(r, I.peso)).toFixed(3), num(val(r, I.cubagem)).toFixed(4)].join('|');
      // vendedora da cotação: quem incluiu (USUARIO INCLUSAO "sistema" = vendedor Sistema); vazio: o campo VENDEDOR; senão, Ostel
      var nome = nomeVendedora(val(r, I.usuario));
      if (!nome) { var vend = txt(val(r, I.vendedor)); nome = vend ? nomeVendedora(vend.split(/\s+/)[0]) : ''; }
      if (!nome) nome = CONFIG.vendedoraPadrao;
      o.vendNome = nome; o.vendedora = registrarNome(nome);
      o._q = r.map(function (x) { return x == null ? '' : String(x); }).join(' ').toLowerCase() + ' ' + nome.toLowerCase();
      if (o.cancelada && CONFIG.ocultarCancelados) { CANC_FORA.cot++; return; }  // cancelada: fora do painel
      out.push(o);
    });
    return out;
  }
  /* PRÊMIOS DAS VENDEDORAS (dados\premios.csv): faixas de faturamento do mês.
     A faixa vale quando o faturamento do mês (CT-e não cancelados) chega ao mínimo; faixa marcada
     "exige bater a meta" só vale se a vendedora também bater a própria meta do metas.csv.
     Participam as vendedoras com meta preenchida no metas.csv do mês. */
  var PREMIOS = [];
  function lerPremios(lista) {
    PREMIOS = (lista || []).filter(function (p) { return Array.isArray(p) && p.length >= 3; })
      .map(function (p) { return { faixa: +p[0], min: +p[1] || 0, premio: +p[2] || 0, exigeMeta: !!p[3] }; })
      .sort(function (a, b) { return a.min - b.min; });
  }
  // situação de uma vendedora: faixa atual, prêmio, próxima faixa e quanto falta
  function premioDe(real, meta) {
    var atual = null, prox = null;
    PREMIOS.forEach(function (p) {
      var precisa = Math.max(p.min, p.exigeMeta && meta ? meta : 0);
      var vale = real >= p.min && (!p.exigeMeta || (meta > 0 && real >= meta));
      if (vale) { atual = p; prox = null; } else if (!prox) prox = { p: p, precisa: precisa, falta: Math.max(0, precisa - real) };
    });
    return { atual: atual, premio: atual ? atual.premio : 0, prox: prox };
  }
  var ORD = function (n) { return n + 'ª'; };
  var brlK = function (v) { return 'R$ ' + fNum.format(Math.round(v / 1000)) + ' mil'; };   // R$ 117 mil (rótulo curto das faixas)
  // COMISSÃO (dados\comissao.csv): percentual sobre o faturamento do mês; "Todas" vale para quem não tem linha própria
  var COMISSAO = { todas: 0, por: {} };
  function lerComissao(lista) {
    COMISSAO = { todas: 0, por: {} };
    (lista || []).forEach(function (c) {
      if (!Array.isArray(c) || c.length < 2) return;
      var k = chave(c[0]), p = (+c[1] || 0) / 100;
      if (k === 'TODAS' || k === 'TODOS') COMISSAO.todas = p; else COMISSAO.por[registrarNome(nomeVendedora(c[0]) || c[0])] = p;
    });
  }
  function pctComissao(k) { return Object.prototype.hasOwnProperty.call(COMISSAO.por, k) ? COMISSAO.por[k] : COMISSAO.todas; }
  var TEM_COMISSAO = function () { return COMISSAO.todas > 0 || Object.keys(COMISSAO.por).length > 0; };
  var fPct = function (p) { return fDec.format(p * 100) + '%'; };

  // METAS[ym][vendedora] = valor | LISTA_CARDS[ym] = vendedoras do metas.csv, na ordem do arquivo
  // (meta em branco também vale: a vendedora ganha card, só fica "sem meta")
  var LISTA_CARDS = {};
  function lerMetas(lista) {
    METAS = {}; LISTA_CARDS = {};
    (lista || []).forEach(function (m) {
      if (!Array.isArray(m) || m.length < 3) return;
      var ym = m[0], k = chave(m[1]) === 'TOTAL' ? 'TOTAL' : registrarNome(nomeVendedora(m[1]) || m[1]);
      (METAS[ym] = METAS[ym] || {})[k] = +m[2] || 0;
      if (k !== 'TOTAL') { var l = LISTA_CARDS[ym] = LISTA_CARDS[ym] || []; if (l.indexOf(k) < 0) l.push(k); }
    });
  }
  // meta de um mês: da vendedora (se informada) ou geral (TOTAL; sem TOTAL = soma das vendedoras)
  function metaDe(ym, vend) {
    var M = METAS[ym]; if (!M) return 0;
    if (vend) return M[vend] || 0;
    if (M.TOTAL) return M.TOTAL;
    return sum(Object.keys(M), function (k) { return k === 'TOTAL' ? 0 : M[k]; });
  }

  /* ===================================================================
     FILTROS
     =================================================================== */
  function filtrarCte(F, semCampo) {
    var q = F.q.trim().toLowerCase();
    var semData = !!F.atraso;   // v50: "Entregas em atraso" mostra todos os meses
    return CTE.filter(function (o) {
      if (F.atraso === '1' && !o.atraso) return false;
      if (F.atraso === 'prazo' && !o.noPrazo) return false;
      if (!semData) {
      if (F.de && o.dt < F.de) return false;
      if (F.ate && o.dt > F.ate) return false;
      if (F.ano && o.ano !== F.ano) return false;
      if (F.mes && o.mes !== F.mes) return false;
      if (F.ym && o.ym !== F.ym) return false;
      if (F.dia && o.dt !== F.dia) return false;
      }
      if (F.status && o.baixa !== F.status) return false;
      if (F.vendedora && semCampo !== 'vendedora' && o.vendedora !== F.vendedora) return false;
      if (F.uf && semCampo !== 'uf' && o.uf !== F.uf) return false;
      if (F.cliente && o.cliente !== F.cliente) return false;
      if (F.origem && o.origem !== F.origem) return false;
      if (F.compl && !o.complemento) return false;
      if (F.motivo && o.motivo !== F.motivo) return false;
      if (F.ocorr && o.ocorr !== F.ocorr) return false;
      if (F.codoc && o.codoc !== F.codoc) return false;
      if (q && o._q.indexOf(q) < 0) return false;
      return true;
    });
  }
  function filtrarCot(F, semCampo) {
    var q = F.q.trim().toLowerCase();
    return COT.filter(function (o) {
      if (F.de && o.dt < F.de) return false;
      if (F.ate && o.dt > F.ate) return false;
      if (F.ano && o.ano !== F.ano) return false;
      if (F.mes && o.mes !== F.mes) return false;
      if (F.ym && o.ym !== F.ym) return false;
      if (F.dia && o.dt !== F.dia) return false;
      if (F.vendedora && semCampo !== 'vendedora' && o.vendedora !== F.vendedora) return false;
      if (F.uf && semCampo !== 'uf' && o.uf !== F.uf) return false;
      if (F.cliente && o.cliente !== F.cliente) return false;
      if (F.situacao && o.situacao !== F.situacao) return false;
      if (q && o._q.indexOf(q) < 0) return false;
      return true;
    });
  }
  function aplicarFiltros() { FC = filtrarCte(F); FQ = filtrarCot(F); }
  function copiaF() { var c = {}; Object.keys(F).forEach(function (k) { c[k] = F[k]; }); return c; }

  /* AGRUPAMENTO: soma CT-e e cotações por uma chave (vendedora, UF, cliente, dia...)
     quantidades contam TODOS os documentos; valores só os NÃO cancelados */
  function novoGrupo(k) { return { key: k, qtd: 0, canc: 0, val: 0, fat: 0, fsi: 0, icms: 0, merc: 0, peso: 0, coleta: 0, armazem: 0, cot: 0, cotCanc: 0, cotVal: 0, valor: 0, valorIni: 0 }; }
  function somarCte(g, o) {
    g.qtd++;
    if (o.cancelado) { g.canc++; return; }
    g.val++; g.fat += o.frete; g.fsi += o.freteSI; g.icms += o.icms; g.merc += o.vlMerc; g.peso += o.peso;
    if (o.origem === 'COLETA') g.coleta += o.frete; else if (o.origem === 'ARMAZEM') g.armazem += o.frete;
  }
  function somarCot(g, o) { g.cot++; if (o.cancelada) { g.cotCanc++; return; } g.cotVal++; g.valor += o.valor; g.valorIni += o.valorIni; }
  /* TAXA DE FECHAMENTO = CT-e emitidos NÃO cancelados ÷ cotações NÃO canceladas (mesmo filtro e período)
     VARIAÇÃO DO FRETE  = (frete atual − frete inicial) ÷ frete inicial, das cotações não canceladas
     (frete inicial = PROPOSTA INICIAL | frete atual = PROPOSTA ATUAL) */
  function fechar(g) {
    g.ticket = g.val ? g.fat / g.val : 0;
    g.fechPct = g.cotVal ? g.val / g.cotVal : 0;
    g.varFrete = g.valorIni ? (g.valor - g.valorIni) / g.valorIni : 0;
    return g;
  }
  function agrupar(lc, lq, chC, chQ) {
    var mapa = {}, lista = [];
    var pega = function (k) { var g = mapa[k]; if (!g) { g = mapa[k] = novoGrupo(k); lista.push(g); } return g; };
    lc.forEach(function (o) { var k = typeof chC === 'function' ? chC(o) : o[chC]; if (k !== '' && k != null) somarCte(pega(k), o); });
    if (lq) lq.forEach(function (o) { var k = typeof chQ === 'function' ? chQ(o) : o[chQ]; if (k !== '' && k != null) somarCot(pega(k), o); });
    lista.forEach(fechar);
    return lista;
  }
  function total(lc, lq) { var g = novoGrupo('Total'); lc.forEach(function (o) { somarCte(g, o); }); (lq || []).forEach(function (o) { somarCot(g, o); }); return fechar(g); }

  /* ===================================================================
     FILTROS: preenchimento, sincronização e eventos
     =================================================================== */
  function preencherSelect(id, itens, rotuloTodos) {
    var s = $('#' + id), atual = s.value;
    s.innerHTML = '<option value="">' + rotuloTodos + '</option>' + itens.map(function (v) {
      var vv = typeof v === 'object' ? v.v : v, l = typeof v === 'object' ? v.l : v;
      return '<option value="' + esc(vv) + '">' + esc(l) + '</option>';
    }).join('');
    s.value = atual;
  }
  function montarOpcoesFiltros() {
    var anos = {}; CTE.concat(COT).forEach(function (o) { anos[o.ano] = 1; });
    preencherSelect('f_ano', Object.keys(anos).sort().reverse(), 'Todos');
    preencherSelect('f_mes', MESES_L.map(function (m, i) { return { v: String(i + 1), l: m }; }), 'Todos');
    var st = {}; CTE.forEach(function (o) { st[o.baixa] = (st[o.baixa] || 0) + 1; });
    preencherSelect('f_status', Object.keys(st).sort(function (a, b) { return st[b] - st[a]; }).map(function (k) { return { v: k, l: titulo(k) }; }), 'Todos');
    var vs = agrupar(CTE, COT, 'vendedora', 'vendedora').sort(function (a, b) { return NOMES[a.key].localeCompare(NOMES[b.key], 'pt-BR'); });
    preencherSelect('f_vendedora', vs.map(function (g) { return { v: g.key, l: NOMES[g.key] }; }), 'Todas');
    var datas = CTE.concat(COT).map(function (o) { return o.dt; }).sort();
    if (datas.length) { $('#f_de').min = $('#f_ate').min = datas[0]; $('#f_de').max = $('#f_ate').max = datas[datas.length - 1]; }
  }
  function rotuloFiltro(k, v) {
    if (k === 'mes') return MESES_L[+v - 1];
    if (k === 'de' || k === 'ate' || k === 'dia') return fDate(v);
    if (k === 'ym') return nomeMes(v);
    if (k === 'vendedora') return NOMES[v] || v;
    if (k === 'origem') return ORIGENS[v] || v;
    if (k === 'ocorr') return nomeOcorr(v);
    if (k === 'codoc') return nomeCodOc(v);
    if (k === 'compl') return 'sem saída e frete acima de ' + brl(CONFIG.complementoAcimaDe);
    if (k === 'atraso') return v === 'prazo' ? 'entregas no prazo, ainda não realizadas (todos os meses)' : 'prazo vencido, sem entrega (todos os meses)';
    if (k === 'status' || k === 'motivo') return titulo(v);
    return v;
  }
  function sincronizarFiltros() {
    ['de','ate','ano','mes','status','vendedora','q'].forEach(function (k) {
      var e = $('#f_' + k); if (e && e.value !== F[k]) e.value = F[k];
      var fld = $('.fld[data-f="' + k + '"]'); if (fld) fld.classList.toggle('on', !!F[k] && filtrosLigados().indexOf(k) >= 0);
    });
    var ativos = filtrosLigados();
    $('#chips').innerHTML = ativos.map(function (k) {
      var nota = SO_CTE[k] ? ' title="Vale só para os CT-e (as cotações não mudam)"' : (SO_COT[k] ? ' title="Vale só para as cotações (os CT-e não mudam)"' : '');
      return '<span class="chip"' + nota + '><span><b>' + F_ROTULO[k] + ':</b> ' + esc(rotuloFiltro(k, F[k])) + '</span><button data-rm="' + k + '" aria-label="Remover filtro ' + F_ROTULO[k] + '">×</button></span>';
    }).join('');
    $$('[data-clear]').forEach(function (b) { b.disabled = !ativos.length; });
    $('#fCount').textContent = ativos.length ? '(' + ativos.length + ')' : '';
  }
  function definirFiltro(k, v, alternar) {
    if (alternar && F[k] === v) v = '';
    F[k] = v; UI.pagina = 1; atualizar();
  }
  // PERÍODO PADRÃO: o painel abre no mês atual (ou ano atual, conforme CONFIG.periodoInicial).
  // Se ainda não houver dados no mês/ano atual, usa o mais recente que existir.
  // O período padrão é o estado "normal": não gera etiqueta nem faz o botão piscar.
  var PADRAO = { ano: '', mes: '' }, PADRAO_APLICADO = false;
  function calcularPadrao() {
    if (CONFIG.periodoInicial === 'tudo') return { ano: '', mes: '' };
    var yms = {}; CTE.concat(COT).forEach(function (o) { yms[o.ym] = 1; });
    var lista = Object.keys(yms).sort(), h = new Date();
    var atual = h.getFullYear() + '-' + ('0' + (h.getMonth() + 1)).slice(-2);
    if (CONFIG.periodoInicial === 'ano') {
      var anos = lista.map(function (y) { return y.slice(0, 4); }), a = String(h.getFullYear());
      return { ano: anos.indexOf(a) >= 0 ? a : (anos[anos.length - 1] || ''), mes: '' };
    }
    var ym = lista.indexOf(atual) >= 0 ? atual : (lista[lista.length - 1] || atual);
    return { ano: ym.slice(0, 4), mes: String(+ym.slice(5, 7)) };
  }
  function filtrosLigados() {
    return Object.keys(F).filter(function (k) {
      if (!F[k]) return false;
      if (k === 'ano' && F.ano === PADRAO.ano && F.mes === PADRAO.mes) return false;
      if (k === 'mes' && F.mes === PADRAO.mes && F.ano === PADRAO.ano) return false;
      return true;
    });
  }
  function limparFiltros() {
    Object.keys(F).forEach(function (k) { F[k] = ''; });
    F.ano = PADRAO.ano; F.mes = PADRAO.mes;
    UI.pagina = 1; UI.busca = ''; $('#tblSearch').value = '';
    atualizar();
    aviso(PADRAO.mes ? 'Filtros limpos · ' + MESES_L[+PADRAO.mes - 1] + ' de ' + PADRAO.ano : (PADRAO.ano ? 'Filtros limpos · ano ' + PADRAO.ano : 'Filtros limpos'));
  }
  ['de','ate','ano','mes','status','vendedora'].forEach(function (k) {
    $('#f_' + k).addEventListener('change', function (e) {
      // escolher datas desliga o mês/ano padrão (senão o período ficaria "preso" no mês)
      if ((k === 'de' || k === 'ate') && e.target.value) { F.ano = ''; F.mes = ''; F.ym = ''; }
      definirFiltro(k, e.target.value);
    });
  });
  $('#f_q').addEventListener('input', debounce(function (e) { definirFiltro('q', e.target.value); }, 250));
  $$('[data-clear]').forEach(function (b) { b.addEventListener('click', limparFiltros); });
  $('#chips').addEventListener('click', function (e) { var b = e.target.closest('[data-rm]'); if (b) definirFiltro(b.getAttribute('data-rm'), ''); });
  $('#btnToggleF').addEventListener('click', function () {
    var f = $('#filters'); f.classList.toggle('open'); this.setAttribute('aria-expanded', f.classList.contains('open'));
  });
  function ligarSeg(id, chaveUI, depois) {
    $('#' + id).addEventListener('click', function (e) {
      var b = e.target.closest('button'); if (!b) return;
      UI[chaveUI] = b.getAttribute('data-v');
      $$('button', this).forEach(function (x) { x.classList.toggle('on', x === b); });
      (depois || desenharGraficos)();
    });
  }
  ligarSeg('segEvo', 'evo'); ligarSeg('segEvoMed', 'evoMed'); ligarSeg('segUF', 'uf'); ligarSeg('segVend', 'vend'); ligarSeg('segCli', 'cli'); ligarSeg('segBaixa', 'baixa');
  ligarSeg('segTab', 'aba', function () { UI.pagina = 1; UI.ordem = -1; UI.soDup = false; UI.sel = null; UI.busca = ''; $('#tblSearch').value = ''; desenharTabela(); salvarNaURL(); $('.tbl-wrap').scrollTop = 0; });

  /* ===================================================================
     MÊS DE REFERÊNCIA DA META e DIAS ÚTEIS
     =================================================================== */
  function mesReferencia() {
    if (F.ym) return F.ym;
    if (F.ano && F.mes) return F.ano + '-' + ('0' + F.mes).slice(-2);
    if (F.dia) return F.dia.slice(0, 7);
    if (F.de && F.ate && F.de.slice(0, 7) === F.ate.slice(0, 7)) return F.de.slice(0, 7);
    var atual = hojeIso().slice(0, 7);
    var yms = FC.concat(FQ).map(function (o) { return o.ym; }).sort();
    if (!yms.length || yms.indexOf(atual) >= 0) return atual;
    return yms[yms.length - 1];
  }
  // FERIADOS (dados\Feriados.csv, só os com Folga = S): "aaaa-mm-dd" -> nome
  var FERIADOS = {};
  function lerFeriados(lista) { FERIADOS = {}; (lista || []).forEach(function (f) { if (Array.isArray(f) && f[0]) FERIADOS[String(f[0]).slice(0, 10)] = f[1] || 'Feriado'; }); }
  function ehDiaUtil(d) { var w = d.getDay(); return w !== 0 && w !== 6 && !FERIADOS[iso(d)]; }   // segunda a sexta, menos feriados
  // dias úteis do mês: total, já passados (até hoje, inclusive) e lista de cada dia
  function diasUteis(ym) {
    var y = +ym.slice(0, 4), m = +ym.slice(5, 7) - 1, fim = new Date(y, m + 1, 0).getDate(), hoje = hojeIso();
    var tot = 0, pass = 0, dias = [], fer = [];
    for (var d = 1; d <= fim; d++) {
      var dt = new Date(y, m, d), s = iso(dt), u = ehDiaUtil(dt);
      if (u) { tot++; if (s <= hoje) pass++; }
      dias.push({ iso: s, util: u });
      if (FERIADOS[s] && dt.getDay() !== 0 && dt.getDay() !== 6) fer.push(s.slice(8, 10) + '/' + s.slice(5, 7) + ' ' + FERIADOS[s]);
    }
    return { total: tot, passados: pass, dias: dias, ultimo: fim, feriados: fer };
  }
  // CT-e do mês de referência com os demais filtros (meta é mensal: ignora dia e intervalo de datas)
  function cteDoMes(ym, semVendedora) {
    var F2 = copiaF(); F2.de = ''; F2.ate = ''; F2.dia = ''; F2.ano = ''; F2.mes = ''; F2.ym = ym;
    return filtrarCte(F2, semVendedora ? 'vendedora' : '');
  }
  function situacaoRitmo(real, meta, du) {
    if (!meta) return { c: 'neutro', t: 'Meta não definida' };
    if (real >= meta) return { c: 'bom', t: 'Meta batida' };
    var esperado = du.total ? meta * du.passados / du.total : 0;
    if (!esperado) return { c: 'neutro', t: 'Mês não começou' };
    var r = real / esperado;
    if (r >= 1) return { c: 'bom', t: 'No ritmo da meta' };
    if (r >= 0.85) return { c: 'atencao', t: 'Pouco abaixo do ritmo' };
    return { c: 'ruim', t: 'Abaixo do ritmo' };
  }

  /* ===================================================================
     META DO MÊS (painel em destaque)
     =================================================================== */
  /* FRASE DE INCENTIVO DO DIA: muda todo dia e acompanha o desempenho da meta
     (meta batida, no ritmo, pouco abaixo, abaixo do ritmo). Para trocar as frases, edite as listas. */
  var FRASES = {
    batida: [
      'Meta batida! Parabéns, time. Agora é buscar o recorde do mês.',
      'Objetivo alcançado! Cada venda daqui para frente é bônus.',
      'Meta conquistada com garra. Vamos fechar o mês com chave de ouro!',
      'Missão cumprida! Bora mostrar até onde esse time consegue chegar.',
      'Resultado de quem trabalha junto. Parabéns pela meta batida!'
    ],
    bom: [
      'No ritmo certo! Mantenha o foco que a meta vem.',
      'Desempenho acima do esperado. Continue assim!',
      'Ótimo ritmo! Cada cotação fechada aproxima o time da meta.',
      'Consistência é o que bate meta. Seguimos firmes!',
      'Excelente trabalho! O resultado de hoje constrói o mês.',
      'Time no compasso da meta. Vamos manter a pegada!',
      'Está dando certo: mesma energia até o último dia útil!'
    ],
    atencao: [
      'Estamos perto do ritmo. Um empurrão hoje faz toda a diferença!',
      'Quase lá! Um contato a mais com cada cliente muda o mês.',
      'Hoje é um ótimo dia para recuperar o ritmo: revise as cotações em aberto.',
      'Pouco abaixo do ritmo, muito acima da vontade. Bora virar!',
      'Uma boa venda hoje coloca o time de volta no ritmo da meta.',
      'Foco nas cotações mais quentes: o ritmo está ao alcance!'
    ],
    ruim: [
      'Ainda dá tempo! Foco nas cotações em aberto e nos clientes ativos.',
      'Cada dia útil é uma nova chance. Vamos virar esse jogo!',
      'Desafio grande, time maior ainda. Bora pra cima!',
      'Retome contatos e revise propostas: a meta ainda é possível.',
      'Um passo de cada vez: comece hoje pelas cotações mais quentes.',
      'Grandes viradas começam com o primeiro fechamento do dia. Vamos!',
      'Meta se constrói todo dia. Hoje é dia de acelerar!'
    ],
    neutro: [
      'Defina a meta do mês no metas.csv para acompanhar o ritmo do time.'
    ]
  };
  function fraseDoDia(chaveSit, quem) {
    var lista = FRASES[chaveSit] || FRASES.neutro, h = new Date();
    var diaDoAno = Math.floor((h - new Date(h.getFullYear(), 0, 0)) / 864e5);   // muda a cada dia
    var f = lista[diaDoAno % lista.length];
    return quem ? quem + ', ' + f.charAt(0).toLowerCase() + f.slice(1) : f;
  }

  function desenharMeta() {
    var ym = mesReferencia(), du = diasUteis(ym), lista = cteDoMes(ym), g = total(lista, null);
    var meta = metaDe(ym, F.vendedora), real = g.fat, p = meta ? real / meta : 0;
    var esperadoFrac = du.total ? du.passados / du.total : 0;
    var sit = situacaoRitmo(real, meta, du);
    var falta = Math.max(0, meta - real), restantes = du.total - du.passados;
    var proj = du.passados ? real / du.passados * du.total : real;
    var quem = F.vendedora ? ' · ' + esc(NOMES[F.vendedora]) : '';
    var arq = esc((RAW && RAW.metasArquivo) || 'metas.csv');
    var corBarra = sit.c === 'ruim' ? 'ruim' : (sit.c === 'atencao' ? 'atencao' : '');
    var h = '<div class="mm-prin">' +
      '<div class="mm-tit"><span class="c-ic" aria-hidden="true" style="color:#fff">' + IC.alvo + '</span><h2>Meta do mês · ' + nomeMesLongo(ym) + quem + '</h2>' +
      '<span class="mm-situ ' + sit.c + '">' + sit.t + '</span></div>';
    // PERÍODO FILTRADO (datas De/Até ou dia): a meta é do mês, mas mostra também o resultado do período
    // escolhido contra a meta proporcional aos dias úteis do período
    if (F.de || F.ate || F.dia) {
      var ini = F.dia || F.de || (ym + '-01'), fim = F.dia || F.ate || du.dias[du.dias.length - 1].iso;
      var uteisP = du.dias.filter(function (d) { return d.util && d.iso >= ini && d.iso <= fim; }).length;
      var realP = total(FC).fat, metaP = meta && du.total ? meta * uteisP / du.total : 0;
      h += '<div class="mm-periodo">' + IC.cal + '<span>Período <b>' + esc(fDate(ini)) + (fim !== ini ? ' a ' + esc(fDate(fim)) : '') + '</b>: faturado <b class="num">' + brl(realP) + '</b>' +
        (metaP ? ' · meta proporcional (' + uteisP + ' de ' + du.total + ' dias úteis) <b class="num">' + brl(metaP) + '</b> · <b class="num ' + (realP >= metaP ? 'ok' : 'baixo') + '">' + fDec.format(realP / metaP * 100) + '%</b>' : '') +
        '<small>Os números grandes abaixo continuam sendo do mês inteiro (a meta é mensal).</small></span></div>';
    }
    if (meta) {
      h += '<div class="mm-num"><div class="mm-pct num">' + fDec.format(p * 100) + '<small>%</small></div>' +
           '<div class="mm-de">' + '<b>' + brl(real) + '</b> de <b>' + brl(meta) + '</b></div></div>' +
           '<div class="mm-barra" title="A marca preta é onde o faturamento deveria estar hoje (dias úteis já passados)"><i class="' + corBarra + '" style="width:' + Math.min(100, p * 100).toFixed(1) + '%"></i>' +
           (esperadoFrac > 0 && esperadoFrac < 1 ? '<em style="left:' + (esperadoFrac * 100).toFixed(1) + '%"></em>' : '') + '</div>' +
           '<div class="mm-escala"><span>0%</span><span>Esperado hoje: ' + pct(esperadoFrac) + ' (' + du.passados + ' de ' + du.total + ' dias úteis)</span><span>100%</span></div>' +
           (du.feriados.length ? '<div class="mm-feriado">' + (FERIADOS[hojeIso()] ? 'Hoje é feriado (<b>' + esc(FERIADOS[hojeIso()]) + '</b>) e não conta como dia útil. ' : '') +
             'Feriados do mês, fora dos dias úteis: ' + esc(du.feriados.join(' · ')) + '</div>' : '') +
           // QUANTO FALTA para bater a meta (e quanto por dia útil)
           (falta > 0
             ? '<div class="mm-falta">Falta <b class="num">' + brl(falta) + '</b> para bater a meta' +
               (restantes > 0 ? ' · <b class="num">' + brl(falta / restantes) + '</b> por dia útil nos ' + restantes + ' dias úteis restantes' : '') + '</div>'
             : '<div class="mm-falta ok">Meta batida! <b class="num">' + brl(real - meta) + '</b> acima da meta</div>');
    } else {
      h += '<div class="mm-num"><div class="mm-pct num" style="font-size:40px">' + brl(real) + '</div><div class="mm-de">faturado no mês</div></div>' +
           '<div class="mm-sem">A meta de ' + nomeMesLongo(ym) + (F.vendedora ? ' para ' + esc(NOMES[F.vendedora]) : '') + ' ainda não foi definida. Preencha o arquivo <code>dados\\' + arq + '</code> ' +
           '(colunas Mês; Vendedora; Meta - ex.: <code>' + ym.slice(5) + '/' + ym.slice(0, 4) + ';TOTAL;250000,00</code>) e salve. O painel atualiza sozinho.</div>';
    }
    // FRASE DE INCENTIVO DO DIA (acompanha a situação da meta)
    var chaveFrase = !meta ? 'neutro' : (real >= meta ? 'batida' : (sit.c === 'neutro' ? 'bom' : sit.c));
    h += '<p class="mm-frase ' + chaveFrase + '"><svg viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M7.2 6C4.9 7.3 3.5 9.6 3.5 12.6V18h6v-6H6.6c.1-2 1.1-3.6 3-4.6zM17.7 6c-2.3 1.3-3.7 3.6-3.7 6.6V18h6v-6h-2.9c.1-2 1.1-3.6 3-4.6z"/></svg>' +
         esc(fraseDoDia(chaveFrase, F.vendedora ? NOMES[F.vendedora] : '')) + '</p>';
    // FATURADO HOJE (ou no último dia com CT-e, quando o mês escolhido não é o atual) e no dia útil anterior
    var hj = hojeIso(), diaRef = hj.slice(0, 7) === ym ? hj : (lista.map(function (o) { return o.dt; }).sort().pop() || ym + '-01');
    var doDia = function (d) { return total(lista.filter(function (o) { return o.dt === d; })); };
    var gHoje = doDia(diaRef), antes = du.dias.filter(function (d) { return d.util && d.iso < diaRef; }).map(function (d) { return d.iso; }).pop();
    var gAnt = antes ? doDia(antes) : null;
    var precisaHoje = meta ? Math.max(0, meta - (real - gHoje.fat)) / Math.max(1, du.dias.filter(function (d) { return d.util && d.iso >= diaRef; }).length) : 0;
    var rotHoje = diaRef === hj ? 'Faturado hoje (' + fDate(hj).slice(0, 5) + ')' : 'Faturado em ' + fDate(diaRef).slice(0, 5) + ' (último dia)';
    h += '</div><div class="mm-cels">' +
      '<div class="mm-hoje"><span>' + rotHoje + '</span><b class="num">' + brl(gHoje.fat) + '</b><small>' + nInt(gHoje.val) + ' CT-e válidos' +
        (precisaHoje ? ' · <b class="mm-pc ' + (gHoje.fat >= precisaHoje ? 'ok' : '') + '">' + pct(gHoje.fat / precisaHoje) + '</b> do necessário no dia (' + brl(precisaHoje) + ')' : '') + '</small></div>' +
      '<div><span>Dia útil anterior' + (antes ? ' (' + fDate(antes).slice(0, 5) + ')' : '') + '</span><b class="num">' + (gAnt ? brl(gAnt.fat) : '—') + '</b><small>' + (gAnt ? nInt(gAnt.val) + ' CT-e válidos' : 'sem dia útil anterior no mês') + '</small></div>' +
      '<div><span>Faturado no mês</span><b class="num">' + brl(real) + '</b><small>' + nInt(g.val) + ' CT-e</small></div>' +
      '<div><span>Falta para a meta</span><b class="num">' + (meta ? brl(falta) : '—') + '</b><small>' + (meta ? (falta ? 'faltam ' + restantes + ' dias úteis' : 'meta alcançada') : 'sem meta') + '</small></div>' +
      '<div><span>Precisa vender por dia útil</span><b class="num">' + (meta && falta && restantes ? brl(falta / restantes) : '—') + '</b><small>' + (du.passados ? 'média até agora: ' + brl(real / du.passados) : 'mês ainda não começou') + '</small></div>' +
      '<div><span>Projeção para o fim do mês</span><b class="num">' + brl(proj) + '</b><small>' + (meta ? pct(meta ? proj / meta : 0) + ' da meta, no ritmo atual' : 'no ritmo atual') + '</small></div>' +
      '</div>';
    $('#metaMes').innerHTML = h;
  }

  /* ===================================================================
     INDICADORES (KPIs) — até 8, num painel só
     =================================================================== */
  var KPIS = [
    { id: 'cte',     ic: 'doc',     c: '--s1', l: 'CT-e emitidos',          fmt: nInt, m: function (C) { return C.length; } },
    { id: 'fat',     ic: 'money',   c: '--s1', l: 'Faturamento do mês',      fmt: brl,  m: function (C) { return total(C).fat; } },
    { id: 'cot',     ic: 'receipt', c: '--s2', l: 'Cotações',                fmt: nInt, q: true, m: function (C, Q) { return Q.length; } },
    { id: 'fech',    ic: 'check',   c: '--s2', l: 'Taxa de fechamento',      fmt: pct,  q: true, pp: true, min: 3, m: function (C, Q) { return total(C, Q).fechPct; } },
    { id: 'coleta',  ic: 'truck',   c: '--s3', l: 'Faturamento coletado',    fmt: brl,  m: function (C) { return total(C).coleta; } },
    { id: 'armazem', ic: 'box',     c: '--s3', l: 'Entregue no armazém',     fmt: brl,  m: function (C) { return total(C).armazem; } },
    { id: 'ticket',  ic: 'ticket',  c: '--s5', l: 'Ticket médio por CT-e',   fmt: brl,  min: 2, m: function (C) { return total(C).ticket; } },
    { id: 'valor',   ic: 'tag',     c: '--s2', l: 'Valor cotado',            fmt: brl,  q: true, m: function (C, Q) { return total([], Q).valor; } }
  ];
  var KPI_ULTIMO = {};

  // período anterior equivalente (para a variação dos indicadores)
  function periodoAnterior() {
    var ini, fim, rot, h = new Date();
    var mesAtual = function (y, m) { return y === h.getFullYear() && m === h.getMonth(); };
    var mesmoPeriodoMesAnt = function (y, m) {   // mês em andamento: 1º até o mesmo dia do mês anterior
      ini = new Date(y, m - 1, 1); var ult = new Date(y, m, 0).getDate();
      fim = new Date(y, m - 1, Math.min(h.getDate(), ult)); rot = 'mesmo período de ' + MESES[(m + 11) % 12] + '/' + (m === 0 ? y - 1 : y);
    };
    if (F.dia) { var d = dataDe(F.dia); ini = fim = new Date(d.getTime() - 864e5); rot = 'dia anterior'; }
    else if (F.ym || (F.ano && F.mes)) {
      var y = F.ym ? +F.ym.slice(0, 4) : +F.ano, m = (F.ym ? +F.ym.slice(5) : +F.mes) - 1;
      if (mesAtual(y, m)) mesmoPeriodoMesAnt(y, m);
      else { ini = new Date(y, m - 1, 1); fim = new Date(y, m, 0); rot = 'mês anterior'; }
    }
    else if (F.ano && !F.mes) {
      if (+F.ano === h.getFullYear()) { ini = new Date(+F.ano - 1, 0, 1); fim = new Date(+F.ano - 1, h.getMonth(), h.getDate()); rot = 'mesmo período de ' + (+F.ano - 1); }
      else { ini = new Date(+F.ano - 1, 0, 1); fim = new Date(+F.ano - 1, 11, 31); rot = 'ano anterior'; }
    }
    else if (F.de && F.ate) {
      var d0 = dataDe(F.de), d1 = dataDe(F.ate), dias = Math.round((d1 - d0) / 864e5) + 1;
      fim = new Date(d0.getTime() - 864e5); ini = new Date(fim.getTime() - (dias - 1) * 864e5); rot = 'período anterior';
    } else return null;
    var F2 = copiaF(); F2.ano = ''; F2.mes = ''; F2.ym = ''; F2.dia = ''; F2.de = iso(ini); F2.ate = iso(fim);
    return { c: filtrarCte(F2), q: filtrarCot(F2), rotulo: rot };
  }
  function contar(el, de, para, fmt) {
    if (!isFinite(de) || de === para || matchMedia('(prefers-reduced-motion: reduce)').matches) { el.textContent = fmt(para); return; }
    var t0 = performance.now(), dur = 700;
    (function passo(t) {
      var k = Math.min(1, (t - t0) / dur), e = 1 - Math.pow(1 - k, 3);
      el.textContent = fmt(de + (para - de) * e);
      if (k < 1) requestAnimationFrame(passo);
    })(t0);
  }
  function desenharKPIs() {
    var C = FC, Q = FQ, ant = periodoAnterior(), g = total(C, Q);
    var ehMes = !!(F.ym || (F.ano && F.mes));
    KPIS[1].l = ehMes ? 'Faturamento do mês' : 'Faturamento do período';
    var extra = {
      cte:     { s: 'CT-e emitidos, sem os cancelados (' + nInt(CANC_FORA.cte) + ' cancelados fora do painel)' },
      fat:     { s: 'Sem ICMS: <b>' + brl(g.fsi) + '</b> · ICMS: <b>' + brl(g.icms) + '</b>' },
      cot:     { s: 'Cotações, sem as canceladas (' + nInt(CANC_FORA.cot) + ' canceladas fora do painel)' },
      fech:    { s: '<b>' + nInt(g.val) + '</b> CT-e válidos ÷ <b>' + nInt(g.cotVal) + '</b> cotações válidas<br>Frete inicial <b>' + brl(g.valorIni) + '</b> · frete atual <b>' + brl(g.valor) + '</b> · ' +
                        '<b class="' + (g.varFrete < 0 ? 'neg' : 'pos') + '">' + (g.varFrete > 0 ? '+' : '') + fDec.format(g.varFrete * 100) + '%</b>', bar: Math.min(1, g.fechPct) },
      coleta:  { s: '<b>' + pct(g.fat ? g.coleta / g.fat : 0) + '</b> do faturamento · coluna E com placa', bar: g.fat ? g.coleta / g.fat : 0 },
      armazem: { s: '<b>' + pct(g.fat ? g.armazem / g.fat : 0) + '</b> do faturamento · coluna E = ARMAZEM', bar: g.fat ? g.armazem / g.fat : 0 },
      ticket:  { s: 'Média de ' + nInt(g.val) + ' CT-e válidos · mercadoria <b>' + brl(g.merc) + '</b>' },
      valor:   { s: 'Proposta atual das cotações válidas · média <b>' + brl(g.cotVal ? g.valor / g.cotVal : 0) + '</b>' }
    };
    $('#kpis').innerHTML = KPIS.map(function (k) {
      var v = k.m(C, Q), x = extra[k.id], delta = '';
      if (ant) {
        var va = k.m(ant.c, ant.q), base = k.q ? ant.q.length : ant.c.length, dif, txtd, bom;
        if (!base) delta = '<div class="k-delta neutro">sem dados no ' + ant.rotulo + '</div>';
        else {
          if (k.pp) { dif = v - va; txtd = fDec.format(Math.abs(dif) * 100) + ' p.p.'; }
          else { dif = va ? (v - va) / Math.abs(va) : 0; txtd = fDec.format(Math.abs(dif) * 100) + '%'; }
          bom = dif >= 0;
          if (!k.pp && (!va || Math.abs(dif) > 9.99)) { delta = '<div class="k-delta neutro" title="' + esc(ant.rotulo + ': ' + k.fmt(va)) + '">' + ant.rotulo + ': ' + k.fmt(va) + ' (pouca base para comparar)</div>'; }
          else
          delta = '<div class="k-delta ' + (Math.abs(dif) < 1e-9 ? 'neutro' : (bom ? 'up' : 'down')) + '" title="' + esc(ant.rotulo + ': ' + k.fmt(va)) + '">' +
                  '<b>' + (dif >= 0 ? '▲ ' : '▼ ') + txtd + '</b><span>' + (bom ? 'acima' : 'abaixo') + ' do ' + ant.rotulo + '</span></div>';
        }
      }
      return '<button type="button" class="kpi" data-kpi="' + k.id + '" style="--k:var(' + k.c + ')" title="Clique para detalhar">' +
        '<div class="k-top"><span class="k-ic">' + IC[k.ic] + '</span>' + k.l +
        '<span class="k-go" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><path d="M7 17L17 7M9 7h8v8"/></svg></span></div>' +
        '<div class="k-val num" data-v="' + v + '"></div>' + delta + '<div class="k-sub">' + x.s + '</div>' +
        (x.bar != null ? '<div class="bar-mini"><i style="width:' + (Math.max(0, Math.min(1, x.bar)) * 100).toFixed(1) + '%"></i></div>' : '') + '</button>';
    }).join('');
    $$('#kpis .kpi').forEach(function (el) {
      var k = KPIS.filter(function (x) { return x.id === el.getAttribute('data-kpi'); })[0];
      var v = +el.querySelector('.k-val').getAttribute('data-v');
      contar(el.querySelector('.k-val'), KPI_ULTIMO[k.id] == null ? 0 : KPI_ULTIMO[k.id], v, k.fmt);
      el.querySelector('.k-val').title = k.fmt(v);
      KPI_ULTIMO[k.id] = v;
    });
  }
  $('#kpis').addEventListener('click', function (e) { var b = e.target.closest('.kpi'); if (b) abrirDetalheKPI(b.getAttribute('data-kpi')); });

  /* ===================================================================
     CARDS DAS VENDEDORAS
     =================================================================== */
  function desenharVendedoras() {
    CACHE_PR = {}; CACHE_SI = {};
    var ym = mesReferencia(), du = diasUteis(ym);
    // todos os filtros, menos o da vendedora (para continuar vendo as outras, mais apagadas)
    var C = filtrarCte(F, 'vendedora'), Q = filtrarCot(F, 'vendedora'), mesC = cteDoMes(ym, true);
    var grupos = {};
    agrupar(C, Q, 'vendedora', 'vendedora').forEach(function (g) { grupos[g.key] = g; });
    var M = METAS[ym] || {}, lista = LISTA_CARDS[ym] || [], ostel = chave(CONFIG.vendedoraPadrao), ordem;
    var noMes = {}; agrupar(mesC, null, 'vendedora').forEach(function (g) { noMes[g.key] = g.fat; });
    if (lista.length) {
      // QUAIS CARDS: as vendedoras do metas.csv do mês (mesmo sem vendas ou sem meta),
      // mais a Ostel (CT-e sem vendedora) quando tiver dados. Quem tem nome mas não está na lista fica sem
      // card, mas continua nos indicadores, gráficos e tabela (nenhum total muda).
      ordem = lista.slice();
      if (ordem.indexOf(ostel) < 0 && grupos[ostel]) ordem.push(ostel);
      ordem.forEach(function (k) { if (!grupos[k]) grupos[k] = fechar(novoGrupo(k)); });
    } else {
      // mês sem vendedoras no metas.csv: mostra todas as que têm dados
      ordem = Object.keys(grupos);
    }
    // REGRA: sempre maior faturamento primeiro (empate: mais cotações; depois a ordem do metas.csv)
    var pos = ordem.slice();
    ordem.sort(function (a, b) { return grupos[b].fat - grupos[a].fat || grupos[b].cot - grupos[a].cot || pos.indexOf(a) - pos.indexOf(b); });
    var ehMes = !!(F.ym || (F.ano && F.mes)) && !F.de && !F.ate && !F.dia;
    $('#vendInfo').innerHTML = 'Cotações e vendas ' + (ehMes ? 'do mês' : 'do período filtrado') + '; meta de <b>' + nomeMes(ym) + '</b>';
    if (!ordem.length) { $('#vendGrid').innerHTML = '<div class="vend-vazio">Nenhuma vendedora com dados nos filtros escolhidos.</div>'; return; }
    var esperadoFrac = du.total ? du.passados / du.total : 0;
    $('#vendGrid').innerHTML = ordem.map(function (k) {
      var g = grupos[k], meta = M[k] || 0, real = noMes[k] || 0, p = meta ? real / meta : 0, sit = situacaoRitmo(real, meta, du);
      var cls = F.vendedora ? (F.vendedora === k ? ' on' : ' dim') : '';
      var nome = NOMES[k] || k;
      return '<button type="button" class="vc' + cls + '" data-v="' + esc(k) + '" title="Clique para filtrar o painel por ' + esc(nome) + '">' +
        '<div class="vc-top"><span class="vc-av" aria-hidden="true">' + esc(nome.charAt(0)) + '</span>' +
          '<div class="vc-nome">' + esc(nome) + '<small>' + (meta ? sit.t : 'meta não definida') + '</small></div>' +
          '<div class="vc-pct num' + (meta ? '' : ' sem') + '">' + (meta ? fDec.format(p * 100) + '%' : 'sem meta') + '</div></div>' +
        (meta ? '<div class="mm-barra"><i class="' + (sit.c === 'ruim' ? 'ruim' : sit.c === 'atencao' ? 'atencao' : '') + '" style="width:' + Math.min(100, p * 100).toFixed(1) + '%;background:' + (sit.c === 'bom' ? 'var(--s4)' : '') + '"></i>' +
                (esperadoFrac > 0 && esperadoFrac < 1 ? '<em style="left:' + (esperadoFrac * 100).toFixed(1) + '%"></em>' : '') + '</div>' : '') +
        '<div class="vc-meta"><span>No mês: <b>' + brl(real) + '</b></span><span>Meta: <b>' + (meta ? brl(meta) : '—') + '</b></span></div>' +
        (meta ? '<div class="vc-falta' + (real >= meta ? ' ok' : '') + '">' + (real >= meta ? 'Meta batida! +' + brl(real - meta) : 'Falta <b>' + brl(meta - real) + '</b> para a meta') + '</div>' : '') +
        (meta && TEM_COMISSAO() ? '<div class="vc-premio vc-comis"><svg class="ic-trofeu ic-com" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="9"/><path d="M15 9.5c-.5-1-1.6-1.5-3-1.5-1.7 0-3 .9-3 2.2 0 3 6 1.6 6 4.6 0 1.3-1.3 2.2-3 2.2-1.5 0-2.6-.6-3.1-1.6M12 6.5v11"/></svg><span>Comissão (' + fPct(pctComissao(k)) + '): <b>' + brl(comissaoBase(ym, k) * pctComissao(k)) + '</b> <small class="vc-base">sobre ' + brl(comissaoBase(ym, k)) + ' sem ICMS</small></span></div>' : '') +
        (meta && PREMIOS.length ? (function () { var pr = premioDe(premioReal(ym, k), meta);
          return '<div class="vc-premio' + (pr.atual ? ' ganhou' : '') + '">' + IC_TROFEU + (pr.atual ? '<span>Prêmio: <b>' + brl(pr.premio) + '</b> (' + ORD(pr.atual.faixa) + ' faixa)</span>' : '<span>Sem prêmio ainda</span>') +
            (pr.prox ? '<small>faltam ' + brl(pr.prox.falta) + ' sem ICMS para a ' + ORD(pr.prox.p.faixa) + ' (' + brl(pr.prox.p.premio) + ')</small>' : '<small>faixa máxima!</small>') + '</div>'; })() : '') +
        '<div class="vc-cels">' +
          '<div><span>' + (F.de || F.ate || F.dia ? 'Faturamento no período' : 'Faturamento') + '</span><b class="num">' + brl(g.fat) + '</b></div>' +
          '<div><span>CT-e</span><b>' + nInt(g.qtd) + '</b></div>' +
          '<div><span>Cotações</span><b>' + nInt(g.cot) + '</b></div>' +
          '<div><span>Frete inicial</span><b class="num">' + brl(g.valorIni) + '</b></div>' +
          '<div><span>Frete atual</span><b class="num">' + brl(g.valor) + '</b></div>' +
          '<div><span>Variação do frete</span><b>' + (g.varFrete > 0 ? '+' : '') + fDec.format(g.varFrete * 100) + '%</b></div>' +
          '<div><span>Fechamento</span><b>' + pct(g.fechPct) + '</b></div>' +
          '<div><span>Ticket médio</span><b>' + brl(g.ticket) + '</b></div>' +
        '</div></button>';
    }).join('');
  }
  $('#vendGrid').addEventListener('click', function (e) { var b = e.target.closest('.vc'); if (b) definirFiltro('vendedora', b.getAttribute('data-v'), true); });

  /* ===================================================================
     PRÊMIOS DO MÊS (faixas do premios.csv x faturamento do mês de cada vendedora)
     =================================================================== */
  var IC_TROFEU = '<svg class="ic-trofeu" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M8 21h8M12 17v4M7 4h10v5a5 5 0 0 1-10 0z"/><path d="M17 5h3v2a3 3 0 0 1-3 3M7 5H4v2a3 3 0 0 0 3 3"/></svg>';
  // base do PRÊMIO: Valor do Frete SEM ICMS do MÊS INTEIRO da vendedora (CT-e não cancelados), sem os outros filtros
  var CACHE_PR = {};
  function premioReal(ym, k) {
    var c = CACHE_PR[ym];
    if (!c) { c = CACHE_PR[ym] = {}; CTE.forEach(function (o) { if (o.ym === ym && o.soma) c[o.vendedora] = (c[o.vendedora] || 0) + o.freteSI; }); }   // PRÊMIO: pago sobre o frete SEM ICMS
    return c[k] || 0;
  }
  // BASE DA COMISSÃO: Valor do Frete SEM ICMS do mês inteiro (CT-e não cancelados)
  var CACHE_SI = {};
  function comissaoBase(ym, k) {
    var c = CACHE_SI[ym];
    if (!c) { c = CACHE_SI[ym] = {}; CTE.forEach(function (o) { if (o.ym === ym && o.soma) c[o.vendedora] = (c[o.vendedora] || 0) + o.freteSI; }); }
    return c[k] || 0;
  }
  /* ===================================================================
     COMPLEMENTO DE CARGA: CT-e sem saída (aguardando embarque) com frete acima de CONFIG.complementoAcimaDe
     =================================================================== */
  var COMPL_ORDEM = 'frete';
  function desenharComplemento() {
    var card = $('#cardCompl'); if (!card) return;
    var lista = FC.filter(function (o) { return o.complemento; });
    lista.sort(COMPL_ORDEM === 'data' ? function (a, b) { return a.dt < b.dt ? -1 : a.dt > b.dt ? 1 : b.frete - a.frete; } : function (a, b) { return b.frete - a.frete; });
    var t = total(lista), pesoT = sum(lista, function (o) { return o.peso; });
    $('#complInfo').innerHTML = 'CT-e <b>sem saída</b> (aguardando embarque, sem manifesto) com Valor do Frete <b>acima de ' + brl(CONFIG.complementoAcimaDe) + '</b>. Clique na linha para ver o CT-e no detalhamento.';
    $('#complResumo').innerHTML = '<div><span>CT-e</span><b class="num">' + nInt(lista.length) + '</b></div><div><span>Valor do frete</span><b class="num">' + brl(t.fat) + '</b></div>' +
      '<div><span>Peso real</span><b class="num">' + fNum.format(pesoT) + ' kg</b></div>' +
      '<button type="button" class="btn-sec' + (F.compl ? ' on' : '') + '" id="complFiltro">' + (F.compl ? 'Mostrar todos os CT-e' : 'Filtrar o painel só por eles') + '</button>';
    var corpo = lista.length ? lista.map(function (o) {
      return '<tr data-n="' + esc(o.numero) + '" tabindex="0"><td>' + esc(o.numero) + '</td><td>' + esc(fDate(o.dt)) + '</td><td><span class="ell" title="' + esc(o.cliente) + '">' + esc(o.cliente) + '</span></td>' +
        '<td>' + esc(titulo(o.cidade)) + ' / ' + esc(o.uf) + '</td><td>' + esc(o.vendNome) + '</td><td><span class="ell" title="' + esc(nomeOcorr(o.ocorr)) + '">' + esc(nomeOcorr(o.ocorr)) + '</span></td>' +
        '<td class="r num">' + fNum.format(o.peso) + '</td><td class="r num"><b>' + brl(o.frete) + '</b></td></tr>';
    }).join('') : '<tr><td colspan="8" class="empty">Nenhum CT-e sem saída acima de ' + brl(CONFIG.complementoAcimaDe) + ' nos filtros selecionados.</td></tr>';
    $('#complTab').innerHTML = '<thead><tr><th>CT-e</th><th>Emissão</th><th>Cliente pagador</th><th>Destino</th><th>Vendedora</th><th>Ocorrência</th><th class="r">Peso (kg)</th><th class="r">Valor do frete</th></tr></thead><tbody>' + corpo + '</tbody>' +
      (lista.length ? '<tfoot><tr><td>TOTAL</td><td colspan="5">' + nInt(lista.length) + ' CT-e</td><td class="r num">' + fNum.format(pesoT) + '</td><td class="r num">' + brl(t.fat) + '</td></tr></tfoot>' : '');
  }
  $('#cardCompl').addEventListener('click', function (e) {
    if (e.target.closest('#complFiltro')) { definirFiltro('compl', '1', true); return; }
    var o = e.target.closest('[data-ord]'); if (o) { COMPL_ORDEM = o.getAttribute('data-ord'); $$('#cardCompl [data-ord]').forEach(function (b) { b.classList.toggle('on', b === o); }); desenharComplemento(); return; }
    var tr = e.target.closest('tr[data-n]'); if (!tr) return;
    definirFiltro('q', tr.getAttribute('data-n'));
    var tb = $('#cardTable'); if (tb) tb.scrollIntoView({ behavior: 'smooth', block: 'start' });
  });
  $('#cardCompl').addEventListener('keydown', function (e) { if (e.key === 'Enter' && e.target.closest('tr[data-n]')) e.target.click(); });

  /* ===================================================================
     v50 ENTREGAS EM ATRASO: CT-e sem entrega realizada com prazo vencido (todos os meses;
     respeita vendedora, UF, cliente e busca, mas NÃO o período)
     =================================================================== */
  var ATR_ORDEM = 'dias', ATR_FAIXA = -1, ATR_MODO = 'atraso';   // v51: 'atraso' = em atraso | 'prazo' = no prazo
  function faixasAtr() { return ATR_MODO === 'prazo' ? CONFIG.prazoFaixas : CONFIG.atrasoFaixas; }
  function diasAtr(o) { return ATR_MODO === 'prazo' ? o.diasFalta : o.diasAtraso; }
  function faixaAtraso(d) { var f = faixasAtr(); for (var i = 0; i < f.length; i++) if (d <= f[i]) return i; return f.length; }
  function rotuloFaixa(i) {
    var f = faixasAtr(), pz = ATR_MODO === 'prazo', ini = i ? f[i - 1] + 1 : (pz ? 0 : 1);
    if (i >= f.length) return 'mais de ' + f[f.length - 1] + ' dias';
    if (pz && f[i] === 0) return 'vence hoje';
    if (pz && ini === f[i] && ini === 1) return 'vence amanhã';
    return ini === f[i] ? ini + ' dia' + (ini > 1 ? 's' : '') : ini + ' a ' + f[i] + ' dias';
  }
  function textoDias(d) {
    if (ATR_MODO === 'prazo') return d === 0 ? 'vence hoje' : d === 1 ? 'falta 1 dia' : 'faltam ' + nInt(d) + ' dias';
    return nInt(d) + (d > 1 ? ' dias' : ' dia');
  }
  function desenharAtraso() {
    var card = $('#cardAtraso'); if (!card) return;
    var pz = ATR_MODO === 'prazo';
    var G = {}; Object.keys(F).forEach(function (k) { G[k] = F[k]; });
    G.atraso = '1'; var nAtr = filtrarCte(G);
    G.atraso = 'prazo'; var nPz = filtrarCte(G);
    var todos = pz ? nPz : nAtr;
    var nF = faixasAtr().length + 1, porFaixa = [];
    for (var i = 0; i < nF; i++) porFaixa.push({ q: 0, v: 0 });
    todos.forEach(function (o) { var b = porFaixa[faixaAtraso(diasAtr(o))]; b.q++; b.v += o.frete; });
    if (ATR_FAIXA >= 0 && !porFaixa[ATR_FAIXA].q) ATR_FAIXA = -1;
    var lista = ATR_FAIXA < 0 ? todos : todos.filter(function (o) { return faixaAtraso(diasAtr(o)) === ATR_FAIXA; });
    lista.sort(ATR_ORDEM === 'frete' ? function (a, b) { return b.frete - a.frete; }
      : pz ? function (a, b) { return a.diasFalta - b.diasFalta || b.frete - a.frete; }
           : function (a, b) { return b.diasAtraso - a.diasAtraso || b.frete - a.frete; });
    var t = total(todos), fAtivo = F.atraso === (pz ? 'prazo' : '1');
    // v52 INDICADORES DO PERÍODO: % entregue e % entregue dentro do prazo (CT-e emitidos no período filtrado)
    G.atraso = ''; var base = filtrarCte(G).filter(function (o) { return !o.cancelado && CONFIG.atrasoIgnorarOcorr.indexOf(o.codOcorr) < 0; });
    var ent = base.filter(function (o) { return o.entregueEm; });
    var entP = ent.filter(function (o) { return o.prazo; }), ok = entP.filter(function (o) { return o.entregueEm <= o.prazo; });
    var dts = base.map(function (o) { return o.dt; }).sort();
    var perTxt = dts.length ? (dts[0] === dts[dts.length - 1] ? fDate(dts[0]) : fDate(dts[0]) + ' a ' + fDate(dts[dts.length - 1])) : '—';
    var pE = base.length ? ent.length / base.length : NaN, pP = entP.length ? ok.length / entP.length : NaN;
    // v53: indicadores em etiquetas compactas (mesmo estilo das faixas de Prêmios)
    var kpi = function (rot, p, det, cls) { return '<span class="atr-k ' + cls + '"><b>' + rot + '</b> <em class="num">' + pct(p) + '</em> <i>' + det + '</i><u><s style="width:' + (isFinite(p) ? Math.max(0, Math.min(100, p * 100)).toFixed(1) : 0) + '%"></s></u></span>'; };
    $('#atrKpis').innerHTML = kpi('Entregues', pE, nInt(ent.length) + ' de ' + nInt(base.length) + ' CT-e', 'k-ent') +
      kpi('Dentro do prazo', pP, nInt(ok.length) + ' de ' + nInt(entP.length) + ' entregues', 'k-pz') +
      '<span class="pr2-nota">emitidos em ' + esc(perTxt) + ' (período do filtro)</span>';
    $('#atrTit').textContent = 'Entregas';
    $('#atrModo').innerHTML = '<button type="button" data-modo="atraso" class="m-atr' + (pz ? '' : ' on') + '">Em atraso <b>' + nInt(nAtr.length) + '</b></button>' +
      '<button type="button" data-modo="prazo" class="m-pz' + (pz ? ' on' : '') + '">No prazo <b>' + nInt(nPz.length) + '</b></button>';
    $('#atrInfo').innerHTML = 'Prazo = Entrega Programada ou Previsão de Entrega. Listas com CT-e <b>sem entrega</b> de todos os meses.';
    var hoje = todos.filter(function (o) { return o.diasFalta === 0; }).length;
    var maior = todos.reduce(function (m, o) { return Math.max(m, o.diasAtraso); }, 0);
    $('#atrResumo').innerHTML = '<span class="atr-res"><b class="num">' + nInt(todos.length) + '</b> CT-e · <b class="num">' + brl(t.fat) + '</b> de frete · ' +
      (pz ? '<b class="num">' + nInt(hoje) + '</b> vencem hoje' : 'maior atraso <b class="num">' + nInt(maior) + ' dias</b>') + '</span>' +
      '<button type="button" class="btn-sec' + (fAtivo ? ' on' : '') + '" id="atrFiltro">' + (fAtivo ? 'Mostrar todos os CT-e' : 'Filtrar o painel só por eles') + '</button>';
    $('#atrFaixas').innerHTML = '<button type="button" data-fx="-1" class="' + (ATR_FAIXA < 0 ? 'on' : '') + '">Todos <b>' + nInt(todos.length) + '</b></button>' +
      porFaixa.map(function (b, i) {
        return '<button type="button" data-fx="' + i + '" class="' + (pz ? 'pz' : 'fx') + i + (ATR_FAIXA === i ? ' on' : '') + '"' + (b.q ? '' : ' disabled') + ' title="' + esc(brl(b.v)) + ' de frete"><i></i>' + rotuloFaixa(i) + ' <b>' + nInt(b.q) + '</b></button>';
      }).join('');
    var corpo = lista.length ? lista.map(function (o) {
      return '<tr data-n="' + esc(o.numero) + '" tabindex="0"><td>' + esc(o.numero) + '</td><td>' + esc(fDate(o.dt)) + '</td><td><span class="ell" title="' + esc(o.cliente) + '">' + esc(o.cliente) + '</span></td>' +
        '<td>' + esc(titulo(o.cidade)) + ' / ' + esc(o.uf) + '</td><td>' + esc(o.vendNome) + '</td><td>' + esc(fDate(o.prazo)) + '</td>' +
        '<td class="r"><span class="atr-d ' + (pz ? 'pz' : 'fx') + faixaAtraso(diasAtr(o)) + '">' + textoDias(diasAtr(o)) + '</span></td>' +
        '<td><span class="ell" title="' + esc(o.codOcorrTxt) + '">' + esc(o.codOcorrTxt || '—') + '</span></td>' +
        '<td class="r num"><b>' + brl(o.frete) + '</b></td></tr>';
    }).join('') : '<tr><td colspan="9" class="empty">' + (pz ? 'Nenhuma entrega no prazo aguardando nos filtros selecionados.' : 'Nenhuma entrega em atraso nos filtros selecionados. 👍') + '</td></tr>';
    $('#atrTab').innerHTML = '<thead><tr><th>CT-e</th><th>Emissão</th><th>Cliente pagador</th><th>Destino</th><th>Vendedora</th><th>Prazo de entrega</th><th class="r">' + (pz ? 'Faltam' : 'Atraso') + '</th><th>Última ocorrência</th><th class="r">Valor do frete</th></tr></thead><tbody>' + corpo + '</tbody>' +
      (lista.length ? '<tfoot><tr><td>TOTAL</td><td colspan="7">' + nInt(lista.length) + ' CT-e' + (ATR_FAIXA >= 0 ? ' · ' + rotuloFaixa(ATR_FAIXA) : '') + '</td><td class="r num">' + brl(total(lista).fat) + '</td></tr></tfoot>' : '');
    // v54: dados para os botões Excel e PDF (a lista que está na tela: aba + faixa)
    TABELAS.atrTab = { cols: [{ l: 'CT-e', t: 't' }, { l: 'Emissão', t: 't' }, { l: 'Cliente pagador', t: 't' }, { l: 'Destino', t: 't' }, { l: 'Vendedora', t: 't' },
        { l: 'Prazo de entrega', t: 't' }, { l: pz ? 'Faltam (dias)' : 'Atraso (dias)', t: 'i' }, { l: 'Última ocorrência', t: 't' }, { l: 'Valor do frete', t: 'm' }],
      linhas: lista.map(function (o) { return [o.numero, fDate(o.dt), o.cliente, titulo(o.cidade) + ' / ' + o.uf, o.vendNome, fDate(o.prazo), diasAtr(o), o.codOcorrTxt || '', o.frete]; }),
      total: ['TOTAL', nInt(lista.length) + ' CT-e', '', '', '', '', null, '', total(lista).fat] };
    ATR_TITULO = (pz ? 'Entregas no prazo' : 'Entregas em atraso') + (ATR_FAIXA >= 0 ? ' - ' + rotuloFaixa(ATR_FAIXA) : '') + ' (todos os meses)';
  }
  var ATR_TITULO = 'Entregas em atraso';
  $('#cardAtraso').addEventListener('click', function (e) {
    var modo = 'prazo' === ATR_MODO ? 'prazo' : '1';
    if (e.target.closest('#atrFiltro')) { definirFiltro('atraso', modo, true); return; }
    var m = e.target.closest('[data-modo]'); if (m) { ATR_MODO = m.getAttribute('data-modo'); ATR_FAIXA = -1; desenharAtraso(); return; }
    var fx = e.target.closest('[data-fx]'); if (fx) { ATR_FAIXA = +fx.getAttribute('data-fx'); desenharAtraso(); return; }
    var bx = e.target.closest('[data-a]');
    if (bx) {
      if (bx.getAttribute('data-a') === 'xls') baixarExcel('atrTab', ATR_TITULO);
      else baixarPDF($('#cardAtraso'), { id: 'atrTab', style: { display: 'none' } }, ATR_TITULO);
      return;
    }
    var tr = e.target.closest('tr[data-n]'); if (!tr) return;
    F.atraso = modo;
    definirFiltro('q', tr.getAttribute('data-n'));
    var tb = $('#cardTable'); if (tb) tb.scrollIntoView({ behavior: 'smooth', block: 'start' });
  });
  $('#cardAtraso').addEventListener('keydown', function (e) { if (e.key === 'Enter' && e.target.closest('tr[data-n]')) e.target.click(); });

  function desenharPremios() {
    var card = $('#cardPremios'); if (!card) return;
    CACHE_PR = {}; CACHE_SI = {};
    var ym = mesReferencia(), M = METAS[ym] || {};
    var quem = (LISTA_CARDS[ym] || []).filter(function (k) { return M[k] > 0; });
    var comOn = TEM_COMISSAO();
    card.hidden = !PREMIOS.length && !comOn;
    if (card.hidden) return;
    $('#premTit').textContent = PREMIOS.length && comOn ? 'Prêmios e comissão do mês' : (comOn ? 'Comissão do mês' : 'Prêmios do mês');
    $('#premInfo').innerHTML = 'Sobre o faturamento <b>sem ICMS</b> de ' + nomeMes(ym) + ' (mês inteiro).' +
      (F.de || F.ate || F.dia ? ' <b>Prêmio e comissão são sempre do mês inteiro</b> (as datas De/Até não mudam estes valores).' : '');
    var tons = ['#8BD06A', '#3DB46D', '#2E8B57', '#2C6E49', '#1F4D33', '#163A26'];
    var maxF = PREMIOS.length ? PREMIOS[PREMIOS.length - 1].min : 1;
    // LAYOUT COMPACTO (v46): faixas numa linha fina + tabela com uma linha por vendedora
    var exige = PREMIOS.filter(function (p) { return p.exigeMeta; }).map(function (p) { return ORD(p.faixa); });
    var h = !PREMIOS.length ? '' : '<div class="pr2-faixas">' + PREMIOS.map(function (p, i) {
      return '<span class="pr2-fx" style="--t:' + tons[Math.min(i, tons.length - 1)] + '"><b>' + ORD(p.faixa) + '</b> ' + brlK(p.min) + ' <i>→</i> <em class="num">' + brl(p.premio) + '</em>' + (p.exigeMeta ? '<sup title="Só é pago se a vendedora bater a meta">*</sup>' : '') + '</span>';
    }).join('') + (exige.length ? '<span class="pr2-nota">* só com a meta batida</span>' : '') + '</div>';
    var linhas = quem.map(function (k) { var real = premioReal(ym, k), meta = M[k], pc = pctComissao(k), base = comissaoBase(ym, k); return { k: k, real: real, meta: meta, pr: premioDe(real, meta), pc: pc, base: base, com: base * pc }; })
      .sort(function (a, b) { return b.real - a.real; });
    if (!linhas.length) { $('#premBody').innerHTML = h + '<div class="vend-vazio">Nenhuma vendedora com meta em ' + nomeMes(ym) + ' no metas.csv.</div>'; return; }
    var pcTxt = comOn ? fPct(linhas[0].pc) : '';
    var cab = '<tr><th>Vendedora</th><th class="r">Faturado s/ ICMS</th>' + (PREMIOS.length ? '<th class="pr2-barra-th">Faixas</th><th>Próxima faixa</th><th class="r">Prêmio</th>' : '') +
      (comOn ? '<th class="r">Comissão' + (linhas.every(function (l) { return l.pc === linhas[0].pc; }) ? ' (' + pcTxt + ')' : '') + '</th>' : '') + (PREMIOS.length && comOn ? '<th class="r">A receber</th>' : '') + '</tr>';
    var corpo = linhas.map(function (l) {
      var nome = NOMES[l.k] || l.k, pos = function (v) { return Math.min(100, v / maxF * 100).toFixed(1); };
      var barra = '<div class="pr2-barra" title="' + esc(brl(l.real) + ' sem ICMS de ' + brl(maxF) + ' (última faixa)') + '"><i style="width:' + pos(l.real) + '%"></i>' +
        PREMIOS.map(function (p) { return '<em class="' + (l.real >= p.min ? 'ok' : '') + '" style="left:' + pos(p.min) + '%" title="' + ORD(p.faixa) + ' faixa: ' + brl(p.min) + '"></em>'; }).join('') + '</div>';
      var premio = l.pr.atual ? '<span class="pr2-badge" style="--t:' + tons[Math.min(PREMIOS.indexOf(l.pr.atual), tons.length - 1)] + '">' + ORD(l.pr.atual.faixa) + '</span> <b class="num">' + brl(l.pr.premio) + '</b>' : '<span class="pr2-zero">—</span>';
      var prox = l.pr.prox ? 'faltam <b class="num">' + brl(l.pr.prox.falta) + '</b> → ' + ORD(l.pr.prox.p.faixa) + (l.pr.prox.p.exigeMeta && l.meta > l.pr.prox.p.min ? '*' : '') : '<span class="pr2-max">faixa máxima!</span>';
      return '<tr class="' + (l.pr.atual ? 'ganhou' : '') + '"><td><span class="pr2-nome"><span class="vc-av">' + esc(nome.charAt(0)) + '</span>' + esc(nome) + '</span></td>' +
        '<td class="r num">' + brl(l.real) + '</td>' + (PREMIOS.length ? '<td class="pr2-barra-td">' + barra + '</td><td class="pr2-prox">' + prox + '</td><td class="r">' + premio + '</td>' : '') +
        (comOn ? '<td class="r num" title="' + esc(fPct(l.pc) + ' de ' + brl(l.base)) + '">' + brl(l.com) + '</td>' : '') + (PREMIOS.length && comOn ? '<td class="r num"><b>' + brl(l.com + l.pr.premio) + '</b></td>' : '') + '</tr>';
    }).join('');
    var totPr = sum(linhas, function (l) { return l.pr.premio; }), totCom = sum(linhas, function (l) { return l.com; });
    var rod = '<tr><td>Total</td><td class="r num">' + brl(sum(linhas, function (l) { return l.real; })) + '</td>' + (PREMIOS.length ? '<td></td><td>' + (linhas.some(function (l) { return l.pr.atual; }) ? '' : '<span class="pr2-zero">nenhuma ganhadora ainda</span>') + '</td><td class="r"><b class="num">' + brl(totPr) + '</b></td>' : '') +
      (comOn ? '<td class="r num">' + brl(totCom) + '</td>' : '') + (PREMIOS.length && comOn ? '<td class="r num"><b>' + brl(totCom + totPr) + '</b></td>' : '') + '</tr>';
    h += '<div class="pr2-wrap"><table class="pr2-tab"><thead>' + cab + '</thead><tbody>' + corpo + '</tbody><tfoot>' + rod + '</tfoot></table></div>';
    $('#premBody').innerHTML = h;
  }

  /* ===================================================================
     GRÁFICOS (Chart.js)
     =================================================================== */
  var TEM_CHART = function () { return typeof window.Chart === 'function'; };
  function temaGraficos() {
    var C = window.Chart;
    C.defaults.font.family = '"Segoe UI", system-ui, -apple-system, Roboto, Arial, sans-serif';
    C.defaults.font.size = 12;
    C.defaults.color = css('--muted');
    C.defaults.borderColor = css('--grid');
    C.defaults.animation.duration = 600;
    var tt = C.defaults.plugins.tooltip;
    tt.backgroundColor = css('--surface'); tt.titleColor = css('--ink'); tt.bodyColor = css('--ink-2');
    tt.footerColor = css('--ink-2'); tt.borderColor = css('--line'); tt.borderWidth = 1;
    tt.padding = 10; tt.boxPadding = 4; tt.usePointStyle = true; tt.cornerRadius = 10;
    tt.titleFont = { weight: '700' }; tt.footerFont = { weight: '400', size: 11.5 };
  }
  // plugin: escreve o total no centro das roscas (a letra diminui para caber)
  var pluginCentro = {
    id: 'centroTexto',
    afterDraw: function (chart, args, opts) {
      if (!opts || !opts.texto) return;
      var a = chart.chartArea, ctx = chart.ctx, x = (a.left + a.right) / 2, y = (a.top + a.bottom) / 2;
      var raio = Math.min(a.right - a.left, a.bottom - a.top) / 2 * 0.68, tam = 24;
      ctx.save(); ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      ctx.font = '600 ' + tam + 'px Bahnschrift, "Segoe UI", system-ui, Arial';
      while (tam > 11 && ctx.measureText(opts.texto).width > raio * 1.8) { tam--; ctx.font = '600 ' + tam + 'px Bahnschrift, "Segoe UI", system-ui, Arial'; }
      ctx.fillStyle = css('--ink'); ctx.fillText(opts.texto, x, y - 6);
      ctx.fillStyle = css('--muted'); ctx.font = '400 11px "Segoe UI", system-ui, Arial';
      ctx.fillText(opts.sub || '', x, y + 14);
      ctx.restore();
    }
  };
  /* VALORES ESCRITOS NOS GRÁFICOS: cada barra/ponto mostra o número da medida escolhida
     (Faturamento em R$ completo, CT-e ou Cotações em quantidade). Barras em pé: o valor fica em
     cima da barra (deitado, se couber; senão em pé, dentro ou acima). Barras deitadas: no fim da
     barra. Linhas: em cima de cada ponto, só quando há espaço (poucos pontos). */
  var pluginValores = {
    id: 'valores',
    afterDatasetsDraw: function (chart, args, o) {
      if (!o || !o.fmt) return;
      var ctx = chart.ctx, deitada = chart.options.indexAxis === 'y', area = chart.chartArea, tinta = css('--ink-2');
      ctx.save();
      ctx.font = '600 11px "Segoe UI", system-ui, Arial'; ctx.textBaseline = 'middle';
      chart.data.datasets.forEach(function (ds, di) {
        var meta = chart.getDatasetMeta(di);
        if (meta.hidden || ds.semValores) return;
        if (meta.type === 'line') {
          var n = meta.data.length;
          if (!o.linhas || !n || (area.right - area.left) / n < (o.minLargura || 78)) return;
        }
        meta.data.forEach(function (el, i) {
          var v = ds.data[i];
          if (v == null || !isFinite(v) || (v === 0 && !o.zero)) return;
          var t = o.fmt(v), tw = ctx.measureText(t).width;
          if (meta.type === 'line') {
            var lx = Math.min(area.right - tw / 2, Math.max(area.left + tw / 2, el.x)), ly = Math.max(area.top + 6, el.y - 12);
            ctx.textAlign = 'center'; ctx.lineWidth = 3; ctx.strokeStyle = css('--surface'); ctx.strokeText(t, lx, ly);   // contorno: lê-se mesmo sobre a linha
            ctx.fillStyle = ds.borderColor || tinta; ctx.fillText(t, lx, ly);
          } else if (deitada) {
            ctx.fillStyle = tinta; ctx.textAlign = 'left';
            ctx.fillText(t, el.x + 6, el.y);
          } else {
            var larg = Math.abs(el.width || 0), alt = Math.abs((el.base || 0) - el.y);
            if (tw <= larg + 10) { ctx.fillStyle = tinta; ctx.textAlign = 'center'; ctx.fillText(t, el.x, el.y - 9); }
            else {
              ctx.save(); ctx.translate(el.x, el.y); ctx.rotate(-Math.PI / 2);
              if (alt > tw + 12) { ctx.fillStyle = '#FFFFFF'; ctx.textAlign = 'right'; ctx.fillText(t, -6, 0); }
              else { ctx.fillStyle = tinta; ctx.textAlign = 'left'; ctx.fillText(t, 5, 0); }
              ctx.restore();
            }
          }
        });
      });
      ctx.restore();
    }
  };
  var VALORES_REGISTRADO = false;
  function montar(id, cfg, recriar) {
    if (!VALORES_REGISTRADO && window.Chart && window.Chart.register) { window.Chart.register(pluginValores); VALORES_REGISTRADO = true; }
    // espaço extra para os valores escritos (fim das barras deitadas / acima das barras em pé)
    cfg.options = cfg.options || {}; cfg.options.plugins = cfg.options.plugins || {};
    cfg.options.plugins.valores = cfg.valores || false;
    if (cfg.valores) {
      var deit = cfg.options.indexAxis === 'y';
      cfg.options.layout = { padding: deit ? { right: cfg.valores.folga || 96 } : { top: 6 } };
      var eixo = cfg.options.scales && cfg.options.scales[deit ? 'x' : 'y'];
      if (eixo) eixo.grace = deit ? 0 : (cfg.type === 'bar' ? '22%' : '12%');
    }
    var canvas = document.getElementById(id), g = GRAFICOS[id];
    var eixo = (cfg.options && cfg.options.indexAxis) || 'x';
    if (g && !recriar && g.config.type === cfg.type && (g.options.indexAxis || 'x') === eixo) {
      g.data.labels = cfg.data.labels; g.data.datasets = cfg.data.datasets; g.options = cfg.options; g.update();
      return g;
    }
    if (g) g.destroy();
    GRAFICOS[id] = new window.Chart(canvas, cfg);
    return GRAFICOS[id];
  }
  function msgNoGrafico(id, msg) {
    var box = document.getElementById(id).parentNode, m = box.querySelector('.chart-msg');
    if (msg) {
      if (GRAFICOS[id]) { GRAFICOS[id].destroy(); delete GRAFICOS[id]; }
      if (!m) { m = document.createElement('div'); m.className = 'chart-msg'; box.appendChild(m); }
      m.textContent = msg; document.getElementById(id).style.display = 'none';
    } else { if (m) m.remove(); document.getElementById(id).style.display = ''; }
  }
  function cursor(e, els) { if (e.native && e.native.target) e.native.target.style.cursor = els.length ? 'pointer' : 'default'; }
  function opcoesHBar(fmt, maxChars, onClickItem) {
    return {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { callbacks: {} } },
      scales: {
        x: { beginAtZero: true, grid: { color: css('--grid') }, ticks: { callback: function (v) { return fmt(v); }, maxTicksLimit: 4, maxRotation: 0 }, border: { display: false } },
        y: { grid: { display: false }, border: { display: false }, ticks: { color: css('--ink-2'), autoSkip: false, callback: function (v) { return cortar(this.getLabelForValue(v), maxChars); } } }
      },
      onHover: cursor,
      onClick: function (e, els) { if (els.length) onClickItem(els[0].index); }
    };
  }
  function larguraDe(id) { return document.getElementById(id).parentNode.clientWidth || 300; }
  function alturaRanking(boxId, n) { document.getElementById(boxId).style.height = Math.max(160, n * 30 + 40) + 'px'; }
  function coresDestaque(itens, ativo, cor) { return itens.map(function (d) { return !ativo || ativo === d.key ? cor : alfa(cor, .28); }); }

  // A CONTA COMPLETA que aparece ao passar o mouse
  function rodapeGrupo(g, semClique) {
    var l = ['CT-e: ' + nInt(g.qtd) + (g.canc ? ' (' + nInt(g.canc) + ' cancelados, só contagem)' : ''),
             'Faturamento: ' + brl(g.fat), '(−) ICMS: ' + brl(g.icms), '(=) Frete sem ICMS: ' + brl(g.fsi), 'Ticket médio: ' + brl(g.ticket),
             'Cotações: ' + nInt(g.cot) + (g.cotCanc ? ' (' + nInt(g.cotCanc) + ' canceladas)' : ''),
             'Frete inicial: ' + brl(g.valorIni), 'Frete atual (valor cotado): ' + brl(g.valor), 'Variação do frete: ' + (g.varFrete > 0 ? '+' : '') + fDec.format(g.varFrete * 100) + '%',
             'Fechamento: ' + pct(g.fechPct) + ' (' + nInt(g.val) + ' CT-e válidos ÷ ' + nInt(g.cotVal) + ' cotações válidas)'];
    if (!semClique) l.push('', 'Clique para filtrar');
    return l;
  }
  function valorMetrica(g, k) { return k === 'qtd' ? g.qtd : k === 'cot' ? g.cot : k === 'meta' ? (g.pMeta || 0) : g.fat; }
  function fmtMetrica(k) { return k === 'qtd' || k === 'cot' ? nInt : k === 'meta' ? pct : brlc; }
  function fmtMetricaCheia(k) { return k === 'qtd' || k === 'cot' ? nInt : k === 'meta' ? pct : brl; }

  function corBaixa(s) {
    s = s.toUpperCase();
    if (/CANCEL/.test(s)) return css('--bad');
    if (/LIQUID/.test(s)) return css('--good');
    if (/FATUR/.test(s)) return css('--s1');
    if (/PENDEN/.test(s)) return css('--warn');
    return css('--muted');
  }
  function corSituacao(s) { return s === 'Cancelada' ? css('--bad') : s === 'CT-e emitido' ? css('--good') : s === 'Contratada' ? css('--s1') : s === 'Cotada' ? css('--warn') : css('--muted'); }
  function legendaRosca(ulId, itens, totalV, ativo, onClick, fmt) {
    var ul = document.getElementById(ulId);
    ul.innerHTML = itens.map(function (d, i) {
      return '<li data-i="' + i + '" class="' + (ativo === d.key ? 'on' : (ativo ? 'dim' : '')) + '"><i style="background:' + d.cor + '"></i><span>' + esc(d.rotulo).replace(/\u0001(.*?)\u0002/, '<em class="dl-cod">$1</em>') + (d.det ? '<small class="dl-det">' + d.det + '</small>' : '') +
        '</span><span class="v num">' + fmt(d.valor) + '</span><span class="p num">' + pct(totalV ? d.valor / totalV : 0) + '</span></li>';
    }).join('');
    ul.onclick = function (e) { var li = e.target.closest('li'); if (li) onClick(itens[+li.getAttribute('data-i')]); };
  }
  function rosca(id, ulId, itens, ativo, onClick, recriar, fmt, sub, tipFn) {
    fmt = fmt || nInt;
    var totalV = sum(itens, function (d) { return d.valor; });
    legendaRosca(ulId, itens, totalV, ativo, onClick, fmt);
    if (!itens.length || !totalV) { msgNoGrafico(id, 'Sem dados para os filtros selecionados'); return; }
    msgNoGrafico(id, '');
    montar(id, {
      type: 'doughnut',
      data: { labels: itens.map(function (d) { return d.rotulo; }),
              datasets: [{ data: itens.map(function (d) { return d.valor; }),
                           backgroundColor: itens.map(function (d) { return !ativo || ativo === d.key ? d.cor : alfa(d.cor, .25); }),
                           borderColor: css('--surface'), borderWidth: 2, hoverOffset: 6 }] },
      options: {
        responsive: true, maintainAspectRatio: false, cutout: '68%',
        plugins: {
          legend: { display: false },
          centroTexto: { texto: fmt(totalV), sub: sub },
          tooltip: { callbacks: {
            label: function (c) { return ' ' + fmt(c.parsed) + ' (' + pct(totalV ? c.parsed / totalV : 0) + ')'; },
            footer: function (it) { var l = tipFn ? tipFn(itens[it[0].dataIndex]) : []; return l.concat(['', 'Clique para filtrar']); } } }
        },
        onHover: cursor,
        onClick: function (e, els) { if (els.length) onClick(itens[els[0].index]); }
      },
      plugins: [pluginCentro]
    }, recriar);
  }

  // granularidade da evolução: por dia (até ~2 meses) ou por mês
  function porDia() {
    var ds = FC.concat(FQ).map(function (o) { return o.dt; }).sort();
    if (!ds.length) return true;
    return (dataDe(ds[ds.length - 1]) - dataDe(ds[0])) / 864e5 <= 62;
  }
  function serieEvolucao() {
    var diario = porDia(), campo = diario ? 'dt' : 'ym';
    var gs = agrupar(FC, FQ, campo, campo).sort(function (a, b) { return a.key < b.key ? -1 : 1; });
    return { diario: diario, grupos: gs, rot: gs.map(function (g) { return diario ? g.key.slice(8, 10) + '/' + g.key.slice(5, 7) : nomeMes(g.key); }) };
  }

  function desenharGraficos(recriar) {
    var ids = ['chEvo','chUF','chRitmo','chOrigem','chVend','chCli','chSit','chOcor'];
    desenharMapa();
    montarTabelasDosGraficos();
    if (!TEM_CHART()) {
      ids.forEach(function (id) { msgNoGrafico(id, 'Biblioteca de gráficos indisponível: sem internet e sem o arquivo assets\\chart.umd.min.js. Veja o LEIAME.txt.'); });
      return;
    }
    temaGraficos();
    var C1 = css('--s1'), C2 = css('--s2'), C3 = css('--s3'), C4 = css('--s4');
    var acum = function (arr) { var s = 0; return arr.map(function (v) { s += v; return s; }); };
    var linha = function (rot, cor, dados, preencher, n) {
      return { label: rot, data: dados, borderColor: cor, backgroundColor: preencher ? alfa(cor, .14) : cor, fill: !!preencher,
               tension: .3, cubicInterpolationMode: 'monotone', borderWidth: 2.2, pointRadius: n > 40 ? 0 : 3, pointHoverRadius: 6,
               pointBackgroundColor: css('--surface'), pointBorderColor: cor, pointBorderWidth: 2 };
    };

    /* ---------- 1. Evolução (linhas) ---------- */
    var E = serieEvolucao(), gs = E.grupos, valores = UI.evoMed === 'fat';
    $('#evoInfo').textContent = (E.diario ? 'Por dia' : 'Por mês') + ' de emissão do CT-e e de inclusão da cotação. Clique em um ponto para filtrar.';
    if (!gs.length) msgNoGrafico('chEvo', 'Sem dados para os filtros selecionados');
    else {
      msgNoGrafico('chEvo', '');
      var s1 = gs.map(function (g) { return valores ? g.fat : g.qtd; }), s2 = gs.map(function (g) { return valores ? g.valor : g.cot; });
      if (UI.evo === 'a') { s1 = acum(s1); s2 = acum(s2); }
      var dsEvo = [linha(valores ? 'Faturamento (CT-e)' : 'CT-e emitidos', C1, s1, true, gs.length), linha(valores ? 'Valor cotado' : 'Cotações', C2, s2, false, gs.length)];
      var ym = mesReferencia(), meta = metaDe(ym, F.vendedora);
      if (valores && UI.evo === 'a' && meta && gs.every(function (g) { return g.key.slice(0, 7) === ym; }))
        dsEvo.push({ label: 'Meta de ' + nomeMes(ym), data: gs.map(function () { return meta; }), borderColor: C3, borderDash: [6, 5], borderWidth: 1.6, pointRadius: 0, pointHoverRadius: 0, fill: false, semValores: true });
      var fmtY = valores ? brl : nInt;
      montar('chEvo', {
        type: 'line', data: { labels: E.rot, datasets: dsEvo },
        options: {
          responsive: true, maintainAspectRatio: false, interaction: { mode: 'index', intersect: false },
          plugins: {
            legend: { position: 'top', align: 'start', labels: { usePointStyle: true, pointStyle: 'line', color: css('--ink-2'), boxWidth: 18 } },
            tooltip: { callbacks: {
              title: function (it) { var g = gs[it[0].dataIndex]; return E.diario ? fDate(g.key) : nomeMesLongo(g.key); },
              label: function (c) { return ' ' + c.dataset.label + (UI.evo === 'a' ? ' (acumulado)' : '') + ': ' + fmtY(c.parsed.y); },
              footer: function (it) { return [''].concat(rodapeGrupo(gs[it[0].dataIndex])); } } }
          },
          scales: { y: { beginAtZero: true, ticks: { callback: valores ? brlc : nInt, maxTicksLimit: 6 }, border: { display: false } }, x: { grid: { display: false }, ticks: { maxRotation: 0, autoSkipPadding: 12 } } },
          onHover: cursor,
          onClick: function (e, els) { if (!els.length) return; var k = gs[els[0].index].key; if (E.diario) definirFiltro('dia', k, true); else definirFiltro('ym', k, true); }
        },
        valores: { fmt: fmtY, linhas: true, minLargura: valores ? 96 : 40 }
      }, recriar);
    }

    /* ---------- 2. Vendas por UF (barras ao lado do mapa) ---------- */
    var porUF = agrupar(FC, FQ, 'uf', 'uf').filter(function (g) { return g.key !== '—'; }).map(function (g) { g.v = valorMetrica(g, UI.uf); return g; })
      .filter(function (g) { return g.v > 0; }).sort(function (a, b) { return b.v - a.v; });
    var estreito = larguraDe('chUF') < 560;
    if (!porUF.length) msgNoGrafico('chUF', 'Sem dados para os filtros selecionados');
    else {
      msgNoGrafico('chUF', '');
      var ufs = estreito ? porUF.slice(0, 12) : porUF, fmtUF = fmtMetrica(UI.uf);
      var clicaUF = function (i) { definirFiltro('uf', ufs[i].key, true); };
      var dsUF = { label: 'UF', data: ufs.map(function (g) { return g.v; }), backgroundColor: coresDestaque(ufs, F.uf, UI.uf === 'cot' ? C2 : C1), borderRadius: 4, borderSkipped: false, maxBarThickness: 44 };
      var tipUF = { callbacks: { title: function (it) { var k = ufs[it[0].dataIndex].key; return (MAPA_NOMES[k] || k) + ' (' + k + ')'; }, label: function () { return ''; }, footer: function (it) { return rodapeGrupo(ufs[it[0].dataIndex]); } } };
      if (estreito) {
        alturaRanking('boxUF', ufs.length);
        var oh = opcoesHBar(fmtUF, 6, clicaUF); oh.plugins.tooltip = tipUF;
        montar('chUF', { type: 'bar', data: { labels: ufs.map(function (g) { return g.key; }), datasets: [dsUF] }, options: oh, valores: { fmt: fmtMetricaCheia(UI.uf) } }, recriar);
      } else {
        document.getElementById('boxUF').style.height = '';
        montar('chUF', { type: 'bar', data: { labels: ufs.map(function (g) { return g.key; }), datasets: [dsUF] },
          options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false }, tooltip: tipUF },
            scales: { y: { beginAtZero: true, ticks: { callback: fmtUF, maxTicksLimit: 6 }, border: { display: false } }, x: { grid: { display: false }, ticks: { autoSkip: false, maxRotation: 0, font: { size: 11 } } } },
            onHover: cursor, onClick: function (e, els) { if (els.length) clicaUF(els[0].index); } }, valores: { fmt: fmtMetricaCheia(UI.uf) } }, recriar);
      }
    }

    /* ---------- 3. Ritmo da meta (acumulado do mês x ritmo necessário) ---------- */
    var R = dadosRitmo();
    $('#ritmoInfo').textContent = R.meta
      ? nomeMesLongo(R.ym) + (F.vendedora ? ' · ' + NOMES[F.vendedora] : '') + ': faturamento acumulado contra o ritmo necessário para a meta de ' + brl(R.meta) + ' (dias úteis).'
      : nomeMesLongo(R.ym) + ': faturamento acumulado. Defina a meta em dados\\metas.csv para ver o ritmo necessário.';
    if (!R.temDados && !R.meta) msgNoGrafico('chRitmo', 'Sem CT-e no mês para os filtros selecionados');
    else {
      msgNoGrafico('chRitmo', '');
      // v47: cores bem diferentes e dica (tooltip) explicando cada número em português simples
      var CR = C1, CM = css('--warn'), CP = css('--muted');
      var dsR = [linha('Vendido (acumulado)', CR, R.real, true, R.rot.length)];
      if (R.meta) dsR.push({ label: 'Meta até o dia', data: R.ideal, borderColor: CM, backgroundColor: CM, borderDash: [7, 5], borderWidth: 2, pointRadius: 0, pointHoverRadius: 5, fill: false, stepped: false });
      if (R.proj.some(function (v) { return v != null; })) dsR.push({ label: 'Projeção (se continuar neste ritmo)', data: R.proj, borderColor: CP, backgroundColor: CP, borderDash: [2, 4], borderWidth: 2, pointRadius: 0, pointHoverRadius: 5, fill: false });
      var hojeR = hojeIso();
      montar('chRitmo', {
        type: 'line', data: { labels: R.rot, datasets: dsR },
        options: {
          responsive: true, maintainAspectRatio: false, interaction: { mode: 'index', intersect: false }, spanGaps: false,
          plugins: {
            legend: { position: 'top', align: 'start', labels: { usePointStyle: true, pointStyle: 'line', color: css('--ink-2'), boxWidth: 22, font: { size: 12.5 } } },
            tooltip: {
              bodyFont: { size: 12.5 }, footerFont: { size: 12, weight: '600' }, titleFont: { size: 13, weight: '700' }, padding: 12, boxPadding: 6, usePointStyle: false,
              callbacks: {
                title: function (it) { var d = R.dias[it[0].dataIndex]; return fDate(d.iso) + (d.iso === hojeR ? ' · hoje' : d.iso > hojeR ? ' · dia futuro' : '') + (d.util ? '' : ' (não é dia útil)'); },
                label: function (c) { return c.parsed.y == null ? null : ' ' + c.dataset.label + ': ' + brl(c.parsed.y); },
                labelColor: function (c) { var cor = c.dataset.borderColor; return { borderColor: cor, backgroundColor: cor, borderWidth: 0, borderRadius: 2 }; },
                footer: function (it) {
                  var i = it[0].dataIndex, d = R.dias[i], dia = R.porDia[d.iso], l = [];
                  if (R.real[i] != null) {
                    l.push('Vendido neste dia: ' + brl(dia ? dia.fat : 0) + ' (' + nInt(dia ? dia.val : 0) + ' CT-e)');
                    if (R.meta) { var dif = R.real[i] - R.ideal[i]; l.push(dif >= 0 ? '✔ ' + brl(dif) + ' à frente da meta' : '✖ ' + brl(-dif) + ' atrás da meta');
                      var fimP = R.proj[R.proj.length - 1]; if (fimP != null) { var dfp = fimP - R.meta; l.push('Neste ritmo (fim do mês): ' + brl(Math.abs(dfp)) + (dfp >= 0 ? ' acima da meta' : ' abaixo da meta')); } }
                    l.push('', 'Clique para filtrar este dia');
                  } else if (R.meta && R.proj[i] != null) {
                    var dp = R.proj[i] - R.ideal[i];
                    l.push('Neste ritmo: ' + brl(Math.abs(dp)) + (dp >= 0 ? ' acima da meta' : ' abaixo da meta'));
                  }
                  return l;
                } } }
          },
          scales: { y: { beginAtZero: true, ticks: { callback: brlc, maxTicksLimit: 6 }, border: { display: false } }, x: { grid: { display: false }, ticks: { maxRotation: 0, autoSkipPadding: 10 } } },
          onHover: cursor,
          onClick: function (e, els) { if (els.length) definirFiltro('dia', R.dias[els[0].index].iso, true); }
        }
      }, recriar);
    }

    /* ---------- 4. Coleta x armazém (rosca, em R$) ---------- */
    var gOr = {}; agrupar(FC, null, 'origem').forEach(function (g) { gOr[g.key] = g; });
    var itensOr = ['COLETA', 'ARMAZEM', 'SEM'].filter(function (k) { return gOr[k] && gOr[k].fat > 0; })
      .map(function (k) { return { key: k, rotulo: ORIGENS[k], valor: gOr[k].fat, cor: k === 'COLETA' ? C1 : k === 'ARMAZEM' ? C3 : css('--muted'), g: gOr[k],
        det: '<b>' + nInt(gOr[k].qtd) + '</b> CT-e' + (gOr[k].canc ? ' (' + nInt(gOr[k].canc) + ' cancelado' + (gOr[k].canc > 1 ? 's' : '') + ')' : '') }; });
    var qtdOr = sum(itensOr, function (d) { return d.g.qtd; });
    rosca('chOrigem', 'lgOrigem', itensOr, F.origem, function (d) { definirFiltro('origem', d.key, true); }, recriar, brl, 'faturamento · ' + nInt(qtdOr) + ' CT-e',
      function (d) { return ['CT-e: ' + nInt(d.g.qtd) + (d.g.canc ? ' (' + nInt(d.g.canc) + ' cancelados, fora das somas)' : ''), 'CT-e válidos: ' + nInt(d.g.val), 'Ticket médio: ' + brl(d.g.ticket)]; });

    /* ---------- 5 e 6. Rankings ---------- */
    function ranking(id, boxId, grupos, metrica, cor, ativo, campoFiltro, rotuloFn, maxChars) {
      if (!grupos.length) { document.getElementById(boxId).style.height = ''; msgNoGrafico(id, 'Sem dados para os filtros selecionados'); return; }
      msgNoGrafico(id, '');
      alturaRanking(boxId, grupos.length);
      var op = opcoesHBar(fmtMetrica(metrica), maxChars, function (i) { definirFiltro(campoFiltro, grupos[i].key, true); });
      op.plugins.tooltip = { callbacks: {
        title: function (it) { return rotuloFn(grupos[it[0].dataIndex]); },
        label: function (c) { return metrica === 'meta' ? ' Meta: ' + pct(c.parsed.x) + ' (' + brl(grupos[c.dataIndex].realMes || 0) + ' de ' + brl(grupos[c.dataIndex].meta || 0) + ')' : ''; },
        footer: function (it) { return rodapeGrupo(grupos[it[0].dataIndex]); } } };
      montar(id, { type: 'bar',
        data: { labels: grupos.map(rotuloFn), datasets: [{ label: '', data: grupos.map(function (g) { return g.v; }),
                backgroundColor: coresDestaque(grupos, ativo, cor), borderRadius: 4, borderSkipped: false, barThickness: 16 }] },
        options: op, valores: { fmt: fmtMetricaCheia(metrica), folga: metrica === 'fat' ? 104 : 54 } }, recriar);
    }
    var largRot = function (id) { return Math.max(12, Math.min(34, Math.floor(larguraDe(id) / 16))); };
    var porVend = gruposVendedoras();
    porVend.forEach(function (g) { g.v = valorMetrica(g, UI.vend); });
    porVend = porVend.filter(function (g) { return UI.vend !== 'meta' || g.meta; }).sort(function (a, b) { return b.v - a.v; });
    if (UI.vend === 'meta' && !porVend.length) { document.getElementById('boxVend').style.height = ''; msgNoGrafico('chVend', 'Nenhuma meta definida para ' + nomeMes(mesReferencia()) + ' (dados\\metas.csv)'); }
    else ranking('chVend', 'boxVend', porVend, UI.vend, C4, F.vendedora, 'vendedora', function (g) { return NOMES[g.key] || g.key; }, largRot('chVend'));

    var porCli = agrupar(FC, FQ, 'cliente', 'cliente').filter(function (g) { return g.key !== '—'; }).map(function (g) { g.v = valorMetrica(g, UI.cli); return g; })
      .filter(function (g) { return g.v > 0; }).sort(function (a, b) { return b.v - a.v; }).slice(0, 10);
    ranking('chCli', 'boxCli', porCli, UI.cli, UI.cli === 'cot' ? C2 : C1, F.cliente, 'cliente', function (g) { return g.key; }, largRot('chCli'));

    /* ---------- 7. Situação das cotações (rosca) ---------- */
    var ordemSit = ['CT-e emitido', 'Contratada', 'Cotada', 'Cancelada'];
    var gSit = agrupar([], FQ, null, 'situacao').sort(function (a, b) { var ia = ordemSit.indexOf(a.key), ib = ordemSit.indexOf(b.key); return (ia < 0 ? 9 : ia) - (ib < 0 ? 9 : ib) || b.cot - a.cot; });
    rosca('chSit', 'lgSit', gSit.map(function (g) { return { key: g.key, rotulo: g.key, valor: g.cot, cor: corSituacao(g.key), g: g }; }), F.situacao,
      function (d) { definirFiltro('situacao', d.key, true); }, recriar, nInt, 'cotações',
      function (d) { if (d.key === 'Cancelada') return ['Canceladas não entram no valor cotado']; var ds = desconto(d.g);
        return ['Frete inicial: ' + brl(d.g.valorIni), 'Frete atual: ' + brl(d.g.valor), (ds.r >= 0 ? 'Desconto: ' : 'Acréscimo: ') + brl(Math.abs(ds.r)) + ' (' + pct(Math.abs(ds.p)) + ')']; });
    // (v27: quadro de valor inicial / valor atual / desconto retirado do card)

    /* ---------- 8. Status dos CT-e (só a relação, sem gráfico) ---------- */
    var listaStatus = function (itens, ativo, onClick) {
      var tot = sum(itens, function (d) { return d.valor; });
      legendaRosca('lgBaixa', itens, tot, ativo, onClick, nInt);
      $('#baixaTotal').innerHTML = itens.length ? '<b class="num">' + nInt(tot) + '</b> CT-e' : 'Sem dados para os filtros selecionados';
    };
    if (UI.baixa === 'ocor') {
      $('#baixaInfo').textContent = 'Código da Última Ocorrência';
      var gOc = agrupar(FC, null, 'codoc').sort(function (a, b) { return b.qtd - a.qtd || b.fat - a.fat; });
      var PAL = ['--s1', '--good', '--warn', '--s4', '--s3', '--s5', '--s2'].map(css);
      // na relação: código em destaque + descrição em letras normais ("62 · Saída da unidade"), mais leve que tudo em maiúsculas
      var frase = function (t) { t = String(t || '').toLowerCase().replace(/\bct-?e\b/g, 'CT-e').replace(/\bctrc\b/g, 'CTRC'); return t.charAt(0).toUpperCase() + t.slice(1); };
      var rotOc = function (k) { if (k === '—' || !OCORR[k]) return nomeCodOc(k); return '\u0001' + k + '\u0002' + frase(OCORR[k]); };
      listaStatus(gOc.map(function (g, i) { return { key: g.key, rotulo: rotOc(g.key), valor: g.qtd, cor: g.key === '—' ? css('--muted') : PAL[i % PAL.length], g: g }; }),
        F.codoc, function (d) { definirFiltro('codoc', d.key, true); });
    } else {
      $('#baixaInfo').textContent = 'Tipo de baixa';
      var gBx = agrupar(FC, null, 'baixa').sort(function (a, b) { return b.qtd - a.qtd; });
      listaStatus(gBx.map(function (g) { return { key: g.key, rotulo: titulo(g.key), valor: g.qtd, cor: corBaixa(g.key), g: g }; }),
        F.status, function (d) { definirFiltro('status', d.key, true); });
    }

    /* ---------- 9. Ocorrências dos CT-e aguardando embarque (barras horizontais) ---------- */
    var agu = FC.filter(function (o) { return o.aguardando; });
    var porOc = agrupar(agu, null, 'ocorr').sort(function (a, b) { return b.qtd - a.qtd || b.fat - a.fat; }).slice(0, 10);   // as 10 maiores (a tabela traz todas)
    $('#ocorInfo').textContent = nInt(agu.length) + ' CT-e ainda na origem (sem manifesto)';
    if (!porOc.length) { document.getElementById('boxOcor').style.height = ''; msgNoGrafico('chOcor', 'Nenhum CT-e aguardando embarque nos filtros selecionados'); }
    else {
      msgNoGrafico('chOcor', '');
      alturaRanking('boxOcor', porOc.length);
      var opO = opcoesHBar(nInt, largRot('chOcor'), function (i) { definirFiltro('ocorr', porOc[i].key, true); });
      opO.plugins.tooltip = { callbacks: { title: function (it) { return nomeOcorr(porOc[it[0].dataIndex].key); }, label: function (c) { return ' ' + nInt(c.parsed.x) + ' CT-e aguardando embarque'; },
        footer: function (it) { var g = porOc[it[0].dataIndex]; return ['Faturamento: ' + brl(g.fat), '% dos aguardando: ' + pct(agu.length ? g.qtd / agu.length : 0), '', 'Clique para filtrar']; } } };
      montar('chOcor', { type: 'bar', data: { labels: porOc.map(function (g) { return nomeOcorr(g.key); }),
        datasets: [{ data: porOc.map(function (g) { return g.qtd; }), backgroundColor: coresDestaque(porOc, F.ocorr, css('--warn')), borderRadius: 4, borderSkipped: false, barThickness: 16 }] },
        options: opO, valores: { fmt: nInt, folga: 36 } }, recriar);
    }
  }
  function redesenharGraficos(recriar) { if (RAW && (CTE.length || COT.length)) { desenharGraficos(recriar); desenharMeta(); } }

  // vendedoras com faturamento, cotações e % da meta do mês de referência
  // vendedoras do ranking: só as do metas.csv do mês (na falta de lista no mês, todas). Os totais do painel não mudam.
  function gruposVendedoras() {
    var ym = mesReferencia(), M = METAS[ym] || {}, noMes = {}, lista = LISTA_CARDS[ym] || [];
    agrupar(cteDoMes(ym), null, 'vendedora').forEach(function (g) { noMes[g.key] = g.fat; });
    return agrupar(FC, FQ, 'vendedora', 'vendedora').filter(function (g) { return !lista.length || lista.indexOf(g.key) >= 0; }).map(function (g) {
      g.meta = M[g.key] || 0; g.realMes = noMes[g.key] || 0; g.pMeta = g.meta ? g.realMes / g.meta : 0; return g;
    });
  }
  // série do gráfico "Ritmo da meta": acumulado por dia x meta proporcional aos dias úteis x projeção
  function dadosRitmo() {
    var ym = mesReferencia(), du = diasUteis(ym), meta = metaDe(ym, F.vendedora), lista = cteDoMes(ym), hoje = hojeIso();
    var porDiaG = {}; agrupar(lista, null, 'dt').forEach(function (g) { porDiaG[g.key] = g; });
    var ultimoDado = lista.reduce(function (m, o) { return o.dt > m ? o.dt : m; }, '');
    var corte = ym === hoje.slice(0, 7) ? (hoje > ultimoDado ? hoje : ultimoDado) : (ym < hoje.slice(0, 7) ? du.dias[du.dias.length - 1].iso : ultimoDado);
    var s = 0, uteis = 0, real = [], ideal = [], proj = [];
    du.dias.forEach(function (d) {
      if (d.util) uteis++;
      s += porDiaG[d.iso] ? porDiaG[d.iso].fat : 0;
      real.push(corte && d.iso <= corte ? s : null);
      ideal.push(du.total ? meta * uteis / du.total : 0);
    });
    // projeção: do último dia com dado até o fim do mês, no ritmo médio por dia útil
    var iCorte = du.dias.map(function (d) { return d.iso; }).indexOf(corte), uteisAteCorte = 0;
    du.dias.forEach(function (d, i) { if (d.util && i <= iCorte) uteisAteCorte++; });
    var ritmo = uteisAteCorte ? (real[iCorte] || 0) / uteisAteCorte : 0, u2 = uteisAteCorte;
    du.dias.forEach(function (d, i) {
      if (iCorte < 0 || i < iCorte || ym < hoje.slice(0, 7)) { proj.push(null); return; }
      if (i > iCorte && d.util) u2++;
      proj.push(i === iCorte ? real[iCorte] : ritmo * u2);
    });
    return { ym: ym, meta: meta, dias: du.dias, porDia: porDiaG, real: real, ideal: ideal, proj: proj, temDados: lista.length > 0,
             rot: du.dias.map(function (d) { return d.iso.slice(8, 10) + '/' + d.iso.slice(5, 7); }) };
  }

  /* ===================================================================
     GRÁFICO EM FORMA DE TABELA + EXCEL + PDF
     TABELAS[id do gráfico] = { cols, linhas, total, filtro, chaves }
     Tipos de coluna: t = texto | i = inteiro | m = dinheiro | p = percentual | n = decimal
     A tabela traz TODOS os itens (não só os 10 primeiros do gráfico).
     =================================================================== */
  var TABELAS = {}, MODO_TABELA = {};
  var COLS_G = function (rot) {
    return [{ l: rot, t: 't' }, { l: 'CT-e', t: 'i' }, { l: 'Faturamento', t: 'm' }, { l: 'ICMS', t: 'm' }, { l: 'Frete sem ICMS', t: 'm' },
            { l: 'Ticket médio', t: 'm' }, { l: 'Cotações', t: 'i' }, { l: 'Frete inicial', t: 'm' }, { l: 'Frete atual (valor cotado)', t: 'm' },
            { l: 'Variação do frete', t: 'p' }, { l: 'Fechamento (CT-e ÷ cotações)', t: 'p' }];
  };
  var LINHA_G = function (nome, g) { return [nome, g.qtd, g.fat, g.icms, g.fsi, g.ticket, g.cot, g.valorIni, g.valor, g.varFrete, g.fechPct]; };
  function montarTabelasDosGraficos() {
    var T = {}, tot = total(FC, FQ), totalG = LINHA_G('Total', tot);
    var chaves = function (gs) { return gs.map(function (g) { return g.key; }); };
    var ordenar = function (gs, k) { return gs.map(function (g) { g.v = valorMetrica(g, k); return g; }).sort(function (a, b) { return b.v - a.v; }); };
    // 1. evolução (segue Valores/Quantidade e No período/Acumulado)
    var E = serieEvolucao(), ac = novoGrupo('a'), acumula = UI.evo === 'a';
    T.chEvo = { cols: COLS_G(E.diario ? 'Dia' : 'Mês').map(function (c, i) { return i === 0 && acumula ? { l: c.l + ' (acumulado)', t: 't' } : c; }),
      filtro: E.diario ? 'dia' : 'ym', chaves: chaves(E.grupos), total: acumula ? null : totalG,
      linhas: E.grupos.map(function (g) {
        var nome = E.diario ? fDate(g.key) : nomeMes(g.key);
        if (!acumula) return LINHA_G(nome, g);
        Object.keys(ac).forEach(function (k) { if (typeof ac[k] === 'number') ac[k] += g[k] || 0; });
        fechar(ac); return LINHA_G(nome, ac);
      }) };
    // 2. por UF
    var porUF = ordenar(agrupar(FC, FQ, 'uf', 'uf').filter(function (g) { return g.key !== '—'; }), UI.uf), cu = COLS_G('UF'); cu.splice(1, 0, { l: 'Estado', t: 't' });
    T.chUF = { cols: cu, filtro: 'uf', chaves: chaves(porUF), linhas: porUF.map(function (g) { var l = LINHA_G(g.key, g); l.splice(1, 0, MAPA_NOMES[g.key] || ''); return l; }),
               total: [totalG[0], ''].concat(totalG.slice(1)) };
    // 3. ritmo da meta
    var R = dadosRitmo();
    T.chRitmo = { cols: [{ l: 'Dia', t: 't' }, { l: 'Dia útil', t: 't' }, { l: 'Faturamento no dia', t: 'm' }, { l: 'CT-e válidos no dia', t: 'i' }, { l: 'Acumulado', t: 'm' }, { l: 'Ritmo necessário (meta)', t: 'm' }, { l: 'Diferença', t: 'm' }],
      filtro: 'dia', chaves: R.dias.map(function (d) { return d.iso; }),
      linhas: R.dias.map(function (d, i) { var g = R.porDia[d.iso]; return [fDate(d.iso), d.util ? 'Sim' : 'Não', g ? g.fat : 0, g ? g.val : 0, R.real[i], R.meta ? R.ideal[i] : null, R.meta && R.real[i] != null ? R.real[i] - R.ideal[i] : null]; }),
      total: ['Total do mês', '', sum(Object.keys(R.porDia), function (k) { return R.porDia[k].fat; }), sum(Object.keys(R.porDia), function (k) { return R.porDia[k].val; }), null, R.meta || null, null] };
    // 4. coleta x armazém
    var gOr = agrupar(FC, null, 'origem').sort(function (a, b) { return b.fat - a.fat; }), tC = total(FC);
    T.chOrigem = { cols: [{ l: 'Origem da carga (coluna E)', t: 't' }, { l: 'CT-e', t: 'i' }, { l: 'Faturamento', t: 'm' }, { l: '% do faturamento', t: 'p' }, { l: 'Ticket médio', t: 'm' }],
      filtro: 'origem', chaves: chaves(gOr), linhas: gOr.map(function (g) { return [ORIGENS[g.key], g.qtd, g.fat, tC.fat ? g.fat / tC.fat : 0, g.ticket]; }),
      total: ['Total', tC.qtd, tC.fat, tC.fat ? 1 : 0, tC.ticket] };
    // 5. vendedoras (com meta)
    var gv = ordenar(gruposVendedoras(), UI.vend === 'meta' ? 'fat' : UI.vend), cv = COLS_G('Vendedora');
    cv.push({ l: 'Meta ' + nomeMes(mesReferencia()), t: 'm' }, { l: 'Faturado no mês', t: 'm' }, { l: '% da meta', t: 'p' });
    T.chVend = { cols: cv, filtro: 'vendedora', chaves: chaves(gv),
      linhas: gv.map(function (g) { return LINHA_G(NOMES[g.key] || g.key, g).concat([g.meta || null, g.realMes, g.meta ? g.pMeta : null]); }),
      // total = soma só das vendedoras da lista (se o metas.csv tiver lista no mês)
      total: (function () { var t = novoGrupo('Total'); gv.forEach(function (g) { Object.keys(t).forEach(function (k) { if (k !== 'key') t[k] += g[k] || 0; }); }); fechar(t);
        return LINHA_G('Total', t).concat([metaDe(mesReferencia(), F.vendedora) || null, sum(gv, function (g) { return g.realMes; }), null]); })() };
    // 6. clientes (todos)
    var gc = ordenar(agrupar(FC, FQ, 'cliente', 'cliente').filter(function (g) { return g.key !== '—'; }), UI.cli);
    T.chCli = { cols: COLS_G('Cliente pagador'), filtro: 'cliente', chaves: chaves(gc), linhas: gc.map(function (g) { return LINHA_G(g.key, g); }), total: totalG };
    // 7. situação das cotações
    var gs2 = agrupar([], FQ, null, 'situacao').sort(function (a, b) { return b.cot - a.cot; }), tq = total([], FQ);
    var linSit = function (nome, g, p) { var d = desconto(g); return [nome, g.cot, p, g.valorIni, g.valor, d.r, d.p]; };
    T.chSit = { cols: [{ l: 'Situação', t: 't' }, { l: 'Cotações', t: 'i' }, { l: '% das cotações', t: 'p' }, { l: 'Valor inicial (Proposta Inicial)', t: 'm' },
                       { l: 'Valor atual (Proposta Atual)', t: 'm' }, { l: 'Desconto R$ (inicial − atual)', t: 'm' }, { l: 'Desconto %', t: 'p' }],
      filtro: 'situacao', chaves: chaves(gs2), linhas: gs2.map(function (g) { return linSit(g.key, g, tq.cot ? g.cot / tq.cot : 0); }), total: linSit('Total', tq, tq.cot ? 1 : 0) };
    // 8. status dos CT-e
    var gb = agrupar(FC, null, UI.baixa === 'ocor' ? 'codoc' : 'baixa').sort(function (a, b) { return b.qtd - a.qtd; });
    if (UI.baixa === 'ocor') T.chBaixa = { cols: [{ l: 'Código da Ultima Ocorrencia - descrição', t: 't' }, { l: 'CT-e', t: 'i' }, { l: '% dos CT-e', t: 'p' }, { l: 'Faturamento', t: 'm' }],
      filtro: 'codoc', chaves: chaves(gb), linhas: gb.map(function (g) { return [nomeCodOc(g.key), g.qtd, tC.qtd ? g.qtd / tC.qtd : 0, g.fat]; }), total: ['Total', tC.qtd, tC.qtd ? 1 : 0, tC.fat] };
    else T.chBaixa = { cols: [{ l: 'Tipo de baixa', t: 't' }, { l: 'CT-e', t: 'i' }, { l: '% dos CT-e', t: 'p' }, { l: 'Faturamento', t: 'm' }],
      filtro: 'status', chaves: chaves(gb), linhas: gb.map(function (g) { return [titulo(g.key), g.qtd, tC.qtd ? g.qtd / tC.qtd : 0, g.fat]; }), total: ['Total', tC.qtd, tC.qtd ? 1 : 0, tC.fat] };
    // 9. ocorrências dos CT-e aguardando embarque
    var agu = FC.filter(function (o) { return o.aguardando; }), go = agrupar(agu, null, 'ocorr').sort(function (a, b) { return b.qtd - a.qtd; }), ta = total(agu);
    T.chOcor = { cols: [{ l: 'Ocorrência (Complemento da Ultima Ocorrencia)', t: 't' }, { l: 'CT-e aguardando embarque', t: 'i' }, { l: '% dos aguardando', t: 'p' }, { l: 'Faturamento', t: 'm' }],
      filtro: 'ocorr', chaves: chaves(go), linhas: go.map(function (g) { return [nomeOcorr(g.key), g.qtd, agu.length ? g.qtd / agu.length : 0, g.fat]; }), total: ['Total', agu.length, agu.length ? 1 : 0, ta.fat] };
    TABELAS = T;
    Object.keys(MODO_TABELA).forEach(function (id) { if (MODO_TABELA[id]) desenharTabelaDoGrafico(id); });
  }
  function fmtCelula(v, t) {
    if (v === null || v === undefined || v === '') return t === 't' ? '' : '—';
    return t === 'm' ? brl(v) : t === 'p' ? pct(v) : t === 'i' ? nInt(v) : t === 'n' ? fDec.format(v) : String(v);
  }
  function areaDoGrafico(card) { return card.querySelector('.donut-wrap') || card.querySelector('.destinos') || card.querySelector('.chart-box'); }
  function htmlTabela(T, comFiltro) {
    var linha = function (cels, tag, extra) {
      return '<tr' + (extra || '') + '>' + cels.map(function (v, i) { var c = T.cols[i]; return '<' + tag + (c.t === 't' ? '' : ' class="r num"') + '>' + esc(tag === 'th' ? v : fmtCelula(v, c.t)) + '</' + tag + '>'; }).join('') + '</tr>';
    };
    return '<table><thead>' + linha(T.cols.map(function (c) { return c.l; }), 'th') + '</thead><tbody>' +
      (T.linhas.length ? T.linhas.map(function (l, i) {
        return linha(l, 'td', comFiltro && T.filtro ? ' data-i="' + i + '" tabindex="0" title="Clique para filtrar"' + (F[T.filtro] === T.chaves[i] ? ' class="on"' : '') : '');
      }).join('') : '<tr><td colspan="' + T.cols.length + '" class="empty">Sem dados para os filtros selecionados</td></tr>') +
      '</tbody>' + (T.total && T.linhas.length ? '<tfoot>' + linha(T.total, 'td') + '</tfoot>' : '') + '</table>';
  }
  function desenharTabelaDoGrafico(id) {
    var canvas = document.getElementById(id), T = TABELAS[id]; if (!canvas || !T) return;
    var box = canvas.closest('.card').querySelector('.g-tab'); if (!box) return;
    var rolagem = box.scrollTop; box.innerHTML = htmlTabela(T, true); box.scrollTop = rolagem;
  }
  function alternarTabela(card, id, botao) {
    var area = areaDoGrafico(card), box = card.querySelector('.g-tab');
    if (!box) {
      box = document.createElement('div'); box.className = 'g-tab'; box.hidden = true;
      area.parentNode.insertBefore(box, area.nextSibling);
      var filtrar = function (e) {
        var tr = e.target.closest('tr[data-i]'), T = TABELAS[id]; if (!tr || !T || !T.filtro) return;
        definirFiltro(T.filtro, T.chaves[+tr.getAttribute('data-i')], true);
      };
      box.addEventListener('click', filtrar);
      box.addEventListener('keydown', function (e) { if (e.key === 'Enter') filtrar(e); });
    }
    var ligar = !MODO_TABELA[id];
    if (ligar) box.style.height = Math.max(240, area.offsetHeight) + 'px';
    MODO_TABELA[id] = ligar; area.hidden = ligar; box.hidden = !ligar;
    card.classList.toggle('modo-tabela', ligar);
    botao.innerHTML = ligar ? ICO_GRAF : ICO_TAB;
    botao.title = ligar ? 'Ver como gráfico' : 'Ver como tabela'; botao.setAttribute('aria-label', botao.title); botao.setAttribute('aria-pressed', ligar);
    if (ligar) desenharTabelaDoGrafico(id); else redesenharGraficos(false);
  }
  function descreverFiltros() {
    var partes = Object.keys(F).filter(function (k) { return F[k]; }).map(function (k) { return F_ROTULO[k] + ': ' + rotuloFiltro(k, F[k]); });
    return partes.length ? 'Filtros: ' + partes.join(' | ') : 'Sem filtros (todos os registros)';
  }
  var REGRA_SOMAS = 'CT-e e cotações cancelados ficam fora do painel (não entram nas quantidades nem nos valores). Faturamento = Valor do Frete; Frete inicial = Proposta Inicial; Frete atual (valor cotado) = Proposta Atual; Fechamento = CT-e ÷ cotações.';

  /* ---------- EXCEL: arquivo .xlsx de verdade, sem biblioteca (ZIP "stored" com XML) ---------- */
  var CRC_TAB = null;
  function crc32(bytes) {
    if (!CRC_TAB) { CRC_TAB = []; for (var n = 0; n < 256; n++) { var c = n; for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; CRC_TAB[n] = c >>> 0; } }
    var crc = 0xFFFFFFFF; for (var i = 0; i < bytes.length; i++) crc = CRC_TAB[(crc ^ bytes[i]) & 0xFF] ^ (crc >>> 8);
    return (crc ^ 0xFFFFFFFF) >>> 0;
  }
  function montarZip(arquivos) {
    var enc = new TextEncoder(), partes = [], central = [], pos = 0, d = new Date();
    var hora = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1), data = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
    var cab = function (assinatura, tam) { var b = new DataView(new ArrayBuffer(tam)); b.setUint32(0, assinatura, true); return b; };
    arquivos.forEach(function (a) {
      var nome = enc.encode(a.nome), dados = enc.encode(a.texto), crc = crc32(dados);
      var l = cab(0x04034b50, 30);
      l.setUint16(4, 20, true); l.setUint16(6, 0x0800, true); l.setUint16(8, 0, true); l.setUint16(10, hora, true); l.setUint16(12, data, true);
      l.setUint32(14, crc, true); l.setUint32(18, dados.length, true); l.setUint32(22, dados.length, true); l.setUint16(26, nome.length, true);
      var c = cab(0x02014b50, 46);
      c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true); c.setUint16(10, 0, true); c.setUint16(12, hora, true); c.setUint16(14, data, true);
      c.setUint32(16, crc, true); c.setUint32(20, dados.length, true); c.setUint32(24, dados.length, true); c.setUint16(28, nome.length, true); c.setUint32(42, pos, true);
      partes.push(l.buffer, nome, dados); central.push(c.buffer, nome);
      pos += 30 + nome.length + dados.length;
    });
    var tamCentral = central.reduce(function (s, b) { return s + b.byteLength; }, 0);
    var fim = cab(0x06054b50, 22);
    fim.setUint16(8, arquivos.length, true); fim.setUint16(10, arquivos.length, true); fim.setUint32(12, tamCentral, true); fim.setUint32(16, pos, true);
    return new Blob(partes.concat(central, [fim.buffer]), { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }
  function xmlEsc(t) { return String(t == null ? '' : t).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
  function letraColuna(i) { var s = ''; i++; while (i > 0) { var r = (i - 1) % 26; s = String.fromCharCode(65 + r) + s; i = Math.floor((i - 1) / 26); } return s; }
  // planilha: linha 1 = título | 2 = filtros ligados | 3 = regra das somas | 5 = cabeçalho (congelado) | dados | TOTAL em negrito
  function gerarXlsx(tituloPlan, T) {
    var estiloDe = { m: 2, p: 3, i: 4, n: 5, t: 0 }, linhas = [], n = 0;
    var celTexto = function (col, lin, v, est) { return '<c r="' + letraColuna(col) + lin + '" t="inlineStr"' + (est ? ' s="' + est + '"' : '') + '><is><t xml:space="preserve">' + xmlEsc(v) + '</t></is></c>'; };
    var celNum = function (col, lin, v, est) { return '<c r="' + letraColuna(col) + lin + '" s="' + est + '"><v>' + (isFinite(v) ? +(+v).toFixed(6) : 0) + '</v></c>'; };
    var linhaDados = function (vals, negrito) {
      n++; return '<row r="' + n + '">' + vals.map(function (v, i) {
        var t = T.cols[i].t;
        if (t === 't' || v === null || v === undefined || v === '') return celTexto(i, n, v == null ? '' : v, negrito ? 6 : 0);
        return celNum(i, n, v, estiloDe[t] + (negrito ? 5 : 0));
      }).join('') + '</row>';
    };
    n++; linhas.push('<row r="1">' + celTexto(0, 1, 'Ostel Transportes - Dashboard Vendas - ' + tituloPlan, 11) + '</row>');
    n++; linhas.push('<row r="2">' + celTexto(0, 2, descreverFiltros() + ' | Gerado em ' + new Date().toLocaleString('pt-BR')) + '</row>');
    n++; linhas.push('<row r="3">' + celTexto(0, 3, REGRA_SOMAS) + '</row>');
    n++;
    n++; linhas.push('<row r="5">' + T.cols.map(function (c, i) { return celTexto(i, 5, c.l, 1); }).join('') + '</row>');
    T.linhas.forEach(function (l) { linhas.push(linhaDados(l, false)); });
    if (T.total && T.linhas.length) linhas.push(linhaDados(T.total, true));
    var larguras = T.cols.map(function (c, i) {
      var m = c.l.length; T.linhas.forEach(function (l) { var t = c.t === 't' ? String(l[i] == null ? '' : l[i]).length : (c.t === 'm' ? 16 : 10); if (t > m) m = t; });
      return '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + Math.min(60, m + 3) + '" customWidth="1"/>';
    }).join('');
    var X = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n', NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    var aba = xmlEsc(tituloPlan.replace(/[\[\]:*?\/\\]/g, ' ').slice(0, 31));
    var fonte = function (negrito, cor, tam) { return '<font>' + (negrito ? '<b/>' : '') + '<sz val="' + (tam || 11) + '"/>' + (cor ? '<color rgb="' + cor + '"/>' : '') + '<name val="Calibri"/></font>'; };
    var xf = function (fmt, fonteId, preench) { return '<xf numFmtId="' + fmt + '" fontId="' + fonteId + '" fillId="' + (preench || 0) + '" borderId="0" xfId="0"' + (fmt ? ' applyNumberFormat="1"' : '') + (fonteId ? ' applyFont="1"' : '') + (preench ? ' applyFill="1"' : '') + '/>'; };
    return montarZip([
      { nome: '[Content_Types].xml', texto: X + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>' },
      { nome: '_rels/.rels', texto: X + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>' },
      { nome: 'xl/workbook.xml', texto: X + '<workbook xmlns="' + NS + '" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="' + aba + '" sheetId="1" r:id="rId1"/></sheets></workbook>' },
      { nome: 'xl/_rels/workbook.xml.rels', texto: X + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>' },
      { nome: 'xl/styles.xml', texto: X + '<styleSheet xmlns="' + NS + '">' +
          '<numFmts count="2"><numFmt numFmtId="164" formatCode="&quot;R$&quot;\\ #,##0.00;[Red]\\-&quot;R$&quot;\\ #,##0.00"/><numFmt numFmtId="165" formatCode="0.0%"/></numFmts>' +
          '<fonts count="4">' + fonte(false) + fonte(true) + fonte(true, 'FFFFFFFF') + fonte(true, null, 14) + '</fonts>' +
          '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF0A2233"/><bgColor indexed="64"/></patternFill></fill></fills>' +
          '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
          '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
          '<cellXfs count="12">' + xf(0, 0) + xf(0, 2, 2) + xf(164, 0) + xf(165, 0) + xf(3, 0) + xf(4, 0) + xf(0, 1) + xf(164, 1) + xf(165, 1) + xf(3, 1) + xf(4, 1) + xf(0, 3) + '</cellXfs>' +
          '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>' },
      { nome: 'xl/worksheets/sheet1.xml', texto: X + '<worksheet xmlns="' + NS + '"><sheetViews><sheetView workbookViewId="0"><pane ySplit="5" topLeftCell="A6" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
          '<cols>' + larguras + '</cols><sheetData>' + linhas.join('') + '</sheetData></worksheet>' }
    ]);
  }
  function nomeArquivo(t) { return 'ostel_' + semAcento(t).replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, ''); }
  function baixar(blob, nome) {
    var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = nome;
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 500);
    if (emVisualizador()) { avisoBloqueio(COPIA_TSV); SEM_AVISO = true; }   // não diz "baixado" se pode ter sido bloqueado
    COPIA_TSV = '';
  }
  /* v57 LINK DO ONEDRIVE: quando o painel abre DENTRO da visualização do OneDrive (ou outro site),
     o navegador bloqueia downloads e impressão. Aí o painel explica e oferece "Copiar a tabela" para colar no Excel. */
  var COPIA_TSV = '', SEM_AVISO = false;
  function emVisualizador() { try { return window.self !== window.top; } catch (e) { return true; } }
  function tsvDe(cols, linhas) {
    var cel = function (v) { if (v == null) return ''; if (typeof v === 'number') return String(+v.toFixed(6)).replace('.', ','); return String(v).replace(/[\t\r\n]+/g, ' '); };
    return [cols.map(cel).join('\t')].concat(linhas.map(function (l) { return l.map(cel).join('\t'); })).join('\r\n');
  }
  function tsvDaTabela(T) { return T ? tsvDe(T.cols.map(function (c) { return c.l; }), T.linhas.concat(T.total ? [T.total] : [])) : ''; }
  function copiarTexto(txt) {
    var ok = false, ta = document.createElement('textarea');
    ta.value = txt; ta.setAttribute('readonly', ''); ta.style.position = 'fixed'; ta.style.left = '-9999px';
    document.body.appendChild(ta); ta.select();
    try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
    ta.remove();
    if (!ok && navigator.clipboard) { navigator.clipboard.writeText(txt).then(function () { aviso('Tabela copiada: abra o Excel e cole (Ctrl+V)'); }, function () { aviso('Não foi possível copiar'); }); return; }
    aviso(ok ? 'Tabela copiada: abra o Excel e cole (Ctrl+V)' : 'Não foi possível copiar');
  }
  function avisoBloqueio(tsv) {
    abrirModal('<div class="bloq">' +
      '<h2>O download pode ter sido bloqueado</h2>' +
      '<p>O painel está aberto <b>dentro da visualização do OneDrive</b> (ou de outro site). Essa visualização não deixa baixar arquivos nem imprimir.</p>' +
      '<h3>Para baixar o Excel ou o PDF</h3>' +
      '<ol><li>No OneDrive, clique em <b>Baixar</b> (ou nos três pontinhos <b>⋯ → Baixar</b>).</li>' +
      '<li>Abra o arquivo <b>Vendas.html</b> baixado (ele abre no Chrome ou Edge).</li>' +
      '<li>Clique de novo no botão Excel ou PDF: agora funciona.</li></ol>' +
      (tsv ? '<h3>Ou copie a tabela agora</h3><p>Copia a mesma lista para colar direto no Excel (Ctrl+V).</p>' +
        '<button type="button" class="btn-copiar" id="btnCopiarTsv">Copiar a tabela</button>' : '') +
      '</div>');
    var b = $('#btnCopiarTsv'); if (b) b.onclick = function () { copiarTexto(tsv); };
  }
  function baixarExcel(id, tituloPlan) {
    var T = TABELAS[id]; if (!T || !T.linhas.length) { aviso('Sem dados para exportar'); return; }
    COPIA_TSV = tsvDaTabela(T);
    baixar(gerarXlsx(tituloPlan, T), nomeArquivo(tituloPlan) + '_' + hojeIso() + '.xlsx');
    aviso('Planilha Excel baixada com ' + nInt(T.linhas.length) + ' linhas');
  }

  /* ---------- PDF: abre a impressão do navegador com o gráfico + a tabela (escolha "Salvar como PDF") ---------- */
  function imagemDoCard(card, canvas) {
    var partes = [];
    if (!MODO_TABELA[canvas.id] && canvas.style.display !== 'none' && canvas.width) {
      var c = document.createElement('canvas'); c.width = canvas.width; c.height = canvas.height;
      var ctx = c.getContext('2d'); ctx.fillStyle = '#FFFFFF'; ctx.fillRect(0, 0, c.width, c.height); ctx.drawImage(canvas, 0, 0);
      partes.push('<img src="' + c.toDataURL('image/png') + '" alt="">');
    }
    var svg = card.querySelector('.mapa svg');
    if (svg && !MODO_TABELA[canvas.id]) {
      var clone = svg.cloneNode(true);
      $$('text', clone).forEach(function (t) { t.setAttribute('fill', t.getAttribute('class') === 'branco' ? '#fff' : '#0A2233'); t.setAttribute('font-size', '19'); t.setAttribute('font-weight', '700'); t.setAttribute('text-anchor', 'middle'); t.setAttribute('dominant-baseline', 'central'); t.setAttribute('font-family', 'Segoe UI, Arial'); });
      $$('path', clone).forEach(function (p) { p.setAttribute('stroke', '#fff'); p.setAttribute('stroke-width', '1.2'); });
      partes.unshift('<div class="mapa">' + new XMLSerializer().serializeToString(clone) + '</div>');
    }
    return partes.join('');
  }
  function baixarPDF(card, canvas, tituloPlan) {
    var T = TABELAS[canvas.id];
    if (emVisualizador()) { avisoBloqueio(tsvDaTabela(T)); return; }
    var estilo = 'body{font-family:"Segoe UI",Arial,sans-serif;color:#0A1C29;margin:18px}h1{font-family:Bahnschrift,"Segoe UI",Arial;font-weight:600;font-size:20px;margin:0 0 4px}' +
      'p{margin:2px 0;font-size:11px;color:#43586A}img{display:block;max-width:100%;max-height:330px;margin:12px auto;border:1px solid #D9E2E9;border-radius:8px}' +
      '.mapa svg{display:block;height:300px;margin:12px auto}table{border-collapse:collapse;width:100%;font-size:10.5px;margin-top:12px}' +
      'th{background:#0A2233;color:#fff;text-align:left;padding:5px 7px}td{padding:4px 7px;border-bottom:1px solid #E9EFF3}.r{text-align:right;font-variant-numeric:tabular-nums}' +
      'tfoot td{font-weight:700;border-top:2px solid #0492D0}td:first-child{white-space:nowrap}@page{size:A4 landscape;margin:12mm}';
    var html = '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="utf-8"><title>' + esc('Ostel - ' + tituloPlan) + '</title><style>' + estilo + '</style></head><body>' +
      '<h1>Ostel Transportes · Dashboard Vendas · ' + esc(tituloPlan) + '</h1><p>' + esc(descreverFiltros()) + ' · Gerado em ' + esc(new Date().toLocaleString('pt-BR')) + '</p><p>' + esc(REGRA_SOMAS) + '</p>' +
      imagemDoCard(card, canvas) + (T ? htmlTabela(T, false) : '') + '</body></html>';
    var fr = document.createElement('iframe'); fr.className = 'pdf-frame'; fr.setAttribute('aria-hidden', 'true');
    document.body.appendChild(fr);
    var doc = fr.contentWindow.document; doc.open(); doc.write(html); doc.close();
    setTimeout(function () {
      try { fr.contentWindow.focus(); fr.contentWindow.print(); } catch (e) { aviso('Não foi possível abrir a impressão'); }
      setTimeout(function () { fr.remove(); }, 60000);
    }, 350);
    aviso('Na janela de impressão, escolha "Salvar como PDF"');
  }

  /* ===================================================================
     MAPA DO BRASIL POR ESTADO (SVG próprio, contornos em assets\mapa-brasil.js)
     Um só tom de azul: claro = pouco, escuro = muito (botões Faturamento | CT-e | Cotações)
     =================================================================== */
  var MAPA_PRONTO = false, MAPA_GRUPOS = {}, MAPA_NOMES = {};
  function rgbDe(hex) { hex = hex.replace('#', ''); if (hex.length === 3) hex = hex.replace(/(.)/g, '$1$1'); return [parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16)]; }
  function corEscala(cores, t) {
    t = Math.max(0, Math.min(1, t)) * (cores.length - 1);
    var i = Math.min(cores.length - 2, Math.floor(t)), f = t - i, a = cores[i], b = cores[i + 1];
    return [0, 1, 2].map(function (k) { return Math.round(a[k] + (b[k] - a[k]) * f); });
  }
  function montarMapa() {
    var box = $('#mapaUF'), M = window.OSTEL_MAPA;
    if (!box) return false;
    if (!M || !M.ufs) { box.innerHTML = '<div class="chart-msg">Mapa indisponível: falta o arquivo assets\\mapa-brasil.js.</div>'; return false; }
    if (MAPA_PRONTO) return true;
    M.ufs.forEach(function (u) { MAPA_NOMES[u.uf] = u.nome; });
    box.innerHTML =
      '<svg viewBox="' + M.viewBox + '" role="group" aria-label="Mapa do Brasil por estado de entrega">' +
        '<g class="m-ufs">' + M.ufs.map(function (u) { return '<path d="' + u.d + '" data-uf="' + u.uf + '" tabindex="0" role="button"></path>'; }).join('') + '</g>' +
        '<g class="m-rot" aria-hidden="true">' + M.ufs.filter(function (u) { return u.rot; }).map(function (u) { return '<text x="' + u.x + '" y="' + u.y + '" data-uf="' + u.uf + '">' + u.uf + '</text>'; }).join('') + '</g>' +
      '</svg>' +
      '<div class="mapa-leg"><span class="t" id="mapaLegT"></span><span class="num">0</span><i></i><span class="num" id="mapaLegMax"></span></div>' +
      '<div class="mapa-tip" id="mapaTip" hidden></div>';
    var tip = $('#mapaTip');
    function mostrar(p, x, y) {
      var uf = p.getAttribute('data-uf'), g = MAPA_GRUPOS[uf];
      tip.innerHTML = '<b>' + esc(MAPA_NOMES[uf]) + ' (' + uf + ')</b>' + (g ? rodapeGrupo(g).filter(Boolean).map(esc).join('<br>') : 'Sem vendas nem cotações neste filtro');
      tip.hidden = false;
      var w = tip.offsetWidth, h = tip.offsetHeight, bw = box.clientWidth;
      tip.style.left = Math.max(0, x + 14 + w > bw ? x - w - 14 : x + 14) + 'px';
      tip.style.top = Math.max(0, y - h - 10 < 0 ? y + 16 : y - h - 10) + 'px';
    }
    function esconder() { tip.hidden = true; }
    var ufs = box.querySelector('.m-ufs');
    ufs.addEventListener('pointermove', function (e) { var p = e.target.closest('path'); if (!p) return esconder(); var r = box.getBoundingClientRect(); mostrar(p, e.clientX - r.left, e.clientY - r.top); });
    ufs.addEventListener('pointerover', function (e) { var p = e.target.closest('path'); if (p && p !== ufs.lastChild && e.pointerType === 'mouse') ufs.appendChild(p); });
    ufs.addEventListener('pointerleave', esconder);
    ufs.addEventListener('focusin', function (e) { var p = e.target.closest('path'); if (!p) return; var r = box.getBoundingClientRect(), b = p.getBoundingClientRect(); mostrar(p, b.left + b.width / 2 - r.left, b.top + b.height / 2 - r.top); });
    ufs.addEventListener('focusout', esconder);
    function filtrar(p) { var uf = p.getAttribute('data-uf'); if (MAPA_GRUPOS[uf] || F.uf === uf) { esconder(); definirFiltro('uf', uf, true); } }
    ufs.addEventListener('click', function (e) { var p = e.target.closest('path'); if (p) filtrar(p); });
    ufs.addEventListener('keydown', function (e) { if (e.key !== 'Enter' && e.key !== ' ') return; var p = e.target.closest('path'); if (p) { e.preventDefault(); filtrar(p); } });
    MAPA_PRONTO = true; return true;
  }
  function desenharMapa() {
    if (!montarMapa()) return;
    var k = UI.uf, fmt = fmtMetricaCheia(k), max = 0, grupos = {};
    // com um estado filtrado, os outros continuam visíveis (mais apagados) para trocar com um clique
    var C = F.uf ? filtrarCte(F, 'uf') : FC, Q = F.uf ? filtrarCot(F, 'uf') : FQ;
    agrupar(C, Q, 'uf', 'uf').forEach(function (g) { g.v = valorMetrica(g, k); if (g.v > 0) { grupos[g.key] = g; if (g.v > max) max = g.v; } });
    MAPA_GRUPOS = grupos;
    var cores = [1, 2, 3, 4, 5].map(function (i) { return rgbDe(css('--map-' + i)); }), vazio = css('--map-0'), claro = {};
    $$('#mapaUF path').forEach(function (p) {
      var uf = p.getAttribute('data-uf'), g = grupos[uf];
      var c = g ? corEscala(cores, max > 0 ? g.v / max : 0) : null;
      p.style.fill = c ? 'rgb(' + c.join(',') + ')' : vazio;
      claro[uf] = !c || (0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2]) / 255 > 0.55;
      p.classList.toggle('vazio', !g); p.classList.toggle('on', F.uf === uf); p.classList.toggle('dim', !!F.uf && F.uf !== uf);
      p.setAttribute('aria-label', MAPA_NOMES[uf] + ': ' + (g ? fmt(g.v) : 'sem dados neste filtro'));
    });
    var sel = F.uf && $('#mapaUF path[data-uf="' + F.uf + '"]'); if (sel) sel.parentNode.appendChild(sel);
    $$('#mapaUF text').forEach(function (t) { var uf = t.getAttribute('data-uf'); t.setAttribute('class', !grupos[uf] ? 'apagado' : (claro[uf] ? 'tinta' : 'branco')); });
    $('#mapaLegT').textContent = k === 'qtd' ? 'CT-e' : k === 'cot' ? 'Cotações' : 'Faturamento';
    $('#mapaLegMax').textContent = fmt(max);
  }

  /* ===================================================================
     TABELA DE DETALHAMENTO (CT-e | Cotações)
     Colunas com os MESMOS NOMES e a MESMA ORDEM da planilha. A única
     coluna a mais é "Vendedora (painel)", logo depois da primeira,
     calculada pela regra da vendedora (coluna V > coluna BC > Ostel).
     =================================================================== */
  var TABELA_ORDEM = [], TABELA_INI = 0;
  function fonteAtual() { return UI.aba === 'cot' ? RAW && RAW.cotacoes : RAW && RAW.cte; }
  function colunasTabela() {
    var fonte = fonteAtual(); if (!fonte || !fonte.ok) return [];
    var somar = (UI.aba === 'cot' ? MAPA.somarCot : MAPA.somarCte).map(semAcento), dinheiro = (UI.aba === 'cot' ? MAPA.dinheiroCot : MAPA.dinheiroCte).map(semAcento);
    var ocultar = (UI.aba === 'cot' ? MAPA.ocultarCot : MAPA.ocultarCte).map(semAcento);
    var cols = fonte.colunas.map(function (nome, i) {
      var n = semAcento(nome);
      return { i: i, l: nome, n: n, ocorr: UI.aba !== 'cot' && i === IDX.cte.codOcorr, tipo: (fonte.tipos || '')[i] || 'a', letra: (fonte.letras || [])[i] || '', soma: somar.indexOf(n) >= 0, dinheiro: dinheiro.indexOf(n) >= 0 };
    }).filter(function (c) { return ocultar.indexOf(c.n) < 0; });
    cols.splice(1, 0, { calc: 'vend', l: 'Vendedora (painel)', tipo: 't', dica: 'Calculada pelo painel: coluna V (Tipo de Calculo), senão coluna BC (Observação), senão Ostel' });
    // v56 COTAÇÕES: PROPOSTA INICIAL − PROPOSTA ATUAL: desconto ou acréscimo, valor e % (logo depois de PROPOSTA ATUAL)
    if (UI.aba === 'cot') {
      var pa = -1; cols.forEach(function (c, k) { if (c.n === 'proposta atual') pa = k; });
      cols.splice(pa >= 0 ? pa + 1 : cols.length, 0,
        { calc: 'res', l: 'Resultado', tipo: 't', dica: 'Desconto (proposta atual menor que a inicial), Acréscimo (atual maior) ou Sem alteração' },
        { calc: 'dif', l: 'Valor do desconto / acréscimo', tipo: 'n', dinheiro: true, soma: true, dica: 'PROPOSTA INICIAL menos PROPOSTA ATUAL (em R$)' },
        { calc: 'pct', l: '% do desconto / acréscimo', tipo: 'n', dica: 'Diferença dividida pela PROPOSTA INICIAL' });
    }
    return cols;
  }
  function valorCol(o, c) {
    if (c.calc === 'vend') return o.vendNome + (o.vendCorrigida ? ' (corrigida)' : '');
    if (c.calc === 'dif' || c.calc === 'res' || c.calc === 'pct') {
      var ini = o.v[IDX.cot.propostaIni], atu = o.v[IDX.cot.proposta];
      if (typeof ini !== 'number' || typeof atu !== 'number') return null;
      var df = Math.round((ini - atu) * 100) / 100;
      if (c.calc === 'dif') return df;
      if (c.calc === 'pct') return ini ? df / ini : null;
      return df > 0 ? 'Desconto' : df < 0 ? 'Acréscimo' : 'Sem alteração';
    }
    if (c.ocorr) return o.codOcorrTxt || null;            // "62 - SAIDA DA UNIDADE"
    return o.v[c.i];
  }
  function fmtValorCol(v, c) {
    if (v == null || v === '') return '—';
    if (c.tipo === 'd') return fDataHora(v);
    if (typeof v === 'number') return c.dinheiro ? brl(v) : fNum.format(v);
    return String(v);
  }
  function linhasTabela(cols) {
    var lista = UI.aba === 'cot' ? FQ : FC, q = UI.busca.trim().toLowerCase();
    var rows = lista.filter(function (o) { return !q || o._q.indexOf(q) >= 0; });
    // botão "Indícios de duplicidade": mostra só essas linhas, com as do mesmo grupo uma embaixo da outra
    var nAlerta = rows.filter(function (o) { return o.dup || o.ind; }).length, bd = $('#btnDup');
    if (!nAlerta) UI.soDup = false;
    if (bd) {
      bd.hidden = !nAlerta; bd.classList.toggle('on', !!UI.soDup); bd.setAttribute('aria-pressed', !!UI.soDup);
      bd.querySelector('span').textContent = UI.soDup ? 'Mostrar todas' : 'Indícios de duplicidade: ' + nInt(nAlerta);
    }
    if (UI.soDup) return rows.filter(function (o) { return o.dup || o.ind; }).sort(function (a, b) { return a.dupGrupo < b.dupGrupo ? -1 : a.dupGrupo > b.dupGrupo ? 1 : (a.dt < b.dt ? -1 : 1); });
    var c = UI.ordem >= 0 ? cols[UI.ordem] : null, d = UI.dir;
    if (!c) return rows.slice().sort(function (a, b) { return a.dt < b.dt ? 1 : a.dt > b.dt ? -1 : 0; });
    return rows.slice().sort(function (a, b) {
      var x = valorCol(a, c), y = valorCol(b, c), vx = (x == null || x === ''), vy = (y == null || y === '');
      if (vx || vy) return vx === vy ? 0 : (vx ? 1 : -1);
      if (typeof x === 'string' && typeof y === 'string') return x.localeCompare(y, 'pt-BR', { numeric: true }) * d;
      return (x > y ? 1 : x < y ? -1 : 0) * d;
    });
  }
  function desenharTabela() {
    var fonte = fonteAtual(), cols = colunasTabela();
    if (!fonte || !fonte.ok) {
      $('#tbl thead').innerHTML = ''; $('#tbl tfoot').innerHTML = ''; $('#pager').innerHTML = '';
      $('#tbl tbody').innerHTML = '<tr><td class="empty">A planilha de ' + (UI.aba === 'cot' ? 'cotações' : 'CT-e') + ' não pôde ser lida' + (fonte && fonte.erro ? ': ' + esc(fonte.erro) : '') + '</td></tr>';
      $('#tblInfo').textContent = '—'; TABELA_ORDEM = []; return;
    }
    var rows = linhasTabela(cols), totalN = rows.length, paginas = Math.max(1, Math.ceil(totalN / UI.porPagina));
    if (UI.pagina > paginas) UI.pagina = paginas;
    var fatia = rows.slice((UI.pagina - 1) * UI.porPagina, UI.pagina * UI.porPagina);
    $('#tbl thead').innerHTML = '<tr>' + cols.map(function (c, j) {
      var dica = c.dica || ('Coluna ' + c.letra + ' da planilha');
      // alinhamento do título igual ao dos valores: à direita só se a coluna tem números nesta página
      var ehNum = fatia.some(function (o) { return typeof valorCol(o, c) === 'number'; });
      return '<th data-j="' + j + '" class="' + (ehNum ? 'r ' : '') + (UI.ordem === j ? 'sorted' : '') + '" scope="col" title="' + esc(dica) + '"><span class="th-t">' + esc(c.l) + '</span>' +
             '<span class="ar">' + (UI.ordem === j ? (UI.dir > 0 ? '▲' : '▼') : '↕') + '</span></th>';
    }).join('') + '</tr>';
    TABELA_ORDEM = rows; TABELA_INI = (UI.pagina - 1) * UI.porPagina;
    $('#tbl tbody').innerHTML = fatia.length ? fatia.map(function (o, i) {
      var fora = !o.soma, alerta = o.dup || o.ind;
      return '<tr class="clicavel' + (fora ? ' fora-soma' : '') + (alerta ? ' alerta-dup' : '') + (o === UI.sel ? ' sel' : '') + '" data-i="' + (TABELA_INI + i) + '" tabindex="0" title="' +
        (alerta ? (o.dup ? 'DUPLICADA: ' : 'INDÍCIO DE DUPLICIDADE (conferir): ') + esc(o.dupMotivo) + '. ' : '') +
        (fora ? 'Cancelado: entra só na contagem, fora das somas. ' : '') + 'Clique para selecionar e ver a ficha completa">' +
        cols.map(function (c, j) {
          var v = valorCol(o, c), t = fmtValorCol(v, c), dir = typeof v === 'number';
          // número da linha (1, 2, 3...) antes do primeiro valor, para contar as linhas na tela
          if ((c.calc === 'dif' || c.calc === 'pct') && typeof v === 'number') {
            var cl = v > 0 ? ' res-desc-t' : v < 0 ? ' res-acr-t' : '';
            return '<td class="r num' + cl + '">' + (c.calc === 'pct' ? pct(Math.abs(v)) : brl(Math.abs(v))) + '</td>';
          }
          if (c.calc === 'res' && v) return '<td><span class="res-b ' + (v === 'Desconto' ? 'res-desc' : v === 'Acréscimo' ? 'res-acr' : 'res-igual') + '">' + esc(v) + '</span></td>';
          return '<td' + (dir ? ' class="r num"' : '') + '>' + (j === 0 ? '<span class="ln num">' + nInt(TABELA_INI + i + 1) + '</span>' : '') + (String(t).length > 38 ? '<span class="ell" title="' + esc(t) + '">' + esc(t) + '</span>' : esc(t)) + '</td>';
        }).join('') + '</tr>';
    }).join('') : '<tr><td colspan="' + cols.length + '" class="empty">Nenhum registro encontrado</td></tr>';
    // TOTAL: todas as linhas filtradas (não só a página), só as NÃO canceladas (mesma regra dos indicadores)
    var somaveis = rows.filter(function (o) { return o.soma; });
    $('#tbl tfoot').innerHTML = totalN ? '<tr title="Somas só dos documentos não cancelados (' + nInt(somaveis.length) + ' de ' + nInt(totalN) + ')">' + cols.map(function (c, j) {
      if (j === 0) return '<td>TOTAL</td>';
      if (c.calc === 'pct' || c.calc === 'dif' || c.calc === 'res') {
        var tI = 0, tA = 0;
        somaveis.forEach(function (o) { var a = o.v[IDX.cot.propostaIni], b = o.v[IDX.cot.proposta]; if (typeof a === 'number' && typeof b === 'number') { tI += a; tA += b; } });
        var dT = tI - tA, cT = dT > 0 ? ' res-desc-t' : dT < 0 ? ' res-acr-t' : '';
        if (c.calc === 'res') return '<td>' + (dT > 0 ? 'Desconto' : dT < 0 ? 'Acréscimo' : '') + '</td>';
        return '<td class="r num' + cT + '">' + (c.calc === 'pct' ? (tI ? pct(Math.abs(dT / tI)) : '—') : brl(Math.abs(dT))) + '</td>';
      }
      if (!c.soma) return '<td></td>';
      var t = sum(somaveis, function (o) { var v = o.v[c.i]; return typeof v === 'number' ? v : 0; });
      return '<td class="r num">' + (c.dinheiro ? brl(t) : fNum.format(t)) + '</td>';
    }).join('') + '</tr>' : '';
    var nCanc = totalN - somaveis.length, nInd = rows.filter(function (o) { return o.dup || o.ind; }).length;
    if (UI.aba === 'cot') {
      var gq = total(FC, rows);
      $('#tblInfo').innerHTML = nInt(totalN) + ' cotações (sem as canceladas) · frete inicial ' + brl(gq.valorIni) + ' · frete atual ' + brl(gq.valor) +
        ' (' + (gq.varFrete > 0 ? '+' : '') + fDec.format(gq.varFrete * 100) + '%) · fechamento ' + pct(total(FC, FQ).fechPct) + (nInd ? ' · <b class="neg">' + nInt(nInd) + ' com indício de duplicidade</b>' : '');
    } else {
      var gc = total(rows);
      $('#tblInfo').innerHTML = nInt(totalN) + ' CT-e (sem os cancelados) · faturamento ' + brl(gc.fat) + ' · coletado ' + brl(gc.coleta) + ' · armazém ' + brl(gc.armazem) + (nInd ? ' · <b class="neg">' + nInt(nInd) + ' com indício de duplicidade</b>' : '');
    }
    var de = totalN ? (UI.pagina - 1) * UI.porPagina + 1 : 0, ate = Math.min(totalN, UI.pagina * UI.porPagina);
    $('#pager').innerHTML = '<div class="pg-info"><b class="pg-qtd">' + nInt(fatia.length) + '</b> ' + (fatia.length === 1 ? 'linha' : 'linhas') + ' na tela · linhas <b>' + nInt(de) + '–' + nInt(ate) + '</b> de <b>' + nInt(totalN) + '</b></div>' +
      '<div class="pg"><label for="porPag" class="pg-lbl">Linhas por página</label><select id="porPag" aria-label="Linhas por página">' + [10, 25, 50, 100].map(function (n) { return '<option' + (n === UI.porPagina ? ' selected' : '') + '>' + n + '</option>'; }).join('') + '</select>' +
      '<button data-p="ini" ' + (UI.pagina <= 1 ? 'disabled' : '') + ' aria-label="Primeira página">«</button>' +
      '<button data-p="ant" ' + (UI.pagina <= 1 ? 'disabled' : '') + ' aria-label="Página anterior">‹</button>' +
      '<span class="num">' + UI.pagina + ' / ' + paginas + '</span>' +
      '<button data-p="prox" ' + (UI.pagina >= paginas ? 'disabled' : '') + ' aria-label="Próxima página">›</button>' +
      '<button data-p="fim" ' + (UI.pagina >= paginas ? 'disabled' : '') + ' aria-label="Última página">»</button></div>';
    $('#pager').setAttribute('data-paginas', paginas);
    ajustarTabela();
  }

  // ALTURA DO CABEÇALHO FIXO (barra superior + barra de filtros no computador)
  function alturaFixa() {
    var t = $('.topbar').offsetHeight, f = $('#filters'), raiz = document.documentElement;
    var fixo = t + (f && !f.hidden && getComputedStyle(f).position === 'sticky' ? f.offsetHeight : 0);
    raiz.style.setProperty('--h-barra', t + 'px'); raiz.style.setProperty('--topo', fixo + 'px');
    return fixo;
  }
  if ('ResizeObserver' in window) new ResizeObserver(function () { ajustarTabela(); }).observe($('#filters'));
  // a área de rolagem da tabela cabe abaixo do cabeçalho: ao chegar na tabela, título e busca ficam
  // parados e só as linhas rolam
  function ajustarTabela() {
    var wrap = $('.tbl-wrap'), card = $('#cardTable');
    if (!wrap || !card || $('#conteudo').hidden) return;
    if (innerWidth < 768 || innerHeight < 600) { alturaFixa(); wrap.style.maxHeight = ''; return; }
    var topo = alturaFixa() + 16, r = card.getBoundingClientRect(), w = wrap.getBoundingClientRect();
    var acima = w.top - r.top, abaixo = r.bottom - w.bottom;
    var rodape = document.documentElement.scrollHeight - (r.bottom + window.pageYOffset);
    wrap.style.maxHeight = Math.max(260, Math.floor(innerHeight - topo - acima - abaixo - rodape)) + 'px';
  }
  $('#tbl tbody').addEventListener('click', function (e) { var tr = e.target.closest('tr[data-i]'); if (tr) abrirFicha(+tr.getAttribute('data-i')); });
  $('#tbl tbody').addEventListener('keydown', function (e) { if (e.key === 'Enter') { var tr = e.target.closest('tr[data-i]'); if (tr) abrirFicha(+tr.getAttribute('data-i')); } });
  $('#tbl thead').addEventListener('click', function (e) {
    var th = e.target.closest('th'); if (!th) return;
    var j = +th.getAttribute('data-j'), c = colunasTabela()[j];
    if (UI.ordem === j) UI.dir *= -1; else { UI.ordem = j; UI.dir = c && (c.tipo === 'd' || c.soma) ? -1 : 1; }
    desenharTabela();
  });
  $('#pager').addEventListener('click', function (e) {
    var b = e.target.closest('button'); if (!b) return;
    var p = +this.getAttribute('data-paginas'), a = b.getAttribute('data-p');
    UI.pagina = a === 'ini' ? 1 : a === 'ant' ? UI.pagina - 1 : a === 'prox' ? UI.pagina + 1 : p;
    desenharTabela();
  });
  $('#pager').addEventListener('change', function (e) { if (e.target.id === 'porPag') { UI.porPagina = +e.target.value; UI.pagina = 1; desenharTabela(); } });
  $('#btnDup').addEventListener('click', function () { UI.soDup = !UI.soDup; UI.pagina = 1; desenharTabela(); $('.tbl-wrap').scrollTop = 0; });
  $('#tblSearch').addEventListener('input', debounce(function (e) { UI.busca = e.target.value; UI.pagina = 1; desenharTabela(); }, 200));

  // EXPORTAR CSV: as linhas filtradas, todas as colunas, no formato do Excel brasileiro (; e vírgula)
  $('#btnCsv').addEventListener('click', function () {
    var cols = colunasTabela(), rows = linhasTabela(cols);
    if (!rows.length) { aviso('Sem dados para exportar'); return; }
    var q = function (v) { v = v == null ? '' : String(v); return '"' + v.replace(/"/g, '""') + '"'; };
    var linhas = [cols.map(function (c) { return q(c.l); }).concat([q('Entra nas somas')]).join(';')].concat(rows.map(function (o) {
      return cols.map(function (c) {
        var v = valorCol(o, c);
        if (v == null || v === '') return '';
        if ((c.calc === 'dif' || c.calc === 'pct') && typeof v === 'number') v = Math.abs(v);
        if (typeof v === 'number') return String(+v.toFixed(6)).replace('.', ',');
        if (c.tipo === 'd') return q(fDataHora(v));
        return q(v);
      }).concat([o.soma ? 'Sim' : 'Não (cancelado)']).join(';');
    }));
    COPIA_TSV = linhas.map(function (l) { return l.split(';').map(function (x) { return x.replace(/^"|"$/g, '').replace(/""/g, '"'); }).join('\t'); }).join('\r\n');
    baixar(new Blob(['﻿' + linhas.join('\r\n')], { type: 'text/csv;charset=utf-8' }), 'ostel_vendas_' + (UI.aba === 'cot' ? 'cotacoes' : 'cte') + '_' + hojeIso() + '.csv');
    aviso('CSV exportado com ' + nInt(rows.length) + ' linhas');
  });

  /* ===================================================================
     RODAPÉ, ROTA DO CABEÇALHO E AVISOS
     =================================================================== */
  function desenharResumo() {
    var datas = FC.concat(FQ).map(function (o) { return o.dt; }).sort();
    var periodo = datas.length ? fDate(datas[0]) + ' a ' + fDate(datas[datas.length - 1]) : '—';
    var ufs = {}, clis = {}; FC.forEach(function (o) { if (o.uf !== '—') ufs[o.uf] = 1; if (o.cliente !== '—') clis[o.cliente] = 1; });
    var nUf = Object.keys(ufs).length, nCli = Object.keys(clis).length;
    var mi = function (ic, valor, rotulo) { return '<span>' + IC[ic] + '<b class="num">' + valor + '</b><em>' + rotulo + '</em></span>'; };
    $('#heroMeta').innerHTML = mi('cal', periodo, 'período') + mi('doc', nInt(FC.length), 'CT-e') + mi('pin', nUf, 'estados') + mi('building', nInt(nCli), 'clientes');
    var mt = function (ic, valor, rotulo) { return '<span>' + IC[ic] + '<b class="num">' + valor + '</b>' + rotulo + '</span>'; };
    $('#topMeta').innerHTML = mt('cal', periodo, '') + mt('doc', nInt(FC.length), 'CT-e') + mt('pin', nUf, 'estados') + mt('building', nInt(nCli), 'clientes');
    var fonteTxt = function (f, nome) { return f && f.ok ? esc(f.arquivo) + ' (salva em ' + fmtStamp(f.arquivoModificado) + ')' : nome + ': não lida'; };
    // rodapé: versão e quando a planilha de CT-e foi salva (a origem completa fica no "title", ao passar o mouse)
    var salva = RAW.cte && RAW.cte.ok ? RAW.cte.arquivoModificado : (RAW.cotacoes && RAW.cotacoes.ok ? RAW.cotacoes.arquivoModificado : RAW.geradoEm);
    $('#footInfo').innerHTML = 'Versão ' + CONFIG.versao + ' · (salva em ' + fmtStamp(salva) + ')';
    $('#footInfo').title = 'Origem: ' + fonteTxt(RAW.cte, 'Comercial CTe') + ' e ' + fonteTxt(RAW.cotacoes, 'Comercial Cotações') + ' · dados de ' + fmtStamp(RAW.geradoEm);
    var av = (RAW.avisos || []).slice();
    var naoAchados = CORR_LISTA.filter(function (n) { return !CORR_USADAS[chaveDoc(n)]; });
    if (naoAchados.length) av.push('correcoes_vendedora.csv: CT-e não encontrado(s) na planilha: ' + naoAchados.join(', '));
    var box = $('#avisos');
    box.hidden = !av.length;
    if (av.length) { box.querySelector('summary').textContent = av.length + ' aviso(s) na leitura das planilhas'; box.querySelector('ul').innerHTML = av.map(function (a) { return '<li>' + esc(a) + '</li>'; }).join(''); }
    // uma das planilhas com problema: o painel continua com a outra
    var ruins = [['cte', 'Comercial CTe (vendas)'], ['cotacoes', 'Comercial Cotações']].filter(function (p) { return RAW[p[0]] && RAW[p[0]].ok === false; });
    var af = $('#alertaFonte'); af.hidden = !ruins.length; af.className = 'alerta atencao';
    if (ruins.length) af.innerHTML = IC.alert + '<div>' + ruins.map(function (p) { var f = RAW[p[0]]; return '<b>Planilha ' + p[1] + ' não lida:</b> ' + esc(f.erro || '') + (f.dica ? ' ' + esc(f.dica) : ''); }).join('<br>') + '</div>';
  }

  /* ===================================================================
     DESENHO GERAL
     =================================================================== */
  function desenharTudo() { desenharMeta(); desenharKPIs(); desenharVendedoras(); desenharPremios(); desenharGraficos(); desenharComplemento(); desenharAtraso(); desenharTabela(); desenharResumo(); ajustarTabela(); }
  function atualizar() {
    try { aplicarFiltros(); sincronizarFiltros(); salvarNaURL(); desenharTudo(); }
    catch (e) { console.error(e); mostrarEstado('erro', 'Ocorreu um erro ao montar o painel', '<p class="det">' + esc(e.message) + '</p><p>Execute <code>1_atualizar_agora.bat</code> e recarregue a página.</p>'); }
  }
  var larguraAnterior = innerWidth;
  window.addEventListener('resize', debounce(function () {
    if (innerWidth !== larguraAnterior) { larguraAnterior = innerWidth; redesenharGraficos(false); }
    ajustarTabela();
  }, 200));

  /* ===================================================================
     MENSAGENS AMIGÁVEIS (nunca tela em branco)
     =================================================================== */
  function mostrarEstado(tipo, tit, html) {
    $('#conteudo').hidden = true; $('#filters').hidden = true;
    var e = $('#estado'); e.hidden = false; e.className = 'estado ' + tipo;
    e.innerHTML = '<div class="box" role="alert"><div class="ic">' + (tipo === 'erro' ? IC.alert : IC.file) + '</div><h2>' + esc(tit) + '</h2>' + html +
      (EMBUTIDO ? '' : '<button class="btn primary" id="btnTentar"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M21 12a9 9 0 1 1-2.64-6.36"/><path d="M21 3v6h-6"/></svg>Tentar novamente</button>' +
      '<p class="det" style="margin-top:12px">O painel tenta de novo sozinho a cada ' + CONFIG.recarregarSegundos + ' segundos.</p>') + '</div>';
    if ($('#btnTentar')) $('#btnTentar').onclick = function () { carregarDados(true); };
    $('#heroMeta').innerHTML = ''; $('#topMeta').innerHTML = '';
  }
  function esconderEstado() { $('#estado').hidden = true; $('#conteudo').hidden = false; $('#filters').hidden = false; }
  var PASSOS = '<ol>' +
    '<li>Confira no <code>config.ps1</code> o caminho das planilhas <b>Comercial CTe</b> e <b>Comercial Cotações</b> (ou coloque os arquivos na pasta <code>dados</code>).</li>' +
    '<li>Dê dois cliques em <code>1_atualizar_agora.bat</code>.</li>' +
    '<li>Esta página se atualiza sozinha em alguns segundos.</li></ol>';

  /* ===================================================================
     CARREGAMENTO E ATUALIZAÇÃO AUTOMÁTICA (dados.js + dados_status.js)
     Usa <script> em vez de fetch para funcionar abrindo o arquivo direto
     do disco (file://), sem servidor. O parâmetro ?v= evita o cache.
     =================================================================== */
  var contagem = CONFIG.recarregarSegundos, carregando = false, ultimoErro = '';
  var EMBUTIDO = !!window.OSTEL_EMBUTIDO, DADOS_EMBUTIDOS = EMBUTIDO ? window.OSTEL_DATA : null;
  if (EMBUTIDO) document.body.classList.add('embutido');

  function carregarScript(arquivo, id, pronto) {
    var antigo = document.getElementById(id); if (antigo) antigo.remove();
    var s = document.createElement('script'); s.id = id; s.charset = 'utf-8';
    s.src = arquivo + '?v=' + Date.now();
    s.onload = function () { pronto(true); }; s.onerror = function () { pronto(false); };
    document.head.appendChild(s);
  }
  function carregarStatus(depois) {
    if (EMBUTIDO) { depois(); return; }
    var anterior = window.OSTEL_STATUS; window.OSTEL_STATUS = undefined;
    carregarScript(CONFIG.arquivoStatus, 'scriptStatus', function (ok) {
      STATUS = ok && window.OSTEL_STATUS ? window.OSTEL_STATUS : (anterior || null);
      depois();
    });
  }
  function processarDados(d, manual) {
    if (!d || typeof d !== 'object') { statusAoVivo(false, 'dados.js inválido'); if (!RAW) mostrarEstado('erro', 'O arquivo dados.js está corrompido', '<p>O arquivo existe, mas não pôde ser lido.</p>' + PASSOS); return; }
    if (d.ok === false) {
      var carimboErro = 'erro|' + d.geradoEm;
      statusAoVivo(false, 'planilhas com problema');
      if (carimboErro !== CARIMBO) {
        CARIMBO = carimboErro; RAW = null; CTE = []; COT = [];
        mostrarEstado('erro', 'Não foi possível ler as planilhas', '<p><b>' + esc(d.erro || 'Erro desconhecido') + '</b></p>' + (d.dica ? '<p>' + esc(d.dica) + '</p>' : '') + PASSOS +
          '<p class="det">Verificado em ' + fmtStamp(d.geradoEm) + '. Detalhes em <code>logs\\atualizacao.log</code>.</p>');
      }
      return;
    }
    var carimbo = d.geradoEm + '|' + (d.cte && d.cte.qtd) + '|' + (d.cotacoes && d.cotacoes.qtd) + '|' + JSON.stringify(d.metas || []).length + '|' + JSON.stringify(d.correcoes || []).length + '|' + JSON.stringify(d.ocorrencias || {}).length + '|' + JSON.stringify(d.feriados || []).length + '|' + JSON.stringify(d.premios || []).length + '|' + JSON.stringify(d.comissao || []).length;
    if (carimbo === CARIMBO) { if (manual) aviso('Nenhuma alteração desde ' + fmtStamp(d.geradoEm)); statusAoVivo(true); return; }
    var primeira = !RAW;
    NOMES = {};
    IDX.cte = d.cte && d.cte.ok ? resolverColunas(d.cte, MAPA.cte) : {};
    IDX.cot = d.cotacoes && d.cotacoes.ok ? resolverColunas(d.cotacoes, MAPA.cot) : {};
    montarConhecidas(d.cte && d.cte.ok ? d.cte.linhas : [], d.cotacoes && d.cotacoes.ok ? d.cotacoes.linhas : []);
    lerCorrecoes(d.correcoes);
    OCORR = d.ocorrencias || {};
    lerFeriados(d.feriados);
    lerPremios(d.premios);
    lerComissao(d.comissao);
    var cte = d.cte && d.cte.ok ? normalizarCte(d.cte) : [], cot = d.cotacoes && d.cotacoes.ok ? normalizarCot(d.cotacoes) : [];
    registrarNome(CONFIG.vendedoraPadrao);
    if (!cte.length && !cot.length) {
      CARIMBO = carimbo; RAW = null; CTE = []; COT = [];
      mostrarEstado('erro', 'As planilhas não têm registros válidos', '<p>Nenhuma linha possui data válida (Data de Emissao do CT-e ou DATA HORA INCLUSAO da cotação).</p>' + PASSOS);
      statusAoVivo(false, 'sem registros'); return;
    }
    marcarDuplicidades(cte, true); marcarDuplicidades(cot, false);
    CARIMBO = carimbo; RAW = d; CTE = cte; COT = cot;
    lerMetas(d.metas);
    esconderEstado();
    montarOpcoesFiltros();
    PADRAO = calcularPadrao();
    if (!PADRAO_APLICADO) {
      PADRAO_APLICADO = true;
      if (!F.ano && !F.mes && !F.de && !F.ate && !F.ym && !F.dia) { F.ano = PADRAO.ano; F.mes = PADRAO.mes; }
    }
    atualizar();
    if (!primeira) aviso('Dados atualizados · ' + nInt(CTE.length) + ' CT-e e ' + nInt(COT.length) + ' cotações');
    statusAoVivo(true);
  }
  function carregarDados(manual) {
    if (carregando) return; carregando = true;
    carregarStatus(function () {
      if (EMBUTIDO) { carregando = false; processarDados(DADOS_EMBUTIDOS, manual); return; }
      var anterior = window.OSTEL_DATA; window.OSTEL_DATA = undefined;
      carregarScript(CONFIG.arquivoDados, 'scriptDados', function (ok) {
        carregando = false;
        if (!ok) {
          window.OSTEL_DATA = anterior; statusAoVivo(false, 'dados.js não encontrado');
          if (!RAW) mostrarEstado('info', 'Os dados ainda não foram gerados', '<p>O arquivo <code>dados.js</code> não foi encontrado nesta pasta.</p>' + PASSOS);
          return;
        }
        var d = window.OSTEL_DATA; if (!d) window.OSTEL_DATA = anterior;
        processarDados(d, manual);
      });
    });
  }
  // "Última atualização" e alerta de dados desatualizados
  function horaVerificacao() { return STATUS && STATUS.verificadoEm ? STATUS.verificadoEm : (RAW && RAW.geradoEm); }
  var VIVO_OK = true;
  function statusAoVivo(ok, msg) {
    if (ok !== undefined) VIVO_OK = ok; else ok = VIVO_OK;
    if (!ok && msg) ultimoErro = msg;
    var lv = $('#live'), ref = horaVerificacao(), velho = false, horas = 0;
    if (!EMBUTIDO && ref) { horas = (Date.now() - new Date(ref).getTime()) / 36e5; velho = horas > CONFIG.alertaDesatualizadoHoras; }
    lv.classList.toggle('err', !ok || velho || (STATUS && STATUS.ok === false));
    var al = $('#alertaVelho');
    if (velho && RAW) {
      al.hidden = false; al.className = 'alerta';
      al.innerHTML = IC.alert + '<div><b>Dados desatualizados:</b> as planilhas não são conferidas há ' + fDec.format(horas) + ' horas (última vez em ' + fmtStamp(ref) + '). ' +
        'Confira se o computador está ligado e se a tarefa agendada existe: rode <code>2_agendar_atualizacao.bat</code> ou <code>1_atualizar_agora.bat</code>.</div>';
    } else if (STATUS && STATUS.ok === false && RAW) {
      al.hidden = false; al.className = 'alerta atencao';
      al.innerHTML = IC.alert + '<div><b>Última conferência com problema</b> (' + fmtStamp(STATUS.verificadoEm) + '): ' + esc(STATUS.mensagem || '') + '. O painel continua com os últimos dados válidos. Detalhes em <code>logs\\atualizacao.log</code>.</div>';
    } else al.hidden = true;
    if (EMBUTIDO) {
      $('#liveTxt').textContent = ok ? 'cópia dos dados de ' + fmtStamp(RAW && RAW.geradoEm) : ultimoErro;
      lv.title = 'Arquivo para envio: mostra os dados do momento em que foi gerado.';
      return;
    }
    $('#liveTxt').textContent = ok ? 'Última atualização: ' + fmtStamp(ref) + ' · próxima em ' + contagem + 's' : ultimoErro + ' · nova tentativa em ' + contagem + 's';
    lv.title = (RAW ? 'Dados gerados em ' + fmtStamp(RAW.geradoEm) + '. ' : '') + (STATUS ? 'Planilhas conferidas pela última vez em ' + fmtStamp(STATUS.verificadoEm) + '.' : '') + ' O painel relê os dados a cada ' + CONFIG.recarregarSegundos + ' segundos.';
  }
  if (!EMBUTIDO) setInterval(function () {
    contagem--;
    if (contagem <= 0) { contagem = CONFIG.recarregarSegundos; carregarDados(false); }
    if (!carregando) statusAoVivo();
  }, 1000);
  $('#btnRefresh').addEventListener('click', function () { contagem = CONFIG.recarregarSegundos; carregarDados(true); });
  document.addEventListener('visibilitychange', function () { if (!document.hidden && !EMBUTIDO) { contagem = CONFIG.recarregarSegundos; carregarDados(false); } });

  /* ===================================================================
     AVISO RÁPIDO (toast)
     =================================================================== */
  var tAviso;
  function aviso(msg) {
    if (SEM_AVISO) { SEM_AVISO = false; return; }
    var t = $('#toast');
    t.innerHTML = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4"><path d="M20 6L9 17l-5-5"/></svg>' + esc(msg);
    t.classList.add('show'); clearTimeout(tAviso);
    tAviso = setTimeout(function () { t.classList.remove('show'); }, 2600);
  }

  /* ===================================================================
     FILTROS NO ENDEREÇO DA PÁGINA (para salvar como favorito)
     =================================================================== */
  function salvarNaURL() {
    var p = Object.keys(F).filter(function (k) { return F[k]; }).map(function (k) { return k + '=' + encodeURIComponent(F[k]); });
    if (UI.aba !== 'cte') p.push('aba=' + UI.aba);
    try { history.replaceState(null, '', p.length ? '#' + p.join('&') : location.pathname + location.search); } catch (e) { /* ignora */ }
  }
  function lerDaURL() {
    var h = location.hash.replace(/^#/, ''); if (!h) return;
    h.split('&').forEach(function (par) {
      var i = par.indexOf('='); if (i < 0) return;
      var k = par.slice(0, i), v = decodeURIComponent(par.slice(i + 1));
      if (Object.prototype.hasOwnProperty.call(F, k)) F[k] = v;
      if (k === 'aba' && (v === 'cot' || v === 'cte')) { UI.aba = v; $$('#segTab button').forEach(function (b) { b.classList.toggle('on', b.getAttribute('data-v') === v); }); }
    });
  }
  lerDaURL();

  /* ===================================================================
     JANELA DE DETALHAMENTO DO INDICADOR
     melhor e pior período, média, evolução e os 5 maiores clientes
     =================================================================== */
  var graficoModal = null;
  function abrirModal(html) {
    var m = $('#modal'); m.querySelector('.m-body').innerHTML = html;
    m.hidden = false; document.body.classList.add('travado');
    requestAnimationFrame(function () { m.classList.add('show'); });
    m.querySelector('.m-close').focus();
  }
  function fecharModal() {
    var m = $('#modal'); if (m.hidden) return;
    m.classList.remove('show'); document.body.classList.remove('travado');
    if (graficoModal) { graficoModal.destroy(); graficoModal = null; }
    setTimeout(function () { m.hidden = true; }, 200);
  }
  $('#modal').addEventListener('click', function (e) { if (e.target === this || e.target.closest('.m-close')) fecharModal(); });
  function abrirDetalheKPI(id) {
    var k = KPIS.filter(function (x) { return x.id === id; })[0]; if (!k) return;
    var diario = porDia(), campo = diario ? 'dt' : 'ym', grupos = {};
    FC.forEach(function (o) { (grupos[o[campo]] = grupos[o[campo]] || { c: [], q: [] }).c.push(o); });
    FQ.forEach(function (o) { (grupos[o[campo]] = grupos[o[campo]] || { c: [], q: [] }).q.push(o); });
    var chaves = Object.keys(grupos).sort(), valores = chaves.map(function (c) { return k.m(grupos[c].c, grupos[c].q); });
    var nomeP = function (c) { return diario ? fDate(c) : nomeMes(c); }, unidade = diario ? 'dia' : 'mês';
    var cli = {};
    FC.forEach(function (o) { (cli[o.cliente] = cli[o.cliente] || { c: [], q: [] }).c.push(o); });
    FQ.forEach(function (o) { (cli[o.cliente] = cli[o.cliente] || { c: [], q: [] }).q.push(o); });
    var top = Object.keys(cli).filter(function (c) { var n = k.q ? cli[c].q.length : cli[c].c.length; return c !== '—' && n && (!k.min || n >= k.min); })
      .map(function (c) { return { c: c, v: k.m(cli[c].c, cli[c].q), n: k.q ? cli[c].q.length : cli[c].c.length }; })
      .sort(function (a, b) { return b.v - a.v; }).slice(0, 5);
    var melhor = null, pior = null;
    chaves.forEach(function (c, i) { if (!melhor || valores[i] > melhor.v) melhor = { c: c, v: valores[i] }; if (!pior || valores[i] < pior.v) pior = { c: c, v: valores[i] }; });
    var media = chaves.length ? valores.reduce(function (a, b) { return a + b; }, 0) / chaves.length : 0;
    abrirModal(
      '<div class="m-head" style="--k:' + css(k.c) + '"><span class="k-ic">' + IC[k.ic] + '</span><div><h2>' + k.l + '</h2></div>' +
      '<div class="m-total num">' + k.fmt(k.m(FC, FQ)) + '</div></div>' +
      '<div class="m-cards">' +
        '<div><span>Melhor ' + unidade + '</span><b>' + (melhor ? nomeP(melhor.c) : '—') + '</b><small>' + (melhor ? k.fmt(melhor.v) : '') + '</small></div>' +
        '<div><span>Pior ' + unidade + '</span><b>' + (pior ? nomeP(pior.c) : '—') + '</b><small>' + (pior ? k.fmt(pior.v) : '') + '</small></div>' +
        '<div><span>Média por ' + unidade + '</span><b>' + (chaves.length ? k.fmt(media) : '—') + '</b><small>' + chaves.length + (diario ? ' dias com movimento' : ' meses') + '</small></div>' +
      '</div>' +
      '<h3>Evolução ' + (diario ? 'diária' : 'mensal') + '</h3><div class="chart-box"><canvas id="chModal"></canvas></div>' +
      '<h3>5 maiores clientes' + (k.min ? ' <small>(mínimo ' + k.min + (k.q ? ' cotações' : ' CT-e') + ')</small>' : '') + '</h3>' +
      '<ol class="m-top">' + (top.length ? top.map(function (t) {
        return '<li><button type="button" data-cli="' + esc(t.c) + '" title="Filtrar o painel por este cliente"><span>' + esc(t.c) + '</span><em>' + nInt(t.n) + (k.q ? ' cotações' : ' CT-e') + '</em><b class="num">' + k.fmt(t.v) + '</b></button></li>';
      }).join('') : '<li class="empty">Sem dados</li>') + '</ol>' +
      '<p class="m-hint">Clique em um ' + unidade + ' do gráfico ou em um cliente para filtrar o painel.</p>');
    $$('.m-top button').forEach(function (b) { b.onclick = function () { fecharModal(); definirFiltro('cliente', b.getAttribute('data-cli')); }; });
    if (TEM_CHART() && chaves.length) {
      temaGraficos();
      var cor = css(k.c);
      if (!VALORES_REGISTRADO && window.Chart.register) { window.Chart.register(pluginValores); VALORES_REGISTRADO = true; }
      graficoModal = new window.Chart($('#chModal'), {
        type: 'line',
        data: { labels: chaves.map(function (c) { return diario ? c.slice(8, 10) + '/' + c.slice(5, 7) : nomeMes(c); }),
                datasets: [{ data: valores, borderColor: cor, backgroundColor: alfa(cor, .14), fill: true, tension: .3, borderWidth: 2.2, pointRadius: 3, pointBackgroundColor: css('--surface'), pointBorderColor: cor, pointBorderWidth: 2 }] },
        options: { responsive: true, maintainAspectRatio: false, layout: { padding: { top: 8 } }, plugins: { legend: { display: false }, valores: { fmt: k.fmt, linhas: true, minLargura: 70 },
                   tooltip: { callbacks: { label: function (c) { return ' ' + k.fmt(c.parsed.y); }, footer: function () { return 'Clique para filtrar'; } } } },
                   scales: { y: { beginAtZero: true, ticks: { callback: function (v) { return k.fmt === pct ? fInt.format(v * 100) + '%' : (k.fmt === nInt ? nInt(v) : brlc(v)); } }, border: { display: false } },
                             x: { grid: { display: false }, ticks: { maxRotation: 0, autoSkipPadding: 10 } } },
                   onHover: cursor,
                   onClick: function (e, els) { if (els.length) { fecharModal(); definirFiltro(diario ? 'dia' : 'ym', chaves[els[0].index]); } } }
      });
    }
  }

  /* ===================================================================
     FICHA COMPLETA DA LINHA (painel lateral; setas passam de linha)
     =================================================================== */
  var FICHA_I = -1;
  function campo(rot, valor, cls) { return '<div class="f-item' + (cls ? ' ' + cls : '') + '"><span>' + esc(rot) + '</span><b>' + (valor == null || valor === '' || valor === '—' ? '—' : esc(valor)) + '</b></div>'; }
  function abrirFicha(i) {
    var o = TABELA_ORDEM[i]; if (!o) return;
    FICHA_I = i; UI.sel = o;      // linha selecionada: continua marcada depois de fechar a ficha
    var d = $('#ficha'), cols = colunasTabela().filter(function (c) { return !c.calc; }), cte = o.t === 'cte';
    var badge = cte ? '<span class="badge"><i style="background:' + corBaixa(o.baixa) + '"></i>' + esc(titulo(o.baixa)) + '</span>'
                    : '<span class="badge"><i style="background:' + corSituacao(o.situacao) + '"></i>' + esc(o.situacao) + '</span>';
    d.querySelector('.d-title').innerHTML = '<div class="c-tag">' + (cte ? 'CT-e ' : 'Cotação ') + esc(o.numero || '') + '</div><h2>' + esc(o.cliente) + '</h2>' + badge;
    d.querySelector('.d-body').innerHTML =
      (o.dup || o.ind ? '<div class="d-dup">' + (o.dup ? 'Linha duplicada: ' : 'Indício de duplicidade (conferir): ') + esc(o.dupMotivo) + '.</div>' : '') +
      '<div class="d-dest">' + IC.pin + '<div><b>' + esc(titulo(o.cidade)) + ' / ' + esc(o.uf) + '</b><span>' + esc(fDate(o.dt)) + (o.hora ? ' ' + o.hora : '') + '</span></div></div>' +
      '<h3>Calculado pelo painel</h3><div class="f-grid2">' +
        campo('Vendedora', o.vendNome) + (cte ? campo('De onde veio a vendedora', o.vendFonte) : campo('De onde veio a vendedora', 'usuário que incluiu a cotação')) +
        (cte ? campo('Aguardando embarque', o.aguardando ? 'Sim · ' + nomeOcorr(o.ocorr) : 'Não') + (o.complemento ? campo('Complemento de carga', 'Sim (sem saída e frete acima de ' + brl(CONFIG.complementoAcimaDe) + ')') : '') + campo('Origem da carga (coluna E)', ORIGENS[o.origem]) + campo('Faturamento (Valor do Frete)', o.cancelado ? 'cancelado: fora das somas' : brl(o.frete)) +
               campo('ICMS (frete − frete sem ICMS)', brl(o.icms)) : campo('Valor cotado (Proposta Atual)', o.cancelada ? 'cancelada: fora das somas' : brl(o.valor)) + campo('Frete inicial (Proposta Inicial)', brl(o.valorIni)) + campo('Variação do frete', o.valorIni ? (o.valor > o.valorIni ? '+' : '') + fDec.format((o.valor - o.valorIni) / o.valorIni * 100) + '%' : null)) +
        campo('Entra nas somas', o.soma ? 'Sim' : 'Não (cancelado: só contagem)') + '</div>' +
      '<h3>Todos os campos da planilha</h3><div class="f-grid2">' +
        cols.map(function (c) { return campo(c.l + ' (' + c.letra + ')', fmtValorCol(valorCol(o, c), c)); }).join('') + '</div>';
    $('#fichaPos').textContent = (i + 1) + ' de ' + nInt(TABELA_ORDEM.length);
    $('#fichaAnt').disabled = i <= 0; $('#fichaProx').disabled = i >= TABELA_ORDEM.length - 1;
    $('#fichaCli').onclick = function () { fecharFicha(); definirFiltro('cliente', o.cliente); };
    if (d.hidden) { d.hidden = false; document.body.classList.add('travado'); requestAnimationFrame(function () { d.classList.add('show'); }); d.querySelector('.d-close').focus(); }
    $$('#tbl tbody tr').forEach(function (tr) { tr.classList.toggle('sel', +tr.getAttribute('data-i') === i); });
  }
  function fecharFicha() {
    var d = $('#ficha'); if (d.hidden) return;
    d.classList.remove('show'); document.body.classList.remove('travado');
    setTimeout(function () { d.hidden = true; }, 220);
  }
  $('#ficha').addEventListener('click', function (e) { if (e.target === this || e.target.closest('.d-close')) fecharFicha(); });
  $('#fichaAnt').addEventListener('click', function () { navegarFicha(-1); });
  $('#fichaProx').addEventListener('click', function () { navegarFicha(1); });
  function navegarFicha(passo) {
    var n = FICHA_I + passo; if (n < 0 || n >= TABELA_ORDEM.length) return;
    var pag = Math.floor(n / UI.porPagina) + 1;
    if (pag !== UI.pagina) { UI.pagina = pag; desenharTabela(); }
    abrirFicha(n);
  }
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') { fecharModal(); fecharFicha(); fecharExpandido(); }
    if (!$('#ficha').hidden && (e.key === 'ArrowRight' || e.key === 'ArrowLeft' || e.key === 'ArrowDown' || e.key === 'ArrowUp') && !/INPUT|SELECT|TEXTAREA/.test(document.activeElement.tagName)) {
      e.preventDefault(); navegarFicha(e.key === 'ArrowRight' || e.key === 'ArrowDown' ? 1 : -1);
    }
  });

  /* ===================================================================
     BOTÕES DE CADA GRÁFICO: tabela, Excel, PDF, tela cheia e imagem
     =================================================================== */
  var ICO_EXP = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5"/></svg>';
  var ICO_DL  = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="16" rx="2"/><circle cx="9" cy="10" r="2"/><path d="M21 16l-5-5-8 9"/></svg>';
  var ICO_TAB = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 10h18M3 15h18M9 4v16"/></svg>';
  var ICO_GRAF = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 20V10M10 20V4M16 20v-7M21 20H3"/></svg>';
  var ICO_XLS = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 3H6a1 1 0 0 0-1 1v16a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V8z"/><path d="M14 3v5h5"/><path d="M9.5 12l5 6M14.5 12l-5 6"/></svg>';
  var ICO_PDF = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 9V3h12v6"/><rect x="3" y="9" width="18" height="8" rx="2"/><path d="M7 14h10v7H7z"/></svg>';
  $$('.grid .card').forEach(function (card) {
    var canvas = card.querySelector('canvas'); if (!canvas) return;
    var semGrafico = card.classList.contains('sem-grafico');     // card só com a relação (Status dos CT-e)
    var t = document.createElement('div'); t.className = 'c-tools';
    t.innerHTML = '<button type="button" class="tbtn" data-a="tab" title="Ver como tabela" aria-label="Ver como tabela" aria-pressed="false">' + ICO_TAB + '</button>' +
                  '<button type="button" class="tbtn" data-a="xls" title="Baixar para Excel (.xlsx)" aria-label="Baixar os dados do gráfico para Excel">' + ICO_XLS + '</button>' +
                  '<button type="button" class="tbtn" data-a="exp" title="Ampliar em tela cheia (Esc para sair)" aria-label="Ampliar gráfico">' + ICO_EXP + '</button>' +
                  (semGrafico ? '' : '<button type="button" class="tbtn" data-a="png" title="Baixar imagem (PNG)" aria-label="Baixar imagem do gráfico">' + ICO_DL + '</button>');
    card.querySelector('.c-head').appendChild(t);
    t.addEventListener('click', function (e) {
      var b = e.target.closest('button'); if (!b) return;
      var acao = b.getAttribute('data-a'), nome = card.querySelector('h3').textContent;
      if (acao === 'exp') alternarExpandido(card);
      else if (acao === 'tab') alternarTabela(card, canvas.id, b);
      else if (acao === 'xls') baixarExcel(canvas.id, nome);
      else if (acao === 'pdf') baixarPDF(card, canvas, nome);
      else if (MODO_TABELA[canvas.id]) aviso('Volte para o gráfico para baixar a imagem');
      else baixarPNG(canvas, nome);
    });
  });
  function alternarExpandido(card) {
    var ja = card.classList.contains('expandido');
    fecharExpandido(); if (ja) return;
    card.classList.add('expandido'); document.body.classList.add('travado', 'com-expandido');
    setTimeout(function () { redesenharGraficos(false); }, 60);
  }
  function fecharExpandido() {
    var c = $('.card.expandido'); if (!c) return;
    c.classList.remove('expandido'); document.body.classList.remove('com-expandido');
    if ($('#ficha').hidden && $('#modal').hidden) document.body.classList.remove('travado');
    setTimeout(function () { redesenharGraficos(false); }, 60);
  }
  document.addEventListener('click', function (e) { if (e.target.classList && e.target.classList.contains('com-expandido')) fecharExpandido(); });
  function baixarPNG(canvas, nome) {
    var c = document.createElement('canvas'); c.width = canvas.width; c.height = canvas.height;
    var ctx = c.getContext('2d'); ctx.fillStyle = css('--surface'); ctx.fillRect(0, 0, c.width, c.height); ctx.drawImage(canvas, 0, 0);
    c.toBlob(function (b) { baixar(b, nomeArquivo(nome) + '.png'); aviso('Imagem do gráfico baixada'); }, 'image/png');
  }

  /* ===================================================================
     CABEÇALHO FIXO: quando o topo "asfalto" sai da tela, a barra superior
     passa a mostrar o logo, o título e o resumo do filtro
     =================================================================== */
  (function () {
    var hero = $('.hero');
    if (!hero || !('IntersectionObserver' in window)) return;
    new IntersectionObserver(function (es) { document.body.classList.toggle('rolou', !es[0].isIntersecting); }, { rootMargin: '-60px 0px 0px 0px' }).observe(hero);
  })();

  /* ===================================================================
     ROLAGEM AUTOMÁTICA (TV): após CONFIG.rolagemAutomaticaSegundos sem
     ninguém mexer, desce a página, rola as linhas da tabela, espera,
     volta ao topo e recomeça. Mouse, clique, tecla ou toque param na hora.
     Não roda no celular nem com janela aberta.
     =================================================================== */
  (function () {
    var espera = +CONFIG.rolagemAutomaticaSegundos || 0, veloc = +CONFIG.rolagemVelocidade || 45;
    if (!espera || !window.requestAnimationFrame) return;
    var ativa = false, quadro = 0, relogio = 0, pausa = 0, fase = 'pagina', pos = 0, antes = 0, mx = -1, my = -1;
    function podeRolar() { return innerWidth >= 768 && !document.hidden && !$('#conteudo').hidden && !document.body.classList.contains('travado'); }
    function fimDaPagina() { return Math.max(0, document.documentElement.scrollHeight - innerHeight); }
    function passo(t) {
      if (!ativa) return;
      if (!podeRolar()) return parar();
      var dt = Math.min(0.1, (t - antes) / 1000); antes = t;
      var wrap = $('.tbl-wrap');
      if (pausa > 0) {
        pausa -= dt;
        if (pausa <= 0 && fase === 'voltar') { if (wrap) wrap.scrollTop = 0; window.scrollTo({ top: 0, behavior: 'smooth' }); fase = 'topo'; pausa = 8; }
        else if (pausa <= 0 && fase === 'topo') { fase = 'pagina'; pos = window.pageYOffset; }
      } else if (fase === 'pagina') {
        pos += veloc * dt; window.scrollTo(0, pos);
        if (pos >= fimDaPagina()) { fase = 'tabela'; pos = wrap ? wrap.scrollTop : 0; }
      } else if (fase === 'tabela') {
        var fim = wrap ? wrap.scrollHeight - wrap.clientHeight : 0;
        pos += veloc * dt; if (wrap) wrap.scrollTop = pos;
        if (pos >= fim) { fase = 'voltar'; pausa = 5; }
      }
      quadro = requestAnimationFrame(passo);
    }
    function iniciar() {
      if (ativa || !podeRolar()) { agendar(); return; }
      ativa = true; fase = 'pagina'; pausa = 0; pos = window.pageYOffset; antes = performance.now();
      document.body.classList.add('auto-rolagem');
      aviso('Rolagem automática ligada. Mexa o mouse para parar.');
      quadro = requestAnimationFrame(passo);
    }
    function parar() {
      if (!ativa) return;
      ativa = false; cancelAnimationFrame(quadro); document.body.classList.remove('auto-rolagem');
      window.scrollTo(0, window.pageYOffset); agendar();
    }
    function agendar() { clearTimeout(relogio); relogio = setTimeout(iniciar, espera * 1000); }
    function mexeu() { if (ativa) parar(); else agendar(); }
    ['mousedown', 'keydown', 'wheel', 'touchstart'].forEach(function (ev) { window.addEventListener(ev, mexeu, { passive: true, capture: true }); });
    window.addEventListener('mousemove', function (e) {
      if (e.screenX === mx && e.screenY === my) return;
      var primeiro = mx < 0; mx = e.screenX; my = e.screenY;
      if (!primeiro || !ativa) mexeu();
    }, { passive: true, capture: true });
    document.addEventListener('visibilitychange', function () { if (document.hidden) parar(); else agendar(); });
    agendar();
  })();

  /* ===================================================================
     INÍCIO
     =================================================================== */
  carregarDados(false);
})();
