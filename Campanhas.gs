// ============================================
// CONFIGURAÇÕES
// ============================================
var SPREADSHEET_ID = '15ZflxC82kDMcK8k4CbZQmX_fO9wDUl86tGYlJ3IMiT4';
var SHEET_NAME = 'SaldoEmpenhos';
var SHEET_CONSOLIDADO = 'Consolidado';
var CACHE_KEY = 'FINALIZADAS_CACHE';
var CACHE_DURATION = 21600; // 6 horas em segundos

// ============================================
// HELPERS
// ============================================
function getSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Aba "' + SHEET_NAME + '" não encontrada');
  return sheet;
}

function getSheetConsolidado_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_CONSOLIDADO);
  if (!sheet) throw new Error('Aba "' + SHEET_CONSOLIDADO + '" não encontrada');
  return sheet;
}

function normalizeString_(str) {
  return String(str ?? "").trim().toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function normalizeHeader_(v) {
  return String(v ?? "").trim().toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, " ");
}

function colIndex_(headers, name) {
  var target = normalizeHeader_(name);
  for (var i = 0; i < headers.length; i++) {
    if (normalizeHeader_(headers[i]) === target) return i;
  }
  throw new Error('Cabeçalho não encontrado: "' + name + '"');
}

function optionalColIndex_(headers, name) {
  try {
    return colIndex_(headers, name);
  } catch (e) {
    return -1;
  }
}

function firstColIndex_(headers, names) {
  for (var i = 0; i < names.length; i++) {
    var idx = optionalColIndex_(headers, names[i]);
    if (idx >= 0) return idx;
  }
  return -1;
}

function colIndexObservacoes_(headers) {
  var idx = firstColIndex_(headers, ['Observações', 'Observacoes', 'Observação', 'Observacao']);
  if (idx >= 0) return idx;
  // fallback fixo: coluna AI (A=0 → AI=34)
  return 34;
}

function buildColMap_(headers) {
  return {
    ORDEM: colIndex_(headers, 'Ordem'),
    TERMO: colIndex_(headers, 'Termo Aditivo'),
    TIPO: colIndex_(headers, 'Tipo'),
    EXERCICIO: colIndex_(headers, 'Exercício Financeiro'),
    CAMPANHA: colIndex_(headers, 'Campanha'),
    AGENCIA: colIndex_(headers, 'Agência'),
    EMPENHO_TOTAL: colIndex_(headers, 'Empenho Total'),
    SALDO_ATUAL: colIndex_(headers, 'Saldo Atual'),
    SITUACAO: colIndex_(headers, 'Situação da Campanha'),
    FATURADO_QTD: colIndex_(headers, 'Faturado pela agência e enviado para a comissão (quantas)'),
    FATURADO_VAL: colIndex_(headers, 'Faturado pela agência e enviado para a comissão'),
    PAGO_QTD: colIndex_(headers, 'Pago (quantas)'),
    PAGO_VAL: colIndex_(headers, 'Pago (valor)'),
    EM_PROCESSO_QTD: colIndex_(headers, 'Em processo de Pagamento (quantas)'),
    EM_PROCESSO_VAL: colIndex_(headers, 'Em processo de Pagamento (valor)'),
    GLOSA_QTD: colIndex_(headers, 'Glosa Total (quantas)'),
    GLOSA_VAL: colIndex_(headers, 'Glosa Total (valor)'),
    ATESTADA_QTD: colIndex_(headers, 'Atestada (quantas)'),
    ATESTADA_VAL: colIndex_(headers, 'Atestada (valor)'),
    EM_ANALISE_QTD: colIndex_(headers, 'Em análise (quantas)'),
    EM_ANALISE_VAL: colIndex_(headers, 'Em Análise (valor)'),
    INCONFORMIDADE_QTD: colIndex_(headers, 'Inconformidade (quantas)'),
    INCONFORMIDADE_VAL: colIndex_(headers, 'Inconformidade (valor)'),
    EM_DILIGENCIA_QTD: colIndex_(headers, 'Em Diligência Interna (quantas)'),
    EM_DILIGENCIA_VAL: colIndex_(headers, 'Em Diligência Interna (valor)'),
    CANCELADA_QTD: colIndex_(headers, 'Cancelada (quantas)'),
    CANCELADA_VAL: colIndex_(headers, 'Cancelada (valor)'),
    A_ATESTAR_QTD: colIndex_(headers, 'A Atestar (quantas)'),
    A_ATESTAR_VAL: colIndex_(headers, 'A Atestar (valor)'),
    PROCESSO_MAE: colIndex_(headers, 'Processo Mãe'),
    PROCESSO: colIndex_(headers, 'Processo'),
    AUTORIZACAO: colIndex_(headers, 'Autorização'),
    INICIO: colIndex_(headers, 'Início'),
    FIM: colIndex_(headers, 'Fim'),

    // opcionais (não quebram se o cabeçalho não existir)
    DESCRICAO: firstColIndex_(headers, ['Resumo da campanha', 'Resumo da Campanha', 'Resumo', 'Descrição', 'Descricao']),
    OBSERVACOES: colIndexObservacoes_(headers)
  };
}

function criarAgenciaVazia_() {
  return {
    processo: '',
    empenhoTotal: 0,
    saldoAtual: 0,
    faturado: { qtd: 0, val: 0 },
    pago: { qtd: 0, val: 0 },
    emProcesso: { qtd: 0, val: 0 },
    glosa: { qtd: 0, val: 0 },
    atestada: { qtd: 0, val: 0 },
    emAnalise: { qtd: 0, val: 0 },
    inconformidade: { qtd: 0, val: 0 },
    emDiligencia: { qtd: 0, val: 0 },
    cancelada: { qtd: 0, val: 0 },
    aAtestar: { qtd: 0, val: 0 }
  };
}

function normalizeAgencia_(agRaw) {
  var ag = normalizeString_(agRaw);
  ag = ag.replace(/[^a-z0-9]/g, '');

  if (ag === 'av' || ag.indexOf('av') === 0) return 'AV';
  if (ag.indexOf('calia') === 0) return 'Calia';
  if (ag.indexOf('ebm') === 0) return 'EBM';

  return String(agRaw ?? '').trim();
}

function formatDateValue_(value) {
  if (!value) return '';
  if (value instanceof Date) {
    var day = ('0' + value.getDate()).slice(-2);
    var month = ('0' + (value.getMonth() + 1)).slice(-2);
    var year = String(value.getFullYear()).slice(-2); // dd/mm/yy
    return day + '/' + month + '/' + year;
  }
  return String(value);
}

// ============================================
// CACHE PARA CAMPANHAS FINALIZADAS
// ============================================
function getCachedFinalizadas_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(CACHE_KEY);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      return null;
    }
  }
  return null;
}

function setCachedFinalizadas_(data) {
  var cache = CacheService.getScriptCache();
  try {
    cache.put(CACHE_KEY, JSON.stringify(data), CACHE_DURATION);
  } catch (e) {
    Logger.log('Erro ao salvar cache: ' + e.message);
  }
}

function buildFinalizadasCache_() {
  var sheet = getSheet_();
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var COL = buildColMap_(headers);
  var finalizadas = {};
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var situacao = normalizeString_(row[COL.SITUACAO]);
    if (situacao === 'finalizada' || situacao === 'encerrada') {
      var campanha = row[COL.CAMPANHA];
      var ag = normalizeAgencia_(row[COL.AGENCIA]);
      if (!finalizadas[campanha]) {
        finalizadas[campanha] = { AV: null, Calia: null, EBM: null };
      }
      if (!finalizadas[campanha][ag]) {
        finalizadas[campanha][ag] = criarAgenciaVazia_();
      }
      var a = finalizadas[campanha][ag];
      if (!a.processo) a.processo = row[COL.PROCESSO] || '';
      a.empenhoTotal += Number(row[COL.EMPENHO_TOTAL]) || 0;
      a.saldoAtual += Number(row[COL.SALDO_ATUAL]) || 0;
      a.faturado.qtd += Number(row[COL.FATURADO_QTD]) || 0;
      a.faturado.val += Number(row[COL.FATURADO_VAL]) || 0;
      a.pago.qtd += Number(row[COL.PAGO_QTD]) || 0;
      a.pago.val += Number(row[COL.PAGO_VAL]) || 0;
      a.emProcesso.qtd += Number(row[COL.EM_PROCESSO_QTD]) || 0;
      a.emProcesso.val += Number(row[COL.EM_PROCESSO_VAL]) || 0;
      a.glosa.qtd += Number(row[COL.GLOSA_QTD]) || 0;
      a.glosa.val += Number(row[COL.GLOSA_VAL]) || 0;
      a.atestada.qtd += Number(row[COL.ATESTADA_QTD]) || 0;
      a.atestada.val += Number(row[COL.ATESTADA_VAL]) || 0;
      a.emAnalise.qtd += Number(row[COL.EM_ANALISE_QTD]) || 0;
      a.emAnalise.val += Number(row[COL.EM_ANALISE_VAL]) || 0;
      a.inconformidade.qtd += Number(row[COL.INCONFORMIDADE_QTD]) || 0;
      a.inconformidade.val += Number(row[COL.INCONFORMIDADE_VAL]) || 0;
      a.emDiligencia.qtd += Number(row[COL.EM_DILIGENCIA_QTD]) || 0;
      a.emDiligencia.val += Number(row[COL.EM_DILIGENCIA_VAL]) || 0;
      a.cancelada.qtd += Number(row[COL.CANCELADA_QTD]) || 0;
      a.cancelada.val += Number(row[COL.CANCELADA_VAL]) || 0;
      a.aAtestar.qtd += Number(row[COL.A_ATESTAR_QTD]) || 0;
      a.aAtestar.val += Number(row[COL.A_ATESTAR_VAL]) || 0;
    }
  }
  setCachedFinalizadas_(finalizadas);
  return finalizadas;
}

function limparCache() {
  var cache = CacheService.getScriptCache();
  cache.remove(CACHE_KEY);
  Logger.log('Cache limpo com sucesso');
}

// ============================================
// FUNÇÕES EXPOSTAS AO FRONTEND
// ============================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Campanhas')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getCampanhas() {
  try {
    var sheet = getSheet_();
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var COL = buildColMap_(headers);
    var campanhasMap = {};
    for (var i = 1; i < values.length; i++) {
      var c = values[i][COL.CAMPANHA];
      var ordem = Number(values[i][COL.ORDEM]) || 0;
      if (c && (!campanhasMap[c] || ordem > campanhasMap[c])) {
        campanhasMap[c] = ordem;
      }
    }
    var campanhasArr = Object.keys(campanhasMap).map(function(nome) {
      return { nome: nome, ordem: campanhasMap[nome] };
    });
    campanhasArr.sort(function(a, b) {
      return b.ordem - a.ordem;
    });
    return campanhasArr.map(function(item) { return item.nome; });
  } catch (e) {
    Logger.log('getCampanhas erro: ' + e.message);
    throw e;
  }
}

function getTermos() {
  try {
    var sheet = getSheet_();
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var COL = buildColMap_(headers);
    var termos = [];
    for (var i = 1; i < values.length; i++) {
      var t = values[i][COL.TERMO];
      if (t && termos.indexOf(t) === -1) termos.push(t);
    }
    return termos.sort();
  } catch (e) {
    Logger.log('getTermos erro: ' + e.message);
    throw e;
  }
}

function getTipos() {
  try {
    var sheet = getSheet_();
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var COL = buildColMap_(headers);
    var tipos = [];
    for (var i = 1; i < values.length; i++) {
      var t = values[i][COL.TIPO];
      if (t && tipos.indexOf(t) === -1) tipos.push(t);
    }
    return tipos.sort();
  } catch (e) {
    Logger.log('getTipos erro: ' + e.message);
    throw e;
  }
}

function getFases() {
  return ['Em Andamento', 'Finalizada'];
}

/**
 * Retorna campanhas organizadas por termo aditivo
 * @returns {Object} { "Termo 1": ["Campanha A", "Campanha B"], "Termo 2": [...] }
 */
function getCampanhasPorTermo() {
  try {
    var sheet = getSheet_();
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var COL = buildColMap_(headers);

    // Map de termo -> { campanha -> ordem }
    var termosCampanhas = {};

    for (var i = 1; i < values.length; i++) {
      var termo = values[i][COL.TERMO];
      var campanha = values[i][COL.CAMPANHA];
      var ordem = Number(values[i][COL.ORDEM]) || 0;

      if (!termo || !campanha) continue;

      if (!termosCampanhas[termo]) {
        termosCampanhas[termo] = {};
      }

      if (!termosCampanhas[termo][campanha] || ordem > termosCampanhas[termo][campanha]) {
        termosCampanhas[termo][campanha] = ordem;
      }
    }

    // Converter para o formato { termo: [campanhas ordenadas] }
    var resultado = {};
    var termosOrdenados = Object.keys(termosCampanhas).sort();

    termosOrdenados.forEach(function(termo) {
      var campanhasMap = termosCampanhas[termo];
      var campanhasArr = Object.keys(campanhasMap).map(function(nome) {
        return { nome: nome, ordem: campanhasMap[nome] };
      });

      // Ordenar campanhas por ordem decrescente
      campanhasArr.sort(function(a, b) {
        return b.ordem - a.ordem;
      });

      resultado[termo] = campanhasArr.map(function(item) { return item.nome; });
    });

    return resultado;
  } catch (e) {
    Logger.log('getCampanhasPorTermo erro: ' + e.message);
    throw e;
  }
}

/**
 * Retorna informações detalhadas de uma campanha
 * Colunas AD a AI (exceto AE):
 * AD = Processo Mãe (col 29)
 * AE = Processo (col 30) - EXCLUÍDO
 * AF = Autorização (col 31)
 * AG = Início (col 32)
 * AH = Fim (col 33)
 * AI = Observações (col 34)
 */
function getCampanhaInfo(campanha) {
  try {
    var sheet = getSheet_();
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var COL = buildColMap_(headers);
    var campanhaNorm = normalizeString_(campanha);

    var descricao = '';
    var observacoes = '';
    var termo = '';
    var tipo = '';
    var processoMae = '';
    var autorizacao = '';
    var inicio = '';
    var fim = '';
    var fase = '';

    for (var i = 1; i < values.length; i++) {
      if (normalizeString_(values[i][COL.CAMPANHA]) === campanhaNorm) {
        // Pegar valores que ainda não foram preenchidos
        if (!termo) termo = values[i][COL.TERMO] || '';
        if (!tipo) tipo = values[i][COL.TIPO] || '';
        if (!processoMae) processoMae = values[i][COL.PROCESSO_MAE] || '';
        if (!autorizacao) autorizacao = values[i][COL.AUTORIZACAO] || '';
        if (!inicio) inicio = formatDateValue_(values[i][COL.INICIO]) || '';
        if (!fim) fim = formatDateValue_(values[i][COL.FIM]) || '';

        // Fase baseada na situação
        if (!fase) {
          var situacao = normalizeString_(values[i][COL.SITUACAO]);
          if (situacao === 'finalizada' || situacao === 'encerrada') {
            fase = 'Finalizada';
          } else {
            fase = 'Em Andamento';
          }
        }

        // Descrição - tentar pegar de qualquer linha que tenha valor
        if (!descricao && COL.DESCRICAO >= 0) {
          var descVal = values[i][COL.DESCRICAO];
          if (descVal && String(descVal).trim()) {
            descricao = String(descVal).trim();
          }
        }
        // Observações - coluna AI (índice 34)
        if (!observacoes && COL.OBSERVACOES >= 0 && values[i].length > COL.OBSERVACOES) {
          var obsVal = values[i][COL.OBSERVACOES];
          if (obsVal && String(obsVal).trim()) {
            observacoes = String(obsVal).trim();
          }
        }
      }
    }

    return {
      campanha: campanha || '',
      descricao: descricao,
      observacoes: observacoes,
      termo: termo,
      tipo: tipo,
      processoMae: processoMae,
      autorizacao: autorizacao,
      inicio: inicio,
      fim: fim,
      fase: fase
    };
  } catch (e) {
    Logger.log('getCampanhaInfo erro: ' + e.message);
    throw e;
  }
}

function getCampanhaFase_(values, COL, campanhaNorm) {
  var situacao = '';

  for (var i = 1; i < values.length; i++) {
    if (normalizeString_(values[i][COL.CAMPANHA]) === campanhaNorm) {
      situacao = normalizeString_(values[i][COL.SITUACAO]);
      if (situacao === 'finalizada' || situacao === 'encerrada') return 'Finalizada';
    }
  }

  return 'Em Andamento';
}

function getCampanhasByFase(fase) {
  try {
    var sheet = getSheet_();
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var COL = buildColMap_(headers);
    var campanhasMap = {};
    for (var i = 1; i < values.length; i++) {
      var c = values[i][COL.CAMPANHA];
      var ordem = Number(values[i][COL.ORDEM]) || 0;
      if (c && (!campanhasMap[c] || ordem > campanhasMap[c].ordem)) {
        campanhasMap[c] = { ordem: ordem };
      }
    }
    var resultado = [];
    for (var campanha in campanhasMap) {
      var campFase = getCampanhaFase_(values, COL, normalizeString_(campanha));
      if (campFase === fase) {
        resultado.push({ nome: campanha, ordem: campanhasMap[campanha].ordem });
      }
    }
    resultado.sort(function(a, b) { return b.ordem - a.ordem; });
    return resultado.map(function(item) { return item.nome; });
  } catch (e) {
    Logger.log('getCampanhasByFase erro: ' + e.message);
    throw e;
  }
}

function getCampanhasByTermo(termo) {
  try {
    var sheet = getSheet_();
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var COL = buildColMap_(headers);
    var termoNorm = normalizeString_(termo);
    var campanhasMap = {};
    for (var i = 1; i < values.length; i++) {
      if (normalizeString_(values[i][COL.TERMO]) === termoNorm) {
        var c = values[i][COL.CAMPANHA];
        var ordem = Number(values[i][COL.ORDEM]) || 0;
        if (c && (!campanhasMap[c] || ordem > campanhasMap[c])) {
          campanhasMap[c] = ordem;
        }
      }
    }
    var resultado = Object.keys(campanhasMap).map(function(nome) {
      return { nome: nome, ordem: campanhasMap[nome] };
    });
    resultado.sort(function(a, b) { return b.ordem - a.ordem; });
    return resultado.map(function(item) { return item.nome; });
  } catch (e) {
    Logger.log('getCampanhasByTermo erro: ' + e.message);
    throw e;
  }
}

function getCampanhasByTipo(tipo) {
  try {
    var sheet = getSheet_();
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var COL = buildColMap_(headers);
    var tipoNorm = normalizeString_(tipo);
    var campanhasMap = {};
    for (var i = 1; i < values.length; i++) {
      if (normalizeString_(values[i][COL.TIPO]) === tipoNorm) {
        var c = values[i][COL.CAMPANHA];
        var ordem = Number(values[i][COL.ORDEM]) || 0;
        if (c && (!campanhasMap[c] || ordem > campanhasMap[c])) {
          campanhasMap[c] = ordem;
        }
      }
    }
    var resultado = Object.keys(campanhasMap).map(function(nome) {
      return { nome: nome, ordem: campanhasMap[nome] };
    });
    resultado.sort(function(a, b) { return b.ordem - a.ordem; });
    return resultado.map(function(item) { return item.nome; });
  } catch (e) {
    Logger.log('getCampanhasByTipo erro: ' + e.message);
    throw e;
  }
}

function getCampanhasNaoFinalizadas() {
  try {
    var sheet = getSheet_();
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var COL = buildColMap_(headers);
    var campanhasMap = {};
    var finalizadas = {};
    for (var i = 1; i < values.length; i++) {
      var c = values[i][COL.CAMPANHA];
      var situacao = normalizeString_(values[i][COL.SITUACAO]);
      if (situacao === 'finalizada' || situacao === 'encerrada') {
        finalizadas[c] = true;
      }
      var ordem = Number(values[i][COL.ORDEM]) || 0;
      if (c && (!campanhasMap[c] || ordem > campanhasMap[c])) {
        campanhasMap[c] = ordem;
      }
    }
    var resultado = [];
    for (var campanha in campanhasMap) {
      if (!finalizadas[campanha]) {
        resultado.push({ nome: campanha, ordem: campanhasMap[campanha] });
      }
    }
    resultado.sort(function(a, b) { return b.ordem - a.ordem; });
    return resultado.map(function(item) { return item.nome; });
  } catch (e) {
    Logger.log('getCampanhasNaoFinalizadas erro: ' + e.message);
    throw e;
  }
}

function getDados(filtros) {
  try {
    var sheet = getSheet_();
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var COL = buildColMap_(headers);
    var finalizadasCache = getCachedFinalizadas_();
    if (!finalizadasCache) {
      finalizadasCache = buildFinalizadasCache_();
    }
    var campanhaNorm = filtros.campanha ? normalizeString_(filtros.campanha) : null;
    var termoNorm = filtros.termo ? normalizeString_(filtros.termo) : null;
    var tipoNorm = filtros.tipo ? normalizeString_(filtros.tipo) : null;
    var fase = filtros.fase || null;
    if (campanhaNorm && finalizadasCache[filtros.campanha]) {
      var cached = finalizadasCache[filtros.campanha];
      return {
        AV: cached.AV || criarAgenciaVazia_(),
        Calia: cached.Calia || criarAgenciaVazia_(),
        EBM: cached.EBM || criarAgenciaVazia_()
      };
    }
    var agencias = { AV: criarAgenciaVazia_(), Calia: criarAgenciaVazia_(), EBM: criarAgenciaVazia_() };
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var rowCampanha = row[COL.CAMPANHA];
      var situacao = normalizeString_(row[COL.SITUACAO]);
      if (campanhaNorm && normalizeString_(rowCampanha) !== campanhaNorm) continue;
      if (termoNorm && normalizeString_(row[COL.TERMO]) !== termoNorm) continue;
      if (tipoNorm && normalizeString_(row[COL.TIPO]) !== tipoNorm) continue;
      if (fase) {
        var campFase = getCampanhaFase_(values, COL, normalizeString_(rowCampanha));
        if (campFase !== fase) continue;
      }
      if (!campanhaNorm && (situacao === 'finalizada' || situacao === 'encerrada') && finalizadasCache[rowCampanha]) {
        continue;
      }
      var ag = normalizeAgencia_(row[COL.AGENCIA]);
      if (!agencias[ag]) continue;
      var a = agencias[ag];
      if (!a.processo) a.processo = row[COL.PROCESSO] || '';
      a.empenhoTotal += Number(row[COL.EMPENHO_TOTAL]) || 0;
      a.saldoAtual += Number(row[COL.SALDO_ATUAL]) || 0;
      a.faturado.qtd += Number(row[COL.FATURADO_QTD]) || 0;
      a.faturado.val += Number(row[COL.FATURADO_VAL]) || 0;
      a.pago.qtd += Number(row[COL.PAGO_QTD]) || 0;
      a.pago.val += Number(row[COL.PAGO_VAL]) || 0;
      a.emProcesso.qtd += Number(row[COL.EM_PROCESSO_QTD]) || 0;
      a.emProcesso.val += Number(row[COL.EM_PROCESSO_VAL]) || 0;
      a.glosa.qtd += Number(row[COL.GLOSA_QTD]) || 0;
      a.glosa.val += Number(row[COL.GLOSA_VAL]) || 0;
      a.atestada.qtd += Number(row[COL.ATESTADA_QTD]) || 0;
      a.atestada.val += Number(row[COL.ATESTADA_VAL]) || 0;
      a.emAnalise.qtd += Number(row[COL.EM_ANALISE_QTD]) || 0;
      a.emAnalise.val += Number(row[COL.EM_ANALISE_VAL]) || 0;
      a.inconformidade.qtd += Number(row[COL.INCONFORMIDADE_QTD]) || 0;
      a.inconformidade.val += Number(row[COL.INCONFORMIDADE_VAL]) || 0;
      a.emDiligencia.qtd += Number(row[COL.EM_DILIGENCIA_QTD]) || 0;
      a.emDiligencia.val += Number(row[COL.EM_DILIGENCIA_VAL]) || 0;
      a.cancelada.qtd += Number(row[COL.CANCELADA_QTD]) || 0;
      a.cancelada.val += Number(row[COL.CANCELADA_VAL]) || 0;
      a.aAtestar.qtd += Number(row[COL.A_ATESTAR_QTD]) || 0;
      a.aAtestar.val += Number(row[COL.A_ATESTAR_VAL]) || 0;
    }
    if (!campanhaNorm && !fase) {
      for (var campKey in finalizadasCache) {
        var campData = finalizadasCache[campKey];
        ['AV', 'Calia', 'EBM'].forEach(function(ag) {
          if (campData[ag]) {
            var a = agencias[ag];
            var c = campData[ag];
            a.empenhoTotal += c.empenhoTotal;
            a.saldoAtual += c.saldoAtual;
            a.faturado.qtd += c.faturado.qtd;
            a.faturado.val += c.faturado.val;
            a.pago.qtd += c.pago.qtd;
            a.pago.val += c.pago.val;
            a.emProcesso.qtd += c.emProcesso.qtd;
            a.emProcesso.val += c.emProcesso.val;
            a.glosa.qtd += c.glosa.qtd;
            a.glosa.val += c.glosa.val;
            a.atestada.qtd += c.atestada.qtd;
            a.atestada.val += c.atestada.val;
            a.emAnalise.qtd += c.emAnalise.qtd;
            a.emAnalise.val += c.emAnalise.val;
            a.inconformidade.qtd += c.inconformidade.qtd;
            a.inconformidade.val += c.inconformidade.val;
            a.emDiligencia.qtd += c.emDiligencia.qtd;
            a.emDiligencia.val += c.emDiligencia.val;
            a.cancelada.qtd += c.cancelada.qtd;
            a.cancelada.val += c.cancelada.val;
            a.aAtestar.qtd += c.aAtestar.qtd;
            a.aAtestar.val += c.aAtestar.val;
          }
        });
      }
    }
    return agencias;
  } catch (e) {
    Logger.log('getDados erro: ' + e.message);
    throw e;
  }
}

function getDadosPorCampanha(filtros) {
  try {
    var sheet = getSheet_();
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var COL = buildColMap_(headers);

    var termoNorm = filtros.termo ? normalizeString_(filtros.termo) : null;
    var tipoNorm = filtros.tipo ? normalizeString_(filtros.tipo) : null;
    var fase = filtros.fase || null;

    var campanhasMap = {};

    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var rowCampanha = row[COL.CAMPANHA];
      if (!rowCampanha) continue;

      if (termoNorm && normalizeString_(row[COL.TERMO]) !== termoNorm) continue;
      if (tipoNorm && normalizeString_(row[COL.TIPO]) !== tipoNorm) continue;

      if (fase) {
        var campFase = getCampanhaFase_(values, COL, normalizeString_(rowCampanha));
        if (campFase !== fase) continue;
      }

      var ordem = Number(row[COL.ORDEM]) || 0;
      if (!campanhasMap[rowCampanha]) {
        campanhasMap[rowCampanha] = {
          nome: rowCampanha,
          descricao: '',
          ordem: ordem,
          AV: criarAgenciaVazia_(),
          Calia: criarAgenciaVazia_(),
          EBM: criarAgenciaVazia_()
        };
      }
      if (ordem > campanhasMap[rowCampanha].ordem) {
        campanhasMap[rowCampanha].ordem = ordem;
      }

      // preencher descrição (se houver)
      if (COL.DESCRICAO >= 0 && !campanhasMap[rowCampanha].descricao) {
        var descVal = row[COL.DESCRICAO];
        if (descVal && String(descVal).trim()) {
          campanhasMap[rowCampanha].descricao = String(descVal).trim();
        }
      }

      var ag = normalizeAgencia_(row[COL.AGENCIA]);
      if (!campanhasMap[rowCampanha][ag]) continue;

      var a = campanhasMap[rowCampanha][ag];
      if (!a.processo) a.processo = row[COL.PROCESSO] || '';
      a.empenhoTotal += Number(row[COL.EMPENHO_TOTAL]) || 0;
      a.saldoAtual += Number(row[COL.SALDO_ATUAL]) || 0;
      a.faturado.qtd += Number(row[COL.FATURADO_QTD]) || 0;
      a.faturado.val += Number(row[COL.FATURADO_VAL]) || 0;
      a.pago.qtd += Number(row[COL.PAGO_QTD]) || 0;
      a.pago.val += Number(row[COL.PAGO_VAL]) || 0;
      a.emProcesso.qtd += Number(row[COL.EM_PROCESSO_QTD]) || 0;
      a.emProcesso.val += Number(row[COL.EM_PROCESSO_VAL]) || 0;
      a.glosa.qtd += Number(row[COL.GLOSA_QTD]) || 0;
      a.glosa.val += Number(row[COL.GLOSA_VAL]) || 0;
      a.atestada.qtd += Number(row[COL.ATESTADA_QTD]) || 0;
      a.atestada.val += Number(row[COL.ATESTADA_VAL]) || 0;
      a.emAnalise.qtd += Number(row[COL.EM_ANALISE_QTD]) || 0;
      a.emAnalise.val += Number(row[COL.EM_ANALISE_VAL]) || 0;
      a.inconformidade.qtd += Number(row[COL.INCONFORMIDADE_QTD]) || 0;
      a.inconformidade.val += Number(row[COL.INCONFORMIDADE_VAL]) || 0;
      a.emDiligencia.qtd += Number(row[COL.EM_DILIGENCIA_QTD]) || 0;
      a.emDiligencia.val += Number(row[COL.EM_DILIGENCIA_VAL]) || 0;
      a.cancelada.qtd += Number(row[COL.CANCELADA_QTD]) || 0;
      a.cancelada.val += Number(row[COL.CANCELADA_VAL]) || 0;
      a.aAtestar.qtd += Number(row[COL.A_ATESTAR_QTD]) || 0;
      a.aAtestar.val += Number(row[COL.A_ATESTAR_VAL]) || 0;
    }

    var campanhasArr = Object.keys(campanhasMap).map(function(nome) {
      return campanhasMap[nome];
    });

    campanhasArr.sort(function(a, b) {
      return b.ordem - a.ordem;
    });

    return campanhasArr;
  } catch (e) {
    Logger.log('getDadosPorCampanha erro: ' + e.message);
    throw e;
  }
}

// ============================================
// PAGAMENTOS - ABA CONSOLIDADO
// ============================================
/**
 * Busca pagamentos da aba Consolidado por campanha
 * Colunas do Consolidado:
 *   A (0) = Agência
 *   B (1) = Data
 *   C (2) = Campanha
 *   M (12) = Status (PAGA)
 *   T (19) = Controle
 *   AB (27) = Valor em reais
 *
 * Também busca o empenho da aba SaldoEmpenhos (coluna G) para calcular o saldo
 *
 * @param {string} campanha - Nome da campanha
 * @returns {Object} { AV: { pagamentos: [], empenho: 0 }, Calia: {...}, EBM: {...} }
 */
function getPagamentosPorCampanha(campanha) {
  try {
    // Índices fixos das colunas do Consolidado
    var COL_CONSOLIDADO = {
      AGENCIA: 0,    // A
      DATA: 1,       // B (Data consolidado)
      CAMPANHA: 2,   // C
      STATUS: 12,    // M
      ATESTO: 18,    // S (Data atesto)
      CONTROLE: 19,  // T
      DATA_PAGO: 20, // U (Data pago)
      VALOR: 27      // AB
    };

    var resultado = {
      AV: { pagamentos: [], empenho: 0 },
      Calia: { pagamentos: [], empenho: 0 },
      EBM: { pagamentos: [], empenho: 0 }
    };

    var campanhaNorm = normalizeString_(campanha);

    // 1. Buscar empenho da aba SaldoEmpenhos
    var sheetSaldo = getSheet_();
    var valuesSaldo = sheetSaldo.getDataRange().getValues();
    var headersSaldo = valuesSaldo[0];
    var COL = buildColMap_(headersSaldo);

    for (var i = 1; i < valuesSaldo.length; i++) {
      var row = valuesSaldo[i];
      if (normalizeString_(row[COL.CAMPANHA]) !== campanhaNorm) continue;

      var ag = normalizeAgencia_(row[COL.AGENCIA]);
      if (resultado[ag]) {
        resultado[ag].empenho += Number(row[COL.EMPENHO_TOTAL]) || 0;
      }
    }

    // 2. Buscar pagamentos da aba Consolidado
    var sheetConsolidado = getSheetConsolidado_();
    var valuesConsolidado = sheetConsolidado.getDataRange().getValues();

    // Map para agrupar por controle: { agencia: { controle: { ... } } }
    var pagamentosMap = {
      AV: {},
      Calia: {},
      EBM: {}
    };

    for (var i = 1; i < valuesConsolidado.length; i++) {
      var row = valuesConsolidado[i];
      
      // Filtrar apenas linhas onde coluna T (controle) está preenchida
      var controle = String(row[COL_CONSOLIDADO.CONTROLE] || '').trim();
      if (!controle) continue;
      
      var rowCampanha = String(row[COL_CONSOLIDADO.CAMPANHA] || '').trim();
      if (normalizeString_(rowCampanha) !== campanhaNorm) continue;

      var agencia = normalizeAgencia_(row[COL_CONSOLIDADO.AGENCIA]);
      if (!pagamentosMap[agencia]) continue;

      var dataAtesto = formatDateValue_(row[COL_CONSOLIDADO.ATESTO]);
      var dataPago = formatDateValue_(row[COL_CONSOLIDADO.DATA_PAGO]);
      var controle = String(row[COL_CONSOLIDADO.CONTROLE] || '').trim();
      var status = normalizeString_(row[COL_CONSOLIDADO.STATUS]);
      var valor = Number(row[COL_CONSOLIDADO.VALOR]) || 0;
      var paga = (status === 'paga' || status === 'pago');

      var key = controle;

      if (!pagamentosMap[agencia][key]) {
        pagamentosMap[agencia][key] = {
          dataAtesto: dataAtesto,
          dataPago: dataPago,
          controle: controle,
          qtd: 0,
          valor: 0,
          paga: false
        };
      }

      pagamentosMap[agencia][key].qtd += 1;
      pagamentosMap[agencia][key].valor += valor;

      // Atualizar datas se ainda não preenchidas
      if (!pagamentosMap[agencia][key].dataAtesto && dataAtesto) {
        pagamentosMap[agencia][key].dataAtesto = dataAtesto;
      }
      if (!pagamentosMap[agencia][key].dataPago && dataPago) {
        pagamentosMap[agencia][key].dataPago = dataPago;
      }

      // Se qualquer linha do controle está paga, marcar como pago
      if (paga) {
        pagamentosMap[agencia][key].paga = true;
      }
    }

    // 3. Converter map para array ordenado por data
    ['AV', 'Calia', 'EBM'].forEach(function(ag) {
      var pagamentos = [];
      for (var key in pagamentosMap[ag]) {
        pagamentos.push(pagamentosMap[ag][key]);
      }

      // Ordenar por data (mais antigas primeiro)
      pagamentos.sort(function(a, b) {
        if (!a.data) return 1;
        if (!b.data) return -1;
        // Converter dd/mm/yyyy para comparação
        var partsA = a.data.split('/');
        var partsB = b.data.split('/');
        var dateA = partsA.length === 3 ? new Date(partsA[2], partsA[1] - 1, partsA[0]) : new Date(0);
        var dateB = partsB.length === 3 ? new Date(partsB[2], partsB[1] - 1, partsB[0]) : new Date(0);
        return dateA - dateB;
      });

      resultado[ag].pagamentos = pagamentos;
    });

    return resultado;
  } catch (e) {
    Logger.log('getPagamentosPorCampanha erro: ' + e.message);
    return {
      AV: { pagamentos: [], empenho: 0 },
      Calia: { pagamentos: [], empenho: 0 },
      EBM: { pagamentos: [], empenho: 0 }
    };
  }
}

// ============================================
// EMPENHOS - ABA EMPENHOS
// ============================================
/**
 * Busca empenhos da aba Empenhos por campanha e agência
 * Colunas da aba Empenhos:
 *   A (0) = Campanha
 *   B (1) = Agência
 *   C (2) = SEI
 *   D (3) = Empenho
 *   E (4) = Data
 *   F (5) = Valor
 *
 * @param {string} campanha - Nome da campanha
 * @returns {Object} { AV: [ { sei, empenho, data, valor }, ... ], Calia: [...], EBM: [...] }
 */
function getEmpenhosPorCampanha(campanha) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Empenhos');
    if (!sheet) {
      Logger.log('Aba Empenhos não encontrada');
      return { AV: [], Calia: [], EBM: [] };
    }

    var values = sheet.getDataRange().getValues();
    var resultado = { AV: [], Calia: [], EBM: [] };
    var campanhaNorm = normalizeString_(campanha);

    // Pular cabeçalho (linha 0)
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var rowCampanha = String(row[0] || '').trim();

      if (normalizeString_(rowCampanha) !== campanhaNorm) continue;

      var agencia = normalizeAgencia_(row[1]);
      if (!resultado[agencia]) continue;

      var valor = Number(row[5]) || 0;

      resultado[agencia].push({
        sei: String(row[2] || '').trim(),
        empenho: String(row[3] || '').trim(),
        data: formatDateValue_(row[4]),
        valor: valor
      });
    }

    // Ordenar por data
    ['AV', 'Calia', 'EBM'].forEach(function(ag) {
      resultado[ag].sort(function(a, b) {
        if (!a.data) return 1;
        if (!b.data) return -1;
        var partsA = a.data.split('/');
        var partsB = b.data.split('/');
        var dateA = partsA.length === 3 ? new Date(partsA[2], partsA[1] - 1, partsA[0]) : new Date(0);
        var dateB = partsB.length === 3 ? new Date(partsB[2], partsB[1] - 1, partsB[0]) : new Date(0);
        return dateA - dateB;
      });
    });

    return resultado;
  } catch (e) {
    Logger.log('getEmpenhosPorCampanha erro: ' + e.message);
    return { AV: [], Calia: [], EBM: [] };
  }
}

// ============================================
// MENU
// ============================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Dashboard')
    .addItem('Abrir Campanhas', 'abrirDashboard')
    .addItem('Limpar Cache', 'limparCache')
    .addToUi();
}

function abrirDashboard() {
  var html = HtmlService.createHtmlOutputFromFile('Index')
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Dashboard Campanhas');
}
