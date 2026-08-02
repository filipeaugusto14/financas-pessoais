/**
 * Config.gs — Configuração central, utilidades e detecção automática de abas.
 *
 * A planilha tem 3 abas com papéis distintos. Em vez de depender dos nomes
 * exatos, o app detecta cada aba pela "assinatura" das colunas. Se por algum
 * motivo a detecção falhar, preencha os OVERRIDES abaixo com os nomes exatos.
 */

// ID da planilha "Finanças pessoais".
const SPREADSHEET_ID = '1ycgVNO2wLI4OSM25Enwfif3kLqQrorL-ulrp7E7JMcY';

// Nome exibido na interface.
const APP_NOME = 'Minhas Finanças';

// Fuso usado para cálculos/formatação de datas.
const TZ = 'America/Sao_Paulo';

/**
 * OVERRIDES (opcionais). Deixe '' para detecção automática.
 * Ex.: SHEET_OVERRIDES.entrada = 'Lançamentos';
 */
const SHEET_OVERRIDES = {
  entrada: '',    // aba onde você digita (tem a coluna "ENVIADO PARA DADOS")
  dados: '',      // aba consolidada que o relatório lê (valores com sinal)
  taxonomia: ''   // aba de referência SUBCATEGORIA -> CATEGORIA -> TIPO
};

// Nome da coluna técnica adicionada pelo app (para editar/excluir com segurança).
const COL_ID = 'ID';

// Abas gerenciadas pelo app (criadas pelo setup) para o controle de despesas fixas.
const SHEET_FIXAS = 'FIXAS';            // modelos de despesa fixa
const SHEET_FIXAS_PROJ = 'FIXAS_PROJECAO'; // validações/ajustes por mês
const LIMITE_PARCELA_LONGA = 12;        // parcelamentos ACIMA disso contam como obrigação longa

// ------------------------------------------------------------------
// Utilidades gerais
// ------------------------------------------------------------------

/** Abre a planilha alvo. */
function _ss() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/** Normaliza texto: sem acento, maiúsculas, trim. Útil para comparar cabeçalhos. */
function _norm(v) {
  return String(v == null ? '' : v)
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .trim().toUpperCase();
}

/** Arredonda para centavos (2 casas), evitando erros de ponto flutuante. */
function _round2(n) {
  return Math.round((Number(n) + Number.EPSILON) * 100) / 100;
}

/** Cria um ID único para vincular Lançamento <-> linhas da aba DADOS. */
function _novoId() {
  return 'L' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

/**
 * Converte um valor de célula/entrada em Date LOCAL.
 * Aceita Date, ISO 'yyyy-MM-dd' (do <input type=date>) e 'dd/MM/yyyy'.
 * IMPORTANTE: 'yyyy-MM-dd' é tratado como data local — nunca via new Date(str),
 * que interpretaria como UTC e voltaria 1 dia no fuso de São Paulo.
 */
function _paraData(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === 'string') {
    var s = v.trim();
    var iso = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (iso) {
      return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
    }
    var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
    if (m) {
      var ano = m[3].length === 2 ? 2000 + Number(m[3]) : Number(m[3]);
      return new Date(ano, Number(m[2]) - 1, Number(m[1]));
    }
    var d = new Date(s);
    if (!isNaN(d)) return d;
  }
  return null;
}

/** Soma meses a uma data, mantendo o dia (com ajuste para meses curtos). */
function _addMeses(data, n) {
  var dia = data.getDate();
  var alvo = new Date(data.getFullYear(), data.getMonth() + n, 1);
  var ultimoDia = new Date(alvo.getFullYear(), alvo.getMonth() + 1, 0).getDate();
  alvo.setDate(Math.min(dia, ultimoDia));
  return alvo;
}

/** Chave de mês "yyyy-MM" para agrupamentos. */
function _chaveMes(data) {
  return Utilities.formatDate(data, TZ, 'yyyy-MM');
}

/** Converte número em texto "R$ 1.234,56". */
function _brl(n) {
  var s = (Math.abs(Number(n)) || 0).toFixed(2).replace('.', ',');
  s = s.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return (n < 0 ? '-R$ ' : 'R$ ') + s;
}

// ------------------------------------------------------------------
// Detecção de abas
// ------------------------------------------------------------------

/**
 * Retorna { entrada, dados, taxonomia } com objetos Sheet.
 * A detecção é feita a cada chamada (barata) e respeita os OVERRIDES.
 */
function _detectarAbas() {
  var ss = _ss();
  var abas = ss.getSheets();
  var info = abas.map(function (sh) {
    var head = _cabecalho(sh);
    return { sheet: sh, nome: sh.getName(), head: head, headNorm: head.map(_norm) };
  });

  function porNome(nome) {
    if (!nome) return null;
    var alvo = _norm(nome);
    for (var i = 0; i < info.length; i++) if (_norm(info[i].nome) === alvo) return info[i].sheet;
    return null;
  }
  function temToken(headNorm, token) {
    return headNorm.some(function (h) { return h.indexOf(token) !== -1; });
  }

  // 1) Overrides explícitos
  var entrada = porNome(SHEET_OVERRIDES.entrada);
  var dados = porNome(SHEET_OVERRIDES.dados);
  var taxonomia = porNome(SHEET_OVERRIDES.taxonomia);

  // 2) Entrada: cabeçalho contém "ENVIADO" (de "ENVIADO PARA DADOS")
  if (!entrada) {
    var e = info.filter(function (x) { return temToken(x.headNorm, 'ENVIADO'); });
    if (e.length) entrada = e[0].sheet;
  }

  // 3) DADOS: aba chamada "DADOS", senão aba (não-entrada) com DATA+VALOR e sem ENVIADO
  if (!dados) dados = porNome('DADOS');
  if (!dados) {
    var d = info.filter(function (x) {
      if (entrada && x.sheet.getSheetId() === entrada.getSheetId()) return false;
      return temToken(x.headNorm, 'DATA') && temToken(x.headNorm, 'VALOR') &&
             !temToken(x.headNorm, 'ENVIADO');
    });
    if (d.length) dados = d[0].sheet;
  }
  // 3b) Fallback: aba (não-entrada) cujo corpo começa com datas na col 1 e números na col 2
  if (!dados) {
    var d2 = info.filter(function (x) {
      if (entrada && x.sheet.getSheetId() === entrada.getSheetId()) return false;
      return _pareceDados(x.sheet);
    });
    if (d2.length) dados = d2[0].sheet;
  }

  // 4) Taxonomia: sobra que contém a coluna TIPO com Receita/Despesa
  if (!taxonomia) {
    var usados = {};
    if (entrada) usados[entrada.getSheetId()] = 1;
    if (dados) usados[dados.getSheetId()] = 1;
    var candidatos = info.filter(function (x) { return !usados[x.sheet.getSheetId()]; });
    // preferir a que tem valores "DESPESA"/"RECEITA" no corpo
    var comTipo = candidatos.filter(function (x) { return _temValoresTipo(x.sheet); });
    taxonomia = (comTipo[0] && comTipo[0].sheet) || (candidatos[0] && candidatos[0].sheet) || null;
  }

  if (!entrada) throw new Error('Não encontrei a aba de entrada (com a coluna "ENVIADO PARA DADOS").');
  if (!dados) throw new Error('Não encontrei a aba DADOS.');

  return { entrada: entrada, dados: dados, taxonomia: taxonomia };
}

/** Lê a primeira linha não vazia como cabeçalho (procura nas 3 primeiras linhas). */
function _cabecalho(sheet) {
  var maxC = Math.max(1, sheet.getLastColumn());
  var linhas = Math.min(3, Math.max(1, sheet.getLastRow()));
  if (linhas < 1) return [];
  var vals = sheet.getRange(1, 1, linhas, maxC).getValues();
  // escolhe a linha com mais células preenchidas
  var melhor = vals[0], melhorN = -1;
  for (var i = 0; i < vals.length; i++) {
    var n = vals[i].filter(function (c) { return String(c).trim() !== ''; }).length;
    if (n > melhorN) { melhorN = n; melhor = vals[i]; }
  }
  return melhor;
}

/** Verifica se a aba tem "DESPESA"/"RECEITA" em alguma célula (marca de taxonomia). */
function _temValoresTipo(sheet) {
  var lr = Math.min(sheet.getLastRow(), 60);
  var lc = Math.min(sheet.getLastColumn(), 12);
  if (lr < 1 || lc < 1) return false;
  var vals = sheet.getRange(1, 1, lr, lc).getValues();
  var achouD = false, achouR = false;
  for (var r = 0; r < vals.length; r++) for (var c = 0; c < vals[r].length; c++) {
    var v = _norm(vals[r][c]);
    if (v === 'DESPESA') achouD = true;
    if (v === 'RECEITA') achouR = true;
  }
  return achouD && achouR;
}

/** Heurística: a aba parece a DADOS? (col 1 com datas, col 2 com números). */
function _pareceDados(sheet) {
  var lr = Math.min(sheet.getLastRow(), 12);
  if (lr < 1 || sheet.getLastColumn() < 2) return false;
  var vals = sheet.getRange(1, 1, lr, 2).getValues();
  var datas = 0, nums = 0;
  for (var i = 0; i < vals.length; i++) {
    if (_paraData(vals[i][0])) datas++;
    if (typeof vals[i][1] === 'number' && vals[i][1] !== '') nums++;
  }
  return datas >= 2 && nums >= 2;
}

/**
 * Mapa de índices de coluna (1-based) por nome normalizado, para uma aba
 * que tenha cabeçalho. Retorna { headerRow, firstDataRow, col: {NOME: idx}, ncols }.
 */
function _mapaColunas(sheet) {
  var maxC = Math.max(1, sheet.getLastColumn());
  var linhas = Math.min(3, Math.max(1, sheet.getLastRow()));
  var vals = sheet.getRange(1, 1, linhas, maxC).getValues();
  var headerRow = -1, header = vals[0];
  for (var i = 0; i < vals.length; i++) {
    var row = vals[i].map(_norm);
    if (row.indexOf('DATA') !== -1 && row.indexOf('VALOR') !== -1) {
      headerRow = i + 1; header = vals[i]; break;
    }
  }
  var col = {};
  if (headerRow > 0) {
    header.forEach(function (h, i) {
      var k = _norm(h);
      if (k) col[k] = i + 1;
    });
    return { headerFound: true, headerRow: headerRow, firstDataRow: headerRow + 1, col: col, header: header, ncols: maxC };
  }
  // Sem cabeçalho reconhecível: dados começam na linha 1; colunas por posição.
  return { headerFound: false, headerRow: 1, firstDataRow: 1, col: {}, header: header, ncols: maxC };
}

/** Garante que a aba tenha a coluna ID (cria ao final se faltar). Retorna o índice. */
function _garantirColunaId(sheet) {
  var m = _mapaColunas(sheet);
  if (m.col[COL_ID]) return m.col[COL_ID];
  if (m.headerFound) {
    var novaCol = sheet.getLastColumn() + 1;
    sheet.getRange(m.headerRow, novaCol).setValue(COL_ID);
    return novaCol;
  }
  // Sem cabeçalho (ex.: DADOS sem título de coluna): ID por convenção na coluna 7,
  // sem escrever rótulo em cima de uma linha de dados.
  return Math.max(7, sheet.getLastColumn() + 1);
}

// ------------------------------------------------------------------
// setup() e diagnostico()
// ------------------------------------------------------------------

/**
 * Prepara a planilha para o app. Idempotente — pode rodar quantas vezes quiser.
 * - Confirma as abas
 * - Garante a coluna ID nas abas de entrada e DADOS
 */
function setup() {
  var abas = _detectarAbas();
  _garantirColunaId(abas.entrada);
  _garantirColunaId(abas.dados);
  _sheetFixas();      // cria a aba FIXAS se faltar
  _sheetProjecao();   // cria a aba FIXAS_PROJECAO se faltar
  var msg = 'setup OK\n' +
    '- Entrada: ' + abas.entrada.getName() + '\n' +
    '- DADOS:   ' + abas.dados.getName() + '\n' +
    '- Taxonomia: ' + (abas.taxonomia ? abas.taxonomia.getName() : '(não encontrada)') + '\n' +
    '- ' + SHEET_FIXAS + ' e ' + SHEET_FIXAS_PROJ + ' prontas.';
  Logger.log(msg);
  return msg;
}

/** Diagnóstico: mostra as abas detectadas e seus cabeçalhos. Rode e veja o log. */
function diagnostico() {
  var ss = _ss();
  var out = { planilha: ss.getName(), abas: [] };
  ss.getSheets().forEach(function (sh) {
    out.abas.push({ nome: sh.getName(), linhas: sh.getLastRow(), colunas: sh.getLastColumn(), cabecalho: _cabecalho(sh) });
  });
  try {
    var det = _detectarAbas();
    out.detectado = {
      entrada: det.entrada.getName(),
      dados: det.dados.getName(),
      taxonomia: det.taxonomia ? det.taxonomia.getName() : null
    };
  } catch (e) {
    out.erro = e.message;
  }
  Logger.log(JSON.stringify(out, null, 2));
  return out;
}
