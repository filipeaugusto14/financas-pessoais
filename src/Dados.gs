/**
 * Dados.gs — Geração da aba DADOS (consolidado que o relatório/Looker lê) e
 * helpers de baixo nível para ler/escrever nas abas.
 *
 * Regras da aba DADOS:
 *  - VALOR com sinal: despesa negativa, receita positiva.
 *  - Parcelas explodidas: 1 linha por mês (data +1 mês a cada parcela),
 *    descrição com sufixo "(n/total)".
 *  - Arredondamento em centavos; a diferença cai na última parcela.
 */

// ------------------------------------------------------------------
// Resolução de colunas (por nome, com fallback por posição)
// ------------------------------------------------------------------

/** Índices (1-based) das colunas da aba de entrada. */
function _colsEntrada(sheet) {
  var m = _mapaColunas(sheet);
  var c = m.col;
  function pick(nome, pos) { return c[nome] || pos; }
  var enviado = c['ENVIADO PARA DADOS'];
  if (!enviado) { // procura por token ENVIADO
    Object.keys(c).forEach(function (k) { if (k.indexOf('ENVIADO') !== -1) enviado = c[k]; });
  }
  return {
    map: m,
    data: pick('DATA', 1),
    valor: pick('VALOR', 2),
    desc: pick('DESCRICAO', 3),
    subcat: pick('SUBCATEGORIA', 4),
    cat: pick('CATEGORIA', 5),
    tipo: pick('TIPO', 6),
    parcelas: pick('PARCELAS', 7),
    enviado: enviado || 8,
    id: c[COL_ID] || 0
  };
}

/** Índices (1-based) das colunas da aba DADOS. */
function _colsDados(sheet) {
  var m = _mapaColunas(sheet);
  var c = m.col;
  function pick(nome, pos) { return c[nome] || pos; }
  return {
    map: m,
    data: pick('DATA', 1),
    valor: pick('VALOR', 2),
    desc: pick('DESCRICAO', 3),
    subcat: pick('SUBCATEGORIA', 4),
    cat: pick('CATEGORIA', 5),
    tipo: pick('TIPO', 6),
    // ID pela coluna rotulada; se não houver rótulo, por convenção coluna 7.
    id: c[COL_ID] || 7
  };
}

// ------------------------------------------------------------------
// Explosão de parcelas
// ------------------------------------------------------------------

/**
 * Gera as linhas (objetos) que representam um lançamento na aba DADOS.
 * @param {Object} lanc {data:Date, valor:Number(+), descricao, subcategoria, categoria, tipo, parcelas, id}
 * @return {Array<Object>} cada item { data, valor(sinal), descricao, subcategoria, categoria, tipo, id }
 */
function _gerarLinhasDados(lanc) {
  var sinal = _norm(lanc.tipo) === 'RECEITA' ? 1 : -1;
  var p = Math.max(1, parseInt(lanc.parcelas, 10) || 1);
  var total = _round2(lanc.valor);
  var linhas = [];

  if (p <= 1) {
    linhas.push(_linhaDados(lanc, lanc.data, sinal * total, lanc.descricao));
    return linhas;
  }

  var base = _round2(total / p);
  var ultima = _round2(total - base * (p - 1)); // absorve a diferença de arredondamento
  for (var i = 1; i <= p; i++) {
    var valorParcela = (i < p ? base : ultima);
    var data = _addMeses(lanc.data, i - 1);
    var desc = lanc.descricao + ' (' + i + '/' + p + ')';
    linhas.push(_linhaDados(lanc, data, sinal * valorParcela, desc));
  }
  return linhas;
}

function _linhaDados(lanc, data, valor, descricao) {
  return {
    data: data,
    valor: valor,
    descricao: descricao,
    subcategoria: lanc.subcategoria,
    categoria: lanc.categoria,
    tipo: lanc.tipo,
    id: lanc.id
  };
}

/** Anexa várias linhas à aba DADOS de uma vez. */
function _appendLinhasDados(sheet, linhas) {
  if (!linhas.length) return;
  var C = _colsDados(sheet);
  var maxCol = Math.max(C.data, C.valor, C.desc, C.subcat, C.cat, C.tipo, C.id || 0);
  var matriz = linhas.map(function (l) {
    var row = new Array(maxCol).fill('');
    row[C.data - 1] = l.data;
    row[C.valor - 1] = l.valor;
    row[C.desc - 1] = l.descricao;
    row[C.subcat - 1] = l.subcategoria;
    row[C.cat - 1] = l.categoria;
    row[C.tipo - 1] = l.tipo;
    if (C.id) row[C.id - 1] = l.id;
    return row;
  });
  var inicio = sheet.getLastRow() + 1;
  sheet.getRange(inicio, 1, matriz.length, maxCol).setValues(matriz);
}

/** Remove todas as linhas da aba DADOS que pertençam a um lançamento (por ID). */
function _removerLinhasDadosPorId(sheet, id) {
  var C = _colsDados(sheet);
  if (!C.id) return 0;
  var ult = sheet.getLastRow();
  if (ult < C.map.firstDataRow) return 0;
  var n = ult - C.map.firstDataRow + 1;
  var ids = sheet.getRange(C.map.firstDataRow, C.id, n, 1).getValues();
  var aRemover = [];
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) aRemover.push(C.map.firstDataRow + i);
  }
  // remove de baixo para cima
  for (var j = aRemover.length - 1; j >= 0; j--) sheet.deleteRow(aRemover[j]);
  return aRemover.length;
}
