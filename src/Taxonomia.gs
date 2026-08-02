/**
 * Taxonomia.gs — Lê a aba de referência SUBCATEGORIA -> CATEGORIA -> TIPO
 * para alimentar os menus do formulário, e permite cadastrar novas
 * subcategorias. Também gera sugestões de descrição a partir do histórico.
 */

/**
 * Estrutura para o formulário.
 * @return {Object} {
 *   lista: [{subcategoria, categoria, tipo}],
 *   porTipo: { Despesa: {Categoria: [subs...]}, Receita: {...} },
 *   catTipo: { Categoria: 'Despesa'|'Receita' },
 *   subInfo: { subcategoria: {categoria, tipo} }
 * }
 */
function getTaxonomia() {
  var abas = _detectarAbas();
  var sh = abas.taxonomia;
  var resultado = { lista: [], porTipo: { Despesa: {}, Receita: {} }, catTipo: {}, subInfo: {} };
  if (!sh) return resultado;

  var lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 1 || lc < 1) return resultado;
  var vals = sh.getRange(1, 1, lr, lc).getValues();

  var loc = _localizarColunasTaxonomia(vals);
  if (!loc) return resultado;

  var vistos = {};
  for (var r = 0; r < vals.length; r++) {
    var tipoRaw = _norm(vals[r][loc.tipoCol]);
    if (tipoRaw !== 'DESPESA' && tipoRaw !== 'RECEITA') continue;
    var sub = String(vals[r][loc.subCol] || '').trim();
    var cat = String(vals[r][loc.catCol] || '').trim();
    if (!sub || !cat) continue;
    var tipo = tipoRaw === 'RECEITA' ? 'Receita' : 'Despesa';

    var chave = tipo + '|' + cat + '|' + sub;
    if (vistos[chave]) continue;
    vistos[chave] = true;

    resultado.lista.push({ subcategoria: sub, categoria: cat, tipo: tipo });
    resultado.catTipo[cat] = tipo;
    resultado.subInfo[sub] = { categoria: cat, tipo: tipo };
    if (!resultado.porTipo[tipo][cat]) resultado.porTipo[tipo][cat] = [];
    if (resultado.porTipo[tipo][cat].indexOf(sub) === -1) resultado.porTipo[tipo][cat].push(sub);
  }
  return resultado;
}

/**
 * Localiza as colunas SUBCATEGORIA/CATEGORIA/TIPO na aba de taxonomia.
 * Estratégia: acha a coluna cujo corpo contém "DESPESA"/"RECEITA" (coluna TIPO);
 * as duas colunas à esquerda são CATEGORIA e SUBCATEGORIA (layout SUB | CAT | TIPO).
 * @return {Object|null} { tipoCol, catCol, subCol } (0-based)
 */
function _localizarColunasTaxonomia(vals) {
  var nCols = vals[0] ? vals[0].length : 0;
  var contagem = new Array(nCols).fill(0);
  for (var r = 0; r < vals.length; r++) {
    for (var c = 0; c < nCols; c++) {
      var v = _norm(vals[r][c]);
      if (v === 'DESPESA' || v === 'RECEITA') contagem[c]++;
    }
  }
  var tipoCol = -1, melhor = 0;
  for (var i = 0; i < nCols; i++) if (contagem[i] > melhor) { melhor = contagem[i]; tipoCol = i; }
  if (tipoCol < 2) {
    // fallback: talvez layout CATEGORIA | SUBCATEGORIA sem coluna à esquerda suficiente
    if (tipoCol === 1) return { tipoCol: tipoCol, catCol: 0, subCol: 0 };
    return null;
  }
  return { tipoCol: tipoCol, catCol: tipoCol - 1, subCol: tipoCol - 2 };
}

/**
 * Cadastra uma nova subcategoria na taxonomia (para reaparecer nos menus).
 * @return {Object} {ok}
 */
function addSubcategoria(subcategoria, categoria, tipo) {
  subcategoria = String(subcategoria || '').trim();
  categoria = String(categoria || '').trim();
  tipo = _norm(tipo) === 'RECEITA' ? 'Receita' : 'Despesa';
  if (!subcategoria || !categoria) throw new Error('Informe subcategoria e categoria.');

  var abas = _detectarAbas();
  var sh = abas.taxonomia;
  if (!sh) throw new Error('Aba de taxonomia não encontrada.');

  var lr = sh.getLastRow(), lc = sh.getLastColumn();
  var vals = sh.getRange(1, 1, Math.max(1, lr), Math.max(1, lc)).getValues();
  var loc = _localizarColunasTaxonomia(vals) || { subCol: 0, catCol: 1, tipoCol: 2 };

  var novaLinha = lr + 1;
  sh.getRange(novaLinha, loc.subCol + 1).setValue(subcategoria);
  sh.getRange(novaLinha, loc.catCol + 1).setValue(categoria);
  sh.getRange(novaLinha, loc.tipoCol + 1).setValue(tipo);
  return { ok: true };
}

/**
 * Sugestões de descrição a partir do histórico da aba de entrada.
 * Cada descrição única traz a última categoria/subcategoria/tipo usada e a
 * frequência — usado para autocompletar e autopreencher o formulário.
 * @return {Array<Object>} [{descricao, subcategoria, categoria, tipo, parcelas, count}]
 */
function getSugestoesDescricao() {
  var abas = _detectarAbas();
  var sh = abas.entrada;
  var E = _colsEntrada(sh);
  var ult = sh.getLastRow();
  if (ult < E.map.firstDataRow) return [];
  var n = ult - E.map.firstDataRow + 1;
  var maxCol = Math.max(E.desc, E.subcat, E.cat, E.tipo, E.parcelas);
  var vals = sh.getRange(E.map.firstDataRow, 1, n, maxCol).getValues();

  var mapa = {};
  for (var i = 0; i < vals.length; i++) {
    var desc = String(vals[i][E.desc - 1] || '').trim();
    if (!desc) continue;
    var k = _norm(desc);
    if (!mapa[k]) {
      mapa[k] = {
        descricao: desc,
        subcategoria: String(vals[i][E.subcat - 1] || '').trim(),
        categoria: String(vals[i][E.cat - 1] || '').trim(),
        tipo: String(vals[i][E.tipo - 1] || '').trim() || 'Despesa',
        parcelas: parseInt(vals[i][E.parcelas - 1], 10) || 1,
        count: 0
      };
    }
    // linhas mais recentes sobrescrevem categoria/subcategoria (estão no fim da planilha)
    mapa[k].subcategoria = String(vals[i][E.subcat - 1] || '').trim() || mapa[k].subcategoria;
    mapa[k].categoria = String(vals[i][E.cat - 1] || '').trim() || mapa[k].categoria;
    mapa[k].tipo = String(vals[i][E.tipo - 1] || '').trim() || mapa[k].tipo;
    mapa[k].count++;
  }
  var arr = Object.keys(mapa).map(function (k) { return mapa[k]; });
  arr.sort(function (a, b) { return b.count - a.count; });
  return arr.slice(0, 200);
}
