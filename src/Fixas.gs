/**
 * Fixas.gs — Controle de despesas fixas (obrigações que você paga todo mês) e
 * de parcelamentos longos (> LIMITE_PARCELA_LONGA).
 *
 * Modelo de dados (abas criadas pelo setup):
 *  - FIXAS:           ID | DESCRIÇÃO | CATEGORIA | SUBCATEGORIA | VALOR_BASE | DIA | ATIVO
 *  - FIXAS_PROJECAO:  FIXA_ID | MES (yyyy-MM) | VALOR | STATUS | ATUALIZADO_EM
 *
 * A projeção de 12 meses é CALCULADA na hora (previsto = VALOR_BASE do modelo).
 * A aba FIXAS_PROJECAO guarda apenas os meses que você validou/ajustou
 * (overrides), mantendo tudo leve.
 */

var HEADER_FIXAS = ['ID', 'DESCRIÇÃO', 'CATEGORIA', 'SUBCATEGORIA', 'VALOR_BASE', 'DIA', 'ATIVO'];
var HEADER_PROJ = ['FIXA_ID', 'MES', 'VALOR', 'STATUS', 'ATUALIZADO_EM'];

// ------------------------------------------------------------------
// Acesso às abas
// ------------------------------------------------------------------
function _ensureSheet(nome, header) {
  var ss = _ss();
  var sh = ss.getSheetByName(nome);
  if (!sh) {
    sh = ss.insertSheet(nome);
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, header.length).setValues([header]);
  }
  return sh;
}
function _sheetFixas() { return _ensureSheet(SHEET_FIXAS, HEADER_FIXAS); }
function _sheetProjecao() { return _ensureSheet(SHEET_FIXAS_PROJ, HEADER_PROJ); }

/** Lê uma aba simples (cabeçalho na linha 1) como array de objetos + índices. */
function _lerTabela(sheet, header) {
  var lc = header.length;
  var lr = sheet.getLastRow();
  var col = {}; header.forEach(function (h, i) { col[h] = i; });
  var rows = [];
  if (lr >= 2) {
    var vals = sheet.getRange(2, 1, lr - 1, lc).getValues();
    for (var i = 0; i < vals.length; i++) rows.push({ row: i + 2, vals: vals[i] });
  }
  return { col: col, rows: rows };
}

// ------------------------------------------------------------------
// Modelos de despesa fixa
// ------------------------------------------------------------------

/** Cria/atualiza um modelo de despesa fixa (dedup por descrição normalizada). */
function upsertFixa(p) {
  var descricao = String(p.descricao || '').trim();
  var categoria = String(p.categoria || '').trim();
  var subcategoria = String(p.subcategoria || '').trim();
  var valorBase = _round2(Number(String(p.valorBase).replace(',', '.')) || 0);
  var dia = Math.min(28, Math.max(1, parseInt(p.dia, 10) || 1));
  var ativo = p.ativo === false ? false : true;
  if (!descricao) throw new Error('Informe a descrição da despesa fixa.');

  var sh = _sheetFixas();
  var t = _lerTabela(sh, HEADER_FIXAS);
  var alvo = _norm(descricao);
  var existente = t.rows.find(function (r) { return _norm(r.vals[t.col['DESCRIÇÃO']]) === alvo; });

  if (existente) {
    var r = existente.row;
    sh.getRange(r, t.col['CATEGORIA'] + 1).setValue(categoria);
    sh.getRange(r, t.col['SUBCATEGORIA'] + 1).setValue(subcategoria);
    if (valorBase > 0) sh.getRange(r, t.col['VALOR_BASE'] + 1).setValue(valorBase);
    sh.getRange(r, t.col['DIA'] + 1).setValue(dia);
    sh.getRange(r, t.col['ATIVO'] + 1).setValue(ativo);
    return String(existente.vals[t.col['ID']]);
  }
  var id = 'F' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  sh.appendRow([id, descricao, categoria, subcategoria, valorBase, dia, ativo]);
  return id;
}

/** Adiciona/edita modelo pela tela de Fixas. */
function addFixaManual(payload) {
  var id = upsertFixa(payload);
  return { ok: true, id: id };
}

function updateFixa(id, payload) {
  var sh = _sheetFixas();
  var t = _lerTabela(sh, HEADER_FIXAS);
  var alvo = t.rows.find(function (r) { return String(r.vals[t.col['ID']]) === String(id); });
  if (!alvo) throw new Error('Despesa fixa não encontrada.');
  var r = alvo.row;
  if (payload.descricao != null) sh.getRange(r, t.col['DESCRIÇÃO'] + 1).setValue(String(payload.descricao).trim());
  if (payload.categoria != null) sh.getRange(r, t.col['CATEGORIA'] + 1).setValue(String(payload.categoria).trim());
  if (payload.subcategoria != null) sh.getRange(r, t.col['SUBCATEGORIA'] + 1).setValue(String(payload.subcategoria).trim());
  if (payload.valorBase != null) sh.getRange(r, t.col['VALOR_BASE'] + 1).setValue(_round2(Number(String(payload.valorBase).replace(',', '.')) || 0));
  if (payload.dia != null) sh.getRange(r, t.col['DIA'] + 1).setValue(Math.min(28, Math.max(1, parseInt(payload.dia, 10) || 1)));
  if (payload.ativo != null) sh.getRange(r, t.col['ATIVO'] + 1).setValue(payload.ativo === true || payload.ativo === 'true');
  return { ok: true };
}

function toggleFixaAtivo(id, ativo) {
  return updateFixa(id, { ativo: !!ativo });
}

/** Exclui um modelo e suas validações. */
function deleteFixa(id) {
  var sh = _sheetFixas();
  var t = _lerTabela(sh, HEADER_FIXAS);
  var alvo = t.rows.find(function (r) { return String(r.vals[t.col['ID']]) === String(id); });
  if (alvo) sh.deleteRow(alvo.row);
  // remove overrides
  var shp = _sheetProjecao();
  var tp = _lerTabela(shp, HEADER_PROJ);
  var linhas = tp.rows.filter(function (r) { return String(r.vals[tp.col['FIXA_ID']]) === String(id); })
    .map(function (r) { return r.row; });
  for (var i = linhas.length - 1; i >= 0; i--) shp.deleteRow(linhas[i]);
  return { ok: true };
}

// ------------------------------------------------------------------
// Validação por mês (overrides)
// ------------------------------------------------------------------

function _acharProjRow(tp, shp, fixaId, mes) {
  return tp.rows.find(function (r) {
    return String(r.vals[tp.col['FIXA_ID']]) === String(fixaId) && String(r.vals[tp.col['MES']]) === String(mes);
  });
}

/** Valida (confirma/ajusta) o valor de uma despesa fixa num mês. */
function validarFixa(fixaId, mes, valor) {
  if (!fixaId || !mes) throw new Error('Dados incompletos.');
  var v = _round2(Number(String(valor).replace(',', '.')) || 0);
  var shp = _sheetProjecao();
  var tp = _lerTabela(shp, HEADER_PROJ);
  var existente = _acharProjRow(tp, shp, fixaId, mes);
  var agora = new Date();
  if (existente) {
    shp.getRange(existente.row, tp.col['VALOR'] + 1).setValue(v);
    shp.getRange(existente.row, tp.col['STATUS'] + 1).setValue('validado');
    shp.getRange(existente.row, tp.col['ATUALIZADO_EM'] + 1).setValue(agora);
  } else {
    shp.appendRow([fixaId, mes, v, 'validado', agora]);
  }
  return { ok: true, valor: v };
}

/** Remove a validação de um mês (volta a "previsto"). */
function desvalidarFixa(fixaId, mes) {
  var shp = _sheetProjecao();
  var tp = _lerTabela(shp, HEADER_PROJ);
  var existente = _acharProjRow(tp, shp, fixaId, mes);
  if (existente) shp.deleteRow(existente.row);
  return { ok: true };
}

// ------------------------------------------------------------------
// Painel consolidado (12 meses)
// ------------------------------------------------------------------

/**
 * Retorna o controle completo de fixas para os próximos 12 meses.
 * @return {Object} { mesAtual, meses:[{mes,label,totalFixas,totalParcelas,total,fixas:[...],parcelas:[...]}], templates:[...] }
 */
function getControleFixas() {
  var shf = _sheetFixas();
  var tf = _lerTabela(shf, HEADER_FIXAS);
  var templates = tf.rows.map(function (r) {
    return {
      id: String(r.vals[tf.col['ID']]),
      descricao: String(r.vals[tf.col['DESCRIÇÃO']] || ''),
      categoria: String(r.vals[tf.col['CATEGORIA']] || ''),
      subcategoria: String(r.vals[tf.col['SUBCATEGORIA']] || ''),
      valorBase: _round2(r.vals[tf.col['VALOR_BASE']] || 0),
      dia: parseInt(r.vals[tf.col['DIA']], 10) || 1,
      ativo: r.vals[tf.col['ATIVO']] === true || _norm(r.vals[tf.col['ATIVO']]) === 'TRUE' || r.vals[tf.col['ATIVO']] === 'VERDADEIRO'
    };
  });

  // overrides de validação
  var shp = _sheetProjecao();
  var tp = _lerTabela(shp, HEADER_PROJ);
  var overrides = {};
  tp.rows.forEach(function (r) {
    var k = String(r.vals[tp.col['FIXA_ID']]) + '|' + String(r.vals[tp.col['MES']]);
    overrides[k] = { valor: _round2(r.vals[tp.col['VALOR']] || 0), status: String(r.vals[tp.col['STATUS']] || 'validado') };
  });

  // parcelas longas (> LIMITE) por mês, a partir da aba DADOS
  var parcelasPorMes = _parcelasLongasPorMes(_lerDados());

  // janela de 12 meses a partir do mês atual
  var hoje = new Date();
  var meses = [];
  for (var i = 0; i < 12; i++) {
    var d = new Date(hoje.getFullYear(), hoje.getMonth() + i, 1);
    var mes = _chaveMes(d);
    var fixasMes = [];
    var totalFixas = 0;
    templates.filter(function (t) { return t.ativo; }).forEach(function (t) {
      var ov = overrides[t.id + '|' + mes];
      var previsto = t.valorBase;
      var validado = ov ? ov.valor : null;
      var valor = ov ? ov.valor : previsto;
      totalFixas += valor;
      fixasMes.push({
        fixaId: t.id, descricao: t.descricao, categoria: t.categoria, subcategoria: t.subcategoria,
        previsto: previsto, validado: validado, status: ov ? 'validado' : 'previsto', valor: _round2(valor), dia: t.dia
      });
    });
    fixasMes.sort(function (a, b) { return b.valor - a.valor; });

    var parcelas = parcelasPorMes[mes] || [];
    var totalParcelas = parcelas.reduce(function (a, p) { return a + p.valor; }, 0);

    meses.push({
      mes: mes, label: _rotuloMesLongo(d),
      totalFixas: _round2(totalFixas), totalParcelas: _round2(totalParcelas), total: _round2(totalFixas + totalParcelas),
      fixas: fixasMes, parcelas: parcelas
    });
  }

  return { mesAtual: _chaveMes(hoje), meses: meses, templates: templates };
}

/** Agrupa parcelas longas (total > LIMITE) por mês, lendo a aba DADOS. */
function _parcelasLongasPorMes(linhas) {
  var mapa = {};
  linhas.forEach(function (l) {
    if (l.valor >= 0) return; // só despesas
    var m = String(l.descricao).match(/\((\d+)\s*\/\s*(\d+)\)\s*$/);
    if (!m) return;
    var total = parseInt(m[2], 10);
    if (!(total > LIMITE_PARCELA_LONGA)) return;
    if (!mapa[l.mes]) mapa[l.mes] = [];
    mapa[l.mes].push({
      descricao: l.descricao, categoria: l.categoria, subcategoria: l.subcategoria,
      valor: l.abs, parcela: parseInt(m[1], 10), total: total
    });
  });
  Object.keys(mapa).forEach(function (m) { mapa[m].sort(function (a, b) { return b.valor - a.valor; }); });
  return mapa;
}

/** 'MMMM/yyyy' capitalizado para um Date. */
function _rotuloMesLongo(d) {
  var s = Utilities.formatDate(d, TZ, 'MMM/yy');
  return s.charAt(0).toUpperCase() + s.slice(1);
}
