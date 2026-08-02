/**
 * Fixas.gs — Controle de despesas fixas (obrigações que você paga sempre) e de
 * parcelamentos longos (> LIMITE_PARCELA_LONGA).
 *
 * Abas (criadas pelo setup):
 *  - FIXAS:  ID | DESCRIÇÃO | CATEGORIA | SUBCATEGORIA | VALOR_BASE | DIA | ATIVO | FREQUENCIA | MES_BASE
 *  - FIXAS_PROJECAO: FIXA_ID | MES (yyyy-MM) | VALOR | STATUS | ATUALIZADO_EM
 *
 * FREQUENCIA: 'mensal' (projeta todo mês) ou 'anual' (projeta só no mês MES_BASE, 1-12).
 * A projeção de 12 meses é calculada na hora; a FIXAS_PROJECAO guarda só os meses
 * que você validou/ajustou (overrides).
 */

var HEADER_FIXAS = ['ID', 'DESCRIÇÃO', 'CATEGORIA', 'SUBCATEGORIA', 'VALOR_BASE', 'DIA', 'ATIVO', 'FREQUENCIA', 'MES_BASE'];
var HEADER_PROJ = ['FIXA_ID', 'MES', 'VALOR', 'STATUS', 'ATUALIZADO_EM'];

// ------------------------------------------------------------------
// Acesso às abas (com migração automática de colunas)
// ------------------------------------------------------------------
function _ensureSheet(nome, header) {
  var ss = _ss();
  var sh = ss.getSheetByName(nome);
  if (!sh) sh = ss.insertSheet(nome);
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sh.setFrozenRows(1);
    return sh;
  }
  // migração: garante que todas as colunas do header existam
  var lc = Math.max(sh.getLastColumn(), 1);
  var atual = sh.getRange(1, 1, 1, lc).getValues()[0].map(_norm);
  var faltantes = header.filter(function (h) { return atual.indexOf(_norm(h)) === -1; });
  if (faltantes.length) sh.getRange(1, lc + 1, 1, faltantes.length).setValues([faltantes]);
  return sh;
}
function _sheetFixas() { return _ensureSheet(SHEET_FIXAS, HEADER_FIXAS); }
function _sheetProjecao() { return _ensureSheet(SHEET_FIXAS_PROJ, HEADER_PROJ); }

/** Lê uma aba simples (cabeçalho na linha 1) mapeando colunas por NOME normalizado. */
function _lerTabela(sheet) {
  var lc = Math.max(sheet.getLastColumn(), 1);
  var header = sheet.getRange(1, 1, 1, lc).getValues()[0];
  var col = {};
  header.forEach(function (h, i) { var k = _norm(h); if (k) col[k] = i; });
  var rows = [];
  var lr = sheet.getLastRow();
  if (lr >= 2) {
    var vals = sheet.getRange(2, 1, lr - 1, lc).getValues();
    for (var i = 0; i < vals.length; i++) rows.push({ row: i + 2, vals: vals[i] });
  }
  return { col: col, rows: rows, ncols: lc };
}

function _boolCelula(v) {
  return v === true || _norm(v) === 'TRUE' || _norm(v) === 'VERDADEIRO';
}
function _normFreq(v) { return _norm(v) === 'ANUAL' ? 'anual' : 'mensal'; }
function _mesBaseValido(v) { var n = parseInt(v, 10); return (n >= 1 && n <= 12) ? n : null; }

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
  var frequencia = _normFreq(p.frequencia);
  var mesBase = frequencia === 'anual' ? (_mesBaseValido(p.mesBase) || '') : '';
  if (!descricao) throw new Error('Informe a descrição da despesa fixa.');

  var sh = _sheetFixas();
  var t = _lerTabela(sh);
  var alvo = _norm(descricao);
  var existente = t.rows.find(function (r) { return _norm(r.vals[t.col['DESCRICAO']]) === alvo; });

  if (existente) {
    var r = existente.row;
    _setCel(sh, r, t.col['CATEGORIA'], categoria);
    _setCel(sh, r, t.col['SUBCATEGORIA'], subcategoria);
    if (valorBase > 0) _setCel(sh, r, t.col['VALOR_BASE'], valorBase);
    _setCel(sh, r, t.col['DIA'], dia);
    _setCel(sh, r, t.col['ATIVO'], ativo);
    _setCel(sh, r, t.col['FREQUENCIA'], frequencia);
    _setCel(sh, r, t.col['MES_BASE'], mesBase);
    return String(existente.vals[t.col['ID']]);
  }
  var id = 'F' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  var linha = new Array(t.ncols).fill('');
  linha[t.col['ID']] = id;
  linha[t.col['DESCRICAO']] = descricao;
  linha[t.col['CATEGORIA']] = categoria;
  linha[t.col['SUBCATEGORIA']] = subcategoria;
  linha[t.col['VALOR_BASE']] = valorBase;
  linha[t.col['DIA']] = dia;
  linha[t.col['ATIVO']] = ativo;
  linha[t.col['FREQUENCIA']] = frequencia;
  linha[t.col['MES_BASE']] = mesBase;
  sh.appendRow(linha);
  return id;
}
function _setCel(sh, row, col0, valor) { if (col0 != null) sh.getRange(row, col0 + 1).setValue(valor); }

/** Adiciona/edita modelo pela tela de Fixas. */
function addFixaManual(payload) { return { ok: true, id: upsertFixa(payload) }; }

function updateFixa(id, payload) {
  var sh = _sheetFixas();
  var t = _lerTabela(sh);
  var alvo = t.rows.find(function (r) { return String(r.vals[t.col['ID']]) === String(id); });
  if (!alvo) throw new Error('Despesa fixa não encontrada.');
  var r = alvo.row;
  if (payload.descricao != null) _setCel(sh, r, t.col['DESCRICAO'], String(payload.descricao).trim());
  if (payload.categoria != null) _setCel(sh, r, t.col['CATEGORIA'], String(payload.categoria).trim());
  if (payload.subcategoria != null) _setCel(sh, r, t.col['SUBCATEGORIA'], String(payload.subcategoria).trim());
  if (payload.valorBase != null) _setCel(sh, r, t.col['VALOR_BASE'], _round2(Number(String(payload.valorBase).replace(',', '.')) || 0));
  if (payload.dia != null) _setCel(sh, r, t.col['DIA'], Math.min(28, Math.max(1, parseInt(payload.dia, 10) || 1)));
  if (payload.ativo != null) _setCel(sh, r, t.col['ATIVO'], payload.ativo === true || payload.ativo === 'true');
  if (payload.frequencia != null) {
    var freq = _normFreq(payload.frequencia);
    _setCel(sh, r, t.col['FREQUENCIA'], freq);
    _setCel(sh, r, t.col['MES_BASE'], freq === 'anual' ? (_mesBaseValido(payload.mesBase) || '') : '');
  } else if (payload.mesBase != null) {
    _setCel(sh, r, t.col['MES_BASE'], _mesBaseValido(payload.mesBase) || '');
  }
  return { ok: true };
}

function toggleFixaAtivo(id, ativo) { return updateFixa(id, { ativo: !!ativo }); }

/** Exclui um modelo e suas validações. */
function deleteFixa(id) {
  var sh = _sheetFixas();
  var t = _lerTabela(sh);
  var alvo = t.rows.find(function (r) { return String(r.vals[t.col['ID']]) === String(id); });
  if (alvo) sh.deleteRow(alvo.row);
  var shp = _sheetProjecao();
  var tp = _lerTabela(shp);
  var linhas = tp.rows.filter(function (r) { return String(r.vals[tp.col['FIXA_ID']]) === String(id); }).map(function (r) { return r.row; });
  for (var i = linhas.length - 1; i >= 0; i--) shp.deleteRow(linhas[i]);
  return { ok: true };
}

// ------------------------------------------------------------------
// Validação por mês (overrides)
// ------------------------------------------------------------------
function _acharProjRow(tp, fixaId, mes) {
  return tp.rows.find(function (r) {
    return String(r.vals[tp.col['FIXA_ID']]) === String(fixaId) && String(r.vals[tp.col['MES']]) === String(mes);
  });
}

function validarFixa(fixaId, mes, valor) {
  if (!fixaId || !mes) throw new Error('Dados incompletos.');
  var v = _round2(Number(String(valor).replace(',', '.')) || 0);
  var shp = _sheetProjecao();
  var tp = _lerTabela(shp);
  var existente = _acharProjRow(tp, fixaId, mes);
  var agora = new Date();
  if (existente) {
    _setCel(shp, existente.row, tp.col['VALOR'], v);
    _setCel(shp, existente.row, tp.col['STATUS'], 'validado');
    _setCel(shp, existente.row, tp.col['ATUALIZADO_EM'], agora);
  } else {
    var linha = new Array(tp.ncols).fill('');
    linha[tp.col['FIXA_ID']] = fixaId; linha[tp.col['MES']] = mes;
    linha[tp.col['VALOR']] = v; linha[tp.col['STATUS']] = 'validado'; linha[tp.col['ATUALIZADO_EM']] = agora;
    shp.appendRow(linha);
  }
  return { ok: true, valor: v };
}

function desvalidarFixa(fixaId, mes) {
  var shp = _sheetProjecao();
  var tp = _lerTabela(shp);
  var existente = _acharProjRow(tp, fixaId, mes);
  if (existente) shp.deleteRow(existente.row);
  return { ok: true };
}

// ------------------------------------------------------------------
// Painel consolidado (12 meses)
// ------------------------------------------------------------------
function getControleFixas() {
  var shf = _sheetFixas();
  var tf = _lerTabela(shf);
  var templates = tf.rows.map(function (r) {
    var freq = _normFreq(r.vals[tf.col['FREQUENCIA']]);
    var mesBase = _mesBaseValido(r.vals[tf.col['MES_BASE']]);
    return {
      id: String(r.vals[tf.col['ID']]),
      descricao: String(r.vals[tf.col['DESCRICAO']] || ''),
      categoria: String(r.vals[tf.col['CATEGORIA']] || ''),
      subcategoria: String(r.vals[tf.col['SUBCATEGORIA']] || ''),
      valorBase: _round2(r.vals[tf.col['VALOR_BASE']] || 0),
      dia: parseInt(r.vals[tf.col['DIA']], 10) || 1,
      ativo: _boolCelula(r.vals[tf.col['ATIVO']]),
      frequencia: freq,
      mesBase: mesBase,
      precisaConfigurar: freq === 'anual' && !mesBase
    };
  });

  var shp = _sheetProjecao();
  var tp = _lerTabela(shp);
  var overrides = {};
  tp.rows.forEach(function (r) {
    var k = String(r.vals[tp.col['FIXA_ID']]) + '|' + String(r.vals[tp.col['MES']]);
    overrides[k] = { valor: _round2(r.vals[tp.col['VALOR']] || 0), status: String(r.vals[tp.col['STATUS']] || 'validado') };
  });

  var parcelasPorMes = _parcelasLongasPorMes(_lerDados());

  var hoje = new Date();
  var meses = [];
  for (var i = 0; i < 12; i++) {
    var d = new Date(hoje.getFullYear(), hoje.getMonth() + i, 1);
    var mes = _chaveMes(d);
    var mesNum = d.getMonth() + 1;
    var fixasMes = [];
    var totalFixas = 0;
    templates.filter(function (t) { return t.ativo; }).forEach(function (t) {
      // filtro de frequência
      if (t.frequencia === 'anual') {
        if (!t.mesBase || t.mesBase !== mesNum) return;
      }
      var ov = overrides[t.id + '|' + mes];
      var previsto = t.valorBase;
      var valor = ov ? ov.valor : previsto;
      totalFixas += valor;
      fixasMes.push({
        fixaId: t.id, descricao: t.descricao, categoria: t.categoria, subcategoria: t.subcategoria,
        previsto: previsto, validado: ov ? ov.valor : null, status: ov ? 'validado' : 'previsto',
        valor: _round2(valor), dia: t.dia, frequencia: t.frequencia
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
    if (l.valor >= 0) return;
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

function _rotuloMesLongo(d) {
  var s = Utilities.formatDate(d, TZ, 'MMM/yy');
  return s.charAt(0).toUpperCase() + s.slice(1);
}

// ------------------------------------------------------------------
// Carga inicial (rode UMA vez no editor)
// ------------------------------------------------------------------
/**
 * Cadastra as despesas fixas atuais. Idempotente (upsert por descrição).
 * As duas anuais ficam com o mês de cobrança EM BRANCO — defina o mês de cada
 * uma na aba "Fixas" do app (ou me diga os meses que eu preencho aqui).
 */
function seedFixasIniciais() {
  var lista = [
    { dia: 8,  valorBase: 69.90,   descricao: 'App Grit',               subcategoria: 'Apps',                   categoria: 'Assinaturas',   frequencia: 'mensal' },
    { dia: 15, valorBase: 1542.86, descricao: 'Seguro do carro',        subcategoria: 'Seguro veículo',         categoria: 'Transporte',    frequencia: 'anual' },
    { dia: 16, valorBase: 118.58,  descricao: 'Claude',                 subcategoria: 'IA',                     categoria: 'Assinaturas',   frequencia: 'mensal' },
    { dia: 17, valorBase: 20.90,   descricao: 'Assinatura Netflix',     subcategoria: 'Streaming',              categoria: 'Lazer',         frequencia: 'mensal' },
    { dia: 18, valorBase: 11.48,   descricao: 'Seguro Residencial',     subcategoria: 'Seguro residencial',     categoria: 'Moradia',       frequencia: 'mensal' },
    { dia: 19, valorBase: 45.00,   descricao: 'Mensalidade de Celular', subcategoria: 'Telefonia',              categoria: 'Assinaturas',   frequencia: 'mensal' },
    { dia: 21, valorBase: 149.90,  descricao: 'Mensalidade Academia',   subcategoria: 'Academia',               categoria: 'Saúde',         frequencia: 'mensal' },
    { dia: 25, valorBase: 19.90,   descricao: 'Icloud',                 subcategoria: 'Cloud',                  categoria: 'Assinaturas',   frequencia: 'mensal' },
    { dia: 26, valorBase: 19.90,   descricao: 'Smart Nutri',            subcategoria: 'Exames',                 categoria: 'Saúde',         frequencia: 'mensal' },
    { dia: 27, valorBase: 149.90,  descricao: 'Strava',                 subcategoria: 'Esportes',               categoria: 'Saúde',         frequencia: 'mensal' },
    { dia: 30, valorBase: 169.90,  descricao: 'Barbearia Paradise',     subcategoria: 'Outros gastos de lazer', categoria: 'Lazer',         frequencia: 'mensal' },
    { dia: 30, valorBase: 340.00,  descricao: 'Psicologo',              subcategoria: 'Terapia',                categoria: 'Saúde',         frequencia: 'mensal' },
    { dia: 6,  valorBase: 2420.12, descricao: 'Plano Prudential',       subcategoria: 'Seguro de vida',         categoria: 'Investimentos', frequencia: 'anual' },
    { dia: 8,  valorBase: 2435.13, descricao: 'Aluguel do apartamento', subcategoria: 'Aluguel',                categoria: 'Moradia',       frequencia: 'mensal' }
  ];
  var n = 0;
  lista.forEach(function (x) {
    upsertFixa({ descricao: x.descricao, categoria: x.categoria, subcategoria: x.subcategoria,
      valorBase: x.valorBase, dia: x.dia, ativo: true, frequencia: x.frequencia });
    n++;
  });
  var msg = 'Cadastradas/atualizadas ' + n + ' despesas fixas (2 anuais sem mês definido: Seguro do carro e Plano Prudential).';
  Logger.log(msg);
  return msg;
}
