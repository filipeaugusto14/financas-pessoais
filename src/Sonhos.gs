/**
 * Sonhos.gs — Metas financeiras de longo prazo (comprar casa, trocar de carro,
 * viagem dos sonhos, etc.). Planejamento: acompanha quanto já foi guardado,
 * quanto falta, quanto guardar por mês para chegar na data alvo e a previsão
 * de conclusão com o aporte planejado.
 *
 * Abas (criadas pelo setup):
 *  - SONHOS: ID | NOME | ICONE | VALOR_ALVO | VALOR_GUARDADO | DATA_ALVO | APORTE_MENSAL | STATUS | CRIADO_EM
 *  - SONHOS_APORTES (histórico): SONHO_ID | DATA | VALOR | OBS
 */

var HEADER_SONHOS = ['ID', 'NOME', 'ICONE', 'VALOR_ALVO', 'VALOR_GUARDADO', 'DATA_ALVO', 'APORTE_MENSAL', 'STATUS', 'CRIADO_EM'];
var HEADER_APORTES = ['SONHO_ID', 'DATA', 'VALOR', 'OBS'];

function _sheetSonhos() { return _ensureSheet(SHEET_SONHOS, HEADER_SONHOS); }
function _sheetAportes() { return _ensureSheet(SHEET_SONHOS_APORTES, HEADER_APORTES); }

// ------------------------------------------------------------------
// CRUD
// ------------------------------------------------------------------

function listSonhos() {
  var sh = _sheetSonhos();
  var t = _lerTabela(sh);
  var hoje = new Date();
  var sonhos = t.rows.map(function (r) {
    return _montarSonho(r.vals, t.col, hoje);
  }).filter(function (s) { return s.nome; });

  // ativos primeiro, depois por progresso
  sonhos.sort(function (a, b) {
    if (a.concluido !== b.concluido) return a.concluido ? 1 : -1;
    return b.pct - a.pct;
  });

  var resumo = { totalAlvo: 0, totalGuardado: 0, totalRestante: 0, aporteNecessarioTotal: 0, quantidade: sonhos.length };
  sonhos.forEach(function (s) {
    resumo.totalAlvo += s.valorAlvo;
    resumo.totalGuardado += s.valorGuardado;
    resumo.totalRestante += s.restante;
    if (s.aporteNecessario) resumo.aporteNecessarioTotal += s.aporteNecessario;
  });
  ['totalAlvo', 'totalGuardado', 'totalRestante', 'aporteNecessarioTotal'].forEach(function (k) { resumo[k] = _round2(resumo[k]); });
  resumo.pct = resumo.totalAlvo > 0 ? _round2(Math.min(100, resumo.totalGuardado / resumo.totalAlvo * 100)) : 0;

  return { sonhos: sonhos, resumo: resumo };
}

function _montarSonho(vals, col, hoje) {
  var alvo = _round2(vals[col['VALOR_ALVO']] || 0);
  var guardado = _round2(vals[col['VALOR_GUARDADO']] || 0);
  var restante = _round2(Math.max(0, alvo - guardado));
  var dataAlvo = _paraData(vals[col['DATA_ALVO']]);
  var aporteMensal = _round2(vals[col['APORTE_MENSAL']] || 0);
  var concluido = alvo > 0 && guardado >= alvo;

  var mesesRestantes = null, aporteNecessario = null;
  if (dataAlvo) {
    mesesRestantes = Math.max(0, (dataAlvo.getFullYear() - hoje.getFullYear()) * 12 + (dataAlvo.getMonth() - hoje.getMonth()));
    aporteNecessario = restante > 0 ? _round2(restante / Math.max(1, mesesRestantes)) : 0;
  }

  var previsao = null;
  if (!concluido && aporteMensal > 0 && restante > 0) {
    var n = Math.ceil(restante / aporteMensal);
    var d = new Date(hoje.getFullYear(), hoje.getMonth() + n, 1);
    previsao = Utilities.formatDate(d, TZ, 'MM/yyyy');
  }

  return {
    id: String(vals[col['ID']]),
    nome: String(vals[col['NOME']] || ''),
    icone: String(vals[col['ICONE']] || '🎯'),
    valorAlvo: alvo,
    valorGuardado: guardado,
    restante: restante,
    pct: alvo > 0 ? _round2(Math.min(100, guardado / alvo * 100)) : 0,
    dataAlvo: dataAlvo ? Utilities.formatDate(dataAlvo, TZ, 'yyyy-MM-dd') : '',
    dataAlvoBr: dataAlvo ? Utilities.formatDate(dataAlvo, TZ, 'MM/yyyy') : '',
    aporteMensal: aporteMensal,
    mesesRestantes: mesesRestantes,
    aporteNecessario: aporteNecessario,
    previsaoConclusao: previsao,
    concluido: concluido,
    status: concluido ? 'concluido' : (String(vals[col['STATUS']] || 'ativo') || 'ativo')
  };
}

function addSonho(payload) {
  var nome = String(payload.nome || '').trim();
  if (!nome) throw new Error('Dê um nome ao sonho.');
  var sh = _sheetSonhos();
  var t = _lerTabela(sh);
  var id = 'S' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  var linha = new Array(t.ncols).fill('');
  linha[t.col['ID']] = id;
  linha[t.col['NOME']] = nome;
  linha[t.col['ICONE']] = String(payload.icone || '🎯');
  linha[t.col['VALOR_ALVO']] = _round2(_num(payload.valorAlvo));
  linha[t.col['VALOR_GUARDADO']] = _round2(_num(payload.valorGuardado));
  linha[t.col['DATA_ALVO']] = payload.dataAlvo ? Utilities.formatDate(_paraData(payload.dataAlvo), TZ, 'yyyy-MM-dd') : '';
  linha[t.col['APORTE_MENSAL']] = _round2(_num(payload.aporteMensal));
  linha[t.col['STATUS']] = 'ativo';
  linha[t.col['CRIADO_EM']] = new Date();
  sh.appendRow(linha);
  return { ok: true, id: id };
}

function updateSonho(id, payload) {
  var sh = _sheetSonhos();
  var t = _lerTabela(sh);
  var alvo = t.rows.find(function (r) { return String(r.vals[t.col['ID']]) === String(id); });
  if (!alvo) throw new Error('Sonho não encontrado.');
  var r = alvo.row;
  if (payload.nome != null) sh.getRange(r, t.col['NOME'] + 1).setValue(String(payload.nome).trim());
  if (payload.icone != null) sh.getRange(r, t.col['ICONE'] + 1).setValue(String(payload.icone));
  if (payload.valorAlvo != null) sh.getRange(r, t.col['VALOR_ALVO'] + 1).setValue(_round2(_num(payload.valorAlvo)));
  if (payload.valorGuardado != null) sh.getRange(r, t.col['VALOR_GUARDADO'] + 1).setValue(_round2(_num(payload.valorGuardado)));
  if (payload.dataAlvo != null) sh.getRange(r, t.col['DATA_ALVO'] + 1).setValue(payload.dataAlvo ? Utilities.formatDate(_paraData(payload.dataAlvo), TZ, 'yyyy-MM-dd') : '');
  if (payload.aporteMensal != null) sh.getRange(r, t.col['APORTE_MENSAL'] + 1).setValue(_round2(_num(payload.aporteMensal)));
  if (payload.status != null) sh.getRange(r, t.col['STATUS'] + 1).setValue(String(payload.status));
  return { ok: true };
}

function deleteSonho(id) {
  var sh = _sheetSonhos();
  var t = _lerTabela(sh);
  var alvo = t.rows.find(function (r) { return String(r.vals[t.col['ID']]) === String(id); });
  if (alvo) sh.deleteRow(alvo.row);
  // remove aportes do sonho
  var sha = _sheetAportes();
  var ta = _lerTabela(sha);
  var linhas = ta.rows.filter(function (r) { return String(r.vals[ta.col['SONHO_ID']]) === String(id); }).map(function (r) { return r.row; });
  for (var i = linhas.length - 1; i >= 0; i--) sha.deleteRow(linhas[i]);
  return { ok: true };
}

// ------------------------------------------------------------------
// Aportes
// ------------------------------------------------------------------

/** Registra um aporte (guardei R$ X) e soma ao valor guardado do sonho. */
function registrarAporte(sonhoId, valor, data, obs) {
  var v = _round2(_num(valor));
  if (!(v !== 0)) throw new Error('Informe um valor.');
  var sh = _sheetSonhos();
  var t = _lerTabela(sh);
  var alvo = t.rows.find(function (r) { return String(r.vals[t.col['ID']]) === String(sonhoId); });
  if (!alvo) throw new Error('Sonho não encontrado.');

  var atual = _round2(alvo.vals[t.col['VALOR_GUARDADO']] || 0);
  var novo = _round2(atual + v);
  sh.getRange(alvo.row, t.col['VALOR_GUARDADO'] + 1).setValue(novo);
  // conclui automaticamente se atingiu a meta
  var metaAlvo = _round2(alvo.vals[t.col['VALOR_ALVO']] || 0);
  if (metaAlvo > 0 && novo >= metaAlvo) sh.getRange(alvo.row, t.col['STATUS'] + 1).setValue('concluido');

  var sha = _sheetAportes();
  var ta = _lerTabela(sha);
  var linha = new Array(ta.ncols).fill('');
  linha[ta.col['SONHO_ID']] = sonhoId;
  linha[ta.col['DATA']] = data ? _paraData(data) : new Date();
  linha[ta.col['VALOR']] = v;
  linha[ta.col['OBS']] = String(obs || '');
  sha.appendRow(linha);
  return { ok: true, valorGuardado: novo };
}

/** Histórico de aportes de um sonho (mais recentes primeiro). */
function listAportes(sonhoId) {
  var sha = _sheetAportes();
  var ta = _lerTabela(sha);
  var out = ta.rows.filter(function (r) { return String(r.vals[ta.col['SONHO_ID']]) === String(sonhoId); })
    .map(function (r) {
      var d = _paraData(r.vals[ta.col['DATA']]);
      return {
        row: r.row,
        data: d ? Utilities.formatDate(d, TZ, 'dd/MM/yy') : '',
        valor: _round2(r.vals[ta.col['VALOR']] || 0),
        obs: String(r.vals[ta.col['OBS']] || '')
      };
    });
  out.reverse();
  return out;
}

function _num(v) { return Number(String(v == null ? 0 : v).replace(',', '.')) || 0; }
