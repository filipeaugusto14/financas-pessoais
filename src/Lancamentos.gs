/**
 * Lancamentos.gs — CRUD dos lançamentos (aba de entrada) mantendo a aba DADOS
 * sincronizada. Todas as funções aqui são chamadas pelo cliente via
 * google.script.run.
 */

/**
 * Adiciona um lançamento: grava na aba de entrada e explode na aba DADOS.
 * @param {Object} payload {data:'yyyy-MM-dd', valor:Number, descricao, subcategoria, categoria, tipo, parcelas}
 * @return {Object} { ok, id, linhasDados }
 */
function addLancamento(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var dados = _validarPayload(payload);
    var abas = _detectarAbas();
    var idCol = _garantirColunaId(abas.entrada);
    _garantirColunaId(abas.dados);

    var id = _novoId();
    dados.id = id;

    // 1) grava na aba de entrada
    var E = _colsEntrada(abas.entrada);
    var maxCol = Math.max(E.data, E.valor, E.desc, E.subcat, E.cat, E.tipo, E.parcelas, E.enviado, idCol);
    var row = new Array(maxCol).fill('');
    row[E.data - 1] = dados.data;
    row[E.valor - 1] = dados.valor;               // total positivo
    row[E.desc - 1] = dados.descricao;
    row[E.subcat - 1] = dados.subcategoria;
    row[E.cat - 1] = dados.categoria;
    row[E.tipo - 1] = dados.tipo;
    row[E.parcelas - 1] = dados.parcelas;
    row[E.enviado - 1] = true;
    row[idCol - 1] = id;
    var inicio = abas.entrada.getLastRow() + 1;
    abas.entrada.getRange(inicio, 1, 1, maxCol).setValues([row]);

    // 2) explode na aba DADOS
    var linhas = _gerarLinhasDados(dados);
    _appendLinhasDados(abas.dados, linhas);

    // 3) se marcado como despesa fixa (e não parcelado), cria/atualiza o modelo
    var fixaId = null;
    if (payload.fixa && dados.parcelas <= 1 && _norm(dados.tipo) === 'DESPESA') {
      try {
        fixaId = upsertFixa({
          descricao: dados.descricao || dados.subcategoria,
          categoria: dados.categoria, subcategoria: dados.subcategoria,
          valorBase: dados.valor, dia: dados.data.getDate(), ativo: true
        });
      } catch (e) { /* não impede o lançamento */ }
    }

    return { ok: true, id: id, linhasDados: linhas.length, fixaId: fixaId };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Lista os lançamentos mais recentes (para a lista rápida e edição).
 * @param {number} limite quantidade (padrão 20)
 * @return {Array<Object>}
 */
function listLancamentos(limite) {
  limite = limite || 20;
  var abas = _detectarAbas();
  var sh = abas.entrada;
  var E = _colsEntrada(sh);
  var ult = sh.getLastRow();
  if (ult < E.map.firstDataRow) return [];
  var n = ult - E.map.firstDataRow + 1;
  var maxCol = Math.max(E.data, E.valor, E.desc, E.subcat, E.cat, E.tipo, E.parcelas, E.id || 0);
  var vals = sh.getRange(E.map.firstDataRow, 1, n, maxCol).getValues();

  var out = [];
  for (var i = vals.length - 1; i >= 0 && out.length < limite; i--) {
    var r = vals[i];
    var data = _paraData(r[E.data - 1]);
    if (!data) continue;
    out.push({
      id: E.id ? String(r[E.id - 1] || '') : '',
      rowIndex: E.map.firstDataRow + i,
      data: Utilities.formatDate(data, TZ, 'yyyy-MM-dd'),
      dataBr: Utilities.formatDate(data, TZ, 'dd/MM/yy'),
      valor: _round2(r[E.valor - 1]),
      descricao: String(r[E.desc - 1] || ''),
      subcategoria: String(r[E.subcat - 1] || ''),
      categoria: String(r[E.cat - 1] || ''),
      tipo: String(r[E.tipo - 1] || ''),
      parcelas: parseInt(r[E.parcelas - 1], 10) || 1
    });
  }
  return out;
}

/**
 * Edita um lançamento existente (identificado por ID). Atualiza a linha de
 * entrada no lugar e regenera as linhas na aba DADOS.
 */
function editLancamento(id, payload) {
  if (!id) throw new Error('Lançamento sem ID não pode ser editado pelo app.');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var dados = _validarPayload(payload);
    dados.id = id;
    var abas = _detectarAbas();
    var E = _colsEntrada(abas.entrada);
    if (!E.id) throw new Error('A aba de entrada não tem coluna ID. Rode setup().');

    var linha = _acharLinhaPorId(abas.entrada, E.id, E.map.firstDataRow, id);
    if (linha < 0) throw new Error('Lançamento não encontrado (ID ' + id + ').');

    var maxCol = Math.max(E.data, E.valor, E.desc, E.subcat, E.cat, E.tipo, E.parcelas, E.enviado, E.id);
    var row = abas.entrada.getRange(linha, 1, 1, maxCol).getValues()[0];
    row[E.data - 1] = dados.data;
    row[E.valor - 1] = dados.valor;
    row[E.desc - 1] = dados.descricao;
    row[E.subcat - 1] = dados.subcategoria;
    row[E.cat - 1] = dados.categoria;
    row[E.tipo - 1] = dados.tipo;
    row[E.parcelas - 1] = dados.parcelas;
    row[E.enviado - 1] = true;
    row[E.id - 1] = id;
    abas.entrada.getRange(linha, 1, 1, maxCol).setValues([row]);

    // regenera DADOS
    _removerLinhasDadosPorId(abas.dados, id);
    _appendLinhasDados(abas.dados, _gerarLinhasDados(dados));
    return { ok: true, id: id };
  } finally {
    lock.releaseLock();
  }
}

/** Exclui um lançamento (linha de entrada + linhas correspondentes em DADOS). */
function deleteLancamento(id) {
  if (!id) throw new Error('Lançamento sem ID não pode ser excluído pelo app.');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var abas = _detectarAbas();
    var E = _colsEntrada(abas.entrada);
    if (!E.id) throw new Error('A aba de entrada não tem coluna ID. Rode setup().');
    var linha = _acharLinhaPorId(abas.entrada, E.id, E.map.firstDataRow, id);
    var removidasDados = _removerLinhasDadosPorId(abas.dados, id);
    if (linha > 0) abas.entrada.deleteRow(linha);
    return { ok: true, entradaRemovida: linha > 0, dadosRemovidos: removidasDados };
  } finally {
    lock.releaseLock();
  }
}

// ------------------------------------------------------------------
// Helpers
// ------------------------------------------------------------------

function _validarPayload(p) {
  if (!p) throw new Error('Dados vazios.');
  var data = _paraData(p.data);
  if (!data) throw new Error('Data inválida.');
  var valor = _round2(Number(String(p.valor).replace(',', '.')));
  if (!(valor > 0)) throw new Error('Informe um valor maior que zero.');
  var tipo = _norm(p.tipo) === 'RECEITA' ? 'Receita' : 'Despesa';
  var categoria = String(p.categoria || '').trim();
  var subcategoria = String(p.subcategoria || '').trim();
  if (!categoria) throw new Error('Escolha a categoria.');
  if (!subcategoria) throw new Error('Escolha a subcategoria.');
  var parcelas = Math.max(1, parseInt(p.parcelas, 10) || 1);
  return {
    data: data,
    valor: valor,
    descricao: String(p.descricao || '').trim(),
    subcategoria: subcategoria,
    categoria: categoria,
    tipo: tipo,
    parcelas: parcelas
  };
}

function _acharLinhaPorId(sheet, idCol, firstDataRow, id) {
  var ult = sheet.getLastRow();
  if (ult < firstDataRow) return -1;
  var n = ult - firstDataRow + 1;
  var ids = sheet.getRange(firstDataRow, idCol, n, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return firstDataRow + i;
  }
  return -1;
}
