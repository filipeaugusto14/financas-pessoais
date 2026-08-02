/**
 * Relatorio.gs — Agregações do dashboard a partir da aba DADOS (valores com
 * sinal, parcelas já explodidas). Uma única leitura da planilha por chamada.
 */

/** Lê a aba DADOS uma vez e devolve linhas normalizadas. */
function _lerDados() {
  var abas = _detectarAbas();
  var sh = abas.dados;
  var C = _colsDados(sh);
  var ult = sh.getLastRow();
  var linhas = [];
  if (ult < C.map.firstDataRow) return linhas;
  var n = ult - C.map.firstDataRow + 1;
  var maxCol = Math.max(C.data, C.valor, C.desc, C.subcat, C.cat, C.tipo);
  var vals = sh.getRange(C.map.firstDataRow, 1, n, maxCol).getValues();
  for (var i = 0; i < vals.length; i++) {
    var data = _paraData(vals[i][C.data - 1]);
    if (!data) continue;
    var valor = Number(vals[i][C.valor - 1]);
    if (isNaN(valor)) continue;
    var tipo = String(vals[i][C.tipo - 1] || '').trim();
    if (!tipo) tipo = valor >= 0 ? 'Receita' : 'Despesa';
    linhas.push({
      data: data,
      ts: data.getTime(),
      mes: _chaveMes(data),
      valor: valor,
      descricao: String(vals[i][C.desc - 1] || ''),
      subcategoria: String(vals[i][C.subcat - 1] || ''),
      categoria: String(vals[i][C.cat - 1] || ''),
      tipo: tipo,
      receita: valor > 0,
      abs: Math.abs(valor)
    });
  }
  return linhas;
}

/**
 * Relatório completo para um período.
 * @param {string} inicio 'yyyy-MM-dd' (opcional; padrão = 1º dia do mês atual)
 * @param {string} fim 'yyyy-MM-dd' (opcional; padrão = último dia do mês atual)
 * @return {Object}
 */
function getRelatorio(inicio, fim) {
  var linhas = _lerDados();
  var per = _resolverPeriodo(inicio, fim);
  var prev = _periodoAnterior(per);

  var noPeriodo = linhas.filter(function (l) { return l.ts >= per.ini.getTime() && l.ts <= per.fim.getTime(); });
  var noAnterior = linhas.filter(function (l) { return l.ts >= prev.ini.getTime() && l.ts <= prev.fim.getTime(); });

  var totais = _totais(noPeriodo);
  var totaisPrev = _totais(noAnterior);

  // Acumulado do ano (1º jan do ano do fim até o fim do período)
  var anoIni = new Date(per.fim.getFullYear(), 0, 1);
  var acumuladoAno = _somaResultado(linhas.filter(function (l) {
    return l.ts >= anoIni.getTime() && l.ts <= per.fim.getTime();
  }));
  var mesesDecorridos = per.fim.getMonth() + 1;
  var projecaoAnual = mesesDecorridos > 0 ? acumuladoAno / mesesDecorridos * 12 : 0;

  // Série mensal (12 meses terminando no mês do fim)
  var meses = _ultimosMeses(per.fim, 12);
  var serieMensal = _serieMensal(linhas, meses);

  // Patrimônio/saldo acumulado (soma acumulada de resultado até cada mês)
  var patrimonio = _patrimonioAcumulado(linhas, meses);

  // Categorias — despesas e receitas do período (com variação vs anterior)
  var despCatPeriodo = _porCategoria(noPeriodo, 'Despesa');
  var despCatPrev = _porCategoria(noAnterior, 'Despesa');
  var recCatPeriodo = _porCategoria(noPeriodo, 'Receita');
  var recCatPrev = _porCategoria(noAnterior, 'Receita');

  // Subcategorias do período
  var despSub = _porSubcategoria(noPeriodo, 'Despesa');
  var recSub = _porSubcategoria(noPeriodo, 'Receita');

  // Composição mensal por categoria (top 8 + Outros)
  var compDesp = _composicaoMensal(linhas, meses, 'Despesa', 8);
  var compRec = _composicaoMensal(linhas, meses, 'Receita', 8);

  // Indicadores
  var taxaPoupancaMes = totais.receita > 0 ? (totais.resultado / totais.receita) * 100 : null;
  var meses12 = _ultimosMeses(per.fim, 12);
  var doze = linhas.filter(function (l) { return meses12.indexOf(l.mes) !== -1; });
  var tot12 = _totais(doze);
  var taxaPoupanca12 = tot12.receita > 0 ? (tot12.resultado / tot12.receita) * 100 : null;
  var custoVida3m = _mediaMovelDespesa(linhas, per.fim, 3);

  // Extras — contas futuras (despesas com data > hoje) por mês
  var contasFuturas = _contasFuturas(linhas, 6);

  // Lançamentos do período (ordenados por valor absoluto)
  var lancPeriodo = noPeriodo.slice().sort(function (a, b) { return b.abs - a.abs; }).map(function (l) {
    return {
      data: Utilities.formatDate(l.data, TZ, 'dd/MM/yy'),
      descricao: l.descricao, categoria: l.categoria, subcategoria: l.subcategoria,
      tipo: l.tipo, valor: l.valor, abs: l.abs
    };
  });

  return {
    periodo: {
      inicio: Utilities.formatDate(per.ini, TZ, 'yyyy-MM-dd'),
      fim: Utilities.formatDate(per.fim, TZ, 'yyyy-MM-dd'),
      label: _labelPeriodo(per)
    },
    kpis: {
      receita: totais.receita,
      despesa: totais.despesa,
      resultado: totais.resultado,
      acumuladoAno: acumuladoAno,
      variacao: {
        receita: _variacao(totais.receita, totaisPrev.receita),
        despesa: _variacao(totais.despesa, totaisPrev.despesa),
        resultado: _variacao(totais.resultado, totaisPrev.resultado)
      }
    },
    serieMensal: serieMensal,
    patrimonio: patrimonio,
    despesas: { categorias: _rankingCategorias(despCatPeriodo, despCatPrev, totais.despesa), subcategorias: despSub, composicao: compDesp },
    receitas: { categorias: _rankingCategorias(recCatPeriodo, recCatPrev, totais.receita), subcategorias: recSub, composicao: compRec },
    indicadores: {
      taxaPoupancaMes: taxaPoupancaMes,
      taxaPoupanca12: taxaPoupanca12,
      acumuladoAno: acumuladoAno,
      projecaoAnual: projecaoAnual,
      custoVidaMedio: custoVida3m
    },
    contasFuturas: contasFuturas,
    lancamentos: lancPeriodo
  };
}

// ------------------------------------------------------------------
// Cálculos auxiliares
// ------------------------------------------------------------------

function _totais(linhas) {
  var receita = 0, despesa = 0;
  linhas.forEach(function (l) {
    if (l.valor > 0) receita += l.valor; else despesa += -l.valor;
  });
  return { receita: _round2(receita), despesa: _round2(despesa), resultado: _round2(receita - despesa) };
}

function _somaResultado(linhas) {
  var s = 0;
  linhas.forEach(function (l) { s += l.valor; });
  return _round2(s);
}

function _variacao(atual, anterior) {
  if (!anterior) return null;
  return _round2(((atual - anterior) / Math.abs(anterior)) * 100);
}

function _resolverPeriodo(inicio, fim) {
  var hoje = new Date();
  var ini = inicio ? _paraData(inicio) : new Date(hoje.getFullYear(), hoje.getMonth(), 1);
  var f = fim ? _paraData(fim) : new Date(hoje.getFullYear(), hoje.getMonth() + 1, 0);
  ini = new Date(ini.getFullYear(), ini.getMonth(), ini.getDate(), 0, 0, 0);
  f = new Date(f.getFullYear(), f.getMonth(), f.getDate(), 23, 59, 59);
  return { ini: ini, fim: f };
}

/** Período imediatamente anterior. Se for um mês-calendário completo, usa o mês anterior. */
function _periodoAnterior(per) {
  var ini = per.ini, fim = per.fim;
  var ehMesCompleto = ini.getDate() === 1 &&
    fim.getDate() === new Date(fim.getFullYear(), fim.getMonth() + 1, 0).getDate() &&
    ini.getMonth() === fim.getMonth() && ini.getFullYear() === fim.getFullYear();
  if (ehMesCompleto) {
    var pIni = new Date(ini.getFullYear(), ini.getMonth() - 1, 1, 0, 0, 0);
    var pFim = new Date(ini.getFullYear(), ini.getMonth(), 0, 23, 59, 59);
    return { ini: pIni, fim: pFim };
  }
  var dur = fim.getTime() - ini.getTime();
  var pFim2 = new Date(ini.getTime() - 1);
  var pIni2 = new Date(pFim2.getTime() - dur);
  return { ini: pIni2, fim: pFim2 };
}

function _labelPeriodo(per) {
  var ini = per.ini, fim = per.fim;
  var mesmoMes = ini.getMonth() === fim.getMonth() && ini.getFullYear() === fim.getFullYear();
  var ehMesCompleto = mesmoMes && ini.getDate() === 1 &&
    fim.getDate() === new Date(fim.getFullYear(), fim.getMonth() + 1, 0).getDate();
  if (ehMesCompleto) {
    var s = Utilities.formatDate(ini, TZ, 'MMMM \'de\' yyyy');
    return s.charAt(0).toUpperCase() + s.slice(1);
  }
  return Utilities.formatDate(ini, TZ, 'dd/MM/yy') + ' – ' + Utilities.formatDate(fim, TZ, 'dd/MM/yy');
}

/** Array de chaves 'yyyy-MM' dos últimos n meses terminando no mês de refData. */
function _ultimosMeses(refData, n) {
  var out = [];
  var y = refData.getFullYear(), m = refData.getMonth();
  for (var i = n - 1; i >= 0; i--) {
    var d = new Date(y, m - i, 1);
    out.push(_chaveMes(d));
  }
  return out;
}

function _serieMensal(linhas, meses) {
  var idx = {}; meses.forEach(function (m, i) { idx[m] = i; });
  var receita = meses.map(function () { return 0; });
  var despesa = meses.map(function () { return 0; });
  linhas.forEach(function (l) {
    if (!(l.mes in idx)) return;
    if (l.valor > 0) receita[idx[l.mes]] += l.valor; else despesa[idx[l.mes]] += -l.valor;
  });
  return meses.map(function (m, i) {
    return {
      mes: m, label: _rotuloMes(m),
      receita: _round2(receita[i]), despesa: _round2(despesa[i]),
      resultado: _round2(receita[i] - despesa[i])
    };
  });
}

function _patrimonioAcumulado(linhas, meses) {
  // soma acumulada de todos os resultados até o fim de cada mês da janela
  var porMes = {};
  linhas.forEach(function (l) { porMes[l.mes] = (porMes[l.mes] || 0) + l.valor; });
  var chaves = Object.keys(porMes).sort();
  var acumulado = 0, mapaAcum = {};
  chaves.forEach(function (m) { acumulado += porMes[m]; mapaAcum[m] = acumulado; });
  // para cada mês da janela, pegar o acumulado até ele (ou o último anterior)
  var ordenadas = chaves;
  return meses.map(function (m) {
    // acumulado até o mês m
    var val = 0;
    for (var i = 0; i < ordenadas.length; i++) { if (ordenadas[i] <= m) val = mapaAcum[ordenadas[i]]; else break; }
    return { mes: m, label: _rotuloMes(m), valor: _round2(val) };
  });
}

function _porCategoria(linhas, tipo) {
  var alvoReceita = _norm(tipo) === 'RECEITA';
  var mapa = {};
  linhas.forEach(function (l) {
    var ehReceita = l.valor > 0;
    if (ehReceita !== alvoReceita) return;
    mapa[l.categoria] = (mapa[l.categoria] || 0) + l.abs;
  });
  return mapa; // {categoria: valor(+)}
}

function _rankingCategorias(mapaAtual, mapaPrev, total) {
  var arr = Object.keys(mapaAtual).map(function (cat) {
    var v = _round2(mapaAtual[cat]);
    return {
      categoria: cat, valor: v,
      pctTotal: total > 0 ? _round2((v / total) * 100) : 0,
      variacao: _variacao(v, _round2(mapaPrev[cat] || 0))
    };
  });
  arr.sort(function (a, b) { return b.valor - a.valor; });
  return arr;
}

function _porSubcategoria(linhas, tipo) {
  var alvoReceita = _norm(tipo) === 'RECEITA';
  var mapa = {}, total = 0;
  linhas.forEach(function (l) {
    var ehReceita = l.valor > 0;
    if (ehReceita !== alvoReceita) return;
    var chave = l.subcategoria || '(sem subcategoria)';
    if (!mapa[chave]) mapa[chave] = { subcategoria: chave, categoria: l.categoria, valor: 0 };
    mapa[chave].valor += l.abs; total += l.abs;
  });
  var arr = Object.keys(mapa).map(function (k) {
    mapa[k].valor = _round2(mapa[k].valor);
    mapa[k].pctTotal = total > 0 ? _round2((mapa[k].valor / total) * 100) : 0;
    return mapa[k];
  });
  arr.sort(function (a, b) { return b.valor - a.valor; });
  return arr;
}

/** Composição mensal por categoria: top N categorias do período + "Outros". */
function _composicaoMensal(linhas, meses, tipo, topN) {
  var alvoReceita = _norm(tipo) === 'RECEITA';
  var idxMes = {}; meses.forEach(function (m, i) { idxMes[m] = i; });
  var relevantes = linhas.filter(function (l) { return (l.valor > 0) === alvoReceita && (l.mes in idxMes); });

  // ranking de categorias na janela
  var totalCat = {};
  relevantes.forEach(function (l) { totalCat[l.categoria] = (totalCat[l.categoria] || 0) + l.abs; });
  var cats = Object.keys(totalCat).sort(function (a, b) { return totalCat[b] - totalCat[a]; });
  var top = cats.slice(0, topN);
  var usarOutros = cats.length > topN;

  var series = {}; top.forEach(function (c) { series[c] = meses.map(function () { return 0; }); });
  if (usarOutros) series['Outros'] = meses.map(function () { return 0; });

  relevantes.forEach(function (l) {
    var c = top.indexOf(l.categoria) !== -1 ? l.categoria : (usarOutros ? 'Outros' : null);
    if (!c) return;
    series[c][idxMes[l.mes]] += l.abs;
  });

  var categorias = top.concat(usarOutros ? ['Outros'] : []);
  return {
    meses: meses.map(_rotuloMes),
    categorias: categorias,
    series: categorias.map(function (c) { return { categoria: c, valores: series[c].map(_round2) }; })
  };
}

function _mediaMovelDespesa(linhas, refData, n) {
  var meses = _ultimosMeses(refData, n);
  var idx = {}; meses.forEach(function (m) { idx[m] = 0; });
  linhas.forEach(function (l) { if (l.mes in idx && l.valor < 0) idx[l.mes] += -l.valor; });
  var soma = 0; meses.forEach(function (m) { soma += idx[m]; });
  return _round2(soma / n);
}

function _contasFuturas(linhas, nMeses) {
  var hoje = new Date();
  var hojeTs = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate()).getTime();
  var meses = [];
  for (var i = 1; i <= nMeses; i++) meses.push(_chaveMes(new Date(hoje.getFullYear(), hoje.getMonth() + i, 1)));
  var idx = {}; meses.forEach(function (m) { idx[m] = 0; });
  linhas.forEach(function (l) {
    if (l.valor < 0 && l.ts > hojeTs && (l.mes in idx)) idx[l.mes] += -l.valor;
  });
  return meses.map(function (m) { return { mes: m, label: _rotuloMes(m), valor: _round2(idx[m]) }; });
}

/** 'yyyy-MM' -> 'MM/yy'. */
function _rotuloMes(chave) {
  var p = chave.split('-');
  return p[1] + '/' + p[0].slice(2);
}
