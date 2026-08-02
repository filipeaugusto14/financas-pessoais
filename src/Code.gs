/**
 * Code.gs — Ponto de entrada do web app.
 */

/** Serve a interface do app. */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(APP_NOME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Inclui um arquivo HTML dentro de outro (para separar CSS/JS). */
function include(nome) {
  return HtmlService.createHtmlOutputFromFile(nome).getContent();
}

/**
 * Dados iniciais entregues ao carregar o app (evita várias idas ao servidor).
 * @return {Object} { app, taxonomia, sugestoes, ultimos, hoje }
 */
function bootstrap() {
  return {
    app: APP_NOME,
    hoje: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'),
    taxonomia: getTaxonomia(),
    sugestoes: getSugestoesDescricao(),
    ultimos: listLancamentos(15)
  };
}
