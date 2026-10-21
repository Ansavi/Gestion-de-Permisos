/**
 * Codigo.gs
 * Punto de entrada del Web App. Sirve login.html o sistema.html según el
 * parámetro ?page=, igual que en la maqueta estática, pero ahora real.
 */

function doGet(e) {
  var pagina = (e.parameter.page || 'login');
  var plantilla = HtmlService.createTemplateFromFile(pagina === 'sistema' ? 'sistema' : 'login');
  return plantilla.evaluate()
    .setTitle('Sistema de Gestión de Accesos — MVCLM')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Permite usar <?!= include('archivo'); ?> dentro de los HTML si hace falta partirlos en piezas. */
function include(nombreArchivo) {
  return HtmlService.createHtmlOutputFromFile(nombreArchivo).getContent();
}
