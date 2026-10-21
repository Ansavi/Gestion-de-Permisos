/**
 * Codigo.gs
 * Punto de entrada del Web App. Sirve login.html o sistema.html según el
 * parámetro ?page=, igual que en la maqueta estática, pero ahora real.
 */

function doGet(e) {
  var params = (e && e.parameter) || {};
  var pagina = params.page || 'login';
  var plantilla = HtmlService.createTemplateFromFile(pagina === 'sistema' ? 'sistema' : 'login');
  // El HTML corre dentro de un iframe: no puede leer ?token= ni conocer la URL /exec
  // por sí mismo, así que el servidor se las inyecta.
  plantilla.token = params.token || '';
  plantilla.baseUrl = ScriptApp.getService().getUrl();
  return plantilla.evaluate()
    .setTitle('Sistema de Gestión de Accesos — MVCLM')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Permite usar <?!= include('archivo'); ?> dentro de los HTML si hace falta partirlos en piezas. */
function include(nombreArchivo) {
  return HtmlService.createHtmlOutputFromFile(nombreArchivo).getContent();
}
