/**
 * Codigo.gs
 * Punto de entrada del Web App. Sirve login.html o sistema.html según el
 * parámetro ?page=, igual que en la maqueta estática, pero ahora real.
 */

function doGet(e) {
  var params = (e && e.parameter) || {};
  var pagina = params.page === 'sistema' ? 'sistema' : 'login';

  var plantilla = codigo_cargarPlantilla_(pagina);
  if (!plantilla) {
    return HtmlService.createHtmlOutput(
      '<p style="font-family:sans-serif;padding:24px;">No se encontró el archivo HTML "' + pagina + '" en el proyecto de Apps Script.<br>' +
      'Debe existir un archivo de tipo HTML llamado exactamente <b>' + pagina + '</b> (sin escribir ".html" en el nombre).</p>');
  }

  // El HTML corre dentro de un iframe: no puede leer ?token= ni conocer la URL /exec
  // por sí mismo, así que el servidor se las inyecta.
  plantilla.token = params.token || '';
  plantilla.baseUrl = ScriptApp.getService().getUrl();
  return plantilla.evaluate()
    .setTitle('Sistema de Gestión de Accesos — MVCLM')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Apps Script distingue mayúsculas y el nombre debe coincidir exacto con el archivo.
 * Prueba las variantes habituales (sistema, Sistema, sistema.html, Sistema.html)
 * y devuelve la primera plantilla que exista, o null si ninguna.
 */
function codigo_cargarPlantilla_(nombreBase) {
  var capitalizado = nombreBase.charAt(0).toUpperCase() + nombreBase.slice(1);
  var candidatos = [nombreBase, capitalizado, nombreBase + '.html', capitalizado + '.html'];
  for (var i = 0; i < candidatos.length; i++) {
    try {
      return HtmlService.createTemplateFromFile(candidatos[i]);
    } catch (err) {
      // no existe con ese nombre; se prueba el siguiente
    }
  }
  return null;
}

/** Permite usar <?!= include('archivo'); ?> dentro de los HTML si hace falta partirlos en piezas. */
function include(nombreArchivo) {
  return HtmlService.createHtmlOutputFromFile(nombreArchivo).getContent();
}
