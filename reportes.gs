/**
 * Reportes.gs
 * NOTA (ver "Revisión de requerimientos", punto 2 pendiente): todavía no
 * definimos exactamente qué columnas exporta cada área. Esta es una primera
 * versión funcional: Administración exporta SUS solicitudes, Informática
 * exporta el detalle completo. Ajustar cuando se confirme el punto pendiente.
 */

function reportes_exportarSolicitudes(token) {
  var sesion = login_validarToken_(token);
  var datos = sesion.area === 'Administración'
    ? solicitudes_listar(token, {})
    : util_leerComoObjetos(HOJA_SOLICITUDES);

  var hojaTemporal = SpreadsheetApp.create('Reporte_Solicitudes_' + util_fechaHoy());
  var hoja = hojaTemporal.getSheets()[0];
  if (datos.length === 0) return { ok: false, mensaje: 'No hay datos para exportar.' };

  var encabezados = Object.keys(datos[0]).filter(function (k) { return k !== '_fila'; });
  hoja.appendRow(encabezados);
  datos.forEach(function (fila) {
    hoja.appendRow(encabezados.map(function (e) { return fila[e]; }));
  });

  auditoria_registrar(sesion.usuario, 'Reporte de solicitudes exportado', '', sesion.area);
  return { ok: true, urlHoja: hojaTemporal.getUrl() };
}
