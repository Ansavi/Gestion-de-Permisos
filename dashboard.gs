/**
 * Dashboard.gs
 * Resumen de la pantalla Dashboard. Administración ve el resumen de SUS solicitudes;
 * Informática ve el de todas las solicitudes ya notificadas.
 */

function dashboard_obtenerResumen(token) {
  var sesion = login_validarToken_(token);
  var propias = sesion.area === 'Administración';

  var mapa = solicitudes_mapaEtapas_();
  var porItem = solicitudes_porcentajesPorItem_(mapa.porcentajes);

  var solicitudes = util_leerComoObjetos(HOJA_SOLICITUDES).filter(function (s) {
    return propias
      ? String(s.CREADO_POR || '').toLowerCase() === String(sesion.usuario).toLowerCase()
      : s.NOTIFICADO === 'SI';
  });

  var resumen = { alcance: propias ? 'propias' : 'todas', pendientes: 0, enProceso: 0, completadas: 0, total: solicitudes.length };
  var porSede = {};
  solicitudes.forEach(function (s) {
    var r = solicitudes_resumenAvance_(porItem[s.ITEM] || [], mapa.primeraEtapa);
    if (r.estado === 'Completado') resumen.completadas++;
    else if (r.estado === 'En proceso') resumen.enProceso++;
    else resumen.pendientes++;
    var sede = String(s.SEDE || '').trim() || 'SIN SEDE';
    porSede[sede] = (porSede[sede] || 0) + 1;
  });
  resumen.porSede = Object.keys(porSede).sort().map(function (k) { return { SEDE: k, TOTAL: porSede[k] }; });
  return resumen;
}
