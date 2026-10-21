/**
 * Auditoria.gs
 * Log único de acciones sensibles. Lo llaman TODAS las páginas, por eso no
 * lleva el prefijo de una sola pantalla. Nunca recibe ni guarda contraseñas.
 */

var HOJA_AUDITORIA = 'Auditoria';

/**
 * Registra un evento de auditoría.
 * @param {string} usuario      Usuario de Credenciales que ejecuta la acción (viene del token de sesión validado)
 * @param {string} accion       Ej: "Notificación enviada a Informática"
 * @param {string} itemAfectado Ej: ITEM de Solicitudes, o el usuario afectado
 * @param {string} area         "Administración" | "Informática" | "Sistema"
 * @param {string} [detalle]    Texto libre — nunca contraseñas ni datos sensibles
 */
function auditoria_registrar(usuario, accion, itemAfectado, area, detalle) {
  var hoja = util_obtenerHoja(HOJA_AUDITORIA);
  hoja.appendRow([
    util_siguienteId(HOJA_AUDITORIA),
    util_fechaHoraAhora(),
    usuario || 'sistema',
    area || '',
    accion,
    itemAfectado || '',
    detalle || ''
  ]);
}

/**
 * Devuelve el historial de auditoría, más reciente primero.
 * @param {number} [limite] Cuántas filas devolver (por defecto 200, para no sobrecargar la vista)
 */
function auditoria_listar(limite) {
  var filas = util_leerComoObjetos(HOJA_AUDITORIA);
  filas.reverse();
  return filas.slice(0, limite || 200);
}
