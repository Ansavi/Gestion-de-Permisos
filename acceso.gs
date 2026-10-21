/**
 * Accesos.gs
 * Todo lo que gestiona Informática sobre los accesos ya notificados.
 * Nota: "credenciales" aquí es la del EMPLEADO NUEVO (usuario de red del
 * solicitado), no confundir con Credenciales.gs / Usuarios.gs, que son
 * quienes entran a ESTE sistema.
 */

/** Bandeja de solicitudes notificadas (para la vista "Gestión de accesos"). */
function accesos_listarPendientes(token) {
  login_exigirArea_(token, 'Informática');
  return util_leerComoObjetos(HOJA_SOLICITUDES)
    .filter(function (s) { return s.NOTIFICADO === 'SI' && s.ESTADO_GENERAL !== 'Completado'; });
}

/** Accesos requeridos de una solicitud puntual (lo que Administración marcó). */
function accesos_obtenerDetalle(token, itemSolicitud) {
  login_exigirArea_(token, 'Informática');
  return util_leerComoObjetos(HOJA_DETALLE_ACCESOS)
    .filter(function (d) { return d.ITEM_SOLICITUD === itemSolicitud && d.REQUERIDO === 'SI'; });
}

/**
 * Cambia la etapa de un acceso puntual (Registrado/Asignado/En progreso/En espera/Resuelto)
 * y recalcula el avance general de la solicitud.
 */
function accesos_actualizarEstado(token, idDetalle, nuevoEstado, valorAsignado) {
  var sesion = login_exigirArea_(token, 'Informática');
  var hoja = util_obtenerHoja(HOJA_DETALLE_ACCESOS);
  var detalle = util_leerComoObjetos(HOJA_DETALLE_ACCESOS).filter(function (d) { return d.ID_DETALLE === idDetalle; })[0];
  if (!detalle) throw new Error('No se encontró el detalle de acceso ' + idDetalle);

  hoja.getRange(detalle._fila, 5).setValue(nuevoEstado); // ESTADO
  if (valorAsignado) hoja.getRange(detalle._fila, 6).setValue(valorAsignado); // VALOR_ASIGNADO
  hoja.getRange(detalle._fila, 7).setValue(util_fechaHoy());  // FECHA_ACTUALIZACION
  hoja.getRange(detalle._fila, 8).setValue(sesion.usuario);   // ACTUALIZADO_POR

  solicitudes_recalcularAvance_(detalle.ITEM_SOLICITUD);
  auditoria_registrar(sesion.usuario, 'Avance actualizado a "' + nuevoEstado + '"', 'ITEM ' + detalle.ITEM_SOLICITUD, sesion.area, 'Sistema: ' + detalle.SISTEMA);

  return { ok: true };
}

/**
 * Genera usuario de red + contraseña temporal para el EMPLEADO NUEVO (no para
 * el personal de este sistema). La contraseña se devuelve una sola vez al
 * cliente para que Informática la comunique por un canal seguro — no se
 * guarda en ninguna hoja.
 */
function accesos_generarCredencialesEmpleado(token, idDetalle, nombres, apellidos) {
  var sesion = login_exigirArea_(token, 'Informática');
  var usuario = (nombres.charAt(0) + apellidos.split(' ')[0])
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  var claveTemporal = util_generarClaveTemporal();

  accesos_actualizarEstado(token, idDetalle, 'Resuelto', usuario);
  auditoria_registrar(sesion.usuario, 'Credenciales generadas para Usuario de red (' + usuario + ')', 'DETALLE ' + idDetalle, sesion.area, 'Contraseña no almacenada');

  return { usuario: usuario, claveTemporal: claveTemporal }; // se muestra una sola vez en pantalla
}

/**
 * Genera el documento a partir de la plantilla asociada al sistema (ej. RENIEC
 * — Consulta de menores), reemplazando los marcadores {{CAMPO}} con los datos
 * de la solicitud, y lo guarda en Drive.
 */
function accesos_generarDocumento(token, itemSolicitud, nombreSistema) {
  var sesion = login_exigirArea_(token, 'Informática');

  var plantilla = mantenimiento_listarPlantillas()
    .filter(function (p) { return p.SISTEMA_ASOCIADO === nombreSistema && p.ESTADO === 'Activo'; })[0];
  if (!plantilla) throw new Error('No hay una plantilla activa asociada a "' + nombreSistema + '".');

  var solicitud = util_leerComoObjetos(HOJA_SOLICITUDES).filter(function (s) { return s.ITEM === itemSolicitud; })[0];
  if (!solicitud) throw new Error('No se encontró la solicitud ITEM ' + itemSolicitud);

  var archivoPlantilla = DriveApp.getFileById(plantilla.ID_DOC_DRIVE_PLANTILLA);
  var copia = archivoPlantilla.makeCopy(plantilla.NOMBRE_PLANTILLA + ' — ' + solicitud.DNI);
  var doc = DocumentApp.openById(copia.getId());
  var cuerpo = doc.getBody();

  var valores = {
    NOMBRES: solicitud.NOMBRES,
    APELLIDOS: solicitud.APELLIDOS,
    DNI: solicitud.DNI,
    FECHA: util_fechaHoy()
  };
  String(plantilla.CAMPOS_A_COMBINAR).split(',').forEach(function (campo) {
    campo = campo.trim();
    cuerpo.replaceText('{{' + campo + '}}', valores[campo] || '');
  });
  doc.saveAndClose();

  // TODO: mover `copia` a la carpeta restringida de documentos generados
  // (ver "Notas de revisión": estos documentos pueden contener datos de menores).
  auditoria_registrar(sesion.usuario, 'Documento generado: ' + plantilla.NOMBRE_PLANTILLA, 'ITEM ' + itemSolicitud, sesion.area);

  return { ok: true, urlDocumento: copia.getUrl() };
}
