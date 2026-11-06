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
    .filter(function (p) { return util_mayus(p.SISTEMA_ASOCIADO) === util_mayus(nombreSistema) && p.ESTADO === 'Activo'; })[0];
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

  // El documento puede contener datos de menores: se guarda en la carpeta restringida creada por
  // config_crearCarpetas (si existe), no suelto en "Mi unidad".
  var carpetaId = PropertiesService.getScriptProperties().getProperty('CARPETA_DOCUMENTOS_GENERADOS_ID');
  if (carpetaId) copia.moveTo(DriveApp.getFolderById(carpetaId));
  auditoria_registrar(sesion.usuario, 'Documento generado: ' + plantilla.NOMBRE_PLANTILLA, 'ITEM ' + itemSolicitud, sesion.area);

  return { ok: true, urlDocumento: copia.getUrl() };
}


// ---------- GESTIÓN DE ACCESOS (pantalla de Informática) ----------
// Funciones nuevas que devuelven solo texto y números: las anteriores de este archivo devuelven
// filas con fechas de Sheets (Date), que google.script.run no puede llevar al navegador.

/** Bandeja: solicitudes ya notificadas. Por defecto solo las que aún no están completadas. */
function accesos_listarBandeja(token, incluirCompletadas) {
  login_exigirArea_(token, 'Informática');
  var mapa = solicitudes_mapaEtapas_();
  var porItem = solicitudes_porcentajesPorItem_(mapa.porcentajes);

  var lista = util_leerComoObjetos(HOJA_SOLICITUDES)
    .filter(function (s) { return s.NOTIFICADO === 'SI'; })
    .map(function (s) {
      var pcts = porItem[s.ITEM] || [];
      var r = solicitudes_resumenAvance_(pcts, mapa.primeraEtapa);
      return {
        ITEM: s.ITEM, RESOLUCION: s.RESOLUCION, FECHA_SOLICITUD: util_fechaTexto(s.FECHA_SOLICITUD),
        DNI: util_dniTexto(s.DNI), APELLIDOS: s.APELLIDOS, NOMBRES: s.NOMBRES, SEDE: s.SEDE, PERFIL: s.PERFIL,
        AVANCE: r.avance, ESTADO_GENERAL: r.estado, ACCESOS_TOTAL: pcts.length, ACCESOS_RESUELTOS: r.resueltos
      };
    });
  if (!incluirCompletadas) lista = lista.filter(function (r) { return r.ESTADO_GENERAL !== 'Completado'; });
  return lista; // en orden de llegada: la más antigua primero
}

/** Datos de una solicitud y de sus accesos, listos para gestionarlos en pantalla. */
function accesos_obtenerDetalleSolicitud(token, itemSolicitud) {
  login_exigirArea_(token, 'Informática');
  var solicitud = util_leerComoObjetos(HOJA_SOLICITUDES)
    .filter(function (s) { return Number(s.ITEM) === Number(itemSolicitud); })[0];
  if (!solicitud) throw new Error('No se encontró la solicitud ITEM ' + itemSolicitud);

  var etapas = mantenimiento_listarEstadosProgreso()
    .sort(function (x, y) { return Number(x.ORDEN) - Number(y.ORDEN); })
    .map(function (e) { return { ESTADO: e.ESTADO, PORCENTAJE: Number(e.PORCENTAJE) || 0 }; });

  var requiereValor = {};
  util_leerComoObjetos(HOJA_SISTEMAS).forEach(function (sis) { requiereValor[util_mayus(sis.NOMBRE_SISTEMA)] = sis.REQUIERE_VALOR === 'SI'; });
  var conPlantilla = mantenimiento_listarPlantillas()
    .filter(function (p) { return p.ESTADO === 'Activo'; })
    .map(function (p) { return util_mayus(p.SISTEMA_ASOCIADO); });

  var accesos = util_leerComoObjetos(HOJA_DETALLE_ACCESOS)
    .filter(function (d) { return Number(d.ITEM_SOLICITUD) === Number(itemSolicitud) && d.REQUERIDO === 'SI'; })
    .map(function (d) {
      var nombre = util_mayus(d.SISTEMA);
      return {
        ID_DETALLE: d.ID_DETALLE, SISTEMA: d.SISTEMA, ESTADO: d.ESTADO, VALOR_ASIGNADO: d.VALOR_ASIGNADO || '',
        REQUIERE_VALOR: !!requiereValor[nombre], TIENE_PLANTILLA: conPlantilla.indexOf(nombre) !== -1,
        ES_USUARIO_RED: nombre === 'USUARIO DE RED'
      };
    });

  return {
    solicitud: {
      ITEM: solicitud.ITEM, RESOLUCION: solicitud.RESOLUCION, FECHA_RESOLUCION: util_fechaTexto(solicitud.FECHA_RESOLUCION),
      FECHA_SOLICITUD: util_fechaTexto(solicitud.FECHA_SOLICITUD), DNI: util_dniTexto(solicitud.DNI),
      TIENE_ADJUNTO: !!String(solicitud.ADJUNTO_URL || '').trim(),
      APELLIDOS: solicitud.APELLIDOS, NOMBRES: solicitud.NOMBRES, SEDE: solicitud.SEDE, PERFIL: solicitud.PERFIL
    },
    etapas: etapas,
    accesos: accesos
  };
}

/**
 * Guarda de una vez los cambios de etapa (y del valor asignado, ej. usuario de red o correo) de una solicitud.
 * @param {Object[]} cambios  [{ idDetalle, estado, valor }]
 * Valida TODO antes de escribir: una etapa que no exista en Mantenimiento_Estados_Progreso contaría como 0%.
 */
function accesos_guardarCambios(token, itemSolicitud, cambios) {
  login_exigirArea_(token, 'Informática');
  cambios = cambios || [];

  var etapasValidas = {};
  mantenimiento_listarEstadosProgreso().forEach(function (e) { etapasValidas[e.ESTADO] = true; });
  var idsDeLaSolicitud = util_leerComoObjetos(HOJA_DETALLE_ACCESOS)
    .filter(function (d) { return Number(d.ITEM_SOLICITUD) === Number(itemSolicitud); })
    .map(function (d) { return Number(d.ID_DETALLE); });

  cambios.forEach(function (c) {
    if (!etapasValidas[c.estado]) throw new Error('La etapa "' + c.estado + '" no existe en Mantenimiento.');
    if (idsDeLaSolicitud.indexOf(Number(c.idDetalle)) === -1) throw new Error('Ese acceso no pertenece a la solicitud ' + itemSolicitud + '.');
  });

  cambios.forEach(function (c) {
    accesos_actualizarEstado(token, Number(c.idDetalle), c.estado, String(c.valor || '').trim());
  });
  return { ok: true, guardados: cambios.length };
}

/** Acciones registradas sobre una solicitud (lo más reciente primero). Solo texto, sin fechas Date. */
function accesos_historialSolicitud(token, itemSolicitud) {
  login_exigirArea_(token, 'Informática');
  var clave = 'ITEM ' + itemSolicitud;
  return util_leerComoObjetos(HOJA_AUDITORIA)
    .filter(function (r) { return String(r.ITEM_AFECTADO) === clave; })
    .reverse().slice(0, 30)
    .map(function (r) {
      return { FECHA_HORA: util_fechaHoraTexto(r.FECHA_HORA), USUARIO: r.USUARIO, ACCION: r.ACCION, DETALLE: r.DETALLE || '' };
    });
}
