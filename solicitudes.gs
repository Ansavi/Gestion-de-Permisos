/**
 * Solicitudes.gs
 * Todo lo que ve y hace Administración sobre sus propias solicitudes.
 */

var HOJA_SOLICITUDES = 'Solicitudes';
var HOJA_DETALLE_ACCESOS = 'Detalle_Accesos';

/** true si el DNI ya tiene una solicitud registrada (para el aviso en tiempo real del formulario). */
function solicitudes_validarDni(token, dni) {
  login_validarToken_(token);
  var dniLimpio = util_normalizarDni(dni);
  var existe = util_leerComoObjetos(HOJA_SOLICITUDES).some(function (s) {
    return util_normalizarDni(s.DNI) === dniLimpio;
  });
  return { dni: dniLimpio, existe: existe };
}

/**
 * Crea una solicitud en estado "borrador" (aún no notificada).
 * @param {Object} datos { resolucion, fechaResolucion, sede, dni, apellidos, nombres, perfil, adjuntoUrl }
 */
function solicitudes_crear(token, datos) {
  var sesion = login_exigirArea_(token, 'Administración');

  if (!util_soloLetras(datos.apellidos) || !util_soloLetras(datos.nombres)) {
    throw new Error('Nombres y apellidos no pueden contener números.');
  }
  var dni = util_normalizarDni(datos.dni);
  if (dni.length !== 8) throw new Error('El DNI debe tener 8 dígitos.');

  var hoja = util_obtenerHoja(HOJA_SOLICITUDES);
  var id = util_siguienteId(HOJA_SOLICITUDES);
  hoja.appendRow([
    id, util_mayus(datos.resolucion), datos.fechaResolucion, '', datos.adjuntoUrl || '',
    util_mayus(datos.sede), dni, util_mayus(datos.apellidos), util_mayus(datos.nombres), util_mayus(datos.perfil),
    0, 'Pendiente', 'NO', ''
  ]);

  auditoria_registrar(sesion.usuario, 'Solicitud creada', 'ITEM ' + id, sesion.area);
  return { ok: true, item: id };
}

/**
 * Notifica a Informática: registra FECHA_SOLICITUD (recién ahora, no al crear),
 * crea una fila en Detalle_Accesos por cada sistema marcado como requerido,
 * y envía el correo. Esto es lo único que dispara el correo — abrir o
 * guardar un borrador NO notifica.
 * @param {string[]} sistemasRequeridos Nombres tal como están en Mantenimiento_Sistemas
 */
function solicitudes_notificar(token, itemSolicitud, sistemasRequeridos) {
  var sesion = login_exigirArea_(token, 'Administración');
  var hoja = util_obtenerHoja(HOJA_SOLICITUDES);
  var solicitudes = util_leerComoObjetos(HOJA_SOLICITUDES);
  var solicitud = solicitudes.filter(function (s) { return s.ITEM === itemSolicitud; })[0];
  if (!solicitud) throw new Error('No se encontró la solicitud ITEM ' + itemSolicitud);

  var hoy = util_fechaHoy();
  // Columnas: FECHA_SOLICITUD=4, NOTIFICADO=13, FECHA_NOTIFICACION=14 (1-based)
  hoja.getRange(solicitud._fila, 4).setValue(hoy);
  hoja.getRange(solicitud._fila, 13).setValue('SI');
  hoja.getRange(solicitud._fila, 14).setValue(hoy);

  var hojaDetalle = util_obtenerHoja(HOJA_DETALLE_ACCESOS);
  sistemasRequeridos = sistemasRequeridos.map(util_mayus);
  sistemasRequeridos.forEach(function (sistema) {
    var idDetalle = util_siguienteId(HOJA_DETALLE_ACCESOS);
    hojaDetalle.appendRow([idDetalle, itemSolicitud, sistema, 'SI', 'Registrado', '', hoy, sesion.usuario]);
  });

  solicitudes_recalcularAvance_(itemSolicitud);
  auditoria_registrar(sesion.usuario, 'Notificación enviada a Informática', 'ITEM ' + itemSolicitud, sesion.area);

  MailApp.sendEmail({
    to: 'informatica@pj.gob.pe', // TODO: reemplazar por el correo real del área
    subject: 'Nueva solicitud de acceso — ' + solicitud.RESOLUCION,
    body: 'Se registró una nueva solicitud de acceso.\n\n' +
      'Resolución: ' + solicitud.RESOLUCION + '\n' +
      'DNI: ' + solicitud.DNI + '\n' +
      'Nombres: ' + solicitud.NOMBRES + ' ' + solicitud.APELLIDOS + '\n' +
      'Accesos requeridos: ' + sistemasRequeridos.join(', ') + '\n\n' +
      'Ingresa al sistema para gestionarla.'
  });

  return { ok: true };
}

/** Lista las solicitudes de Administración, con filtros opcionales. */
function solicitudes_listar(token, filtros) {
  login_exigirArea_(token, 'Administración');
  filtros = filtros || {};
  var lista = util_leerComoObjetos(HOJA_SOLICITUDES);
  if (filtros.sede) lista = lista.filter(function (s) { return String(s.SEDE).toUpperCase() === util_mayus(filtros.sede); });
  if (filtros.estado) lista = lista.filter(function (s) { return s.ESTADO_GENERAL === filtros.estado; });
  if (filtros.busqueda) {
    var q = filtros.busqueda.toLowerCase();
    lista = lista.filter(function (s) {
      return String(s.DNI).includes(q) || String(s.APELLIDOS).toLowerCase().includes(q);
    });
  }
  return lista;
}

/**
 * Recalcula AVANCE_% y ESTADO_GENERAL de una solicitud a partir del promedio
 * de etapa de sus accesos en Detalle_Accesos. Se llama después de cualquier
 * cambio de estado hecho por Informática (ver accesos_actualizarEstado).
 */
function solicitudes_recalcularAvance_(itemSolicitud) {
  var estados = mantenimiento_listarEstadosProgreso(); // {ESTADO: PORCENTAJE}
  var mapaPorcentaje = {};
  estados.forEach(function (e) { mapaPorcentaje[e.ESTADO] = e.PORCENTAJE; });

  var accesos = util_leerComoObjetos(HOJA_DETALLE_ACCESOS)
    .filter(function (d) { return d.ITEM_SOLICITUD === itemSolicitud && d.REQUERIDO === 'SI'; });

  var avance = 0;
  if (accesos.length > 0) {
    var suma = accesos.reduce(function (acc, d) { return acc + (mapaPorcentaje[d.ESTADO] || 0); }, 0);
    avance = Math.round(suma / accesos.length);
  }
  var estadoGeneral = avance >= 100 ? 'Completado' : (avance > 0 ? 'En proceso' : 'Pendiente');

  var hoja = util_obtenerHoja(HOJA_SOLICITUDES);
  var solicitud = util_leerComoObjetos(HOJA_SOLICITUDES).filter(function (s) { return s.ITEM === itemSolicitud; })[0];
  if (!solicitud) return;
  hoja.getRange(solicitud._fila, 11).setValue(avance);        // AVANCE_%
  hoja.getRange(solicitud._fila, 12).setValue(estadoGeneral); // ESTADO_GENERAL
}
