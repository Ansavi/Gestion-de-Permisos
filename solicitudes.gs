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
 * @param {Object} datos { resolucion, fechaResolucion, sede, dni, apellidos, nombres, perfil, adjunto: { nombre, base64 } }
 * El PDF de la resolución es obligatorio: se valida, se guarda en la carpeta restringida de Drive y su enlace
 * queda en ADJUNTO_URL. Si algo falla después de guardarlo, el archivo se manda a la papelera (no quedan huérfanos).
 */
function solicitudes_crear(token, datos) {
  var sesion = login_exigirArea_(token, 'Administración');

  if (!util_soloLetras(datos.apellidos) || !util_soloLetras(datos.nombres)) {
    throw new Error('Nombres y apellidos no pueden contener números.');
  }
  var dni = util_normalizarDni(datos.dni);
  if (dni.length !== 8) throw new Error('El DNI debe tener 8 dígitos.');

  var adjuntoUrl = solicitudes_guardarAdjunto_(datos.adjunto, util_mayus(datos.resolucion), dni);
  try {
    var hoja = util_obtenerHoja(HOJA_SOLICITUDES);
    var id = util_siguienteId(HOJA_SOLICITUDES);
    var fila = [
      id, util_mayus(datos.resolucion), datos.fechaResolucion, '', adjuntoUrl,
      util_mayus(datos.sede), dni, util_mayus(datos.apellidos), util_mayus(datos.nombres), util_mayus(datos.perfil),
      0, 'Pendiente', 'NO', ''
    ];
    var colCreador = solicitudes_asegurarColumnaCreadoPor_(); // quién la registró: define quién la ve después
    while (fila.length < colCreador) fila.push('');
    fila[colCreador - 1] = String(sesion.usuario).toLowerCase();
    hoja.appendRow(fila);
  } catch (error) {
    solicitudes_borrarAdjunto_(adjuntoUrl); // no dejar un PDF suelto si la solicitud no se pudo guardar
    throw error;
  }

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
  if (String(solicitud.CREADO_POR || '').toLowerCase() !== String(sesion.usuario).toLowerCase()) {
    throw new Error('Solo puedes notificar las solicitudes que tú registraste.');
  }

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

/** Lista las solicitudes que registró la persona que consulta (Administración), con filtros opcionales. */
function solicitudes_listar(token, filtros) {
  var sesion = login_exigirArea_(token, 'Administración');
  filtros = filtros || {};
  // Solo las solicitudes que registró esta persona (las anteriores a CREADO_POR no tienen dueño y no se muestran).
  var lista = util_leerComoObjetos(HOJA_SOLICITUDES).filter(function (s) {
    return String(s.CREADO_POR || '').toLowerCase() === String(sesion.usuario).toLowerCase();
  });
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
  // Los accesos se entregan de forma escalonada: "Pendiente" = ningún acceso ha salido de la primera etapa;
  // "En proceso" = al menos uno avanzó pero faltan otros; "Completado" = todos resueltos.
  var etapas = Object.keys(mapaPorcentaje).map(function (k) { return Number(mapaPorcentaje[k]) || 0; });
  var primeraEtapa = Math.min.apply(null, etapas);
  var ningunoAvanzo = accesos.every(function (d) { return (Number(mapaPorcentaje[d.ESTADO]) || 0) <= primeraEtapa; });
  var estadoGeneral = avance >= 100 ? 'Completado' : (ningunoAvanzo ? 'Pendiente' : 'En proceso');

  var hoja = util_obtenerHoja(HOJA_SOLICITUDES);
  var solicitud = util_leerComoObjetos(HOJA_SOLICITUDES).filter(function (s) { return s.ITEM === itemSolicitud; })[0];
  if (!solicitud) return;
  hoja.getRange(solicitud._fila, 11).setValue(avance);        // AVANCE_%
  hoja.getRange(solicitud._fila, 12).setValue(estadoGeneral); // ESTADO_GENERAL
}

/**
 * Lista para "Gestión de solicitudes" (Administración): cada solicitud con su avance general
 * y el DETALLE por acceso (etapa y %), para saber qué ya está resuelto y qué falta.
 * Reutiliza solicitudes_listar (sin modificarla) y devuelve todo como texto/números simples,
 * porque google.script.run no puede transportar fechas (Date).
 * @param {Object} filtros { sede, estado, busqueda } — todos opcionales
 */
function solicitudes_listarConAvance(token, filtros) {
  filtros = filtros || {};
  var lista = solicitudes_listar(token, { sede: filtros.sede, busqueda: filtros.busqueda }); // valida el área

  var porcentajes = {};
  mantenimiento_listarEstadosProgreso().forEach(function (e) { porcentajes[e.ESTADO] = Number(e.PORCENTAJE) || 0; });
  var minimo = Math.min.apply(null, Object.keys(porcentajes).map(function (k) { return porcentajes[k]; }));

  var accesosPorItem = {};
  util_leerComoObjetos(HOJA_DETALLE_ACCESOS).forEach(function (d) {
    if (d.REQUERIDO !== 'SI') return;
    (accesosPorItem[d.ITEM_SOLICITUD] = accesosPorItem[d.ITEM_SOLICITUD] || []).push({
      SISTEMA: d.SISTEMA, ESTADO: d.ESTADO, PORCENTAJE: porcentajes[d.ESTADO] || 0
    });
  });

  var resultado = lista.map(function (s) {
    var accesos = accesosPorItem[s.ITEM] || [];
    var suma = accesos.reduce(function (a, x) { return a + x.PORCENTAJE; }, 0);
    var avance = accesos.length ? Math.round(suma / accesos.length) : 0;
    var todosEnLaPrimeraEtapa = accesos.every(function (x) { return x.PORCENTAJE <= minimo; });
    return {
      ITEM: s.ITEM,
      RESOLUCION: s.RESOLUCION,
      FECHA_RESOLUCION: util_fechaTexto(s.FECHA_RESOLUCION),
      FECHA_SOLICITUD: util_fechaTexto(s.FECHA_SOLICITUD),
      SEDE: s.SEDE,
      DNI: util_dniTexto(s.DNI),
      APELLIDOS: s.APELLIDOS,
      NOMBRES: s.NOMBRES,
      PERFIL: s.PERFIL,
      NOTIFICADO: s.NOTIFICADO,
      TIENE_ADJUNTO: !!String(s.ADJUNTO_URL || '').trim(),
      AVANCE: avance,
      ESTADO_GENERAL: avance >= 100 ? 'Completado' : (accesos.length === 0 || todosEnLaPrimeraEtapa ? 'Pendiente' : 'En proceso'),
      ACCESOS_RESUELTOS: accesos.filter(function (x) { return x.PORCENTAJE >= 100; }).length,
      ACCESOS_TOTAL: accesos.length,
      ACCESOS: accesos
    };
  });

  if (filtros.estado) resultado = resultado.filter(function (r) { return r.ESTADO_GENERAL === filtros.estado; });
  return resultado.reverse(); // la más reciente primero
}

/**
 * Columna CREADO_POR (usuario que registró la solicitud). Si la hoja es anterior y no la tiene,
 * la agrega sola en la columna 15. Devuelve el número de columna.
 */
function solicitudes_asegurarColumnaCreadoPor_() {
  var hoja = util_obtenerHoja(HOJA_SOLICITUDES);
  var encabezados = hoja.getRange(1, 1, 1, Math.max(hoja.getLastColumn(), 15)).getValues()[0];
  var col = encabezados.indexOf('CREADO_POR') + 1;
  if (col > 0) return col;
  if (encabezados[14] !== '' && encabezados[14] != null) {
    throw new Error('La columna 15 de Solicitudes ya tiene el encabezado "' + encabezados[14] + '"; no se pudo agregar CREADO_POR.');
  }
  hoja.getRange(1, 15).setValue('CREADO_POR').setFontWeight('bold');
  return 15;
}

/** Porcentaje de cada etapa (Mantenimiento_Estados_Progreso) y cuál es la primera etapa. */
function solicitudes_mapaEtapas_() {
  var porcentajes = {};
  mantenimiento_listarEstadosProgreso().forEach(function (e) { porcentajes[e.ESTADO] = Number(e.PORCENTAJE) || 0; });
  var valores = Object.keys(porcentajes).map(function (k) { return porcentajes[k]; });
  return { porcentajes: porcentajes, primeraEtapa: valores.length ? Math.min.apply(null, valores) : 0 };
}

/**
 * Avance y estado general a partir de los % de cada acceso. Misma regla en toda la app:
 * Pendiente = ningún acceso salió de la primera etapa; En proceso = al menos uno avanzó pero faltan;
 * Completado = todos resueltos.
 */
function solicitudes_resumenAvance_(pcts, primeraEtapa) {
  var avance = pcts.length ? Math.round(pcts.reduce(function (a, x) { return a + x; }, 0) / pcts.length) : 0;
  var ningunoAvanzo = pcts.every(function (x) { return x <= primeraEtapa; });
  return {
    avance: avance,
    estado: avance >= 100 ? 'Completado' : (ningunoAvanzo ? 'Pendiente' : 'En proceso'),
    resueltos: pcts.filter(function (x) { return x >= 100; }).length
  };
}

/** % de cada acceso REQUERIDO, agrupado por solicitud: { ITEM: [pct, pct, ...] }. */
function solicitudes_porcentajesPorItem_(porcentajes) {
  var porItem = {};
  util_leerComoObjetos(HOJA_DETALLE_ACCESOS).forEach(function (d) {
    if (d.REQUERIDO !== 'SI') return;
    (porItem[d.ITEM_SOLICITUD] = porItem[d.ITEM_SOLICITUD] || []).push(porcentajes[d.ESTADO] || 0);
  });
  return porItem;
}


// ---------- PDF DE LA RESOLUCIÓN (adjunto) ----------

var ADJUNTO_MAX_BYTES = 10 * 1024 * 1024; // 10 MB

/**
 * Carpeta restringida donde se guardan los PDF de las resoluciones. Si todavía no existe, se crea sola
 * y su ID queda en las propiedades del script.
 */
function solicitudes_carpetaResoluciones_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('CARPETA_RESOLUCIONES_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* la borraron: se vuelve a crear */ }
  }
  var existentes = DriveApp.getFoldersByName('MVCLM - Resoluciones (RESTRINGIDO)');
  var carpeta = existentes.hasNext() ? existentes.next() : DriveApp.createFolder('MVCLM - Resoluciones (RESTRINGIDO)');
  props.setProperty('CARPETA_RESOLUCIONES_ID', carpeta.getId());
  return carpeta;
}

/** Valida el PDF recibido del navegador y lo guarda en Drive. Devuelve la URL del archivo. */
function solicitudes_guardarAdjunto_(adjunto, resolucion, dni) {
  if (!adjunto || !adjunto.base64) throw new Error('Adjunta el PDF de la resolución.');

  var contenido = String(adjunto.base64).replace(/^data:[^,]*,/, ''); // por si llega como "data:...;base64,"
  var bytes;
  try { bytes = Utilities.base64Decode(contenido); } catch (e) { throw new Error('El archivo adjunto no se pudo leer.'); }
  if (!bytes.length) throw new Error('El archivo adjunto está vacío.');
  if (bytes.length > ADJUNTO_MAX_BYTES) {
    throw new Error('El PDF pesa más de ' + (ADJUNTO_MAX_BYTES / 1024 / 1024) + ' MB. Reduce su tamaño y vuelve a intentar.');
  }
  // Un PDF real empieza con "%PDF-": evita aceptar otro tipo de archivo renombrado a .pdf
  var inicio = [37, 80, 68, 70, 45];
  for (var i = 0; i < inicio.length; i++) {
    if (bytes[i] !== inicio[i]) throw new Error('El archivo adjunto no es un PDF válido.');
  }

  var limpio = function (t) { return String(t || '').replace(/[^A-Za-z0-9._-]+/g, '_').replace(/^_+|_+$/g, ''); };
  var nombre = 'RESOLUCION_' + (limpio(resolucion) || 'SIN-NUMERO') + '_DNI_' + dni + '_' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmm') + '.pdf';

  var archivo = solicitudes_carpetaResoluciones_().createFile(Utilities.newBlob(bytes, 'application/pdf', nombre));
  return archivo.getUrl();
}

/** Manda a la papelera un PDF recién guardado (se usa si la solicitud no llegó a registrarse). */
function solicitudes_borrarAdjunto_(url) {
  try {
    var m = String(url || '').match(/[-\w]{25,}/);
    if (m) DriveApp.getFileById(m[0]).setTrashed(true);
  } catch (e) { /* si no se pudo, no se oculta el error original */ }
}

/**
 * Entrega el PDF de una solicitud al navegador (los archivos quedan privados en Drive: nadie los abre
 * por enlace). Administración solo puede ver los de sus propias solicitudes; Informática, los de las
 * solicitudes ya notificadas. Cada consulta queda en la auditoría.
 */
function solicitudes_obtenerAdjunto(token, itemSolicitud) {
  var sesion = login_validarToken_(token);
  var solicitud = util_leerComoObjetos(HOJA_SOLICITUDES)
    .filter(function (s) { return Number(s.ITEM) === Number(itemSolicitud); })[0];
  if (!solicitud) throw new Error('No se encontró la solicitud ITEM ' + itemSolicitud);

  if (sesion.area === 'Administración') {
    if (String(solicitud.CREADO_POR || '').toLowerCase() !== String(sesion.usuario).toLowerCase()) {
      throw new Error('Solo puedes ver el PDF de las solicitudes que tú registraste.');
    }
  } else if (sesion.area === 'Informática') {
    if (solicitud.NOTIFICADO !== 'SI') throw new Error('Esta solicitud todavía no fue notificada a Informática.');
  } else {
    throw new Error('No tienes permiso para ver este documento.');
  }

  var m = String(solicitud.ADJUNTO_URL || '').match(/[-\w]{25,}/);
  if (!m) throw new Error('Esta solicitud no tiene PDF adjunto.');
  var archivo = DriveApp.getFileById(m[0]);

  auditoria_registrar(sesion.usuario, 'Resolución (PDF) consultada', 'ITEM ' + itemSolicitud, sesion.area);
  return { nombre: archivo.getName(), base64: Utilities.base64Encode(archivo.getBlob().getBytes()) };
}
