/**
 * Correo.gs
 * Configuración de correo (a quiénes se avisa y de qué) y envío de los avisos.
 *
 * IMPORTANTE: los correos salen desde la cuenta de Google con la que se PUBLICA la web app
 * ("Ejecutar como: yo"). MailApp no usa usuario ni contraseña de correo, así que este sistema
 * no guarda ninguna. Si la contraseña de esa cuenta cambia, no hay nada que actualizar aquí.
 * Para cambiar la cuenta remitente hay que publicar la aplicación con esa otra cuenta.
 *
 * Hojas (se crean solas la primera vez que se usan):
 *  - Correo_Contactos:     ID_CONTACTO, NOMBRE, CORREO, EVENTOS, ESTADO
 *  - Correo_Configuracion: CLAVE, VALOR   (ENVIO_ACTIVO, NOMBRE_REMITENTE, RESPONDER_A)
 */

var HOJA_CORREO_CONTACTOS = 'Correo_Contactos';
var HOJA_CORREO_CONFIG = 'Correo_Configuracion';

/** Avisos que el sistema puede enviar. Cada contacto elige cuáles recibe. */
var CORREO_EVENTOS_ = [
  { codigo: 'NUEVA_SOLICITUD', descripcion: 'Cuando Administración notifica una solicitud nueva' },
  { codigo: 'SOLICITUD_COMPLETADA', descripcion: 'Cuando todos los accesos de una solicitud quedan resueltos' }
];

// ---------------------------------------------------------------------
// Hojas y lectura de datos
// ---------------------------------------------------------------------

/** Crea las hojas de correo si faltan, y la configuración inicial (envío DESACTIVADO hasta que Informática lo active). */
function correo_asegurarHojas_() {
  [
    { nombre: HOJA_CORREO_CONTACTOS, encabezados: ['ID_CONTACTO', 'NOMBRE', 'CORREO', 'EVENTOS', 'ESTADO'] },
    { nombre: HOJA_CORREO_CONFIG, encabezados: ['CLAVE', 'VALOR'] }
  ].forEach(function (d) {
    if (SS.getSheetByName(d.nombre)) return;
    var hoja = SS.insertSheet(d.nombre);
    hoja.getRange(1, 1, 1, d.encabezados.length).setValues([d.encabezados]).setFontWeight('bold');
    hoja.setFrozenRows(1);
  });

  var existentes = {};
  util_leerComoObjetos(HOJA_CORREO_CONFIG).forEach(function (f) { existentes[f.CLAVE] = true; });
  var hojaConfig = util_obtenerHoja(HOJA_CORREO_CONFIG);
  [['ENVIO_ACTIVO', 'NO'], ['NOMBRE_REMITENTE', 'GESTIÓN DE ALTA DE USUARIO — MVCLM'], ['RESPONDER_A', '']].forEach(function (par) {
    if (!existentes[par[0]]) hojaConfig.appendRow(par);
  });
}

function correo_leerConfig_() {
  correo_asegurarHojas_();
  var c = {};
  util_leerComoObjetos(HOJA_CORREO_CONFIG).forEach(function (f) { c[f.CLAVE] = f.VALOR; });
  return {
    envioActivo: String(c.ENVIO_ACTIVO || '').toUpperCase() === 'SI',
    remitente: String(c.NOMBRE_REMITENTE || ''),
    responderA: String(c.RESPONDER_A || '')
  };
}

function correo_escribirConfig_(clave, valor) {
  var hoja = util_obtenerHoja(HOJA_CORREO_CONFIG);
  var fila = util_leerComoObjetos(HOJA_CORREO_CONFIG).filter(function (f) { return f.CLAVE === clave; })[0];
  if (fila) hoja.getRange(fila._fila, 2).setValue(valor);
  else hoja.appendRow([clave, valor]);
}

function correo_leerContactos_() {
  correo_asegurarHojas_();
  return util_leerComoObjetos(HOJA_CORREO_CONTACTOS)
    .filter(function (c) { return String(c.CORREO || '').trim() !== ''; })
    .map(function (c) {
      return {
        _fila: c._fila,
        ID_CONTACTO: c.ID_CONTACTO,
        NOMBRE: c.NOMBRE,
        CORREO: String(c.CORREO).trim().toLowerCase(),
        EVENTOS: String(c.EVENTOS || '').split(',').map(function (e) { return e.trim(); }).filter(Boolean),
        ESTADO: c.ESTADO
      };
    });
}

function correo_correoValido_(texto) {
  var t = String(texto || '').trim();
  return t.length <= 120 && /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(t);
}

/** Valida los datos de un contacto. Devuelve { nombre, correo, eventos } limpios o lanza un error con el motivo. */
function correo_validarContacto_(datos, idPropio) {
  var nombre = util_mayus(datos.nombre);
  var correo = String(datos.correo || '').trim().toLowerCase();
  if (!nombre) throw new Error('Escribe el nombre del contacto.');
  if (!correo_correoValido_(correo)) throw new Error('El correo no es válido.');

  var codigos = CORREO_EVENTOS_.map(function (e) { return e.codigo; });
  var eventos = (datos.eventos || []).filter(function (e, i, a) { return a.indexOf(e) === i; });
  if (!eventos.length) throw new Error('Elige al menos un aviso para este contacto.');
  eventos.forEach(function (e) { if (codigos.indexOf(e) === -1) throw new Error('Aviso desconocido: ' + e); });

  var repetido = correo_leerContactos_().some(function (c) {
    return c.CORREO === correo && Number(c.ID_CONTACTO) !== Number(idPropio);
  });
  if (repetido) throw new Error('Ya existe un contacto con ese correo.');
  return { nombre: nombre, correo: correo, eventos: eventos };
}

// ---------------------------------------------------------------------
// Pantalla "Configuración de correo" (solo Informática)
// ---------------------------------------------------------------------

function correo_obtenerConfiguracion(token) {
  login_exigirArea_(token, 'Informática');
  var cfg = correo_leerConfig_();
  var cuenta = '';
  try { cuenta = Session.getEffectiveUser().getEmail(); } catch (e) { cuenta = ''; }
  var cuota = null;
  try { cuota = MailApp.getRemainingDailyQuota(); } catch (e) { cuota = null; }

  return {
    envioActivo: cfg.envioActivo,
    remitente: cfg.remitente,
    responderA: cfg.responderA,
    cuentaEnvio: cuenta,
    cuotaRestante: cuota,
    eventos: CORREO_EVENTOS_,
    contactos: correo_leerContactos_().map(function (c) {
      return { ID_CONTACTO: c.ID_CONTACTO, NOMBRE: c.NOMBRE, CORREO: c.CORREO, EVENTOS: c.EVENTOS, ESTADO: c.ESTADO };
    })
  };
}

/** @param {Object} ajustes { envioActivo: boolean, remitente: string, responderA: string } */
function correo_guardarAjustes(token, ajustes) {
  var sesion = login_exigirArea_(token, 'Informática');
  var remitente = String(ajustes.remitente || '').trim();
  var responderA = String(ajustes.responderA || '').trim().toLowerCase();
  if (!remitente) throw new Error('Escribe el nombre que verán quienes reciban los correos.');
  if (remitente.length > 80) throw new Error('El nombre del remitente es demasiado largo (máximo 80 caracteres).');
  if (responderA && !correo_correoValido_(responderA)) throw new Error('El correo para responder no es válido.');

  correo_asegurarHojas_();
  correo_escribirConfig_('ENVIO_ACTIVO', ajustes.envioActivo ? 'SI' : 'NO');
  correo_escribirConfig_('NOMBRE_REMITENTE', util_mayus(remitente));
  correo_escribirConfig_('RESPONDER_A', responderA);

  auditoria_registrar(sesion.usuario, 'Ajustes de correo modificados', '', sesion.area, 'Envío ' + (ajustes.envioActivo ? 'activado' : 'desactivado'));
  return { ok: true };
}

/** @param {Object} datos { nombre, correo, eventos: string[] } */
function correo_crearContacto(token, datos) {
  var sesion = login_exigirArea_(token, 'Informática');
  var d = correo_validarContacto_(datos, null);
  var id = util_siguienteId(HOJA_CORREO_CONTACTOS);
  util_obtenerHoja(HOJA_CORREO_CONTACTOS).appendRow([id, d.nombre, d.correo, d.eventos.join(','), 'Activo']);
  auditoria_registrar(sesion.usuario, 'Contacto de correo creado: ' + d.nombre, '', sesion.area);
  return { ok: true, id: id };
}

function correo_editarContacto(token, idContacto, datos) {
  var sesion = login_exigirArea_(token, 'Informática');
  var fila = correo_leerContactos_().filter(function (c) { return Number(c.ID_CONTACTO) === Number(idContacto); })[0];
  if (!fila) throw new Error('No se encontró el contacto ' + idContacto);
  var d = correo_validarContacto_(datos, idContacto);
  util_obtenerHoja(HOJA_CORREO_CONTACTOS).getRange(fila._fila, 2, 1, 3).setValues([[d.nombre, d.correo, d.eventos.join(',')]]);
  auditoria_registrar(sesion.usuario, 'Contacto de correo editado: ' + d.nombre, '', sesion.area);
  return { ok: true };
}

function correo_cambiarEstadoContacto(token, idContacto, nuevoEstado) {
  var sesion = login_exigirArea_(token, 'Informática');
  if (['Activo', 'Inactivo'].indexOf(nuevoEstado) === -1) throw new Error('Estado no válido: ' + nuevoEstado);
  var fila = correo_leerContactos_().filter(function (c) { return Number(c.ID_CONTACTO) === Number(idContacto); })[0];
  if (!fila) throw new Error('No se encontró el contacto ' + idContacto);
  util_obtenerHoja(HOJA_CORREO_CONTACTOS).getRange(fila._fila, 5).setValue(nuevoEstado);
  auditoria_registrar(sesion.usuario, 'Contacto de correo ' + fila.NOMBRE + ' marcado como ' + nuevoEstado, '', sesion.area);
  return { ok: true };
}

/**
 * Envía un correo de prueba a UN contacto, para comprobar que llega. Funciona aunque el envío general
 * esté desactivado (justamente sirve para probar antes de activarlo).
 */
function correo_enviarPrueba(token, idContacto) {
  var sesion = login_exigirArea_(token, 'Informática');
  var contacto = correo_leerContactos_().filter(function (c) { return Number(c.ID_CONTACTO) === Number(idContacto); })[0];
  if (!contacto) throw new Error('No se encontró el contacto ' + idContacto);

  var cfg = correo_leerConfig_();
  try {
    if (MailApp.getRemainingDailyQuota() < 1) throw new Error('Hoy ya no quedan correos disponibles en la cuenta que envía.');
    MailApp.sendEmail(correo_opciones_(cfg, contacto.CORREO,
      '[PRUEBA] Configuración de correo — Alta de Usuario',
      'Este es un correo de prueba enviado desde Gestión de Alta de Usuario (MVCLM).\n\n' +
      'Si lo estás leyendo, el envío de correos funciona para esta dirección.\n\n—\nMensaje automático.'));
  } catch (e) {
    throw new Error('No se pudo enviar la prueba: ' + e.message);
  }
  auditoria_registrar(sesion.usuario, 'Correo de prueba enviado a ' + contacto.NOMBRE, '', sesion.area);
  return { ok: true, destino: contacto.CORREO };
}

// ---------------------------------------------------------------------
// Envío de avisos (lo llaman Solicitudes.gs y Accesos.gs)
// ---------------------------------------------------------------------

function correo_opciones_(cfg, destino, asunto, cuerpo) {
  var opciones = { to: destino, subject: asunto, body: cuerpo };
  if (cfg.remitente) opciones.name = cfg.remitente;
  if (cfg.responderA) opciones.replyTo = cfg.responderA;
  return opciones;
}

/** Mensaje de cada aviso. Deliberadamente SIN nombres ni DNI: los datos personales se consultan en el sistema. */
function correo_armarMensaje_(evento, d) {
  var url = '';
  try { url = ScriptApp.getService().getUrl(); } catch (e) { url = ''; }
  var enlace = url ? ':\n' + url : '.';
  var pie = '\n\n—\nMensaje automático de Gestión de Alta de Usuario (MVCLM). No incluye datos personales: consúltalos en el sistema.';

  if (evento === 'NUEVA_SOLICITUD') {
    return {
      asunto: '[Alta de Usuario] Nueva solicitud — ' + d.resolucion,
      cuerpo: 'Se registró una nueva solicitud de acceso.\n\nResolución: ' + d.resolucion + '\nSede: ' + d.sede +
        '\nAccesos requeridos: ' + d.accesos + '\n\nIngresa al sistema para gestionarla' + enlace + pie
    };
  }
  if (evento === 'SOLICITUD_COMPLETADA') {
    return {
      asunto: '[Alta de Usuario] Solicitud completada — ' + d.resolucion,
      cuerpo: 'Todos los accesos de la solicitud quedaron resueltos.\n\nResolución: ' + d.resolucion + '\nSede: ' + d.sede +
        '\n\nPuedes ver el detalle en el sistema' + enlace + pie
    };
  }
  throw new Error('Aviso desconocido: ' + evento);
}

/**
 * Envía el aviso a los contactos ACTIVOS que lo tienen marcado. NUNCA lanza error: un correo que no sale
 * no debe impedir que la solicitud se registre. Devuelve { enviado, destinatarios, motivo } para que la
 * pantalla pueda avisar si no salió nadie.
 */
function correo_enviarEvento_(evento, datos) {
  try {
    var cfg = correo_leerConfig_();
    if (!cfg.envioActivo) {
      return { enviado: false, destinatarios: 0, motivo: 'El envío de correos está desactivado (Configuración de correo).' };
    }
    var destinos = correo_leerContactos_().filter(function (c) {
      return c.ESTADO === 'Activo' && c.EVENTOS.indexOf(evento) !== -1;
    });
    if (!destinos.length) {
      return { enviado: false, destinatarios: 0, motivo: 'No hay contactos activos para este aviso (Configuración de correo).' };
    }
    var disponibles = MailApp.getRemainingDailyQuota();
    if (disponibles < destinos.length) {
      return { enviado: false, destinatarios: 0, motivo: 'Hoy no quedan correos suficientes en la cuenta que envía (' + disponibles + ').' };
    }

    var mensaje = correo_armarMensaje_(evento, datos);
    var enviados = 0, fallidos = 0;
    destinos.forEach(function (c) { // un correo por persona: nadie ve las direcciones de los demás
      try { MailApp.sendEmail(correo_opciones_(cfg, c.CORREO, mensaje.asunto, mensaje.cuerpo)); enviados++; }
      catch (e) { fallidos++; }
    });

    auditoria_registrar(datos.usuario || 'sistema', 'Correo de aviso: ' + evento, datos.item ? 'ITEM ' + datos.item : '', datos.area || 'Sistema',
      enviados + ' enviado(s)' + (fallidos ? ', ' + fallidos + ' con error' : ''));
    return {
      enviado: enviados > 0,
      destinatarios: enviados,
      motivo: enviados ? (fallidos ? fallidos + ' correo(s) no se pudieron enviar.' : '') : 'No se pudo enviar ningún correo.'
    };
  } catch (e) {
    return { enviado: false, destinatarios: 0, motivo: 'No se pudo enviar el correo: ' + e.message };
  }
}
