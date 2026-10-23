/**
 * Mantenimiento.gs
 * CRUD de los catálogos que alimentan el resto del sistema. Todas siguen el
 * mismo patrón: Nuevo / Editar / Activar / Inactivar — nunca se borra una fila,
 * para no romper el historial de solicitudes que ya usaron ese valor.
 */

var HOJA_SEDE = 'Mantenimiento_Sede';
var HOJA_PERFIL = 'Mantenimiento_Perfil';
var HOJA_SISTEMAS = 'Mantenimiento_Sistemas';
var HOJA_ESTADOS_PROGRESO = 'Mantenimiento_Estados_Progreso';
var HOJA_PLANTILLAS = 'Mantenimiento_Plantillas';

// ---------- SEDES ----------
function mantenimiento_listarSedes() { return util_leerComoObjetos(HOJA_SEDE); }

function mantenimiento_crearSede(token, nombreSede) {
  var sesion = login_exigirArea_(token, 'Informática');
  nombreSede = util_mayus(nombreSede);
  util_obtenerHoja(HOJA_SEDE).appendRow([nombreSede, 'Activo']);
  auditoria_registrar(sesion.usuario, 'Sede creada: ' + nombreSede, nombreSede, sesion.area);
  return { ok: true };
}

function mantenimiento_cambiarEstadoSede(token, nombreSede, nuevoEstado) {
  var sesion = login_exigirArea_(token, 'Informática');
  var hoja = util_obtenerHoja(HOJA_SEDE);
  var fila = util_leerComoObjetos(HOJA_SEDE).filter(function (s) { return s.SEDE === nombreSede; })[0];
  if (!fila) throw new Error('No se encontró la sede ' + nombreSede);
  hoja.getRange(fila._fila, 2).setValue(nuevoEstado);
  auditoria_registrar(sesion.usuario, 'Sede ' + nombreSede + ' marcada como ' + nuevoEstado, nombreSede, sesion.area);
  return { ok: true };
}

// ---------- PERFILES ----------
function mantenimiento_listarPerfiles() { return util_leerComoObjetos(HOJA_PERFIL); }

function mantenimiento_crearPerfil(token, nombrePerfil) {
  var sesion = login_exigirArea_(token, 'Informática');
  nombrePerfil = util_mayus(nombrePerfil);
  util_obtenerHoja(HOJA_PERFIL).appendRow([nombrePerfil, 'Activo']);
  auditoria_registrar(sesion.usuario, 'Perfil creado: ' + nombrePerfil, nombrePerfil, sesion.area);
  return { ok: true };
}

function mantenimiento_cambiarEstadoPerfil(token, nombrePerfil, nuevoEstado) {
  var sesion = login_exigirArea_(token, 'Informática');
  var hoja = util_obtenerHoja(HOJA_PERFIL);
  var fila = util_leerComoObjetos(HOJA_PERFIL).filter(function (p) { return p.PERFIL === nombrePerfil; })[0];
  if (!fila) throw new Error('No se encontró el perfil ' + nombrePerfil);
  hoja.getRange(fila._fila, 2).setValue(nuevoEstado);
  auditoria_registrar(sesion.usuario, 'Perfil ' + nombrePerfil + ' marcado como ' + nuevoEstado, nombrePerfil, sesion.area);
  return { ok: true };
}

// ---------- SISTEMAS ----------
// Agregar una fila aquí hace aparecer un acceso nuevo en "¿Qué accesos
// requiere?" de Administración, sin tocar código.
function mantenimiento_listarSistemas() {
  return util_leerComoObjetos(HOJA_SISTEMAS).filter(function (s) { return s.ESTADO === 'Activo'; });
}

function mantenimiento_listarSistemasTodos(token) {
  login_exigirArea_(token, 'Informática');
  return util_leerComoObjetos(HOJA_SISTEMAS);
}

function mantenimiento_crearSistema(token, nombreSistema, requiereValor) {
  var sesion = login_exigirArea_(token, 'Informática');
  var id = util_siguienteId(HOJA_SISTEMAS);
  var orden = util_leerComoObjetos(HOJA_SISTEMAS).length + 1;
  nombreSistema = util_mayus(nombreSistema);
  util_obtenerHoja(HOJA_SISTEMAS).appendRow([id, nombreSistema, requiereValor ? 'SI' : 'NO', orden, 'Activo']);
  auditoria_registrar(sesion.usuario, 'Sistema creado: ' + nombreSistema, nombreSistema, sesion.area);
  return { ok: true, id: id };
}

function mantenimiento_cambiarEstadoSistema(token, idSistema, nuevoEstado) {
  var sesion = login_exigirArea_(token, 'Informática');
  var hoja = util_obtenerHoja(HOJA_SISTEMAS);
  var fila = util_leerComoObjetos(HOJA_SISTEMAS).filter(function (s) { return s.ID_SISTEMA === idSistema; })[0];
  if (!fila) throw new Error('No se encontró el sistema ' + idSistema);
  hoja.getRange(fila._fila, 5).setValue(nuevoEstado); // ESTADO
  auditoria_registrar(sesion.usuario, 'Sistema ' + fila.NOMBRE_SISTEMA + ' marcado como ' + nuevoEstado, fila.NOMBRE_SISTEMA, sesion.area);
  return { ok: true };
}

// ---------- ETAPAS DE PROGRESO ----------
// El % de cada etapa vive aquí, no en código — cambiarlo recalcula el avance
// de TODAS las solicitudes automáticamente (ver solicitudes_recalcularAvance_).
function mantenimiento_listarEstadosProgreso() { return util_leerComoObjetos(HOJA_ESTADOS_PROGRESO); }

function mantenimiento_editarPorcentajeEstado(token, idEstado, nuevoPorcentaje) {
  var sesion = login_exigirArea_(token, 'Informática');
  var hoja = util_obtenerHoja(HOJA_ESTADOS_PROGRESO);
  var fila = util_leerComoObjetos(HOJA_ESTADOS_PROGRESO).filter(function (e) { return e.ID_ESTADO === idEstado; })[0];
  if (!fila) throw new Error('No se encontró la etapa ' + idEstado);
  hoja.getRange(fila._fila, 3).setValue(nuevoPorcentaje); // PORCENTAJE
  auditoria_registrar(sesion.usuario, 'Etapa "' + fila.ESTADO + '" cambiada a ' + nuevoPorcentaje + '%', fila.ESTADO, sesion.area);
  return { ok: true };
}

// ---------- PLANTILLAS ----------
function mantenimiento_listarPlantillas() { return util_leerComoObjetos(HOJA_PLANTILLAS); }

function mantenimiento_crearPlantilla(token, datos) {
  // datos: { nombrePlantilla, sistemaAsociado, idDocDrivePlantilla, camposACombinar }
  var sesion = login_exigirArea_(token, 'Informática');
  var id = util_siguienteId(HOJA_PLANTILLAS);
  util_obtenerHoja(HOJA_PLANTILLAS).appendRow([
    id, util_mayus(datos.nombrePlantilla), util_mayus(datos.sistemaAsociado), datos.idDocDrivePlantilla, util_mayus(datos.camposACombinar), 'Activo'
  ]);
  auditoria_registrar(sesion.usuario, 'Plantilla creada: ' + datos.nombrePlantilla, datos.nombrePlantilla, sesion.area);
  return { ok: true, id: id };
}

function mantenimiento_cambiarEstadoPlantilla(token, idPlantilla, nuevoEstado) {
  var sesion = login_exigirArea_(token, 'Informática');
  var hoja = util_obtenerHoja(HOJA_PLANTILLAS);
  var fila = util_leerComoObjetos(HOJA_PLANTILLAS).filter(function (p) { return p.ID_PLANTILLA === idPlantilla; })[0];
  if (!fila) throw new Error('No se encontró la plantilla ' + idPlantilla);
  hoja.getRange(fila._fila, 6).setValue(nuevoEstado); // ESTADO
  auditoria_registrar(sesion.usuario, 'Plantilla ' + fila.NOMBRE_PLANTILLA + ' marcada como ' + nuevoEstado, fila.NOMBRE_PLANTILLA, sesion.area);
  return { ok: true };
}

// ---------- EDICIÓN (agregado: las funciones de crear y cambiar estado de arriba no se modificaron) ----------

/**
 * Reemplaza un valor en una columna (por nombre de encabezado) de otra hoja.
 * Se usa al renombrar una sede/perfil/sistema para que el historial no quede
 * apuntando al nombre viejo. Devuelve cuántas celdas cambió.
 */
function mantenimiento_reemplazarEnColumna_(nombreHoja, encabezado, valorViejo, valorNuevo) {
  var hoja = util_obtenerHoja(nombreHoja);
  var ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return 0;
  var encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  var col = encabezados.indexOf(encabezado) + 1;
  if (col < 1) return 0;
  var rango = hoja.getRange(2, col, ultimaFila - 1, 1);
  var valores = rango.getValues();
  var cambios = 0;
  for (var i = 0; i < valores.length; i++) {
    if (String(valores[i][0]) === String(valorViejo)) { valores[i][0] = valorNuevo; cambios++; }
  }
  if (cambios) rango.setValues(valores);
  return cambios;
}

/** true si ya hay OTRA fila (distinta de filaExcluida) con ese texto en esa columna. */
function mantenimiento_nombreRepetido_(filas, campo, texto, filaExcluida) {
  return filas.some(function (f) {
    return f._fila !== filaExcluida && String(f[campo]).toLowerCase() === String(texto).toLowerCase();
  });
}

function mantenimiento_editarSede(token, nombreActual, nombreNuevo) {
  var sesion = login_exigirArea_(token, 'Informática');
  nombreNuevo = util_mayus(nombreNuevo);
  if (!nombreNuevo) throw new Error('El nombre de la sede no puede quedar vacío.');
  var filas = util_leerComoObjetos(HOJA_SEDE);
  var fila = filas.filter(function (s) { return s.SEDE === nombreActual; })[0];
  if (!fila) throw new Error('No se encontró la sede ' + nombreActual);
  if (mantenimiento_nombreRepetido_(filas, 'SEDE', nombreNuevo, fila._fila)) {
    throw new Error('Ya existe una sede llamada "' + nombreNuevo + '".');
  }
  util_obtenerHoja(HOJA_SEDE).getRange(fila._fila, 1).setValue(nombreNuevo);
  var actualizadas = nombreActual !== nombreNuevo
    ? mantenimiento_reemplazarEnColumna_(HOJA_SOLICITUDES, 'SEDE', nombreActual, nombreNuevo) : 0;
  auditoria_registrar(sesion.usuario, 'Sede editada: "' + nombreActual + '" → "' + nombreNuevo + '"', nombreNuevo, sesion.area, actualizadas + ' solicitudes actualizadas');
  return { ok: true };
}

function mantenimiento_editarPerfil(token, nombreActual, nombreNuevo) {
  var sesion = login_exigirArea_(token, 'Informática');
  nombreNuevo = util_mayus(nombreNuevo);
  if (!nombreNuevo) throw new Error('El nombre del perfil no puede quedar vacío.');
  var filas = util_leerComoObjetos(HOJA_PERFIL);
  var fila = filas.filter(function (p) { return p.PERFIL === nombreActual; })[0];
  if (!fila) throw new Error('No se encontró el perfil ' + nombreActual);
  if (mantenimiento_nombreRepetido_(filas, 'PERFIL', nombreNuevo, fila._fila)) {
    throw new Error('Ya existe un perfil llamado "' + nombreNuevo + '".');
  }
  util_obtenerHoja(HOJA_PERFIL).getRange(fila._fila, 1).setValue(nombreNuevo);
  var actualizadas = nombreActual !== nombreNuevo
    ? mantenimiento_reemplazarEnColumna_(HOJA_SOLICITUDES, 'PERFIL', nombreActual, nombreNuevo) : 0;
  auditoria_registrar(sesion.usuario, 'Perfil editado: "' + nombreActual + '" → "' + nombreNuevo + '"', nombreNuevo, sesion.area, actualizadas + ' solicitudes actualizadas');
  return { ok: true };
}

function mantenimiento_editarSistema(token, idSistema, nombreNuevo, requiereValor) {
  var sesion = login_exigirArea_(token, 'Informática');
  nombreNuevo = util_mayus(nombreNuevo);
  if (!nombreNuevo) throw new Error('El nombre del sistema no puede quedar vacío.');
  var filas = util_leerComoObjetos(HOJA_SISTEMAS);
  var fila = filas.filter(function (s) { return Number(s.ID_SISTEMA) === Number(idSistema); })[0];
  if (!fila) throw new Error('No se encontró el sistema ' + idSistema);
  if (mantenimiento_nombreRepetido_(filas, 'NOMBRE_SISTEMA', nombreNuevo, fila._fila)) {
    throw new Error('Ya existe un sistema llamado "' + nombreNuevo + '".');
  }
  var nombreViejo = fila.NOMBRE_SISTEMA;
  var hoja = util_obtenerHoja(HOJA_SISTEMAS);
  hoja.getRange(fila._fila, 2).setValue(nombreNuevo);            // NOMBRE_SISTEMA
  hoja.getRange(fila._fila, 3).setValue(requiereValor ? 'SI' : 'NO'); // REQUIERE_VALOR

  var actualizadas = 0;
  if (nombreViejo !== nombreNuevo) {
    actualizadas += mantenimiento_reemplazarEnColumna_(HOJA_DETALLE_ACCESOS, 'SISTEMA', nombreViejo, nombreNuevo);
    mantenimiento_reemplazarEnColumna_(HOJA_PLANTILLAS, 'SISTEMA_ASOCIADO', nombreViejo, nombreNuevo);
  }
  auditoria_registrar(sesion.usuario, 'Sistema editado: "' + nombreViejo + '" → "' + nombreNuevo + '"', nombreNuevo, sesion.area, actualizadas + ' accesos actualizados');
  return { ok: true };
}

function mantenimiento_editarPlantilla(token, idPlantilla, datos) {
  // datos: { nombrePlantilla, sistemaAsociado, idDocDrivePlantilla, camposACombinar }
  var sesion = login_exigirArea_(token, 'Informática');
  if (!String(datos.nombrePlantilla || '').trim()) throw new Error('El nombre de la plantilla no puede quedar vacío.');
  var fila = util_leerComoObjetos(HOJA_PLANTILLAS)
    .filter(function (p) { return Number(p.ID_PLANTILLA) === Number(idPlantilla); })[0];
  if (!fila) throw new Error('No se encontró la plantilla ' + idPlantilla);
  util_obtenerHoja(HOJA_PLANTILLAS).getRange(fila._fila, 2, 1, 4).setValues([[
    util_mayus(datos.nombrePlantilla), util_mayus(datos.sistemaAsociado), datos.idDocDrivePlantilla, util_mayus(datos.camposACombinar)
  ]]);
  auditoria_registrar(sesion.usuario, 'Plantilla editada: ' + datos.nombrePlantilla, datos.nombrePlantilla, sesion.area);
  return { ok: true };
}
