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
    id, datos.nombrePlantilla, datos.sistemaAsociado, datos.idDocDrivePlantilla, datos.camposACombinar, 'Activo'
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
