/**
 * Usuarios.gs
 * Personal con acceso a ESTE sistema (Administración e Informática).
 * No confundir con accesos_generarCredencialesEmpleado (Accesos.gs), que es
 * para el empleado que RECIBE los accesos, no para quien usa esta app.
 */

/** Lista el personal con acceso al sistema, sin exponer el hash de la contraseña. */
function usuarios_listar(token) {
  login_exigirArea_(token, 'Informática');
  return util_leerComoObjetos(HOJA_CREDENCIALES).map(function (u) {
    return {
      ID_USUARIO: u.ID_USUARIO, NOMBRES: u.NOMBRES, APELLIDOS: u.APELLIDOS,
      DNI: u.DNI, CARGO: u.CARGO, AREA: u.AREA, USUARIO: u.USUARIO, ESTADO: u.ESTADO
      // Deliberadamente NO se incluye SALT ni HASH_CONTRASENA en la respuesta al cliente.
    };
  });
}

/**
 * Crea un usuario del sistema. AREA decide qué menú verá al iniciar sesión.
 * @param {Object} datos { nombres, apellidos, dni, cargo, area, usuario }
 */
function usuarios_crear(token, datos) {
  var sesion = login_exigirArea_(token, 'Informática');
  if (!util_soloLetras(datos.nombres) || !util_soloLetras(datos.apellidos)) {
    throw new Error('Nombres y apellidos no pueden contener números.');
  }
  if (['Administración', 'Informática'].indexOf(datos.area) === -1) {
    throw new Error('AREA debe ser "Administración" o "Informática".');
  }

  var claveTemporal = util_generarClaveTemporal();
  var salt = util_generarSalt();
  var hash = util_hashClave(claveTemporal, salt);

  var hoja = util_obtenerHoja(HOJA_CREDENCIALES);
  var id = util_siguienteId(HOJA_CREDENCIALES);
  hoja.appendRow([
    id, datos.nombres, datos.apellidos, util_normalizarDni(datos.dni), datos.cargo,
    datos.area, datos.usuario, salt, hash, util_fechaHoy(), 'Activo'
  ]);

  auditoria_registrar(sesion.usuario, 'Usuario del sistema creado (' + datos.usuario + ', área: ' + datos.area + ')', datos.usuario, sesion.area);
  return { ok: true, usuario: datos.usuario, claveTemporal: claveTemporal }; // se muestra una sola vez
}

/** Restablece la clave de un usuario del sistema. Nunca reutiliza ni expone la anterior. */
function usuarios_restablecerClave(token, idUsuario) {
  var sesion = login_validarToken_(token); // cualquier usuario autenticado puede restablecer la suya;
  // TODO: si se quiere permitir restablecer la clave de OTRO usuario, exigir aquí login_exigirArea_(token, 'Informática').

  var hoja = util_obtenerHoja(HOJA_CREDENCIALES);
  var fila = util_leerComoObjetos(HOJA_CREDENCIALES).filter(function (u) { return u.ID_USUARIO === idUsuario; })[0];
  if (!fila) throw new Error('No se encontró el usuario ' + idUsuario);

  var claveTemporal = util_generarClaveTemporal();
  var salt = util_generarSalt();
  var hash = util_hashClave(claveTemporal, salt);

  hoja.getRange(fila._fila, 8).setValue(salt);              // SALT
  hoja.getRange(fila._fila, 9).setValue(hash);              // HASH_CONTRASENA
  hoja.getRange(fila._fila, 10).setValue(util_fechaHoy());  // FECHA_ULTIMO_CAMBIO

  auditoria_registrar(sesion.usuario, 'Clave restablecida para ' + fila.USUARIO, fila.USUARIO, sesion.area, 'Contraseña no almacenada');
  return { usuario: fila.USUARIO, claveTemporal: claveTemporal }; // se muestra una sola vez
}

/** Activa/inactiva un usuario del sistema (no se borra nunca la fila). */
function usuarios_cambiarEstado(token, idUsuario, nuevoEstado) {
  var sesion = login_exigirArea_(token, 'Informática');
  var hoja = util_obtenerHoja(HOJA_CREDENCIALES);
  var fila = util_leerComoObjetos(HOJA_CREDENCIALES).filter(function (u) { return u.ID_USUARIO === idUsuario; })[0];
  if (!fila) throw new Error('No se encontró el usuario ' + idUsuario);

  hoja.getRange(fila._fila, 11).setValue(nuevoEstado); // ESTADO
  auditoria_registrar(sesion.usuario, 'Usuario ' + fila.USUARIO + ' marcado como ' + nuevoEstado, fila.USUARIO, sesion.area);
  return { ok: true };
}
