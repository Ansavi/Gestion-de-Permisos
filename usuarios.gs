/**
 * Usuarios.gs
 * Personal con acceso a ESTE sistema (Administración e Informática).
 * No confundir con accesos_generarCredencialesEmpleado (Accesos.gs), que es
 * para el empleado que RECIBE los accesos, no para quien usa esta app.
 */

/** DNI como texto de 8 dígitos (recupera los ceros a la izquierda si la celda quedó como número). */
function usuarios_dniTexto_(valor) {
  var digitos = String(valor == null ? '' : valor).replace(/\D/g, '');
  return digitos ? ('00000000' + digitos).slice(-8) : '';
}

/** Lista el personal con acceso al sistema, sin exponer el hash de la contraseña. */
function usuarios_listar(token) {
  login_exigirArea_(token, 'Informática');
  return util_leerComoObjetos(HOJA_CREDENCIALES).map(function (u) {
    return {
      ID_USUARIO: u.ID_USUARIO, NOMBRES: u.NOMBRES, APELLIDOS: u.APELLIDOS,
      DNI: usuarios_dniTexto_(u.DNI), CARGO: u.CARGO, AREA: u.AREA, USUARIO: u.USUARIO, ESTADO: u.ESTADO
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
  var dni = util_normalizarDni(datos.dni);
  var usuario = String(datos.usuario || '').trim().toLowerCase();
  if (dni.length !== 8) throw new Error('El DNI debe tener 8 dígitos.');
  if (!util_mayus(datos.cargo)) throw new Error('El cargo es obligatorio.');
  if (!/^[a-z0-9._-]{3,30}$/.test(usuario)) {
    throw new Error('El usuario debe tener de 3 a 30 caracteres: letras minúsculas, números, punto, guion o guion bajo.');
  }
  var existentes = util_leerComoObjetos(HOJA_CREDENCIALES);
  if (existentes.some(function (u) { return String(u.USUARIO).toLowerCase() === usuario; })) {
    throw new Error('Ya existe el usuario "' + usuario + '".');
  }
  if (existentes.some(function (u) { return usuarios_dniTexto_(u.DNI) === dni; })) {
    throw new Error('Ya existe un usuario con ese DNI.');
  }

  var claveTemporal = util_generarClaveTemporal();
  var salt = util_generarSalt();
  var hash = util_hashClave(claveTemporal, salt);

  var hoja = util_obtenerHoja(HOJA_CREDENCIALES);
  var id = util_siguienteId(HOJA_CREDENCIALES);
  hoja.appendRow([
    id, util_mayus(datos.nombres), util_mayus(datos.apellidos), util_normalizarDni(datos.dni), util_mayus(datos.cargo),
    datos.area, usuario, salt, hash, util_fechaHoy(), 'Activo'
  ]);
  hoja.getRange(hoja.getLastRow(), 4).setNumberFormat('@').setValue(dni); // DNI como texto: conserva los ceros

  auditoria_registrar(sesion.usuario, 'Usuario del sistema creado (' + usuario + ', área: ' + datos.area + ')', usuario, sesion.area);
  return { ok: true, usuario: usuario, claveTemporal: claveTemporal }; // se muestra una sola vez
}

/** Restablece la clave de un usuario del sistema. Nunca reutiliza ni expone la anterior. */
function usuarios_restablecerClave(token, idUsuario) {
  var sesion = login_exigirArea_(token, 'Informática'); // solo Informática puede restablecer claves

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
  if (['Activo', 'Bloqueado'].indexOf(nuevoEstado) === -1) throw new Error('Estado no válido: ' + nuevoEstado);
  if (nuevoEstado === 'Bloqueado' && String(fila.USUARIO).toLowerCase() === String(sesion.usuario).toLowerCase()) {
    throw new Error('No puedes bloquear tu propio usuario.');
  }

  hoja.getRange(fila._fila, 11).setValue(nuevoEstado); // ESTADO
  auditoria_registrar(sesion.usuario, 'Usuario ' + fila.USUARIO + ' marcado como ' + nuevoEstado, fila.USUARIO, sesion.area);
  return { ok: true };
}

/**
 * Edita los datos de una persona (el usuario de acceso no se cambia: es su identidad).
 * @param {Object} datos { nombres, apellidos, dni, cargo, area }
 */
function usuarios_editar(token, idUsuario, datos) {
  var sesion = login_exigirArea_(token, 'Informática');
  var filas = util_leerComoObjetos(HOJA_CREDENCIALES);
  var fila = filas.filter(function (u) { return Number(u.ID_USUARIO) === Number(idUsuario); })[0];
  if (!fila) throw new Error('No se encontró el usuario ' + idUsuario);

  if (!util_soloLetras(datos.nombres) || !util_soloLetras(datos.apellidos)) {
    throw new Error('Nombres y apellidos no pueden contener números.');
  }
  if (['Administración', 'Informática'].indexOf(datos.area) === -1) {
    throw new Error('AREA debe ser "Administración" o "Informática".');
  }
  var dni = util_normalizarDni(datos.dni);
  if (dni.length !== 8) throw new Error('El DNI debe tener 8 dígitos.');
  if (!util_mayus(datos.cargo)) throw new Error('El cargo es obligatorio.');
  if (filas.some(function (u) { return u._fila !== fila._fila && usuarios_dniTexto_(u.DNI) === dni; })) {
    throw new Error('Ya existe otro usuario con ese DNI.');
  }
  var esMismoUsuario = String(fila.USUARIO).toLowerCase() === String(sesion.usuario).toLowerCase();
  if (esMismoUsuario && datos.area !== fila.AREA) {
    throw new Error('No puedes cambiar tu propia área.');
  }

  var hoja = util_obtenerHoja(HOJA_CREDENCIALES);
  hoja.getRange(fila._fila, 4).setNumberFormat('@'); // DNI como texto antes de escribir
  hoja.getRange(fila._fila, 2, 1, 5).setValues([[
    util_mayus(datos.nombres), util_mayus(datos.apellidos), dni, util_mayus(datos.cargo), datos.area
  ]]);

  auditoria_registrar(sesion.usuario, 'Usuario del sistema editado: ' + fila.USUARIO, fila.USUARIO, sesion.area);
  return { ok: true };
}
