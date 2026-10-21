/**
 * Autenticacion.gs
 * (Antes "Login.gs" — se renombró el archivo porque Apps Script no permite
 * un .gs y un .html con el mismo nombre "login" en el mismo proyecto.
 * Las funciones siguen con el prefijo login_ tal como se definió.)
 * Valida usuario/contraseña contra Credenciales y crea una sesión por token
 * (CacheService), porque un web app de Apps Script no tiene sesión de
 * navegador propia cuando el login es nuestro y no el de Google.
 *
 * Cada función de las demás páginas recibe este token como primer parámetro
 * y lo valida con login_validarToken_ antes de hacer cualquier cosa.
 */

var HOJA_CREDENCIALES = 'Credenciales';
var DURACION_SESION_SEGUNDOS = 60 * 60 * 4; // 4 horas

/**
 * Valida usuario y contraseña. Si son correctos, crea una sesión y devuelve
 * el token + los datos que la interfaz necesita para armar el menú.
 */
function login_validar(usuario, claveTexto) {
  var credenciales = util_leerComoObjetos(HOJA_CREDENCIALES);
  var fila = credenciales.filter(function (c) {
    return String(c.USUARIO).toLowerCase() === String(usuario).toLowerCase();
  })[0];

  if (!fila || fila.ESTADO !== 'Activo') {
    return { ok: false, mensaje: 'Usuario o contraseña incorrectos.' };
  }

  var claveOk = util_verificarClave(claveTexto, fila.SALT, fila.HASH_CONTRASENA);
  if (!claveOk) {
    return { ok: false, mensaje: 'Usuario o contraseña incorrectos.' };
  }

  var token = Utilities.getUuid();
  var sesion = {
    usuario: fila.USUARIO,
    nombre: fila.NOMBRES + ' ' + fila.APELLIDOS,
    cargo: fila.CARGO,
    area: fila.AREA // "Administración" | "Informática" — decide el menú
  };
  CacheService.getScriptCache().put('sesion_' + token, JSON.stringify(sesion), DURACION_SESION_SEGUNDOS);

  return { ok: true, token: token, sesion: sesion };
}

/** Cierra la sesión (botón "Cerrar sesión"). */
function login_cerrarSesion(token) {
  CacheService.getScriptCache().remove('sesion_' + token);
  return { ok: true };
}

/**
 * Igual que login_validarToken_, pero sin lanzar error — para que el cliente
 * (sistema.html al cargar) pueda preguntar "¿tengo sesión válida?" y decidir
 * qué mostrar, sin necesitar manejar una excepción.
 */
function login_obtenerSesion(token) {
  var crudo = CacheService.getScriptCache().get('sesion_' + token);
  if (!crudo) return { ok: false };
  return { ok: true, sesion: JSON.parse(crudo) };
}

/**
 * Valida el token y devuelve los datos de sesión, o lanza error si no es válido.
 * TODA función de Solicitudes/Accesos/Usuarios/Mantenimiento/Reportes debe
 * empezar llamando a esto con el token que le llega del cliente.
 */
function login_validarToken_(token) {
  var crudo = CacheService.getScriptCache().get('sesion_' + token);
  if (!crudo) throw new Error('Sesión inválida o expirada. Vuelve a iniciar sesión.');
  return JSON.parse(crudo);
}

/** Exige que la sesión pertenezca a un área específica (control de permisos por pestaña). */
function login_exigirArea_(token, areaRequerida) {
  var sesion = login_validarToken_(token);
  if (sesion.area !== areaRequerida) {
    throw new Error('No tienes permiso para esta acción (requiere área: ' + areaRequerida + ').');
  }
  return sesion;
}
