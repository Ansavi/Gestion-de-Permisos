/**
 * Utilidades.gs
 * Funciones transversales usadas por todas las páginas.
 * No llevan prefijo de página porque no pertenecen a una sola pantalla.
 * Convención del proyecto: <pagina>_accion() para lo propio de cada pantalla,
 * auditoria_ para todo lo que escribe en el log, util_ para lo demás.
 */

// ID del Spreadsheet del sistema (hojas Solicitudes, Credenciales, Mantenimiento_*, etc.).
// Se usa openById en vez de getActiveSpreadsheet() porque este proyecto de Apps Script
// puede no estar "vinculado" (contenedor) directamente a ese Sheet — openById funciona
// siempre, esté el script vinculado o sea un proyecto independiente.
var SPREADSHEET_ID = '1fdTykVYeatROh0I4cUD-pC4U2g2dVTUIpQUfpgY7j3Y';
var SS = SpreadsheetApp.openById(SPREADSHEET_ID);

/** Devuelve la hoja por nombre, o lanza error claro si no existe (evita fallos silenciosos). */
function util_obtenerHoja(nombre) {
  var hoja = SS.getSheetByName(nombre);
  if (!hoja) throw new Error('No existe la hoja "' + nombre + '". Revisa el nombre en Mantenimiento o en el Spreadsheet.');
  return hoja;
}

/**
 * Lee una hoja completa y la devuelve como arreglo de objetos {ENCABEZADO: valor},
 * usando la fila 1 como encabezados. Evita tener que acordarse de índices de columna.
 */
function util_leerComoObjetos(nombreHoja) {
  var hoja = util_obtenerHoja(nombreHoja);
  var datos = hoja.getDataRange().getValues();
  var encabezados = datos[0];
  var filas = [];
  for (var i = 1; i < datos.length; i++) {
    var obj = { _fila: i + 1 }; // fila real en la hoja, útil para editar después
    for (var c = 0; c < encabezados.length; c++) {
      obj[encabezados[c]] = datos[i][c];
    }
    filas.push(obj);
  }
  return filas;
}

/** Devuelve el siguiente ID autonumérico de una hoja, según su primera columna (ID). */
function util_siguienteId(nombreHoja) {
  var hoja = util_obtenerHoja(nombreHoja);
  var ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return 1;
  var valores = hoja.getRange(2, 1, ultimaFila - 1, 1).getValues();
  var max = 0;
  valores.forEach(function (fila) {
    var n = Number(fila[0]);
    if (!isNaN(n) && n > max) max = n;
  });
  return max + 1;
}

/** Genera una clave temporal aleatoria. Nunca se guarda en ninguna hoja — solo se muestra una vez. */
function util_generarClaveTemporal() {
  var chars = 'ABCDEFGHJKMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789!@#$';
  var clave = '';
  for (var i = 0; i < 10; i++) {
    clave += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return clave;
}

/** Genera una sal aleatoria para el hash de contraseñas. */
function util_generarSalt() {
  return Utilities.getUuid();
}

/**
 * Hash SHA-256 de "clave+salt". Esto es lo único que se guarda en Credenciales —
 * NUNCA la contraseña en texto plano, ni siquiera temporalmente.
 */
function util_hashClave(claveTexto, salt) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, claveTexto + salt);
  return bytes.map(function (b) { return (b < 0 ? b + 256 : b).toString(16).padStart(2, '0'); }).join('');
}

/** Verifica una contraseña ingresada contra el hash + salt guardados. */
function util_verificarClave(claveTexto, salt, hashGuardado) {
  return util_hashClave(claveTexto, salt) === hashGuardado;
}

/** Normaliza un DNI: solo dígitos, máximo 8, se guarda como texto para no perder ceros a la izquierda. */
function util_normalizarDni(dni) {
  return String(dni).replace(/\D/g, '').slice(0, 8);
}

/** true si el texto no contiene números (para Nombres/Apellidos). */
function util_soloLetras(texto) {
  return /^[a-zA-ZÀ-ÿñÑ\s]+$/.test(texto);
}

function util_fechaHoy() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function util_fechaHoraAhora() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}

/**
 * Texto en MAYÚSCULAS y sin espacios sobrantes. Se usa en todo texto que escribe
 * una persona (resolución, nombres, apellidos, cargo, catálogos de Mantenimiento).
 * NO usar en contraseñas, usuario de red/correo ni IDs o URLs de Drive (distinguen mayúsculas).
 */
function util_mayus(texto) {
  return String(texto == null ? '' : texto).trim().toUpperCase();
}
