/**
 * Script.gs
 * Configuración inicial del proyecto — se ejecuta UNA SOLA VEZ, a mano,
 * desde el editor de Apps Script (seleccionar la función y "Ejecutar"),
 * para dejar todo listo antes de usar el sistema por primera vez:
 *   1) Crea las HOJAS y CAMPOS (encabezados) del alta de usuario y del resto
 *      del sistema, con los mismos nombres que ya usa el código (Solicitudes.gs,
 *      Accesos.gs, Usuarios.gs, Mantenimiento.gs) — si los nombres no coinciden
 *      exactos, esas funciones fallan al leer/escribir.
 *   2) Precarga los catálogos de Mantenimiento ya acordados (Sedes, Perfiles,
 *      Sistemas, Etapas de progreso) para no arrancar con todo vacío.
 *   3) Crea las carpetas de Drive del sistema.
 *   4) Crea el usuario administrador inicial en Credenciales.
 *
 * Es idempotente: se puede volver a ejecutar sin duplicar hojas, catálogos
 * ni el usuario admin.
 *
 * No se llama desde ninguna pantalla — es un script de arranque, no una
 * función de la app.
 */

/**
 * DIAGNÓSTICO — ejecutar esta función sola (seleccionarla en el desplegable
 * y darle Ejecutar) para saber si el problema de login es del BACKEND
 * (esta función) o del DESPLIEGUE/FRONTEND (la web app).
 * Revisa el resultado en Ver → Registros de ejecución.
 */
function diag_probarLoginAdmin() {
  var filas = util_leerComoObjetos(HOJA_CREDENCIALES);
  var admin = filas.filter(function (f) { return String(f.USUARIO).toLowerCase() === 'admin'; })[0];

  if (!admin) {
    Logger.log('❌ NO existe ninguna fila con USUARIO = "admin" en Credenciales. ' +
      'Corre config_inicializarSistema (o config_crearUsuarioAdminInicial) primero.');
    return;
  }

  Logger.log('Fila encontrada: ' + JSON.stringify(admin));
  if (admin.ESTADO !== 'Activo') {
    Logger.log('❌ El usuario existe pero ESTADO = "' + admin.ESTADO + '" (debe ser exactamente "Activo", con mayúscula).');
  }
  if (!admin.SALT || !admin.HASH_CONTRASENA) {
    Logger.log('❌ Falta SALT o HASH_CONTRASENA — revisa que el ENCABEZADO de esa columna diga exactamente ' +
      '"HASH_CONTRASENA" (SIN Ñ). Si dice "HASH_CONTRASEÑA" (con Ñ), el código nunca la encuentra y el login falla siempre.');
  }

  var resultado = login_validar('admin', 'admin123');
  Logger.log('Resultado de login_validar("admin","admin123"): ' + JSON.stringify(resultado));
  Logger.log(resultado.ok
    ? '✅ El backend funciona. Si en el navegador sigue sin entrar, el problema es el DESPLIEGUE: hay que publicar una Nueva versión (Implementar → Administrar implementaciones → editar → Nueva versión) y entrar por la URL que termina en /exec.'
    : '❌ El backend rechaza el login — revisa los mensajes de arriba.');
}

/** Ejecuta todo lo necesario para dejar el sistema listo para usarse. */
function config_inicializarSistema() {
  var hojas = config_crearHojasYCampos();
  var carpetas = config_crearCarpetas();
  var admin = config_crearUsuarioAdminInicial();
  Logger.log('Hojas: ' + JSON.stringify(hojas));
  Logger.log('Carpetas: ' + JSON.stringify(carpetas));
  Logger.log('Usuario admin: ' + JSON.stringify(admin));
  return { hojas: hojas, carpetas: carpetas, admin: admin };
}

/**
 * Crea cada hoja del sistema con sus encabezados exactos (columna por columna,
 * en el mismo orden que usa el código en Utilidades.gs/Solicitudes.gs/etc.)
 * y precarga los catálogos de Mantenimiento. No borra ni toca una hoja que
 * ya tenga datos — solo crea la hoja y pone encabezados si está vacía.
 */
function config_crearHojasYCampos() {
  var definicion = {
    'Solicitudes': ['ITEM','RESOLUCION','FECHA_RESOLUCION','FECHA_SOLICITUD','ADJUNTO_URL',
      'SEDE','DNI','APELLIDOS','NOMBRES','PERFIL','AVANCE_%','ESTADO_GENERAL','NOTIFICADO','FECHA_NOTIFICACION'],

    'Detalle_Accesos': ['ID_DETALLE','ITEM_SOLICITUD','SISTEMA','REQUERIDO','ESTADO',
      'VALOR_ASIGNADO','FECHA_ACTUALIZACION','ACTUALIZADO_POR'],

    'Mantenimiento_Sede': ['SEDE','ESTADO'],
    'Mantenimiento_Perfil': ['PERFIL','ESTADO'],
    'Mantenimiento_Sistemas': ['ID_SISTEMA','NOMBRE_SISTEMA','REQUIERE_VALOR','ORDEN','ESTADO'],
    'Mantenimiento_Estados_Progreso': ['ID_ESTADO','ESTADO','PORCENTAJE','ORDEN'],
    'Mantenimiento_Plantillas': ['ID_PLANTILLA','NOMBRE_PLANTILLA','SISTEMA_ASOCIADO',
      'ID_DOC_DRIVE_PLANTILLA','CAMPOS_A_COMBINAR','ESTADO'],

    // Sin Ñ a propósito: es la columna que Autenticacion.gs lee como fila.HASH_CONTRASENA.
    'Credenciales': ['ID_USUARIO','NOMBRES','APELLIDOS','DNI','CARGO','AREA','USUARIO',
      'SALT','HASH_CONTRASENA','FECHA_ULTIMO_CAMBIO','ESTADO'],

    'Auditoria': ['ID_EVENTO','FECHA_HORA','USUARIO','AREA','ACCION','ITEM_AFECTADO','DETALLE']
  };

  var resultado = {};
  Object.keys(definicion).forEach(function (nombreHoja) {
    resultado[nombreHoja] = config_crearHojaSiNoExiste_(nombreHoja, definicion[nombreHoja]);
  });

  config_precargarCatalogos_();
  return resultado;
}

/** Crea la hoja si no existe y le pone encabezados en negrita; si ya existe, la deja tal cual. */
function config_crearHojaSiNoExiste_(nombre, encabezados) {
  var hoja = SS.getSheetByName(nombre); // SS apunta al Spreadsheet por ID (ver Utilidades.gs)
  if (hoja) return 'ya existía, no se tocó';

  hoja = SS.insertSheet(nombre);
  hoja.getRange(1, 1, 1, encabezados.length).setValues([encabezados]).setFontWeight('bold');
  hoja.setFrozenRows(1);
  return 'creada con ' + encabezados.length + ' columnas';
}

/** Precarga Sedes, Perfiles, Sistemas y Etapas de progreso — solo si la hoja está vacía (sin filas de datos). */
function config_precargarCatalogos_() {
  config_precargarSiVacia_('Mantenimiento_Sede', [
    ['CARABAYLLO', 'Activo'],
    ['INDEPENDENCIA', 'Activo']
  ]);

  config_precargarSiVacia_('Mantenimiento_Perfil', [
    ['ADMINISTRADOR','Activo'], ['ANALISTA','Activo'], ['ASISTENTE','Activo'],
    ['AUXILIAR','Activo'], ['COORDINADOR','Activo'], ['ESPECIALISTA','Activo'],
    ['JUEZ','Activo'], ['SECRETARIO','Activo'], ['SESIGRISTA','Activo']
  ]);

  config_precargarSiVacia_('Mantenimiento_Sistemas', [
    [1,'USUARIO DE RED','SI',1,'Activo'],
    [2,'CORREO ELECTRÓNICO CORPORATIVO','SI',2,'Activo'],
    [3,'SISTEMA DE EXPEDIENTES (SIJ)','NO',3,'Activo'],
    [4,'CERTIFICADO ELECTRÓNICO','NO',4,'Activo'],
    [5,'WHATSAPP WEB','NO',5,'Activo'],
    [6,'VPN 24HRS','NO',6,'Activo'],
    [7,'AMAUTA PRO','NO',7,'Activo'],
    [8,'RENIEC','NO',8,'Activo']
  ]);

  config_precargarSiVacia_('Mantenimiento_Estados_Progreso', [
    [1,'Registrado',10,1],
    [2,'Asignado',25,2],
    [3,'En progreso',50,3],
    [4,'En espera',75,4],
    [5,'Resuelto',100,5]
  ]);
}

/** Agrega filas de ejemplo/catálogo solo si la hoja todavía no tiene datos (fila 2 vacía). */
function config_precargarSiVacia_(nombreHoja, filas) {
  var hoja = util_obtenerHoja(nombreHoja);
  if (hoja.getLastRow() > 1) return; // ya tiene datos, no se toca
  hoja.getRange(2, 1, filas.length, filas[0].length).setValues(filas);
}

/**
 * Crea (o reutiliza si ya existen) las carpetas de Drive del sistema y guarda
 * sus IDs en las Propiedades del Script, para que el resto del código
 * (ej. accesos_generarDocumento en Accesos.gs) las use sin tener el ID
 * escrito a mano.
 */
function config_crearCarpetas() {
  var props = PropertiesService.getScriptProperties();

  var idPlantillas = config_obtenerOCrearCarpeta_('MVCLM - Plantillas');
  var idGenerados = config_obtenerOCrearCarpeta_('MVCLM - Documentos Generados (RESTRINGIDO)');

  props.setProperty('CARPETA_PLANTILLAS_ID', idPlantillas);
  props.setProperty('CARPETA_DOCUMENTOS_GENERADOS_ID', idGenerados);

  return {
    plantillas: { id: idPlantillas, url: 'https://drive.google.com/drive/folders/' + idPlantillas },
    documentosGenerados: { id: idGenerados, url: 'https://drive.google.com/drive/folders/' + idGenerados },
    aviso: 'Revisa manualmente los permisos de "Documentos Generados" — puede contener datos de menores (MVCLM). No debe ser de acceso general.'
  };
}

/** Busca una carpeta por nombre en la raíz de Drive; si no existe, la crea. */
function config_obtenerOCrearCarpeta_(nombre) {
  var existentes = DriveApp.getFoldersByName(nombre);
  if (existentes.hasNext()) return existentes.next().getId();
  return DriveApp.createFolder(nombre).getId();
}

/**
 * Crea el usuario administrador inicial para poder entrar al sistema por
 * primera vez. Es idempotente: si "admin" ya existe, no lo duplica.
 *
 * Usuario: admin
 * Clave inicial: admin123  (se guarda como hash, nunca en texto plano —
 * ver util_hashClave en Utilidades.gs)
 *
 * ⚠️ Cambiar esta clave desde "Usuarios" apenas se entre por primera vez
 * (usuarios_restablecerClave). admin123 es solo para el arranque.
 */
function config_crearUsuarioAdminInicial() {
  var hoja = util_obtenerHoja(HOJA_CREDENCIALES);
  var yaExiste = util_leerComoObjetos(HOJA_CREDENCIALES)
    .some(function (u) { return String(u.USUARIO).toLowerCase() === 'admin'; });

  if (yaExiste) {
    return { ok: true, mensaje: 'El usuario "admin" ya existía — no se creó de nuevo.' };
  }

  var claveInicial = 'admin123';
  var salt = util_generarSalt();
  var hash = util_hashClave(claveInicial, salt);
  var id = util_siguienteId(HOJA_CREDENCIALES);

  hoja.appendRow([
    id, 'ADMIN', 'SISTEMA', '00000000', 'ADMINISTRADOR DEL SISTEMA',
    'Informática', 'admin', salt, hash, util_fechaHoy(), 'Activo'
  ]);

  auditoria_registrar('admin', 'Usuario administrador inicial creado (config_crearUsuarioAdminInicial)', 'admin', 'Sistema');

  return {
    ok: true,
    usuario: 'admin',
    claveInicial: claveInicial,
    aviso: 'Cambia esta clave desde "Usuarios" apenas inicies sesión por primera vez.'
  };
}


/**
 * MIGRACIÓN (se ejecuta UNA vez a mano): convierte a MAYÚSCULAS los textos que ya
 * estaban guardados en minúscula. Es idempotente: volver a correrla no cambia nada más.
 *
 * Convierte solo textos que escribe una persona. NO toca, a propósito:
 *  - Estados y códigos que el programa compara tal cual (Activo, Pendiente, En progreso,
 *    Registrado, Resuelto, SI/NO) ni la columna AREA (Administración / Informática).
 *  - Contraseñas, SALT, HASH, usuario de acceso, usuario de red/correo, DNI, ni IDs/URLs de Drive.
 *
 * Los nombres de sistemas se convierten en TODAS las hojas donde aparecen
 * (Mantenimiento_Sistemas, Detalle_Accesos y Mantenimiento_Plantillas) para que sigan coincidiendo.
 */
function config_convertirTextosAMayusculas() {
  var columnas = {
    'Solicitudes': ['RESOLUCION', 'SEDE', 'APELLIDOS', 'NOMBRES', 'PERFIL'],
    'Detalle_Accesos': ['SISTEMA'],
    'Mantenimiento_Sede': ['SEDE'],
    'Mantenimiento_Perfil': ['PERFIL'],
    'Mantenimiento_Sistemas': ['NOMBRE_SISTEMA'],
    'Mantenimiento_Plantillas': ['NOMBRE_PLANTILLA', 'SISTEMA_ASOCIADO', 'CAMPOS_A_COMBINAR'],
    'Credenciales': ['NOMBRES', 'APELLIDOS', 'CARGO']
  };

  var resumen = {};
  Object.keys(columnas).forEach(function (nombreHoja) {
    var hoja = SS.getSheetByName(nombreHoja);
    if (!hoja || hoja.getLastRow() < 2) { resumen[nombreHoja] = 'sin datos'; return; }

    var encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
    var convertidas = 0;
    columnas[nombreHoja].forEach(function (encabezado) {
      var col = encabezados.indexOf(encabezado) + 1;
      if (col < 1) return;
      var rango = hoja.getRange(2, col, hoja.getLastRow() - 1, 1);
      var valores = rango.getValues();
      var huboCambio = false;
      valores.forEach(function (fila) {
        var v = fila[0];
        if (typeof v === 'string' && v !== '') {
          var mayus = v.trim().toUpperCase();
          if (mayus !== v) { fila[0] = mayus; convertidas++; huboCambio = true; }
        }
      });
      if (huboCambio) rango.setValues(valores);
    });
    resumen[nombreHoja] = convertidas + ' celdas convertidas';
  });

  Logger.log('Conversión a mayúsculas: ' + JSON.stringify(resumen));
  try {
    auditoria_registrar('sistema', 'Textos convertidos a mayúsculas (config_convertirTextosAMayusculas)', '', 'Sistema', JSON.stringify(resumen));
  } catch (e) { /* si Auditoria aún no existe, no bloquea la conversión */ }
  return resumen;
}
