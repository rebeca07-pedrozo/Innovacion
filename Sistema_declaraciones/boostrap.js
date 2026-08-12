/**
 * ============================================================
 * BOOTSTRAP DEL SISTEMA
 * Crea la estructura inicial de carpetas y hojas.
 * EJECUTAR UNA SOLA VEZ: función bootstrapSistema()
 * ============================================================
 */

/**
 * Función principal de instalación.
 */
function bootstrapSistema() {
  const carpetaBase = localizarCarpetaBase_();
  const idsCarpetas = crearEstructuraCarpetas_(carpetaBase);

  const libroRegistro  = SpreadsheetApp.openById(ID_REGISTRO);
  const libroOperacion = SpreadsheetApp.openById(ID_OPERACION);

  crearHojasDelLibro_(libroRegistro,  HOJAS_REGISTRO);
  crearHojasDelLibro_(libroOperacion, HOJAS_OPERACION);

  escribirConfiguracion_(libroRegistro, idsCarpetas);
  sembrarDatosIniciales_(libroRegistro);

  Logger.log("Bootstrap completado correctamente.");
}

/**
 * Verifica que el libro de registro esté dentro de la carpeta base.
 */
function localizarCarpetaBase_() {
  const padres = DriveApp.getFileById(ID_REGISTRO).getParents();

  if (!padres.hasNext()) {
    throw new Error("REGISTRO_FORMULARIOS no está dentro de ninguna carpeta.");
  }

  const carpeta = padres.next();

  if (carpeta.getName() !== NOMBRE_CARPETA_BASE) {
    throw new Error(
      "REGISTRO_FORMULARIOS debe estar dentro de '" + NOMBRE_CARPETA_BASE +
      "'. Se encontró dentro de '" + carpeta.getName() + "'."
    );
  }

  return carpeta;
}

/**
 * Crea las carpetas del sistema y devuelve sus identificadores.
 */
function crearEstructuraCarpetas_(carpetaBase) {
  const ids = {};

  Object.keys(ESTRUCTURA_CARPETAS).forEach(function (nombrePrincipal) {
    const definicion = ESTRUCTURA_CARPETAS[nombrePrincipal];
    const carpeta = obtenerOCrearCarpeta(carpetaBase, nombrePrincipal);
    ids[nombrePrincipal] = carpeta.getId();

    definicion.subcarpetas.forEach(function (nombreSub) {
      const sub = obtenerOCrearCarpeta(carpeta, nombreSub);
      ids[nombrePrincipal + "/" + nombreSub] = sub.getId();

      definicion.nietos.forEach(function (nombreNieto) {
        obtenerOCrearCarpeta(sub, nombreNieto);
      });
    });
  });

  return ids;
}

/**
 * Crea todas las hojas definidas para un libro.
 */
function crearHojasDelLibro_(libro, definicionHojas) {
  Object.keys(definicionHojas).forEach(function (nombreHoja) {
    obtenerOCrearHoja(libro, nombreHoja, definicionHojas[nombreHoja]);
  });

  eliminarHojaPorDefecto(libro);
}

/**
 * Escribe los identificadores de libros y carpetas en la hoja CONFIG.
 */
function escribirConfiguracion_(libroRegistro, idsCarpetas) {
  const hoja = libroRegistro.getSheetByName("CONFIG");
  hoja.clear();

  hoja.getRange(1, 1, 1, 3)
      .setValues([["clave", "valor", "descripcion"]])
      .setFontWeight("bold")
      .setBackground(COLOR_ENCABEZADO);

  const filas = [
    ["ID_REGISTRO_FORMULARIOS",    ID_REGISTRO,  "Libro de catálogo"],
    ["ID_OPERACION_DECLARACIONES", ID_OPERACION, "Libro transaccional"]
  ];

  Object.keys(idsCarpetas).forEach(function (ruta) {
    filas.push([
      "CARPETA_" + ruta.replace(/\//g, "_").toUpperCase(),
      idsCarpetas[ruta],
      "Carpeta " + ruta
    ]);
  });

  hoja.getRange(2, 1, filas.length, 3).setValues(filas);
  hoja.setFrozenRows(1);
}

/**
 * Registra el formulario y las entidades iniciales.
 */
function sembrarDatosIniciales_(libroRegistro) {
  const hojaForm = libroRegistro.getSheetByName("FORMULARIOS");

  if (hojaForm.getLastRow() <= 1) {
    hojaForm.appendRow([
      FORMULARIO_INICIAL.codigo,
      FORMULARIO_INICIAL.nombre,
      FORMULARIO_INICIAL.version,
      "2026-01-01", "",
      FORMULARIO_INICIAL.periodicidad,
      FORMULARIO_INICIAL.paginas,
      "ACTIVO",
      new Date(),
      usuarioActual()
    ]);
  }

  const hojaEnt = libroRegistro.getSheetByName("ENTIDADES");

  if (hojaEnt.getLastRow() <= 1) {
    ENTIDADES_INICIALES.forEach(function (entidad) {
      hojaEnt.appendRow(entidad);
    });
  }
}