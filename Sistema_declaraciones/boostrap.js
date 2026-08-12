/**
 * ============================================================
 * MAPEO DE CARPETAS
 * Recorre la estructura existente en Drive y registra
 * los identificadores en la hoja CONFIG.
 * No crea ni modifica carpetas.
 * Ejecutar: mapearCarpetas()
 * ============================================================
 */

function mapearCarpetas() {
  const carpetaBase = localizarCarpetaBase_();
  const encontradas = recorrerCarpetas_(carpetaBase, "");

  if (encontradas.length === 0) {
    throw new Error("No se encontró ninguna subcarpeta dentro de " + NOMBRE_CARPETA_BASE);
  }

  escribirConfiguracion_(encontradas, carpetaBase.getId());
  Logger.log("Carpetas registradas: " + encontradas.length);
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
 * Recorre recursivamente las subcarpetas y devuelve su ruta e identificador.
 */
function recorrerCarpetas_(carpeta, rutaPadre) {
  let resultado = [];
  const hijas = carpeta.getFolders();

  while (hijas.hasNext()) {
    const hija = hijas.next();
    const ruta = rutaPadre ? rutaPadre + "/" + hija.getName() : hija.getName();

    resultado.push({ ruta: ruta, id: hija.getId() });
    resultado = resultado.concat(recorrerCarpetas_(hija, ruta));
  }

  return resultado;
}

/**
 * Escribe los identificadores de libros y carpetas en la hoja CONFIG.
 */
function escribirConfiguracion_(carpetas, idCarpetaBase) {
  const libro = SpreadsheetApp.openById(ID_REGISTRO);
  const hoja = obtenerOCrearHoja(libro, "CONFIG", ["clave", "valor", "descripcion"]);

  hoja.clear();
  hoja.getRange(1, 1, 1, 3)
      .setValues([["clave", "valor", "descripcion"]])
      .setFontWeight("bold")
      .setBackground(COLOR_ENCABEZADO);

  const filas = [
    ["ID_REGISTRO_FORMULARIOS",    ID_REGISTRO,    "Libro de catálogo"],
    ["ID_OPERACION_DECLARACIONES", ID_OPERACION,   "Libro transaccional"],
    ["CARPETA_BASE",               idCarpetaBase,  "Carpeta raíz del sistema"]
  ];

  carpetas.forEach(function (c) {
    const clave = "CARPETA_" + normalizarClave_(c.ruta);
    filas.push([clave, c.id, c.ruta]);
  });

  hoja.getRange(2, 1, filas.length, 3).setValues(filas);
  hoja.setFrozenRows(1);
  hoja.autoResizeColumns(1, 3);
}

/**
 * Convierte una ruta de carpeta en una clave válida para CONFIG.
 */
function normalizarClave_(ruta) {
  return ruta
    .toUpperCase()
    .replace(/\//g, "_")   // separador de nivel
    .replace(/\./g, "")    // puntos de numeración
    .replace(/\s+/g, "_")  // espacios
    .replace(/_+/g, "_");  // guiones bajos repetidos
}