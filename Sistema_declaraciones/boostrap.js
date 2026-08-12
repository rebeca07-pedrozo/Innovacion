/**
 * ============================================================
 * INSTALACIÓN DEL SISTEMA DE DECLARACIONES
 *
 * Hace dos cosas:
 *   1. Registra en CONFIG los IDs de las carpetas de Drive
 *   2. Crea las hojas de trabajo en los dos libros
 *
 * No crea carpetas: solo lee las que ya existen.
 * Ejecutar: instalarSistema()
 * ============================================================
 */

// ---------- Identificadores de los libros ----------
const ID_REGISTRO  = "1qXjlhZgsK-jOyGxlf0Vk65WLDK-Oypm6yXbF5SF7H5s";
const ID_OPERACION = "1ghhI0GK8HFI1lP-WMs6Ce2pBHNH_fnSspJ9wIZWectU";

const NOMBRE_CARPETA_BASE = "SISTEMA_DECLARACIONES";
const COLOR_ENCABEZADO    = "#D9E2F3";

// ---------- Hojas del libro de catálogo ----------
const HOJAS_REGISTRO = {
  CONFIG: ["clave", "valor", "descripcion"],

  FORMULARIOS: [
    "cod_formulario", "nombre", "version", "periodicidad",
    "num_paginas", "estado", "fecha_alta"
  ],

  RENGLONES: [
    "cod_formulario", "version", "pagina", "nro_renglon", "etiqueta",
    "tipo_persona", "tipo_valor", "grupo_concepto", "seccion", "editable"
  ],

  ENTIDADES: [
    "cod_entidad", "nombre", "nit", "dv",
    "cod_direccion_seccional", "actividad_economica", "activa"
  ]
};

// ---------- Hojas del libro transaccional ----------
const HOJAS_OPERACION = {
  PLANTILLAS_EMITIDAS: [
    "id_plantilla", "cod_formulario", "version", "cod_entidad",
    "anio", "periodo", "id_archivo_drive", "generada_por",
    "fecha_generacion", "estado"
  ],

  ENTREGAS: [
    "radicado", "id_plantilla", "cod_formulario", "cod_entidad",
    "anio", "periodo", "nombre_archivo", "id_archivo_drive",
    "cargada_por", "fecha_carga", "estado", "num_errores"
  ],

  DATOS_CARGADOS: [
    "radicado", "cod_formulario", "version",
    "nro_renglon", "valor", "fecha_registro"
  ]
};


/**
 * Función principal. Ejecutar una sola vez.
 */
function instalarSistema() {
  const carpetaBase = localizarCarpetaBase_();
  const carpetas    = recorrerCarpetas_(carpetaBase, "");

  crearHojas_(SpreadsheetApp.openById(ID_REGISTRO),  HOJAS_REGISTRO);
  crearHojas_(SpreadsheetApp.openById(ID_OPERACION), HOJAS_OPERACION);

  escribirConfig_(carpetas, carpetaBase.getId());

  Logger.log("Instalación completada.");
  Logger.log("Carpetas registradas: " + carpetas.length);
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
 * Recorre las subcarpetas y devuelve su ruta e identificador.
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
 * Crea las hojas definidas para un libro, respetando las existentes.
 */
function crearHojas_(libro, definicion) {
  Object.keys(definicion).forEach(function (nombreHoja) {
    let hoja = libro.getSheetByName(nombreHoja);

    if (!hoja) {
      hoja = libro.insertSheet(nombreHoja);
    }

    if (hoja.getLastRow() === 0) {
      const encabezados = definicion[nombreHoja];
      hoja.getRange(1, 1, 1, encabezados.length)
          .setValues([encabezados])
          .setFontWeight("bold")
          .setBackground(COLOR_ENCABEZADO);
      hoja.setFrozenRows(1);
    }
  });

  // Se elimina la hoja por defecto si quedó vacía
  ["Hoja 1", "Hoja1", "Sheet1"].forEach(function (nombre) {
    const hoja = libro.getSheetByName(nombre);
    if (hoja && libro.getSheets().length > 1 && hoja.getLastRow() === 0) {
      libro.deleteSheet(hoja);
    }
  });
}


/**
 * Escribe los identificadores de libros y carpetas en la hoja CONFIG.
 */
function escribirConfig_(carpetas, idCarpetaBase) {
  const hoja = SpreadsheetApp.openById(ID_REGISTRO).getSheetByName("CONFIG");

  hoja.clear();
  hoja.getRange(1, 1, 1, 3)
      .setValues([["clave", "valor", "descripcion"]])
      .setFontWeight("bold")
      .setBackground(COLOR_ENCABEZADO);

  const filas = [
    ["ID_REGISTRO_FORMULARIOS",    ID_REGISTRO,   "Libro de catálogo"],
    ["ID_OPERACION_DECLARACIONES", ID_OPERACION,  "Libro transaccional"],
    ["CARPETA_BASE",               idCarpetaBase, "Carpeta raíz del sistema"]
  ];

  carpetas.forEach(function (c) {
    filas.push(["CARPETA_" + normalizarClave_(c.ruta), c.id, c.ruta]);
  });

  hoja.getRange(2, 1, filas.length, 3).setValues(filas);
  hoja.setFrozenRows(1);
  hoja.autoResizeColumns(1, 3);
}


/**
 * Convierte una ruta de carpeta en una clave válida.
 */
function normalizarClave_(ruta) {
  return ruta
    .toUpperCase()
    .replace(/\//g, "_")
    .replace(/\./g, "")
    .replace(/\s+/g, "_")
    .replace(/_+/g, "_");
}