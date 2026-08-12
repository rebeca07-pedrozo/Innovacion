/**
 * ============================================================
 * PROMOCIÓN DEL CATÁLOGO
 *
 * Pasa los renglones de STAGING_EXTRACCION a RENGLONES,
 * previa verificación de integridad.
 *
 * Ejecutar en orden:
 *   1. verificarStaging()   → revisa sin escribir nada
 *   2. promoverCatalogo()   → escribe en RENGLONES
 * ============================================================
 */

// ---------- Renglones que debe tener el F350 ----------
const RENGLONES_ESPERADOS_F350 = (function () {
  const lista = [];
  for (let i = 29; i <= 138; i++) lista.push(i);
  for (let i = 148; i <= 155; i++) lista.push(i);
  return lista;
})();

// ---------- Datos maestros del formulario ----------
const DATOS_FORMULARIO = [
  "F350", "Declaración de retenciones en la fuente", "v2026",
  "MENSUAL", 2, "ACTIVO", new Date()
];

const DATOS_ENTIDADES = [
  ["DAVIVIENDA", "BANCO DAVIVIENDA S.A.", "860034313", "7", "31", "6412", "SI"],
  ["DAVIBANK",   "DAVIBANK",              "",          "",  "",   "",     "SI"]
];


/**
 * Revisa el contenido de staging sin modificar nada.
 */
function verificarStaging() {
  const resultado = analizarStaging_();

  Logger.log("--- VERIFICACIÓN DE STAGING ---");
  Logger.log("Renglones encontrados: " + resultado.total);
  Logger.log("Esperados: " + RENGLONES_ESPERADOS_F350.length);
  Logger.log("Faltantes: "  + (resultado.faltantes.length  ? resultado.faltantes.join(", ")  : "ninguno"));
  Logger.log("Sobrantes: "  + (resultado.sobrantes.length  ? resultado.sobrantes.join(", ")  : "ninguno"));
  Logger.log("Duplicados: " + (resultado.duplicados.length ? resultado.duplicados.join(", ") : "ninguno"));
  Logger.log("Sin etiqueta: " + (resultado.sinEtiqueta.length ? resultado.sinEtiqueta.join(", ") : "ninguno"));

  Logger.log(resultado.valido
    ? "\nCatálogo íntegro. Se puede promover."
    : "\nHay inconsistencias. Corregir en STAGING_EXTRACCION antes de promover.");

  return resultado;
}


/**
 * Promueve el catálogo a RENGLONES si la verificación es correcta.
 */
function promoverCatalogo() {
  const resultado = analizarStaging_();

  if (!resultado.valido) {
    throw new Error(
      "El staging tiene inconsistencias. Ejecutar verificarStaging() para ver el detalle."
    );
  }

  const libro = SpreadsheetApp.openById(ID_REGISTRO);
  const hojaRenglones = libro.getSheetByName("RENGLONES");

  // Se reemplaza el catálogo de esta versión, conservando otras versiones
  eliminarVersionExistente_(hojaRenglones, "F350", "v2026");

  const filas = resultado.filas.map(function (f) {
    return [
      f.cod_formulario, f.version, f.pagina, f.nro_renglon, f.etiqueta,
      f.tipo_persona, f.tipo_valor, f.grupo_concepto, f.seccion, f.editable
    ];
  });

  hojaRenglones.getRange(
    hojaRenglones.getLastRow() + 1, 1, filas.length, filas[0].length
  ).setValues(filas);

  registrarFormulario_(libro);
  registrarEntidades_(libro);
  marcarStagingPromovido_(libro);

  Logger.log("Renglones promovidos: " + filas.length);
}


/**
 * Lee staging y evalúa su integridad.
 */
function analizarStaging_() {
  const hoja = SpreadsheetApp.openById(ID_REGISTRO).getSheetByName("STAGING_EXTRACCION");

  if (!hoja) {
    throw new Error("No se encontró la hoja STAGING_EXTRACCION.");
  }

  const datos = hoja.getDataRange().getValues();
  const encabezados = datos[0].map(function (h) { return String(h).trim(); });

  // Los índices se resuelven por nombre para no depender del orden de columnas
  const col = {};
  encabezados.forEach(function (nombre, i) { col[nombre] = i; });

  const requeridas = ["cod_formulario", "version", "pagina", "nro_renglon",
                      "etiqueta", "tipo_persona", "tipo_valor",
                      "grupo_concepto", "seccion", "editable"];

  requeridas.forEach(function (c) {
    if (col[c] === undefined) {
      throw new Error("Falta la columna '" + c + "' en STAGING_EXTRACCION.");
    }
  });

  const filas = [];
  const vistos = [];
  const duplicados = [];
  const sinEtiqueta = [];

  for (let i = 1; i < datos.length; i++) {
    const nro = parseInt(datos[i][col.nro_renglon], 10);
    if (isNaN(nro)) continue;

    if (vistos.indexOf(nro) >= 0) {
      duplicados.push(nro);
    } else {
      vistos.push(nro);
    }

    const etiqueta = String(datos[i][col.etiqueta]).trim();
    if (etiqueta.length < 3) sinEtiqueta.push(nro);

    filas.push({
      cod_formulario: datos[i][col.cod_formulario],
      version:        datos[i][col.version],
      pagina:         datos[i][col.pagina],
      nro_renglon:    nro,
      etiqueta:       etiqueta,
      tipo_persona:   datos[i][col.tipo_persona],
      tipo_valor:     datos[i][col.tipo_valor],
      grupo_concepto: datos[i][col.grupo_concepto],
      seccion:        datos[i][col.seccion],
      editable:       datos[i][col.editable]
    });
  }

  const faltantes = RENGLONES_ESPERADOS_F350.filter(function (n) {
    return vistos.indexOf(n) < 0;
  });

  const sobrantes = vistos.filter(function (n) {
    return RENGLONES_ESPERADOS_F350.indexOf(n) < 0;
  });

  return {
    filas: filas,
    total: filas.length,
    faltantes: faltantes,
    sobrantes: sobrantes,
    duplicados: duplicados,
    sinEtiqueta: sinEtiqueta,
    valido: !faltantes.length && !sobrantes.length &&
            !duplicados.length && !sinEtiqueta.length
  };
}


/**
 * Elimina de RENGLONES las filas de un formulario y versión concretos.
 */
function eliminarVersionExistente_(hoja, codFormulario, version) {
  if (hoja.getLastRow() <= 1) return;

  const datos = hoja.getDataRange().getValues();

  // Se recorre de abajo hacia arriba para que los índices no se desplacen
  for (let i = datos.length - 1; i >= 1; i--) {
    if (datos[i][0] === codFormulario && datos[i][1] === version) {
      hoja.deleteRow(i + 1);
    }
  }
}


/**
 * Registra el formulario en FORMULARIOS si aún no existe.
 */
function registrarFormulario_(libro) {
  const hoja = libro.getSheetByName("FORMULARIOS");
  const datos = hoja.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === "F350" && datos[i][2] === "v2026") return;
  }

  hoja.appendRow(DATOS_FORMULARIO);
}


/**
 * Registra las entidades si la hoja está vacía.
 */
function registrarEntidades_(libro) {
  const hoja = libro.getSheetByName("ENTIDADES");
  if (hoja.getLastRow() > 1) return;

  DATOS_ENTIDADES.forEach(function (entidad) {
    hoja.appendRow(entidad);
  });
}


/**
 * Marca las filas de staging como ya promovidas.
 */
function marcarStagingPromovido_(libro) {
  const hoja = libro.getSheetByName("STAGING_EXTRACCION");
  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const idx = encabezados.indexOf("estado_revision");

  if (idx < 0 || hoja.getLastRow() <= 1) return;

  const numFilas = hoja.getLastRow() - 1;
  const valores = [];
  for (let i = 0; i < numFilas; i++) valores.push(["PROMOVIDO"]);

  hoja.getRange(2, idx + 1, numFilas, 1).setValues(valores);
}