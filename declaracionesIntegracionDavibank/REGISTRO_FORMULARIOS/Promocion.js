/**
 * ============================================================
 * PROMOCIÓN DEL CATÁLOGO
 *
 * Copia los renglones de STAGING_EXTRACCION a RENGLONES,
 * previa verificación de integridad.
 *
 * Ejecutar en orden:
 *   1. verificarStaging()   → revisa, no escribe
 *   2. promoverCatalogo()   → escribe en RENGLONES
 * ============================================================
 */

// Renglones esperados por formulario
const RENGLONES_ESPERADOS = {
  F350: (function () {
    const l = [];
    for (let i = 29; i <= 138; i++) l.push(i);
    for (let i = 148; i <= 155; i++) l.push(i);
    return l;
  })(),

  F300: (function () {
    const l = [];
    for (let i = 1; i <= 6; i++) l.push(i);
    for (let i = 27; i <= 93; i++) l.push(i);
    l.push(100);
    return l;
  })()
};


/**
 * Revisa el contenido de staging sin modificar nada.
 */
function verificarStaging() {
  const r = analizarStaging_();

  Logger.log("--- VERIFICACIÓN DE STAGING ---");
  Logger.log("Formulario: " + r.codFormulario + " " + r.version);
  Logger.log("Renglones encontrados: " + r.total + " de " + r.esperados);
  Logger.log("Faltantes: "    + (r.faltantes.length   ? r.faltantes.join(", ")   : "ninguno"));
  Logger.log("Sobrantes: "    + (r.sobrantes.length   ? r.sobrantes.join(", ")   : "ninguno"));
  Logger.log("Duplicados: "   + (r.duplicados.length  ? r.duplicados.join(", ")  : "ninguno"));
  Logger.log("Sin etiqueta: " + (r.sinEtiqueta.length ? r.sinEtiqueta.join(", ") : "ninguno"));

  Logger.log(r.valido
    ? "\nCatálogo íntegro. Se puede promover."
    : "\nHay inconsistencias. Corregir en STAGING_EXTRACCION antes de promover.");
}


/**
 * Promueve el catálogo a RENGLONES si la verificación es correcta.
 */
function promoverCatalogo() {
  const r = analizarStaging_();

  if (!r.valido) {
    throw new Error("El staging tiene inconsistencias. Ejecutar verificarStaging().");
  }

  const hoja = SpreadsheetApp.openById(ID_REGISTRO).getSheetByName("RENGLONES");

  // Se reemplaza el catálogo de este formulario y versión
  eliminarVersion_(hoja, r.codFormulario, r.version);

  const filas = r.filas.map(function (f) {
    return [
      f.cod_formulario, f.version, f.pagina, f.nro_renglon, f.etiqueta,
      f.tipo_persona, f.tipo_valor, f.grupo_concepto, f.seccion, f.editable
    ];
  });

  hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, filas[0].length)
      .setValues(filas);

  registrarFormulario_(r.codFormulario, r.version);

  Logger.log("Renglones promovidos: " + filas.length + " (" + r.codFormulario + ")");
}


/**
 * Lee staging y evalúa su integridad.
 */
function analizarStaging_() {
  const hoja = SpreadsheetApp.openById(ID_REGISTRO).getSheetByName("STAGING_EXTRACCION");

  if (!hoja) throw new Error("No se encontró la hoja STAGING_EXTRACCION.");

  const datos = hoja.getDataRange().getValues();

  const col = {};
  datos[0].forEach(function (nombre, i) { col[String(nombre).trim()] = i; });

  ["cod_formulario", "version", "pagina", "nro_renglon", "etiqueta",
   "tipo_persona", "tipo_valor", "grupo_concepto", "seccion", "editable"
  ].forEach(function (c) {
    if (col[c] === undefined) {
      throw new Error("Falta la columna '" + c + "' en STAGING_EXTRACCION.");
    }
  });

  const filas = [];
  const vistos = [];
  const duplicados = [];
  const sinEtiqueta = [];
  let codFormulario = "";
  let version = "";

  for (let i = 1; i < datos.length; i++) {
    const nro = parseInt(datos[i][col.nro_renglon], 10);
    if (isNaN(nro)) continue;

    // El formulario se toma de los propios datos
    if (!codFormulario) {
      codFormulario = String(datos[i][col.cod_formulario]).trim();
      version = String(datos[i][col.version]).trim();
    }

    if (vistos.indexOf(nro) >= 0) duplicados.push(nro);
    else vistos.push(nro);

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

  const esperados = RENGLONES_ESPERADOS[codFormulario];

  if (!esperados) {
    throw new Error("No hay lista de renglones esperados para " + codFormulario +
                    ". Agregarla en RENGLONES_ESPERADOS.");
  }

  const faltantes = esperados.filter(function (n) { return vistos.indexOf(n) < 0; });
  const sobrantes = vistos.filter(function (n) { return esperados.indexOf(n) < 0; });

  return {
    codFormulario: codFormulario,
    version: version,
    filas: filas,
    total: filas.length,
    esperados: esperados.length,
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
function eliminarVersion_(hoja, codFormulario, version) {
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
function registrarFormulario_(codFormulario, version) {
  const hoja = SpreadsheetApp.openById(ID_REGISTRO).getSheetByName("FORMULARIOS");
  const datos = hoja.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === codFormulario && datos[i][2] === version) return;
  }

  const catalogo = {
    F350: ["Declaración de retenciones en la fuente", "MENSUAL",   2],
    F300: ["Declaración del impuesto sobre las ventas - IVA", "BIMESTRAL", 1]
  };

  const datosForm = catalogo[codFormulario] || [codFormulario, "MENSUAL", 1];

  hoja.appendRow([
    codFormulario, datosForm[0], version, datosForm[1],
    datosForm[2], "ACTIVO", new Date()
  ]);
}