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
function promoverCatalogo() {
  const r = analizarStaging_();

  if (!r.valido) {
    throw new Error("El staging tiene inconsistencias. Ejecutar verificarStaging().");
  }

  const hoja = SpreadsheetApp.openById(ID_REGISTRO).getSheetByName("RENGLONES");
  eliminarVersion_(hoja, r.codFormulario, r.version);

  const filas = r.filas.map(function (f) {
    return [
      f.cod_formulario, f.version, f.pagina, f.nro_renglon, f.etiqueta,
      f.tipo_persona, f.tipo_valor, f.grupo_concepto, f.seccion, f.editable
    ];
  });

  hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, filas[0].length)
      .setValues(filas);

  Logger.log("Renglones promovidos: " + filas.length + " (" + r.codFormulario + ")");
}
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
  const esperados = definicionFormulario(codFormulario).esperados;
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
function eliminarVersion_(hoja, codFormulario, version) {
  if (hoja.getLastRow() <= 1) return;
  const datos = hoja.getDataRange().getValues();
  for (let i = datos.length - 1; i >= 1; i--) {
    if (datos[i][0] === codFormulario && datos[i][1] === version) {
      hoja.deleteRow(i + 1);
    }
  }
}