function onEditBuzon(e) {
  const hoja = e.range.getSheet();
  if (hoja.getName() !== "home") return;
  if (e.range.getColumn() !== 1) return; // solo reacciona si el pegado fue en la columna A

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila === 0) return;

  const valoresColumnaA = hoja.getRange(1, 1, ultimaFila, 1).getValues().map(f => String(f[0]).trim());

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hojaDestino = ss.getSheetByName("motorDeBusqueda");
  if (!hojaDestino) {
    hojaDestino = ss.insertSheet("motorDeBusqueda");
    hojaDestino.appendRow([
      "nombre_archivo", "tipo_documento", "entidad_emisora", "tipo_impuesto",
      "tema", "subtema", "radicado", "fecha", "pregunta_consulta",
      "resumen_respuesta", "articulos_citados", "normas_referenciadas",
      "conclusion_clave", "palabras_clave"
    ]);
  }

  const datosExistentes = hojaDestino.getDataRange().getValues();
  const columnaNombre = datosExistentes[0].indexOf("nombre_archivo");
  const nombresYaCargados = new Set(datosExistentes.slice(1).map(fila => fila[columnaNombre]));

  const filasNuevas = agruparPorBloquesDe14_(valoresColumnaA, nombresYaCargados);

  if (filasNuevas.length > 0) {
    hojaDestino.getRange(hojaDestino.getLastRow() + 1, 1, filasNuevas.length, filasNuevas[0].length)
      .setValues(filasNuevas);
  }

  // Limpia TODA la columna A usada (no solo A1)
  hoja.getRange(1, 1, ultimaFila, 1).clearContent();

  hoja.getRange("B1").setValue(
    filasNuevas.length > 0
      ? `✅ ${filasNuevas.length} documento(s) agregado(s) — ${new Date().toLocaleString()}`
      : `⚠️ No se detectaron filas nuevas válidas. Revisa el contenido pegado. — ${new Date().toLocaleString()}`
  );
}

function agruparPorBloquesDe14_(lineas, nombresYaCargados) {
  const NUM_COLUMNAS = 14;

  lineas = lineas.filter(l => l !== "");

  // Salta el bloque de encabezado si está presente al inicio
  if (lineas[0] === "nombre_archivo") {
    lineas = lineas.slice(NUM_COLUMNAS);
  }

  const filas = [];
  for (let i = 0; i + NUM_COLUMNAS <= lineas.length; i += NUM_COLUMNAS) {
    const bloque = lineas.slice(i, i + NUM_COLUMNAS);
    const nombreArchivo = bloque[0];

    if (nombreArchivo && !nombresYaCargados.has(nombreArchivo)) {
      filas.push(bloque);
      nombresYaCargados.add(nombreArchivo);
    }
  }

  return filas;
}