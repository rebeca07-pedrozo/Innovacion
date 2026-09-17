// ============================================================
// BUZÓN NOTEBOOKLM → motorDeBusqueda
// ============================================================

const NUM_COLUMNAS_BUZON = 14;

// --- Se dispara SOLA cada vez que se pega/edita algo en la columna A de "home" ---
function onEditBuzon(e) {
  const hoja = e.range.getSheet();
  if (hoja.getName() !== "home") return;
  if (e.range.getColumn() !== 1) return;

  procesarBuzon_();
}

// --- Puedes correr esta manualmente desde el editor cuando quieras forzar el procesamiento ---
function procesarBuzonManual() {
  procesarBuzon_();
}

// --- Lógica compartida por las dos anteriores ---
function procesarBuzon_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaHome = ss.getSheetByName("home");
  if (!hojaHome) {
    Logger.log("No existe la pestaña 'home'.");
    return;
  }

  const ultimaFila = hojaHome.getLastRow();
  if (ultimaFila === 0) {
    Logger.log("La pestaña 'home' está vacía.");
    return;
  }

  const valoresColumnaA = hojaHome.getRange(1, 1, ultimaFila, 1).getValues().map(f => String(f[0]).trim());

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

  hojaHome.getRange(1, 1, ultimaFila, 1).clearContent();

  const mensaje = filasNuevas.length > 0
    ? `✅ ${filasNuevas.length} documento(s) agregado(s) — ${new Date().toLocaleString()}`
    : `⚠️ No se detectaron filas nuevas válidas. Revisa el contenido pegado. — ${new Date().toLocaleString()}`;

  hojaHome.getRange("B1").setValue(mensaje);
  Logger.log(mensaje);
}

function agruparPorBloquesDe14_(lineas, nombresYaCargados) {
  lineas = lineas.filter(l => l !== "");

  if (lineas[0] === "nombre_archivo") {
    lineas = lineas.slice(NUM_COLUMNAS_BUZON);
  }

  const filas = [];
  for (let i = 0; i + NUM_COLUMNAS_BUZON <= lineas.length; i += NUM_COLUMNAS_BUZON) {
    const bloque = lineas.slice(i, i + NUM_COLUMNAS_BUZON);
    const nombreArchivo = bloque[0];

    if (nombreArchivo && !nombresYaCargados.has(nombreArchivo)) {
      filas.push(bloque);
      nombresYaCargados.add(nombreArchivo);
    }
  }

  return filas;
}