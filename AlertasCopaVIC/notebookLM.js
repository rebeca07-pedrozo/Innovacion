function onEditBuzon(e) {
  const hoja = e.range.getSheet();

  if (hoja.getName() !== "home") return;
  if (e.range.getRow() !== 1 || e.range.getColumn() !== 1) return;

  const textoPegado = hoja.getRange("A1").getValue();
  if (!textoPegado || String(textoPegado).trim() === "") return;

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

  const filasNuevas = parsearTablaMarkdown_(textoPegado, nombresYaCargados);

  if (filasNuevas.length > 0) {
    hojaDestino.getRange(hojaDestino.getLastRow() + 1, 1, filasNuevas.length, filasNuevas[0].length)
      .setValues(filasNuevas);
  }

  hoja.getRange("A1").clearContent();

  hoja.getRange("B1").setValue(
    filasNuevas.length > 0
      ? `✅ ${filasNuevas.length} documento(s) agregado(s) — ${new Date().toLocaleString()}`
      : `⚠️ No se detectaron filas nuevas válidas. Revisa el formato de lo pegado. — ${new Date().toLocaleString()}`
  );
}

function parsearTablaMarkdown_(texto, nombresYaCargados) {
  const lineas = String(texto).split("\n").map(linea => linea.trim()).filter(linea => linea.startsWith("|"));

  const filas = [];
  lineas.forEach(linea => {
    if (/^\|[\s\-|:]+\|$/.test(linea)) return; // salta línea separadora |---|---|

    const celdas = linea.split("|").map(celda => celda.trim())
      .filter((celda, i, arr) => !(i === 0 && celda === "") && !(i === arr.length - 1 && celda === ""));

    if (celdas[0] === "nombre_archivo") return; // salta encabezado repetido

    if (celdas[0] && !nombresYaCargados.has(celdas[0])) {
      filas.push(celdas);
      nombresYaCargados.add(celdas[0]);
    }
  });

  return filas;
}