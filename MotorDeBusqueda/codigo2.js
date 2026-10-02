// ============================================================
// CONFIGURACIÓN
// ============================================================
const CONFIG = {
  // --- Normativas propias (ahora estructurado, ya no texto crudo por página) ---
  SHEET_MOTOR: "motorDeBusqueda",
  COL_ARCHIVO: "nombre_archivo",
  COL_TIPO_DOC: "tipo_documento",
  COL_ENTIDAD: "entidad_emisora",
  COL_TIPO_IMPUESTO_PROPIO: "tipo_impuesto",
  COL_TEMA: "tema",
  COL_SUBTEMA: "subtema",
  COL_RADICADO: "radicado",
  COL_FECHA_PROPIO: "fecha",
  COL_PREGUNTA: "pregunta_consulta",
  COL_RESUMEN_RESP: "resumen_respuesta",
  COL_ARTICULOS: "articulos_citados",
  COL_NORMAS_REF: "normas_referenciadas",
  COL_CONCLUSION: "conclusion_clave",
  COL_PALABRAS: "palabras_clave",

  // --- Compendio DIAN / Sentencias (Excel del abogado) ---
  COMPENDIO_SPREADSHEET_ID: "1R4gZpTwd1PBaE8yj3ruJezQYoqmrOUuGUfSRmMOFtGE",
  HOJA_COMPENDIO: "Compendio Completo Doctrina y Conceptos DIAN",
  HOJA_SENTENCIAS: "Sentencias",
  COL_NUMERO: "Número de Documento / Sentencia",
  COL_FECHA: "Fecha",
  COL_TITULO: "Título",
  COL_TIPO_IMPUESTO: "Tipo de Impuesto Evaluado",
  COL_RESUMEN: "Resumen",

  MAX_RESULTADOS: 50,
  CARACTERES_CONTEXTO: 240
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Interfaz')
    .setTitle('Buscador de Normativas - Davivienda')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================
// BÚSQUEDA: Compendio DIAN
// ============================================================
function buscarCompendio(termino, tipoImpuesto) {
  const resultado = buscarEnHojaExterna_(CONFIG.HOJA_COMPENDIO, termino, tipoImpuesto);
  registrarBusqueda_("Compendio DIAN", termino, resultado.total);
  return resultado;
}

// ============================================================
// BÚSQUEDA: Sentencias
// ============================================================
function buscarSentencias(termino, tipoImpuesto) {
  const resultado = buscarEnHojaExterna_(CONFIG.HOJA_SENTENCIAS, termino, tipoImpuesto);
  registrarBusqueda_("Sentencias", termino, resultado.total);
  return resultado;
}

// ============================================================
// BÚSQUEDA: Normativas Propias (estructurado)
// ============================================================
function buscarNormativasPropias(termino, tipoDocumento, tipoImpuesto) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_MOTOR);
  if (!hoja) return { resultados: [], total: 0 };

  const datos = hoja.getDataRange().getValues();
  const enc = datos[0].map(h => String(h).trim());
  const idx = (col) => enc.indexOf(col);

  const iNombre = idx(CONFIG.COL_ARCHIVO);
  const iTipoDoc = idx(CONFIG.COL_TIPO_DOC);
  const iEntidad = idx(CONFIG.COL_ENTIDAD);
  const iImpuesto = idx(CONFIG.COL_TIPO_IMPUESTO_PROPIO);
  const iTema = idx(CONFIG.COL_TEMA);
  const iSubtema = idx(CONFIG.COL_SUBTEMA);
  const iRadicado = idx(CONFIG.COL_RADICADO);
  const iFecha = idx(CONFIG.COL_FECHA_PROPIO);
  const iPregunta = idx(CONFIG.COL_PREGUNTA);
  const iResumen = idx(CONFIG.COL_RESUMEN_RESP);
  const iArticulos = idx(CONFIG.COL_ARTICULOS);
  const iNormasRef = idx(CONFIG.COL_NORMAS_REF);
  const iConclusion = idx(CONFIG.COL_CONCLUSION);
  const iPalabras = idx(CONFIG.COL_PALABRAS);

  const tieneTermino = termino && termino.trim().length >= 2;
  const terminoNorm = tieneTermino ? quitarTildes_(termino.trim()) : "";
  const resultados = [];

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    if (!fila[iNombre]) continue;

    if (tipoDocumento && tipoDocumento !== "TODOS" && String(fila[iTipoDoc]).trim() !== tipoDocumento) continue;
    if (tipoImpuesto && tipoImpuesto !== "TODOS" && String(fila[iImpuesto]).trim() !== tipoImpuesto) continue;

    if (tieneTermino) {
      const bolsaTexto = quitarTildes_([
        fila[iTema], fila[iSubtema], fila[iPregunta], fila[iResumen],
        fila[iConclusion], fila[iPalabras], fila[iArticulos], fila[iNormasRef]
      ].join(" "));
      if (bolsaTexto.indexOf(terminoNorm) === -1) continue;
    }

    const nombreArchivo = String(fila[iNombre]);
    const infoDrive = obtenerInfoDrivePorNombre_(nombreArchivo);

    resultados.push({
      archivo: nombreArchivo,
      tipoDocumento: fila[iTipoDoc],
      entidad: fila[iEntidad],
      tipoImpuesto: fila[iImpuesto],
      tema: fila[iTema],
      subtema: fila[iSubtema],
      radicado: fila[iRadicado],
      fecha: formatearFecha_(fila[iFecha]),
      pregunta: fila[iPregunta],
      resumen: fila[iResumen],
      articulos: fila[iArticulos],
      normasRef: fila[iNormasRef],
      conclusion: fila[iConclusion],
      palabras: fila[iPalabras],
      urlPreview: infoDrive ? infoDrive.previewUrl : null
    });

    if (resultados.length >= CONFIG.MAX_RESULTADOS) break;
  }

  registrarBusqueda_("Normativas Propias", termino, resultados.length);
  return { resultados: resultados, total: resultados.length };
}

// ============================================================
// Búsqueda genérica sobre Compendio DIAN / Sentencias
// ============================================================
function buscarEnHojaExterna_(nombreHoja, termino, tipoImpuesto) {
  const tieneTermino = termino && termino.trim().length >= 2;
  const tieneFiltro = tipoImpuesto && tipoImpuesto !== "TODOS";

  if (!tieneTermino && !tieneFiltro) return { resultados: [], total: 0 };

  const ss = SpreadsheetApp.openById(CONFIG.COMPENDIO_SPREADSHEET_ID);
  const hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) throw new Error("No se encontró la hoja '" + nombreHoja + "' en el Excel del compendio.");

  const datos = hoja.getDataRange().getValues();
  const encabezados = datos[0].map(h => String(h).trim());

  const idxNumero = encabezados.indexOf(CONFIG.COL_NUMERO);
  const idxFecha  = encabezados.indexOf(CONFIG.COL_FECHA);
  const idxTitulo = encabezados.indexOf(CONFIG.COL_TITULO);
  const idxTipo   = encabezados.indexOf(CONFIG.COL_TIPO_IMPUESTO);
  const idxResumen = encabezados.indexOf(CONFIG.COL_RESUMEN);

  const terminoNorm = tieneTermino ? quitarTildes_(termino.trim()) : "";
  const resultados = [];

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const titulo = String(fila[idxTitulo] || "");
    const resumen = String(fila[idxResumen] || "");
    const tipo = String(fila[idxTipo] || "");

    if (tieneFiltro && tipo.trim() !== tipoImpuesto) continue;

    if (tieneTermino) {
      const combinado = quitarTildes_(titulo + " " + resumen);
      if (combinado.indexOf(terminoNorm) === -1) continue;
    }

    resultados.push({
      numero: fila[idxNumero],
      fecha: formatearFecha_(fila[idxFecha]),
      titulo: titulo,
      tipoImpuesto: tipo,
      resumen: resumen
    });

    if (resultados.length >= CONFIG.MAX_RESULTADOS) break;
  }

  return { resultados: resultados, total: resultados.length };
}

// ============================================================
// VALORES ÚNICOS PARA DESPLEGABLES
// ============================================================
function obtenerTiposImpuesto(nombreHoja) {
  const ss = SpreadsheetApp.openById(CONFIG.COMPENDIO_SPREADSHEET_ID);
  const hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) return [];
  return obtenerValoresUnicosDeHoja_(hoja, CONFIG.COL_TIPO_IMPUESTO);
}

function obtenerTiposDocumentoPropio() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_MOTOR);
  if (!hoja) return [];
  return obtenerValoresUnicosDeHoja_(hoja, CONFIG.COL_TIPO_DOC);
}

function obtenerTiposImpuestoPropio() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_MOTOR);
  if (!hoja) return [];
  return obtenerValoresUnicosDeHoja_(hoja, CONFIG.COL_TIPO_IMPUESTO_PROPIO);
}

function obtenerValoresUnicosDeHoja_(hoja, nombreColumna) {
  const datos = hoja.getDataRange().getValues();
  const idx = datos[0].map(h => String(h).trim()).indexOf(nombreColumna);
  if (idx === -1) return [];
  const set = new Set();
  for (let i = 1; i < datos.length; i++) {
    const v = String(datos[i][idx] || "").trim();
    if (v) set.add(v);
  }
  return Array.from(set).sort();
}

// ============================================================
// REGISTRO DE USO (para los KPIs)
// ============================================================
function registrarBusqueda_(categoria, termino, totalResultados) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hojaLogs = ss.getSheetByName("Logs_Busquedas");
    if (!hojaLogs) {
      hojaLogs = ss.insertSheet("Logs_Busquedas");
      hojaLogs.appendRow(["fecha_hora", "categoria", "termino_buscado", "total_resultados"]);
    }
    hojaLogs.appendRow([new Date(), categoria, termino || "", totalResultados]);
  } catch (e) {
    Logger.log("Error registrando búsqueda: " + e.message);
  }
}

// ============================================================
// DASHBOARD (KPIs)
// ============================================================
function actualizarKPIs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hojaKPIs = ss.getSheetByName("KPIs");
  if (!hojaKPIs) hojaKPIs = ss.insertSheet("KPIs");
  hojaKPIs.clear();

  const ssCompendio = SpreadsheetApp.openById(CONFIG.COMPENDIO_SPREADSHEET_ID);
  const totalCompendio = ssCompendio.getSheetByName(CONFIG.HOJA_COMPENDIO).getDataRange().getNumRows() - 1;
  const totalSentencias = ssCompendio.getSheetByName(CONFIG.HOJA_SENTENCIAS).getDataRange().getNumRows() - 1;

  const hojaMotor = ss.getSheetByName(CONFIG.SHEET_MOTOR);
  const totalPropias = hojaMotor ? hojaMotor.getDataRange().getNumRows() - 1 : 0;

  const hojaLogs = ss.getSheetByName("Logs_Busquedas");
  const logs = hojaLogs ? hojaLogs.getDataRange().getValues().slice(1) : [];
  const totalBusquedas = logs.length;
  const sinResultados = logs.filter(f => Number(f[3]) === 0).length;

  const porTermino = {};
  logs.forEach(f => {
    const t = String(f[2] || "").trim().toLowerCase();
    if (t) porTermino[t] = (porTermino[t] || 0) + 1;
  });
  const top10Terminos = Object.entries(porTermino).sort((a, b) => b[1] - a[1]).slice(0, 10);

  const filas = [
    ["MÉTRICA", "VALOR"],
    ["Total documentos - Compendio DIAN", totalCompendio],
    ["Total documentos - Sentencias", totalSentencias],
    ["Total documentos - Normativas propias", totalPropias],
    ["", ""],
    ["Total búsquedas realizadas", totalBusquedas],
    ["Búsquedas sin resultados", sinResultados],
    ["", ""],
    ["TOP 10 TÉRMINOS MÁS BUSCADOS", "VECES BUSCADO"],
  ];
  top10Terminos.forEach(([termino, veces]) => filas.push([termino, veces]));

  hojaKPIs.getRange(1, 1, filas.length, 2).setValues(filas);
  hojaKPIs.getRange(1, 1, 1, 2).setFontWeight("bold");
  hojaKPIs.getRange(9, 1, 1, 2).setFontWeight("bold");
  hojaKPIs.autoResizeColumns(1, 2);

  Logger.log("KPIs actualizados.");
}

// ============================================================
// BUZÓN NOTEBOOKLM → motorDeBusqueda
// ============================================================
const NUM_COLUMNAS_BUZON = 14;

function onEditBuzon(e) {
  const hoja = e.range.getSheet();
  if (hoja.getName() !== "home") return;
  if (e.range.getColumn() !== 1) return;
  procesarBuzon_();
}

function procesarBuzonManual() {
  procesarBuzon_();
}

function procesarBuzon_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaHome = ss.getSheetByName("home");
  if (!hojaHome) return;

  const ultimaFila = hojaHome.getLastRow();
  if (ultimaFila === 0) return;

  const valoresColumnaA = hojaHome.getRange(1, 1, ultimaFila, 1).getValues().map(f => String(f[0]).trim());

  let hojaDestino = ss.getSheetByName(CONFIG.SHEET_MOTOR);
  if (!hojaDestino) {
    hojaDestino = ss.insertSheet(CONFIG.SHEET_MOTOR);
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
    : `⚠️ No se detectaron filas nuevas válidas. — ${new Date().toLocaleString()}`;

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

// ============================================================
// UTILIDADES
// ============================================================
function quitarTildes_(texto) {
  return String(texto).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
}

function escapeHtml_(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function formatearFecha_(valor) {
  if (!valor) return "";
  if (Object.prototype.toString.call(valor) === "[object Date]") {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  return String(valor);
}

function obtenerInfoDrivePorNombre_(nombreArchivo) {
  const cache = CacheService.getScriptCache();
  const claveCache = "info_" + nombreArchivo;
  const cacheado = cache.get(claveCache);
  if (cacheado) return cacheado === "NO_ENCONTRADO" ? null : JSON.parse(cacheado);

  try {
    const archivos = DriveApp.getFilesByName(nombreArchivo);
    if (archivos.hasNext()) {
      const archivo = archivos.next();
      const id = archivo.getId();
      try {
        archivo.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (errorPermisos) {}

      const info = {
        url: archivo.getUrl(),
        previewUrl: "https://drive.google.com/file/d/" + id + "/preview"
      };
      cache.put(claveCache, JSON.stringify(info), 21600);
      return info;
    }
  } catch (e) {}

  cache.put(claveCache, "NO_ENCONTRADO", 120);
  return null;
}
