// ============================================================
// CONFIGURACIÓN
// ============================================================
const CONFIG = {
  SHEET_MOTOR: "texto_detallado",
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

  COMPENDIO_SPREADSHEET_ID: "1R4gZpTwd1PBaE8yj3ruJezQYoqmrOUuGUfSRmMOFtGE",
  HOJA_COMPENDIO: "Compendio Completo Doctrina y Conceptos DIAN",
  HOJA_SENTENCIAS: "Sentencias",
  COL_NUMERO: "Número de Documento / Sentencia",
  COL_FECHA: "Fecha",
  COL_TITULO: "Título",
  COL_TIPO_IMPUESTO: "Tipo de Impuesto Evaluado",
  COL_RESUMEN: "Resumen",

  MAX_RESULTADOS: 50
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Interfaz')
    .setTitle('Buscador de Normativas - Davivienda')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================
// BÚSQUEDA UNIFICADA (lo que llama la interfaz)
// ============================================================
function buscarTodo(termino, tipoImpuesto) {
  if (!termino || termino.trim().length < 2) return { resultados: [] };

  const filtro = tipoImpuesto || "TODOS";

  const compendio = buscarEnHojaExterna_(CONFIG.HOJA_COMPENDIO, termino, filtro);
  const sentencias = buscarEnHojaExterna_(CONFIG.HOJA_SENTENCIAS, termino, filtro);
  const propias = buscarNormativasPropias(termino, "TODOS", filtro).resultados;

  const resultados = [];

  compendio.resultados.forEach(r => resultados.push({
    titulo: r.titulo,
    fecha: r.fecha,
    detalle: r.resumen,
    urlPreview: null
  }));

  sentencias.resultados.forEach(r => resultados.push({
    titulo: r.titulo,
    fecha: r.fecha,
    detalle: r.resumen,
    urlPreview: null
  }));

  propias.forEach(r => resultados.push({
    titulo: r.tema,
    fecha: r.fecha,
    detalle: r.resumen,
    urlPreview: r.urlPreview,
    archivo: r.archivo
  }));

  registrarBusqueda_("Búsqueda unificada", termino, resultados.length);
  return { resultados: resultados, total: resultados.length };
}

function obtenerTodosLosTiposImpuesto() {
  const deCompendio = obtenerTiposImpuesto(CONFIG.HOJA_COMPENDIO);
  const deSentencias = obtenerTiposImpuesto(CONFIG.HOJA_SENTENCIAS);
  const dePropias = obtenerTiposImpuestoPropio();
  const set = new Set([...deCompendio, ...deSentencias, ...dePropias]);
  return Array.from(set).sort();
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
  const palabrasBuscadas = tieneTermino ? tokenizar_(termino) : [];
  const resultados = [];

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    if (!fila[iNombre]) continue;

    if (tipoDocumento && tipoDocumento !== "TODOS" && String(fila[iTipoDoc]).trim() !== tipoDocumento) continue;
    if (tipoImpuesto && tipoImpuesto !== "TODOS" && String(fila[iImpuesto]).trim() !== tipoImpuesto) continue;

    const bolsaTexto = quitarTildes_([
      fila[iTema], fila[iSubtema], fila[iPregunta], fila[iResumen],
      fila[iConclusion], fila[iPalabras], fila[iArticulos], fila[iNormasRef]
    ].join(" "));

    if (tieneTermino && !contieneTodasLasPalabras_(bolsaTexto, palabrasBuscadas)) continue;

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

  const palabrasBuscadas = tieneTermino ? tokenizar_(termino) : [];
  const resultados = [];

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const titulo = String(fila[idxTitulo] || "");
    const resumen = String(fila[idxResumen] || "");
    const tipo = String(fila[idxTipo] || "");

    if (tieneFiltro && tipo.trim() !== tipoImpuesto) continue;

    if (tieneTermino) {
      const combinado = quitarTildes_(titulo + " " + resumen);
      if (!contieneTodasLasPalabras_(combinado, palabrasBuscadas)) continue;
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
// UTILIDADES
// ============================================================
function quitarTildes_(texto) {
  return String(texto).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
}

// Separa el término buscado en palabras sueltas (sin tildes, sin palabras de 1 letra)
function tokenizar_(termino) {
  return quitarTildes_(termino).split(/\s+/).filter(p => p.length > 1);
}

// Verdadero solo si TODAS las palabras buscadas aparecen en algún lugar del texto (sin importar el orden)
function contieneTodasLasPalabras_(textoNormalizado, palabras) {
  return palabras.every(p => textoNormalizado.indexOf(p) !== -1);
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