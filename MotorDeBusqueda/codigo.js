
const CONFIG = {
  SHEET_NAME: "texto_detallado",           
  COL_ARCHIVO: "nombre_archivo",
  COL_RUTA: "ruta_completa",
  COL_PAGINA: "pagina",
  COL_TEXTO: "texto",

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

function buscarTermino(termino) {
  if (!termino || termino.trim().length < 2) return { resultados: [], total: 0 };

  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!hoja) throw new Error("No se encontró la hoja '" + CONFIG.SHEET_NAME + "'.");

  const datos = hoja.getDataRange().getValues();
  const encabezados = datos[0].map(h => String(h).trim());

  const idxArchivo = encabezados.indexOf(CONFIG.COL_ARCHIVO);
  const idxPagina  = encabezados.indexOf(CONFIG.COL_PAGINA);
  const idxTexto   = encabezados.indexOf(CONFIG.COL_TEXTO);

  if (idxArchivo === -1 || idxPagina === -1 || idxTexto === -1) {
    throw new Error("No se encontraron las columnas esperadas en '" + CONFIG.SHEET_NAME + "'.");
  }

  const terminoNorm = quitarTildes_(termino.trim());
  const resultados = [];

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const textoOriginal = String(fila[idxTexto] || "");
    const textoNorm = quitarTildes_(textoOriginal);
    const posicion = textoNorm.indexOf(terminoNorm);

    if (posicion !== -1) {
      const fragmento = construirFragmento_(textoOriginal, posicion, terminoNorm.length, termino);
      const nombreArchivo = String(fila[idxArchivo]);

      resultados.push({
        archivo: nombreArchivo,
        pagina: fila[idxPagina],
        fragmento: fragmento,
        urlPdf: obtenerUrlDrivePorNombre_(nombreArchivo)
      });

      if (resultados.length >= CONFIG.MAX_RESULTADOS) break;
    }
  }

  return { resultados: resultados, total: resultados.length };
}

function buscarCompendio(termino, tipoImpuesto) {
  return buscarEnHojaExterna_(CONFIG.HOJA_COMPENDIO, termino, tipoImpuesto);
}

function buscarSentencias(termino, tipoImpuesto) {
  return buscarEnHojaExterna_(CONFIG.HOJA_SENTENCIAS, termino, tipoImpuesto);
}

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


function obtenerTiposImpuesto(nombreHoja) {
  const ss = SpreadsheetApp.openById(CONFIG.COMPENDIO_SPREADSHEET_ID);
  const hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) return [];

  const datos = hoja.getDataRange().getValues();
  const encabezados = datos[0].map(h => String(h).trim());
  const idxTipo = encabezados.indexOf(CONFIG.COL_TIPO_IMPUESTO);
  if (idxTipo === -1) return [];

  const set = new Set();
  for (let i = 1; i < datos.length; i++) {
    const valor = String(datos[i][idxTipo] || "").trim();
    if (valor) set.add(valor);
  }
  return Array.from(set).sort();
}

function quitarTildes_(texto) {
  return texto.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
}

function escapeHtml_(s) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function escapeRegExp_(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function construirFragmento_(textoOriginal, posicion, largoTermino, terminoOriginal) {
  const mitad = CONFIG.CARACTERES_CONTEXTO / 2;
  const inicio = Math.max(0, posicion - mitad);
  const fin = Math.min(textoOriginal.length, posicion + largoTermino + mitad);
  let fragmento = (inicio > 0 ? "…" : "") + textoOriginal.substring(inicio, fin) + (fin < textoOriginal.length ? "…" : "");
  const fragmentoEscapado = escapeHtml_(fragmento);
  const regex = new RegExp(escapeRegExp_(terminoOriginal), "gi");
  return fragmentoEscapado.replace(regex, (m) => "<mark>" + m + "</mark>");
}

function formatearFecha_(valor) {
  if (!valor) return "";
  if (Object.prototype.toString.call(valor) === "[object Date]") {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  return String(valor);
}

function obtenerUrlDrivePorNombre_(nombreArchivo) {
  const cache = CacheService.getScriptCache();
  const claveCache = "url_" + nombreArchivo;
  const cacheado = cache.get(claveCache);
  if (cacheado) return cacheado === "NO_ENCONTRADO" ? null : cacheado;

  try {
    const archivos = DriveApp.getFilesByName(nombreArchivo);
    if (archivos.hasNext()) {
      const url = archivos.next().getUrl();
      cache.put(claveCache, url, 21600);
      return url;
    }
  } catch (e) {
  }
  cache.put(claveCache, "NO_ENCONTRADO", 21600);
  return null;
}