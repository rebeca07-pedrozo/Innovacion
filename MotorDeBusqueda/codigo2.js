// ============================================================
// CONFIGURACIÓN
// ============================================================
const CONFIG = {
  SHEET_NAME: "Hoja1",              // <-- ajusta al nombre real de tu pestaña
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
  MAX_RESULTADOS_APOYO: 5,
  CARACTERES_CONTEXTO: 240
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Interfaz')
    .setTitle('Buscador de Normativas - Davivienda')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================
// BÚSQUEDAS PRINCIPALES
// ============================================================
function buscarCompendio(termino, tipoImpuesto) {
  const principal = buscarEnHojaExterna_(CONFIG.HOJA_COMPENDIO, termino, tipoImpuesto);
  const apoyo = (termino && termino.trim().length >= 2) ? buscarApoyoInterno_(termino) : { resultados: [], total: 0 };
  registrarBusqueda_("Compendio DIAN", termino, tipoImpuesto, principal.total);
  return { principal: principal, apoyo: apoyo };
}

function buscarSentencias(termino, tipoImpuesto) {
  const principal = buscarEnHojaExterna_(CONFIG.HOJA_SENTENCIAS, termino, tipoImpuesto);
  const apoyo = (termino && termino.trim().length >= 2) ? buscarApoyoInterno_(termino) : { resultados: [], total: 0 };
  registrarBusqueda_("Sentencias", termino, tipoImpuesto, principal.total);
  return { principal: principal, apoyo: apoyo };
}
// ============================================================
// Fuente interna de apoyo (normativas propias) — no visible como categoría
// ============================================================
function buscarApoyoInterno_(termino) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!hoja) return { resultados: [], total: 0 };

  const datos = hoja.getDataRange().getValues();
  const encabezados = datos[0].map(h => String(h).trim());

  const idxArchivo = encabezados.indexOf(CONFIG.COL_ARCHIVO);
  const idxPagina  = encabezados.indexOf(CONFIG.COL_PAGINA);
  const idxTexto   = encabezados.indexOf(CONFIG.COL_TEXTO);
  if (idxArchivo === -1 || idxPagina === -1 || idxTexto === -1) return { resultados: [], total: 0 };

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
      const infoDrive = obtenerInfoDrivePorNombre_(nombreArchivo);

      resultados.push({
        archivo: nombreArchivo,
        pagina: fila[idxPagina],
        fragmento: fragmento,
        urlPreview: infoDrive ? infoDrive.previewUrl : null,
        urlDescarga: infoDrive ? infoDrive.url : null
      });

      if (resultados.length >= CONFIG.MAX_RESULTADOS_APOYO) break;
    }
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

// ============================================================
// UTILIDADES
// ============================================================
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

/**
 * Busca el archivo en Drive por nombre, le asegura permiso de visualización
 * dentro del dominio (para que nadie tope con pantalla de "solicitar acceso"),
 * y devuelve tanto el link de descarga como el link de previsualización embebible.
 * Usa caché (6h) para no repetir esta operación en cada búsqueda.
 */
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
      } catch (errorPermisos) {
        // no rompemos la búsqueda si esto falla
      }

      const info = {
        url: archivo.getUrl(),
        previewUrl: "https://drive.google.com/file/d/" + id + "/preview"
      };

      cache.put(claveCache, JSON.stringify(info), 21600); // 6 horas para resultados positivos, está bien
      return info;
    }
  } catch (e) {
    // si falla la búsqueda, seguimos sin romper el resto
  }

  // Antes: 21600 (6 horas) también para "no encontrado" — eso causaba el bug.
  // Ahora: solo 120 segundos, así un falso negativo se autocorrige rápido.
  cache.put(claveCache, "NO_ENCONTRADO", 120);
  return null;
  // ============================================================
// NUEVO: Registro de uso del buscador (para KPIs)
// ============================================================
function registrarBusqueda_(categoria, termino, tipoImpuestoFiltro, totalResultados) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hojaLogs = ss.getSheetByName("Logs_Busquedas");

    if (!hojaLogs) {
      hojaLogs = ss.insertSheet("Logs_Busquedas");
      hojaLogs.appendRow(["fecha_hora", "categoria", "termino_buscado", "filtro_tipo_impuesto", "total_resultados"]);
    }

    hojaLogs.appendRow([
      new Date(),
      categoria,
      termino || "",
      (tipoImpuestoFiltro && tipoImpuestoFiltro !== "TODOS") ? tipoImpuestoFiltro : "",
      totalResultados
    ]);
  } catch (e) {
    // el logging nunca debe romper la búsqueda del usuario
    Logger.log("Error registrando búsqueda: " + e.message);
  }
}
}