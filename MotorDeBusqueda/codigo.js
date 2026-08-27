
const CONFIG = {
  SHEET_NAME: "texto_detallado",         
  COL_ARCHIVO: "nombre_archivo",
  COL_RUTA: "ruta_completa",
  COL_PAGINA: "pagina",
  COL_TEXTO: "texto",
  MAX_RESULTADOS: 50,
  CARACTERES_CONTEXTO: 240
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Interfaz')
    .setTitle('Buscador de Normativas - Davivienda')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function buscarTermino(termino) {
  if (!termino || termino.trim().length < 2) {
    return { resultados: [], total: 0 };
  }

  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!hoja) {
    throw new Error("No se encontró la hoja '" + CONFIG.SHEET_NAME + "'. Revisa CONFIG.SHEET_NAME en Code.gs");
  }

  const datos = hoja.getDataRange().getValues();
  const encabezados = datos[0].map(h => String(h).trim());

  const idxArchivo = encabezados.indexOf(CONFIG.COL_ARCHIVO);
  const idxPagina  = encabezados.indexOf(CONFIG.COL_PAGINA);
  const idxTexto   = encabezados.indexOf(CONFIG.COL_TEXTO);

  if (idxArchivo === -1 || idxPagina === -1 || idxTexto === -1) {
    throw new Error("No se encontraron las columnas esperadas. Revisa los nombres en CONFIG.");
  }

  const terminoNorm = quitarTildes_(termino.trim());
  const resultados = [];

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const textoOriginal = String(fila[idxTexto] || "");
    const textoNorm = quitarTildes_(textoOriginal);
    const posicion = textoNorm.indexOf(terminoNorm);

    if (posicion !== -1) {
      const mitad = CONFIG.CARACTERES_CONTEXTO / 2;
      const inicio = Math.max(0, posicion - mitad);
      const fin = Math.min(textoOriginal.length, posicion + terminoNorm.length + mitad);

      let fragmento = (inicio > 0 ? "…" : "") +
                       textoOriginal.substring(inicio, fin) +
                       (fin < textoOriginal.length ? "…" : "");

      fragmento = resaltarTermino_(fragmento, termino);

      resultados.push({
        archivo: fila[idxArchivo],
        pagina: fila[idxPagina],
        fragmento: fragmento
      });

      if (resultados.length >= CONFIG.MAX_RESULTADOS) break;
    }
  }

  return { resultados: resultados, total: resultados.length };
}

// --- Utilidades internas ---

function quitarTildes_(texto) {
  return texto.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
}

function escapeHtml_(s) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function escapeRegExp_(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function resaltarTermino_(fragmento, termino) {
  const fragmentoEscapado = escapeHtml_(fragmento);
  const regex = new RegExp(escapeRegExp_(termino), "gi");
  return fragmentoEscapado.replace(regex, (m) => "<mark>" + m + "</mark>");
}