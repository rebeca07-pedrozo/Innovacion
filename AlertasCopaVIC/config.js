const CONFIG = {
  HOJA_EXTRACT: 'EXTRACT',
  HOJA_TRANSFORM: 'TRANSFORM',
  HOJA_LOAD: 'LOAD',
  HOJA_PARAMETROS: 'PARAMETROS_ALERTA',
  HOJA_ANOMALIAS: 'LOG_ANOMALIAS',

  ANIO_REFERENCIA: null,

  URL_WEBAPP: 'https://script.google.com/a/macros/davivienda.com/s/AKfycbxaZBFfABWbNo5OLy0KA4wu2QghTBGNd9Eyho-PGOsqU1ycrdiepDLaISd5aTAGDhUV/exec',
  LOGO_BASE64: 'PEGA_AQUI_TU_BASE64_QUE_YA_GENERASTE_DESDE_EL_DOC',
  URL_ICONO_CALENDARIO: 'https://cdn-icons-png.flaticon.com/512/747/747310.png',

  DIAS_UMBRAL_URGENTE: 2,

  COL_EXTRACT: {
    COMPANIA: 'Compañía',
    NIT: 'NIT',
    IMPUESTO: 'Impuesto',
    FECHA_MAXIMA: 'Fecha máxima de presentación',
    ENCARGADO: 'Persona encargada',
    JEFE1: 'Jefe  1',
    JEFE2: 'Jefe 2 (Opcional)',
    CIFRAS: 'Cifras',
    MUNICIPIO: 'Municipio'
  },

  COL_TRANSFORM: [
    'ID',
    'Compañía (Normalizada)', 'Compañía (Original)', 'NIT', 'Impuesto',
    'Fecha máxima de presentación', 'Fecha Válida',
    'Encargado Nombre', 'Encargado Email',
    'Jefe1 Nombre', 'Jefe1 Email',
    'Jefe2 Nombre', 'Jefe2 Email',
    'Cifras', 'Municipio'
  ],

  COL_LOAD: [
    'ID',
    'Compañía (Normalizada)', 'Compañía (Original)', 'NIT', 'Impuesto',
    'Fecha máxima de presentación', 'Fecha Válida',
    'Encargado Nombre', 'Encargado Email',
    'Jefe1 Nombre', 'Jefe1 Email',
    'Jefe2 Nombre', 'Jefe2 Email',
    'Cifras', 'Municipio',
    'Fecha Notificado', 'Fecha En Proceso', 'Fecha Presentado',
    'Estado Actual', 'Semáforo', 'Días Restantes', 'Extemporaneidad (días)',
    'Última Fecha Envío Recordatorio'
  ],

  COL_ANOMALIAS: ['Fila EXTRACT', 'Compañía', 'NIT', 'Impuesto', 'Motivo', 'Valor Original'],

  IMPUESTOS_CON_MUNICIPIO: ['retencion de ica', 'autorretencion de ica', 'autoretencion de ica', 'exogenas'],

  ESTADOS: {
    PENDIENTE: 'Pendiente',
    NOTIFICADO: 'Notificado',
    EN_PROCESO: 'En proceso',
    PRESENTADO: 'Presentado'
  },

  COLORES_ESTADO: {
    'Pendiente': '#F2F2F2',
    'Notificado': '#D9E8F5',
    'En proceso': '#FCE8D5',
    'Presentado': '#DCF0E1'
  }
};

function obtenerMapaColumnas(hoja) {
  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const mapa = {};
  encabezados.forEach((nombre, idx) => {
    if (nombre) mapa[String(nombre).trim()] = idx;
  });
  return mapa;
}

function valorPorColumna(fila, mapa, nombreColumna) {
  const idx = mapa[nombreColumna];
  if (idx === undefined) return null;
  return fila[idx];
}

function asegurarHojaParametros() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(CONFIG.HOJA_PARAMETROS);
  if (hoja) return hoja;

  hoja = libro.insertSheet(CONFIG.HOJA_PARAMETROS);
  hoja.getRange(1, 1, 1, 2).setValues([['Días antes del vencimiento', 'Enviar recordatorio (TRUE/FALSE)']]);
  const umbralesDefault = [7, 3, 2, 1, 0].map(d => [d, true]);
  hoja.getRange(2, 1, umbralesDefault.length, 2).setValues(umbralesDefault);
  hoja.getRange(1, 1, 1, 2).setFontWeight('bold');
  hoja.autoResizeColumns(1, 2);
  return hoja;
}

function asegurarHojaAnomalias() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(CONFIG.HOJA_ANOMALIAS);
  if (!hoja) hoja = libro.insertSheet(CONFIG.HOJA_ANOMALIAS);
  hoja.clearContents();
  hoja.getRange(1, 1, 1, CONFIG.COL_ANOMALIAS.length).setValues([CONFIG.COL_ANOMALIAS]);
  hoja.getRange(1, 1, 1, CONFIG.COL_ANOMALIAS.length).setFontWeight('bold');
  return hoja;
}

function escribirAnomalias(listaAnomalias) {
  const hoja = asegurarHojaAnomalias();
  if (listaAnomalias.length > 0) {
    hoja.getRange(2, 1, listaAnomalias.length, CONFIG.COL_ANOMALIAS.length).setValues(listaAnomalias);
  }
  hoja.autoResizeColumns(1, CONFIG.COL_ANOMALIAS.length);
}

function generarIdObligacion(compania, nit, impuesto, fechaMaximaTexto, municipio) {
  const partes = [compania, nit, impuesto, fechaMaximaTexto];
  if (municipio) partes.push(municipio);
  const textoClave = partes.join('|').toUpperCase().trim();
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, textoClave);
  const hash = bytes.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
  return 'OBL-' + hash.substring(0, 10).toUpperCase();
}

function buscarFilaPorId(hoja, id) {
  const mapa = obtenerMapaColumnas(hoja);
  const idxId = mapa['ID'];
  if (idxId === undefined) return null;
  const datos = hoja.getRange(2, 1, Math.max(hoja.getLastRow() - 1, 0), hoja.getLastColumn()).getValues();
  for (let i = 0; i < datos.length; i++) {
    if (datos[i][idxId] === id) {
      return { numeroFila: i + 2, valores: datos[i], mapa: mapa };
    }
  }
  return null;
}

function calcularSemaforo(estado, diasRestantes) {
  if (estado === CONFIG.ESTADOS.PRESENTADO) return '🟢';
  if (typeof diasRestantes === 'number' && diasRestantes <= CONFIG.DIAS_UMBRAL_URGENTE) return '🔴';
  return '🟡';
}