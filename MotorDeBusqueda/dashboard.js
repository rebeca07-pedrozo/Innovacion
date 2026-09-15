// ============================================================
// DASHBOARD.GS — genera y actualiza el Sheet de KPIs para Looker Studio
// ============================================================

const CONFIG_DASHBOARD = {
  NOMBRE_DASHBOARD: "Dashboard_MotorBusqueda_Tributario",
  // Se completa solo la primera vez que corras crearOActualizarDashboard()
  DASHBOARD_SPREADSHEET_ID: "" // <-- déjalo vacío la primera vez
};

function crearOActualizarDashboard() {
  const ss = obtenerOCrearDashboard_();

  actualizarKpisVolumen_(ss);
  actualizarKpisUso_(ss);

  Logger.log("Dashboard actualizado: " + ss.getUrl());
  Logger.log("Guarda este ID para futuras corridas: " + ss.getId());
}

function obtenerOCrearDashboard_() {
  if (CONFIG_DASHBOARD.DASHBOARD_SPREADSHEET_ID) {
    return SpreadsheetApp.openById(CONFIG_DASHBOARD.DASHBOARD_SPREADSHEET_ID);
  }
  const nuevo = SpreadsheetApp.create(CONFIG_DASHBOARD.NOMBRE_DASHBOARD);
  Logger.log("Se creó un Dashboard nuevo. ID: " + nuevo.getId());
  Logger.log("Cópialo en CONFIG_DASHBOARD.DASHBOARD_SPREADSHEET_ID para no crear uno nuevo cada vez.");
  return nuevo;
}

// --- KPIs de VOLUMEN: cuántos documentos hay, por fuente y por tipo de impuesto ---
function actualizarKpisVolumen_(ssDashboard) {
  const ssCompendio = SpreadsheetApp.openById(CONFIG.COMPENDIO_SPREADSHEET_ID);
  const hojaCompendio = ssCompendio.getSheetByName(CONFIG.HOJA_COMPENDIO);
  const hojaSentencias = ssCompendio.getSheetByName(CONFIG.HOJA_SENTENCIAS);

  const datosCompendio = hojaCompendio.getDataRange().getValues();
  const datosSentencias = hojaSentencias.getDataRange().getValues();

  const encCompendio = datosCompendio[0].map(h => String(h).trim());
  const encSentencias = datosSentencias[0].map(h => String(h).trim());
  const idxTipoC = encCompendio.indexOf(CONFIG.COL_TIPO_IMPUESTO);
  const idxTipoS = encSentencias.indexOf(CONFIG.COL_TIPO_IMPUESTO);

  const totalCompendio = datosCompendio.length - 1;
  const totalSentencias = datosSentencias.length - 1;

  // Normativas propias: contar archivos únicos y páginas totales
  const ssPropio = SpreadsheetApp.getActiveSpreadsheet();
  const hojaPropia = ssPropio.getSheetByName(CONFIG.SHEET_NAME);
  const datosPropios = hojaPropia.getDataRange().getValues();
  const encPropio = datosPropios[0].map(h => String(h).trim());
  const idxArchivoPropio = encPropio.indexOf(CONFIG.COL_ARCHIVO);
  const archivosUnicosPropios = new Set(datosPropios.slice(1).map(f => f[idxArchivoPropio]));

  // --- Tabla 1: resumen general ---
  const resumen = [
    ["fuente", "total_documentos"],
    ["Compendio DIAN", totalCompendio],
    ["Sentencias", totalSentencias],
    ["Normativas propias (archivos)", archivosUnicosPropios.size],
    ["Normativas propias (páginas)", datosPropios.length - 1],
  ];
  escribirTabla_(ssDashboard, "KPI_Resumen", resumen);

  // --- Tabla 2: distribución por tipo de impuesto (Compendio + Sentencias combinados) ---
  const conteoTipos = {};
  if (idxTipoC !== -1) {
    for (let i = 1; i < datosCompendio.length; i++) {
      const tipo = String(datosCompendio[i][idxTipoC] || "Sin clasificar").trim() || "Sin clasificar";
      conteoTipos[tipo] = (conteoTipos[tipo] || 0) + 1;
    }
  }
  if (idxTipoS !== -1) {
    for (let i = 1; i < datosSentencias.length; i++) {
      const tipo = String(datosSentencias[i][idxTipoS] || "Sin clasificar").trim() || "Sin clasificar";
      conteoTipos[tipo] = (conteoTipos[tipo] || 0) + 1;
    }
  }
  const tablaTipos = [["tipo_impuesto", "cantidad_documentos"]];
  Object.keys(conteoTipos).sort().forEach(tipo => {
    tablaTipos.push([tipo, conteoTipos[tipo]]);
  });
  escribirTabla_(ssDashboard, "KPI_Por_Tipo_Impuesto", tablaTipos);
}

// --- KPIs de USO: total de búsquedas, términos más buscados, búsquedas por día ---
function actualizarKpisUso_(ssDashboard) {
  const ssPropio = SpreadsheetApp.getActiveSpreadsheet();
  const hojaLogs = ssPropio.getSheetByName("Logs_Busquedas");

  if (!hojaLogs) {
    escribirTabla_(ssDashboard, "KPI_Uso_Resumen", [["metrica", "valor"], ["total_busquedas", 0]]);
    return;
  }

  const datos = hojaLogs.getDataRange().getValues();
  const filas = datos.slice(1); // sin encabezado

  const totalBusquedas = filas.length;

  // --- Búsquedas por categoría ---
  const porCategoria = {};
  filas.forEach(f => {
    const cat = f[1] || "Sin categoría";
    porCategoria[cat] = (porCategoria[cat] || 0) + 1;
  });
  const tablaCategoria = [["categoria", "total_busquedas"]];
  Object.keys(porCategoria).forEach(cat => tablaCategoria.push([cat, porCategoria[cat]]));

  // --- Términos más buscados (top 30) ---
  const porTermino = {};
  filas.forEach(f => {
    const termino = String(f[2] || "").trim().toLowerCase();
    if (!termino) return;
    porTermino[termino] = (porTermino[termino] || 0) + 1;
  });
  const tablaTerminos = [["termino_buscado", "veces_buscado"]];
  Object.entries(porTermino)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 30)
    .forEach(([termino, veces]) => tablaTerminos.push([termino, veces]));

  // --- Búsquedas por día ---
  const porDia = {};
  filas.forEach(f => {
    const fecha = new Date(f[0]);
    const clave = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "yyyy-MM-dd");
    porDia[clave] = (porDia[clave] || 0) + 1;
  });
  const tablaDias = [["fecha", "total_busquedas"]];
  Object.keys(porDia).sort().forEach(dia => tablaDias.push([dia, porDia[dia]]));

  // --- Búsquedas sin resultados (para saber qué le falta al sistema) ---
  const sinResultados = filas.filter(f => Number(f[4]) === 0).length;

  escribirTabla_(ssDashboard, "KPI_Uso_Resumen", [
    ["metrica", "valor"],
    ["total_busquedas", totalBusquedas],
    ["busquedas_sin_resultados", sinResultados],
  ]);
  escribirTabla_(ssDashboard, "KPI_Uso_Por_Categoria", tablaCategoria);
  escribirTabla_(ssDashboard, "KPI_Terminos_Mas_Buscados", tablaTerminos);
  escribirTabla_(ssDashboard, "KPI_Busquedas_Por_Dia", tablaDias);
}

// --- Utilidad: escribe una tabla completa en una pestaña, creándola si no existe ---
function escribirTabla_(ss, nombrePestaña, filas) {
  let hoja = ss.getSheetByName(nombrePestaña);
  if (!hoja) {
    hoja = ss.insertSheet(nombrePestaña);
  } else {
    hoja.clear();
  }
  if (filas.length > 0) {
    hoja.getRange(1, 1, filas.length, filas[0].length).setValues(filas);
  }
}