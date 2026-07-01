/***** LÓGICA DEL TABLERO VISUAL *****/

function obtenerDatosTablero_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const obl = ss.getSheetByName(CONFIG.HOJA_OBL);
  if (!obl || obl.getLastRow() < 2) return null;
  
  const vals = obl.getDataRange().getValues();
  const h = vals[0], idx = n => h.indexOf(n);
  
  let datos = {
    kpis: { total: 0, proximas: 0, vencidas: 0, companias: new Set() },
    tabla: [],
    fechas: {},
    impuestos: {}
  };

  const hoy = new Date();
  hoy.setHours(0,0,0,0);

  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    if (r[idx('estado')] === 'Presentado' || !r[idx('id_vencimiento')]) continue;

    const dias = Number(r[idx('dias_restantes')]);
    const compania = String(r[idx('compania')]);
    const impuesto = r[idx('tipo_obligacion')] + ' · ' + r[idx('subtipo')];
    const fecha = r[idx('fecha_limite')];

    // Acumular KPIs
    datos.kpis.total++;
    datos.kpis.companias.add(compania);
    if (dias < 0) datos.kpis.vencidas++;
    if (dias >= 0 && dias <= 15) datos.kpis.proximas++;

    // Preparar filas para la tabla
    datos.tabla.push({
      compania: compania,
      impuesto: impuesto,
      fechaStr: Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'd MMM yyyy').toLowerCase(),
      dias: dias,
      estado: r[idx('estado')],
      colorSemaforo: dias <= 3 ? '#E24B4A' : (dias <= 15 ? '#EF9F27' : '#639922')
    });

    // Agrupar para los gráficos
    if (fecha instanceof Date) {
      const dia = fecha.getDate();
      datos.fechas[dia] = (datos.fechas[dia] || 0) + 1;
    }
    datos.impuestos[impuesto] = (datos.impuestos[impuesto] || 0) + 1;
  }

  // Ordenar la tabla por días restantes (los más urgentes primero)
  datos.tabla.sort((a, b) => a.dias - b.dias);
  datos.totalOcultas = Math.max(0, datos.tabla.length - 8);
  datos.tabla = datos.tabla.slice(0, 8); // Mostrar solo el top 8
  datos.kpis.companias = datos.kpis.companias.size;

  return datos;
}

// Lanza esta función desde el editor para ver el tablero en tu navegador
function probarTablero() {
  const t = HtmlService.createTemplateFromFile('TableroUI');
  t.datos = obtenerDatosTablero_();
  const html = t.evaluate().setTitle('Tablero de Vencimientos · Davivienda').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  SpreadsheetApp.getUi().showModalDialog(html, 'Vista Previa del Tablero');
}