
function cargarDatosLoad() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hojaTransform = libro.getSheetByName(CONFIG.HOJA_TRANSFORM);
  const hojaLoad = libro.getSheetByName(CONFIG.HOJA_LOAD) || libro.insertSheet(CONFIG.HOJA_LOAD);

  if (!hojaTransform) {
    throw new Error('No se encontró la hoja "' + CONFIG.HOJA_TRANSFORM + '". Corre primero transformarDatosETL().');
  }

  const mapaTransform = obtenerMapaColumnas(hojaTransform);
  const datosTransform = hojaTransform.getLastRow() > 1
    ? hojaTransform.getRange(2, 1, hojaTransform.getLastRow() - 1, hojaTransform.getLastColumn()).getValues()
    : [];

  if (hojaLoad.getLastRow() === 0) {
    hojaLoad.getRange(1, 1, 1, CONFIG.COL_LOAD.length).setValues([CONFIG.COL_LOAD]);
    hojaLoad.getRange(1, 1, 1, CONFIG.COL_LOAD.length).setFontWeight('bold');
  }
  const mapaLoad = obtenerMapaColumnas(hojaLoad);
  const datosLoadExistentes = hojaLoad.getLastRow() > 1
    ? hojaLoad.getRange(2, 1, hojaLoad.getLastRow() - 1, hojaLoad.getLastColumn()).getValues()
    : [];

  const controlPorId = {};
  datosLoadExistentes.forEach(fila => {
    const id = valorPorColumna(fila, mapaLoad, 'ID');
    if (!id) return;
    controlPorId[id] = {
      notificado: valorPorColumna(fila, mapaLoad, 'Fecha Notificado'),
      enProceso: valorPorColumna(fila, mapaLoad, 'Fecha En Proceso'),
      presentado: valorPorColumna(fila, mapaLoad, 'Fecha Presentado'),
      ultimoEnvio: valorPorColumna(fila, mapaLoad, 'Última Fecha Envío Recordatorio')
    };
  });

  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);

  const filasLoad = datosTransform.map(fila => {
    const id = valorPorColumna(fila, mapaTransform, 'ID');
    const fechaMaxima = valorPorColumna(fila, mapaTransform, 'Fecha máxima de presentación');
    const control = controlPorId[id] || {};

    let estado = CONFIG.ESTADOS.PENDIENTE;
    if (control.presentado) estado = CONFIG.ESTADOS.PRESENTADO;
    else if (control.enProceso) estado = CONFIG.ESTADOS.EN_PROCESO;
    else if (control.notificado) estado = CONFIG.ESTADOS.NOTIFICADO;

    const fechaMaximaDate = fechaMaxima instanceof Date ? fechaMaxima : null;
    const diasRestantes = fechaMaximaDate ? Math.round((fechaMaximaDate - hoy) / 86400000) : '';

    let extemporaneidad = '';
    if (control.presentado && fechaMaximaDate) {
      const fechaPres = control.presentado instanceof Date ? control.presentado : new Date(control.presentado);
      extemporaneidad = Math.round((fechaPres - fechaMaximaDate) / 86400000);
    }

    const semaforo = calcularSemaforo(estado, diasRestantes);
    const filaBase = CONFIG.COL_TRANSFORM.map(nombreCol => valorPorColumna(fila, mapaTransform, nombreCol));

    return filaBase.concat([
      control.notificado || '', control.enProceso || '', control.presentado || '',
      estado, semaforo, diasRestantes, extemporaneidad, control.ultimoEnvio || ''
    ]);
  });

  const filasAntiguas = hojaLoad.getLastRow();
  if (filasAntiguas > 1) {
    hojaLoad.getRange(2, 1, filasAntiguas - 1, hojaLoad.getLastColumn()).clearContent();
  }
  if (filasLoad.length > 0) {
    hojaLoad.getRange(2, 1, filasLoad.length, CONFIG.COL_LOAD.length).setValues(filasLoad);
    hojaLoad.getRange(2, 6, filasLoad.length, 1).setNumberFormat('dd/mm/yyyy');
    aplicarSemaforoHoja(hojaLoad);
  }

  Logger.log('LOAD completado: ' + filasLoad.length + ' filas.');
  return filasLoad.length;
}

function ejecutarProcesoCompletoETL() {
  transformarDatosETL();
  cargarDatosLoad();
  asegurarHojaParametros();
  SpreadsheetApp.getActiveSpreadsheet().toast('Proceso ETL completo (Extract → Transform → Load) terminado', 'Listo', 6);
}

function aplicarSemaforoHoja(hoja) {
  const mapa = obtenerMapaColumnas(hoja);
  const idxEstado = mapa['Estado Actual'];
  if (idxEstado === undefined || hoja.getLastRow() < 2) return;

  const rango = hoja.getRange(2, idxEstado + 1, hoja.getLastRow() - 1, 1);
  const valores = rango.getValues();
  const colores = valores.map(fila => [CONFIG.COLORES_ESTADO[fila[0]] || '#FFFFFF']);
  rango.setBackgrounds(colores);
}