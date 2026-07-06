function transformarDatosETL() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hojaExtract = libro.getSheetByName(CONFIG.HOJA_EXTRACT);
  const hojaTransform = libro.getSheetByName(CONFIG.HOJA_TRANSFORM) || libro.insertSheet(CONFIG.HOJA_TRANSFORM);

  if (!hojaExtract) {
    throw new Error('No se encontró la hoja "' + CONFIG.HOJA_EXTRACT + '". Revisa el nombre exacto.');
  }

  const anioReferencia = CONFIG.ANIO_REFERENCIA || new Date().getFullYear();
  const anomalias = [];

  const mapaExtract = obtenerMapaColumnas(hojaExtract);
  const columnasObligatorias = [
    CONFIG.COL_EXTRACT.COMPANIA, CONFIG.COL_EXTRACT.NIT, CONFIG.COL_EXTRACT.IMPUESTO,
    CONFIG.COL_EXTRACT.FECHA_MAXIMA, CONFIG.COL_EXTRACT.ENCARGADO
  ];
  columnasObligatorias.forEach(nombre => {
    if (mapaExtract[nombre] === undefined) {
      throw new Error('Falta la columna "' + nombre + '" en EXTRACT. Encabezados encontrados: ' + Object.keys(mapaExtract).join(', '));
    }
  });

  const datosExtract = hojaExtract.getRange(2, 1, Math.max(hojaExtract.getLastRow() - 1, 0), hojaExtract.getLastColumn()).getValues();
  const mapaCanonicoNIT = construirMapaCanonicoNIT(datosExtract, mapaExtract);

  if (hojaTransform.getLastRow() === 0) {
    hojaTransform.getRange(1, 1, 1, CONFIG.COL_TRANSFORM.length).setValues([CONFIG.COL_TRANSFORM]);
    hojaTransform.getRange(1, 1, 1, CONFIG.COL_TRANSFORM.length).setFontWeight('bold');
  }
  const mapaTransform = obtenerMapaColumnas(hojaTransform);
  const filasTransformExistentes = hojaTransform.getLastRow() > 1
    ? hojaTransform.getRange(2, 1, hojaTransform.getLastRow() - 1, hojaTransform.getLastColumn()).getValues()
    : [];

  const controlPorId = {};
  filasTransformExistentes.forEach(fila => {
    const id = valorPorColumna(fila, mapaTransform, 'ID');
    if (!id) return;
    controlPorId[id] = {
      notificado: valorPorColumna(fila, mapaTransform, 'Fecha Notificado'),
      enProceso: valorPorColumna(fila, mapaTransform, 'Fecha En Proceso'),
      presentado: valorPorColumna(fila, mapaTransform, 'Fecha Presentado'),
      ultimoEnvio: valorPorColumna(fila, mapaTransform, 'Última Fecha Envío Recordatorio')
    };
  });

  const conteoFilasExactas = {};
  datosExtract.forEach(fila => {
    const clave = fila.join('|').toUpperCase();
    conteoFilasExactas[clave] = (conteoFilasExactas[clave] || 0) + 1;
  });

  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);

  const filasNuevas = [];
  const idsYaIncluidos = new Set();
  datosExtract.forEach((fila, idx) => {
    const numeroFilaReal = idx + 2; 
    const companiaOriginal = valorPorColumna(fila, mapaExtract, CONFIG.COL_EXTRACT.COMPANIA);
    if (!companiaOriginal) return;

    const nit = String(valorPorColumna(fila, mapaExtract, CONFIG.COL_EXTRACT.NIT) || '').trim();
    const impuesto = String(valorPorColumna(fila, mapaExtract, CONFIG.COL_EXTRACT.IMPUESTO) || '').trim();
    const companiaNormalizada = mapaCanonicoNIT[nit] || limpiarNombreCompania(companiaOriginal);

    const fechaMaximaRaw = valorPorColumna(fila, mapaExtract, CONFIG.COL_EXTRACT.FECHA_MAXIMA);
    const infoFecha = parsearFecha(fechaMaximaRaw, anioReferencia);
    if (!infoFecha.valida) {
      anomalias.push([numeroFilaReal, companiaOriginal, nit, impuesto, infoFecha.motivo, String(fechaMaximaRaw)]);
    }

    const encargadoContactos = extraerContactos(valorPorColumna(fila, mapaExtract, CONFIG.COL_EXTRACT.ENCARGADO));
    const jefe1Contactos = extraerContactos(valorPorColumna(fila, mapaExtract, CONFIG.COL_EXTRACT.JEFE1));
    const jefe2Contactos = extraerContactos(valorPorColumna(fila, mapaExtract, CONFIG.COL_EXTRACT.JEFE2));

    if (encargadoContactos.length === 0) {
      anomalias.push([numeroFilaReal, companiaOriginal, nit, impuesto, 'Sin persona encargada asignada', String(valorPorColumna(fila, mapaExtract, CONFIG.COL_EXTRACT.ENCARGADO))]);
    }

    const encargadoStr = contactosAString(encargadoContactos);
    const jefe1Str = contactosAString(jefe1Contactos);
    const jefe2Str = contactosAString(jefe2Contactos);

     const requiereMunicipio = CONFIG.IMPUESTOS_CON_MUNICIPIO.includes(normalizarTexto(impuesto));
    const municipio = requiereMunicipio ? String(valorPorColumna(fila, mapaExtract, CONFIG.COL_EXTRACT.MUNICIPIO) || '').trim() : '';
    if (requiereMunicipio && !municipio) {
      anomalias.push([numeroFilaReal, companiaOriginal, nit, impuesto, 'Impuesto requiere Municipio pero viene vacío', '']);
    }

    const cifras = valorPorColumna(fila, mapaExtract, CONFIG.COL_EXTRACT.CIFRAS) || '';

    const claveExacta = fila.join('|').toUpperCase();
    if (conteoFilasExactas[claveExacta] > 1) {
      anomalias.push([numeroFilaReal, companiaOriginal, nit, impuesto, 'Fila duplicada exactamente en EXTRACT (aparece ' + conteoFilasExactas[claveExacta] + ' veces)', '']);
    }

    const fechaParaId = infoFecha.fecha ? infoFecha.fecha.toISOString() : String(fechaMaximaRaw);
    const id = generarIdObligacion(companiaNormalizada, nit, impuesto, fechaParaId, municipio);
    const control = controlPorId[id] || {};

    let estado = CONFIG.ESTADOS.PENDIENTE;
    if (control.presentado) estado = CONFIG.ESTADOS.PRESENTADO;
    else if (control.enProceso) estado = CONFIG.ESTADOS.EN_PROCESO;
    else if (control.notificado) estado = CONFIG.ESTADOS.NOTIFICADO;

    const diasRestantes = infoFecha.fecha ? Math.round((infoFecha.fecha - hoy) / 86400000) : '';

    let extemporaneidad = '';
    if (control.presentado && infoFecha.fecha) {
      const fechaPres = control.presentado instanceof Date ? control.presentado : new Date(control.presentado);
      extemporaneidad = Math.round((fechaPres - infoFecha.fecha) / 86400000);
    }

// Evita filas duplicadas en TRANSFORM: si el ID ya se usó en esta
    // misma corrida, se omite (ya quedó registrada como anomalía arriba).
    if (idsYaIncluidos.has(id)) return;
    idsYaIncluidos.add(id);

    filasNuevas.push([
      id,
      companiaNormalizada, companiaOriginal, nit, impuesto,
      infoFecha.fecha || '', infoFecha.valida,
      encargadoStr.nombres, encargadoStr.emails,
      jefe1Str.nombres, jefe1Str.emails,
      jefe2Str.nombres, jefe2Str.emails,
      cifras, municipio,
      control.notificado || '', control.enProceso || '', control.presentado || '',
      estado, diasRestantes, extemporaneidad, control.ultimoEnvio || ''
    ]);
  });

  const filasAntiguas = hojaTransform.getLastRow();
  if (filasAntiguas > 1) {
    hojaTransform.getRange(2, 1, filasAntiguas - 1, hojaTransform.getLastColumn()).clearContent();
  }
  if (filasNuevas.length > 0) {
    hojaTransform.getRange(2, 1, filasNuevas.length, CONFIG.COL_TRANSFORM.length).setValues(filasNuevas);
    hojaTransform.getRange(2, 6, filasNuevas.length, 1).setNumberFormat('dd/mm/yyyy'); // Fecha máxima
  }

  escribirAnomalias(anomalias);
  asegurarHojaParametros();

  Logger.log('ETL completado: ' + filasNuevas.length + ' obligaciones procesadas, ' + anomalias.length + ' anomalías detectadas.');
  SpreadsheetApp.getActiveSpreadsheet().toast(
    filasNuevas.length + ' filas procesadas, ' + anomalias.length + ' anomalías en LOG_ANOMALIAS',
    'ETL completado', 8
  );
  return { procesadas: filasNuevas.length, anomalias: anomalias.length };
}