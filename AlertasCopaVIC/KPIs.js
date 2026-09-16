/**
 * KPIs.js — Cálculo de indicadores sobre la hoja LOAD.
 *
 * Genera dos hojas:
 *   - KPIs           : registro histórico, UNA FILA POR CORRIDA (append, nunca borra).
 *                      Es la hoja tidy que se conecta a Looker Studio / Data Studio.
 *   - KPIs_TABLERO   : vista legible para el jefe (tarjetas + desgloses + comparativo
 *                      contra el control manual en Excel). Se reconstruye en cada corrida.
 *
 * Función principal:  calcularKPIs()
 */

const KPI_CONFIG = {
  HOJA_KPIS: 'KPIs',
  HOJA_TABLERO: 'KPIs_TABLERO',

  // Días que se consideran "próximo" para el KPI de anticipación.
  DIAS_PROXIMOS: 7,

  /**
   * Línea base medida sobre "Copia de Control Cumplimiento ICA Y RETEICA 2026.xlsx",
   * que era el control manual anterior. Se contaron las celdas de seguimiento
   * (columnas REVISIÓN y PRESENTACIÓN de cada mes) realmente diligenciadas.
   * Esta es la cifra contra la que comparamos la usabilidad del sistema nuevo.
   */
  BASELINE_MANUAL: {
    ARCHIVO: 'Copia de Control Cumplimiento ICA Y RETEICA 2026.xlsx',
    HOJAS: [
      { hoja: 'Consolidado',    filas: 380, celdas: 6080, llenas: 288 },
      { hoja: 'Retención ICA',  filas: 338, celdas: 3718, llenas: 273 },
      { hoja: 'Autorreteica',   filas: 156, celdas: 1560, llenas: 0 }
    ]
  }
};

/** Columnas del registro histórico. El orden manda: la fila se arma con este array. */
const KPI_COLUMNAS = [
  // --- Trazabilidad de la corrida ---
  'Fecha Corrida', 'Hora Corrida', 'Año', 'Mes', 'Semana ISO', 'Ejecutado Por', 'Hoja Origen',

  // --- Volumen ---
  'Total Obligaciones', 'Compañías Únicas', 'Impuestos Únicos', 'Municipios Únicos',
  'Encargados Únicos', 'Jefes Únicos',

  // --- Estado ---
  'Pendientes', 'Notificadas', 'En Proceso', 'Presentadas',
  '% Pendientes', '% Notificadas', '% En Proceso', '% Presentadas',

  // --- Vencimiento / riesgo ---
  'Vencidas Sin Presentar', 'Vencen Hoy', 'Críticas (<=2 días)',
  'Próximas 7 días', 'Vencen Este Mes', 'Sin Fecha Válida',
  '% Vencidas Sin Presentar', 'Días Restantes Promedio',

  // --- Cumplimiento ---
  'Ya Vencidas (total)', 'Presentadas A Tiempo', 'Presentadas Extemporáneas',
  '% Cumplimiento', '% Puntualidad', 'Días Extemporaneidad Promedio', 'Peor Extemporaneidad',
  'Días Anticipación Promedio',

  // --- Usabilidad / adopción (los 3 botones) ---
  'Obligaciones Notificadas Por Sistema', 'Con Respuesta Del Profesional',
  '% Tasa De Respuesta', 'Clics Botón Notificado', 'Clics Botón En Proceso',
  'Clics Botón Presentado', 'Total Interacciones',
  'Horas Promedio Reacción', 'Días Promedio Ciclo Notificado A Presentado',

  // --- Calidad de datos ---
  'Anomalías Detectadas', 'Sin Encargado Asignado', '% Completitud Datos',

  // --- Comparativo contra el control manual ---
  '% Diligenciamiento Manual (baseline)', '% Captura Automática', 'Mejora En Puntos'
];


/* ============================================================
   FUNCIÓN PRINCIPAL
   ============================================================ */

function calcularKPIs() {
  return calcularKPIsDesdeHoja(CONFIG.HOJA_LOAD, false);
}

/**
 * Versión de prueba: corre los KPIs sobre TRANSFORM2 SIN tocar las hojas reales.
 * Escribe en "KPIs_PRUEBA" y "KPIs_TABLERO_PRUEBA" en vez de "KPIs" y "KPIs_TABLERO",
 * igual al patrón que ya usan enviarCorreosDiariosPrueba() y enviarReportePruebaJefes().
 */
function calcularKPIsPrueba() {
  return calcularKPIsDesdeHoja('TRANSFORM2', true);
}

function calcularKPIsDesdeHoja(nombreHoja, esPrueba) {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = libro.getSheetByName(nombreHoja);
  if (!hoja) {
    throw new Error('No se encontró la hoja "' + nombreHoja + '". Corre primero ejecutarProcesoCompletoETL().');
  }

  const mapa = obtenerMapaColumnas(hoja);
  const datos = hoja.getLastRow() > 1
    ? hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getValues()
    : [];

  const m = kpiCalcularMetricas(datos, mapa, nombreHoja);

  const sufijo = esPrueba ? '_PRUEBA' : '';
  const hojaHistorico = KPI_CONFIG.HOJA_KPIS + sufijo;
  const hojaTablero = KPI_CONFIG.HOJA_TABLERO + sufijo;

  kpiEscribirHistorico(m, hojaHistorico);
 // kpiEscribirTablero(m, datos, mapa, hojaTablero);

  Logger.log('KPIs calculados sobre ' + nombreHoja + (esPrueba ? ' (PRUEBA, no afecta datos reales)' : '') +
             ': ' + m['Total Obligaciones'] + ' obligaciones, ' +
             m['% Cumplimiento'] + '% cumplimiento, ' + m['% Tasa De Respuesta'] + '% tasa de respuesta.');
  libro.toast(
    (esPrueba ? '[PRUEBA] ' : '') +
    m['Total Obligaciones'] + ' obligaciones · ' + m['% Cumplimiento'] + '% cumplimiento · ' +
    m['% Tasa De Respuesta'] + '% respuesta',
    esPrueba ? 'KPIs de prueba actualizados (' + hojaHistorico + ')' : 'KPIs actualizados', 8
  );
  return m;
}


/* ============================================================
   CÁLCULO DE MÉTRICAS
   ============================================================ */

function kpiCalcularMetricas(datos, mapa, nombreHoja) {
  const ahora = new Date();
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);

  const finMes = new Date(hoy.getFullYear(), hoy.getMonth() + 1, 0);
  const limiteProximos = new Date(hoy);
  limiteProximos.setDate(hoy.getDate() + KPI_CONFIG.DIAS_PROXIMOS);

  // Acumuladores
  const companias = {}, impuestos = {}, municipios = {}, encargados = {}, jefes = {};
  let pendientes = 0, notificadas = 0, enProceso = 0, presentadas = 0;
  let vencidasSinPresentar = 0, vencenHoy = 0, criticas = 0, proximas = 0, venceEsteMes = 0;
  let sinFechaValida = 0, sinEncargado = 0;
  let yaVencidas = 0, aTiempo = 0, extemporaneas = 0;
  let sumaExtemporaneidad = 0, contExtemporaneidad = 0, peorExtemporaneidad = 0;
  let sumaAnticipacion = 0, contAnticipacion = 0;
  let sumaDiasRestantes = 0, contDiasRestantes = 0;
  let notificadasPorSistema = 0, conRespuesta = 0;
  let clicNotificado = 0, clicEnProceso = 0, clicPresentado = 0;
  let sumaReaccionHoras = 0, contReaccion = 0;
  let sumaCiclo = 0, contCiclo = 0;
  let celdasEsperadas = 0, celdasLlenas = 0;

  datos.forEach(fila => {
    const id = valorPorColumna(fila, mapa, 'ID');
    if (!id) return;

    const estado       = valorPorColumna(fila, mapa, 'Estado Actual');
    const fechaMaxima  = kpiADate(valorPorColumna(fila, mapa, 'Fecha máxima de presentación'));
    const fechaValida  = valorPorColumna(fila, mapa, 'Fecha Válida');
    const fNotificado  = kpiADate(valorPorColumna(fila, mapa, 'Fecha Notificado'));
    const fEnProceso   = kpiADate(valorPorColumna(fila, mapa, 'Fecha En Proceso'));
    const fPresentado  = kpiADate(valorPorColumna(fila, mapa, 'Fecha Presentado'));
    const fEnvio       = kpiADate(valorPorColumna(fila, mapa, 'Última Fecha Envío Recordatorio'));
    const diasRest     = valorPorColumna(fila, mapa, 'Días Restantes');

    kpiContar(companias,  valorPorColumna(fila, mapa, 'Compañía (Normalizada)'));
    kpiContar(impuestos,  valorPorColumna(fila, mapa, 'Impuesto'));
    kpiContar(municipios, valorPorColumna(fila, mapa, 'Municipio'));
    kpiContarLista(encargados, valorPorColumna(fila, mapa, 'Encargado Email'));
    kpiContarLista(jefes, valorPorColumna(fila, mapa, 'Jefe1 Email'));
    kpiContarLista(jefes, valorPorColumna(fila, mapa, 'Jefe2 Email'));

    // --- Estado ---
    if (estado === CONFIG.ESTADOS.PRESENTADO)      presentadas++;
    else if (estado === CONFIG.ESTADOS.EN_PROCESO) enProceso++;
    else if (estado === CONFIG.ESTADOS.NOTIFICADO) notificadas++;
    else                                           pendientes++;

    // --- Calidad de datos ---
    if (fechaValida === false || !fechaMaxima) sinFechaValida++;
    if (!valorPorColumna(fila, mapa, 'Encargado Email')) sinEncargado++;
    celdasEsperadas += 3;
    if (fechaMaxima) celdasLlenas++;
    if (valorPorColumna(fila, mapa, 'Encargado Email')) celdasLlenas++;
    if (valorPorColumna(fila, mapa, 'Impuesto')) celdasLlenas++;

    // --- Vencimiento / riesgo ---
    if (typeof diasRest === 'number') {
      sumaDiasRestantes += diasRest;
      contDiasRestantes++;
    }
    if (fechaMaxima) {
      const dias = Math.round((fechaMaxima - hoy) / 86400000);
      const presentado = estado === CONFIG.ESTADOS.PRESENTADO;

      if (dias < 0 && !presentado) vencidasSinPresentar++;
      if (dias === 0 && !presentado) vencenHoy++;
      if (dias >= 0 && dias <= CONFIG.DIAS_UMBRAL_URGENTE && !presentado) criticas++;
      if (dias >= 0 && fechaMaxima <= limiteProximos && !presentado) proximas++;
      if (fechaMaxima >= hoy && fechaMaxima <= finMes) venceEsteMes++;

      // --- Cumplimiento: solo sobre las que ya vencieron ---
      if (fechaMaxima < hoy) {
        yaVencidas++;
        if (presentado) {
          const retraso = fPresentado ? Math.round((kpiSoloFecha(fPresentado) - fechaMaxima) / 86400000) : 0;
          if (retraso > 0) {
            extemporaneas++;
            sumaExtemporaneidad += retraso;
            contExtemporaneidad++;
            if (retraso > peorExtemporaneidad) peorExtemporaneidad = retraso;
          } else {
            aTiempo++;
            sumaAnticipacion += Math.abs(retraso);
            contAnticipacion++;
          }
        }
      }
    }

    // --- Usabilidad: los 3 botones ---
    if (fEnvio) notificadasPorSistema++;
    if (fNotificado) clicNotificado++;
    if (fEnProceso)  clicEnProceso++;
    if (fPresentado) clicPresentado++;
    if (fNotificado || fEnProceso || fPresentado) {
      if (fEnvio) conRespuesta++;
    }

    // Tiempo de reacción: del correo enviado al primer clic.
    const primerClic = kpiMenorFecha([fNotificado, fEnProceso, fPresentado]);
    if (fEnvio && primerClic && primerClic >= fEnvio) {
      sumaReaccionHoras += (primerClic - fEnvio) / 3600000;
      contReaccion++;
    }

    // Ciclo completo: del clic "Notificado" al clic "Presentado".
    if (fNotificado && fPresentado && fPresentado >= fNotificado) {
      sumaCiclo += (fPresentado - fNotificado) / 86400000;
      contCiclo++;
    }
  });

  const total = datos.filter(f => valorPorColumna(f, mapa, 'ID')).length;
  const totalInteracciones = clicNotificado + clicEnProceso + clicPresentado;

  // Captura automática: de todo lo que el sistema notificó, cuánto quedó registrado.
  const capturaAutomatica = kpiPorcentaje(conRespuesta, notificadasPorSistema);
  const baseline = kpiBaselineManual();

  const anomalias = kpiContarAnomalias();

  const metricas = {
    'Fecha Corrida': Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    'Hora Corrida': Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'HH:mm:ss'),
    'Año': ahora.getFullYear(),
    'Mes': ahora.getMonth() + 1,
    'Semana ISO': kpiSemanaISO(ahora),
    'Ejecutado Por': kpiUsuarioActual(),
    'Hoja Origen': nombreHoja,

    'Total Obligaciones': total,
    'Compañías Únicas': Object.keys(companias).length,
    'Impuestos Únicos': Object.keys(impuestos).length,
    'Municipios Únicos': Object.keys(municipios).length,
    'Encargados Únicos': Object.keys(encargados).length,
    'Jefes Únicos': Object.keys(jefes).length,

    'Pendientes': pendientes,
    'Notificadas': notificadas,
    'En Proceso': enProceso,
    'Presentadas': presentadas,
    '% Pendientes': kpiPorcentaje(pendientes, total),
    '% Notificadas': kpiPorcentaje(notificadas, total),
    '% En Proceso': kpiPorcentaje(enProceso, total),
    '% Presentadas': kpiPorcentaje(presentadas, total),

    'Vencidas Sin Presentar': vencidasSinPresentar,
    'Vencen Hoy': vencenHoy,
    'Críticas (<=2 días)': criticas,
    'Próximas 7 días': proximas,
    'Vencen Este Mes': venceEsteMes,
    'Sin Fecha Válida': sinFechaValida,
    '% Vencidas Sin Presentar': kpiPorcentaje(vencidasSinPresentar, total),
    'Días Restantes Promedio': kpiPromedio(sumaDiasRestantes, contDiasRestantes),

    'Ya Vencidas (total)': yaVencidas,
    'Presentadas A Tiempo': aTiempo,
    'Presentadas Extemporáneas': extemporaneas,
    '% Cumplimiento': kpiPorcentaje(aTiempo + extemporaneas, yaVencidas),
    '% Puntualidad': kpiPorcentaje(aTiempo, aTiempo + extemporaneas),
    'Días Extemporaneidad Promedio': kpiPromedio(sumaExtemporaneidad, contExtemporaneidad),
    'Peor Extemporaneidad': peorExtemporaneidad,
    'Días Anticipación Promedio': kpiPromedio(sumaAnticipacion, contAnticipacion),

    'Obligaciones Notificadas Por Sistema': notificadasPorSistema,
    'Con Respuesta Del Profesional': conRespuesta,
    '% Tasa De Respuesta': capturaAutomatica,
    'Clics Botón Notificado': clicNotificado,
    'Clics Botón En Proceso': clicEnProceso,
    'Clics Botón Presentado': clicPresentado,
    'Total Interacciones': totalInteracciones,
    'Horas Promedio Reacción': kpiPromedio(sumaReaccionHoras, contReaccion),
    'Días Promedio Ciclo Notificado A Presentado': kpiPromedio(sumaCiclo, contCiclo),

    'Anomalías Detectadas': anomalias,
    'Sin Encargado Asignado': sinEncargado,
    '% Completitud Datos': kpiPorcentaje(celdasLlenas, celdasEsperadas),

    '% Diligenciamiento Manual (baseline)': baseline.porcentaje,
    '% Captura Automática': capturaAutomatica,
    'Mejora En Puntos': Math.round((capturaAutomatica - baseline.porcentaje) * 10) / 10
  };

  // Datos auxiliares que usa el tablero pero que no van al histórico.
  metricas._desgloses = {
    companias: companias,
    impuestos: impuestos,
    municipios: municipios,
    encargados: encargados,
    baseline: baseline
  };

  return metricas;
}


/* ============================================================
   HOJA "KPIs" — registro histórico append-only
   ============================================================ */

function kpiEscribirHistorico(metricas, nombreHojaDestino) {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(nombreHojaDestino);

  if (!hoja) {
    hoja = libro.insertSheet(nombreHojaDestino);
  }

  // Encabezado: se escribe solo si la hoja está vacía. Nunca se borra el histórico.
  if (hoja.getLastRow() === 0) {
    hoja.getRange(1, 1, 1, KPI_COLUMNAS.length).setValues([KPI_COLUMNAS]);
    hoja.getRange(1, 1, 1, KPI_COLUMNAS.length)
        .setFontWeight('bold')
        .setBackground('#FDE1E0')
        .setVerticalAlignment('middle')
        .setWrap(true);
    hoja.setFrozenRows(1);
    hoja.setFrozenColumns(1);
  }

  const fila = KPI_COLUMNAS.map(col => {
    const v = metricas[col];
    return (v === undefined || v === null) ? '' : v;
  });

  hoja.appendRow(fila);
  hoja.autoResizeColumns(1, Math.min(KPI_COLUMNAS.length, 10));
}


/* ============================================================
   HOJA "KPIs_TABLERO" — vista para el jefe
   ============================================================ */

function kpiEscribirTablero(m, datos, mapa, nombreHojaDestino) {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(nombreHojaDestino);
  if (!hoja) hoja = libro.insertSheet(nombreHojaDestino);

  hoja.clear();
  const d = m._desgloses;
  const filas = [];

  filas.push(['TABLERO DE INDICADORES — VENCIMIENTOS DIAN', '', '']);
  filas.push(['Generado', m['Fecha Corrida'] + ' ' + m['Hora Corrida'], 'Hoja origen: ' + m['Hoja Origen']]);
  filas.push(['', '', '']);

  kpiBloque(filas, 'RESUMEN GENERAL', [
    ['Total obligaciones', m['Total Obligaciones'], 'Registros activos en LOAD'],
    ['Compañías', m['Compañías Únicas'], ''],
    ['Impuestos distintos', m['Impuestos Únicos'], ''],
    ['Municipios', m['Municipios Únicos'], ''],
    ['Profesionales encargados', m['Encargados Únicos'], ''],
    ['Jefes involucrados', m['Jefes Únicos'], '']
  ]);

  kpiBloque(filas, 'ESTADO ACTUAL', [
    ['Pendientes', m['Pendientes'], m['% Pendientes'] + '%'],
    ['Notificadas', m['Notificadas'], m['% Notificadas'] + '%'],
    ['En proceso', m['En Proceso'], m['% En Proceso'] + '%'],
    ['Presentadas', m['Presentadas'], m['% Presentadas'] + '%']
  ]);

  kpiBloque(filas, 'RIESGO / VENCIMIENTOS', [
    ['Vencidas sin presentar', m['Vencidas Sin Presentar'], 'ALERTA — requiere gestión inmediata'],
    ['Vencen hoy', m['Vencen Hoy'], ''],
    ['Críticas (<= ' + CONFIG.DIAS_UMBRAL_URGENTE + ' días)', m['Críticas (<=2 días)'], ''],
    ['Próximas ' + KPI_CONFIG.DIAS_PROXIMOS + ' días', m['Próximas 7 días'], ''],
    ['Vencen este mes', m['Vencen Este Mes'], ''],
    ['Sin fecha válida', m['Sin Fecha Válida'], 'Revisar LOG_ANOMALIAS'],
    ['Días restantes promedio', m['Días Restantes Promedio'], '']
  ]);

  kpiBloque(filas, 'CUMPLIMIENTO', [
    ['Obligaciones ya vencidas', m['Ya Vencidas (total)'], 'Base de cálculo del cumplimiento'],
    ['Presentadas a tiempo', m['Presentadas A Tiempo'], ''],
    ['Presentadas extemporáneas', m['Presentadas Extemporáneas'], ''],
    ['% Cumplimiento', m['% Cumplimiento'] + '%', 'Presentadas / ya vencidas'],
    ['% Puntualidad', m['% Puntualidad'] + '%', 'A tiempo / presentadas'],
    ['Días de extemporaneidad (prom.)', m['Días Extemporaneidad Promedio'], ''],
    ['Peor extemporaneidad', m['Peor Extemporaneidad'], 'días de retraso'],
    ['Días de anticipación (prom.)', m['Días Anticipación Promedio'], 'Se presenta X días antes del límite']
  ]);

  kpiBloque(filas, 'USABILIDAD DEL SISTEMA (LOS 3 BOTONES)', [
    ['Obligaciones notificadas por el sistema', m['Obligaciones Notificadas Por Sistema'], 'Correos con recordatorio enviado'],
    ['Con respuesta del profesional', m['Con Respuesta Del Profesional'], 'Marcó al menos un botón'],
    ['% Tasa de respuesta', m['% Tasa De Respuesta'] + '%', 'KPI CLAVE de adopción'],
    ['Clics en "Notificado"', m['Clics Botón Notificado'], ''],
    ['Clics en "En proceso"', m['Clics Botón En Proceso'], ''],
    ['Clics en "Presentado"', m['Clics Botón Presentado'], ''],
    ['Total interacciones registradas', m['Total Interacciones'], 'Antes esto se llenaba a mano'],
    ['Horas promedio de reacción', m['Horas Promedio Reacción'], 'Del correo al primer clic'],
    ['Días promedio de ciclo', m['Días Promedio Ciclo Notificado A Presentado'], 'De "Notificado" a "Presentado"']
  ]);

  // --- Comparativo contra el Excel manual ---
  filas.push(['', '', '']);
  filas.push(['COMPARATIVO: CONTROL MANUAL (EXCEL) vs. SISTEMA AUTOMATIZADO', '', '']);
  filas.push(['Fuente del baseline', KPI_CONFIG.BASELINE_MANUAL.ARCHIVO, '']);
  filas.push(['Hoja del Excel', 'Celdas de seguimiento', '% diligenciado a mano']);
  d.baseline.detalle.forEach(x => {
    filas.push([x.hoja, x.llenas + ' / ' + x.celdas, x.porcentaje + '%']);
  });
  filas.push(['TOTAL CONTROL MANUAL', d.baseline.llenas + ' / ' + d.baseline.celdas, d.baseline.porcentaje + '%']);
  filas.push(['TOTAL SISTEMA AUTOMATIZADO',
              m['Con Respuesta Del Profesional'] + ' / ' + m['Obligaciones Notificadas Por Sistema'],
              m['% Captura Automática'] + '%']);
  filas.push(['MEJORA', m['Mejora En Puntos'] + ' puntos porcentuales', '']);

  kpiBloque(filas, 'CALIDAD DE DATOS', [
    ['Anomalías detectadas', m['Anomalías Detectadas'], 'Hoja LOG_ANOMALIAS'],
    ['Sin encargado asignado', m['Sin Encargado Asignado'], ''],
    ['% Completitud de datos', m['% Completitud Datos'] + '%', '']
  ]);

  kpiBloqueTop(filas, 'TOP 10 COMPAÑÍAS CON MÁS OBLIGACIONES', d.companias, 10);
  kpiBloqueTop(filas, 'OBLIGACIONES POR IMPUESTO', d.impuestos, 20);
  kpiBloqueTop(filas, 'TOP 10 MUNICIPIOS', d.municipios, 10);
  kpiBloqueTop(filas, 'CARGA POR PROFESIONAL', d.encargados, 30);

  hoja.getRange(1, 1, filas.length, 3).setValues(filas);
  kpiFormatearTablero(hoja, filas);
}

function kpiBloque(filas, titulo, items) {
  filas.push(['', '', '']);
  filas.push([titulo, '', '']);
  items.forEach(i => filas.push([i[0], i[1], i[2] || '']));
}

function kpiBloqueTop(filas, titulo, conteo, limite) {
  filas.push(['', '', '']);
  filas.push([titulo, '', '']);
  const orden = Object.keys(conteo)
    .map(k => [k, conteo[k]])
    .sort((a, b) => b[1] - a[1])
    .slice(0, limite);
  const total = Object.keys(conteo).reduce((s, k) => s + conteo[k], 0);
  if (orden.length === 0) {
    filas.push(['(sin datos)', '', '']);
    return;
  }
  orden.forEach(x => filas.push([x[0], x[1], kpiPorcentaje(x[1], total) + '%']));
}

function kpiFormatearTablero(hoja, filas) {
  hoja.setColumnWidth(1, 330);
  hoja.setColumnWidth(2, 170);
  hoja.setColumnWidth(3, 340);

  // Título principal
  hoja.getRange(1, 1, 1, 3).merge()
      .setFontSize(15).setFontWeight('bold')
      .setBackground('#FDE1E0').setHorizontalAlignment('center');

  // Encabezados de bloque: fila cuya col A tiene texto y col B/C vacías, y no es una fila en blanco.
  filas.forEach((f, i) => {
    const esTitulo = f[0] && f[1] === '' && f[2] === '' && i > 0;
    if (esTitulo) {
      hoja.getRange(i + 1, 1, 1, 3).merge()
          .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#27729D');
    }
  });

  hoja.getRange(2, 1, filas.length, 3).setVerticalAlignment('middle');
  hoja.setFrozenRows(2);
}


/* ============================================================
   UTILIDADES
   ============================================================ */

function kpiBaselineManual() {
  const hojas = KPI_CONFIG.BASELINE_MANUAL.HOJAS;
  let celdas = 0, llenas = 0;
  const detalle = hojas.map(h => {
    celdas += h.celdas;
    llenas += h.llenas;
    return { hoja: h.hoja, celdas: h.celdas, llenas: h.llenas, porcentaje: kpiPorcentaje(h.llenas, h.celdas) };
  });
  return { celdas: celdas, llenas: llenas, porcentaje: kpiPorcentaje(llenas, celdas), detalle: detalle };
}

function kpiContarAnomalias() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_ANOMALIAS);
  if (!hoja) return 0;
  return Math.max(hoja.getLastRow() - 1, 0);
}

function kpiContar(acumulador, valor) {
  const clave = String(valor || '').trim();
  if (!clave) return;
  acumulador[clave] = (acumulador[clave] || 0) + 1;
}

/** Para campos con varios valores separados por ";" (encargados, jefes). */
function kpiContarLista(acumulador, valor) {
  if (!valor) return;
  String(valor).split(';').map(v => v.trim()).filter(Boolean)
    .forEach(v => { acumulador[v] = (acumulador[v] || 0) + 1; });
}

function kpiADate(valor) {
  if (!valor) return null;
  if (valor instanceof Date) return isNaN(valor.getTime()) ? null : valor;
  const d = new Date(valor);
  return isNaN(d.getTime()) ? null : d;
}

function kpiSoloFecha(fecha) {
  const d = new Date(fecha);
  d.setHours(0, 0, 0, 0);
  return d;
}

function kpiMenorFecha(lista) {
  const validas = lista.filter(f => f instanceof Date);
  if (validas.length === 0) return null;
  return new Date(Math.min.apply(null, validas.map(f => f.getTime())));
}

function kpiPorcentaje(parte, total) {
  if (!total) return 0;
  return Math.round((parte / total) * 1000) / 10;
}

function kpiPromedio(suma, cantidad) {
  if (!cantidad) return 0;
  return Math.round((suma / cantidad) * 10) / 10;
}

function kpiSemanaISO(fecha) {
  const d = new Date(Date.UTC(fecha.getFullYear(), fecha.getMonth(), fecha.getDate()));
  const diaSemana = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - diaSemana);
  const inicioAnio = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - inicioAnio) / 86400000) + 1) / 7);
}

function kpiUsuarioActual() {
  try {
    return Session.getActiveUser().getEmail() || '(automático)';
  } catch (e) {
    return '(automático)';
  }
}


/* ============================================================
   INTEGRACIÓN
   ============================================================ */

/** ETL completo + KPIs en un solo clic. */
function ejecutarProcesoCompletoConKPIs() {
  transformarDatosETL();
  cargarDatosLoad();
  asegurarHojaParametros();
  calcularKPIs();
  SpreadsheetApp.getActiveSpreadsheet().toast('ETL + KPIs completados', 'Listo', 6);
}

/** Menú propio en la hoja de cálculo. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Vencimientos DIAN')
    .addItem('Ejecutar ETL completo', 'ejecutarProcesoCompletoETL')
    .addItem('Calcular KPIs', 'calcularKPIs')
    .addItem('ETL + KPIs', 'ejecutarProcesoCompletoConKPIs')
    .addSeparator()
    .addItem('Calcular KPIs (prueba TRANSFORM2)', 'calcularKPIsPrueba')
    .addSeparator()
    .addItem('Activar trigger KPIs cada 5 min', 'crearTriggerKPIsCada5Min')
    .addItem('Desactivar trigger KPIs', 'eliminarTriggerKPIs')
    .addToUi();
}


/* ============================================================
   TRIGGER — calcularKPIs() cada 5 minutos
   ============================================================ */

/**
 * Instala un trigger de tiempo que corre calcularKPIs() cada 5 minutos.
 * Es idempotente: primero borra cualquier trigger anterior apuntando a
 * 'calcularKPIs' para que nunca queden dos triggers disparando en paralelo
 * (eso duplicaría cada fila del histórico).
 *
 * Nota de volumen: a cada 5 min esto agrega ~288 filas/día a la hoja "KPIs".
 * Es una decisión intencional para tener el histórico lo más fino posible;
 * si más adelante pesa mucho, se puede filtrar o resumir por día en un reporte aparte.
 */
function crearTriggerKPIsCada5Min() {
  eliminarTriggerKPIs();
  ScriptApp.newTrigger('calcularKPIs')
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log('Trigger creado: calcularKPIs() cada 5 minutos.');
  SpreadsheetApp.getActiveSpreadsheet().toast('Trigger activado: KPIs se calculará cada 5 minutos.', 'Listo', 6);
}

/** Quita cualquier trigger existente que apunte a calcularKPIs (evita duplicados). */
function eliminarTriggerKPIs() {
  const triggers = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'calcularKPIs');
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  if (triggers.length > 0) {
    Logger.log('Se eliminaron ' + triggers.length + ' trigger(s) previo(s) de calcularKPIs.');
  }
  return triggers.length;
}

/** Utilidad de diagnóstico: lista todos los triggers activos del proyecto. */
function listarTriggersActivos() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    Logger.log(t.getHandlerFunction() + ' — ' + t.getEventType() + ' — ' + t.getTriggerSourceId());
  });
  return triggers.length;
}
