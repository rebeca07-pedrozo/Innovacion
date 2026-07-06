const PALETA_REPORTE = {
  CARD_TOTAL_BG: '#DFEAF0', CARD_TOTAL_TEXTO: '#27729D',
  CARD_PROXIMAS_BG: '#FDF0DF', CARD_PROXIMAS_TEXTO: '#F4971D',
  CARD_VENCIDAS_BG: '#FDE1E0', CARD_VENCIDAS_TEXTO: '#F0342D',
  BARRA_ROJA: '#EF2C25', BARRA_AZUL: '#4A7FF7', BARRA_NARANJA: '#F3B86B', BARRA_AMARILLA: '#FFDE59'
};

function enviarReportePruebaJefes() {
  return enviarReportesJefes('TRANSFORM2');
}

function enviarReportesJefesProduccion() {
  return enviarReportesJefes(CONFIG.HOJA_LOAD);
}

function enviarReportesJefes(nombreHoja) {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = libro.getSheetByName(nombreHoja);
  if (!hoja) throw new Error('No se encontró la hoja "' + nombreHoja + '".');

  const mapa = obtenerMapaColumnas(hoja);
  const datos = hoja.getRange(2, 1, Math.max(hoja.getLastRow() - 1, 0), hoja.getLastColumn()).getValues();
  const rango = obtenerRangoSemanaActual();

  const grupos = {};

  datos.forEach((fila, idx) => {
    const fechaMaxima = valorPorColumna(fila, mapa, 'Fecha máxima de presentación');
    const estado = valorPorColumna(fila, mapa, 'Estado Actual');
    const dentroDeSemana = fechaMaxima instanceof Date && fechaMaxima >= rango.inicio && fechaMaxima <= rango.fin;
    const yaVencida = fechaMaxima instanceof Date && fechaMaxima < rango.inicio && estado !== CONFIG.ESTADOS.PRESENTADO;
    if (!dentroDeSemana && !yaVencida) return;

    const item = { numeroFila: idx + 2, valores: fila };

    ['Jefe1', 'Jefe2'].forEach(prefijo => {
      const emailCampo = valorPorColumna(fila, mapa, prefijo + ' Email');
      const nombreCampo = valorPorColumna(fila, mapa, prefijo + ' Nombre');
      if (!emailCampo) return;
      const emails = String(emailCampo).split(';').map(e => e.trim()).filter(Boolean);
      const nombres = String(nombreCampo || '').split(';').map(n => n.trim());
      emails.forEach((correo, i) => {
        if (!grupos[correo]) grupos[correo] = { nombre: nombres[i] || nombres[0] || correo, filas: [] };
        grupos[correo].filas.push(item);
      });
    });
  });

  let enviados = 0;
  Object.keys(grupos).forEach(correoJefe => {
    const grupo = grupos[correoJefe];
    const pdfBlob = construirPdfReporteJefe(grupo.nombre, grupo.filas, mapa, rango);

    GmailApp.sendEmail(correoJefe,
      'Informe semanal de vencimientos DIAN - Semana del ' + rango.texto,
      'Hola ' + grupo.nombre + ',\n\nAdjunto el informe semanal de vencimientos ante la DIAN de tu equipo.\n\nSaludos.',
      { attachments: [pdfBlob], name: 'Sistema de Vencimientos DIAN' }
    );
    enviados++;
  });

  Logger.log('Reportes PDF enviados: ' + enviados + ' (hoja: ' + nombreHoja + ')');
  SpreadsheetApp.getActiveSpreadsheet().toast(enviados + ' reportes PDF enviados', 'Reportes Jefes', 6);
  return enviados;
}

function etiquetaEstadoParaJefe(estado) {
  if (estado === CONFIG.ESTADOS.PENDIENTE) return CONFIG.ESTADOS.NOTIFICADO;
  return estado;
}

function construirPdfReporteJefe(nombreJefe, filas, mapa, rango) {
  const total = filas.length;
  const vencidas = filas.filter(item => {
    const f = valorPorColumna(item.valores, mapa, 'Fecha máxima de presentación');
    const estado = valorPorColumna(item.valores, mapa, 'Estado Actual');
    return f instanceof Date && f < rango.inicio && estado !== CONFIG.ESTADOS.PRESENTADO;
  }).length;
  const proximas = total - vencidas;

  const conteoCompanias = {};
  filas.forEach(item => {
    const c = valorPorColumna(item.valores, mapa, 'Compañía (Normalizada)');
    conteoCompanias[c] = (conteoCompanias[c] || 0) + 1;
  });
  const topCompanias = Object.entries(conteoCompanias).sort((a, b) => b[1] - a[1]).slice(0, 5);

  const conteoImpuestos = {};
  filas.forEach(item => {
    const imp = valorPorColumna(item.valores, mapa, 'Impuesto');
    conteoImpuestos[imp] = (conteoImpuestos[imp] || 0) + 1;
  });
  const topImpuestos = Object.entries(conteoImpuestos).sort((a, b) => b[1] - a[1]).slice(0, 5);

  // Paleta de colores vivos, uno distinto por barra, igual que la imagen de referencia
  const paletaBarras = [PALETA_REPORTE.BARRA_ROJA, PALETA_REPORTE.BARRA_AZUL, PALETA_REPORTE.BARRA_NARANJA, PALETA_REPORTE.BARRA_AMARILLA, PALETA_REPORTE.BARRA_AZUL];

  const chartEntidades = generarGraficoBarrasHorizontal(
    topCompanias.map(c => c[0]), topCompanias.map(c => c[1]), 'Top de entidades con más vencimientos', paletaBarras
  );
  const chartImpuestos = generarGraficoColumnas(
    topImpuestos.map(c => c[0]), topImpuestos.map(c => c[1]), 'Top de obligaciones a vencer', PALETA_REPORTE.BARRA_AZUL
  );
  const chartDona = generarGraficoDona(vencidas, proximas, total);

  const filasTablaHtml = filas.map(item => construirFilaDetalleHtml(item, mapa)).join('');

  const html = `
  <html><body style="font-family:Arial, sans-serif; margin:0; padding:0; background:#ffffff;">
    <table style="width:100%; background:linear-gradient(90deg, #FDF3EE, #FDE1E0); border-collapse:collapse;">
      <tr>
        <td style="padding:14px 24px;"><img src="data:image/png;base64,${CONFIG.LOGO_BASE64}" style="height:26px;"></td>
        <td style="padding:14px 24px; text-align:right;">
          <img src="${CONFIG.URL_ICONO_CALENDARIO}" style="height:14px; vertical-align:middle; margin-right:5px;">
          <span style="color:#555; font-size:12px;">Semana del ${rango.texto}</span>
        </td>
      </tr>
    </table>
    <div style="padding:24px;">
      <h1 style="font-size:22px; margin:0 0 14px 0; color:#222;">Informe de resumen vencimientos semanales</h1>
      <p style="border-left:4px solid #EF2C25; padding-left:10px; margin-bottom:22px; font-size:13px; color:#333;">
        Resumen del estado de las obligaciones tributarias de tu equipo para esta semana.
      </p>

      <table style="width:100%; border-collapse:separate; border-spacing:10px 0; margin-bottom:26px;">
        <tr>
          <td style="width:33%; background:${PALETA_REPORTE.CARD_TOTAL_BG}; border-radius:10px; padding:16px;">
            <div style="color:${PALETA_REPORTE.CARD_TOTAL_TEXTO}; font-weight:bold; font-size:13px;">Total obligaciones</div>
            <div style="color:#222; font-size:26px; font-weight:bold;">${total}</div>
          </td>
          <td style="width:33%; background:${PALETA_REPORTE.CARD_PROXIMAS_BG}; border-radius:10px; padding:16px;">
            <div style="color:${PALETA_REPORTE.CARD_PROXIMAS_TEXTO}; font-weight:bold; font-size:13px;">Próximas a vencer</div>
            <div style="color:#222; font-size:26px; font-weight:bold;">${proximas}</div>
          </td>
          <td style="width:33%; background:${PALETA_REPORTE.CARD_VENCIDAS_BG}; border-radius:10px; padding:16px;">
            <div style="color:${PALETA_REPORTE.CARD_VENCIDAS_TEXTO}; font-weight:bold; font-size:13px;">Vencidas a la fecha</div>
            <div style="color:#222; font-size:26px; font-weight:bold;">${vencidas}</div>
          </td>
        </tr>
      </table>

      <table style="width:100%; border-collapse:collapse; margin-bottom:22px;">
        <tr>
          <td style="width:58%; vertical-align:top;">
            <img src="data:image/png;base64,${chartEntidades}" style="width:100%;">
          </td>
          <td style="width:42%; vertical-align:top;">
            <img src="data:image/png;base64,${chartDona}" style="width:100%;">
          </td>
        </tr>
      </table>

      <img src="data:image/png;base64,${chartImpuestos}" style="width:100%; margin-bottom:22px;">

      <p style="border-left:4px solid #EF2C25; padding-left:10px; font-weight:bold; font-size:14px; color:#222;">Detalle de obligaciones:</p>
      <table style="width:100%; border-collapse:collapse; font-size:11px;">
        <tr style="background:#f5f5f5;">
          <th style="padding:7px; border:1px solid #ddd; text-align:left;">Compañía</th>
          <th style="padding:7px; border:1px solid #ddd; text-align:left;">Obligación</th>
          <th style="padding:7px; border:1px solid #ddd; text-align:left;">Responsable</th>
          <th style="padding:7px; border:1px solid #ddd; text-align:left;">Fecha vencimiento</th>
          <th style="padding:7px; border:1px solid #ddd; text-align:center;">Días restantes</th>
          <th style="padding:7px; border:1px solid #ddd; text-align:center;">Estado</th>
        </tr>
        ${filasTablaHtml}
      </table>
    </div>
  </body></html>`;

  const blobHtml = Utilities.newBlob(html, 'text/html', 'reporte_temp.html');
  const nombreArchivo = 'Informe_Vencimientos_' + nombreJefe.replace(/\s+/g, '_') + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd') + '.pdf';
  return blobHtml.getAs('application/pdf').setName(nombreArchivo);
}

function construirFilaDetalleHtml(item, mapa) {
  const fila = item.valores;
  const compania = valorPorColumna(fila, mapa, 'Compañía (Normalizada)');
  const impuesto = valorPorColumna(fila, mapa, 'Impuesto');
  const municipio = valorPorColumna(fila, mapa, 'Municipio');
  const responsable = valorPorColumna(fila, mapa, 'Encargado Nombre');
  const fechaMaxima = valorPorColumna(fila, mapa, 'Fecha máxima de presentación');
  const diasRestantes = valorPorColumna(fila, mapa, 'Días Restantes');
  const estadoReal = valorPorColumna(fila, mapa, 'Estado Actual');
  const estadoMostrado = etiquetaEstadoParaJefe(estadoReal);
  const semaforo = calcularSemaforo(estadoReal, diasRestantes);

  const fechaTexto = fechaMaxima instanceof Date
    ? Utilities.formatDate(fechaMaxima, Session.getScriptTimeZone(), 'dd/MM/yyyy')
    : '(por confirmar)';
  const impuestoTexto = municipio ? impuesto + ' - ' + municipio : impuesto;

  return `
  <tr>
    <td style="padding:7px; border:1px solid #ddd;">${compania}</td>
    <td style="padding:7px; border:1px solid #ddd;">${impuestoTexto}</td>
    <td style="padding:7px; border:1px solid #ddd;">${responsable || '(sin asignar)'}</td>
    <td style="padding:7px; border:1px solid #ddd;">${fechaTexto}</td>
    <td style="padding:7px; border:1px solid #ddd; text-align:center;">${diasRestantes} ${semaforo}</td>
    <td style="padding:7px; border:1px solid #ddd; text-align:center;">${estadoMostrado}</td>
  </tr>`;
}
function generarGraficoBarrasHorizontal(etiquetas, valores, titulo, paletaColores) {
  const dataTable = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, 'Categoria')
    .addColumn(Charts.ColumnType.NUMBER, 'Valor');
  etiquetas.forEach((et, i) => dataTable.addRow([et, valores[i]]));

  const builder = Charts.newBarChart()
    .setDataTable(dataTable.build())
    .setTitle(titulo)
    .setDimensions(430, 260)
    .setOption('legend', { position: 'none' })
    .setOption('colors', paletaColores.slice(0, etiquetas.length))
    .setOption('titleTextStyle', { fontSize: 14, bold: true });

  const chart = builder.build();
  return Utilities.base64Encode(chart.getAs('image/png').getBytes());
}

function generarGraficoColumnas(etiquetas, valores, titulo, colorBarra) {
  const dataTable = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, 'Categoria')
    .addColumn(Charts.ColumnType.NUMBER, 'Valor');
  etiquetas.forEach((et, i) => dataTable.addRow([et, valores[i]]));

  const chart = Charts.newColumnChart()
    .setDataTable(dataTable.build())
    .setTitle(titulo)
    .setDimensions(650, 240)
    .setColors([colorBarra])
    .setOption('legend', { position: 'none' })
    .setOption('titleTextStyle', { fontSize: 14, bold: true })
    .build();

  return Utilities.base64Encode(chart.getAs('image/png').getBytes());
}

function generarGraficoDona(vencidas, proximas, total) {
  const presentadas = Math.max(total - vencidas - proximas, 0);

  const dataTable = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, 'Categoria')
    .addColumn(Charts.ColumnType.NUMBER, 'Valor')
    .addRow(['Vencidas a la fecha', vencidas])
    .addRow(['Próximas a vencer', proximas]);
  if (presentadas > 0) dataTable.addRow(['Total de Obligaciones', presentadas]);

  const chart = Charts.newPieChart()
    .setDataTable(dataTable.build())
    .setTitle('Resumen de métricas')
    .setDimensions(300, 260)
    .setColors([PALETA_REPORTE.BARRA_ROJA, PALETA_REPORTE.BARRA_AMARILLA, PALETA_REPORTE.BARRA_AZUL])
    .setOption('pieHole', 0.5)
    .setOption('titleTextStyle', { fontSize: 14, bold: true })
    .build();

  return Utilities.base64Encode(chart.getAs('image/png').getBytes());
}