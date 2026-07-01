/***** REPORTE PDF SEMANAL PARA JEFES *****/

function enviarReporteSemanalJefes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const obl = ss.getSheetByName(CONFIG.HOJA_OBL);
  if (!obl || obl.getLastRow() < 2) return;
  const vals = obl.getDataRange().getValues();
  const h = vals[0], idx = n => h.indexOf(n);

  const porJefe = {};
  vals.slice(1).forEach(r => {
    const jefe = String(r[idx('correo_jefe')]).trim();
    if (!jefe || jefe.indexOf('@') < 0) return;
    (porJefe[jefe] = porJefe[jefe] || []).push(r);
  });

  Object.keys(porJefe).forEach(jefe => {
    const html = construirReporteHtml_('Reporte de vencimientos tributarios', porJefe[jefe], idx);
    const pdf = Utilities.newBlob(html, MimeType.HTML, 'Reporte.html')
                  .getAs(MimeType.PDF).setName('Reporte_vencimientos.pdf');
    GmailApp.sendEmail(jefe, 'Reporte semanal de vencimientos tributarios',
      'Adjuntamos el reporte de vencimientos de tus compañías.',
      { name: CONFIG.REMITENTE, attachments: [pdf] });
    registrarHistorico_(ss, { id: 'REPORTE', compania: '', tipo: 'Reporte', subtipo: 'Semanal' },
                        jefe, 'Reporte PDF enviado');
  });
}

function enviarReportePrueba() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const obl = ss.getSheetByName(CONFIG.HOJA_OBL);
  const vals = obl.getDataRange().getValues();
  const h = vals[0], idx = n => h.indexOf(n);
  const html = construirReporteHtml_('Reporte de vencimientos (prueba)', vals.slice(1), idx);
  const pdf = Utilities.newBlob(html, MimeType.HTML, 'Reporte.html')
                .getAs(MimeType.PDF).setName('Reporte_vencimientos.pdf');
  GmailApp.sendEmail(CONFIG.CORREO_PRUEBA, 'Reporte de vencimientos (prueba)',
    'Adjunto el reporte de prueba en PDF.', { name: CONFIG.REMITENTE, attachments: [pdf] });
  SpreadsheetApp.getActiveSpreadsheet().toast('PDF de prueba enviado a ' + CONFIG.CORREO_PRUEBA, 'Listo', 6);
}

function construirReporteHtml_(titulo, filas, idx) {
  let total = filas.length, vencidas = 0, proximas = 0, presentadas = 0;
  filas.forEach(r => {
    const dias = Number(r[idx('dias_restantes')]);
    if (r[idx('estado')] === 'Presentado') { presentadas++; return; }
    if (dias < 0) vencidas++;
    if (dias >= 0 && dias <= 15) proximas++;
  });

  const orden = filas.slice().sort((a, b) => {
    const pa = a[idx('estado')] === 'Presentado' ? 1 : 0;
    const pb = b[idx('estado')] === 'Presentado' ? 1 : 0;
    if (pa !== pb) return pa - pb;
    return Number(a[idx('dias_restantes')]) - Number(b[idx('dias_restantes')]);
  });

  let detalle = '', k = 0;
  orden.forEach(r => {
    const dias = Number(r[idx('dias_restantes')]);
    const col = semColorPdf_(r[idx('estado')], dias);
    const diasCol = (r[idx('estado')] !== 'Presentado' && dias <= 3) ? '#A32D2D' : '#333';
    const bg = (k++ % 2 === 0) ? '#ffffff' : '#faf9f7';
    detalle +=
      '<tr style="background:' + bg + ';">' +
      '<td style="padding:8px 10px;border-bottom:1px solid #eee;text-align:center;"><span style="color:' + col + ';font-size:15px;">&#9679;</span></td>' +
      '<td style="padding:8px 10px;border-bottom:1px solid #eee;">' + r[idx('compania')] + '</td>' +
      '<td style="padding:8px 10px;border-bottom:1px solid #eee;color:#555;">' + r[idx('tipo_obligacion')] + ' &middot; ' + r[idx('subtipo')] + '</td>' +
      '<td style="padding:8px 10px;border-bottom:1px solid #eee;">' + fmtFechaEs_(r[idx('fecha_limite')]) + '</td>' +
      '<td style="padding:8px 10px;border-bottom:1px solid #eee;text-align:right;color:' + diasCol + ';font-weight:bold;">' + dias + '</td>' +
      '<td style="padding:8px 10px;border-bottom:1px solid #eee;">' + r[idx('estado')] + '</td></tr>';
  });

  const logo = logoDataUri_();
  const logoTag = logo
    ? '<img src="' + logo + '" alt="Davivienda" height="30" style="display:block;border:0;">'
    : '<span style="color:#fff;font-size:20px;font-weight:bold;">Davivienda</span>';

  return '' +
  '<div style="font-family:Arial,Helvetica,sans-serif;color:#333;max-width:760px;margin:0 auto;">' +
  '<table width="100%" cellpadding="0" cellspacing="0" style="border-radius:10px;overflow:hidden;"><tr>' +
    '<td style="background:#ED1C27;padding:18px 22px;" valign="middle">' + logoTag + '</td>' +
    '<td style="background:#ED1C27;padding:18px 22px;color:#fff;font-size:16px;font-weight:bold;" valign="middle">' + titulo + '</td>' +
    '<td style="background:#ED1C27;padding:18px 22px;color:#ffd7d7;font-size:12px;text-align:right;" valign="middle">Generado<br>' + fmtFechaEs_(new Date()) + '</td>' +
  '</tr></table>' +
  '<p style="font-size:13px;color:#666;margin:16px 4px 12px;">Resumen del estado de las obligaciones tributarias a la fecha.</p>' +
  '<table width="100%" cellspacing="10" cellpadding="0"><tr>' +
    kpi_('Total obligaciones', total, '#185FA5') +
    kpi_('Próximas (&le;15 días)', proximas, '#BA7517') +
    kpi_('Vencidas', vencidas, '#A32D2D') +
    kpi_('Presentadas', presentadas, '#3B6D11') +
  '</tr></table>' +
  '<h3 style="margin:24px 4px 10px;font-size:15px;color:#222;border-left:4px solid #ED1C27;padding-left:10px;">Detalle de obligaciones</h3>' +
  '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:12.5px;border:1px solid #eee;">' +
    '<tr style="background:#2C2C2A;color:#fff;">' +
      '<th style="padding:9px 10px;"></th>' +
      '<th style="padding:9px 10px;text-align:left;">Compañía</th>' +
      '<th style="padding:9px 10px;text-align:left;">Impuesto</th>' +
      '<th style="padding:9px 10px;text-align:left;">Vence</th>' +
      '<th style="padding:9px 10px;text-align:right;">Días</th>' +
      '<th style="padding:9px 10px;text-align:left;">Estado</th></tr>' +
    detalle +
  '</table>' +
  '<p style="margin-top:18px;font-size:11px;color:#999;">Reporte automático del Sistema de Alertas de Vencimientos Tributarios · Davivienda.</p>' +
  '</div>';
}

function kpi_(label, valor, color) {
  return '<td width="25%" valign="top" style="background:#F7F6F1;border:1px solid #ececec;padding:14px 16px;border-radius:8px;">' +
    '<div style="font-size:12px;color:#5F5E5A;">' + label + '</div>' +
    '<div style="font-size:26px;font-weight:bold;color:' + color + ';margin-top:6px;">' + valor + '</div></td>';
}

function semColorPdf_(estado, dias) {
  if (estado === 'Presentado') return '#639922';
  if (dias < 0)  return '#791F1F';
  if (dias <= 3) return '#E24B4A';
  if (dias <= 15) return '#EF9F27';
  return '#639922';
}

function logoDataUri_() {
  try {
    const blob = DriveApp.getFileById(CONFIG.LOGO_FILE_ID).getBlob();
    return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch (e) { return ''; }
}

function setupTriggerReporte() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'enviarReporteSemanalJefes') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('enviarReporteSemanalJefes')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
  SpreadsheetApp.getActiveSpreadsheet().toast('Reporte semanal programado: lunes 7am', 'Listo', 6);
}