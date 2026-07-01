/***** SISTEMA DE ALERTAS DE VENCIMIENTOS TRIBUTARIOS · DAVIVIENDA *****/

const CONFIG = {
  HOJA_RAW:  'RAW_CARGA',
  HOJA_OBL:  'OBLIGACIONES',
  HOJA_HIST: 'HISTORICO_ALERTAS',
  UMBRALES: [15, 7, 3],             // días antes del vencimiento en que se avisa
  UMBRAL_ESCALA: 3,                 // a partir de aquí también se copia al jefe
  REMITENTE: 'Alertas Tributarias · Davivienda',
  ESTADOS: ['Notificado', 'En proceso', 'Presentado']
};

// Columnas de RAW_CARGA (empezando en 0):
// 0:Vencimientos 1:Compañia 2:Correo 3:Correo jefe 4:NIT 5:DV 6:P
// 7:Declaración 8:1a Cuota 9:2da Cuota 10:3a Cuota 11:Exógena
const EVENTOS = [
  { col: 8,  tipo: 'Renta',   subtipo: '1a Cuota',           key: 'RENTA-C1',   criticidad: 4 },
  { col: 7,  tipo: 'Renta',   subtipo: 'Declaración',        key: 'RENTA-DECL', criticidad: 5 },
  { col: 9,  tipo: 'Renta',   subtipo: '2da Cuota',          key: 'RENTA-C2',   criticidad: 4 },
  { col: 10, tipo: 'Renta',   subtipo: '3a Cuota',           key: 'RENTA-C3',   criticidad: 4 },
  { col: 11, tipo: 'Exógena', subtipo: 'Información exógena', key: 'EXOGENA',    criticidad: 5 }
];

const COLS_OBL = ['id_vencimiento','codigo','compania','correo_responsable','correo_jefe',
  'nit','dv','perfil','tipo_obligacion','subtipo','fecha_limite','criticidad','multa_estimada',
  'dias_restantes','estado','semaforo','score_riesgo',
  'ultimo_umbral_enviado','fecha_estado','actualizado_por'];


/* =============================== MENÚ =============================== */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🔔 Alertas DIAN')
    .addItem('Normalizar catálogo', 'normalizarCatalogo')
    .addItem('Actualizar estados', 'actualizarEstados')
    .addItem('Enviar alertas', 'enviarAlertas')
    .addSeparator()
    .addItem('Crear triggers diarios', 'setupTriggers')
    .addToUi();
}

/* ============= NORMALIZA AUTOMÁTICAMENTE AL EDITAR RAW ============== */
function onEdit(e) {
  if (!e || !e.range) return;
  if (e.range.getSheet().getName() !== CONFIG.HOJA_RAW) return;
  if (e.range.getRow() === 1) return;
  normalizarCatalogo();
}

/* ===================== NORMALIZAR (unpivot) ========================= */
function normalizarCatalogo() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const raw = ss.getSheetByName(CONFIG.HOJA_RAW);
  if (!raw) return;
  const data  = raw.getDataRange().getValues();
  const filas = data.slice(1).filter(r => String(r[0]).trim() !== '');
  const previos = leerEstadoPrevio_(ss);

  const salida = [COLS_OBL];
  filas.forEach(r => {
    const codigo = r[0], compania = r[1], correo = r[2], correoJefe = r[3],
          nit = r[4], dv = r[5], perfil = r[6];
    EVENTOS.forEach(ev => {
      const fecha = parseFecha_(r[ev.col]);
      if (!fecha) return;
      const id = codigo + '-' + ev.key;
      const p  = previos[id] || {};
      salida.push([
        id, codigo, compania, correo, correoJefe, nit, dv, perfil,
        ev.tipo, ev.subtipo, fecha,
        p.criticidad || ev.criticidad,
        p.multa || 0,
        '',                         // dias_restantes
        p.estado || 'Pendiente',    // estado inicial antes de notificar
        '',                         // semaforo
        '',                         // score_riesgo
        p.umbral || '',
        p.festado || '',
        p.actpor || ''
      ]);
    });
  });

  const obl = obtenerHoja_(ss, CONFIG.HOJA_OBL);
  obl.clearContents();
  obl.getRange(1, 1, salida.length, COLS_OBL.length).setValues(salida);
  obl.getRange(1, 1, 1, COLS_OBL.length).setFontWeight('bold');
  obl.setFrozenRows(1);
  actualizarEstados();
}

function leerEstadoPrevio_(ss) {
  const map = {};
  const obl = ss.getSheetByName(CONFIG.HOJA_OBL);
  if (!obl || obl.getLastRow() < 2) return map;
  const vals = obl.getDataRange().getValues();
  const h = vals[0], idx = n => h.indexOf(n);
  for (let i = 1; i < vals.length; i++) {
    const id = vals[i][idx('id_vencimiento')];
    if (!id) continue;
    map[id] = {
      criticidad: vals[i][idx('criticidad')],
      multa:      vals[i][idx('multa_estimada')],
      estado:     vals[i][idx('estado')],
      umbral:     vals[i][idx('ultimo_umbral_enviado')],
      festado:    vals[i][idx('fecha_estado')],
      actpor:     vals[i][idx('actualizado_por')]
    };
  }
  return map;
}

function parseFecha_(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (Object.prototype.toString.call(v) === '[object Date]') return v;
  const partes = String(v).trim().toLowerCase().replace(/\./g, '').split(/\s+/);
  if (partes.length < 3) return null;
  const meses = { ene:0, feb:1, mar:2, abr:3, may:4, jun:5,
                  jul:6, ago:7, sep:8, set:8, oct:9, nov:10, dic:11 };
  const dia  = parseInt(partes[0], 10);
  const mes  = meses[partes[1].substring(0, 3)];
  const anio = parseInt(partes[2], 10);
  if (isNaN(dia) || mes === undefined || isNaN(anio)) return null;
  return new Date(anio, mes, dia);
}

/* ===================== ESTADOS Y SEMÁFORO =========================== */
function actualizarEstados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const obl = ss.getSheetByName(CONFIG.HOJA_OBL);
  if (!obl || obl.getLastRow() < 2) return;
  const vals = obl.getDataRange().getValues();
  const h = vals[0], idx = n => h.indexOf(n);
  const hoy = new Date(); hoy.setHours(0, 0, 0, 0);

  for (let i = 1; i < vals.length; i++) {
    const row = vals[i];
    const fecha = row[idx('fecha_limite')];
    if (!(fecha instanceof Date)) continue;
    const dias = Math.round((fecha.getTime() - hoy.getTime()) / 86400000);
    row[idx('dias_restantes')] = dias;
    row[idx('semaforo')]       = semaforo_(row[idx('estado')], dias);
    row[idx('score_riesgo')]   = scoreRiesgo_(row[idx('criticidad')], dias, row[idx('estado')]);
  }
  obl.getRange(1, 1, vals.length, h.length).setValues(vals);
}

function semaforo_(estado, dias) {
  if (estado === 'Presentado') return '✅ Presentado';
  if (dias < 0)  return '⛔ Vencida';
  if (dias <= 3) return '🔴 Crítico';
  if (dias <= 15) return '🟡 Próximo';
  return '🟢 A tiempo';
}

function scoreRiesgo_(crit, dias, estado) {
  if (estado === 'Presentado') return 0;
  const prox = dias < 0 ? 6 : dias <= 3 ? 5 : dias <= 7 ? 4 : dias <= 15 ? 3 : dias <= 30 ? 2 : 1;
  return (Number(crit) || 1) * prox;
}

/* ===================== ENVÍO DE ALERTAS ============================= */
function enviarAlertas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const obl = ss.getSheetByName(CONFIG.HOJA_OBL);
  if (!obl || obl.getLastRow() < 2) return;
  const vals = obl.getDataRange().getValues();
  const h = vals[0], idx = n => h.indexOf(n);
  let cambios = false;

  for (let i = 1; i < vals.length; i++) {
    const row = vals[i];
    if (row[idx('estado')] === 'Presentado') continue;
    const dias = Number(row[idx('dias_restantes')]);
    const umbral = umbralAplicable_(dias);
    if (umbral === null) continue;
    const ultimo = row[idx('ultimo_umbral_enviado')];
    if (ultimo !== '' && Number(ultimo) <= umbral) continue;

    const correo = String(row[idx('correo_responsable')]).trim();
    const correoJefe = String(row[idx('correo_jefe')]).trim();
    const d = {
      id: row[idx('id_vencimiento')],
      compania: row[idx('compania')],
      nit: row[idx('nit')] + '-' + row[idx('dv')],
      tipo: row[idx('tipo_obligacion')],
      subtipo: row[idx('subtipo')],
      fecha: Utilities.formatDate(row[idx('fecha_limite')], Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      dias: dias, umbral: umbral
    };
    if (!correo || correo.indexOf('@') < 0) {
      Logger.log('Sin correo válido para ' + d.id + ', se omite.');
      continue;
    }

    enviarCorreoAlerta_(correo, d, false);
    registrarHistorico_(ss, d, correo, 'Notificación enviada');

    // Escalamiento: al jefe de ESA compañía (de la columna Correo jefe)
    if (dias <= CONFIG.UMBRAL_ESCALA && correoJefe && correoJefe.indexOf('@') >= 0 &&
        correoJefe !== correo) {
      enviarCorreoAlerta_(correoJefe, d, true);
      registrarHistorico_(ss, d, correoJefe, 'Escalada al jefe');
    }

    if (row[idx('estado')] === 'Pendiente') row[idx('estado')] = 'Notificado';
    row[idx('ultimo_umbral_enviado')] = umbral;
    cambios = true;
  }
  if (cambios) obl.getRange(1, 1, vals.length, h.length).setValues(vals);
}

function umbralAplicable_(dias) {
  if (dias < 0) return null;
  const us = CONFIG.UMBRALES.slice().sort((a, b) => a - b);
  for (const u of us) if (dias <= u) return u;
  return null;
}

/* ===================== CORREO (usa CorreoAlerta.html) =============== */
function enviarCorreoAlerta_(destinatario, d, esEscalamiento) {
  const base = ScriptApp.getService().getUrl();
  const link = est => base + '?id=' + encodeURIComponent(d.id) + '&estado=' + encodeURIComponent(est);

  const t = HtmlService.createTemplateFromFile('CorreoAlerta');
  t.v = {
    compania: d.compania, nit: d.nit, tipo: d.tipo, subtipo: d.subtipo,
    fecha: d.fecha, dias: d.dias, esc: !!esEscalamiento,
    urg:    d.dias <= 3 ? '#C0392B' : (d.dias <= 7 ? '#E67E22' : '#D4AC0D'),
    urgTxt: d.dias <= 3 ? 'CRÍTICO' : (d.dias <= 7 ? 'URGENTE' : 'PRÓXIMO'),
    urlNotificado: link('Notificado'),
    urlProceso:    link('En proceso'),
    urlPresentado: link('Presentado')
  };
  const html = t.evaluate().getContent();
  const asunto = (esEscalamiento ? '[ESCALAMIENTO] ' : '') +
    'Vencimiento ' + d.tipo + ' · ' + d.subtipo + ' · ' + d.compania +
    ' (' + d.dias + ' día' + (d.dias === 1 ? '' : 's') + ')';

  GmailApp.sendEmail(destinatario, asunto, textoPlano_(d, link('Presentado')),
    { htmlBody: html, name: CONFIG.REMITENTE });
}

function textoPlano_(d, urlPresentado) {
  return 'Vencimiento próximo ante la DIAN\n\n' +
    d.tipo + ' - ' + d.subtipo + '\n' + d.compania + ' (' + d.nit + ')\n' +
    'Fecha límite: ' + d.fecha + '\nDías restantes: ' + d.dias + '\n\n' +
    'Marcar como presentado: ' + urlPresentado;
}

/* ===================== LOOP: CAMBIO DE ESTADO (Web App) ============= */
function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  const id = p.id || '', estado = p.estado || '';
  if (id && CONFIG.ESTADOS.indexOf(estado) >= 0) {
    return paginaConfirmacion_(cambiarEstado_(id, estado), id, estado);
  }
  return HtmlService.createHtmlOutput('Solicitud no válida.');
}

function cambiarEstado_(id, estado) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const obl = ss.getSheetByName(CONFIG.HOJA_OBL);
  const vals = obl.getDataRange().getValues();
  const h = vals[0], idx = n => h.indexOf(n);
  let usuario = '';
  try { usuario = Session.getActiveUser().getEmail(); } catch (e) {}

  for (let i = 1; i < vals.length; i++) {
    if (vals[i][idx('id_vencimiento')] === id) {
      const fila = i + 1, dias = Number(vals[i][idx('dias_restantes')]);
      obl.getRange(fila, idx('estado') + 1).setValue(estado);
      obl.getRange(fila, idx('fecha_estado') + 1).setValue(new Date());
      obl.getRange(fila, idx('actualizado_por') + 1).setValue(usuario || 'vía correo');
      obl.getRange(fila, idx('semaforo') + 1).setValue(semaforo_(estado, dias));
      obl.getRange(fila, idx('score_riesgo') + 1).setValue(scoreRiesgo_(vals[i][idx('criticidad')], dias, estado));
      registrarHistorico_(ss, {
        id: id, compania: vals[i][idx('compania')], tipo: vals[i][idx('tipo_obligacion')],
        subtipo: vals[i][idx('subtipo')]
      }, usuario || 'vía correo', 'Estado → ' + estado);
      return true;
    }
  }
  return false;
}

function paginaConfirmacion_(ok, id, estado) {
  const t = HtmlService.createTemplateFromFile('PaginaConfirmacion');
  t.ok = ok; t.id = id; t.estado = estado;
  return t.evaluate().setTitle('Davivienda · Alertas Tributarias');
}

/* ===================== HISTÓRICO Y UTILIDADES ====================== */
function registrarHistorico_(ss, d, destinatario, accion) {
  const hist = obtenerHoja_(ss, CONFIG.HOJA_HIST);
  if (hist.getLastRow() === 0) {
    hist.appendRow(['fecha_hora','id_vencimiento','compania','tipo','subtipo','destinatario','accion']);
    hist.getRange(1, 1, 1, 7).setFontWeight('bold');
    hist.setFrozenRows(1);
  }
  hist.appendRow([new Date(), d.id, d.compania || '', d.tipo || '', d.subtipo || '', destinatario, accion]);
}

function obtenerHoja_(ss, nombre) {
  return ss.getSheetByName(nombre) || ss.insertSheet(nombre);
}

function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['actualizarEstados', 'enviarAlertas'].indexOf(t.getHandlerFunction()) >= 0) {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('actualizarEstados').timeBased().everyDays(1).atHour(6).create();
  ScriptApp.newTrigger('enviarAlertas').timeBased().everyDays(1).atHour(7).create();
  SpreadsheetApp.getActiveSpreadsheet().toast('Triggers diarios creados (6am y 7am).', 'Listo', 6);
}


/* =================== REPORTE PDF SEMANAL PARA JEFES =================== */

// Envía un PDF a cada jefe con SOLO las obligaciones de sus compañías.
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
      'Adjuntamos el reporte de vencimientos de tus compañías. No respondas a este correo.',
      { name: CONFIG.REMITENTE, attachments: [pdf] });
    registrarHistorico_(ss, { id: 'REPORTE', compania: '', tipo: 'Reporte', subtipo: 'Semanal' },
                        jefe, 'Reporte PDF enviado');
  });
}

// Para PROBAR el PDF ya mismo: te lo manda a ti con TODAS las obligaciones.
function enviarReportePrueba() {
  const CORREO_PRUEBA = 'rebeca.pedrozo@davivienda.com';   // cámbialo si quieres
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const obl = ss.getSheetByName(CONFIG.HOJA_OBL);
  const vals = obl.getDataRange().getValues();
  const h = vals[0], idx = n => h.indexOf(n);
  const html = construirReporteHtml_('Reporte de vencimientos (prueba)', vals.slice(1), idx);
  const pdf = Utilities.newBlob(html, MimeType.HTML, 'Reporte.html')
                .getAs(MimeType.PDF).setName('Reporte_vencimientos.pdf');
  GmailApp.sendEmail(CORREO_PRUEBA, 'Reporte de vencimientos (prueba)',
    'Adjunto el reporte de prueba en PDF.', { name: CONFIG.REMITENTE, attachments: [pdf] });
}

// Construye el HTML del reporte (resumen + detalle) a partir de un grupo de filas.
function construirReporteHtml_(titulo, filas, idx) {
  let total = filas.length, vencidas = 0, criticas = 0, proximas = 0, presentadas = 0;
  filas.forEach(r => {
    const dias = Number(r[idx('dias_restantes')]);
    if (r[idx('estado')] === 'Presentado') { presentadas++; return; }
    if (dias < 0) vencidas++; else if (dias <= 3) criticas++;
    if (dias >= 0 && dias <= 15) proximas++;
  });

  const orden = filas.slice().sort((a, b) => {
    const pa = a[idx('estado')] === 'Presentado' ? 1 : 0;
    const pb = b[idx('estado')] === 'Presentado' ? 1 : 0;
    if (pa !== pb) return pa - pb;
    return Number(a[idx('dias_restantes')]) - Number(b[idx('dias_restantes')]);
  });

  let detalle = '';
  orden.forEach(r => {
    const dias = Number(r[idx('dias_restantes')]);
    const col = semColorPdf_(r[idx('estado')], dias);
    detalle +=
      '<tr>' +
      '<td style="padding:7px 6px;border-bottom:1px solid #eee;"><span style="color:' + col + ';font-size:14px;">&#9679;</span></td>' +
      '<td style="padding:7px 6px;border-bottom:1px solid #eee;">' + r[idx('compania')] + '</td>' +
      '<td style="padding:7px 6px;border-bottom:1px solid #eee;">' + r[idx('tipo_obligacion')] + ' &middot; ' + r[idx('subtipo')] + '</td>' +
      '<td style="padding:7px 6px;border-bottom:1px solid #eee;">' + fmtFechaEs_(r[idx('fecha_limite')]) + '</td>' +
      '<td style="padding:7px 6px;border-bottom:1px solid #eee;text-align:right;">' + dias + '</td>' +
      '<td style="padding:7px 6px;border-bottom:1px solid #eee;">' + r[idx('estado')] + '</td>' +
      '</tr>';
  });

  return '' +
  '<div style="font-family:Arial,Helvetica,sans-serif;color:#333;">' +
  '<table width="100%" cellpadding="0" cellspacing="0"><tr>' +
    '<td style="background:#ED1C27;padding:16px 20px;color:#fff;font-size:18px;font-weight:bold;">Davivienda &nbsp;&middot;&nbsp; ' + titulo + '</td>' +
    '<td style="background:#ED1C27;padding:16px 20px;color:#fff;font-size:12px;text-align:right;">Generado el ' + fmtFechaEs_(new Date()) + '</td>' +
  '</tr></table>' +
  '<table width="100%" cellspacing="8" cellpadding="0" style="margin-top:14px;"><tr>' +
    kpi_('Total obligaciones', total, '#185FA5') +
    kpi_('Próximas (&le;15 días)', proximas, '#BA7517') +
    kpi_('Vencidas', vencidas, '#A32D2D') +
    kpi_('Presentadas', presentadas, '#3B6D11') +
  '</tr></table>' +
  '<h3 style="margin:22px 4px 8px;font-size:15px;color:#222;">Detalle de obligaciones</h3>' +
  '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:13px;">' +
    '<tr style="color:#888;font-size:12px;">' +
      '<th style="padding:6px;"></th>' +
      '<th style="padding:6px;text-align:left;">Compañía</th>' +
      '<th style="padding:6px;text-align:left;">Impuesto</th>' +
      '<th style="padding:6px;text-align:left;">Vence</th>' +
      '<th style="padding:6px;text-align:right;">Días</th>' +
      '<th style="padding:6px;text-align:left;">Estado</th></tr>' +
    detalle +
  '</table>' +
  '<p style="margin-top:16px;font-size:11px;color:#999;">Reporte automático del Sistema de Alertas de Vencimientos Tributarios &middot; Davivienda.</p>' +
  '</div>';
}

function kpi_(label, valor, color) {
  return '<td width="25%" valign="top" style="background:#F1EFE8;padding:12px 14px;border-radius:8px;">' +
    '<div style="font-size:12px;color:#5F5E5A;">' + label + '</div>' +
    '<div style="font-size:24px;font-weight:bold;color:' + color + ';margin-top:4px;">' + valor + '</div></td>';
}

function semColorPdf_(estado, dias) {
  if (estado === 'Presentado') return '#639922';
  if (dias < 0)  return '#791F1F';
  if (dias <= 3) return '#E24B4A';
  if (dias <= 15) return '#EF9F27';
  return '#639922';
}

function fmtFechaEs_(d) {
  if (!(d instanceof Date)) return '';
  const m = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
  return d.getDate() + ' ' + m[d.getMonth()] + ' ' + d.getFullYear();
}

// Programa el envío automático cada lunes a las 7am.
function setupTriggerReporte() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'enviarReporteSemanalJefes') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('enviarReporteSemanalJefes')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
  SpreadsheetApp.getActiveSpreadsheet().toast('Reporte semanal programado: lunes 7am', 'Listo', 6);
}

//datos nuevos

/* ===== CARGA DATOS REALES DE JULIO (correr una sola vez para la demo) ===== */
function cargarDatosJulio() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const obl = obtenerHoja_(ss, CONFIG.HOJA_OBL);
  obl.clear();

  // [compania, nit, tipo, subtipo, dia_julio]
  const base = [
    ['FIDUCIARIA DAVIVIENDA','800182281', 9],
    ['EDICIONES GAMMA S.A.','860062001', 9],
    ['EPAYCO.COM SAS','900471052', 10],
    ['Davivienda Capital S.A.','901929182', 10],
    ['PROMOCIONES Y COBRANZAS BEL S.A.S.','860354473', 13],
    ['BANCO DAVIVIENDA','860034313', 13],
    ['VC INVESTMENTS S.A.S.','901321213', 13],
    ['Multiacciones','900122793', 13],
    ['CORREDORES DAVIVIENDA','860079174', 14],
    ['Davibank S.A.S.','860034594', 14],
    ['CORPORACION FINANCIERA DAVIVIENDA','901323565', 15],
    ['Fiduciaria Davibank','800144467', 17],
    ['INVERSIONES DATIO S.A.S.','901667807', 17],
    ['Davivienda Group S.A.','901929057', 17],
    ['COBRANZAS SIGMA','900383098', 21],
    ['RENTING DAVIVIENDA','901913509', 22],
    ['Davibank Securities SA','830504700', 23],
    ['INVERSIONES CFD S.A.S.','901475500', 23]
  ];

  const impuestos = [
    { tipo: 'IVA Bimestral',        subtipo: '3B 2026',   key: 'IVA-3B',  crit: 4 },
    { tipo: 'Retención en la Fuente', subtipo: 'Junio 2026', key: 'RTF-JUN', crit: 4 }
  ];

  const filas = [COLS_OBL];
  base.forEach((c, i) => {
    const codigo = 'JUL-' + String(i + 1).padStart(2, '0');
    impuestos.forEach(imp => {
      filas.push([
        codigo + '-' + imp.key, codigo, c[0], '', '', c[1], '', 'GC',
        imp.tipo, imp.subtipo, new Date(2026, 6, c[2]),  // 6 = julio
        imp.crit, 0, '', 'Pendiente', '', '', '', '', ''
      ]);
    });
  });

  obl.getRange(1, 1, filas.length, COLS_OBL.length).setValues(filas);
  obl.getRange(1, 1, 1, COLS_OBL.length).setFontWeight('bold');
  obl.setFrozenRows(1);
  actualizarEstados();  // calcula días, semáforo y score
  ss.toast((filas.length - 1) + ' obligaciones de julio cargadas', 'Listo', 6);
}



