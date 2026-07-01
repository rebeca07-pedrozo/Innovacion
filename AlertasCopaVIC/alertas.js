/***** CONFIGURACIÓN Y MOTOR DE ALERTAS · DAVIVIENDA *****/

const CONFIG = {
  HOJA_RAW:  'RAW_CARGA',
  HOJA_OBL:  'OBLIGACIONES',
  HOJA_HIST: 'HISTORICO_ALERTAS',
  UMBRALES: [15, 7, 3],
  UMBRAL_ESCALA: 3,
  URL_WEBAPP: 'https://script.google.com/a/macros/davivienda.com/s/AKfycbwg2yWgU2o8cUJRbiqEfQoRv4aUSxsjpOGTQKaQjc3WmYMWMXeo7Y5aotQ8nHLPlZL4pg/exec',
  LOGO_FILE_ID: '1-yqbDOFcf0WNPG935xahfwZAS2T18zFE',
  CORREO_PRUEBA: 'rebeca.pedrozo@davivienda.com',
  REMITENTE: 'Alertas Tributarias · Davivienda',
  ESTADOS: ['Notificado', 'En proceso', 'Presentado']
};

const EVENTOS = [
  { col: 8,  tipo: 'Renta',   subtipo: '1a Cuota',           key: 'RENTA-C1',   criticidad: 4 },
  { col: 7,  tipo: 'Renta',   subtipo: 'Declaración',        key: 'RENTA-DECL', criticidad: 5 },
  { col: 9,  tipo: 'Renta',   subtipo: '2da Cuota',          key: 'RENTA-C2',   criticidad: 4 },
  { col: 10, tipo: 'Renta',   subtipo: '3a Cuota',           key: 'RENTA-C3',   criticidad: 4 },
  { col: 11, tipo: 'Exógena', subtipo: 'Información exógena', key: 'EXOGENA',    criticidad: 5 }
];

const COLS_OBL = ['id_vencimiento','codigo','compania','correo_responsable','correo_jefe',
  'nit','dv','perfil','tipo_obligacion','subtipo','fecha_limite','criticidad','multa_estimada',
  'dias_restantes','estado','semaforo','score_riesgo','ultimo_umbral_enviado',
  'fecha_estado','actualizado_por','fecha_presentacion','radicado'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🔔 Alertas DIAN')
    .addItem('Cargar datos de julio (demo)', 'cargarDatosJulio')
    .addItem('Normalizar catálogo', 'normalizarCatalogo')
    .addItem('Actualizar estados', 'actualizarEstados')
    .addSeparator()
    .addItem('Enviar alerta de prueba (a mí)', 'enviarAlertaPrueba')
    .addItem('Enviar alertas (real)', 'enviarAlertas')
    .addItem('Enviar PDF de prueba (a mí)', 'enviarReportePrueba')
    .addSeparator()
    .addItem('Crear triggers diarios', 'setupTriggers')
    .addItem('Programar reporte semanal (lunes)', 'setupTriggerReporte')
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.range) return;
  if (e.range.getSheet().getName() !== CONFIG.HOJA_RAW) return;
  if (e.range.getRow() === 1) return;
  normalizarCatalogo();
}

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
        p.criticidad || ev.criticidad, p.multa || 0,
        '', p.estado || 'Pendiente', '', '',
        p.umbral || '', p.festado || '', p.actpor || '',
        p.fpres || '', p.rad || ''
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
      criticidad: vals[i][idx('criticidad')], multa: vals[i][idx('multa_estimada')],
      estado: vals[i][idx('estado')], umbral: vals[i][idx('ultimo_umbral_enviado')],
      festado: vals[i][idx('fecha_estado')], actpor: vals[i][idx('actualizado_por')],
      fpres: vals[i][idx('fecha_presentacion')], rad: vals[i][idx('radicado')]
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
  const dia = parseInt(partes[0], 10), mes = meses[partes[1].substring(0, 3)],
        anio = parseInt(partes[2], 10);
  if (isNaN(dia) || mes === undefined || isNaN(anio)) return null;
  return new Date(anio, mes, dia);
}

function actualizarEstados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const obl = ss.getSheetByName(CONFIG.HOJA_OBL);
  if (!obl || obl.getLastRow() < 2) return;
  const vals = obl.getDataRange().getValues();
  const h = vals[0], idx = n => h.indexOf(n);
  const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  for (let i = 1; i < vals.length; i++) {
    const row = vals[i], fecha = row[idx('fecha_limite')];
    if (!(fecha instanceof Date)) continue;
    const dias = Math.round((fecha.getTime() - hoy.getTime()) / 86400000);
    row[idx('dias_restantes')] = dias;
    row[idx('semaforo')]     = semaforo_(row[idx('estado')], dias);
    row[idx('score_riesgo')] = scoreRiesgo_(row[idx('criticidad')], dias, row[idx('estado')]);
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
      id: row[idx('id_vencimiento')], compania: row[idx('compania')],
      nit: row[idx('nit')] + '-' + row[idx('dv')], tipo: row[idx('tipo_obligacion')],
      subtipo: row[idx('subtipo')],
      fecha: Utilities.formatDate(row[idx('fecha_limite')], Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      dias: dias, umbral: umbral
    };
    if (!correo || correo.indexOf('@') < 0) continue;

    enviarCorreoAlerta_(correo, d, false);
    registrarHistorico_(ss, d, correo, 'Notificación enviada');
    if (dias <= CONFIG.UMBRAL_ESCALA && correoJefe && correoJefe.indexOf('@') >= 0 && correoJefe !== correo) {
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

function enviarCorreoAlerta_(destinatario, d, esEscalamiento) {
  const base = CONFIG.URL_WEBAPP;
  const link = est => base + '?id=' + encodeURIComponent(d.id) + '&estado=' + encodeURIComponent(est);
  const t = HtmlService.createTemplateFromFile('CorreoAlerta');
  t.v = {
    compania: d.compania, nit: d.nit, tipo: d.tipo, subtipo: d.subtipo,
    fecha: d.fecha, dias: d.dias, esc: !!esEscalamiento,
    urg: d.dias <= 3 ? '#C0392B' : (d.dias <= 7 ? '#E67E22' : '#D4AC0D'),
    urgTxt: d.dias <= 3 ? 'CRÍTICO' : (d.dias <= 7 ? 'URGENTE' : 'PRÓXIMO'),
    urlNotificado: link('Notificado'), urlProceso: link('En proceso'), urlPresentado: link('Presentado')
  };
  const html = t.evaluate().getContent();
  const asunto = (esEscalamiento ? '[ESCALAMIENTO] ' : '') +
    'Vencimiento ' + d.tipo + ' · ' + d.subtipo + ' · ' + d.compania +
    ' (' + d.dias + ' día' + (d.dias === 1 ? '' : 's') + ')';
  GmailApp.sendEmail(destinatario, asunto, textoPlano_(d, link('Presentado')),
    { htmlBody: html, name: CONFIG.REMITENTE });
}

function textoPlano_(d, urlPresentado) {
  return 'Vencimiento próximo ante la DIAN\n\n' + d.tipo + ' - ' + d.subtipo + '\n' +
    d.compania + ' (' + d.nit + ')\nFecha límite: ' + d.fecha + '\nDías restantes: ' + d.dias +
    '\n\nMarcar como presentado: ' + urlPresentado;
}

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

function fmtFechaEs_(d) {
  if (!(d instanceof Date)) return '';
  const m = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
  return d.getDate() + ' ' + m[d.getMonth()] + ' ' + d.getFullYear();
}

function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['actualizarEstados', 'enviarAlertas'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('actualizarEstados').timeBased().everyDays(1).atHour(6).create();
  ScriptApp.newTrigger('enviarAlertas').timeBased().everyDays(1).atHour(7).create();
  SpreadsheetApp.getActiveSpreadsheet().toast('Triggers diarios creados (6am y 7am).', 'Listo', 6);
}