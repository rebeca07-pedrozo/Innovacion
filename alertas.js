/***********************************************************************
 * SISTEMA DE ALERTAS DE VENCIMIENTOS TRIBUTARIOS · DAVIVIENDA
 * Stack: Google Sheets + Google Apps Script + Gmail (+ Looker Studio)
 *
 * Flujo:
 *   RAW_CARGA (Excel ancho)  --normalizarCatalogo()-->  OBLIGACIONES (largo)
 *   OBLIGACIONES  --actualizarEstados() / enviarAlertas()-->  correos + HISTORICO
 *   Correo (botón 1 clic) --doGet()-->  marca "Presentada" en OBLIGACIONES
 *
 * Orden de instalación: usa el menú "🔔 Alertas DIAN" (pasos 1 a 5).
 ***********************************************************************/

/* ============================ CONFIGURACIÓN ============================ */
const CONFIG = {
  HOJA_RAW:  'RAW_CARGA',
  HOJA_OBL:  'OBLIGACIONES',
  HOJA_HIST: 'HISTORICO_ALERTAS',

  // Sube el logo a tu carpeta "documentacion" en Drive, abre el archivo,
  // copia el ID (el código que va en la URL después de /d/) y pégalo aquí.
  LOGO_FILE_ID: 'PEGA_AQUI_EL_ID_DEL_LOGO',

  // A quién llegan las alertas por defecto y a quién se escala lo crítico.
  RESPONSABLE_DEFECTO: 'rebeca.pedrozo@davivienda.com',
  CORREO_ESCALAMIENTO: 'rebeca.pedrozo@davivienda.com',

  UMBRALES: [15, 7, 3],        // días antes del vencimiento en que se avisa
  UMBRAL_ESCALA: 3,            // a partir de aquí también se notifica a escalamiento
  REMITENTE: 'Alertas Tributarias · Davivienda',
  COLOR_DAVIVIENDA: '#ED1C27'
};

// Cada fila ancha de RAW_CARGA genera estos eventos (esto es el "unpivot").
// 'col' = índice de la columna de fecha en RAW_CARGA (empezando en 0).
const EVENTOS = [
  { col: 6, tipo: 'Renta',   subtipo: '1a Cuota',           key: 'RENTA-C1',   criticidad: 4 },
  { col: 5, tipo: 'Renta',   subtipo: 'Declaración',        key: 'RENTA-DECL', criticidad: 5 },
  { col: 7, tipo: 'Renta',   subtipo: '2da Cuota',          key: 'RENTA-C2',   criticidad: 4 },
  { col: 8, tipo: 'Renta',   subtipo: '3a Cuota',           key: 'RENTA-C3',   criticidad: 4 },
  { col: 9, tipo: 'Exógena', subtipo: 'Información exógena', key: 'EXOGENA',    criticidad: 5 }
];
// Columnas de RAW_CARGA: 0:Codigo 1:Compañia 2:NIT 3:DV 4:Perfil
//                        5:Declaración 6:1aCuota 7:2daCuota 8:3aCuota 9:Exógena

const COLS_OBL = ['id_vencimiento','codigo','compania','nit','dv','perfil',
  'tipo_obligacion','subtipo','fecha_limite','responsable','correo_responsable',
  'criticidad','multa_estimada','dias_restantes','estado','semaforo','score_riesgo',
  'ultimo_umbral_enviado','fecha_confirmacion','confirmado_por'];


/* ============================== MENÚ ================================== */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🔔 Alertas DIAN')
    .addItem('1. Crear datos de prueba', 'crearDatosDePrueba')
    .addItem('2. Normalizar catálogo',   'normalizarCatalogo')
    .addItem('3. Actualizar estados',    'actualizarEstados')
    .addItem('4. Enviar alertas (prueba)','enviarAlertas')
    .addSeparator()
    .addItem('5. Crear triggers diarios','setupTriggers')
    .addToUi();
}


/* ===================== 1. DATOS DE PRUEBA ============================= */
function crearDatosDePrueba() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const raw = obtenerHoja_(ss, CONFIG.HOJA_RAW);
  raw.clear();
  const enc = ['Vencimientos','Compañia','NIT','DV','P',
               'Declaración','1a Cuota','2da Cuota','3a Cuota','Exógena'];
  const datos = [
    enc,
    // --- Tus dos compañías reales ---
    ['DAV061','FIDUCIARIA DAVIVIENDA','800182281','1','GC','13 abr 2026','10 feb 2026','13 abr 2026','10 jun 2026','28 abr 2026'],
    ['DAV113','EDICIONES GAMMA S.A.','860062001','1','GC','13 abr 2026','10 feb 2026','13 abr 2026','10 jun 2026','28 abr 2026'],
    // --- Cinco de prueba con fechas variadas (rojo / amarillo / verde) ---
    ['DAV205','INVERSIONES EL ROBLE S.A.S.','830111222','7','GC','12 jun 2026','11 feb 2026','12 jun 2026','24 jul 2026','30 jun 2026'],
    ['DAV310','COMERCIAL ANDINA LTDA','900445566','3','GC','19 jun 2026','09 feb 2026','19 jun 2026','22 jul 2026','14 jul 2026'],
    ['DAV418','TRANSPORTES DEL SUR S.A.','860524654','5','GC','30 jun 2026','12 feb 2026','30 jun 2026','28 jul 2026','24 jul 2026'],
    ['DAV522','CONSTRUCTORA NORTE S.A.S.','901223344','8','GC','24 jun 2026','11 feb 2026','24 jun 2026','27 jul 2026','31 jul 2026'],
    ['DAV630','AGRO EXPORT COLOMBIA S.A.','830998877','2','GC','03 jul 2026','13 feb 2026','03 jul 2026','29 jul 2026','21 jul 2026']
  ];
  raw.getRange(1, 1, datos.length, enc.length).setValues(datos);
  raw.getRange(1, 1, 1, enc.length).setFontWeight('bold');
  raw.setFrozenRows(1);
  ss.toast(datos.length - 1 + ' compañías cargadas en RAW_CARGA', 'Listo', 5);
}


/* ===================== 2. NORMALIZAR (unpivot) ======================= */
function normalizarCatalogo() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const raw = ss.getSheetByName(CONFIG.HOJA_RAW);
  if (!raw) throw new Error('Falta la hoja ' + CONFIG.HOJA_RAW + ' (corre el paso 1).');

  const data  = raw.getDataRange().getValues();
  const filas = data.slice(1).filter(r => String(r[0]).trim() !== '');
  const previos = leerEstadoPrevio_(ss);   // conserva lo que ya editaste/confirmaste

  const salida = [COLS_OBL];
  filas.forEach(r => {
    const codigo = r[0], compania = r[1], nit = r[2], dv = r[3], perfil = r[4];
    EVENTOS.forEach(ev => {
      const fecha = parseFecha_(r[ev.col]);
      if (!fecha) return;                   // sin fecha no se crea el evento
      const id   = codigo + '-' + ev.key;
      const p    = previos[id] || {};
      salida.push([
        id, codigo, compania, nit, dv, perfil,
        ev.tipo, ev.subtipo, fecha,
        p.responsable || CONFIG.RESPONSABLE_DEFECTO,
        p.correo      || CONFIG.RESPONSABLE_DEFECTO,
        p.criticidad  || ev.criticidad,
        p.multa       || 0,
        '',                                  // dias_restantes  (lo calcula el paso 3)
        p.estado || 'Pendiente',
        '',                                  // semaforo
        '',                                  // score_riesgo
        p.umbral || '',
        p.fconf  || '',
        p.confpor|| ''
      ]);
    });
  });

  const obl = obtenerHoja_(ss, CONFIG.HOJA_OBL);
  obl.clearContents();
  obl.getRange(1, 1, salida.length, COLS_OBL.length).setValues(salida);
  obl.getRange(1, 1, 1, COLS_OBL.length).setFontWeight('bold');
  obl.setFrozenRows(1);
  actualizarEstados();
  ss.toast((salida.length - 1) + ' vencimientos en OBLIGACIONES', 'Catálogo normalizado', 5);
}

function leerEstadoPrevio_(ss) {
  const map = {};
  const obl = ss.getSheetByName(CONFIG.HOJA_OBL);
  if (!obl || obl.getLastRow() < 2) return map;
  const vals = obl.getDataRange().getValues();
  const h = vals[0], idx = n => h.indexOf(n);
  for (let i = 1; i < vals.length; i++) {
    const row = vals[i], id = row[idx('id_vencimiento')];
    if (!id) continue;
    map[id] = {
      responsable: row[idx('responsable')],
      correo:      row[idx('correo_responsable')],
      criticidad:  row[idx('criticidad')],
      multa:       row[idx('multa_estimada')],
      estado:      row[idx('estado')],
      umbral:      row[idx('ultimo_umbral_enviado')],
      fconf:       row[idx('fecha_confirmacion')],
      confpor:     row[idx('confirmado_por')]
    };
  }
  return map;
}

// Acepta una fecha real o el texto "13 abr 2026" / "13 abril 2026".
function parseFecha_(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (Object.prototype.toString.call(v) === '[object Date]') return v;
  const s = String(v).trim().toLowerCase().replace(/\./g, '');
  const partes = s.split(/\s+/);
  if (partes.length < 3) return null;
  const meses = { ene:0, feb:1, mar:2, abr:3, may:4, jun:5,
                  jul:6, ago:7, sep:8, set:8, oct:9, nov:10, dic:11 };
  const dia  = parseInt(partes[0], 10);
  const mes  = meses[partes[1].substring(0, 3)];
  const anio = parseInt(partes[2], 10);
  if (isNaN(dia) || mes === undefined || isNaN(anio)) return null;
  return new Date(anio, mes, dia);
}


/* ===================== 3. ESTADOS Y SEMÁFORO ========================= */
function actualizarEstados() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
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

    let estado = row[idx('estado')];
    if (estado !== 'Presentada') estado = dias < 0 ? 'Vencida' : 'Pendiente';
    row[idx('estado')]      = estado;
    row[idx('semaforo')]    = semaforo_(estado, dias);
    row[idx('score_riesgo')]= scoreRiesgo_(row[idx('criticidad')], dias, estado);
  }
  obl.getRange(1, 1, vals.length, h.length).setValues(vals);
}

function semaforo_(estado, dias) {
  if (estado === 'Presentada') return '✅ Presentada';
  if (estado === 'Vencida')    return '⛔ Vencida';
  if (dias <= 3)  return '🔴 Crítico';
  if (dias <= 15) return '🟡 Próximo';
  return '🟢 A tiempo';
}

function scoreRiesgo_(criticidad, dias, estado) {
  if (estado === 'Presentada') return 0;
  const prox = dias < 0 ? 6 : dias <= 3 ? 5 : dias <= 7 ? 4 : dias <= 15 ? 3 : dias <= 30 ? 2 : 1;
  return (Number(criticidad) || 1) * prox;
}


/* ===================== 4. ENVÍO DE ALERTAS ============================ */
function enviarAlertas() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const obl = ss.getSheetByName(CONFIG.HOJA_OBL);
  if (!obl || obl.getLastRow() < 2) return;
  const vals = obl.getDataRange().getValues();
  const h = vals[0], idx = n => h.indexOf(n);
  let cambios = false;

  for (let i = 1; i < vals.length; i++) {
    const row = vals[i];
    if (row[idx('estado')] !== 'Pendiente') continue;
    const dias   = Number(row[idx('dias_restantes')]);
    const umbral = umbralAplicable_(dias);
    if (umbral === null) continue;

    const ultimo = row[idx('ultimo_umbral_enviado')];
    if (ultimo !== '' && Number(ultimo) <= umbral) continue; // ese umbral ya se avisó

    const d = {
      id:       row[idx('id_vencimiento')],
      compania: row[idx('compania')],
      nit:      row[idx('nit')] + '-' + row[idx('dv')],
      tipo:     row[idx('tipo_obligacion')],
      subtipo:  row[idx('subtipo')],
      fecha:    Utilities.formatDate(row[idx('fecha_limite')], Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      dias:     dias,
      umbral:   umbral
    };
    const correo = row[idx('correo_responsable')] || CONFIG.RESPONSABLE_DEFECTO;
    enviarCorreoAlerta_(correo, d, false);
    registrarHistorico_(ss, d, correo, 'enviada');

    if (dias <= CONFIG.UMBRAL_ESCALA && CONFIG.CORREO_ESCALAMIENTO &&
        CONFIG.CORREO_ESCALAMIENTO !== correo) {
      enviarCorreoAlerta_(CONFIG.CORREO_ESCALAMIENTO, d, true);
      registrarHistorico_(ss, d, CONFIG.CORREO_ESCALAMIENTO, 'escalada');
    }
    row[idx('ultimo_umbral_enviado')] = umbral;
    cambios = true;
  }
  if (cambios) obl.getRange(1, 1, vals.length, h.length).setValues(vals);
}

function umbralAplicable_(dias) {
  if (dias < 0) return null;
  const us = CONFIG.UMBRALES.slice().sort((a, b) => a - b); // [3,7,15]
  for (const u of us) if (dias <= u) return u;
  return null;
}


/* ===================== CORREO PROFESIONAL ============================ */
function enviarCorreoAlerta_(destinatario, d, esEscalamiento) {
  const url    = urlConfirmacion_(d.id);
  const html   = plantillaCorreo_(d, url, esEscalamiento);
  const asunto = (esEscalamiento ? '[ESCALAMIENTO] ' : '') +
    'Vencimiento ' + d.tipo + ' · ' + d.subtipo + ' · ' + d.compania +
    ' (' + d.dias + ' día' + (d.dias === 1 ? '' : 's') + ')';

  const opciones = { htmlBody: html, name: CONFIG.REMITENTE };

  // Logo embebido (solo si ya pegaste el ID). Si falla, el correo igual se envía.
  try {
    if (CONFIG.LOGO_FILE_ID && CONFIG.LOGO_FILE_ID.indexOf('PEGA') === -1) {
      const blob = DriveApp.getFileById(CONFIG.LOGO_FILE_ID).getBlob().setName('logo');
      opciones.inlineImages = { logoDavivienda: blob };
    }
  } catch (e) {
    Logger.log('No se pudo cargar el logo: ' + e);
  }
  GmailApp.sendEmail(destinatario, asunto, textoPlano_(d, url), opciones);
}

function plantillaCorreo_(d, url, esc) {
  const rojo   = CONFIG.COLOR_DAVIVIENDA;
  const urg    = d.dias <= 3 ? '#C0392B' : (d.dias <= 7 ? '#E67E22' : '#D4AC0D');
  const urgTxt = d.dias <= 3 ? 'CRÍTICO' : (d.dias <= 7 ? 'URGENTE' : 'PRÓXIMO');
  const usaLogo = CONFIG.LOGO_FILE_ID && CONFIG.LOGO_FILE_ID.indexOf('PEGA') === -1;
  const logo = usaLogo
    ? '<img src="cid:logoDavivienda" alt="Davivienda" height="36" style="display:block;border:0;outline:none;">'
    : '<span style="color:#ffffff;font-size:22px;font-weight:bold;font-family:Arial,sans-serif;">Davivienda</span>';
  const banner = esc
    ? '<tr><td style="background:#7B1113;color:#fff;font-family:Arial,sans-serif;font-size:13px;padding:9px 24px;">Alerta escalada · obligación muy próxima a vencer</td></tr>'
    : '';

  return '' +
  '<div style="background:#f2f3f5;padding:24px 0;font-family:Arial,Helvetica,sans-serif;">' +
    '<table align="center" width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;background:#ffffff;border-radius:10px;overflow:hidden;border:1px solid #e6e6e6;">' +
      '<tr><td style="background:' + rojo + ';padding:18px 24px;">' + logo + '</td></tr>' +
      banner +
      '<tr><td style="padding:28px 24px 6px 24px;">' +
        '<span style="background:' + urg + ';color:#fff;font-size:12px;font-weight:bold;padding:5px 13px;border-radius:20px;letter-spacing:.5px;">' +
          urgTxt + ' · ' + d.dias + ' DÍA' + (d.dias === 1 ? '' : 'S') + '</span>' +
        '<h1 style="font-size:20px;color:#222;margin:16px 0 4px;">Vencimiento próximo ante la DIAN</h1>' +
        '<p style="font-size:14px;color:#666;margin:0 0 18px;line-height:1.5;">Gestiona esta obligación antes de la fecha límite para evitar sanciones.</p>' +
      '</td></tr>' +
      '<tr><td style="padding:0 24px;">' +
        '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:14px;color:#333;">' +
          fila_('Compañía',       d.compania) +
          fila_('NIT',            d.nit) +
          fila_('Obligación',     d.tipo + ' · ' + d.subtipo) +
          fila_('Fecha límite',   d.fecha) +
          fila_('Días restantes', d.dias) +
        '</table>' +
      '</td></tr>' +
      '<tr><td align="center" style="padding:26px 24px 30px;">' +
        '<a href="' + url + '" style="background:' + rojo + ';color:#fff;text-decoration:none;font-size:15px;font-weight:bold;padding:13px 32px;border-radius:6px;display:inline-block;">Marcar como presentada</a>' +
        '<p style="font-size:12px;color:#999;margin:14px 0 0;">Al confirmar dejarás de recibir alertas de esta obligación.</p>' +
      '</td></tr>' +
      '<tr><td style="background:#fafafa;border-top:1px solid #eee;padding:16px 24px;">' +
        '<p style="font-size:11px;color:#999;margin:0;line-height:1.5;">Mensaje automático del Sistema de Alertas de Vencimientos Tributarios. Por favor no respondas a este correo.</p>' +
      '</td></tr>' +
    '</table>' +
  '</div>';
}

function fila_(k, v) {
  return '<tr>' +
    '<td style="padding:9px 0;border-bottom:1px solid #eee;color:#888;width:140px;">' + k + '</td>' +
    '<td style="padding:9px 0;border-bottom:1px solid #eee;font-weight:bold;">' + v + '</td>' +
  '</tr>';
}

function textoPlano_(d, url) {
  return 'Vencimiento próximo ante la DIAN\n\n' +
    d.tipo + ' - ' + d.subtipo + '\n' +
    d.compania + ' (' + d.nit + ')\n' +
    'Fecha límite: ' + d.fecha + '\n' +
    'Días restantes: ' + d.dias + '\n\n' +
    'Marcar como presentada: ' + url;
}


/* ===================== LOOP: CONFIRMACIÓN (Web App) ================== */
function doGet(e) {
  const id     = e && e.parameter ? e.parameter.id : '';
  const accion = e && e.parameter ? e.parameter.action : '';
  if (accion === 'confirmar' && id) {
    const ok = confirmarPresentacion_(id);
    return paginaConfirmacion_(ok, id);
  }
  return HtmlService.createHtmlOutput('<p style="font-family:Arial">Solicitud no válida.</p>');
}

function confirmarPresentacion_(id) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const obl = ss.getSheetByName(CONFIG.HOJA_OBL);
  const vals = obl.getDataRange().getValues();
  const h = vals[0], idx = n => h.indexOf(n);
  let usuario = '';
  try { usuario = Session.getActiveUser().getEmail(); } catch (e) {}

  for (let i = 1; i < vals.length; i++) {
    if (vals[i][idx('id_vencimiento')] === id) {
      obl.getRange(i + 1, idx('estado') + 1).setValue('Presentada');
      obl.getRange(i + 1, idx('semaforo') + 1).setValue('✅ Presentada');
      obl.getRange(i + 1, idx('score_riesgo') + 1).setValue(0);
      obl.getRange(i + 1, idx('fecha_confirmacion') + 1).setValue(new Date());
      obl.getRange(i + 1, idx('confirmado_por') + 1).setValue(usuario || 'vía correo');
      registrarHistorico_(ss, {
        id: id, compania: vals[i][idx('compania')], tipo: vals[i][idx('tipo_obligacion')],
        subtipo: vals[i][idx('subtipo')]
      }, usuario || 'vía correo', 'confirmada');
      return true;
    }
  }
  return false;
}

function urlConfirmacion_(id) {
  return ScriptApp.getService().getUrl() + '?action=confirmar&id=' + encodeURIComponent(id);
}

function paginaConfirmacion_(ok, id) {
  const rojo = CONFIG.COLOR_DAVIVIENDA;
  const msg  = ok
    ? '<h2 style="color:#1d9e75;margin:0 0 8px;">¡Listo!</h2><p>La obligación <b>' + id + '</b> quedó marcada como <b>presentada</b>. No recibirás más alertas de este vencimiento.</p>'
    : '<h2 style="color:#C0392B;margin:0 0 8px;">No encontrada</h2><p>No se encontró la obligación <b>' + id + '</b>.</p>';
  return HtmlService.createHtmlOutput(
    '<div style="font-family:Arial;max-width:460px;margin:60px auto;text-align:center;border:1px solid #eee;border-radius:10px;overflow:hidden;">' +
      '<div style="background:' + rojo + ';padding:16px;color:#fff;font-weight:bold;font-size:18px;">Davivienda · Alertas Tributarias</div>' +
      '<div style="padding:30px 26px;color:#333;">' + msg + '</div>' +
    '</div>'
  );
}


/* ===================== HISTÓRICO Y UTILIDADES ======================== */
function registrarHistorico_(ss, d, destinatario, accion) {
  const hist = obtenerHoja_(ss, CONFIG.HOJA_HIST);
  if (hist.getLastRow() === 0) {
    hist.appendRow(['timestamp','id_vencimiento','compania','tipo','subtipo',
                    'fecha_limite','dias_restantes','umbral','destinatario','accion']);
    hist.getRange(1, 1, 1, 10).setFontWeight('bold');
    hist.setFrozenRows(1);
  }
  hist.appendRow([new Date(), d.id, d.compania || '', d.tipo || '', d.subtipo || '',
    d.fecha || '', (d.dias === undefined ? '' : d.dias),
    (d.umbral === undefined ? '' : d.umbral), destinatario, accion]);
}

function obtenerHoja_(ss, nombre) {
  return ss.getSheetByName(nombre) || ss.insertSheet(nombre);
}


/* ===================== 5. TRIGGERS DIARIOS =========================== */
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['actualizarEstados', 'enviarAlertas'].indexOf(t.getHandlerFunction()) >= 0) {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('actualizarEstados').timeBased().everyDays(1).atHour(6).create();
  ScriptApp.newTrigger('enviarAlertas').timeBased().everyDays(1).atHour(7).create();
  SpreadsheetApp.getActiveSpreadsheet().toast('Triggers diarios creados: estados 6am, alertas 7am', 'Listo', 6);
}