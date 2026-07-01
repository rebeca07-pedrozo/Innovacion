/***** WEB APP · CAMBIO DE ESTADO Y FICHA DE FINALIZACIÓN *****/

function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  const id = p.id || '', estado = p.estado || '';
  if (!id || CONFIG.ESTADOS.indexOf(estado) < 0) {
    return HtmlService.createHtmlOutput('Solicitud no válida.');
  }
  // Si marca "Presentado" y aún no llenó la ficha, se la mostramos
  if (estado === 'Presentado' && p.confirmar !== '1') {
    return mostrarFicha_(id);
  }
  const extra = { fecha: p.fecha || '', radicado: p.radicado || '', obs: p.obs || '' };
  return paginaConfirmacion_(cambiarEstado_(id, estado, extra), id, estado);
}

function mostrarFicha_(id) {
  const t = HtmlService.createTemplateFromFile('FichaFinalizacion');
  t.id = id;
  t.url = CONFIG.URL_WEBAPP;
  t.hoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return t.evaluate().setTitle('Davivienda · Ficha de finalización');
}

function cambiarEstado_(id, estado, extra) {
  extra = extra || {};
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
      if (estado === 'Presentado') {
        obl.getRange(fila, idx('fecha_presentacion') + 1).setValue(parseFechaIso_(extra.fecha) || new Date());
        obl.getRange(fila, idx('radicado') + 1).setValue(extra.radicado || '');
      }
      registrarHistorico_(ss, {
        id: id, compania: vals[i][idx('compania')], tipo: vals[i][idx('tipo_obligacion')],
        subtipo: vals[i][idx('subtipo')]
      }, usuario || 'vía correo',
        'Estado → ' + estado + (extra.radicado ? ' · radicado ' + extra.radicado : '') +
        (extra.obs ? ' · ' + extra.obs : ''));
      return true;
    }
  }
  return false;
}

function parseFechaIso_(s) {
  if (!s) return null;
  const p = String(s).split('-');
  if (p.length !== 3) return null;
  return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
}

function paginaConfirmacion_(ok, id, estado) {
  const t = HtmlService.createTemplateFromFile('PaginaConfirmacion');
  t.ok = ok; t.id = id; t.estado = estado;
  return t.evaluate().setTitle('Davivienda · Alertas Tributarias');
}