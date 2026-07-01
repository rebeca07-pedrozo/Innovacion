/***** CARGA DE DATOS DE JULIO Y PRUEBAS *****/

function cargarDatosJulio() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const obl = obtenerHoja_(ss, CONFIG.HOJA_OBL);
  obl.clear();
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
    { tipo: 'IVA Bimestral',          subtipo: '3B 2026',    key: 'IVA-3B',  crit: 4 },
    { tipo: 'Retención en la Fuente', subtipo: 'Junio 2026', key: 'RTF-JUN', crit: 4 }
  ];
  const filas = [COLS_OBL];
  base.forEach((c, i) => {
    const codigo = 'JUL-' + String(i + 1).padStart(2, '0');
    impuestos.forEach(imp => {
      filas.push([
        codigo + '-' + imp.key, codigo, c[0], '', '', c[1], '', 'GC',
        imp.tipo, imp.subtipo, new Date(2026, 6, c[2]),
        imp.crit, 0, '', 'Pendiente', '', '', '', '', '', '', ''
      ]);
    });
  });
  obl.getRange(1, 1, filas.length, COLS_OBL.length).setValues(filas);
  obl.getRange(1, 1, 1, COLS_OBL.length).setFontWeight('bold');
  obl.setFrozenRows(1);
  actualizarEstados();
  ss.toast((filas.length - 1) + ' obligaciones de julio cargadas', 'Listo', 6);
}

// Envía UNA alerta de prueba a tu correo (toma la primera obligación pendiente)
function enviarAlertaPrueba() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const obl = ss.getSheetByName(CONFIG.HOJA_OBL);
  const vals = obl.getDataRange().getValues();
  const h = vals[0], idx = n => h.indexOf(n);
  for (let i = 1; i < vals.length; i++) {
    const row = vals[i];
    if (row[idx('estado')] === 'Presentado') continue;
    const dias = Number(row[idx('dias_restantes')]);
    const d = {
      id: row[idx('id_vencimiento')], compania: row[idx('compania')],
      nit: row[idx('nit')] + '-' + row[idx('dv')], tipo: row[idx('tipo_obligacion')],
      subtipo: row[idx('subtipo')],
      fecha: Utilities.formatDate(row[idx('fecha_limite')], Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      dias: dias, umbral: umbralAplicable_(dias)
    };
    enviarCorreoAlerta_(CONFIG.CORREO_PRUEBA, d, false);
    ss.toast('Alerta de prueba enviada a ' + CONFIG.CORREO_PRUEBA, 'Listo', 6);
    return;
  }
}