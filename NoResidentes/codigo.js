/*************************************************
 * Automatización CDT - No residentes
 * Prototipo web app (Apps Script)
 *************************************************/

// ===== CONFIGURACIÓN =====
var ID_HOJA = '';                 // ID del Sheet. Vacío = modo demo (guarda en Propiedades)
var NOMBRE_HOJA = 'CASOS';
var CORREO_BUZON = 'gestioncdtimpuestos@davivienda.com';
var CORREOS_OPERATIVOS = ['leydi@davivienda.com'];
var TARIFA_DEFECTO = 4;
var OFICINAS = ['Principal Bogotá', 'Chapinero', 'Centro Internacional', 'Cali Norte', 'Medellín Poblado', 'Barranquilla Centro'];
var MEDIOS_PAGO = ['Efectivo', 'Débito de cuenta', 'Otras - especifique'];

// Pon tu correo aquí si quieres probar la vista de Operativos desde tu propia cuenta
var FORZAR_VISTA = '';            // '', 'Oficinas' u 'Operativos'


// ===== ENTRADA WEB =====
function doGet(e) {
  var correo = Session.getActiveUser().getEmail() || 'invitado@davivienda.com';
  var vista = FORZAR_VISTA || (esOperativo(correo) ? 'Operativos' : 'Oficinas');

  if (e && e.parameter && e.parameter.vista) {
    vista = e.parameter.vista === 'operativos' ? 'Operativos' : 'Oficinas';
  }

  var pagina = HtmlService.createTemplateFromFile(vista);
  pagina.correo = correo;
  pagina.nombre = nombreDesdeCorreo(correo);
  pagina.oficinas = OFICINAS;
  pagina.mediosPago = MEDIOS_PAGO;
  pagina.buzon = CORREO_BUZON;

  return pagina.evaluate()
    .setTitle('Automatización CDT - No residentes')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function incluir(archivo) {
  return HtmlService.createHtmlOutputFromFile(archivo).getContent();
}

function esOperativo(correo) {
  return CORREOS_OPERATIVOS.indexOf(String(correo).toLowerCase()) > -1;
}

function nombreDesdeCorreo(correo) {
  var usuario = String(correo).split('@')[0].replace(/[._]/g, ' ');
  return usuario.replace(/\b\w/g, function (l) { return l.toUpperCase(); });
}


// ===== CÁLCULO =====
function calcularRetencion(rendimientos, tarifa, retefuenteDigitada) {
  var valorSistema = Math.round(rendimientos * (tarifa / 100));
  var valorRetener = retefuenteDigitada || 0;
  var diferencia = valorSistema - valorRetener;
  return {
    valorSistema: valorSistema,
    valorRetener: valorRetener,
    diferencia: diferencia,
    estado: diferencia === 0 ? 'VÁLIDO' : 'PENDIENTE'
  };
}


// ===== GUARDAR CASO (vista Oficinas) =====
function guardarCaso(datos) {
  var casos = leerCasos();
  var calculo = calcularRetencion(
    Number(datos.rendimientos) || 0,
    TARIFA_DEFECTO,
    Number(datos.retefuente) || 0
  );

  var caso = {
    id: '#CDT-' + ('00' + (casos.length + 1)).slice(-3),
    fecha: new Date().toISOString().substring(0, 10),
    correoOficina: datos.correoOficina,
    oficina: datos.oficina,
    documento: datos.documento,
    cliente: datos.cliente,
    numeroCdt: datos.numeroCdt,
    fechaVencimiento: datos.fechaVencimiento,
    fechaProximo: datos.fechaProximo,
    medioPago: datos.medioPago,
    inversion: Number(datos.inversion) || 0,
    rendimientos: Number(datos.rendimientos) || 0,
    tarifa: TARIFA_DEFECTO,
    valorSistema: calculo.valorSistema,
    valorRetener: calculo.valorRetener,
    diferencia: calculo.diferencia,
    estado: calculo.estado,
    observaciones: datos.observaciones || ''
  };

  casos.push(caso);
  escribirCasos(casos);
  notificarOperativos(caso);
  return caso;
}


// ===== PANEL OPERATIVOS =====
function obtenerPanel() {
  var casos = leerCasos();
  var hoy = new Date().toISOString().substring(0, 10);
  return {
    casos: casos,
    pendientes: casos.filter(function (c) { return c.estado === 'PENDIENTE'; }).length,
    procesadosHoy: casos.filter(function (c) { return c.fecha === hoy; }).length,
    aprobados: casos.filter(function (c) { return c.estado === 'APROBADO'; }).length
  };
}

function actualizarTarifa(id, tarifa) {
  var casos = leerCasos();
  for (var i = 0; i < casos.length; i++) {
    if (casos[i].id === id) {
      var calculo = calcularRetencion(casos[i].rendimientos, Number(tarifa), casos[i].valorRetener);
      casos[i].tarifa = Number(tarifa);
      casos[i].valorSistema = calculo.valorSistema;
      casos[i].diferencia = calculo.diferencia;
      if (casos[i].estado !== 'APROBADO') casos[i].estado = calculo.estado;
    }
  }
  escribirCasos(casos);
  return obtenerPanel();
}

function aprobarCasos(ids, observaciones) {
  var casos = leerCasos();
  var aprobados = [];
  for (var i = 0; i < casos.length; i++) {
    if (ids.indexOf(casos[i].id) > -1) {
      casos[i].estado = 'APROBADO';
      if (observaciones) casos[i].observaciones = observaciones;
      aprobados.push(casos[i]);
    }
  }
  escribirCasos(casos);
  notificarOficinas(aprobados);
  return obtenerPanel();
}


// ===== CORREOS =====
function notificarOperativos(caso) {
  try {
    MailApp.sendEmail({
      to: CORREOS_OPERATIVOS.join(','),
      subject: 'Nuevo caso CDT para revisión: ' + caso.id,
      htmlBody: 'Caso <b>' + caso.id + '</b> de la oficina ' + caso.oficina +
        '<br>Cliente: ' + caso.cliente +
        '<br>Rendimientos: ' + caso.rendimientos +
        '<br>Diferencia a ajustar: ' + caso.diferencia
    });
  } catch (err) { /* en prototipo no interrumpe el flujo */ }
}

function notificarOficinas(casos) {
  try {
    casos.forEach(function (c) {
      if (!c.correoOficina) return;
      MailApp.sendEmail({
        to: c.correoOficina,
        subject: 'Caso ' + c.id + ' aprobado',
        htmlBody: 'Tu caso <b>' + c.id + '</b> fue aprobado.<br>Valor a retener: ' +
          c.valorSistema + '<br>Observaciones: ' + (c.observaciones || 'Sin observaciones') +
          '<br><br>' + CORREO_BUZON
      });
    });
  } catch (err) { /* en prototipo no interrumpe el flujo */ }
}


// ===== ALMACENAMIENTO =====
function leerCasos() {
  if (!ID_HOJA) {
    var guardado = PropertiesService.getScriptProperties().getProperty('casos');
    return guardado ? JSON.parse(guardado) : casosDemo();
  }
  var hoja = SpreadsheetApp.openById(ID_HOJA).getSheetByName(NOMBRE_HOJA);
  var filas = hoja.getDataRange().getValues();
  var titulos = filas.shift();
  return filas.map(function (fila) {
    var caso = {};
    titulos.forEach(function (t, i) { caso[t] = fila[i]; });
    return caso;
  });
}

function escribirCasos(casos) {
  if (!ID_HOJA) {
    PropertiesService.getScriptProperties().setProperty('casos', JSON.stringify(casos));
    return;
  }
  var hoja = SpreadsheetApp.openById(ID_HOJA).getSheetByName(NOMBRE_HOJA);
  var titulos = Object.keys(casos[0]);
  var datos = casos.map(function (c) {
    return titulos.map(function (t) { return c[t]; });
  });
  hoja.clear();
  hoja.getRange(1, 1, 1, titulos.length).setValues([titulos]);
  if (datos.length) hoja.getRange(2, 1, datos.length, titulos.length).setValues(datos);
}

function reiniciarDemo() {
  PropertiesService.getScriptProperties().deleteProperty('casos');
}

function casosDemo() {
  return [
    { id: '#CDT-001', fecha: new Date().toISOString().substring(0, 10), correoOficina: 'oficina1@davivienda.com', oficina: 'Principal Bogotá', documento: '123456789', cliente: 'Rebeca Pedrozo Cueto', numeroCdt: 'CDT-99812', fechaVencimiento: '2026-10-15', fechaProximo: '2027-10-15', medioPago: 'Débito de cuenta', inversion: 1000000000, rendimientos: 1000000, tarifa: 4, valorSistema: 40000, valorRetener: 10000, diferencia: 30000, estado: 'PENDIENTE', observaciones: '' },
    { id: '#CDT-002', fecha: new Date().toISOString().substring(0, 10), correoOficina: 'oficina2@davivienda.com', oficina: 'Chapinero', documento: '321654987', cliente: 'Claudia Milena Cueto', numeroCdt: 'CDT-77210', fechaVencimiento: '2026-11-02', fechaProximo: '2027-11-02', medioPago: 'Efectivo', inversion: 1000000000, rendimientos: 1000000000, tarifa: 14, valorSistema: 140000000, valorRetener: 140000000, diferencia: 0, estado: 'VÁLIDO', observaciones: '' },
    { id: '#CDT-003', fecha: '2026-09-28', correoOficina: 'oficina3@davivienda.com', oficina: 'Cali Norte', documento: '14589', cliente: 'Leonel Pedrozo', numeroCdt: 'CDT-10034', fechaVencimiento: '2026-12-20', fechaProximo: '2027-12-20', medioPago: 'Otras - especifique', inversion: 1000000000, rendimientos: 2500000, tarifa: 4, valorSistema: 100000, valorRetener: 0, diferencia: 100000, estado: 'PENDIENTE', observaciones: '' }
  ];
}
