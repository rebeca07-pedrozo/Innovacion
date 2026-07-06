/**
 * BLOQUE 3: CORREOS DIARIOS A ENCARGADOS
 * -------------------------------------------------------------
 * Agrupa TODAS las obligaciones pendientes (Estado != Presentado)
 * de cada encargado en UN solo correo diario, con botones de acción
 * por cada obligación (porque un encargado puede tener varias).
 *
 * FUNCIONES A EJECUTAR:
 *  - enviarCorreosDiariosPrueba()      -> usa la hoja TRANSFORM2
 *  - enviarCorreosDiariosProduccion()  -> usa la hoja LOAD
 */

function enviarCorreosDiariosPrueba() {
  return enviarCorreosDiarios('TRANSFORM2');
}

function enviarCorreosDiariosProduccion() {
  return enviarCorreosDiarios(CONFIG.HOJA_LOAD);
}

function enviarCorreosDiarios(nombreHoja) {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = libro.getSheetByName(nombreHoja);
  if (!hoja) throw new Error('No se encontró la hoja "' + nombreHoja + '".');

  const mapa = obtenerMapaColumnas(hoja);
  const datos = hoja.getRange(2, 1, Math.max(hoja.getLastRow() - 1, 0), hoja.getLastColumn()).getValues();
  const hoy = new Date();
  const hoyTexto = Utilities.formatDate(hoy, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  // Agrupar filas pendientes por email de encargado
  const grupos = {}; // email -> { nombre, filas: [{numeroFila, valores}] }

  datos.forEach((fila, idx) => {
    const estado = valorPorColumna(fila, mapa, 'Estado Actual');
    const email = valorPorColumna(fila, mapa, 'Encargado Email');
    if (!email || estado === CONFIG.ESTADOS.PRESENTADO) return; // se envía TODOS los días mientras no esté Presentado

    // Puede haber varios encargados en un mismo campo (separados por "; ")
    const emails = String(email).split(';').map(e => e.trim()).filter(Boolean);
    const nombres = String(valorPorColumna(fila, mapa, 'Encargado Nombre') || '').split(';').map(n => n.trim());

    emails.forEach((correoIndividual, i) => {
      if (!grupos[correoIndividual]) {
        grupos[correoIndividual] = { nombre: nombres[i] || nombres[0] || correoIndividual, filas: [] };
      }
      grupos[correoIndividual].filas.push({ numeroFila: idx + 2, valores: fila });
    });
  });

  let correosEnviados = 0;

  Object.keys(grupos).forEach(correo => {
    const grupo = grupos[correo];
    const htmlCorreo = construirHtmlCorreoEncargado(grupo.nombre, grupo.filas, mapa, nombreHoja);

    MailApp.sendEmail({
      to: correo,
      subject: 'Vencimientos DIAN pendientes - ' + grupo.filas.length + ' obligación(es)',
      htmlBody: htmlCorreo
    });

    // Marca la fecha de último envío (solo para auditoría, no bloquea el envío del día siguiente)
    grupo.filas.forEach(item => {
      hoja.getRange(item.numeroFila, mapa['Última Fecha Envío Recordatorio'] + 1).setValue(hoyTexto);
    });

    correosEnviados++;
  });

  Logger.log('Correos enviados: ' + correosEnviados + ' (hoja: ' + nombreHoja + ')');
  SpreadsheetApp.getActiveSpreadsheet().toast(correosEnviados + ' correos enviados (hoja: ' + nombreHoja + ')', 'Notificaciones', 6);
  return correosEnviados;
}

/**
 * Construye el HTML del correo para un encargado, con una fila de
 * tabla por cada obligación pendiente y sus 3 botones de acción.
 */
function construirHtmlCorreoEncargado(nombreEncargado, filas, mapa, nombreHoja) {
  const filasHtml = filas.map(item => construirFilaHtmlObligacion(item, mapa, nombreHoja)).join('');

  return `
  <div style="font-family: Arial, sans-serif; max-width: 700px; margin: 0 auto;">
    <div style="background:#EC6E1F; padding:14px 20px; display:flex; justify-content:space-between; align-items:center;">
      <span style="color:#fff; font-weight:bold; font-size:18px;">DAVIVIENDA</span>
      <span style="color:#fff; font-size:13px;">Semana del ${obtenerRangoSemanaTexto()}</span>
    </div>
    <div style="padding:20px; border:1px solid #eee; border-top:none;">
      <h2 style="margin-top:0;">Informe de vencimientos pendientes</h2>
      <p>Hola ${nombreEncargado || ''},</p>
      <p style="border-left:4px solid #D0021B; padding-left:10px;">
        Tienes <b>${filas.length}</b> obligación(es) ante la DIAN pendiente(s) de presentar.
        Gestiona esto antes de la fecha límite para evitar sanciones.
      </p>
      <table style="width:100%; border-collapse:collapse; font-size:13px; margin-top:15px;">
        <tr style="background:#f5f5f5; text-align:left;">
          <th style="padding:8px; border:1px solid #ddd;">Compañía</th>
          <th style="padding:8px; border:1px solid #ddd;">Obligación</th>
          <th style="padding:8px; border:1px solid #ddd;">Fecha límite</th>
          <th style="padding:8px; border:1px solid #ddd;">Días restantes</th>
          <th style="padding:8px; border:1px solid #ddd;">Estado</th>
          <th style="padding:8px; border:1px solid #ddd;">Acción</th>
        </tr>
        ${filasHtml}
      </table>
    </div>
  </div>`;
}

function construirFilaHtmlObligacion(item, mapa, nombreHoja) {
  const fila = item.valores;
  const id = valorPorColumna(fila, mapa, 'ID');
  const compania = valorPorColumna(fila, mapa, 'Compañía (Normalizada)');
  const impuesto = valorPorColumna(fila, mapa, 'Impuesto');
  const municipio = valorPorColumna(fila, mapa, 'Municipio');
  const fechaMaxima = valorPorColumna(fila, mapa, 'Fecha máxima de presentación');
  const diasRestantes = valorPorColumna(fila, mapa, 'Días Restantes');
  const estado = valorPorColumna(fila, mapa, 'Estado Actual');
  const colorFondo = CONFIG.COLORES_ESTADO[estado] || '#fff';

  const fechaTexto = fechaMaxima instanceof Date
    ? Utilities.formatDate(fechaMaxima, Session.getScriptTimeZone(), 'dd/MM/yyyy')
    : '(fecha por confirmar)';

  const impuestoTexto = municipio ? impuesto + ' - ' + municipio : impuesto;

  const base = CONFIG.URL_WEBAPP + '?hoja=' + encodeURIComponent(nombreHoja) + '&id=' + encodeURIComponent(id);

  return `
  <tr style="background:${colorFondo};">
    <td style="padding:8px; border:1px solid #ddd;">${compania}</td>
    <td style="padding:8px; border:1px solid #ddd;">${impuestoTexto}</td>
    <td style="padding:8px; border:1px solid #ddd;">${fechaTexto}</td>
    <td style="padding:8px; border:1px solid #ddd; text-align:center;">${diasRestantes}</td>
    <td style="padding:8px; border:1px solid #ddd; text-align:center;">${estado}</td>
    <td style="padding:8px; border:1px solid #ddd; white-space:nowrap;">
      <a href="${base}&accion=notificado" style="color:#2C5F8A; text-decoration:none; margin-right:6px;">Notificado</a> |
      <a href="${base}&accion=enproceso" style="color:#C97A1F; text-decoration:none; margin:0 6px;">En proceso</a> |
      <a href="${base}&accion=formPresentado" style="color:#2E8B4F; text-decoration:none; margin-left:6px;">Presentado</a>
    </td>
  </tr>`;
}

function obtenerRangoSemanaTexto() {
  const hoy = new Date();
  const formato = 'd \'de\' MMMM';
  return Utilities.formatDate(hoy, Session.getScriptTimeZone(), formato);
}