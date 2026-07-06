function doGet(e) {
  const nombreHoja = e.parameter.hoja || CONFIG.HOJA_LOAD;
  const id = e.parameter.id;
  const accion = e.parameter.accion;

  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = libro.getSheetByName(nombreHoja);

  if (!hoja || !id) {
    return paginaHtml('Enlace inválido', 'No se pudo identificar la obligación o la hoja.', '#F2F2F2');
  }

  const resultado = buscarFilaPorId(hoja, id);
  if (!resultado) {
    return paginaHtml('No encontrado', 'Esta obligación ya no existe en "' + nombreHoja + '".', '#F2F2F2');
  }

  const mapa = resultado.mapa;
  const numeroFila = resultado.numeroFila;

  if (accion === 'notificado') {
    marcarFecha(hoja, numeroFila, mapa, 'Fecha Notificado');
    SpreadsheetApp.flush();
    return paginaHtml('¡Listo!', 'Marcaste esta obligación como <b>Notificado</b>. Ya quedó registrado.', CONFIG.COLORES_ESTADO['Notificado']);
  }

  if (accion === 'enproceso') {
    marcarFecha(hoja, numeroFila, mapa, 'Fecha En Proceso');
    SpreadsheetApp.flush();
    return paginaHtml('¡Listo!', 'Marcaste esta obligación como <b>En proceso</b>. Ya quedó registrado.', CONFIG.COLORES_ESTADO['En proceso']);
  }

  if (accion === 'formPresentado') {
    return paginaFormularioPresentado(nombreHoja, id);
  }

  if (accion === 'guardarPresentado') {
    const fechaTexto = e.parameter.fecha;
    if (!fechaTexto) {
      return paginaHtml('Falta la fecha', 'Debes ingresar la fecha de presentación.', '#F2F2F2');
    }
    const fecha = new Date(fechaTexto + 'T00:00:00');
    if (isNaN(fecha.getTime())) {
      return paginaHtml('Fecha inválida', 'El formato de fecha no es válido.', '#F2F2F2');
    }
    hoja.getRange(numeroFila, mapa['Fecha Presentado'] + 1).setValue(fecha);
    recalcularEstadoFila(hoja, numeroFila, mapa);
    SpreadsheetApp.flush();
    return paginaHtml('¡Presentado registrado!', 'Se guardó la fecha de presentación: <b>' + fechaTexto + '</b>. Ya quedó actualizado en el sistema.', CONFIG.COLORES_ESTADO['Presentado']);
  }

  const estadoActual = valorPorColumna(resultado.valores, mapa, 'Estado Actual');
  return paginaHtml('Estado actual', 'Esta obligación está: <b>' + estadoActual + '</b>', CONFIG.COLORES_ESTADO[estadoActual] || '#fff');
}

function marcarFecha(hoja, numeroFila, mapa, nombreColumna) {
  const valorActual = hoja.getRange(numeroFila, mapa[nombreColumna] + 1).getValue();
  if (!valorActual) {
    hoja.getRange(numeroFila, mapa[nombreColumna] + 1).setValue(new Date());
  }
  recalcularEstadoFila(hoja, numeroFila, mapa);
}

function recalcularEstadoFila(hoja, numeroFila, mapa) {
  const filaCompleta = hoja.getRange(numeroFila, 1, 1, hoja.getLastColumn()).getValues()[0];
  const notificado = valorPorColumna(filaCompleta, mapa, 'Fecha Notificado');
  const enProceso = valorPorColumna(filaCompleta, mapa, 'Fecha En Proceso');
  const presentado = valorPorColumna(filaCompleta, mapa, 'Fecha Presentado');
  const fechaMaxima = valorPorColumna(filaCompleta, mapa, 'Fecha máxima de presentación');

  let estado = CONFIG.ESTADOS.PENDIENTE;
  if (presentado) estado = CONFIG.ESTADOS.PRESENTADO;
  else if (enProceso) estado = CONFIG.ESTADOS.EN_PROCESO;
  else if (notificado) estado = CONFIG.ESTADOS.NOTIFICADO;

  hoja.getRange(numeroFila, mapa['Estado Actual'] + 1).setValue(estado);

  let diasRestantes = '';
  if (fechaMaxima instanceof Date) {
    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);
    diasRestantes = Math.round((fechaMaxima - hoy) / 86400000);
  }

  if (mapa['Semáforo'] !== undefined) {
    const semaforo = calcularSemaforo(estado, diasRestantes);
    hoja.getRange(numeroFila, mapa['Semáforo'] + 1).setValue(semaforo);
  }

  if (presentado && fechaMaxima instanceof Date) {
    const fechaPres = presentado instanceof Date ? presentado : new Date(presentado);
    const extemporaneidad = Math.round((fechaPres - fechaMaxima) / 86400000);
    hoja.getRange(numeroFila, mapa['Extemporaneidad (días)'] + 1).setValue(extemporaneidad);
  }

  const colorCelda = CONFIG.COLORES_ESTADO[estado] || '#FFFFFF';
  hoja.getRange(numeroFila, mapa['Estado Actual'] + 1).setBackground(colorCelda);
}

function headerWebApp() {
  return `
  <div style="background-color:#FDE1E0; padding:18px 24px; text-align:center;">
    <img src="data:image/png;base64,${CONFIG.LOGO_BASE64}" style="height:48px;">
  </div>`;
}

function paginaHtml(titulo, texto, color) {
  const html = `
  <!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"></head>
  <body style="font-family:Arial, sans-serif; margin:0; padding:0; background:#f4f4f4;">
    ${headerWebApp()}
    <div style="max-width:480px; margin:40px auto; text-align:center; background:${color}; padding:36px; border-radius:12px; box-shadow:0 2px 10px rgba(0,0,0,0.08);">
      <h2 style="margin-top:0; color:#222;">${titulo}</h2>
      <p style="color:#444; font-size:15px;">${texto}</p>
      <p style="color:#999; font-size:12px; margin-top:24px;">Puedes cerrar esta ventana.</p>
    </div>
  </body></html>`;
  return HtmlService.createHtmlOutput(html).setTitle('Sistema Vencimientos DIAN');
}

function paginaFormularioPresentado(nombreHoja, id) {
  const urlActiva = ScriptApp.getService().getUrl(); 

  const html = `
  <!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"></head>
  <body style="font-family:Arial, sans-serif; margin:0; padding:0; background:#f4f4f4;">
    ${headerWebApp()}
    <div style="max-width:480px; margin:40px auto; text-align:center; background:${CONFIG.COLORES_ESTADO['Presentado']}; padding:36px; border-radius:12px; box-shadow:0 2px 10px rgba(0,0,0,0.08);">
      <h2 style="margin-top:0; color:#222;">Registrar presentación</h2>
      <p style="color:#444;">Inserte la fecha en la que se presentó esta obligación:</p>
      
      <form action="${urlActiva}" method="get" target="_top">
        
        <input type="hidden" name="hoja" value="${nombreHoja}">
        <input type="hidden" name="id" value="${id}">
        <input type="hidden" name="accion" value="guardarPresentado">
        <input type="date" name="fecha" required style="padding:10px; font-size:15px; border:1px solid #ccc; border-radius:6px;">
        <br><br>
        <button type="submit" style="padding:12px 28px; background:#2E8B4F; color:#fff; border:none; border-radius:6px; cursor:pointer; font-size:15px; font-weight:bold;">Guardar</button>
      </form>
    </div>
  </body></html>`;
  
  return HtmlService.createHtmlOutput(html).setTitle('Registrar presentación');
}