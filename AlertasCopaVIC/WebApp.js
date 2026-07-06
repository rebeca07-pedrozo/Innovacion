/**
 * BLOQUE 4: WEB APP - BOTONES DE ESTADO + SEMÁFORO
 * -------------------------------------------------------------
 * Recibe clics desde el correo. Parámetros esperados en la URL:
 *   ?hoja=LOAD (o TRANSFORM2 para pruebas)
 *   &id=OBL-XXXXXXXXXX
 *   &accion=notificado | enproceso | formPresentado | guardarPresentado
 *   &fecha=YYYY-MM-DD  (solo para guardarPresentado)
 *
 * DESPLIEGUE: Implementar > Nueva implementación > Aplicación web
 *   - Ejecutar como: Yo (tu cuenta)
 *   - Quién tiene acceso: Cualquier usuario (para que funcione sin login)
 *   Copia la URL resultante y pégala en CONFIG.URL_WEBAPP en Config.gs.
 */

function doGet(e) {
  const nombreHoja = e.parameter.hoja || CONFIG.HOJA_LOAD;
  const id = e.parameter.id;
  const accion = e.parameter.accion;

  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = libro.getSheetByName(nombreHoja);

  if (!hoja || !id) {
    return HtmlService.createHtmlOutput(paginaMensaje('Enlace inválido', 'No se pudo identificar la obligación o la hoja.', '#F2F2F2'));
  }

  const resultado = buscarFilaPorId(hoja, id);
  if (!resultado) {
    return HtmlService.createHtmlOutput(paginaMensaje('No encontrado', 'Esta obligación ya no existe en "' + nombreHoja + '" (puede que el ETL se haya vuelto a correr).', '#F2F2F2'));
  }

  const mapa = resultado.mapa;
  const numeroFila = resultado.numeroFila;

  if (accion === 'notificado') {
    marcarFecha(hoja, numeroFila, mapa, 'Fecha Notificado');
    return HtmlService.createHtmlOutput(paginaMensaje('¡Listo!', 'Marcaste esta obligación como Notificado.', CONFIG.COLORES_ESTADO['Notificado']));
  }

  if (accion === 'enproceso') {
    marcarFecha(hoja, numeroFila, mapa, 'Fecha En Proceso');
    return HtmlService.createHtmlOutput(paginaMensaje('¡Listo!', 'Marcaste esta obligación como En proceso.', CONFIG.COLORES_ESTADO['En proceso']));
  }

  if (accion === 'formPresentado') {
    return HtmlService.createHtmlOutput(paginaFormularioPresentado(nombreHoja, id));
  }

  if (accion === 'guardarPresentado') {
    const fechaTexto = e.parameter.fecha;
    if (!fechaTexto) {
      return HtmlService.createHtmlOutput(paginaMensaje('Falta la fecha', 'Debes ingresar la fecha de presentación.', '#F2F2F2'));
    }
    const fecha = new Date(fechaTexto + 'T00:00:00');
    if (isNaN(fecha.getTime())) {
      return HtmlService.createHtmlOutput(paginaMensaje('Fecha inválida', 'El formato de fecha no es válido.', '#F2F2F2'));
    }
    hoja.getRange(numeroFila, mapa['Fecha Presentado'] + 1).setValue(fecha);
    recalcularEstadoFila(hoja, numeroFila, mapa);
    return HtmlService.createHtmlOutput(paginaMensaje('¡Presentado registrado!', 'Se guardó la fecha de presentación: ' + fechaTexto, CONFIG.COLORES_ESTADO['Presentado']));
  }

  // Sin acción: muestra estado actual
  const estadoActual = valorPorColumna(resultado.valores, mapa, 'Estado Actual');
  return HtmlService.createHtmlOutput(paginaMensaje('Estado actual', 'Esta obligación está: ' + estadoActual, CONFIG.COLORES_ESTADO[estadoActual] || '#fff'));
}

/**
 * Marca la fecha de hoy en la columna indicada (si aún no tenía fecha),
 * recalcula el Estado Actual y aplica el color de semáforo.
 */
function marcarFecha(hoja, numeroFila, mapa, nombreColumna) {
  const valorActual = hoja.getRange(numeroFila, mapa[nombreColumna] + 1).getValue();
  if (!valorActual) {
    hoja.getRange(numeroFila, mapa[nombreColumna] + 1).setValue(new Date());
  }
  recalcularEstadoFila(hoja, numeroFila, mapa);
}

/**
 * Recalcula Estado Actual y Extemporaneidad para una fila, y colorea
 * la celda de Estado Actual según el semáforo.
 */
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

  const semaforo = calcularSemaforo(estado, diasRestantes);
  hoja.getRange(numeroFila, mapa['Semáforo'] + 1).setValue(semaforo);

  if (presentado && fechaMaxima instanceof Date) {
    const fechaPres = presentado instanceof Date ? presentado : new Date(presentado);
    const extemporaneidad = Math.round((fechaPres - fechaMaxima) / 86400000);
    hoja.getRange(numeroFila, mapa['Extemporaneidad (días)'] + 1).setValue(extemporaneidad);
  }

  const colorCelda = CONFIG.COLORES_ESTADO[estado] || '#FFFFFF';
  hoja.getRange(numeroFila, mapa['Estado Actual'] + 1).setBackground(colorCelda);
}

function paginaMensaje(titulo, texto, color) {
  return `
  <div style="font-family:Arial, sans-serif; max-width:480px; margin:60px auto; text-align:center; background:${color}; padding:30px; border-radius:8px;">
    <h2>${titulo}</h2>
    <p>${texto}</p>
  </div>`;
}

function paginaFormularioPresentado(nombreHoja, id) {
  return `
  <div style="font-family:Arial, sans-serif; max-width:480px; margin:60px auto; text-align:center; background:${CONFIG.COLORES_ESTADO['Presentado']}; padding:30px; border-radius:8px;">
    <h2>Registrar presentación</h2>
    <p>Inserte la fecha en la que se presentó esta obligación:</p>
    <form action="${CONFIG.URL_WEBAPP}" method="get">
      <input type="hidden" name="hoja" value="${nombreHoja}">
      <input type="hidden" name="id" value="${id}">
      <input type="hidden" name="accion" value="guardarPresentado">
      <input type="date" name="fecha" required style="padding:8px; font-size:14px;">
      <br><br>
      <button type="submit" style="padding:10px 20px; background:#2E8B4F; color:#fff; border:none; border-radius:5px; cursor:pointer;">Guardar</button>
    </form>
  </div>`;
}