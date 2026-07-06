
function obtenerBlobLogo() {
  return Utilities.newBlob(
    Utilities.base64Decode(CONFIG.LOGO_BASE64),
    'image/png',
    'logo_davivienda.png'
  );
}
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
  const rango = obtenerRangoSemanaActual();

  const grupos = {}; // email -> { nombre, filas: [{numeroFila, valores}] }

  datos.forEach((fila, idx) => {
    const estado = valorPorColumna(fila, mapa, 'Estado Actual');
    const email = valorPorColumna(fila, mapa, 'Encargado Email');
    const fechaMaxima = valorPorColumna(fila, mapa, 'Fecha máxima de presentación');

    const dentroDeSemana = fechaMaxima instanceof Date && fechaMaxima >= rango.inicio && fechaMaxima <= rango.fin;
    const yaVencida = fechaMaxima instanceof Date && fechaMaxima < rango.inicio;
    if (!email || estado === CONFIG.ESTADOS.PRESENTADO || (!dentroDeSemana && !yaVencida)) return;

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
      subject: 'Vencimientos DIAN esta semana - ' + grupo.filas.length + ' obligación(es)',
      htmlBody: htmlCorreo,
      inlineImages: {
        logoDavivienda: obtenerBlobLogo()
      }
    });

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
 * Calcula el lunes y domingo de la semana actual, y el texto
 * "5 - 11 de julio" para el header del correo.
 */
function obtenerRangoSemanaActual() {
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  const diaSemana = hoy.getDay();
  const offsetLunes = diaSemana === 0 ? -6 : 1 - diaSemana;

  const lunes = new Date(hoy);
  lunes.setDate(hoy.getDate() + offsetLunes);

  const domingo = new Date(lunes);
  domingo.setDate(lunes.getDate() + 6);
  domingo.setHours(23, 59, 59, 999);

  const meses = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  const texto = lunes.getMonth() === domingo.getMonth()
    ? `${lunes.getDate()} - ${domingo.getDate()} de ${meses[lunes.getMonth()]}`
    : `${lunes.getDate()} de ${meses[lunes.getMonth()]} - ${domingo.getDate()} de ${meses[domingo.getMonth()]}`;

  return { inicio: lunes, fin: domingo, texto: texto };
}

/**
 * HTML completo del correo: header CON LOGO E ÍCONO DE CALENDARIO,
 * más una tarjeta por cada obligación pendiente esta semana.
 */
function construirHtmlCorreoEncargado(nombreEncargado, filas, mapa, nombreHoja) {
  const rango = obtenerRangoSemanaActual();
  const tarjetas = filas.map(item => construirTarjetaObligacion(item, mapa, nombreHoja)).join('<div style="height:20px;"></div>');

  return `
  <div style="font-family: Arial, sans-serif; max-width: 650px; margin: 0 auto; border:1px solid #eee; border-radius:10px; overflow:hidden;">
    <table style="width:100%; background:linear-gradient(90deg, #FCEFEA, #FBD9CE); border-collapse:collapse;">
      <tr>
        <td style="padding:16px 24px; text-align:left; vertical-align:middle;">
          <img src="cid:logoDavivienda" style="height:26px; vertical-align:middle;" alt="Davivienda">
        </td>
        <td style="padding:16px 24px; text-align:right; vertical-align:middle; white-space:nowrap;">
          <img src="${CONFIG.URL_ICONO_CALENDARIO}" style="height:16px; vertical-align:middle; margin-right:6px;" alt="Calendario">
          <span style="color:#555; font-size:13px; vertical-align:middle;">Semana del ${rango.texto}</span>
        </td>
      </tr>
    </table>
    <div style="padding:24px;">
      <h1 style="font-size:22px; margin:0 0 16px 0;">Informe de resumen vencimientos semanales</h1>
      <p style="border-left:4px solid #D0021B; padding-left:10px; margin-bottom:20px;">
        Vencimiento próximo ante la DIAN. Gestiona esta obligación antes de la fecha límite para evitar sanciones.
      </p>
      ${tarjetas}
    </div>
  </div>`;
}

/**
 * Una tarjeta = tabla de 1 obligación + los 3 botones de color.
 * Cada botón lleva el ID de ESA obligación específica.
 */
function construirTarjetaObligacion(item, mapa, nombreHoja) {
  const fila = item.valores;
  const id = valorPorColumna(fila, mapa, 'ID');
  const compania = valorPorColumna(fila, mapa, 'Compañía (Normalizada)');
  const nit = valorPorColumna(fila, mapa, 'NIT');
  const impuesto = valorPorColumna(fila, mapa, 'Impuesto');
  const municipio = valorPorColumna(fila, mapa, 'Municipio');
  const fechaMaxima = valorPorColumna(fila, mapa, 'Fecha máxima de presentación');
  const diasRestantes = valorPorColumna(fila, mapa, 'Días Restantes');

  const fechaTexto = fechaMaxima instanceof Date
    ? Utilities.formatDate(fechaMaxima, Session.getScriptTimeZone(), 'dd/MM/yyyy')
    : '(fecha por confirmar)';
  const impuestoTexto = municipio ? impuesto + ' - ' + municipio : impuesto;
  const base = CONFIG.URL_WEBAPP + '?hoja=' + encodeURIComponent(nombreHoja) + '&id=' + encodeURIComponent(id);

  return `
  <div style="border:1px solid #e0e0e0; border-radius:8px; padding:16px;">
    <p style="font-weight:bold; margin:0 0 10px 0;">OBLIGATORIO:</p>
    <p style="margin:0 0 12px 0; color:#555;">Indícanos en qué estado se encuentra:</p>
    <table style="width:100%; border-collapse:collapse; font-size:13px; margin-bottom:16px;">
      <tr style="background:#f7f7f7;">
        <th style="padding:8px; border:1px solid #ddd; text-align:left;">Compañía</th>
        <th style="padding:8px; border:1px solid #ddd; text-align:left;">NIT</th>
        <th style="padding:8px; border:1px solid #ddd; text-align:left;">Obligación</th>
        <th style="padding:8px; border:1px solid #ddd; text-align:left;">Fecha límite</th>
        <th style="padding:8px; border:1px solid #ddd; text-align:center;">Días restantes</th>
      </tr>
      <tr>
        <td style="padding:8px; border:1px solid #ddd;">${compania}</td>
        <td style="padding:8px; border:1px solid #ddd;">${nit}</td>
        <td style="padding:8px; border:1px solid #ddd;">${impuestoTexto}</td>
        <td style="padding:8px; border:1px solid #ddd;">${fechaTexto}</td>
        <td style="padding:8px; border:1px solid #ddd; text-align:center;">${diasRestantes}</td>
      </tr>
    </table>
    <table style="width:100%; border-collapse:separate; border-spacing:8px 0;">
      <tr>
        <td style="width:33%; background:#D9E8F5; border-radius:6px; text-align:center; padding:12px 0;">
          <a href="${base}&accion=notificado" style="color:#2C5F8A; font-weight:bold; text-decoration:none;">Notificado</a>
        </td>
        <td style="width:33%; background:#FCE8D5; border-radius:6px; text-align:center; padding:12px 0;">
          <a href="${base}&accion=enproceso" style="color:#C97A1F; font-weight:bold; text-decoration:none;">En proceso</a>
        </td>
        <td style="width:33%; background:#DCF0E1; border-radius:6px; text-align:center; padding:12px 0;">
          <a href="${base}&accion=formPresentado" style="color:#2E8B4F; font-weight:bold; text-decoration:none;">Presentado</a>
        </td>
      </tr>
    </table>
  </div>`;
}


/**
 * FUNCIÓN DE DIAGNÓSTICO - NO es parte del sistema final.
 * Decodifica el base64 guardado en CONFIG.LOGO_BASE64 y lo guarda
 * como un archivo de imagen real en tu Google Drive. Si al abrirlo
 * ahí se ve roto/en blanco, el base64 está corrupto (lo más probable
 * es que se haya cortado al copiar/pegar, por ser tan largo).
 * Si se ve bien en Drive, el problema está en otra parte del envío.
 */
function diagnosticarLogoBase64() {
  const blob = Utilities.newBlob(
    Utilities.base64Decode(CONFIG.LOGO_BASE64),
    'image/png',
    'diagnostico_logo.png'
  );
  const archivo = DriveApp.createFile(blob);
  Logger.log('Archivo creado. Ábrelo aquí: ' + archivo.getUrl());
  Logger.log('Tamaño en bytes: ' + blob.getBytes().length);
  return archivo.getUrl();
}


/**
 * Genera el base64 de una imagen guardada en Drive, dado su ID de
 * archivo. Lo escribe en una hoja (no en el Logger) para que puedas
 * copiarlo completo sin que se corte por ser un string largo.
 *
 * CÓMO USAR:
 * 1. Reemplaza el ID de abajo por el ID real de tu archivo en Drive.
 * 2. Ejecuta esta función.
 * 3. Ve a la hoja "LOGO_BASE64_TEMP", celda A3 -> ahí está el string
 *    completo listo para copiar y pegar en CONFIG.LOGO_BASE64.
 */
function generarBase64DesdeDrive() {
  const idArchivo = 'PEGA_AQUI_EL_ID_DE_TU_ARCHIVO_EN_DRIVE';

  const archivo = DriveApp.getFileById(idArchivo);
  const blob = archivo.getBlob();
  const bytes = blob.getBytes();
  const base64 = Utilities.base64Encode(bytes);

  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName('LOGO_BASE64_TEMP') || libro.insertSheet('LOGO_BASE64_TEMP');
  hoja.clear();
  hoja.getRange('A1').setValue('Nombre archivo: ' + archivo.getName());
  hoja.getRange('A2').setValue('Tipo MIME: ' + blob.getContentType());
  hoja.getRange('A3').setValue('Tamaño en bytes: ' + bytes.length);
  hoja.getRange('A4').setValue(base64);

  Logger.log('Listo. Revisa la hoja LOGO_BASE64_TEMP, celda A4. Bytes: ' + bytes.length);
  SpreadsheetApp.getActiveSpreadsheet().toast('Base64 generado en hoja LOGO_BASE64_TEMP, celda A4 (' + bytes.length + ' bytes)', 'Listo', 8);

  return base64;
}