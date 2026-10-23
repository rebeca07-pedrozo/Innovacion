/**
 * Captura de requerimientos — Herramienta de Impuestos Territoriales
 *
 * PUESTA EN MARCHA
 *   1. Crea un Google Sheets nuevo.
 *   2. Extensiones > Apps Script. Pega este archivo como "Codigo".
 *   3. Crea un archivo HTML llamado "Index" y pega el otro archivo.
 *   4. Corre instalar() una vez.
 *   5. Implementar > Nueva implementación > Aplicación web
 *        Ejecutar como: Yo
 *        Quién tiene acceso: Cualquier usuario de Davivienda
 *      Copia el link y repártelo.
 */

const HOJA = 'REQUERIMIENTOS';
const CARPETA = 'Requerimientos — Adjuntos';

const COLUMNAS = [
  'id', 'fecha', 'solicitante', 'titulo', 'modulo', 'descripcion',
  'justificacion', 'criterio', 'prioridad', 'entrega', 'frecuencia',
  'adjuntos', 'observaciones', 'estado', 'notas'
];


function instalar() {
  const libro = SpreadsheetApp.getActive();
  let hoja = libro.getSheetByName(HOJA);
  if (!hoja) hoja = libro.insertSheet(HOJA);

  hoja.getRange(1, 1, 1, COLUMNAS.length)
      .setValues([COLUMNAS])
      .setFontWeight('bold')
      .setBackground('#E30518')
      .setFontColor('#FFFFFF');
  hoja.setFrozenRows(1);

  carpeta_();
  Logger.log('Listo. Ya puedes implementar la aplicación web.');
}


function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Requerimientos — Impuestos Territoriales')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


/** Carpeta de Drive donde quedan los adjuntos. La crea si no existe. */
function carpeta_() {
  const busca = DriveApp.getFoldersByName(CARPETA);
  return busca.hasNext() ? busca.next() : DriveApp.createFolder(CARPETA);
}


function hoja_() {
  const h = SpreadsheetApp.getActive().getSheetByName(HOJA);
  if (!h) throw new Error('Falta la hoja ' + HOJA + '. Corre instalar().');
  return h;
}


/** Quién está usando la aplicación. */
function quienSoy() {
  try {
    return Session.getActiveUser().getEmail() || '';
  } catch (err) {
    return '';
  }
}


/**
 * Guarda un requerimiento. Lo llama la interfaz.
 * datos.archivos = [{nombre, tipo, contenido}] en base64
 */
function guardarRequerimiento(datos) {
  const candado = LockService.getScriptLock();
  candado.waitLock(20000);

  try {
    const hoja = hoja_();
    const fila = hoja.getLastRow() + 1;
    const id = 'REQ-' + ('00' + (fila - 1)).slice(-3);

    const links = guardarArchivos_(datos.archivos, id);

    const registro = {
      id: id,
      fecha: new Date(),
      solicitante: datos.solicitante || quienSoy(),
      titulo: datos.titulo,
      modulo: datos.modulo,
      descripcion: datos.descripcion,
      justificacion: datos.justificacion,
      criterio: datos.criterio,
      prioridad: datos.prioridad,
      entrega: datos.entrega,
      frecuencia: datos.frecuencia || '',
      adjuntos: links.join('\n'),
      observaciones: datos.observaciones || '',
      estado: 'Nuevo',
      notas: ''
    };

    hoja.getRange(fila, 1, 1, COLUMNAS.length)
        .setValues([COLUMNAS.map(function (c) { return registro[c]; })]);

    avisar_(registro);
    return id;

  } finally {
    candado.releaseLock();
  }
}


/** Convierte los archivos de base64 a Drive y devuelve los links. */
function guardarArchivos_(archivos, id) {
  if (!archivos || !archivos.length) return [];

  const destino = carpeta_();
  return archivos.map(function (a) {
    const blob = Utilities.newBlob(
      Utilities.base64Decode(a.contenido), a.tipo, id + ' — ' + a.nombre
    );
    const archivo = destino.createFile(blob);
    return archivo.getUrl();
  });
}


/** Le confirma al solicitante con su número. */
function avisar_(r) {
  if (!r.solicitante || r.solicitante.indexOf('@') === -1) return;

  MailApp.sendEmail({
    to: r.solicitante,
    subject: r.id + ' · ' + r.titulo,
    body:
      'Tu requerimiento quedó registrado con el número ' + r.id + '.\n\n' +
      'Título: ' + r.titulo + '\n' +
      'Módulo: ' + r.modulo + '\n' +
      'Prioridad: ' + r.prioridad + '\n' +
      'Entrega esperada: ' + r.entrega + '\n\n' +
      'Criterio de aceptación que definiste:\n' + r.criterio + '\n\n' +
      'Lo priorizamos en el próximo weekly. Si quieres preguntar por él, ' +
      'menciona el número ' + r.id + '.\n\n' +
      'Recuerda: el martes 10 de noviembre a las 5:00 p.m. se cierra el ' +
      'alcance de la versión que sale el 1 de diciembre.'
  });
}


/** Lista los requerimientos para la tabla de la interfaz. */
function listarRequerimientos() {
  const hoja = hoja_();
  if (hoja.getLastRow() < 2) return [];

  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, COLUMNAS.length).getValues();

  return datos.map(function (f) {
    const r = {};
    COLUMNAS.forEach(function (c, i) { r[c] = f[i]; });
    r.fecha = r.fecha instanceof Date
      ? Utilities.formatDate(r.fecha, 'America/Bogota', 'dd/MM/yyyy')
      : r.fecha;
    return r;
  }).reverse();
}


/** Números para las tarjetas de arriba. */
function resumen() {
  const lista = listarRequerimientos();
  return {
    total: lista.length,
    diciembre: lista.filter(function (r) {
      return String(r.entrega).indexOf('diciembre') !== -1;
    }).length,
    imprescindibles: lista.filter(function (r) {
      return String(r.prioridad).indexOf('Imprescindible') !== -1;
    }).length,
    nuevos: lista.filter(function (r) { return r.estado === 'Nuevo'; }).length
  };
}
