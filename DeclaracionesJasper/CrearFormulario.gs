/**
 * Crea el formulario de requerimientos completo, con las dos rutas
 * (error / algo nuevo), subida de archivos de cualquier tipo, la hoja
 * de respuestas y la numeración automática REQ-001, REQ-002...
 *
 * Corre crearTodo() UNA sola vez y mira el log: ahí quedan los links.
 */

function crearTodo() {
  const hoja = SpreadsheetApp.create('Requerimientos — Respuestas');
  const form = FormApp.create('Requerimientos — Herramienta de Impuestos Territoriales');

  configurar_(form);
  armarPreguntas_(form);

  form.setDestination(FormApp.DestinationType.SPREADSHEET, hoja.getId());

  ScriptApp.newTrigger('alRecibirRespuesta')
           .forSpreadsheet(hoja.getId())
           .onFormSubmit()
           .create();

  Logger.log('LISTO. Guarda estos tres links:\n');
  Logger.log('1. Para el equipo (responder):\n   ' + form.getPublishedUrl());
  Logger.log('\n2. La hoja — compártela en SOLO LECTURA:\n   ' + hoja.getUrl());
  Logger.log('\n3. Para editar el formulario:\n   ' + form.getEditUrl());
}


function configurar_(form) {
  form.setDescription(
    'Todo lo que necesiten que la herramienta haga se pide por acá. No por ' +
    'correo ni en la reunión: así no se pierde nada y ustedes pueden ver en ' +
    'qué va lo suyo.\n\n' +
    'Al enviarlo les llega un correo con su número de requerimiento.\n\n' +
    'IMPORTANTE: el martes 10 de noviembre a las 5:00 p.m. se cierra la lista ' +
    'de lo que sale el 1 de diciembre. Lo que llegue después entra a la ' +
    'versión 2, que arranca en enero.'
  );

  form.setCollectEmail(true);
  form.setAllowResponseEdits(true);
  form.setPublishingSummary(false);
  form.setConfirmationMessage(
    'Listo. En unos segundos te llega un correo con tu número de ' +
    'requerimiento. Lo revisamos en el próximo weekly.'
  );

  // Solo funciona en dominios de Workspace. Si falla, no pasa nada grave.
  try {
    form.setRequireLogin(true);
  } catch (err) {
    Logger.log('Aviso: no pude restringir el formulario al dominio. ' +
               'Actívalo a mano en Configuración.');
  }
}


function armarPreguntas_(form) {

  // ---------------------------------------------------- SECCIÓN 1

  const tipo = form.addMultipleChoiceItem()
    .setTitle('¿Esto es un error o es algo nuevo?')
    .setHelpText('Un error es algo que ya existe y no está funcionando bien. ' +
                 'Algo nuevo es una función que hoy no tiene.')
    .setRequired(true);

  // ---------------------------------------------------- SECCIÓN 2 · ERROR

  const secError = form.addPageBreakItem()
    .setTitle('Reportar un error')
    .setHelpText('Los errores no esperan al congelamiento. Se arreglan apenas se pueda.');

  form.addParagraphTextItem()
    .setTitle('¿Qué estabas haciendo cuando pasó?')
    .setHelpText('Por ejemplo: "estaba subiendo el soporte de Aguadas y le di a ' +
                 'Enviar". Entre más pasos me des, más rápido lo reproduzco.')
    .setRequired(true);

  form.addParagraphTextItem()
    .setTitle('¿Qué esperabas que pasara?')
    .setRequired(true);

  form.addParagraphTextItem()
    .setTitle('¿Qué pasó en realidad?')
    .setHelpText('Si salió un mensaje de error, cópialo tal cual aunque no se entienda.')
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('¿Pasa siempre o solo a veces?')
    .setChoiceValues([
      'Siempre que lo intento',
      'A veces sí, a veces no',
      'Me pasó una sola vez'
    ])
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('¿Esto te bloquea el trabajo?')
    .setChoiceValues([
      'Sí, no puedo seguir hasta que se arregle',
      'Puedo seguir, pero me toca hacer un rodeo',
      'Es molesto pero no me frena'
    ])
    .setRequired(true);

  archivo_(form,
    'Sube un pantallazo',
    'Un pantallazo vale más que tres párrafos. Si puedes, que se vea la ' +
    'pantalla completa, no solo el error.');

  // ---------------------------------------------------- SECCIÓN 3 · NUEVO

  const secNuevo = form.addPageBreakItem()
    .setTitle('Pedir algo nuevo')
    .setHelpText('Entre mejor contestes esto, más probable es que quede como lo ' +
                 'necesitas a la primera.');

  form.addParagraphTextItem()
    .setTitle('¿Qué necesitas que la herramienta haga y hoy no hace?')
    .setHelpText('Cuéntalo como se lo contarías a un compañero. No tiene que sonar ' +
                 'técnico ni usar palabras de sistemas.')
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('¿En qué parte de la herramienta?')
    .setChoiceValues([
      'Enviar — donde se carga el soporte y se mandan los correos',
      'Respuestas — donde se ve quién ya contestó',
      'Alertas — el tablero de vencimientos',
      'En varias / no estoy seguro'
    ])
    .setRequired(true);

  form.addParagraphTextItem()
    .setTitle('Hoy, ¿cómo resuelves esto?')
    .setHelpText('Por ejemplo: "lo anoto en un Excel aparte", "le escribo a la ' +
                 'oficina por aparte", "no lo resuelvo, lo dejo así". Esto me sirve ' +
                 'para entender qué estoy reemplazando.')
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('¿Cada cuánto necesitas hacer esto?')
    .setHelpText('Algo que haces todos los días pesa más que algo de una vez al año.')
    .setChoiceValues([
      'Varias veces al día',
      'Todos los días',
      'Varias veces por semana',
      'Una vez al mes o por periodo',
      'Pocas veces al año'
    ])
    .setRequired(true);

  form.addCheckboxItem()
    .setTitle('¿A quiénes les sirve esto?')
    .setChoiceValues([
      'Solo a mí',
      'A todo el equipo de territoriales',
      'A las oficinas del banco',
      'A la jefatura',
      'A riesgos / auditoría'
    ])
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('¿Qué pasa si esto no se hace?')
    .setHelpText('Sé honesto. Esto es lo que usamos en el weekly para decidir qué ' +
                 'entra primero, y si todo se marca como crítico, nada lo es.')
    .setChoiceValues([
      'Se puede generar una sanción o un evento de riesgo',
      'Nos toca hacerlo a mano y nos quita tiempo',
      'Sería más cómodo, pero se puede vivir sin eso'
    ])
    .setRequired(true);

  form.addParagraphTextItem()
    .setTitle('¿Cómo sabrías tú que quedó bien hecho?')
    .setHelpText('Termina esta frase: "sabré que quedó listo cuando pueda...". ' +
                 'Por ejemplo: "cuando pueda ver el valor en la lista sin abrir el ' +
                 'correo", "cuando me llegue un solo correo al día y no diez". ' +
                 'Esta es la pregunta más importante: es la que evita que quede ' +
                 'distinto a lo que necesitabas.')
    .setRequired(true);

  archivo_(form,
    'Sube un ejemplo',
    'Un correo real, un pantallazo, o el archivo como lo llenas hoy. ' +
    'Sin un ejemplo me toca adivinar, y ahí es donde queda distinto.');

  form.addMultipleChoiceItem()
    .setTitle('¿Esto tiene que estar para el 1 de diciembre?')
    .setHelpText('Si dices que sí, en el weekly decidimos entre todos qué otra cosa ' +
                 'sale para que esto entre. Nada entra sin que algo salga.')
    .setChoiceValues([
      'Sí — sin esto no podemos arrancar en diciembre',
      'No — puede ser en enero, en la versión 2'
    ])
    .setRequired(true);

  // ---------------------------------------------------- SECCIÓN 4 · CIERRE

  const secCierre = form.addPageBreakItem().setTitle('Para cerrar');

  form.addTextItem()
    .setTitle('¿De dónde salió esto?')
    .setHelpText('Opcional. Por ejemplo: "la reunión del 21 de octubre", "me lo ' +
                 'pidió la oficina de Aguadas". Sirve para rastrearlo después.')
    .setRequired(false);

  form.addParagraphTextItem()
    .setTitle('¿Algo más que deba saber?')
    .setRequired(false);

  // ---------------------------------------------------- LAS BIFURCACIONES
  // Se arman al final, cuando ya existen todas las secciones.

  tipo.setChoices([
    tipo.createChoice('Es un error — algo que ya existe y no funciona como debería', secError),
    tipo.createChoice('Es algo nuevo — necesito que haga algo que hoy no hace', secNuevo)
  ]);

  // Al terminar la ruta de error, saltar al cierre (no a "pedir algo nuevo")
  secError.setGoToPage(secCierre);
}


/** Pregunta de subida de archivos, de cualquier tipo. */
function archivo_(form, titulo, ayuda) {
  try {
    const item = form.addFileUploadItem()
      .setTitle(titulo)
      .setHelpText(ayuda)
      .setRequired(false);

    // Sin setFileTypes queda abierto a cualquier tipo de archivo.
    try { item.setNumberOfFiles(5); } catch (e) {}

  } catch (err) {
    // Algunos dominios no permiten subir archivos en formularios.
    form.addParagraphTextItem()
      .setTitle(titulo + ' — pégalo acá o dime dónde está')
      .setHelpText('No pude habilitar la subida de archivos en este dominio. ' +
                   'Pega el texto, o la ruta del archivo en Drive.')
      .setRequired(false);

    Logger.log('OJO: tu dominio no permite subir archivos en formularios. ' +
               'Puse una pregunta de texto en su lugar ("' + titulo + '").');
  }
}


// ============================================================ NUMERACIÓN

/** Se dispara con cada respuesta: asigna el número y avisa a quien pidió. */
function alRecibirRespuesta(e) {
  const hoja = e.range.getSheet();
  const fila = e.range.getRow();

  const colNumero = columna_(hoja, 'N° requerimiento');
  const colEstado = columna_(hoja, 'Estado');
  columna_(hoja, 'Notas');

  const numero = 'REQ-' + ('00' + (fila - 1)).slice(-3);

  hoja.getRange(fila, colNumero).setValue(numero);
  hoja.getRange(fila, colEstado).setValue('Recibido');

  avisar_(e, numero);
}


/** Busca una columna por su encabezado. Si no existe, la crea al final. */
function columna_(hoja, encabezado) {
  const ultima = hoja.getLastColumn();
  const encabezados = hoja.getRange(1, 1, 1, ultima).getValues()[0];
  const pos = encabezados.indexOf(encabezado);
  if (pos !== -1) return pos + 1;

  hoja.getRange(1, ultima + 1).setValue(encabezado).setFontWeight('bold');
  return ultima + 1;
}


/** Le manda el número a quien lo pidió. */
function avisar_(e, numero) {
  const v = e.namedValues;
  const correo = v['Dirección de correo electrónico'] ||
                 v['Correo electrónico'] ||
                 v['Email Address'];
  if (!correo || !correo[0]) return;

  const esError = (v['¿Esto es un error o es algo nuevo?'] || [''])[0]
                  .indexOf('error') !== -1;

  const que = esError
    ? (v['¿Qué pasó en realidad?'] || [''])[0]
    : (v['¿Qué necesitas que la herramienta haga y hoy no hace?'] || [''])[0];

  const cierre = esError
    ? 'Los errores no esperan al congelamiento: lo miro apenas pueda y te aviso.'
    : 'Lo revisamos en el próximo weekly y ahí decidimos entre todos si entra ' +
      'esta semana.\n\nRecuerda: el martes 10 de noviembre a las 5:00 p.m. se ' +
      'cierra la lista de lo que sale el 1 de diciembre.';

  MailApp.sendEmail({
    to: correo[0],
    subject: numero + ' — recibido',
    body: 'Recibí tu solicitud. Quedó con el número ' + numero + '.\n\n' +
          'Lo que reportaste:\n' + que + '\n\n' +
          cierre + '\n\n' +
          'Si quieres preguntar por esto después, menciona el número ' + numero + '.'
  });
}
