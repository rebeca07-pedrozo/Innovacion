/*************************************************************
 * SISTEMA DE ALERTAS TRIBUTARIAS - Google Apps Script
 * Revisa diariamente las obligaciones y envía correos
 * cuando faltan exactamente 15, 7 o 3 días para el vencimiento.
 *************************************************************/

/* ============================================================
 * CONSTANTES GLOBALES
 * ========================================================== */
const HOJA_OBLIGACIONES = "OBLIGACIONES";
const HOJA_RESPONSABLES = "RESPONSABLES";
const HOJA_HISTORICO    = "HISTORICO_ALERTAS";

// Zona horaria de Colombia para todos los cálculos de fechas.
const ZONA_HORARIA = "America/Bogota";

// Nombre que aparece como remitente del correo.
const REMITENTE_NOMBRE = "Sistema de Alertas Tributarias";

// Reglas de negocio: días restantes -> tipo de alerta.
const REGLAS_ALERTA = {
  15: "Alerta Preventiva",
  7:  "Alerta Próxima",
  3:  "Alerta Crítica"
};

// Colores por tipo de alerta (usados en el HTML del correo).
const COLOR_ALERTA = {
  "Alerta Preventiva": "#2563eb", // azul
  "Alerta Próxima":    "#f59e0b", // ámbar
  "Alerta Crítica":    "#dc2626"  // rojo
};

// Índices de columna (base 0) para acceso a los arreglos de datos.
const COL_OBL = { ID: 0, COMPANIA: 1, NIT: 2, DV: 3, TIPO_DOC: 4, TIPO_OBL: 5, FECHA_VENC: 6 };
const COL_RESP = { COMPANIA: 0, RESPONSABLE: 1, CORREO: 2 };

/* ============================================================
 * FUNCIÓN PRINCIPAL
 * ========================================================== */
/**
 * Punto de entrada del sistema. Recorre todas las obligaciones,
 * determina cuáles requieren alerta hoy y envía los correos.
 * Esta es la función que debe ejecutar el trigger diario.
 */
function revisarVencimientos() {
  try {
    Logger.log("===== INICIO revisarVencimientos =====");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaObl  = obtenerHoja(ss, HOJA_OBLIGACIONES);
    const hojaResp = obtenerHoja(ss, HOJA_RESPONSABLES);
    const hojaHist = obtenerHoja(ss, HOJA_HISTORICO);

    // 1. Cargar responsables (Mapa: COMPAÑIA normalizada -> {responsable, correo}).
    const responsables = obtenerResponsables(hojaResp);

    // 2. Cargar el histórico una sola vez en un Set para validación O(1).
    const clavesEnviadas = cargarClavesHistorico(hojaHist);

    // 3. Leer todas las obligaciones de una sola vez (eficiencia).
    const datosObl = hojaObl.getDataRange().getValues();

    const filasPendientes = []; // Acumula nuevos registros para escribir en lote.
    let enviados = 0;
    let omitidos = 0;

    // Se omite la fila 0 porque es el encabezado.
    for (let i = 1; i < datosObl.length; i++) {
      const fila = datosObl[i];

      // Ignorar filas vacías (sin compañía o sin fecha).
      const compania = String(fila[COL_OBL.COMPANIA] || "").trim();
      const fechaVenc = aFecha(fila[COL_OBL.FECHA_VENC]);
      if (!compania || !fechaVenc) {
        continue;
      }

      // Calcular días restantes respetando la zona horaria.
      const diasRestantes = calcularDiasRestantes(fechaVenc);

      // Solo continuar si coincide exactamente con una regla (15, 7 o 3).
      const tipoAlerta = REGLAS_ALERTA[diasRestantes];
      if (!tipoAlerta) {
        continue;
      }

      const tipoObligacion = String(fila[COL_OBL.TIPO_OBL] || "").trim();

      // Resolver el responsable/correo de la compañía.
      const datosResp = responsables[normalizar(compania)];
      if (!datosResp || !esCorreoValido(datosResp.correo)) {
        Logger.log("OMITIDO (correo inválido o sin responsable): " + compania +
                   " - " + tipoObligacion);
        omitidos++;
        continue;
      }

      // Construir clave única para validar duplicados.
      const clave = construirClave(compania, tipoObligacion, fechaVenc, tipoAlerta);
      if (alertaYaEnviada(clavesEnviadas, clave)) {
        continue; // Ya se envió esta misma alerta; no duplicar.
      }

      // Empaquetar todos los datos para el correo y el registro.
      const datosAlerta = {
        compania: compania,
        nit: String(fila[COL_OBL.NIT] || "").trim(),
        dv: String(fila[COL_OBL.DV] || "").trim(),
        tipoObligacion: tipoObligacion,
        fechaVencimiento: fechaVenc,
        fechaVencimientoTexto: formatearFecha(fechaVenc, "dd/MM/yyyy"),
        diasRestantes: diasRestantes,
        tipoAlerta: tipoAlerta,
        responsable: datosResp.responsable,
        correo: datosResp.correo
      };

      // Enviar el correo (aislado en try/catch para no detener el lote).
      try {
        enviarCorreoAlerta(datosAlerta);
        registrarAlerta(datosAlerta, filasPendientes, clavesEnviadas, clave);
        enviados++;
        Logger.log("ENVIADO: " + compania + " | " + tipoObligacion +
                   " | " + tipoAlerta + " | " + datosResp.correo);
      } catch (errEnvio) {
        Logger.log("ERROR al enviar a " + datosResp.correo + ": " + errEnvio.message);
        omitidos++;
      }
    }

    // 4. Escribir todos los registros nuevos en una sola operación.
    escribirHistorico(hojaHist, filasPendientes);

    Logger.log("RESUMEN -> Enviados: " + enviados + " | Omitidos: " + omitidos);
    Logger.log("===== FIN revisarVencimientos =====");

  } catch (err) {
    Logger.log("ERROR CRÍTICO en revisarVencimientos: " + err.message + "\n" + err.stack);
  }
}

/* ============================================================
 * OBTENER RESPONSABLES
 * ========================================================== */
/**
 * Lee la hoja RESPONSABLES y devuelve un objeto/mapa
 * indexado por nombre de compañía normalizado para búsqueda rápida.
 * @param {Sheet} hojaResp Hoja de responsables.
 * @return {Object} { "COMPAÑIA": {responsable, correo}, ... }
 */
function obtenerResponsables(hojaResp) {
  const mapa = {};
  const datos = hojaResp.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    const compania = String(datos[i][COL_RESP.COMPANIA] || "").trim();
    if (!compania) {
      continue; // Ignorar filas vacías.
    }
    mapa[normalizar(compania)] = {
      responsable: String(datos[i][COL_RESP.RESPONSABLE] || "").trim(),
      correo: String(datos[i][COL_RESP.CORREO] || "").trim()
    };
  }
  return mapa;
}

/* ============================================================
 * CALCULAR DÍAS RESTANTES
 * ========================================================== */
/**
 * Calcula la diferencia en días calendario entre hoy y la fecha
 * de vencimiento, normalizando ambas a medianoche en la zona horaria
 * configurada para evitar errores por horas o DST.
 * @param {Date} fechaVencimiento Fecha de vencimiento.
 * @return {number} Días restantes (positivo = futuro).
 */
function calcularDiasRestantes(fechaVencimiento) {
  const hoyStr = formatearFecha(new Date(), "yyyy-MM-dd");
  const venStr = formatearFecha(fechaVencimiento, "yyyy-MM-dd");

  // Se interpretan como UTC para que la resta sea exacta en días.
  const hoyUTC = new Date(hoyStr + "T00:00:00Z");
  const venUTC = new Date(venStr + "T00:00:00Z");

  const MS_POR_DIA = 24 * 60 * 60 * 1000;
  return Math.round((venUTC - hoyUTC) / MS_POR_DIA);
}

/* ============================================================
 * VALIDAR DUPLICADOS
 * ========================================================== */
/**
 * Indica si una alerta ya fue registrada previamente.
 * @param {Set} clavesEnviadas Conjunto de claves del histórico.
 * @param {string} clave Clave única de la alerta.
 * @return {boolean} true si ya existe.
 */
function alertaYaEnviada(clavesEnviadas, clave) {
  return clavesEnviadas.has(clave);
}

/**
 * Carga el histórico existente en un Set de claves únicas
 * para validar duplicados en O(1).
 * @param {Sheet} hojaHist Hoja HISTORICO_ALERTAS.
 * @return {Set<string>} Conjunto de claves ya registradas.
 */
function cargarClavesHistorico(hojaHist) {
  const claves = new Set();
  const datos = hojaHist.getDataRange().getValues();

  // Columnas del histórico: A=FECHA_ENVIO, B=COMPAÑIA, C=TIPO_OBLIGACION,
  // D=FECHA_VENCIMIENTO, E=TIPO_ALERTA, F=CORREO
  for (let i = 1; i < datos.length; i++) {
    const compania  = datos[i][1];
    const tipoObl   = datos[i][2];
    const fechaVenc = aFecha(datos[i][3]);
    const tipoAlerta = datos[i][4];

    if (!compania || !fechaVenc) {
      continue;
    }
    claves.add(construirClave(compania, tipoObl, fechaVenc, tipoAlerta));
  }
  return claves;
}

/* ============================================================
 * REGISTRAR ALERTA
 * ========================================================== */
/**
 * Prepara una nueva fila para el histórico y la agrega al lote
 * pendiente. También añade su clave al Set para evitar duplicados
 * dentro de la misma ejecución.
 * @param {Object} datos Datos de la alerta.
 * @param {Array[]} filasPendientes Acumulador de filas a escribir.
 * @param {Set} clavesEnviadas Set de claves enviadas.
 * @param {string} clave Clave única de esta alerta.
 */
function registrarAlerta(datos, filasPendientes, clavesEnviadas, clave) {
  const fila = [
    formatearFecha(new Date(), "dd/MM/yyyy HH:mm:ss"), // FECHA_ENVIO
    datos.compania,                                    // COMPAÑIA
    datos.tipoObligacion,                              // TIPO_OBLIGACION
    datos.fechaVencimientoTexto,                       // FECHA_VENCIMIENTO
    datos.tipoAlerta,                                  // TIPO_ALERTA
    datos.correo                                       // CORREO
  ];
  filasPendientes.push(fila);
  clavesEnviadas.add(clave); // Evita doble envío en la misma corrida.
}

/**
 * Escribe todas las filas nuevas en HISTORICO_ALERTAS en una sola
 * operación (eficiente para grandes volúmenes).
 * @param {Sheet} hojaHist Hoja del histórico.
 * @param {Array[]} filas Filas a escribir.
 */
function escribirHistorico(hojaHist, filas) {
  if (!filas || filas.length === 0) {
    return;
  }
  const inicio = hojaHist.getLastRow() + 1;
  hojaHist.getRange(inicio, 1, filas.length, filas[0].length).setValues(filas);
}

/* ============================================================
 * ENVIAR CORREO
 * ========================================================== */
/**
 * Envía el correo de alerta en formato HTML profesional.
 * @param {Object} datos Datos completos de la alerta.
 */
function enviarCorreoAlerta(datos) {
  const asunto = "[ALERTA TRIBUTARIA] " + datos.tipoObligacion + " - " + datos.compania;
  const cuerpoHtml = construirCuerpoHTML(datos);

  MailApp.sendEmail({
    to: datos.correo,
    subject: asunto,
    htmlBody: cuerpoHtml,
    name: REMITENTE_NOMBRE
  });
}

/**
 * Construye el cuerpo HTML del correo con una tabla informativa.
 * @param {Object} datos Datos de la alerta.
 * @return {string} HTML del correo.
 */
function construirCuerpoHTML(datos) {
  const color = COLOR_ALERTA[datos.tipoAlerta] || "#374151";
  const nitCompleto = datos.dv ? (datos.nit + "-" + datos.dv) : datos.nit;

  return ''
    + '<div style="font-family:Arial,Helvetica,sans-serif;max-width:640px;margin:auto;'
    + 'border:1px solid #e5e7eb;border-radius:8px;overflow:hidden;">'
    +   '<div style="background:' + color + ';color:#ffffff;padding:18px 24px;">'
    +     '<h2 style="margin:0;font-size:18px;">' + datos.tipoAlerta + '</h2>'
    +     '<p style="margin:4px 0 0;font-size:13px;opacity:.9;">Alerta Tributaria Automática</p>'
    +   '</div>'
    +   '<div style="padding:24px;color:#111827;font-size:14px;">'
    +     '<p>Estimado(a) <strong>' + (datos.responsable || "responsable") + '</strong>,</p>'
    +     '<p>Le informamos que la siguiente obligación está próxima a vencer. '
    +     'Faltan <strong>' + datos.diasRestantes + ' día(s)</strong> para su vencimiento.</p>'
    +     '<table style="width:100%;border-collapse:collapse;margin-top:16px;font-size:14px;">'
    +       fila("Compañía", datos.compania)
    +       fila("NIT", nitCompleto)
    +       fila("Tipo de obligación", datos.tipoObligacion)
    +       fila("Fecha de vencimiento", datos.fechaVencimientoTexto)
    +       fila("Días restantes", String(datos.diasRestantes))
    +       fila("Tipo de alerta", datos.tipoAlerta)
    +     '</table>'
    +     '<p style="margin-top:24px;color:#6b7280;font-size:12px;">'
    +     'Este es un mensaje automático generado por el Sistema de Alertas Tributarias. '
    +     'Por favor no responda a este correo.</p>'
    +   '</div>'
    + '</div>';

  // Función interna para generar cada fila de la tabla.
  function fila(etiqueta, valor) {
    return '<tr>'
      + '<td style="padding:8px 12px;border:1px solid #e5e7eb;background:#f9fafb;'
      + 'font-weight:bold;width:45%;">' + etiqueta + '</td>'
      + '<td style="padding:8px 12px;border:1px solid #e5e7eb;">' + valor + '</td>'
      + '</tr>';
  }
}

/* ============================================================
 * FUNCIÓN DE PRUEBA MANUAL
 * ========================================================== */
/**
 * Ejecuta el flujo completo de forma manual desde el editor de
 * Apps Script. Útil para validar el funcionamiento y revisar logs.
 * Importante: SÍ envía correos reales si hay coincidencias.
 */
function pruebaAlertas() {
  Logger.log(">>> Ejecutando prueba manual de alertas...");
  revisarVencimientos();
  Logger.log(">>> Prueba finalizada. Revise el menú Ejecuciones / Logs.");
}

/* ============================================================
 * UTILIDADES AUXILIARES
 * ========================================================== */
/**
 * Obtiene una hoja por nombre o lanza un error claro si no existe.
 */
function obtenerHoja(ss, nombre) {
  const hoja = ss.getSheetByName(nombre);
  if (!hoja) {
    throw new Error("No se encontró la hoja: " + nombre);
  }
  return hoja;
}

/**
 * Construye la clave única para validar duplicados.
 * Combina compañía, tipo de obligación, fecha de vencimiento y tipo de alerta.
 */
function construirClave(compania, tipoObligacion, fechaVencDate, tipoAlerta) {
  const fechaKey = formatearFecha(fechaVencDate, "yyyy-MM-dd");
  return [
    normalizar(compania),
    normalizar(tipoObligacion),
    fechaKey,
    normalizar(tipoAlerta)
  ].join("||");
}

/**
 * Normaliza texto para comparaciones (sin espacios extra, en mayúsculas).
 */
function normalizar(valor) {
  return String(valor || "").trim().toUpperCase();
}

/**
 * Formatea una fecha usando la zona horaria configurada.
 */
function formatearFecha(fecha, patron) {
  return Utilities.formatDate(fecha, ZONA_HORARIA, patron);
}

/**
 * Convierte un valor de celda a objeto Date.
 * Acepta objetos Date o cadenas en formato dd/MM/yyyy.
 * @return {Date|null} Fecha válida o null.
 */
function aFecha(valor) {
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return valor;
  }
  if (typeof valor === "string") {
    const s = valor.trim();
    const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); // dd/MM/yyyy
    if (m) {
      return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
    }
    const d = new Date(s);
    if (!isNaN(d.getTime())) {
      return d;
    }
  }
  return null;
}

/**
 * Valida un correo electrónico con una expresión regular básica.
 */
function esCorreoValido(correo) {
  if (!correo) {
    return false;
  }
  const regex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return regex.test(String(correo).trim());
}