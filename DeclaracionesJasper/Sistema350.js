/**
 * ============================================================
 * SISTEMA 350 — generar, cargar y consolidar
 *
 * La plantilla base es la fuente de verdad: los números de
 * renglón se leen del propio archivo.
 *
 * Requiere el servicio avanzado Drive API v3.
 * ============================================================
 */

// ---------- Carpetas ----------
const CARPETA_BASE       = "PEGAR_ID_1_PLANTILLA_BASE";
const CARPETA_GENERADAS  = "PEGAR_ID_2_GENERADAS";
const CARPETA_DAVIVIENDA = "PEGAR_ID_3_CARGAS_DAVIVIENDA";
const CARPETA_DAVIBANK   = "PEGAR_ID_3_CARGAS_DAVIBANK";
const CARPETA_CONSOLIDADOS = "PEGAR_ID_4_CONSOLIDADOS";

// ---------- Entidades ----------
const ENTIDADES = {
  DAVIVIENDA: {
    nombre: "BANCO DAVIVIENDA S.A.",
    nit: "860034313",
    dv: "7",
    dirSeccional: "31",
    carpeta: CARPETA_DAVIVIENDA
  },
  DAVIBANK: {
    nombre: "DAVIBANK",
    nit: "",                    // completar
    dv: "",
    dirSeccional: "",
    carpeta: CARPETA_DAVIBANK
  }
};

// ---------- Mapa de la plantilla ----------
const COLS = [19, 27, 32, 47];              // S, AA, AF, AU
const CALCULADOS = [130, 134, 136, 138];    // los resuelve la fórmula

const ENC = {
  anio: ["F4", "H4", "J4", "L4"],
  periodo: ["Z4", "AA4"],
  nit: "D11",
  dv: "R11",
  razonSocial: "D13",
  dirSeccional: "AZ13"
};

const DET_INI = 11;
const DET_FIN = 55;
const DET = { convenio:3, concepto:17, persona:20, pais:23,
              codPais:31, base:34, tarifa:47, retencion:50 };

const MESES = ["enero","febrero","marzo","abril","mayo","junio",
               "julio","agosto","septiembre","octubre","noviembre","diciembre"];


// ============================================================
// INTERFAZ
// ============================================================

function doGet() {
  return HtmlService.createTemplateFromFile("Interfaz")
    .evaluate()
    .setTitle("Sistema de declaraciones")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}


function obtenerDatosIniciales() {
  const hoy = new Date();
  const anios = [];
  for (let a = hoy.getFullYear(); a >= hoy.getFullYear() - 2; a--) {
    anios.push({ valor: String(a), texto: String(a) });
  }

  const periodos = MESES.map(function (m, i) {
    return { valor: ("0" + (i + 1)).slice(-2), texto: (i + 1) + " — " + m };
  });

  const entidades = Object.keys(ENTIDADES).map(function (k) {
    return { codigo: k, nombre: ENTIDADES[k].nombre };
  });

  return {
    usuario: Session.getActiveUser().getEmail(),
    entidades: entidades,
    anios: anios,
    periodos: periodos
  };
}


/**
 * Genera la plantilla de una entidad y periodo, y devuelve el enlace.
 */
function obtenerPlantilla(codEntidad, anio, periodo) {
  const entidad = ENTIDADES[codEntidad];
  if (!entidad) throw new Error("Entidad no reconocida: " + codEntidad);

  const base = archivoBase();
  const copia = Drive.Files.copy(
    { name: "TMP", mimeType: MimeType.GOOGLE_SHEETS }, base.getId()
  );

  const libro = SpreadsheetApp.openById(copia.id);
  escribirEncabezado(libro, entidad, anio, periodo);
  SpreadsheetApp.flush();

  const nombre = "350_" + codEntidad + "_" + anio + "-" + periodo;
  const archivo = DriveApp.getFileById(copia.id);
  archivo.setName(nombre);

  DriveApp.getFolderById(CARPETA_GENERADAS).addFile(archivo);
  DriveApp.getRootFolder().removeFile(archivo);
  archivo.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    nombre: nombre,
    entidad: entidad.nombre,
    periodo: anio + "-" + periodo,
    urlDescarga: "https://docs.google.com/spreadsheets/d/" + copia.id +
                 "/export?format=xlsx"
  };
}


/**
 * Recibe el archivo diligenciado y lo guarda en la carpeta de la entidad.
 */
function recibirCarga(datosArchivo, nombreArchivo, codEntidad, anio, periodo) {
  try {
    const entidad = ENTIDADES[codEntidad];
    if (!entidad) throw new Error("Entidad no reconocida: " + codEntidad);

    const bytes = Utilities.base64Decode(datosArchivo);
    const blob = Utilities.newBlob(bytes, MimeType.MICROSOFT_EXCEL, nombreArchivo);

    const carpeta = DriveApp.getFolderById(entidad.carpeta);
    const nombre = codEntidad + "_" + anio + "-" + periodo + ".xlsx";

    // Una carga reemplaza a la anterior del mismo periodo
    const previos = carpeta.getFilesByName(nombre);
    while (previos.hasNext()) previos.next().setTrashed(true);

    const archivo = carpeta.createFile(blob).setName(nombre);

    // Comprobación rápida de que el archivo es la plantilla
    const copia = Drive.Files.copy(
      { name: "TMP_REV", mimeType: MimeType.GOOGLE_SHEETS }, archivo.getId()
    );
    const libro = SpreadsheetApp.openById(copia.id);
    const tieneHoja = libro.getSheetByName("350") !== null;
    DriveApp.getFileById(copia.id).setTrashed(true);

    if (!tieneHoja) {
      archivo.setTrashed(true);
      return { ok: false, mensaje: "El archivo no tiene la hoja '350'. " +
                                   "Debe ser la plantilla generada por el sistema." };
    }

    return {
      ok: true,
      mensaje: "Archivo cargado: " + nombre,
      entidad: entidad.nombre
    };

  } catch (e) {
    return { ok: false, mensaje: e.message };
  }
}


/**
 * Devuelve el estado de cada entidad para un periodo.
 */
function obtenerEstadoPeriodo(anio, periodo) {
  const estado = {};

  Object.keys(ENTIDADES).forEach(function (cod) {
    const entidad = ENTIDADES[cod];
    const nombre = cod + "_" + anio + "-" + periodo + ".xlsx";
    const archivos = DriveApp.getFolderById(entidad.carpeta).getFilesByName(nombre);

    if (archivos.hasNext()) {
      const a = archivos.next();
      estado[cod] = {
        entidad: entidad.nombre,
        estado: "CARGADO",
        fecha: Utilities.formatDate(a.getLastUpdated(), "America/Bogota",
                                    "dd/MM/yyyy HH:mm")
      };
    } else {
      estado[cod] = { entidad: entidad.nombre, estado: "PENDIENTE", fecha: "" };
    }
  });

  return { periodo: anio + "-" + periodo, entidades: estado };
}


function consolidarDesdeWeb(anio, periodo) {
  try {
    return { ok: true, resultado: consolidar(anio, periodo) };
  } catch (e) {
    return { ok: false, mensaje: e.message };
  }
}


// ============================================================
// CONSOLIDACIÓN
// ============================================================

function consolidar(anio, periodo) {
  const archivos = [];

  Object.keys(ENTIDADES).forEach(function (cod) {
    const nombre = cod + "_" + anio + "-" + periodo + ".xlsx";
    const lista = DriveApp.getFolderById(ENTIDADES[cod].carpeta)
                    .getFilesByName(nombre);
    if (lista.hasNext()) archivos.push({ cod: cod, archivo: lista.next() });
  });

  if (archivos.length < 2) {
    throw new Error("Faltan cargas. Para el periodo " + anio + "-" + periodo +
                    " hay " + archivos.length + " de " +
                    Object.keys(ENTIDADES).length + " entidades.");
  }

  // --- Leer y sumar ---
  const sumas = {};
  let detalle = [];
  const temporales = [];

  archivos.forEach(function (item) {
    const copia = Drive.Files.copy(
      { name: "TMP_" + item.cod, mimeType: MimeType.GOOGLE_SHEETS },
      item.archivo.getId()
    );
    temporales.push(copia.id);

    const datos = leerArchivo(SpreadsheetApp.openById(copia.id));

    Object.keys(datos.valores).forEach(function (r) {
      sumas[r] = (sumas[r] || 0) + datos.valores[r];
    });

    detalle = detalle.concat(datos.detalle);
  });

  const detalleAgrupado = agruparPaises(detalle);

  if (detalleAgrupado.length > (DET_FIN - DET_INI + 1)) {
    throw new Error("El detalle consolidado da " + detalleAgrupado.length +
                    " filas y el formato solo admite 45.");
  }

  totalesExterior(sumas, detalleAgrupado);

  // --- Escribir ---
  const salida = escribirConsolidado(anio, periodo, sumas, detalleAgrupado);

  temporales.forEach(function (id) {
    try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {}
  });

  return {
    nombre: salida.getName(),
    periodo: anio + "-" + periodo,
    entidades: archivos.map(function (a) { return ENTIDADES[a.cod].nombre; }),
    filasPaises: detalleAgrupado.length,
    urlDescarga: "https://docs.google.com/spreadsheets/d/" + salida.getId() +
                 "/export?format=xlsx"
  };
}


/**
 * Extrae renglones y detalle de países de un archivo.
 */
function leerArchivo(libro) {
  const valores = {};
  const hoja = libro.getSheetByName("350");
  if (!hoja) throw new Error("El archivo no tiene la hoja '350'.");

  const datos = hoja.getDataRange().getValues();

  for (let f = 0; f < datos.length; f++) {
    COLS.forEach(function (col) {
      const n = datos[f][col - 1];
      if (typeof n === "number" && n >= 27 && n <= 138) {
        valores[n] = aNumero(datos[f][col]);
      }
    });
  }

  const detalle = [];
  const hoja2 = libro.getSheetByName("HOJA 2");

  if (hoja2) {
    const d2 = hoja2.getDataRange().getValues();

    for (let f = DET_INI - 1; f < DET_FIN && f < d2.length; f++) {
      const convenio = d2[f][DET.convenio - 1];
      if (!convenio) continue;

      detalle.push({
        convenio:  String(convenio).trim().toUpperCase(),
        concepto:  d2[f][DET.concepto - 1],
        persona:   String(d2[f][DET.persona - 1]).trim(),
        pais:      String(d2[f][DET.pais - 1]).trim(),
        codPais:   d2[f][DET.codPais - 1],
        tarifa:    d2[f][DET.tarifa - 1],
        base:      aNumero(d2[f][DET.base - 1]),
        retencion: aNumero(d2[f][DET.retencion - 1])
      });
    }
  }

  return { valores: valores, detalle: detalle };
}


/**
 * Agrupa las filas que comparten convenio, concepto,
 * tipo de persona, país y tarifa.
 */
function agruparPaises(detalle) {
  const grupos = {};
  const orden = [];

  detalle.forEach(function (fila) {
    const clave = [fila.convenio, fila.concepto, fila.persona,
                   fila.pais, fila.tarifa].join("|");

    if (!grupos[clave]) {
      grupos[clave] = {
        convenio: fila.convenio, concepto: fila.concepto,
        persona: fila.persona, pais: fila.pais, codPais: fila.codPais,
        tarifa: fila.tarifa, base: 0, retencion: 0
      };
      orden.push(clave);
    }

    grupos[clave].base      += fila.base;
    grupos[clave].retencion += fila.retencion;
  });

  return orden.map(function (k) { return grupos[k]; });
}


/**
 * Calcula 148 a 155 desde el detalle agrupado, para que siempre cuadren.
 */
function totalesExterior(sumas, detalle) {
  const t = {148:0, 149:0, 150:0, 151:0, 152:0, 153:0, 154:0, 155:0};

  detalle.forEach(function (fila) {
    const sin = fila.convenio.indexOf("SIN") === 0;
    const jur = String(fila.persona) === "1";

    if (sin && jur)   { t[148] += fila.base; t[150] += fila.retencion; }
    if (sin && !jur)  { t[149] += fila.base; t[151] += fila.retencion; }
    if (!sin && jur)  { t[152] += fila.base; t[154] += fila.retencion; }
    if (!sin && !jur) { t[153] += fila.base; t[155] += fila.retencion; }
  });

  Object.keys(t).forEach(function (r) { sumas[r] = t[r]; });
}


/**
 * Escribe el consolidado sobre una copia limpia de la plantilla base.
 */
function escribirConsolidado(anio, periodo, sumas, detalle) {
  const base = archivoBase();
  const copia = Drive.Files.copy(
    { name: "TMP_CONS", mimeType: MimeType.GOOGLE_SHEETS }, base.getId()
  );

  const libro = SpreadsheetApp.openById(copia.id);

  // Encabezado: periodo sí, entidad no (es la suma de varias)
  escribirEncabezado(libro, { nombre: "CONSOLIDADO DAVIVIENDA + DAVIBANK",
                              nit: "", dv: "", dirSeccional: "" }, anio, periodo);

  // --- Hoja 1 ---
  const hoja = libro.getSheetByName("350");
  const datos = hoja.getDataRange().getValues();

  for (let f = 0; f < datos.length; f++) {
    COLS.forEach(function (col) {
      const n = datos[f][col - 1];
      if (typeof n !== "number" || n < 27 || n > 138) return;
      if (CALCULADOS.indexOf(n) >= 0) return;

      hoja.getRange(f + 1, col + 1).setValue(sumas[n] || 0);
    });
  }

  // --- Hoja 2 ---
  const hoja2 = libro.getSheetByName("HOJA 2");

  if (hoja2) {
    hoja2.getRange(DET_INI, DET.convenio, DET_FIN - DET_INI + 1, 48).clearContent();

    detalle.forEach(function (fila, i) {
      const f = DET_INI + i;
      hoja2.getRange(f, DET.convenio ).setValue(fila.convenio);
      hoja2.getRange(f, DET.concepto ).setValue(fila.concepto);
      hoja2.getRange(f, DET.persona  ).setValue(fila.persona);
      hoja2.getRange(f, DET.pais     ).setValue(fila.pais);
      hoja2.getRange(f, DET.codPais  ).setValue(fila.codPais);
      hoja2.getRange(f, DET.base     ).setValue(fila.base);
      hoja2.getRange(f, DET.tarifa   ).setValue(fila.tarifa);
      hoja2.getRange(f, DET.retencion).setValue(fila.retencion);
    });

    const d2 = hoja2.getDataRange().getValues();

    for (let f = DET_FIN; f < d2.length; f++) {
      [32, 47].forEach(function (col) {
        const n = d2[f][col - 1];
        if (typeof n === "number" && n >= 148 && n <= 155) {
          hoja2.getRange(f + 1, col + 1).setValue(sumas[n] || 0);
        }
      });
    }
  }

  SpreadsheetApp.flush();

  const sello = Utilities.formatDate(new Date(), "America/Bogota", "yyyyMMdd-HHmm");
  const archivo = DriveApp.getFileById(copia.id);
  archivo.setName("CONSOLIDADO_350_" + anio + "-" + periodo + "_" + sello);

  DriveApp.getFolderById(CARPETA_CONSOLIDADOS).addFile(archivo);
  DriveApp.getRootFolder().removeFile(archivo);
  archivo.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);

  return archivo;
}


// ============================================================
// AUXILIARES
// ============================================================

/**
 * Devuelve el archivo maestro de la carpeta de plantilla base.
 */
function archivoBase() {
  const lista = DriveApp.getFolderById(CARPETA_BASE).getFiles();
  const encontrados = [];

  while (lista.hasNext()) encontrados.push(lista.next());

  if (encontrados.length === 0) {
    throw new Error("No hay ningún archivo en la carpeta de plantilla base.");
  }
  if (encontrados.length > 1) {
    throw new Error("Hay " + encontrados.length + " archivos en la carpeta de " +
                    "plantilla base. Debe quedar solo el formato maestro.");
  }

  return encontrados[0];
}


/**
 * Escribe los datos de la entidad y el periodo en el encabezado.
 */
function escribirEncabezado(libro, entidad, anio, periodo) {
  const hoja = libro.getSheetByName("350");
  if (!hoja) return;

  // El año y el periodo van dígito por dígito en sus casillas
  const digitosAnio = String(anio).split("");
  ENC.anio.forEach(function (celda, i) {
    hoja.getRange(celda).setValue(digitosAnio[i] || "");
  });

  const digitosPer = ("0" + parseInt(periodo, 10)).slice(-2).split("");
  ENC.periodo.forEach(function (celda, i) {
    hoja.getRange(celda).setValue(digitosPer[i] || "");
  });

  hoja.getRange(ENC.nit).setValue(entidad.nit || "");
  hoja.getRange(ENC.dv).setValue(entidad.dv || "");
  hoja.getRange(ENC.razonSocial).setValue(entidad.nombre || "");
  hoja.getRange(ENC.dirSeccional).setValue(entidad.dirSeccional || "");
}


function aNumero(valor) {
  if (typeof valor === "number") return valor;
  if (!valor) return 0;

  const n = Number(String(valor).replace(/[.,\s$]/g, ""));
  return isNaN(n) ? 0 : n;
}