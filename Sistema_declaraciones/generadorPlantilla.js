/**
 * ============================================================
 * GENERADOR DE PLANTILLAS
 *
 * Crea una plantilla en Google Sheets a partir del catálogo
 * de renglones, la guarda en 3_PLANTILLAS_EMITIDAS y la
 * registra en PLANTILLAS_EMITIDAS.
 *
 * Ejecutar: generarPlantillaPrueba()
 * ============================================================
 */

const COLOR_SECCION   = "#1F4E79";
const COLOR_BLOQUEADO = "#F2F2F2";
const COLOR_EDITABLE  = "#FFFFFF";


/**
 * Función de prueba. Ajustar los parámetros según se necesite.
 */
function generarPlantillaPrueba() {
  const resultado = generarPlantilla("F350", "v2026", "DAVIVIENDA", 2026, 7);
  Logger.log("Plantilla generada: " + resultado.idPlantilla);
  Logger.log("URL: " + resultado.url);
}


/**
 * Genera una plantilla para un formulario, entidad y periodo dados.
 */
function generarPlantilla(codFormulario, version, codEntidad, anio, periodo) {
  const libroRegistro = SpreadsheetApp.openById(ID_REGISTRO);

  const renglones = leerRenglones_(libroRegistro, codFormulario, version);
  if (!renglones.length) {
    throw new Error("No hay renglones en el catálogo para " + codFormulario + " " + version);
  }

  const entidad = leerEntidad_(libroRegistro, codEntidad);
  const idPlantilla = generarIdPlantilla_();

  // --- Creación del libro de la plantilla ---
  const nombreArchivo = idPlantilla + "_" + codFormulario + "_" + codEntidad +
                        "_" + anio + "-" + formatearPeriodo_(periodo);
  const libroPlantilla = SpreadsheetApp.create(nombreArchivo);

  construirHojaMeta_(libroPlantilla, idPlantilla, codFormulario, version,
                     codEntidad, anio, periodo);
  construirHojaEncabezado_(libroPlantilla, entidad, anio, periodo);
  construirHojaRenglones_(libroPlantilla, renglones);

  // Se elimina la hoja por defecto creada con el libro
  const porDefecto = libroPlantilla.getSheetByName("Hoja 1") ||
                     libroPlantilla.getSheetByName("Sheet1");
  if (porDefecto) libroPlantilla.deleteSheet(porDefecto);

  libroPlantilla.setActiveSheet(libroPlantilla.getSheetByName("RENGLONES"));

  // --- Archivado en Drive ---
  const idArchivo = moverAPlantillasEmitidas_(
    libroPlantilla.getId(), codFormulario, anio, periodo
  );

  registrarPlantillaEmitida_(idPlantilla, codFormulario, version, codEntidad,
                             anio, periodo, idArchivo);

  return {
    idPlantilla: idPlantilla,
    idArchivo: idArchivo,
    url: libroPlantilla.getUrl()
  };
}


/**
 * Lee del catálogo los renglones de un formulario y versión.
 */
function leerRenglones_(libro, codFormulario, version) {
  const hoja = libro.getSheetByName("RENGLONES");
  const datos = hoja.getDataRange().getValues();
  const resultado = [];

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] !== codFormulario || datos[i][1] !== version) continue;

    resultado.push({
      pagina:         datos[i][2],
      nro_renglon:    datos[i][3],
      etiqueta:       datos[i][4],
      tipo_persona:   datos[i][5],
      tipo_valor:     datos[i][6],
      grupo_concepto: datos[i][7],
      seccion:        datos[i][8],
      editable:       datos[i][9]
    });
  }

  resultado.sort(function (a, b) { return a.nro_renglon - b.nro_renglon; });
  return resultado;
}


/**
 * Lee los datos maestros de una entidad.
 */
function leerEntidad_(libro, codEntidad) {
  const datos = libro.getSheetByName("ENTIDADES").getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === codEntidad) {
      return {
        codigo: datos[i][0], nombre: datos[i][1], nit: datos[i][2],
        dv: datos[i][3], direccion_seccional: datos[i][4],
        actividad: datos[i][5]
      };
    }
  }

  throw new Error("Entidad no encontrada en el catálogo: " + codEntidad);
}


/**
 * Construye la hoja oculta con los metadatos de la plantilla.
 */
function construirHojaMeta_(libro, idPlantilla, codFormulario, version,
                            codEntidad, anio, periodo) {
  const hoja = libro.insertSheet("_META");

  hoja.getRange(1, 1, 8, 2).setValues([
    ["id_plantilla",     idPlantilla],
    ["cod_formulario",   codFormulario],
    ["version",          version],
    ["cod_entidad",      codEntidad],
    ["anio",             anio],
    ["periodo",          periodo],
    ["generada_por",     Session.getActiveUser().getEmail()],
    ["fecha_generacion", new Date()]
  ]);

  hoja.getRange(1, 1, 8, 1).setFontWeight("bold");
  hoja.autoResizeColumns(1, 2);
  hoja.hideSheet();
}


/**
 * Construye la hoja de encabezado con los datos del declarante.
 */
function construirHojaEncabezado_(libro, entidad, anio, periodo) {
  const hoja = libro.insertSheet("ENCABEZADO");

  hoja.getRange(1, 1, 1, 2)
      .setValues([["DATOS DEL DECLARANTE", ""]])
      .setFontWeight("bold")
      .setFontSize(12)
      .setBackground(COLOR_SECCION)
      .setFontColor("#FFFFFF");

  const campos = [
    ["Año",                        anio],
    ["Periodo",                    periodo],
    ["NIT",                        entidad.nit],
    ["DV",                         entidad.dv],
    ["Razón social",               entidad.nombre],
    ["Cód. dirección seccional",   entidad.direccion_seccional],
    ["Actividad económica",        entidad.actividad],
    ["Nombre del responsable",     ""],
    ["Correo del responsable",     ""],
    ["Fecha de diligenciamiento",  ""]
  ];

  hoja.getRange(3, 1, campos.length, 2).setValues(campos);
  hoja.getRange(3, 1, campos.length, 1).setFontWeight("bold");

  // Los datos maestros llegan precargados y no deben editarse
  hoja.getRange(3, 2, 7, 1).setBackground(COLOR_BLOQUEADO);
  hoja.getRange(10, 2, 3, 1).setBackground(COLOR_EDITABLE);

  hoja.setColumnWidth(1, 220);
  hoja.setColumnWidth(2, 280);
}


/**
 * Construye la hoja de renglones agrupada por sección.
 */
function construirHojaRenglones_(libro, renglones) {
  const hoja = libro.insertSheet("RENGLONES");

  const encabezados = ["Renglón", "Concepto", "Tipo persona", "Tipo valor", "Valor"];
  hoja.getRange(1, 1, 1, encabezados.length)
      .setValues([encabezados])
      .setFontWeight("bold")
      .setBackground(COLOR_SECCION)
      .setFontColor("#FFFFFF");

  const filas = [];
  const filasSeccion = [];   // posiciones de los separadores
  const filasBloqueadas = [];
  let seccionActual = "";
  let fila = 2;

  renglones.forEach(function (r) {
    if (r.seccion !== seccionActual) {
      seccionActual = r.seccion;
      filas.push([formatearSeccion_(seccionActual), "", "", "", ""]);
      filasSeccion.push(fila);
      fila++;
    }

    filas.push([
      r.nro_renglon, r.etiqueta, r.tipo_persona, r.tipo_valor, ""
    ]);

    if (String(r.editable).toUpperCase() === "NO") {
      filasBloqueadas.push(fila);
    }
    fila++;
  });

  hoja.getRange(2, 1, filas.length, 5).setValues(filas);

  // --- Formato de las filas de sección ---
  filasSeccion.forEach(function (f) {
    hoja.getRange(f, 1, 1, 5)
        .merge()
        .setFontWeight("bold")
        .setBackground("#D9E2F3");
  });

  // --- Celdas de valor: solo esta columna se diligencia ---
  const ultimaFila = hoja.getLastRow();
  hoja.getRange(2, 5, ultimaFila - 1, 1)
      .setBackground(COLOR_EDITABLE)
      .setNumberFormat("#,##0")
      .setHorizontalAlignment("right");

  // --- Renglones calculados: no se diligencian ---
  filasBloqueadas.forEach(function (f) {
    hoja.getRange(f, 5).setBackground(COLOR_BLOQUEADO);
  });

  // --- Protección de todo salvo la columna de valor ---
  const proteccion = hoja.protect().setDescription("Estructura de la plantilla");
  proteccion.setUnprotectedRanges([hoja.getRange(2, 5, ultimaFila - 1, 1)]);
  proteccion.setWarningOnly(true);

  hoja.setColumnWidth(1, 70);
  hoja.setColumnWidth(2, 420);
  hoja.setColumnWidth(3, 110);
  hoja.setColumnWidth(4, 110);
  hoja.setColumnWidth(5, 160);
  hoja.setFrozenRows(1);
}


/**
 * Mueve la plantilla a la carpeta correspondiente en Drive.
 */
function moverAPlantillasEmitidas_(idArchivo, codFormulario, anio, periodo) {
  const idCarpetaRaiz = leerConfig_("CARPETA_3_PLANTILLAS_EMITIDAS");
  const carpetaRaiz = DriveApp.getFolderById(idCarpetaRaiz);

  const carpetaForm    = obtenerOCrearCarpeta_(carpetaRaiz, codFormulario);
  const carpetaAnio    = obtenerOCrearCarpeta_(carpetaForm, String(anio));
  const carpetaPeriodo = obtenerOCrearCarpeta_(carpetaAnio, formatearPeriodo_(periodo));

  const archivo = DriveApp.getFileById(idArchivo);
  carpetaPeriodo.addFile(archivo);
  DriveApp.getRootFolder().removeFile(archivo);

  return archivo.getId();
}


/**
 * Registra la plantilla generada en el libro transaccional.
 */
function registrarPlantillaEmitida_(idPlantilla, codFormulario, version,
                                    codEntidad, anio, periodo, idArchivo) {
  SpreadsheetApp.openById(ID_OPERACION)
    .getSheetByName("PLANTILLAS_EMITIDAS")
    .appendRow([
      idPlantilla, codFormulario, version, codEntidad, anio,
      formatearPeriodo_(periodo), idArchivo,
      Session.getActiveUser().getEmail(), new Date(), "EMITIDA"
    ]);
}


// ---------- Funciones auxiliares ----------

/**
 * Genera un identificador consecutivo de plantilla.
 */
function generarIdPlantilla_() {
  const hoja = SpreadsheetApp.openById(ID_OPERACION).getSheetByName("PLANTILLAS_EMITIDAS");
  const consecutivo = hoja.getLastRow();  // encabezado incluido: el siguiente número
  return "PLT-" + new Date().getFullYear() + "-" +
         ("0000" + consecutivo).slice(-4);
}

/**
 * Devuelve el periodo con dos dígitos.
 */
function formatearPeriodo_(periodo) {
  return ("0" + periodo).slice(-2);
}

/**
 * Convierte el código de sección en un título legible.
 */
function formatearSeccion_(seccion) {
  const titulos = {
    CONCEPTOS:         "CONCEPTOS SUJETOS A RETENCIÓN",
    EXTERIOR:          "PAGOS AL EXTERIOR",
    AUTORRETENCIONES:  "AUTORRETENCIONES",
    TOTALES_RENTA:     "TOTALES RENTA Y COMPLEMENTARIO",
    TOTALES_IVA:       "RETENCIONES A TÍTULO DE IVA",
    TOTALES_TIMBRE:    "IMPUESTO DE TIMBRE NACIONAL",
    TOTALES_GENERAL:   "TOTALES GENERALES",
    TOTALES_EXTERIOR:  "TOTALES PAGOS AL EXTERIOR"
  };
  return titulos[seccion] || seccion;
}

/**
 * Lee un valor de la hoja CONFIG por su clave.
 */
function leerConfig_(clave) {
  const datos = SpreadsheetApp.openById(ID_REGISTRO)
                  .getSheetByName("CONFIG").getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === clave) return datos[i][1];
  }
  throw new Error("Clave no encontrada en CONFIG: " + clave);
}

/**
 * Devuelve una carpeta existente o la crea.
 */
function obtenerOCrearCarpeta_(padre, nombre) {
  const existentes = padre.getFoldersByName(nombre);
  return existentes.hasNext() ? existentes.next() : padre.createFolder(nombre);
}