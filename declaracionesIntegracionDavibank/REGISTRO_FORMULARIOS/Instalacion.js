const HOJAS_REGISTRO = {
  FORMULARIOS: [
    "cod_formulario", "nombre", "version", "periodicidad", "num_paginas",
    "disposicion", "renglones_esperados", "estado", "fecha_alta"
  ],
  RENGLONES: [
    "cod_formulario", "version", "pagina", "nro_renglon", "etiqueta",
    "tipo_persona", "tipo_valor", "grupo_concepto", "seccion", "editable"
  ],
  ENTIDADES: [
    "cod_entidad", "nombre", "nit", "dv",
    "cod_direccion_seccional", "actividad_economica", "activa"
  ]
};

const HOJAS_OPERACION = {
  PLANTILLAS_EMITIDAS: [
    "id_plantilla", "cod_formulario", "version", "cod_entidad",
    "anio", "periodo", "id_archivo_drive", "generada_por",
    "fecha_generacion", "estado"
  ],
  ENTREGAS: [
    "radicado", "id_plantilla", "cod_formulario", "cod_entidad",
    "anio", "periodo", "nombre_archivo", "id_archivo_drive",
    "cargada_por", "fecha_carga", "estado", "num_errores"
  ],
  DATOS_CARGADOS: [
    "radicado", "cod_formulario", "version",
    "nro_renglon", "valor", "fecha_registro"
  ],
  DATOS_DETALLE: [
    "radicado", "cod_formulario", "tabla", "nro_fila",
    "convenio", "concepto_pago", "tipo_persona", "pais", "cod_pais",
    "base", "tarifa", "retencion", "fecha_registro"
  ],
  LOG_VALIDACION: [
    "radicado", "num_error", "tipo", "nro_renglon",
    "valor_esperado", "valor_encontrado", "mensaje", "fecha"
  ],
  CONSOLIDADOS: [
    "id_consolidado", "cod_formulario", "anio", "periodo",
    "radicados_origen", "id_archivo_drive", "generado_por",
    "vigente", "fecha_generacion"
  ]
};

const ENTIDADES_INICIALES = [
  ["DAVIVIENDA", "BANCO DAVIVIENDA S.A.", "860034313", "7", "31", "6412", "SI"],
  ["DAVIBANK",   "DAVIBANK",              "",          "",  "",   "",     "SI"]
];


function instalarHojas() {
  crearHojas_(SpreadsheetApp.openById(ID_REGISTRO),  HOJAS_REGISTRO);
  crearHojas_(SpreadsheetApp.openById(ID_OPERACION), HOJAS_OPERACION);
  sembrarEntidades_();

  Logger.log("Hojas creadas correctamente.");
}


function crearHojas_(libro, definicion) {
  Object.keys(definicion).forEach(function (nombreHoja) {
    let hoja = libro.getSheetByName(nombreHoja);

    if (!hoja) {
      hoja = libro.insertSheet(nombreHoja);
    }

    if (hoja.getLastRow() === 0) {
      const encabezados = definicion[nombreHoja];
      hoja.getRange(1, 1, 1, encabezados.length)
          .setValues([encabezados])
          .setFontWeight("bold")
          .setBackground(COLOR_ENCABEZADO);
      hoja.setFrozenRows(1);
      hoja.autoResizeColumns(1, encabezados.length);
    }
  });

  ["Hoja 1", "Hoja1", "Sheet1"].forEach(function (nombre) {
    const hoja = libro.getSheetByName(nombre);
    if (hoja && libro.getSheets().length > 1 && hoja.getLastRow() === 0) {
      libro.deleteSheet(hoja);
    }
  });
}


function sembrarEntidades_() {
  const hoja = SpreadsheetApp.openById(ID_REGISTRO).getSheetByName("ENTIDADES");
  if (hoja.getLastRow() > 1) return;

  ENTIDADES_INICIALES.forEach(function (entidad) {
    hoja.appendRow(entidad);
  });
}


/**
 * Registra un formulario en el catálogo.
 *
 * disposicion: "MATRIZ" (como el 350) o "LISTA" (como el 300)
 * renglonesEsperados: rangos separados por coma, p. ej. "1-6,27-93,100"
 */
function registrarFormulario(codigo, nombre, version, periodicidad,
                             paginas, disposicion, renglonesEsperados) {
  const hoja = SpreadsheetApp.openById(ID_REGISTRO).getSheetByName("FORMULARIOS");
  const datos = hoja.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === codigo && datos[i][2] === version) {
      Logger.log("Ya estaba registrado: " + codigo + " " + version);
      return;
    }
  }

  hoja.appendRow([
    codigo, nombre, version, periodicidad, paginas,
    String(disposicion).toUpperCase(), renglonesEsperados,
    "ACTIVO", new Date()
  ]);

  Logger.log("Formulario registrado: " + codigo + " " + version);
}


function registrarFormulariosIniciales() {
  registrarFormulario("F350", "Declaración de retenciones en la fuente",
                      "v2026", "MENSUAL", 2, "MATRIZ", "29-138,148-155");

  registrarFormulario("F300", "Declaración del impuesto sobre las ventas - IVA",
                      "v2026", "BIMESTRAL", 1, "LISTA", "1-6,27-93,100");
}