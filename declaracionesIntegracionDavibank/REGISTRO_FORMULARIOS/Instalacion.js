const COLOR_ENCABEZADO = "#D9E2F3";

const HOJAS_REGISTRO = {
  FORMULARIOS: [
    "cod_formulario", "nombre", "version", "periodicidad",
    "num_paginas", "estado", "fecha_alta"
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
  LOG_VALIDACION: [
    "radicado", "num_error", "tipo", "nro_renglon",
    "valor_esperado", "valor_encontrado", "mensaje", "fecha"
  ]
};

const FORMULARIO_F350 = [
  "F350", "Declaración de retenciones en la fuente", "v2026",
  "MENSUAL", 2, "ACTIVO", new Date()
];

const ENTIDADES_INICIALES = [
  ["DAVIVIENDA", "BANCO DAVIVIENDA S.A.", "860034313", "7", "31", "6412", "SI"],
  ["DAVIBANK",   "DAVIBANK",              "",          "",  "",   "",     "SI"]
];

function instalarHojas() {
  const libroRegistro  = SpreadsheetApp.openById(ID_REGISTRO);
  const libroOperacion = SpreadsheetApp.openById(ID_OPERACION);

  crearHojas_(libroRegistro,  HOJAS_REGISTRO);
  crearHojas_(libroOperacion, HOJAS_OPERACION);

  sembrarDatosIniciales_(libroRegistro);

  Logger.log("Instalación completada.");
  Logger.log("REGISTRO_FORMULARIOS: " + Object.keys(HOJAS_REGISTRO).join(", "));
  Logger.log("OPERACION_DECLARACIONES: " + Object.keys(HOJAS_OPERACION).join(", "));
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

  eliminarHojaPorDefecto_(libro);
}
function eliminarHojaPorDefecto_(libro) {
  ["Hoja 1", "Hoja1", "Sheet1"].forEach(function (nombre) {
    const hoja = libro.getSheetByName(nombre);
    if (hoja && libro.getSheets().length > 1 && hoja.getLastRow() === 0) {
      libro.deleteSheet(hoja);
    }
  });
}
function sembrarDatosIniciales_(libroRegistro) {
  const hojaForm = libroRegistro.getSheetByName("FORMULARIOS");
  if (hojaForm.getLastRow() <= 1) {
    hojaForm.appendRow(FORMULARIO_F350);
  }

  const hojaEnt = libroRegistro.getSheetByName("ENTIDADES");
  if (hojaEnt.getLastRow() <= 1) {
    ENTIDADES_INICIALES.forEach(function (entidad) {
      hojaEnt.appendRow(entidad);
    });
  }
}