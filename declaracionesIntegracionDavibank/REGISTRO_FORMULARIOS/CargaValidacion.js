/**
 * Procesa un archivo cargado.
 *
 * @param {string} idArchivo  Identificador del archivo en Drive
 * @return {Object} { estado, radicado, errores, resumen }
 */
function procesarCarga(idArchivo) {
  const archivo = DriveApp.getFileById(idArchivo);
  let idTemporal = null;

  try {
    const conversion = abrirComoSheet_(archivo);
    const libroCarga = conversion.libro;
    idTemporal = conversion.idTemporal;

    const meta     = leerMetadatosCarga_(libroCarga);
    const catalogo = leerCatalogoParaValidar_(meta.cod_formulario, meta.version);
    const valores  = leerValoresDiligenciados_(libroCarga);
    const errores  = validarCarga_(catalogo, valores);

    const aceptada = errores.length === 0;
    const radicado = generarRadicado_();

    const idArchivado = archivarCarga_(archivo, meta.cod_entidad, aceptada);

    registrarEntrega_(radicado, meta, archivo.getName(), idArchivado,
                      aceptada, errores.length);

    if (aceptada) {
      guardarDatos_(radicado, meta, valores);
    } else {
      registrarErrores_(radicado, errores);
    }

    return {
      estado: aceptada ? "APROBADO" : "RECHAZADO",
      radicado: radicado,
      errores: errores,
      resumen: aceptada
        ? "Se cargaron " + Object.keys(valores).length + " renglones."
        : "No se guardó ningún dato. Se encontraron " + errores.length + " inconsistencias."
    };

  } finally {
    if (idTemporal) {
      try { DriveApp.getFileById(idTemporal).setTrashed(true); } catch (e) {}
    }
  }
}

function abrirComoSheet_(archivo) {
  const tipo = archivo.getMimeType();

  if (tipo === MimeType.GOOGLE_SHEETS) {
    return { libro: SpreadsheetApp.openById(archivo.getId()), idTemporal: null };
  }

  const esExcel = tipo === MimeType.MICROSOFT_EXCEL ||
                  tipo === MimeType.MICROSOFT_EXCEL_LEGACY;

  if (!esExcel) {
    throw new Error("Formato no admitido. Se espera un archivo .xlsx, .xls o Google Sheets.");
  }

  const copia = Drive.Files.copy(
    { name: "TEMP_CARGA_" + archivo.getName(), mimeType: MimeType.GOOGLE_SHEETS },
    archivo.getId()
  );

  return { libro: SpreadsheetApp.openById(copia.id), idTemporal: copia.id };
}

function leerMetadatosCarga_(libro) {
  let idPlantilla = null;

  const hojaMeta = libro.getSheetByName("_META");
  if (hojaMeta && hojaMeta.getLastRow() > 0) {
    const datos = hojaMeta.getDataRange().getValues();
    for (let i = 0; i < datos.length; i++) {
      if (String(datos[i][0]).trim() === "id_plantilla") {
        idPlantilla = String(datos[i][1]).trim();
      }
    }
  }

  if (!idPlantilla) {
    const hoja = libro.getSheetByName("FORMULARIO");
    if (hoja) {
      const datos = hoja.getDataRange().getValues();
      for (let i = 0; i < Math.min(datos.length, 20); i++) {
        for (let j = 0; j < datos[i].length; j++) {
          if (String(datos[i][j]).indexOf("PLT-") === 0) {
            idPlantilla = String(datos[i][j]).trim();
          }
        }
      }
    }
  }

  if (!idPlantilla) {
    throw new Error(
      "El archivo no tiene identificador de plantilla. " +
      "Debe diligenciarse sobre la plantilla generada por el sistema."
    );
  }

  return buscarPlantillaEmitida_(idPlantilla);
}

function buscarPlantillaEmitida_(idPlantilla) {
  const datos = SpreadsheetApp.openById(ID_OPERACION)
                  .getSheetByName("PLANTILLAS_EMITIDAS")
                  .getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][0]).trim() === idPlantilla) {
      return {
        id_plantilla:   datos[i][0],
        cod_formulario: datos[i][1],
        version:        datos[i][2],
        cod_entidad:    datos[i][3],
        anio:           datos[i][4],
        periodo:        datos[i][5]
      };
    }
  }

  throw new Error("El identificador " + idPlantilla +
                  " no corresponde a ninguna plantilla emitida.");
}

function leerCatalogoParaValidar_(codFormulario, version) {
  const datos = SpreadsheetApp.openById(ID_REGISTRO)
                  .getSheetByName("RENGLONES").getDataRange().getValues();
  const catalogo = {};

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] !== codFormulario || datos[i][1] !== version) continue;

    catalogo[datos[i][3]] = {
      nro_renglon:  datos[i][3],
      etiqueta:     datos[i][4],
      tipo_persona: datos[i][5],
      tipo_valor:   datos[i][6],
      seccion:      datos[i][8]
    };
  }

  return catalogo;
}

function leerValoresDiligenciados_(libro) {
  const hoja = libro.getSheetByName("FORMULARIO");

  if (!hoja) {
    throw new Error("El archivo no contiene la hoja FORMULARIO.");
  }

  const datos = hoja.getDataRange().getValues();
  const valores = {};

  const COLUMNAS_RENGLON = [1, 3, 5, 7];

  for (let i = 0; i < datos.length; i++) {
    COLUMNAS_RENGLON.forEach(function (c) {
      const nro = parseInt(datos[i][c], 10);
      if (!isNaN(nro) && nro >= 29 && nro <= 155) {
        valores[nro] = datos[i][c + 1];
      }
    });
  }

  for (let i = 0; i < datos.length; i++) {
    const nro = parseInt(datos[i][7], 10);
    if (!isNaN(nro) && nro >= 129 && nro <= 138) {
      valores[nro] = datos[i][8];
    }
  }

  const hojaExt = libro.getSheetByName("EXTERIOR");
  if (hojaExt) {
    const datosExt = hojaExt.getDataRange().getValues();
    for (let i = 0; i < datosExt.length; i++) {
      const nro = parseInt(datosExt[i][6], 10);
      if (!isNaN(nro) && nro >= 148 && nro <= 155) {
        valores[nro] = datosExt[i][7];
      }
    }
  }

  return valores;
}

function validarCarga_(catalogo, valores) {
  const errores = [];

  Object.keys(catalogo).forEach(function (nro) {
    if (!(nro in valores)) {
      errores.push({
        tipo: "ESTRUCTURA", renglon: nro, esperado: "presente", encontrado: "ausente",
        mensaje: "Falta el renglón " + nro + " · " + catalogo[nro].etiqueta
      });
    }
  });

  Object.keys(valores).forEach(function (nro) {
    if (!(nro in catalogo)) {
      errores.push({
        tipo: "ESTRUCTURA", renglon: nro, esperado: "", encontrado: nro,
        mensaje: "El renglón " + nro + " no pertenece al formulario"
      });
    }
  });

  if (errores.length) return errores;

  Object.keys(catalogo).forEach(function (nro) {
    const bruto = valores[nro];

    if (bruto === "" || bruto === null || bruto === undefined) {
      errores.push({
        tipo: "DATO", renglon: nro, esperado: "valor numérico", encontrado: "vacío",
        mensaje: "Renglón " + nro + " · " + catalogo[nro].etiqueta +
                 ": vacío. Debe ir en cero si no aplica"
      });
      return;
    }

    if (typeof bruto !== "number" &&
        isNaN(Number(String(bruto).replace(/[.,\s$]/g, "")))) {
      errores.push({
        tipo: "DATO", renglon: nro, esperado: "valor numérico", encontrado: String(bruto),
        mensaje: "Renglón " + nro + ": el valor no es numérico"
      });
    }
  });

  if (errores.length) return errores;

  return validarTotales_(catalogo, valores);
}


function validarTotales_(catalogo, valores) {
  const errores = [];
  const num = function (n) { return aNumero_(valores[n]); };

  let sumaRetenciones = 0;

  Object.keys(catalogo).forEach(function (nro) {
    const n = parseInt(nro, 10);
    if (catalogo[nro].tipo_valor === "RETENCION" && n <= 128) {
      sumaRetenciones += num(n);
    }
  });
  comparar_(errores, 130, sumaRetenciones - num(129), num(130),
            "Total retenciones renta");
  comparar_(errores, 134, num(131) + num(132) - num(133), num(134),
            "Total retenciones IVA");
  comparar_(errores, 136, num(130) + num(134) + num(135), num(136),
            "Total retenciones");
  comparar_(errores, 138, num(136) + num(137), num(138),
            "Total retenciones más sanciones");
  return errores;
}


function comparar_(errores, renglon, esperado, encontrado, etiqueta) {
  if (Math.abs(esperado - encontrado) < 1) return;

  errores.push({
    tipo: "TOTAL", renglon: renglon, esperado: esperado, encontrado: encontrado,
    mensaje: "Renglón " + renglon + " · " + etiqueta + ": debería ser " +
             formatearMiles_(esperado) + " y el archivo trae " + formatearMiles_(encontrado)
  });
}

function archivarCarga_(archivo, codEntidad, aceptada) {
  const carpetas = carpetasDeEntidad(codEntidad);
  const destino = DriveApp.getFolderById(
    aceptada ? carpetas.aceptadas : carpetas.rechazadas
  );

  return archivo.makeCopy(archivo.getName(), destino).getId();
}

function registrarEntrega_(radicado, meta, nombreArchivo, idArchivado, aceptada, numErrores) {
  SpreadsheetApp.openById(ID_OPERACION)
    .getSheetByName("ENTREGAS")
    .appendRow([
      radicado, meta.id_plantilla, meta.cod_formulario, meta.cod_entidad,
      meta.anio, meta.periodo, nombreArchivo, idArchivado,
      Session.getActiveUser().getEmail(), new Date(),
      aceptada ? "APROBADO" : "RECHAZADO", numErrores
    ]);
}

function guardarDatos_(radicado, meta, valores) {
  const hoja = SpreadsheetApp.openById(ID_OPERACION).getSheetByName("DATOS_CARGADOS");
  const ahora = new Date();
  const filas = [];

  Object.keys(valores).forEach(function (nro) {
    filas.push([radicado, meta.cod_formulario, meta.version,
                parseInt(nro, 10), aNumero_(valores[nro]), ahora]);
  });

  if (filas.length) {
    hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, filas[0].length)
        .setValues(filas);
  }
}

function registrarErrores_(radicado, errores) {
  const hoja = SpreadsheetApp.openById(ID_OPERACION).getSheetByName("LOG_VALIDACION");
  const ahora = new Date();

  const filas = errores.map(function (e, i) {
    return [radicado, i + 1, e.tipo, e.renglon,
            e.esperado, e.encontrado, e.mensaje, ahora];
  });

  hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, filas[0].length)
      .setValues(filas);
}
function generarRadicado_() {
  const hoja = SpreadsheetApp.openById(ID_OPERACION).getSheetByName("ENTREGAS");
  return "ENT-" + new Date().getFullYear() + "-" +
         ("0000" + hoja.getLastRow()).slice(-4);
}

function aNumero_(valor) {
  if (typeof valor === "number") return valor;
  if (valor === "" || valor === null || valor === undefined) return 0;

  const n = Number(String(valor).replace(/[.,\s$]/g, ""));
  return isNaN(n) ? 0 : n;
}

function formatearMiles_(n) {
  return Number(n).toLocaleString("es-CO");
}