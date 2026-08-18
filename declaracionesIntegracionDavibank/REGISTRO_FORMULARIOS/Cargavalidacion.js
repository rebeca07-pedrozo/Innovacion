/**
 * ============================================================
 * CARGA Y VALIDACIÓN
 *
 * Recibe el archivo diligenciado, lo identifica por su
 * id_plantilla y lo acepta o rechaza en su totalidad.
 *
 * Un rechazo no guarda ningún dato.
 * Requiere el servicio avanzado Drive API v3.
 * ============================================================
 */

function procesarCarga(idArchivo) {
  const archivo = DriveApp.getFileById(idArchivo);
  let idTemporal = null;

  try {
    const conversion = abrirComoSheet_(archivo);
    const libroCarga = conversion.libro;
    idTemporal = conversion.idTemporal;

    const meta     = leerMetadatosCarga_(libroCarga);
    const def      = definicionFormulario(meta.cod_formulario);
    const catalogo = leerCatalogoValidacion_(meta.cod_formulario, meta.version);
    const valores  = leerValores_(libroCarga, def.disposicion);
    const errores  = validarCarga_(catalogo, valores, meta.cod_formulario);

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


/**
 * Abre el archivo como Spreadsheet, convirtiéndolo si viene de Excel.
 */
function abrirComoSheet_(archivo) {
  const tipo = archivo.getMimeType();

  if (tipo === MimeType.GOOGLE_SHEETS) {
    return { libro: SpreadsheetApp.openById(archivo.getId()), idTemporal: null };
  }

  const esExcel = tipo === MimeType.MICROSOFT_EXCEL ||
                  tipo === MimeType.MICROSOFT_EXCEL_LEGACY;

  if (!esExcel) {
    throw new Error("Formato no admitido. Se espera .xlsx, .xls o Google Sheets.");
  }

  const copia = Drive.Files.copy(
    { name: "TEMP_CARGA_" + archivo.getName(), mimeType: MimeType.GOOGLE_SHEETS },
    archivo.getId()
  );

  return { libro: SpreadsheetApp.openById(copia.id), idTemporal: copia.id };
}


/**
 * Recupera el identificador de plantilla y sus datos asociados.
 */
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

  // Respaldo: el identificador visible en el encabezado del formulario
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
    throw new Error("El archivo no tiene identificador de plantilla. " +
                    "Debe diligenciarse sobre la plantilla generada por el sistema.");
  }

  const datos = SpreadsheetApp.openById(ID_OPERACION)
                  .getSheetByName("PLANTILLAS_EMITIDAS").getDataRange().getValues();

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


function leerCatalogoValidacion_(codFormulario, version) {
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


/**
 * Extrae los valores diligenciados según la disposición del formulario.
 */
function leerValores_(libro, disposicion) {
  const hoja = libro.getSheetByName("FORMULARIO");
  if (!hoja) throw new Error("El archivo no contiene la hoja FORMULARIO.");

  const datos = hoja.getDataRange().getValues();
  const valores = {};

  if (disposicion === "MATRIZ") {
    // Los números de renglón ocupan las columnas 2, 4, 6 y 8
    const COLUMNAS = [1, 3, 5, 7];

    for (let i = 0; i < datos.length; i++) {
      COLUMNAS.forEach(function (c) {
        const nro = parseInt(datos[i][c], 10);
        if (!isNaN(nro) && nro >= 1 && nro <= 200) {
          valores[nro] = datos[i][c + 1];
        }
      });
    }

    // Los totales de la página 1 usan las columnas 8 y 9
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

  } else {
    // Disposición de lista: renglón en columna 1, valor en columna 3
    for (let i = 0; i < datos.length; i++) {
      const nro = parseInt(datos[i][0], 10);
      if (!isNaN(nro) && nro >= 1 && nro <= 200) {
        valores[nro] = datos[i][2];
      }
    }
  }

  return valores;
}


/**
 * Aplica las validaciones en tres niveles.
 */
function validarCarga_(catalogo, valores, codFormulario) {
  const errores = [];

  // --- Estructura ---
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

  // --- Tipo de dato ---
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

  // --- Cuadre de totales ---
  if (codFormulario === "F350") return validarTotales350_(catalogo, valores);
  if (codFormulario === "F300") return validarTotales300_(valores);

  return errores;
}


/**
 * Aritmética interna del formulario 350.
 */
function validarTotales350_(catalogo, valores) {
  const errores = [];
  const num = function (n) { return aNumero(valores[n]); };

  let sumaRetenciones = 0;
  Object.keys(catalogo).forEach(function (nro) {
    const n = parseInt(nro, 10);
    if (catalogo[nro].tipo_valor === "RETENCION" && n <= 128) {
      sumaRetenciones += num(n);
    }
  });

  comparar_(errores, 130, sumaRetenciones - num(129), num(130), "Total retenciones renta");
  comparar_(errores, 134, num(131) + num(132) - num(133), num(134), "Total retenciones IVA");
  comparar_(errores, 136, num(130) + num(134) + num(135), num(136), "Total retenciones");
  comparar_(errores, 138, num(136) + num(137), num(138), "Total más sanciones");

  return errores;
}

/**
 * Aritmética interna del formulario 300.
 */
function validarTotales300_(valores) {
  const errores = [];
  const num = function (n) { return aNumero(valores[n]); };

  // Total ingresos brutos: renglones 27 a 40
  let ingresos = 0;
  for (let n = 27; n <= 40; n++) ingresos += num(n);
  comparar_(errores, 41, ingresos, num(41), "Total ingresos brutos");

  comparar_(errores, 43, num(41) - num(42), num(43), "Total ingresos netos");

  // Total compras e importaciones: renglones 44 a 54
  let compras = 0;
  for (let n = 44; n <= 54; n++) compras += num(n);
  comparar_(errores, 55, compras, num(55), "Total compras e importaciones brutas");

  comparar_(errores, 57, num(55) - num(56), num(57), "Total compras netas");

  // Impuesto generado: renglones 58 a 66
  let generado = 0;
  for (let n = 58; n <= 66; n++) generado += num(n);
  comparar_(errores, 67, generado, num(67), "Total impuesto generado");

  // Impuesto pagado o facturado: renglones 68 a 76
  let pagado = 0;
  for (let n = 68; n <= 76; n++) pagado += num(n);
  comparar_(errores, 77, pagado, num(77), "Total impuesto pagado o facturado");

  comparar_(errores, 81, num(77) + num(78) + num(79) + num(80), num(81),
            "Total impuestos descontables");

  comparar_(errores, 82, num(67) - num(81), num(82),
            "Saldo a pagar por el período fiscal");

  comparar_(errores, 88, num(86) + num(87), num(88), "Total saldo a pagar");

  return errores;
}

function comparar_(errores, renglon, esperado, encontrado, etiqueta) {
  if (Math.abs(esperado - encontrado) < 1) return;

  errores.push({
    tipo: "TOTAL", renglon: renglon, esperado: esperado, encontrado: encontrado,
    mensaje: "Renglón " + renglon + " · " + etiqueta + ": debería ser " +
             formatearMiles(esperado) + " y el archivo trae " + formatearMiles(encontrado)
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
                parseInt(nro, 10), aNumero(valores[nro]), ahora]);
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