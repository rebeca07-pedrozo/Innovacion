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
    const valores  = leerValores_(libroCarga, def.disposicion, catalogo);
    const detalle  = leerDetalleExterior_(libroCarga);

    const errores  = validarCarga_(catalogo, valores);
    const aceptada = errores.length === 0;

    // Los totales los calcula el sistema, no el usuario
    if (aceptada) {
      calcularTotales_(meta.cod_formulario, catalogo, valores);
    }

    const radicado = generarRadicado_();
    const idArchivado = archivarCarga_(archivo, meta.cod_entidad, aceptada);

    registrarEntrega_(radicado, meta, archivo.getName(), idArchivado,
                      aceptada, errores.length);

    if (aceptada) {
      guardarDatos_(radicado, meta, valores);
      guardarDetalle_(radicado, meta, detalle);
    } else {
      registrarErrores_(radicado, errores);
    }

    return {
      estado: aceptada ? "APROBADO" : "RECHAZADO",
      radicado: radicado,
      errores: errores,
      resumen: aceptada
        ? "Se cargaron " + Object.keys(valores).length + " renglones" +
          (detalle.length ? " y " + detalle.length + " filas de exterior" : "") +
          ". Los totales fueron calculados por el sistema."
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

  // Respaldo: el identificador de control interno del encabezado
  if (!idPlantilla) {
    const hoja = libro.getSheetByName("FORMULARIO");
    if (hoja) {
      const datos = hoja.getDataRange().getValues();
      for (let i = 0; i < Math.min(datos.length, 25); i++) {
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
      seccion:      datos[i][8],
      editable:     String(datos[i][9]).toUpperCase()
    };
  }

  return catalogo;
}


/**
 * Extrae los valores diligenciados. Todo renglón del catálogo
 * queda presente: los que no se encuentren se toman como cero.
 */
function leerValores_(libro, disposicion, catalogo) {
  const hoja = libro.getSheetByName("FORMULARIO");
  if (!hoja) throw new Error("El archivo no contiene la hoja FORMULARIO.");

  const datos = hoja.getDataRange().getValues();
  const valores = {};

  Object.keys(catalogo).forEach(function (nro) {
    valores[parseInt(nro, 10)] = 0;
  });

  if (disposicion === "MATRIZ") {
    // Números de renglón en las columnas 2, 4, 6 y 8
    [1, 3, 5, 7].forEach(function (c) {
      for (let i = 0; i < datos.length; i++) {
        const nro = parseInt(datos[i][c], 10);
        if (!isNaN(nro) && nro in valores) {
          valores[nro] = aNumero(datos[i][c + 1]);
        }
      }
    });

    // Totales de la página 1 en las columnas 8 y 9
    for (let i = 0; i < datos.length; i++) {
      const nro = parseInt(datos[i][7], 10);
      if (!isNaN(nro) && nro in valores) {
        valores[nro] = aNumero(datos[i][8]);
      }
    }

    // La hoja de exterior es opcional
    const hojaExt = libro.getSheetByName("EXTERIOR");
    if (hojaExt) {
      const datosExt = hojaExt.getDataRange().getValues();
      for (let i = 0; i < datosExt.length; i++) {
        const nro = parseInt(datosExt[i][6], 10);
        if (!isNaN(nro) && nro in valores) {
          valores[nro] = aNumero(datosExt[i][7]);
        }
      }
    }

  } else {
    // Dos bloques paralelos: renglón en 2 y 6, valor en 3 y 7
    [1, 5].forEach(function (c) {
      for (let i = 0; i < datos.length; i++) {
        const nro = parseInt(datos[i][c], 10);
        if (!isNaN(nro) && nro in valores) {
          valores[nro] = aNumero(datos[i][c + 1]);
        }
      }
    });
  }

  return valores;
}


/**
 * Lee las filas de detalle de la hoja de exterior.
 * Su diligenciamiento es opcional.
 */
function leerDetalleExterior_(libro) {
  const hoja = libro.getSheetByName("EXTERIOR");
  if (!hoja) return [];

  const datos = hoja.getDataRange().getValues();
  const filas = [];

  for (let i = 0; i < datos.length; i++) {
    const convenio = String(datos[i][0]).trim().toUpperCase();

    // Solo se toman las filas con un convenio válido diligenciado
    if (convenio !== "SIN CONVENIO" && convenio !== "CON CONVENIO") continue;

    const base = aNumero(datos[i][5]);
    const retencion = aNumero(datos[i][7]);

    // Las filas sin cifras no aportan nada al consolidado
    if (base === 0 && retencion === 0) continue;

    filas.push({
      convenio:      convenio,
      concepto_pago: String(datos[i][1]).trim(),
      tipo_persona:  String(datos[i][2]).trim(),
      pais:          String(datos[i][3]).trim(),
      cod_pais:      String(datos[i][4]).trim(),
      base:          base,
      tarifa:        aNumero(datos[i][6]),
      retencion:     retencion
    });
  }

  return filas;
}


/**
 * Comprueba únicamente que los datos sean utilizables.
 * Los totales no se validan: el sistema los recalcula.
 */
function validarCarga_(catalogo, valores) {
  const errores = [];

  Object.keys(catalogo).forEach(function (nro) {
    const valor = valores[nro];

    if (typeof valor !== "number" || isNaN(valor)) {
      errores.push({
        tipo: "DATO", renglon: nro,
        esperado: "valor numérico", encontrado: String(valor),
        mensaje: "Renglón " + nro + " · " + catalogo[nro].etiqueta +
                 ": el valor no es numérico"
      });
      return;
    }

    if (valor < 0) {
      errores.push({
        tipo: "DATO", renglon: nro,
        esperado: "valor positivo", encontrado: String(valor),
        mensaje: "Renglón " + nro + " · " + catalogo[nro].etiqueta +
                 ": no admite valores negativos"
      });
    }
  });

  return errores;
}


/**
 * Calcula los totales del formulario y reemplaza lo que traiga el archivo.
 */
function calcularTotales_(codFormulario, catalogo, valores) {
  if (codFormulario === "F350") calcularTotales350_(catalogo, valores);
  if (codFormulario === "F300") calcularTotales300_(valores);
}


function calcularTotales350_(catalogo, valores) {
  const num = function (n) { return aNumero(valores[n]); };

  let retenciones = 0;
  Object.keys(catalogo).forEach(function (nro) {
    const n = parseInt(nro, 10);
    if (catalogo[nro].tipo_valor === "RETENCION" && n <= 128) {
      retenciones += num(n);
    }
  });

  valores[130] = retenciones - num(129);
  valores[134] = num(131) + num(132) - num(133);
  valores[136] = num(130) + num(134) + num(135);
  valores[138] = num(136) + num(137);
}


function calcularTotales300_(valores) {
  const num = function (n) { return aNumero(valores[n]); };

  let ingresos = 0;
  for (let n = 27; n <= 40; n++) ingresos += num(n);
  valores[41] = ingresos;
  valores[43] = num(41) - num(42);

  let compras = 0;
  for (let n = 44; n <= 54; n++) compras += num(n);
  valores[55] = compras;
  valores[57] = num(55) - num(56);

  let generado = 0;
  for (let n = 58; n <= 66; n++) generado += num(n);
  valores[67] = generado;

  let pagado = 0;
  for (let n = 68; n <= 76; n++) pagado += num(n);
  valores[77] = pagado;

  valores[81] = num(77) + num(78) + num(79) + num(80);

  // El resultado del período va a saldo a pagar o a saldo a favor
  const diferencia = num(67) - num(81);
  valores[82] = diferencia > 0 ? diferencia : 0;
  valores[83] = diferencia < 0 ? -diferencia : 0;

  valores[86] = num(82) - num(84) - num(85);
  if (valores[86] < 0) valores[86] = 0;

  valores[88] = num(86) + num(87);
  valores[89] = num(83) + num(84) + num(85) - num(82);
  if (valores[89] < 0) valores[89] = 0;

  valores[93] = num(91) + num(92);

  let anticipos = 0;
  for (let n = 1; n <= 6; n++) anticipos += num(n);
  valores[100] = anticipos;
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


/**
 * Guarda las filas de detalle de exterior.
 */
function guardarDetalle_(radicado, meta, detalle) {
  if (!detalle.length) return;

  const hoja = SpreadsheetApp.openById(ID_OPERACION).getSheetByName("DATOS_DETALLE");
  const ahora = new Date();

  const filas = detalle.map(function (d, i) {
    return [radicado, meta.cod_formulario, "EXTERIOR", i + 1,
            d.convenio, d.concepto_pago, d.tipo_persona, d.pais, d.cod_pais,
            d.base, d.tarifa, d.retencion, ahora];
  });

  hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, filas[0].length)
      .setValues(filas);
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