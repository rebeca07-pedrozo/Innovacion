/**
 * ============================================================
 * APLICACIÓN WEB
 * Punto de entrada y funciones que invoca la interfaz.
 * ============================================================
 */

function doGet() {
  return HtmlService.createTemplateFromFile("Interfaz")
    .evaluate()
    .setTitle("Sistema de declaraciones")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}


function obtenerDatosIniciales() {
  const formularios = listarFormularios_();
  const primero = formularios.length ? formularios[0].codigo : "F350";

  return {
    usuario:     Session.getActiveUser().getEmail(),
    formularios: formularios,
    entidades:   listarEntidades_(),
    anios:       construirAnios_(),
    detalle:     detallePeriodicidad(primero)
  };
}


/**
 * Devuelve la periodicidad de un formulario y sus periodos disponibles.
 */
function detallePeriodicidad(codFormulario) {
  const def = definicionFormulario(codFormulario);
  const periodicidad = String(def.periodicidad).toUpperCase();

  const etiquetas = {
    MENSUAL: ["enero", "febrero", "marzo", "abril", "mayo", "junio",
              "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"],
    BIMESTRAL: ["ene-feb", "mar-abr", "may-jun", "jul-ago", "sep-oct", "nov-dic"],
    TRIMESTRAL: ["ene-mar", "abr-jun", "jul-sep", "oct-dic"],
    CUATRIMESTRAL: ["ene-abr", "may-ago", "sep-dic"],
    ANUAL: ["año completo"]
  };

  const rotulos = {
    MENSUAL: "Mes",
    BIMESTRAL: "Bimestre",
    TRIMESTRAL: "Trimestre",
    CUATRIMESTRAL: "Cuatrimestre",
    ANUAL: "Año"
  };

  const legibles = {
    MENSUAL: "Mensual",
    BIMESTRAL: "Bimestral",
    TRIMESTRAL: "Trimestral",
    CUATRIMESTRAL: "Cuatrimestral",
    ANUAL: "Anual"
  };

  const nombres = etiquetas[periodicidad] || etiquetas.MENSUAL;

  const periodos = nombres.map(function (nombre, i) {
    return {
      valor: ("0" + (i + 1)).slice(-2),
      texto: (i + 1) + " — " + nombre
    };
  });

  return {
    periodicidad: legibles[periodicidad] || "Mensual",
    rotulo:       rotulos[periodicidad] || "Mes",
    periodos:     periodos
  };
}


/**
 * Entrega la plantilla de una combinación, generándola si no existe.
 */
function obtenerPlantilla(codFormulario, codEntidad, periodoTxt) {
  const anio = periodoTxt.split("-")[0];
  const per  = periodoTxt.split("-")[1];

  const datos = SpreadsheetApp.openById(ID_OPERACION)
                  .getSheetByName("PLANTILLAS_EMITIDAS").getDataRange().getValues();

  // Se recorre de abajo hacia arriba para tomar la emisión más reciente
  for (let i = datos.length - 1; i >= 1; i--) {
    if (datos[i][1] !== codFormulario) continue;
    if (datos[i][3] !== codEntidad) continue;
    if (String(datos[i][4]) !== anio) continue;
    if (String(datos[i][5]) !== per) continue;

    return {
      idPlantilla: datos[i][0],
      generadaPor: datos[i][7],
      fecha:       formatearFecha(datos[i][8]),
      reutilizada: true,
      urlDescarga: "https://docs.google.com/spreadsheets/d/" + datos[i][6] +
                   "/export?format=xlsx"
    };
  }

  // No existía: se genera en el momento
  const r = generarPlantilla(codFormulario, codEntidad,
                             parseInt(anio, 10), parseInt(per, 10));

  return {
    idPlantilla: r.idPlantilla,
    generadaPor: Session.getActiveUser().getEmail(),
    fecha:       formatearFecha(new Date()),
    reutilizada: false,
    urlDescarga: "https://docs.google.com/spreadsheets/d/" + r.idArchivo +
                 "/export?format=xlsx"
  };
}


/**
 * Devuelve el estado de las entidades para un formulario y periodo.
 */
function obtenerEstadoPeriodo(codFormulario, periodoTxt) {
  const anio = periodoTxt.split("-")[0];
  const per  = periodoTxt.split("-")[1];

  const datos = SpreadsheetApp.openById(ID_OPERACION)
                  .getSheetByName("ENTREGAS").getDataRange().getValues();
  const estado = {};

  listarEntidades_().forEach(function (e) {
    estado[e.codigo] = { entidad: e.nombre, estado: "PENDIENTE", fecha: "", radicado: "" };
  });

  // La entrega más reciente define el estado de cada entidad
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][2] !== codFormulario) continue;
    if (String(datos[i][4]) !== anio) continue;
    if (String(datos[i][5]) !== per) continue;

    const codEntidad = datos[i][3];
    if (!estado[codEntidad]) continue;

    estado[codEntidad] = {
      entidad:  estado[codEntidad].entidad,
      estado:   datos[i][10],
      fecha:    formatearFecha(datos[i][9]),
      radicado: datos[i][0]
    };
  }

  return { formulario: codFormulario, periodo: anio + "-" + per, entidades: estado };
}


/**
 * Recibe el archivo enviado desde el navegador y lo procesa.
 */
function recibirCarga(datosArchivo, nombreArchivo) {
  let idTemporal = null;

  try {
    const bytes = Utilities.base64Decode(datosArchivo);
    const blob  = Utilities.newBlob(bytes, MimeType.MICROSOFT_EXCEL, nombreArchivo);

    const temporal = DriveApp.getFolderById(CARPETA_CARGAS).createFile(blob);
    idTemporal = temporal.getId();

    const resultado = procesarCarga(idTemporal);

    return {
      ok: true,
      estado:   resultado.estado,
      radicado: resultado.radicado,
      resumen:  resultado.resumen,
      errores:  resultado.errores.map(function (e) {
        return { renglon: e.renglon, mensaje: e.mensaje };
      })
    };

  } catch (e) {
    return { ok: false, mensaje: e.message };

  } finally {
    // La copia definitiva ya quedó archivada por procesarCarga
    if (idTemporal) {
      try { DriveApp.getFileById(idTemporal).setTrashed(true); } catch (err) {}
    }
  }
}


// ---------- Lecturas del catálogo ----------

function listarFormularios_() {
  const datos = SpreadsheetApp.openById(ID_REGISTRO)
                  .getSheetByName("FORMULARIOS").getDataRange().getValues();
  const lista = [];

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][7]).toUpperCase() !== "ACTIVO") continue;
    lista.push({ codigo: datos[i][0], nombre: datos[i][1], version: datos[i][2] });
  }

  return lista;
}


function listarEntidades_() {
  const datos = SpreadsheetApp.openById(ID_REGISTRO)
                  .getSheetByName("ENTIDADES").getDataRange().getValues();
  const lista = [];

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][6]).toUpperCase() !== "SI") continue;
    lista.push({ codigo: datos[i][0], nombre: datos[i][1] });
  }

  return lista;
}


/**
 * Devuelve los años disponibles, del actual hacia atrás.
 */
function construirAnios_() {
  const actual = new Date().getFullYear();
  const lista = [];

  for (let a = actual; a >= actual - 3; a--) {
    lista.push({ valor: String(a), texto: String(a) });
  }

  return lista;
}