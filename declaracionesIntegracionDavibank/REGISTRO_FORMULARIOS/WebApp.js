function doGet() {
  return HtmlService.createTemplateFromFile("Interfaz")
    .evaluate()
    .setTitle("Sistema de declaraciones")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function obtenerDatosIniciales() {
  return {
    usuario:     Session.getActiveUser().getEmail(),
    formularios: listarFormularios_(),
    entidades:   listarEntidades_(),
    periodos:    construirPeriodos_(),
    estado:      obtenerEstadoPeriodo(null, null)
  };
}
function obtenerEstadoPeriodo(codFormulario, periodoTxt) {
  const hoy = new Date();
  const cod = codFormulario || "F350";
  const anio = periodoTxt ? periodoTxt.split("-")[0] : String(hoy.getFullYear());
  const per  = periodoTxt ? periodoTxt.split("-")[1]
                          : ("0" + (hoy.getMonth() + 1)).slice(-2);

  const datos = SpreadsheetApp.openById(ID_OPERACION)
                  .getSheetByName("ENTREGAS").getDataRange().getValues();

  const estado = {};

  listarEntidades_().forEach(function (e) {
    estado[e.codigo] = { entidad: e.nombre, estado: "PENDIENTE", fecha: "", radicado: "" };
  });
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][2] !== cod) continue;
    if (String(datos[i][4]) !== anio) continue;
    if (String(datos[i][5]) !== per) continue;

    const codEntidad = datos[i][3];
    if (!estado[codEntidad]) continue;

    estado[codEntidad] = {
      entidad:  estado[codEntidad].entidad,
      estado:   datos[i][10],
      fecha:    formatearFecha_(datos[i][9]),
      radicado: datos[i][0]
    };
  }

  return { formulario: cod, periodo: anio + "-" + per, entidades: estado };
}
function buscarPlantillaExistente(codFormulario, codEntidad, periodoTxt) {
  const anio = periodoTxt.split("-")[0];
  const per  = periodoTxt.split("-")[1];

  const datos = SpreadsheetApp.openById(ID_OPERACION)
                  .getSheetByName("PLANTILLAS_EMITIDAS").getDataRange().getValues();

  for (let i = datos.length - 1; i >= 1; i--) {
    if (datos[i][1] !== codFormulario) continue;
    if (datos[i][3] !== codEntidad) continue;
    if (String(datos[i][4]) !== anio) continue;
    if (String(datos[i][5]) !== per) continue;

    return {
      encontrada:  true,
      idPlantilla: datos[i][0],
      generadaPor: datos[i][7],
      fecha:       formatearFecha_(datos[i][8]),
      url:         "https://docs.google.com/spreadsheets/d/" + datos[i][6] + "/edit",
      urlDescarga: "https://docs.google.com/spreadsheets/d/" + datos[i][6] +
                   "/export?format=xlsx"
    };
  }

  return { encontrada: false };
}
function generarPlantillaDesdeWeb(codFormulario, codEntidad, periodoTxt) {
  const anio    = parseInt(periodoTxt.split("-")[0], 10);
  const periodo = parseInt(periodoTxt.split("-")[1], 10);
  const version = versionVigente_(codFormulario);

  const r = generarPlantilla(codFormulario, version, codEntidad, anio, periodo);

  return {
    encontrada:  true,
    idPlantilla: r.idPlantilla,
    generadaPor: Session.getActiveUser().getEmail(),
    fecha:       formatearFecha_(new Date()),
    url:         r.url,
    urlDescarga: "https://docs.google.com/spreadsheets/d/" + r.idArchivo +
                 "/export?format=xlsx"
  };
}
function recibirCarga(datosArchivo, nombreArchivo) {
  let idTemporal = null;

  try {
    const bytes = Utilities.base64Decode(datosArchivo);
    const blob  = Utilities.newBlob(bytes, MimeType.MICROSOFT_EXCEL, nombreArchivo);

    // El archivo se deposita temporalmente antes de validarlo
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
    if (idTemporal) {
      try { DriveApp.getFileById(idTemporal).setTrashed(true); } catch (err) {}
    }
  }
}
function listarFormularios_() {
  const datos = SpreadsheetApp.openById(ID_REGISTRO)
                  .getSheetByName("FORMULARIOS").getDataRange().getValues();
  const lista = [];

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][5] !== "ACTIVO") continue;
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

function versionVigente_(codFormulario) {
  const datos = SpreadsheetApp.openById(ID_REGISTRO)
                  .getSheetByName("FORMULARIOS").getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === codFormulario && datos[i][5] === "ACTIVO") {
      return datos[i][2];
    }
  }
  throw new Error("No hay versión activa del formulario " + codFormulario);
}
function construirPeriodos_() {
  const meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio",
                 "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
  const lista = [];
  const hoy = new Date();

  for (let i = 0; i < 12; i++) {
    const f = new Date(hoy.getFullYear(), hoy.getMonth() - i, 1);
    const anio = f.getFullYear();
    const mes  = ("0" + (f.getMonth() + 1)).slice(-2);

    lista.push({
      valor: anio + "-" + mes,
      texto: anio + " · " + mes + " — " + meses[f.getMonth()]
    });
  }

  return lista;
}

function formatearFecha_(fecha) {
  if (!fecha) return "";
  return Utilities.formatDate(new Date(fecha), "America/Bogota", "dd/MM/yyyy HH:mm");
}