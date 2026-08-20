/**
 * Genera el consolidado de un formulario y periodo.
 *
 * @param {string} codFormulario  Código del formulario
 * @param {string} periodoTxt     Periodo en formato "2026-06"
 * @return {Object} { idConsolidado, url, urlDescarga, entidades, totales }
 */
function generarConsolidado(codFormulario, periodoTxt) {
  const anio = periodoTxt.split("-")[0];
  const per  = periodoTxt.split("-")[1];

  const def = definicionFormulario(codFormulario);
  const entradas = radicadosAprobados_(codFormulario, anio, per);

  if (entradas.length < 2) {
    throw new Error("Se requiere una carga aprobada de cada entidad. " +
                    "Actualmente hay " + entradas.length + ".");
  }

  const catalogo = leerCatalogoValidacion_(codFormulario, def.version);
  const sumas    = sumarEntradas_(entradas, catalogo);

  const idConsolidado = generarIdConsolidado_();
  const renglones = leerRenglones_(codFormulario, def.version);

  const libro = construirLibroConsolidado_(
    idConsolidado, def, anio, per, renglones, sumas, entradas
  );

  const archivo = DriveApp.getFileById(libro.getId());
  DriveApp.getFolderById(CARPETA_CONSOLIDADOS).addFile(archivo);
  DriveApp.getRootFolder().removeFile(archivo);
  archivo.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);

  registrarConsolidado_(idConsolidado, codFormulario, anio, per,
                        entradas, archivo.getId());

  return {
    idConsolidado: idConsolidado,
    formulario:    codFormulario,
    periodo:       anio + "-" + per,
    entidades:     entradas.map(function (e) { return e.entidad; }),
    url:           libro.getUrl(),
    urlDescarga:   "https://docs.google.com/spreadsheets/d/" + archivo.getId() +
                   "/export?format=xlsx"
  };
}


/**
 * Devuelve el radicado aprobado más reciente de cada entidad.
 */
function radicadosAprobados_(codFormulario, anio, per) {
  const datos = SpreadsheetApp.openById(ID_OPERACION)
                  .getSheetByName("ENTREGAS").getDataRange().getValues();
  const porEntidad = {};

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][2] !== codFormulario) continue;
    if (String(datos[i][4]) !== String(anio)) continue;
    if (String(datos[i][5]) !== String(per)) continue;
    if (String(datos[i][10]).toUpperCase() !== "APROBADO") continue;

    // La entrega posterior reemplaza a la anterior de la misma entidad
    porEntidad[datos[i][3]] = {
      radicado: datos[i][0],
      entidad:  datos[i][3],
      fecha:    formatearFecha(datos[i][9])
    };
  }

  return Object.keys(porEntidad).map(function (k) { return porEntidad[k]; });
}


/**
 * Suma los valores de todos los radicados, renglón por renglón.
 */
function sumarEntradas_(entradas, catalogo) {
  const datos = SpreadsheetApp.openById(ID_OPERACION)
                  .getSheetByName("DATOS_CARGADOS").getDataRange().getValues();

  const radicados = entradas.map(function (e) { return e.radicado; });
  const sumas = { TOTAL: {} };

  entradas.forEach(function (e) { sumas[e.entidad] = {}; });

  // Todo renglón del catálogo parte en cero
  Object.keys(catalogo).forEach(function (nro) {
    const n = parseInt(nro, 10);
    sumas.TOTAL[n] = 0;
    entradas.forEach(function (e) { sumas[e.entidad][n] = 0; });
  });

  for (let i = 1; i < datos.length; i++) {
    const posicion = radicados.indexOf(datos[i][0]);
    if (posicion < 0) continue;

    const nro = parseInt(datos[i][3], 10);
    if (!(nro in sumas.TOTAL)) continue;

    const valor = aNumero(datos[i][4]);
    sumas[entradas[posicion].entidad][nro] = valor;
    sumas.TOTAL[nro] += valor;
  }

  return sumas;
}


/**
 * Construye el libro del consolidado.
 */
function construirLibroConsolidado_(idConsolidado, def, anio, per,
                                    renglones, sumas, entradas) {
  const nombre = "CONS_" + def.codigo + "_" + anio + "-" + per + "_" + idConsolidado;
  const libro = SpreadsheetApp.create(nombre);

  construirHojaConsolidado_(libro, idConsolidado, def, anio, per,
                            renglones, sumas, entradas);

  const porDefecto = libro.getSheetByName("Hoja 1") || libro.getSheetByName("Sheet1");
  if (porDefecto) libro.deleteSheet(porDefecto);

  return libro;
}


/**
 * Escribe la hoja del consolidado con una columna por entidad.
 */
function construirHojaConsolidado_(libro, idConsolidado, def, anio, per,
                                   renglones, sumas, entradas) {
  const hoja = libro.insertSheet("CONSOLIDADO");
  const nombresEntidad = entradas.map(function (e) { return e.entidad; });
  const ancho = 2 + nombresEntidad.length + 1;

  let f = 1;

  // ===== Aviso obligatorio =====
  hoja.getRange(f, 1, 1, ancho).merge()
      .setValue("BORRADOR - DOCUMENTO DE TRABAJO INTERNO - " +
                "NO CONSTITUYE DECLARACIÓN TRIBUTARIA")
      .setBackground(AMARILLO_AVISO)
      .setFontSize(10).setFontWeight("bold")
      .setHorizontalAlignment("center");
  f += 2;

  // ===== Título =====
  hoja.getRange(f, 1, 1, ancho - 1).merge()
      .setValue("Consolidado · " + def.nombre)
      .setFontSize(13).setFontWeight("bold")
      .setBackground(AZUL_TITULO).setFontColor(BLANCO)
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
  hoja.getRange(f, ancho)
      .setValue(def.codigo.replace(/\D/g, ""))
      .setFontSize(22).setFontWeight("bold")
      .setBackground(AZUL_TITULO).setFontColor(BLANCO)
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
  hoja.setRowHeight(f, 34);
  f += 2;

  // ===== Datos de la consolidación =====
  const inicioDatos = f;
  const detalle = [
    ["Identificador", idConsolidado],
    ["Formulario", def.codigo + " · " + def.version],
    ["Año", anio],
    ["Período", per],
    ["Generado por", Session.getActiveUser().getEmail()],
    ["Fecha de generación", formatearFecha(new Date())],
    ["Radicados incluidos", entradas.map(function (e) {
      return e.entidad + ": " + e.radicado;
    }).join("  ·  ")]
  ];

  detalle.forEach(function (fila) {
    hoja.getRange(f, 1).setValue(fila[0]).setFontWeight("bold");
    hoja.getRange(f, 2, 1, ancho - 1).merge().setValue(fila[1]);
    f++;
  });

  hoja.getRange(inicioDatos, 1, detalle.length, ancho)
      .setBackground(GRIS_BLOQUEO).setFontSize(9);
  f++;

  // ===== Cabecera de la tabla =====
  const cabecera = ["Renglón", "Concepto"]
    .concat(nombresEntidad)
    .concat(["Consolidado"]);

  hoja.getRange(f, 1, 1, ancho).setValues([cabecera])
      .setFontWeight("bold").setBackground(AZUL_CABECERA)
      .setHorizontalAlignment("center").setWrap(true);
  hoja.setRowHeight(f, 28);
  f++;

  // ===== Cuerpo =====
  const inicioTabla = f;
  const ordenados = renglones.slice().sort(function (a, b) {
    return a.nro_renglon - b.nro_renglon;
  });

  let seccionActual = "";

  ordenados.forEach(function (r) {
    if (r.seccion !== seccionActual) {
      seccionActual = r.seccion;
      hoja.getRange(f, 1, 1, ancho).merge()
          .setValue(titularSeccion(seccionActual))
          .setFontWeight("bold").setBackground(AZUL_SECCION).setFontSize(9);
      f++;
    }

    const calculado = String(r.editable).toUpperCase() === "NO";
    const fila = [r.nro_renglon, r.etiqueta];

    nombresEntidad.forEach(function (ent) {
      fila.push(sumas[ent][r.nro_renglon] || 0);
    });
    fila.push(sumas.TOTAL[r.nro_renglon] || 0);

    hoja.getRange(f, 1, 1, ancho).setValues([fila]).setFontSize(9);
    hoja.getRange(f, 1).setHorizontalAlignment("center").setFontSize(8);
    hoja.getRange(f, 2).setWrap(true);

    if (calculado) {
      hoja.getRange(f, 1, 1, ancho).setFontWeight("bold").setBackground("#F5F5F5");
    }

    f++;
  });

  const ultima = f - 1;

  // ===== Formato =====
  hoja.getRange(inicioTabla, 1, ultima - inicioTabla + 1, ancho)
      .setBorder(true, true, true, true, true, true, "#B0B0B0",
                 SpreadsheetApp.BorderStyle.SOLID);

  hoja.getRange(inicioTabla, 3, ultima - inicioTabla + 1, ancho - 2)
      .setNumberFormat("#,##0").setHorizontalAlignment("right");

  // La columna del consolidado se destaca sobre las de origen
  hoja.getRange(inicioTabla, ancho, ultima - inicioTabla + 1, 1)
      .setBackground("#E8F0FA").setFontWeight("bold");

  hoja.setColumnWidth(1, 70);
  hoja.setColumnWidth(2, 380);
  for (let c = 3; c <= ancho; c++) hoja.setColumnWidth(c, 150);

  hoja.setFrozenRows(inicioTabla - 1);
  hoja.setFrozenColumns(2);
}


/**
 * Registra el consolidado en el libro transaccional.
 */
function registrarConsolidado_(idConsolidado, codFormulario, anio, per,
                               entradas, idArchivo) {
  const hoja = obtenerHojaConsolidados_();

  // Los consolidados anteriores del mismo periodo dejan de ser vigentes
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][1] === codFormulario &&
        String(datos[i][2]) === String(anio) &&
        String(datos[i][3]) === String(per)) {
      hoja.getRange(i + 1, 8).setValue("NO");
    }
  }

  hoja.appendRow([
    idConsolidado, codFormulario, anio, per,
    entradas.map(function (e) { return e.radicado; }).join(", "),
    idArchivo, Session.getActiveUser().getEmail(), "SI", new Date()
  ]);
}


/**
 * Devuelve la hoja de consolidados, creándola la primera vez.
 */
function obtenerHojaConsolidados_() {
  const libro = SpreadsheetApp.openById(ID_OPERACION);
  let hoja = libro.getSheetByName("CONSOLIDADOS");

  if (!hoja) {
    hoja = libro.insertSheet("CONSOLIDADOS");
    const encabezados = ["id_consolidado", "cod_formulario", "anio", "periodo",
                         "radicados_origen", "id_archivo_drive", "generado_por",
                         "vigente", "fecha_generacion"];
    hoja.getRange(1, 1, 1, encabezados.length)
        .setValues([encabezados])
        .setFontWeight("bold")
        .setBackground(COLOR_ENCABEZADO);
    hoja.setFrozenRows(1);
  }

  return hoja;
}


function generarIdConsolidado_() {
  const hoja = obtenerHojaConsolidados_();
  return "CONS-" + new Date().getFullYear() + "-" +
         ("0000" + hoja.getLastRow()).slice(-4);
}


/**
 * Busca el consolidado vigente de un formulario y periodo.
 */
function buscarConsolidadoVigente(codFormulario, periodoTxt) {
  const anio = periodoTxt.split("-")[0];
  const per  = periodoTxt.split("-")[1];

  const hoja = obtenerHojaConsolidados_();
  if (hoja.getLastRow() <= 1) return { encontrado: false };

  const datos = hoja.getDataRange().getValues();

  for (let i = datos.length - 1; i >= 1; i--) {
    if (datos[i][1] !== codFormulario) continue;
    if (String(datos[i][2]) !== anio) continue;
    if (String(datos[i][3]) !== per) continue;
    if (String(datos[i][7]).toUpperCase() !== "SI") continue;

    return {
      encontrado:    true,
      idConsolidado: datos[i][0],
      generadoPor:   datos[i][6],
      fecha:         formatearFecha(datos[i][8]),
      urlDescarga:   "https://docs.google.com/spreadsheets/d/" + datos[i][5] +
                     "/export?format=xlsx"
    };
  }

  return { encontrado: false };
}