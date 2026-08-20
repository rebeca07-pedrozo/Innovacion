function generarConsolidado(codFormulario, periodoTxt) {
  const anio = periodoTxt.split("-")[0];
  const per  = periodoTxt.split("-")[1];

  const def = definicionFormulario(codFormulario);
  const entradas = radicadosAprobados_(codFormulario, anio, per);

  if (entradas.length < 2) {
    throw new Error("Se requiere una carga aprobada de cada entidad. " +
                    "Actualmente hay " + entradas.length + ".");
  }

  const catalogo  = leerCatalogoValidacion_(codFormulario, def.version);
  const renglones = leerRenglones_(codFormulario, def.version);
  const sumas     = sumarEntradas_(entradas, catalogo);
  const detalle   = agruparDetalle_(entradas);

  const idConsolidado = generarIdConsolidado_();

  const nombre = "CONS_" + def.codigo + "_" + anio + "-" + per + "_" + idConsolidado;
  const libro = SpreadsheetApp.create(nombre);

  if (def.disposicion === "MATRIZ") {
    consolidadoMatriz_(libro, renglones, sumas, def, anio, per,
                       idConsolidado, entradas);
    consolidadoExterior_(libro, renglones, sumas, detalle);
  } else {
    consolidadoDobleColumna_(libro, renglones, sumas, def, anio, per,
                             idConsolidado, entradas);
  }

  const porDefecto = libro.getSheetByName("Hoja 1") || libro.getSheetByName("Sheet1");
  if (porDefecto) libro.deleteSheet(porDefecto);

  libro.setActiveSheet(libro.getSheetByName("CONSOLIDADO"));

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
  const sumas = {};

  Object.keys(catalogo).forEach(function (nro) {
    sumas[parseInt(nro, 10)] = 0;
  });

  for (let i = 1; i < datos.length; i++) {
    if (radicados.indexOf(datos[i][0]) < 0) continue;

    const nro = parseInt(datos[i][3], 10);
    if (!(nro in sumas)) continue;

    sumas[nro] += aNumero(datos[i][4]);
  }

  return sumas;
}


/**
 * Agrupa el detalle de exterior de todas las entidades.
 * Las filas con igual convenio, concepto, tipo de persona,
 * país y tarifa se suman en una sola.
 */
function agruparDetalle_(entradas) {
  const libro = SpreadsheetApp.openById(ID_OPERACION);
  const hoja = libro.getSheetByName("DATOS_DETALLE");
  if (!hoja || hoja.getLastRow() <= 1) return [];

  const datos = hoja.getDataRange().getValues();
  const radicados = entradas.map(function (e) { return e.radicado; });
  const grupos = {};
  const orden = [];

  for (let i = 1; i < datos.length; i++) {
    if (radicados.indexOf(datos[i][0]) < 0) continue;
    if (datos[i][2] !== "EXTERIOR") continue;

    const clave = [datos[i][4], datos[i][5], datos[i][6],
                   datos[i][7], datos[i][10]].join("|");

    if (!grupos[clave]) {
      grupos[clave] = {
        convenio:      datos[i][4],
        concepto_pago: datos[i][5],
        tipo_persona:  datos[i][6],
        pais:          datos[i][7],
        cod_pais:      datos[i][8],
        tarifa:        datos[i][10],
        base:          0,
        retencion:     0
      };
      orden.push(clave);
    }

    grupos[clave].base      += aNumero(datos[i][9]);
    grupos[clave].retencion += aNumero(datos[i][11]);
  }

  return orden.map(function (k) { return grupos[k]; });
}


/**
 * Escribe el encabezado del borrador consolidado.
 */
function encabezadoConsolidado_(hoja, def, anio, per, idConsolidado,
                                entradas, ancho) {
  let f = 1;

  hoja.getRange(f, 1, 1, ancho).merge()
      .setValue("BORRADOR - DOCUMENTO DE TRABAJO INTERNO - " +
                "NO CONSTITUYE DECLARACIÓN TRIBUTARIA")
      .setBackground(AMARILLO_AVISO)
      .setFontSize(10).setFontWeight("bold")
      .setHorizontalAlignment("center");
  f += 2;

  hoja.getRange(f, 1, 1, ancho - 1).merge()
      .setValue(def.nombre + " · consolidado")
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

  hoja.getRange(f, 1, 1, ancho).merge()
      .setValue("Datos de la consolidación")
      .setBackground(AZUL_SECCION).setFontWeight("bold");
  f++;

  const inicio = f;
  const campos = [
    ["1. Año", anio],
    ["3. Período", per],
    ["Entidades consolidadas", entradas.map(function (e) {
      return e.entidad;
    }).join("  +  ")],
    ["Radicados de origen", entradas.map(function (e) {
      return e.entidad + ": " + e.radicado;
    }).join("   ·   ")],
    ["Generado por", Session.getActiveUser().getEmail()],
    ["Fecha de generación", formatearFecha(new Date())]
  ];

  campos.forEach(function (campo) {
    hoja.getRange(f, 1).setValue(campo[0]).setFontWeight("bold");
    hoja.getRange(f, 2, 1, ancho - 1).merge().setValue(campo[1]);
    f++;
  });

  hoja.getRange(inicio, 1, campos.length, ancho)
      .setBackground(GRIS_BLOQUEO).setFontSize(9);

  hoja.getRange(f, 1).setValue("Control interno")
      .setFontWeight("bold").setFontSize(8).setFontColor("#888888");
  hoja.getRange(f, 2, 1, ancho - 1).merge().setValue(idConsolidado)
      .setFontSize(8).setFontColor("#888888");
  hoja.getRange(f, 1, 1, ancho).setBackground("#F7F7F7");
  hoja.setRowHeight(f, 16);
  f++;

  return f + 1;
}


/**
 * Consolidado con disposición de matriz, como el formulario 350.
 */
function consolidadoMatriz_(libro, renglones, sumas, def, anio, per,
                            idConsolidado, entradas) {
  const hoja = libro.insertSheet("CONSOLIDADO");
  let f = encabezadoConsolidado_(hoja, def, anio, per, idConsolidado, entradas, 9);

  hoja.getRange(f, 1, 2, 1).merge().setValue("Concepto")
      .setFontWeight("bold").setBackground(AZUL_CABECERA)
      .setVerticalAlignment("middle");
  hoja.getRange(f, 2, 1, 4).merge().setValue("A personas jurídicas")
      .setFontWeight("bold").setBackground(AZUL_CABECERA)
      .setHorizontalAlignment("center");
  hoja.getRange(f, 6, 1, 4).merge().setValue("A personas naturales")
      .setFontWeight("bold").setBackground(AZUL_CABECERA)
      .setHorizontalAlignment("center");
  f++;

  hoja.getRange(f, 1, 1, 9).setBackground(AZUL_CABECERA).setFontSize(9)
      .setHorizontalAlignment("center").setWrap(true);
  hoja.getRange(f, 2, 1, 2).merge().setValue("Base sujeta a retención");
  hoja.getRange(f, 4, 1, 2).merge().setValue("Retenciones a título de renta");
  hoja.getRange(f, 6, 1, 2).merge().setValue("Base sujeta a retención");
  hoja.getRange(f, 8, 1, 2).merge().setValue("Retenciones a título de renta");
  hoja.setRowHeight(f, 30);
  f++;

  const inicio = f;
  const matriz = agruparEnMatriz_(renglones);
  const totales = renglones.filter(function (r) {
    return String(r.seccion).indexOf("TOTALES") === 0 && r.seccion !== "TOTALES_EXTERIOR";
  }).sort(function (a, b) { return a.nro_renglon - b.nro_renglon; });

  let seccionActual = "";

  matriz.forEach(function (fila) {
    if (fila.seccion !== seccionActual) {
      seccionActual = fila.seccion;
      hoja.getRange(f, 1, 1, 9).merge()
          .setValue(titularSeccion(seccionActual))
          .setFontWeight("bold").setBackground(AZUL_SECCION).setFontSize(9);
      f++;
    }

    hoja.getRange(f, 1).setValue(fila.etiqueta).setWrap(true).setFontSize(9);

    [[fila.jurBase, 2], [fila.jurRet, 4], [fila.natBase, 6], [fila.natRet, 8]]
      .forEach(function (par) {
        if (par[0] === null) {
          hoja.getRange(f, par[1], 1, 2).setBackground(GRIS_BLOQUEO);
          return;
        }
        hoja.getRange(f, par[1]).setValue(par[0])
            .setFontSize(8).setHorizontalAlignment("center")
            .setBackground(AZUL_CABECERA);
        hoja.getRange(f, par[1] + 1).setValue(sumas[par[0]] || 0)
            .setBackground(BLANCO);
      });

    f++;
  });

  totales.forEach(function (t) {
    const calculado = String(t.editable).toUpperCase() === "NO";

    hoja.getRange(f, 1, 1, 7).merge().setValue(t.etiqueta)
        .setWrap(true).setFontSize(9)
        .setFontWeight(calculado ? "bold" : "normal");
    hoja.getRange(f, 8).setValue(t.nro_renglon)
        .setFontSize(8).setHorizontalAlignment("center")
        .setBackground(AZUL_CABECERA);
    hoja.getRange(f, 9).setValue(sumas[t.nro_renglon] || 0)
        .setBackground(calculado ? "#EDF3FA" : BLANCO)
        .setFontWeight(calculado ? "bold" : "normal");
    f++;
  });

  const ultima = f - 1;
  hoja.getRange(inicio, 1, ultima - inicio + 1, 9)
      .setBorder(true, true, true, true, true, true, "#B0B0B0",
                 SpreadsheetApp.BorderStyle.SOLID);

  [3, 5, 7, 9].forEach(function (col) {
    hoja.getRange(inicio, col, ultima - inicio + 1, 1)
        .setNumberFormat("#,##0").setHorizontalAlignment("right");
  });

  hoja.setColumnWidth(1, 300);
  [2, 4, 6, 8].forEach(function (c) { hoja.setColumnWidth(c, 32); });
  [3, 5, 7, 9].forEach(function (c) { hoja.setColumnWidth(c, 130); });
}


/**
 * Hoja de exterior del consolidado, con el detalle ya agrupado.
 */
function consolidadoExterior_(libro, renglones, sumas, detalle) {
  const totalesExt = renglones.filter(function (r) {
    return r.seccion === "TOTALES_EXTERIOR";
  }).sort(function (a, b) { return a.nro_renglon - b.nro_renglon; });

  if (!totalesExt.length) return;

  const hoja = libro.insertSheet("EXTERIOR");

  hoja.getRange(1, 1, 1, 8).merge()
      .setValue("Pagos o abonos en cuenta al exterior - consolidado")
      .setFontSize(12).setFontWeight("bold")
      .setBackground(AZUL_TITULO).setFontColor(BLANCO)
      .setHorizontalAlignment("center");
  hoja.setRowHeight(1, 28);

  hoja.getRange(2, 1, 1, 8).merge()
      .setValue("Las filas de las entidades se agruparon por convenio, " +
                "concepto, tipo de persona, país y tarifa.")
      .setFontSize(9).setFontColor("#666666")
      .setHorizontalAlignment("center");

  hoja.getRange(4, 1, 1, 8).setValues([[
    "141. A países", "142. Concepto de pago", "143. Tipo de persona",
    "144. País", "Cód.", "145. Pagos o abonos en cuenta",
    "146. Tarifa (%)", "147. Valor retención"
  ]]).setBackground(AZUL_CABECERA).setFontWeight("bold").setFontSize(9)
     .setWrap(true).setHorizontalAlignment("center");
  hoja.setRowHeight(4, 34);

  let f = 5;

  if (detalle.length) {
    const filas = detalle.map(function (d) {
      return [d.convenio, d.concepto_pago, d.tipo_persona, d.pais,
              d.cod_pais, d.base, d.tarifa, d.retencion];
    });

    hoja.getRange(f, 1, filas.length, 8).setValues(filas).setFontSize(9);
    hoja.getRange(f, 6, filas.length, 1).setNumberFormat("#,##0");
    hoja.getRange(f, 8, filas.length, 1).setNumberFormat("#,##0");
    hoja.getRange(f, 7, filas.length, 1).setNumberFormat("0.00%");
    hoja.getRange(f, 1, filas.length, 8)
        .setBorder(true, true, true, true, true, true, "#B0B0B0",
                   SpreadsheetApp.BorderStyle.SOLID);
    f += filas.length;
  } else {
    hoja.getRange(f, 1, 1, 8).merge()
        .setValue("Ninguna entidad reportó pagos al exterior en el período.")
        .setFontSize(9).setFontColor("#888888")
        .setHorizontalAlignment("center");
    f++;
  }

  f++;

  hoja.getRange(f, 1, 1, 8).merge().setValue("Total pagos al exterior")
      .setBackground(AZUL_SECCION).setFontWeight("bold");
  f++;

  const inicioTotales = f;

  totalesExt.forEach(function (t) {
    hoja.getRange(f, 1, 1, 6).merge().setValue(t.etiqueta).setFontSize(9);
    hoja.getRange(f, 7).setValue(t.nro_renglon)
        .setFontSize(8).setHorizontalAlignment("center")
        .setBackground(AZUL_CABECERA);
    hoja.getRange(f, 8).setValue(sumas[t.nro_renglon] || 0)
        .setBackground("#EDF3FA").setFontWeight("bold")
        .setNumberFormat("#,##0");
    f++;
  });

  hoja.getRange(inicioTotales, 1, f - inicioTotales, 8)
      .setBorder(true, true, true, true, true, true, "#B0B0B0",
                 SpreadsheetApp.BorderStyle.SOLID);

  [120, 90, 90, 160, 60, 150, 80, 150].forEach(function (ancho, i) {
    hoja.setColumnWidth(i + 1, ancho);
  });
  hoja.setFrozenRows(4);
}


/**
 * Consolidado en dos bloques paralelos, como el formulario 300.
 */
function consolidadoDobleColumna_(libro, renglones, sumas, def, anio, per,
                                  idConsolidado, entradas) {
  const hoja = libro.insertSheet("CONSOLIDADO");
  let f = encabezadoConsolidado_(hoja, def, anio, per, idConsolidado, entradas, 7);

  const ordenados = renglones.slice().sort(function (a, b) {
    return a.nro_renglon - b.nro_renglon;
  });

  const corte = Math.ceil(ordenados.length / 2);
  const izquierda = ordenados.slice(0, corte);
  const derecha   = ordenados.slice(corte);

  hoja.getRange(f, 1, 1, 3).merge().setValue("Concepto")
      .setFontWeight("bold").setBackground(AZUL_CABECERA)
      .setHorizontalAlignment("center");
  hoja.getRange(f, 5, 1, 3).merge().setValue("Concepto")
      .setFontWeight("bold").setBackground(AZUL_CABECERA)
      .setHorizontalAlignment("center");
  f++;

  const inicio = f;
  const filas = Math.max(izquierda.length, derecha.length);

  for (let i = 0; i < filas; i++) {
    bloqueConsolidado_(hoja, f, 1, izquierda[i], sumas);
    bloqueConsolidado_(hoja, f, 5, derecha[i], sumas);
    f++;
  }

  const ultima = f - 1;

  hoja.getRange(inicio, 1, ultima - inicio + 1, 3)
      .setBorder(true, true, true, true, true, true, "#B0B0B0",
                 SpreadsheetApp.BorderStyle.SOLID);
  hoja.getRange(inicio, 5, ultima - inicio + 1, 3)
      .setBorder(true, true, true, true, true, true, "#B0B0B0",
                 SpreadsheetApp.BorderStyle.SOLID);

  [3, 7].forEach(function (col) {
    hoja.getRange(inicio, col, ultima - inicio + 1, 1)
        .setNumberFormat("#,##0").setHorizontalAlignment("right");
  });

  hoja.setColumnWidth(1, 290);
  hoja.setColumnWidth(2, 40);
  hoja.setColumnWidth(3, 140);
  hoja.setColumnWidth(4, 20);
  hoja.setColumnWidth(5, 290);
  hoja.setColumnWidth(6, 40);
  hoja.setColumnWidth(7, 140);
}


/**
 * Escribe un renglón consolidado dentro de un bloque.
 */
function bloqueConsolidado_(hoja, fila, colInicial, renglon, sumas) {
  if (!renglon) return;

  const calculado = String(renglon.editable).toUpperCase() === "NO";

  hoja.getRange(fila, colInicial)
      .setValue(renglon.etiqueta)
      .setWrap(true).setFontSize(9)
      .setFontWeight(calculado ? "bold" : "normal")
      .setBackground(calculado ? "#F2F2F2" : BLANCO);

  hoja.getRange(fila, colInicial + 1)
      .setValue(renglon.nro_renglon)
      .setFontSize(8).setHorizontalAlignment("center")
      .setBackground(AZUL_CABECERA);

  hoja.getRange(fila, colInicial + 2)
      .setValue(sumas[renglon.nro_renglon] || 0)
      .setBackground(calculado ? "#EDF3FA" : BLANCO)
      .setFontWeight(calculado ? "bold" : "normal");
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

/**
 * Muestra el contenido de ENTREGAS para diagnosticar
 * por qué el consolidado no encuentra las cargas.
 */
function diagnosticarEntregas() {
  const datos = SpreadsheetApp.openById(ID_OPERACION)
                  .getSheetByName("ENTREGAS").getDataRange().getValues();

  Logger.log("--- ENTREGAS REGISTRADAS ---");
  Logger.log("Total de filas: " + (datos.length - 1));
  Logger.log("");

  for (let i = 1; i < datos.length; i++) {
    Logger.log(
      "Radicado: " + datos[i][0] +
      " | Formulario: [" + datos[i][2] + "]" +
      " | Entidad: [" + datos[i][3] + "]" +
      " | Año: [" + datos[i][4] + "] (" + typeof datos[i][4] + ")" +
      " | Periodo: [" + datos[i][5] + "] (" + typeof datos[i][5] + ")" +
      " | Estado: [" + datos[i][10] + "]"
    );
  }
}