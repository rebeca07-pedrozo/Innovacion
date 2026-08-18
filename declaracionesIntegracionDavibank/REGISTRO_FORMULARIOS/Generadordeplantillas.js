function generarPlantilla(codFormulario, codEntidad, anio, periodo) {
  const def = definicionFormulario(codFormulario);
  const renglones = leerRenglones_(codFormulario, def.version);

  if (!renglones.length) {
    throw new Error("El catálogo no tiene renglones para " + codFormulario);
  }
  const entidad = leerEntidad_(codEntidad);
  const idPlantilla = generarIdPlantilla_();
  const periodoTxt = formatearPeriodo(periodo);

  const nombre = codFormulario + "_" + codEntidad + "_" + anio + "-" +
                 periodoTxt + "_" + idPlantilla;
  const libro = SpreadsheetApp.create(nombre);

  construirMeta_(libro, idPlantilla, def, codEntidad, anio, periodo);

  if (def.disposicion === "MATRIZ") {
    construirMatriz_(libro, renglones, entidad, anio, periodo, idPlantilla, def);
    construirExterior_(libro, renglones);
  } else {
    construirLista_(libro, renglones, entidad, anio, periodo, idPlantilla, def);
  }

  const porDefecto = libro.getSheetByName("Hoja 1") || libro.getSheetByName("Sheet1");
  if (porDefecto) libro.deleteSheet(porDefecto);

  libro.setActiveSheet(libro.getSheetByName("FORMULARIO"));

  const archivo = DriveApp.getFileById(libro.getId());
  DriveApp.getFolderById(CARPETA_PLANTILLAS).addFile(archivo);
  DriveApp.getRootFolder().removeFile(archivo);

  // Debe ser accesible para quien la descargue desde la interfaz
  archivo.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);

  registrarEmision_(idPlantilla, def, codEntidad, anio, periodoTxt, archivo.getId());

  return { idPlantilla: idPlantilla, idArchivo: archivo.getId(), url: libro.getUrl() };
}
function escribirEncabezado_(hoja, def, entidad, anio, periodo, idPlantilla, ancho) {
  let f = 1;

  hoja.getRange(f, 1, 1, ancho).merge()
      .setValue("PLANTILLA DE TRABAJO - NO CONSTITUYE DECLARACIÓN TRIBUTARIA")
      .setBackground(AMARILLO_AVISO)
      .setFontSize(10).setFontWeight("bold")
      .setHorizontalAlignment("center");
  f += 2;

  hoja.getRange(f, 1, 1, ancho - 1).merge()
      .setValue(def.nombre)
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
      .setValue("Datos del declarante")
      .setBackground(AZUL_SECCION).setFontWeight("bold");
  f++;

  const inicio = f;
  const campos = [
    ["1. Año", anio],
    ["3. Período", periodo],
    ["5. NIT", entidad.nit],
    ["6. DV", entidad.dv],
    ["11. Razón social", entidad.nombre],
    ["12. Cód. dirección seccional", entidad.direccion_seccional],
    ["Identificador de plantilla", idPlantilla]
  ];

  campos.forEach(function (campo) {
    hoja.getRange(f, 1).setValue(campo[0]).setFontWeight("bold");
    hoja.getRange(f, 2, 1, ancho - 1).merge().setValue(campo[1]);
    f++;
  });

  hoja.getRange(inicio, 1, campos.length, ancho)
      .setBackground(GRIS_BLOQUEO).setFontSize(9);

  return f + 1;
}
function construirMatriz_(libro, renglones, entidad, anio, periodo, idPlantilla, def) {
  const hoja = libro.insertSheet("FORMULARIO");
  let f = escribirEncabezado_(hoja, def, entidad, anio, periodo, idPlantilla, 9);

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
        hoja.getRange(f, par[1] + 1).setBackground(BLANCO);
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
    hoja.getRange(f, 9).setBackground(calculado ? GRIS_BLOQUEO : BLANCO);
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
function construirLista_(libro, renglones, entidad, anio, periodo, idPlantilla, def) {
  const hoja = libro.insertSheet("FORMULARIO");
  let f = escribirEncabezado_(hoja, def, entidad, anio, periodo, idPlantilla, 3);

  hoja.getRange(f, 1, 1, 3)
      .setValues([["Renglón", "Concepto", "Valor"]])
      .setFontWeight("bold").setBackground(AZUL_CABECERA)
      .setHorizontalAlignment("center");
  f++;

  const inicio = f;
  const ordenados = renglones.slice().sort(function (a, b) {
    return a.nro_renglon - b.nro_renglon;
  });

  let seccionActual = "";

  ordenados.forEach(function (r) {
    if (r.seccion !== seccionActual) {
      seccionActual = r.seccion;
      hoja.getRange(f, 1, 1, 3).merge()
          .setValue(titularSeccion(seccionActual))
          .setFontWeight("bold").setBackground(AZUL_SECCION).setFontSize(9);
      f++;
    }

    const calculado = String(r.editable).toUpperCase() === "NO";

    hoja.getRange(f, 1).setValue(r.nro_renglon)
        .setFontSize(9).setHorizontalAlignment("center")
        .setBackground(AZUL_CABECERA);
    hoja.getRange(f, 2).setValue(r.etiqueta)
        .setWrap(true).setFontSize(9)
        .setFontWeight(calculado ? "bold" : "normal");
    hoja.getRange(f, 3).setBackground(calculado ? GRIS_BLOQUEO : BLANCO);
    f++;
  });

  const ultima = f - 1;
  hoja.getRange(inicio, 1, ultima - inicio + 1, 3)
      .setBorder(true, true, true, true, true, true, "#B0B0B0",
                 SpreadsheetApp.BorderStyle.SOLID);
  hoja.getRange(inicio, 3, ultima - inicio + 1, 1)
      .setNumberFormat("#,##0").setHorizontalAlignment("right");

  hoja.setColumnWidth(1, 70);
  hoja.setColumnWidth(2, 420);
  hoja.setColumnWidth(3, 160);
}
function construirExterior_(libro, renglones) {
  const totalesExt = renglones.filter(function (r) {
    return r.seccion === "TOTALES_EXTERIOR";
  }).sort(function (a, b) { return a.nro_renglon - b.nro_renglon; });

  if (!totalesExt.length) return;

  const hoja = libro.insertSheet("EXTERIOR");
  const FILAS = 45;

  hoja.getRange(1, 1, 1, 8).merge()
      .setValue("Pagos o abonos en cuenta al exterior - detalle por país")
      .setFontSize(12).setFontWeight("bold")
      .setBackground(AZUL_TITULO).setFontColor(BLANCO)
      .setHorizontalAlignment("center");
  hoja.setRowHeight(1, 28);

  hoja.getRange(3, 1, 1, 8).setValues([[
    "141. A países", "142. Concepto de pago", "143. Tipo de persona",
    "144. País", "Cód.", "145. Pagos o abonos en cuenta",
    "146. Tarifa (%)", "147. Valor retención"
  ]]).setBackground(AZUL_CABECERA).setFontWeight("bold").setFontSize(9)
     .setWrap(true).setHorizontalAlignment("center");
  hoja.setRowHeight(3, 34);

  hoja.getRange(4, 1, FILAS, 8).setBackground(BLANCO);
  hoja.getRange(4, 6, FILAS, 1).setNumberFormat("#,##0");
  hoja.getRange(4, 8, FILAS, 1).setNumberFormat("#,##0");
  hoja.getRange(4, 7, FILAS, 1).setNumberFormat("0.00%");

  hoja.getRange(4, 1, FILAS, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["SIN CONVENIO", "CON CONVENIO"], true).build());
  hoja.getRange(4, 3, FILAS, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["1", "2"], true).build());

  hoja.getRange(4, 1, FILAS, 8)
      .setBorder(true, true, true, true, true, true, "#B0B0B0",
                 SpreadsheetApp.BorderStyle.SOLID);

  let f = 4 + FILAS + 1;
  hoja.getRange(f, 1, 1, 8).merge().setValue("Total pagos al exterior")
      .setBackground(AZUL_SECCION).setFontWeight("bold");
  f++;

  totalesExt.forEach(function (t) {
    hoja.getRange(f, 1, 1, 6).merge().setValue(t.etiqueta).setFontSize(9);
    hoja.getRange(f, 7).setValue(t.nro_renglon)
        .setFontSize(8).setHorizontalAlignment("center")
        .setBackground(AZUL_CABECERA);
    hoja.getRange(f, 8).setBackground(GRIS_BLOQUEO).setNumberFormat("#,##0");
    f++;
  });

  [120, 90, 90, 160, 60, 150, 80, 150].forEach(function (ancho, i) {
    hoja.setColumnWidth(i + 1, ancho);
  });
  hoja.setFrozenRows(3);
}
function agruparEnMatriz_(renglones) {
  const grupos = {};
  const orden = [];

  renglones.forEach(function (r) {
    if (String(r.seccion).indexOf("TOTALES") === 0) return;

    if (!grupos[r.grupo_concepto]) {
      grupos[r.grupo_concepto] = {
        clave: r.grupo_concepto, etiqueta: r.etiqueta, seccion: r.seccion,
        jurBase: null, jurRet: null, natBase: null, natRet: null
      };
      orden.push(r.grupo_concepto);
    }

    const g = grupos[r.grupo_concepto];
    if (r.tipo_persona === "JURIDICA" && r.tipo_valor === "BASE")      g.jurBase = r.nro_renglon;
    if (r.tipo_persona === "JURIDICA" && r.tipo_valor === "RETENCION") g.jurRet  = r.nro_renglon;
    if (r.tipo_persona === "NATURAL"  && r.tipo_valor === "BASE")      g.natBase = r.nro_renglon;
    if (r.tipo_persona === "NATURAL"  && r.tipo_valor === "RETENCION") g.natRet  = r.nro_renglon;
  });

  return orden.map(function (k) { return grupos[k]; })
    .sort(function (a, b) { return numeroDeGrupo_(a.clave) - numeroDeGrupo_(b.clave); });
}
function leerRenglones_(codFormulario, version) {
  const datos = SpreadsheetApp.openById(ID_REGISTRO)
                  .getSheetByName("RENGLONES").getDataRange().getValues();
  const resultado = [];

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] !== codFormulario || datos[i][1] !== version) continue;

    resultado.push({
      pagina: datos[i][2], nro_renglon: datos[i][3], etiqueta: datos[i][4],
      tipo_persona: datos[i][5], tipo_valor: datos[i][6],
      grupo_concepto: datos[i][7], seccion: datos[i][8], editable: datos[i][9]
    });
  }

  return resultado;
}
function leerEntidad_(codEntidad) {
  const datos = SpreadsheetApp.openById(ID_REGISTRO)
                  .getSheetByName("ENTIDADES").getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === codEntidad) {
      return {
        codigo: datos[i][0], nombre: datos[i][1], nit: datos[i][2], dv: datos[i][3],
        direccion_seccional: datos[i][4], actividad: datos[i][5]
      };
    }
  }
  throw new Error("Entidad no encontrada: " + codEntidad);
}
function construirMeta_(libro, idPlantilla, def, codEntidad, anio, periodo) {
  const hoja = libro.insertSheet("_META");

  hoja.getRange(1, 1, 8, 2).setValues([
    ["id_plantilla",     idPlantilla],
    ["cod_formulario",   def.codigo],
    ["version",          def.version],
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


function registrarEmision_(idPlantilla, def, codEntidad, anio, periodoTxt, idArchivo) {
  SpreadsheetApp.openById(ID_OPERACION)
    .getSheetByName("PLANTILLAS_EMITIDAS")
    .appendRow([idPlantilla, def.codigo, def.version, codEntidad, anio, periodoTxt,
                idArchivo, Session.getActiveUser().getEmail(), new Date(), "EMITIDA"]);
}


function generarIdPlantilla_() {
  const hoja = SpreadsheetApp.openById(ID_OPERACION).getSheetByName("PLANTILLAS_EMITIDAS");
  return "PLT-" + new Date().getFullYear() + "-" +
         ("0000" + hoja.getLastRow()).slice(-4);
}


function numeroDeGrupo_(clave) {
  const partes = String(clave).split("_");
  const n = parseInt(partes[partes.length - 1], 10);
  return isNaN(n) ? 999 : n;
}
function compartirPlantillasExistentes() {
  const datos = SpreadsheetApp.openById(ID_OPERACION)
                  .getSheetByName("PLANTILLAS_EMITIDAS").getDataRange().getValues();
  let ajustadas = 0;

  for (let i = 1; i < datos.length; i++) {
    if (!datos[i][6]) continue;

    try {
      DriveApp.getFileById(datos[i][6])
        .setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
      ajustadas++;
    } catch (e) {
      Logger.log("No se pudo ajustar " + datos[i][0] + ": " + e.message);
    }
  }

  Logger.log("Plantillas compartidas: " + ajustadas);
}