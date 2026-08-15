const AZUL_TITULO   = "#1F4E79";
const AZUL_CABECERA = "#D6E3F0";
const AZUL_SECCION  = "#BDD0E4";
const GRIS_BLOQUEO  = "#EDEDED";
const BLANCO        = "#FFFFFF";
const AMARILLO_AVISO = "#FFF2CC";

const PARAMETROS_PLANTILLA = {
  codFormulario: "F350",
  version: "v2026",
  codEntidad: "DAVIVIENDA",
  anio: 2026,
  periodo: 8
};
//Plantilla manual para pruebas de integración
function generarPlantillaManual() {
  const p = PARAMETROS_PLANTILLA;
  const r = generarPlantilla(p.codFormulario, p.version, p.codEntidad, p.anio, p.periodo);

  Logger.log("Plantilla: " + r.idPlantilla);
  Logger.log("Enlace: " + r.url);
}
function generarPlantilla(codFormulario, version, codEntidad, anio, periodo) {
  const renglones = leerRenglones_(codFormulario, version);
  if (!renglones.length) {
    throw new Error("El catálogo no tiene renglones para " + codFormulario + " " + version);
  }

  const entidad = leerEntidad_(codEntidad);
  const idPlantilla = generarIdPlantilla_();
  const periodoTxt = formatearPeriodo_(periodo);

  const nombre = codFormulario + "_" + codEntidad + "_" + anio + "-" +
                 periodoTxt + "_" + idPlantilla;

  const libro = SpreadsheetApp.create(nombre);

  construirMeta_(libro, idPlantilla, codFormulario, version, codEntidad, anio, periodo);
  construirFormulario_(libro, renglones, entidad, anio, periodo, idPlantilla);
  construirExterior_(libro, renglones);

  const porDefecto = libro.getSheetByName("Hoja 1") || libro.getSheetByName("Sheet1");
  if (porDefecto) libro.deleteSheet(porDefecto);

  libro.setActiveSheet(libro.getSheetByName("FORMULARIO"));

  const archivo = DriveApp.getFileById(libro.getId());
  DriveApp.getFolderById(CARPETA_PLANTILLAS).addFile(archivo);
  DriveApp.getRootFolder().removeFile(archivo);

  registrarEmision_(idPlantilla, codFormulario, version, codEntidad,
                    anio, periodoTxt, archivo.getId());

  return { idPlantilla: idPlantilla, idArchivo: archivo.getId(), url: libro.getUrl() };
}
function construirFormulario_(libro, renglones, entidad, anio, periodo, idPlantilla) {
  const hoja = libro.insertSheet("FORMULARIO");
  const matriz = agruparEnMatriz_(renglones);
  const totales = filtrarTotales_(renglones);

  let f = 1;

  hoja.getRange(f, 1, 1, 9).merge()
      .setValue("PLANTILLA DE TRABAJO - NO CONSTITUYE DECLARACIÓN TRIBUTARIA")
      .setBackground(AMARILLO_AVISO)
      .setFontSize(10)
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
  f += 2;

  hoja.getRange(f, 1, 1, 7).merge()
      .setValue("Declaración retenciones en la fuente")
      .setFontSize(13).setFontWeight("bold")
      .setBackground(AZUL_TITULO).setFontColor(BLANCO)
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
  hoja.getRange(f, 8, 1, 2).merge()
      .setValue("350")
      .setFontSize(22).setFontWeight("bold")
      .setBackground(AZUL_TITULO).setFontColor(BLANCO)
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
  hoja.setRowHeight(f, 34);
  f += 2;

  hoja.getRange(f, 1, 1, 9).merge()
      .setValue("Datos del declarante")
      .setBackground(AZUL_SECCION).setFontWeight("bold");
  f++;

  const filaDeclarante = f;
  const declarante = [
    ["1. Año", anio, "3. Período", periodo, "4. Identificador de plantilla", idPlantilla],
    ["5. NIT", entidad.nit, "6. DV", entidad.dv, "11. Razón social", entidad.nombre],
    ["12. Cód. dirección seccional", entidad.direccion_seccional,
     "27. Actividad económica", entidad.actividad, "", ""]
  ];

  declarante.forEach(function (fila) {
    hoja.getRange(f, 1, 1, 2).merge().setValue(fila[0]).setFontWeight("bold");
    hoja.getRange(f, 3).setValue(fila[1]);
    hoja.getRange(f, 4, 1, 2).merge().setValue(fila[2]).setFontWeight("bold");
    hoja.getRange(f, 6).setValue(fila[3]);
    hoja.getRange(f, 7).setValue(fila[4]).setFontWeight("bold");
    hoja.getRange(f, 8, 1, 2).merge().setValue(fila[5]);
    f++;
  });

  hoja.getRange(filaDeclarante, 1, 3, 9).setBackground(GRIS_BLOQUEO).setFontSize(9);
  f++;

  hoja.getRange(f, 1).setValue("Concepto")
      .setFontWeight("bold").setBackground(AZUL_CABECERA)
      .setVerticalAlignment("middle");
  hoja.getRange(f, 1, 2, 1).merge();

  hoja.getRange(f, 2, 1, 4).merge().setValue("A personas jurídicas")
      .setFontWeight("bold").setBackground(AZUL_CABECERA)
      .setHorizontalAlignment("center");
  hoja.getRange(f, 6, 1, 4).merge().setValue("A personas naturales")
      .setFontWeight("bold").setBackground(AZUL_CABECERA)
      .setHorizontalAlignment("center");
  f++;

  const subcabecera = [
    "", "Base sujeta a retención", "", "Retenciones a título de renta", "",
    "Base sujeta a retención", "", "Retenciones a título de renta", ""
  ];
  hoja.getRange(f, 1, 1, 9).setValues([subcabecera])
      .setBackground(AZUL_CABECERA).setFontSize(9)
      .setHorizontalAlignment("center").setWrap(true);
  hoja.getRange(f, 2, 1, 2).merge();
  hoja.getRange(f, 4, 1, 2).merge();
  hoja.getRange(f, 6, 1, 2).merge();
  hoja.getRange(f, 8, 1, 2).merge();
  hoja.setRowHeight(f, 30);
  f++;

  const inicioMatriz = f;
  const filasValor = [];
  let seccionActual = "";

  matriz.forEach(function (fila) {
    if (fila.seccion !== seccionActual) {
      seccionActual = fila.seccion;
      hoja.getRange(f, 1, 1, 9).merge()
          .setValue(titularSeccion_(seccionActual))
          .setFontWeight("bold").setBackground(AZUL_SECCION).setFontSize(9);
      f++;
    }

    hoja.getRange(f, 1).setValue(fila.etiqueta).setWrap(true).setFontSize(9);

    // Cada par renglón/valor ocupa dos columnas contiguas
    [[fila.jurBase, 2], [fila.jurRet, 4], [fila.natBase, 6], [fila.natRet, 8]]
      .forEach(function (par) {
        const nro = par[0];
        const col = par[1];

        if (nro === null) {
          hoja.getRange(f, col, 1, 2).setBackground(GRIS_BLOQUEO);
          return;
        }

        hoja.getRange(f, col).setValue(nro)
            .setFontSize(8).setHorizontalAlignment("center")
            .setBackground(AZUL_CABECERA);
        hoja.getRange(f, col + 1).setBackground(BLANCO);
        filasValor.push({ fila: f, col: col + 1 });
      });

    f++;
  });

  totales.forEach(function (t) {
    const esCalculado = String(t.editable).toUpperCase() === "NO";

    hoja.getRange(f, 1, 1, 7).merge()
        .setValue(t.etiqueta).setWrap(true).setFontSize(9)
        .setFontWeight(esCalculado ? "bold" : "normal");
    hoja.getRange(f, 8).setValue(t.nro_renglon)
        .setFontSize(8).setHorizontalAlignment("center")
        .setBackground(AZUL_CABECERA);
    hoja.getRange(f, 9).setBackground(esCalculado ? GRIS_BLOQUEO : BLANCO);

    filasValor.push({ fila: f, col: 9 });
    f++;
  });

  const ultima = f - 1;
  hoja.getRange(inicioMatriz, 1, ultima - inicioMatriz + 1, 9)
      .setBorder(true, true, true, true, true, true, "#B0B0B0",
                 SpreadsheetApp.BorderStyle.SOLID);

  [3, 5, 7, 9].forEach(function (col) {
    hoja.getRange(inicioMatriz, col, ultima - inicioMatriz + 1, 1)
        .setNumberFormat("#,##0").setHorizontalAlignment("right");
  });

  hoja.setColumnWidth(1, 300);
  [2, 4, 6, 8].forEach(function (c) { hoja.setColumnWidth(c, 32); });
  [3, 5, 7, 9].forEach(function (c) { hoja.setColumnWidth(c, 130); });

}
function construirExterior_(libro, renglones) {
  const hoja = libro.insertSheet("EXTERIOR");
  const FILAS_DETALLE = 45;

  hoja.getRange(1, 1, 1, 7).merge()
      .setValue("Pagos o abonos en cuenta al exterior - detalle por país")
      .setFontSize(12).setFontWeight("bold")
      .setBackground(AZUL_TITULO).setFontColor(BLANCO)
      .setHorizontalAlignment("center");
  hoja.setRowHeight(1, 28);

  const cabecera = ["141. A países", "142. Concepto de pago", "143. Tipo de persona",
                    "144. País", "Cód.", "145. Pagos o abonos en cuenta",
                    "146. Tarifa (%)", "147. Valor retención"];

  hoja.getRange(3, 1, 1, cabecera.length).setValues([cabecera])
      .setBackground(AZUL_CABECERA).setFontWeight("bold").setFontSize(9)
      .setWrap(true).setHorizontalAlignment("center");
  hoja.setRowHeight(3, 34);
  const numeros = [];
  for (let i = 1; i <= FILAS_DETALLE; i++) numeros.push([i]);

  hoja.getRange(4, 1, FILAS_DETALLE, 8).setBackground(BLANCO);
  hoja.getRange(4, 6, FILAS_DETALLE, 1).setNumberFormat("#,##0");
  hoja.getRange(4, 8, FILAS_DETALLE, 1).setNumberFormat("#,##0");
  hoja.getRange(4, 7, FILAS_DETALLE, 1).setNumberFormat("0.00%");

  const convenio = SpreadsheetApp.newDataValidation()
    .requireValueInList(["SIN CONVENIO", "CON CONVENIO"], true).build();
  hoja.getRange(4, 1, FILAS_DETALLE, 1).setDataValidation(convenio);

  const tipoPersona = SpreadsheetApp.newDataValidation()
    .requireValueInList(["1", "2"], true).build();
  hoja.getRange(4, 3, FILAS_DETALLE, 1).setDataValidation(tipoPersona);

  hoja.getRange(4, 1, FILAS_DETALLE, 8)
      .setBorder(true, true, true, true, true, true, "#B0B0B0",
                 SpreadsheetApp.BorderStyle.SOLID);
  let f = 4 + FILAS_DETALLE + 1;

  hoja.getRange(f, 1, 1, 8).merge()
      .setValue("Total pagos al exterior")
      .setBackground(AZUL_SECCION).setFontWeight("bold");
  f++;

  const totalesExt = renglones
    .filter(function (r) { return r.seccion === "TOTALES_EXTERIOR"; })
    .sort(function (a, b) { return a.nro_renglon - b.nro_renglon; });

  totalesExt.forEach(function (t) {
    hoja.getRange(f, 1, 1, 6).merge().setValue(t.etiqueta).setFontSize(9);
    hoja.getRange(f, 7).setValue(t.nro_renglon)
        .setFontSize(8).setHorizontalAlignment("center")
        .setBackground(AZUL_CABECERA);
    hoja.getRange(f, 8).setBackground(GRIS_BLOQUEO).setNumberFormat("#,##0");
    f++;
  });

  hoja.getRange(4 + FILAS_DETALLE + 1, 1, f - (4 + FILAS_DETALLE + 1), 8)
      .setBorder(true, true, true, true, true, true, "#B0B0B0",
                 SpreadsheetApp.BorderStyle.SOLID);

  hoja.setColumnWidth(1, 120);
  hoja.setColumnWidth(2, 90);
  hoja.setColumnWidth(3, 90);
  hoja.setColumnWidth(4, 160);
  hoja.setColumnWidth(5, 60);
  hoja.setColumnWidth(6, 150);
  hoja.setColumnWidth(7, 80);
  hoja.setColumnWidth(8, 150);
  hoja.setFrozenRows(3);
}
function agruparEnMatriz_(renglones) {
  const grupos = {};
  const orden = [];

  renglones.forEach(function (r) {
    if (["CONCEPTOS", "EXTERIOR", "AUTORRETENCIONES"].indexOf(r.seccion) < 0) return;

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
    .sort(function (a, b) {
      return numeroDeGrupo_(a.clave) - numeroDeGrupo_(b.clave);
    });
}
function filtrarTotales_(renglones) {
  return renglones
    .filter(function (r) { return String(r.seccion).indexOf("TOTALES") === 0 &&
                                  r.seccion !== "TOTALES_EXTERIOR"; })
    .sort(function (a, b) { return a.nro_renglon - b.nro_renglon; });
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
function construirMeta_(libro, idPlantilla, codFormulario, version, codEntidad, anio, periodo) {
  const hoja = libro.insertSheet("_META");

  hoja.getRange(1, 1, 8, 2).setValues([
    ["id_plantilla",     idPlantilla],
    ["cod_formulario",   codFormulario],
    ["version",          version],
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

function registrarEmision_(idPlantilla, codFormulario, version, codEntidad,
                           anio, periodoTxt, idArchivo) {
  SpreadsheetApp.openById(ID_OPERACION)
    .getSheetByName("PLANTILLAS_EMITIDAS")
    .appendRow([idPlantilla, codFormulario, version, codEntidad, anio, periodoTxt,
                idArchivo, Session.getActiveUser().getEmail(), new Date(), "EMITIDA"]);
}
function generarIdPlantilla_() {
  const hoja = SpreadsheetApp.openById(ID_OPERACION).getSheetByName("PLANTILLAS_EMITIDAS");
  return "PLT-" + new Date().getFullYear() + "-" +
         ("0000" + hoja.getLastRow()).slice(-4);
}

function formatearPeriodo_(periodo) {
  return ("0" + periodo).slice(-2);
}

function numeroDeGrupo_(clave) {
  const partes = String(clave).split("_");
  const n = parseInt(partes[partes.length - 1], 10);
  return isNaN(n) ? 999 : n;
}

function titularSeccion_(seccion) {
  const titulos = {
    CONCEPTOS: "Conceptos sujetos a retención",
    EXTERIOR: "Pagos o abonos en cuenta al exterior",
    AUTORRETENCIONES: "Autorretenciones"
  };
  return titulos[seccion] || seccion;
}