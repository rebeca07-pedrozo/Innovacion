/**
 * ============================================================
 * DIAGNÓSTICO
 * Verifica que la instalación quedó completa.
 * Ejecutar: verificarInstalacion()
 * ============================================================
 */

function verificarInstalacion() {
  const reporte = [];

  // --- Verificación de hojas ---
  const libroRegistro  = SpreadsheetApp.openById(ID_REGISTRO);
  const libroOperacion = SpreadsheetApp.openById(ID_OPERACION);

  reporte.push("--- HOJAS EN REGISTRO_FORMULARIOS ---");
  Object.keys(HOJAS_REGISTRO).forEach(function (nombre) {
    const hoja = libroRegistro.getSheetByName(nombre);
    reporte.push(hoja ? "OK  " + nombre : "FALTA  " + nombre);
  });

  reporte.push("--- HOJAS EN OPERACION_DECLARACIONES ---");
  Object.keys(HOJAS_OPERACION).forEach(function (nombre) {
    const hoja = libroOperacion.getSheetByName(nombre);
    reporte.push(hoja ? "OK  " + nombre : "FALTA  " + nombre);
  });

  // --- Verificación de carpetas ---
  reporte.push("--- CARPETAS ---");
  const hojaConfig = libroRegistro.getSheetByName("CONFIG");
  const config = hojaConfig.getDataRange().getValues();

  for (let i = 1; i < config.length; i++) {
    if (String(config[i][0]).indexOf("CARPETA_") !== 0) continue;

    try {
      const carpeta = DriveApp.getFolderById(config[i][1]);
      reporte.push("OK  " + carpeta.getName());
    } catch (e) {
      reporte.push("ERROR  " + config[i][0] + " (ID inválido)");
    }
  }

  // --- Datos sembrados ---
  reporte.push("--- DATOS INICIALES ---");
  reporte.push("Formularios registrados: " +
    Math.max(0, libroRegistro.getSheetByName("FORMULARIOS").getLastRow() - 1));
  reporte.push("Entidades registradas: " +
    Math.max(0, libroRegistro.getSheetByName("ENTIDADES").getLastRow() - 1));
  reporte.push("Renglones en staging: " +
    (libroRegistro.getSheetByName("STAGING_EXTRACCION")
      ? libroRegistro.getSheetByName("STAGING_EXTRACCION").getLastRow() - 1
      : "hoja no encontrada"));

  Logger.log(reporte.join("\n"));
}