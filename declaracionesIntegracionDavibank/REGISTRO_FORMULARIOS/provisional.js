/**
 * Comparte con el dominio las plantillas ya emitidas.
 * Ejecutar una sola vez tras el cambio de permisos.
 */
function compartirPlantillasExistentes() {
  const datos = SpreadsheetApp.openById(ID_OPERACION)
                  .getSheetByName("PLANTILLAS_EMITIDAS").getDataRange().getValues();
  let ajustadas = 0;

  for (let i = 1; i < datos.length; i++) {
    const idArchivo = datos[i][6];
    if (!idArchivo) continue;

    try {
      DriveApp.getFileById(idArchivo)
        .setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
      ajustadas++;
    } catch (e) {
      Logger.log("No se pudo ajustar " + datos[i][0] + ": " + e.message);
    }
  }

  Logger.log("Plantillas compartidas: " + ajustadas);
}