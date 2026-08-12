/**
 * ============================================================
 * UTILIDADES
 * Funciones auxiliares reutilizables. Este archivo no se ejecuta.
 * ============================================================
 */

/**
 * Devuelve una carpeta existente o la crea si no está.
 */
function obtenerOCrearCarpeta(padre, nombre) {
  const existentes = padre.getFoldersByName(nombre);
  return existentes.hasNext() ? existentes.next() : padre.createFolder(nombre);
}

/**
 * Devuelve una hoja existente o la crea con sus encabezados.
 */
function obtenerOCrearHoja(libro, nombreHoja, encabezados) {
  let hoja = libro.getSheetByName(nombreHoja);

  if (!hoja) {
    hoja = libro.insertSheet(nombreHoja);
  }

  if (hoja.getLastRow() === 0 && encabezados) {
    hoja.getRange(1, 1, 1, encabezados.length)
        .setValues([encabezados])
        .setFontWeight("bold")
        .setBackground(COLOR_ENCABEZADO);
    hoja.setFrozenRows(1);
  }

  return hoja;
}

/**
 * Elimina la hoja por defecto si quedó vacía y hay otras hojas.
 */
function eliminarHojaPorDefecto(libro) {
  const nombres = ["Hoja 1", "Hoja1", "Sheet1"];

  nombres.forEach(function (nombre) {
    const hoja = libro.getSheetByName(nombre);
    if (hoja && libro.getSheets().length > 1 && hoja.getLastRow() === 0) {
      libro.deleteSheet(hoja);
    }
  });
}

/**
 * Lee un valor de la hoja CONFIG por su clave.
 */
function leerConfig(clave) {
  const hoja = SpreadsheetApp.openById(ID_REGISTRO).getSheetByName("CONFIG");
  const datos = hoja.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === clave) {
      return datos[i][1];
    }
  }
  return null;
}

/**
 * Devuelve el correo del usuario que ejecuta la acción.
 */
function usuarioActual() {
  return Session.getActiveUser().getEmail() || "desconocido";
}