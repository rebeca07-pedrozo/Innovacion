const NOMBRE_SALIDA_BASE64 = "IMAGENES_BASE64.txt";
const TAMANO_MAXIMO_KB = 200;
function convertirImagenesABase64() {
  const carpeta = DriveApp.getFolderById(CARPETA_DOCUMENTACION);
  const imagenes = listarImagenes_(carpeta);

  if (!imagenes.length) {
    Logger.log("No se encontraron imágenes en 1_DOCUMENTACION.");
    Logger.log("Formatos admitidos: png, jpg, jpeg, gif.");
    return;
  }

  const bloques = [];

  imagenes.forEach(function (archivo) {
    const nombre = archivo.getName();
    const tamanoKB = Math.round(archivo.getSize() / 1024);

    if (tamanoKB > TAMANO_MAXIMO_KB) {
      Logger.log("OMITIDA: " + nombre + " pesa " + tamanoKB +
                 " KB. El máximo recomendado es " + TAMANO_MAXIMO_KB + " KB.");
      return;
    }

    const blob = archivo.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());
    const dataUri = "data:" + blob.getContentType() + ";base64," + base64;

    bloques.push(
      "============================================================\n" +
      "ARCHIVO:   " + nombre + "\n" +
      "TIPO:      " + blob.getContentType() + "\n" +
      "ORIGINAL:  " + tamanoKB + " KB\n" +
      "BASE64:    " + Math.round(dataUri.length / 1024) + " KB\n" +
      "============================================================\n\n" +
      "USO EN HTML:\n" +
      '<img src="' + "PEGAR_AQUI_LA_CADENA" + '" alt="' + nombre + '">\n\n' +
      "CADENA:\n" +
      dataUri + "\n\n"
    );

    Logger.log("Convertida: " + nombre + " (" + tamanoKB + " KB)");
  });

  if (!bloques.length) {
    Logger.log("Ninguna imagen pudo convertirse.");
    return;
  }

  const contenido =
    "IMÁGENES EN BASE64 - SISTEMA DE DECLARACIONES\n" +
    "Generado: " + new Date().toLocaleString("es-CO") + "\n" +
    "Imágenes: " + bloques.length + "\n\n" +
    bloques.join("\n");

  const idArchivo = guardarSalida_(carpeta, contenido);

  Logger.log("\nArchivo generado: " + NOMBRE_SALIDA_BASE64);
  Logger.log("Enlace: https://drive.google.com/file/d/" + idArchivo + "/view");
}
function listarImagenes_(carpeta) {
  const tiposValidos = ["image/png", "image/jpeg", "image/gif"];
  const encontradas = [];
  const archivos = carpeta.getFiles();

  while (archivos.hasNext()) {
    const archivo = archivos.next();
    if (tiposValidos.indexOf(archivo.getMimeType()) >= 0) {
      encontradas.push(archivo);
    }
  }

  return encontradas;
}
function guardarSalida_(carpeta, contenido) {
  const existentes = carpeta.getFilesByName(NOMBRE_SALIDA_BASE64);

  while (existentes.hasNext()) {
    existentes.next().setTrashed(true);
  }

  const archivo = carpeta.createFile(
    NOMBRE_SALIDA_BASE64, contenido, MimeType.PLAIN_TEXT
  );

  return archivo.getId();
}