const ID_REGISTRO  = "1qXjlhZgsK-jOyGxlf0Vk65WLDK-Oypm6yXbF5SF7H5s";
const ID_OPERACION = "1ghhI0GK8HFI1lP-WMs6Ce2pBHNH_fnSspJ9wIZWectU";
const CARPETA_DOCUMENTACION    = "1_lX1Q2LllzrB-Z3Fjm11uJu11cY698nB";
const CARPETA_FORMULARIOS_BASE = "12vwQBhdo9tM_D1JS_gbfTTRVVB4DYQEW";
const CARPETA_A_PROCESAR       = "19F-wU3TQ_G_je5eZa2hm2PvdbKj5jD05";
const CARPETA_PROCESADOS       = "1ABGxm2C-Nc-rMrqk2gc3UoDTVOywgQNh";
const CARPETA_PLANTILLAS       = "1pfGsUPjYDnNoBLkbuKPBdonb4vr5bR46";
const CARPETA_CARGAS           = "1vdiBlJ-lKwdU23ySuVJSj1ABbsjGGm0y";
const CARPETA_DAVIVIENDA_OK    = "1r6qUk9ZVmFdZ1PvjsqhmVZdsxPS4Tqyj";
const CARPETA_DAVIVIENDA_NO    = "1-mnpjuk4pVEK2efFburkGT0B1L1Twd-0";
const CARPETA_DAVIBANK_OK      = "1YPcq_sPBA3y0VDz0I1CxiJwqM2jXXd0M";
const CARPETA_DAVIBANK_NO      = "1P5oLt10tvRgYQrvdcMT32wMAfbwD6hp7";
const CARPETA_CONSOLIDADOS     = "1omDtwJMma8dnG5gJ6sHoEw0BmA7AnfkW";
const AZUL_TITULO      = "#1F4E79";
const AZUL_CABECERA    = "#D6E3F0";
const AZUL_SECCION     = "#BDD0E4";
const GRIS_BLOQUEO     = "#EDEDED";
const BLANCO           = "#FFFFFF";
const AMARILLO_AVISO   = "#FFF2CC";
const COLOR_ENCABEZADO = "#D9E2F3";
function carpetasDeEntidad(codEntidad) {
  const mapa = {
    DAVIVIENDA: { aceptadas: CARPETA_DAVIVIENDA_OK, rechazadas: CARPETA_DAVIVIENDA_NO },
    DAVIBANK:   { aceptadas: CARPETA_DAVIBANK_OK,   rechazadas: CARPETA_DAVIBANK_NO }
  };

  if (!mapa[codEntidad]) {
    throw new Error("Entidad no reconocida: " + codEntidad);
  }
  return mapa[codEntidad];
}
function definicionFormulario(codFormulario) {
  const datos = SpreadsheetApp.openById(ID_REGISTRO)
                  .getSheetByName("FORMULARIOS").getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] !== codFormulario) continue;
    if (String(datos[i][7]).toUpperCase() !== "ACTIVO") continue;

    return {
      codigo:       datos[i][0],
      nombre:       datos[i][1],
      version:      datos[i][2],
      periodicidad: datos[i][3],
      paginas:      datos[i][4],
      disposicion:  String(datos[i][5]).toUpperCase(),
      esperados:    interpretarRangos(datos[i][6])
    };
  }

  throw new Error("No hay definición activa para el formulario " + codFormulario);
}
function interpretarRangos(especificacion) {
  const lista = [];

  String(especificacion).split(",").forEach(function (parte) {
    const trozo = parte.trim();
    if (!trozo) return;

    if (trozo.indexOf("-") > 0) {
      const limites = trozo.split("-");
      const desde = parseInt(limites[0], 10);
      const hasta = parseInt(limites[1], 10);
      for (let n = desde; n <= hasta; n++) lista.push(n);
    } else {
      lista.push(parseInt(trozo, 10));
    }
  });

  return lista;
}
function titularSeccion(seccion) {
  const texto = String(seccion).replace(/_/g, " ").toLowerCase();
  return texto.charAt(0).toUpperCase() + texto.slice(1);
}
function formatearPeriodo(periodo) {
  return ("0" + periodo).slice(-2);
}
function aNumero(valor) {
  if (typeof valor === "number") return valor;
  if (valor === "" || valor === null || valor === undefined) return 0;

  const n = Number(String(valor).replace(/[.,\s$]/g, ""));
  return isNaN(n) ? 0 : n;
}
function formatearMiles(n) {
  return Number(n).toLocaleString("es-CO");
}
function formatearFecha(fecha) {
  if (!fecha) return "";
  return Utilities.formatDate(new Date(fecha), "America/Bogota", "dd/MM/yyyy HH:mm");
}