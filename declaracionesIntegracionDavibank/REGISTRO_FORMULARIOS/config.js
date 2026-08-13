const ID_REGISTRO  = "1qXjlhZgsK-jOyGxlf0Vk65WLDK-Oypm6yXbF5SF7H5s";
const ID_OPERACION = "1ghhI0GK8HFI1lP-WMs6Ce2pBHNH_fnSspJ9wIZWectU";
const CARPETA_DOCUMENTACION      = "1_lX1Q2LllzrB-Z3Fjm11uJu11cY698nB";
const CARPETA_FORMULARIOS_BASE   = "12vwQBhdo9tM_D1JS_gbfTTRVVB4DYQEW";
const CARPETA_A_PROCESAR         = "19F-wU3TQ_G_je5eZa2hm2PvdbKj5jD05";
const CARPETA_PROCESADOS         = "1ABGxm2C-Nc-rMrqk2gc3UoDTVOywgQNh";
const CARPETA_PLANTILLAS         = "1pfGsUPjYDnNoBLkbuKPBdonb4vr5bR46";
const CARPETA_CARGAS             = "1vdiBlJ-lKwdU23ySuVJSj1ABbsjGGm0y";
const CARPETA_DAVIVIENDA         = "1AbKO7g4cCK5Se2PDQQSAcJI5QCOMJo5g";
const CARPETA_DAVIVIENDA_OK      = "1r6qUk9ZVmFdZ1PvjsqhmVZdsxPS4Tqyj";
const CARPETA_DAVIVIENDA_NO      = "1-mnpjuk4pVEK2efFburkGT0B1L1Twd-0";
const CARPETA_DAVIBANK           = "1RqrIp_7EOfl0REOUl9_HM5ucHw5WriHB";
const CARPETA_DAVIBANK_OK        = "1YPcq_sPBA3y0VDz0I1CxiJwqM2jXXd0M";
const CARPETA_DAVIBANK_NO        = "1P5oLt10tvRgYQrvdcMT32wMAfbwD6hp7";
const CARPETA_CONSOLIDADOS       = "1omDtwJMma8dnG5gJ6sHoEw0BmA7AnfkW";
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
//Davibank - Config