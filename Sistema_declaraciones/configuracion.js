/**
 * ============================================================
 * CONFIGURACIÓN GENERAL DEL SISTEMA
 * Constantes globales. Este archivo no se ejecuta.
 * ============================================================
 */

// --- Libros del sistema ---
const ID_REGISTRO  = "1qXjlhZgsK-jOyGxlf0Vk65WLDK-Oypm6yXbF5SF7H5s";
const ID_OPERACION = "1ghhI0GK8HFI1lP-WMs6Ce2pBHNH_fnSspJ9wIZWectU";

// --- Carpeta raíz en Drive ---
const NOMBRE_CARPETA_BASE = "SISTEMA_DECLARACIONES";

// --- Formulario inicial ---
const FORMULARIO_INICIAL = {
  codigo: "F350",
  nombre: "Declaración de retenciones en la fuente",
  version: "v2026",
  periodicidad: "MENSUAL",
  paginas: 2
};

// --- Entidades participantes ---
const ENTIDADES_INICIALES = [
  ["DAVIVIENDA", "BANCO DAVIVIENDA S.A.", "860034313", "7", "31", "6412", "SI"],
  ["DAVIBANK",   "DAVIBANK",              "",          "",  "",   "",     "SI"]
];

// --- Formato visual de encabezados ---
const COLOR_ENCABEZADO = "#D9E2F3";