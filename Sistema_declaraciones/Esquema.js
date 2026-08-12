/**
 * ============================================================
 * ESQUEMA DE DATOS
 * Define las hojas de cada libro y la estructura de carpetas.
 * Este archivo no se ejecuta.
 * ============================================================
 */

// --- Hojas del libro de catálogo ---
const HOJAS_REGISTRO = {
  FORMULARIOS: [
    "cod_formulario", "nombre", "version", "vigencia_desde", "vigencia_hasta",
    "periodicidad", "num_paginas", "estado", "fecha_alta", "usuario_alta"
  ],
  RENGLONES: [
    "cod_formulario", "version", "pagina", "nro_renglon", "etiqueta",
    "tipo_persona", "tipo_valor", "grupo_concepto", "seccion",
    "tipo_dato", "editable", "obligatorio", "orden"
  ],
  REGLAS: [
    "cod_formulario", "version", "id_regla", "tipo_regla", "renglon_destino",
    "expresion", "tolerancia", "severidad", "mensaje", "activa"
  ],
  TABLAS_DETALLE: [
    "cod_formulario", "version", "id_tabla", "nombre_tabla", "pagina",
    "columna", "etiqueta_columna", "tipo_dato", "catalogo_ref",
    "obligatorio", "orden"
  ],
  CATALOGOS: [
    "id_catalogo", "codigo", "descripcion", "vigencia_desde",
    "vigencia_hasta", "activo"
  ],
  ENTIDADES: [
    "cod_entidad", "nombre", "nit", "dv", "cod_direccion_seccional",
    "actividad_economica", "activa"
  ],
  CONFIG: ["clave", "valor", "descripcion"]
};

// --- Hojas del libro transaccional ---
const HOJAS_OPERACION = {
  PLANTILLAS_EMITIDAS: [
    "id_plantilla", "cod_formulario", "version", "cod_entidad", "periodo",
    "anio", "id_archivo_drive", "hash_estructura", "generada_por",
    "fecha_generacion", "estado"
  ],
  ENTREGAS: [
    "radicado", "id_plantilla", "cod_formulario", "cod_entidad", "periodo",
    "anio", "nombre_archivo", "id_archivo_drive", "hash_archivo",
    "cargada_por", "fecha_carga", "estado", "num_errores", "fecha_validacion"
  ],
  DATOS_CARGADOS: [
    "radicado", "cod_formulario", "version", "nro_renglon", "valor",
    "fecha_registro"
  ],
  DATOS_DETALLE: [
    "radicado", "id_tabla", "nro_fila", "columna", "valor", "fecha_registro"
  ],
  LOG_VALIDACION: [
    "radicado", "id_regla", "severidad", "renglon", "valor_esperado",
    "valor_encontrado", "diferencia", "mensaje", "fecha"
  ],
  CONSOLIDADOS: [
    "id_consolidado", "cod_formulario", "periodo", "anio", "radicados_origen",
    "id_archivo_drive", "generado_por", "fecha_generacion", "vigente"
  ]
};

// --- Carpetas de Drive ---
const ESTRUCTURA_CARPETAS = {
  "1_DOCUMENTACION":       { subcarpetas: [], nietos: [] },
  "2_FORMULARIOS_BASE":    { subcarpetas: ["F350"], nietos: ["v2026"] },
  "3_PLANTILLAS_EMITIDAS": { subcarpetas: ["F350"], nietos: [] },
  "4_CARGAS_RECIBIDAS":    { subcarpetas: ["DAVIVIENDA", "DAVIBANK"],
                             nietos: ["aceptadas", "rechazadas"] },
  "5_CONSOLIDADOS":        { subcarpetas: ["F350"], nietos: [] }
};