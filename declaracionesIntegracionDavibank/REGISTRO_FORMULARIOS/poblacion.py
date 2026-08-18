# BLOQUE 1 — Dependencias
!pip install -q gspread --upgrade

# BLOQUE 2 — Autenticación y configuración

import pandas as pd
import gspread
from google.colab import auth
from google.auth import default

auth.authenticate_user()
credenciales, _ = default()
gc = gspread.authorize(credenciales)

ID_REGISTRO_FORMULARIOS = "1qXjlhZgsK-jOyGxlf0Vk65WLDK-Oypm6yXbF5SF7H5s"
NOMBRE_HOJA_STAGING     = "STAGING_EXTRACCION"

COD_FORMULARIO = "F300"
VERSION_FORM   = "v2026"

libro_registro = gc.open_by_key(ID_REGISTRO_FORMULARIOS)
print("Conectado a:", libro_registro.title)

# BLOQUE 3 — Catálogo de renglones del formulario 300 (IVA)

RENGLONES_F300 = [

    # ---------- Ingresos ----------
    (27, "Por operaciones gravadas al 5%",                             "INGRESOS", "SI"),
    (28, "Por operaciones gravadas a la tarifa general",               "INGRESOS", "SI"),
    (29, "A.I.U por operaciones gravadas (base gravable especial)",    "INGRESOS", "SI"),
    (30, "Por exportación de bienes",                                  "INGRESOS", "SI"),
    (31, "Por exportación de servicios",                               "INGRESOS", "SI"),
    (32, "Por ventas a sociedades de comercialización internacional",  "INGRESOS", "SI"),
    (33, "Por ventas a zonas francas",                                 "INGRESOS", "SI"),
    (34, "Por juegos de suerte y azar",                                "INGRESOS", "SI"),
    (35, "Por operaciones exentas",                                    "INGRESOS", "SI"),
    (36, "Por venta de cerveza de producción nacional o importada",    "INGRESOS", "SI"),
    (37, "Por venta de gaseosas y similares",                          "INGRESOS", "SI"),
    (38, "Por venta de licores, aperitivos, vinos y similares",        "INGRESOS", "SI"),
    (39, "Por operaciones excluidas",                                  "INGRESOS", "SI"),
    (40, "Por operaciones no gravadas",                                "INGRESOS", "SI"),
    (41, "Total ingresos brutos",                                      "INGRESOS", "NO"),
    (42, "Devoluciones en ventas anuladas, rescindidas o resueltas",   "INGRESOS", "SI"),
    (43, "Total ingresos netos recibidos durante el período",          "INGRESOS", "NO"),

    # ---------- Compras: importación ----------
    (44, "De bienes gravados a la tarifa del 5%",                      "COMPRAS_IMPORTACION", "SI"),
    (45, "De bienes gravados a la tarifa general",                     "COMPRAS_IMPORTACION", "SI"),
    (46, "De bienes y servicios gravados provenientes de Zonas Francas",
                                                                       "COMPRAS_IMPORTACION", "SI"),
    (47, "De bienes no gravados",                                      "COMPRAS_IMPORTACION", "SI"),
    (48, "De bienes excluidos, exentos y no gravados provenientes de Zonas Francas",
                                                                       "COMPRAS_IMPORTACION", "SI"),
    (49, "De servicios",                                               "COMPRAS_IMPORTACION", "SI"),

    # ---------- Compras: nacionales ----------
    (50, "De bienes gravados a la tarifa del 5%",                      "COMPRAS_NACIONALES", "SI"),
    (51, "De bienes gravados a la tarifa general",                     "COMPRAS_NACIONALES", "SI"),
    (52, "De servicios gravados a la tarifa del 5%",                   "COMPRAS_NACIONALES", "SI"),
    (53, "De servicios gravados a la tarifa general",                  "COMPRAS_NACIONALES", "SI"),
    (54, "De bienes y servicios excluidos, exentos y no gravados",     "COMPRAS_NACIONALES", "SI"),
    (55, "Total compras e importaciones brutas",                       "COMPRAS_NACIONALES", "NO"),
    (56, "Devoluciones en compras anuladas, rescindidas o resueltas en este período",
                                                                       "COMPRAS_NACIONALES", "SI"),
    (57, "Total compras netas realizadas durante el período",          "COMPRAS_NACIONALES", "NO"),

    # ---------- Liquidación privada: impuesto generado ----------
    (58, "A la tarifa del 5%",                                         "IMPUESTO_GENERADO", "SI"),
    (59, "A la tarifa general",                                        "IMPUESTO_GENERADO", "SI"),
    (60, "Sobre A.I.U en operaciones gravadas (base gravable especial)",
                                                                       "IMPUESTO_GENERADO", "SI"),
    (61, "En juegos de suerte y azar",                                 "IMPUESTO_GENERADO", "SI"),
    (62, "En venta cerveza de producción nacional o importada",        "IMPUESTO_GENERADO", "SI"),
    (63, "En venta de gaseosas y similares",                           "IMPUESTO_GENERADO", "SI"),
    (64, "En venta de licores, aperitivos, vinos y similares 5%",      "IMPUESTO_GENERADO", "SI"),
    (65, "En retiro de inventario para activos fijos, consumo, muestras gratis o donaciones",
                                                                       "IMPUESTO_GENERADO", "SI"),
    (66, "IVA recuperado en devoluciones en compras anuladas, rescindidas o resueltas",
                                                                       "IMPUESTO_GENERADO", "SI"),
    (67, "Total impuesto generado por operaciones gravadas",           "IMPUESTO_GENERADO", "NO"),

    # ---------- Liquidación privada: impuesto descontable ----------
    (68, "Por importaciones gravadas a tarifa del 5%",                 "IMPUESTO_DESCONTABLE", "SI"),
    (69, "Por importaciones gravadas la tarifa general",               "IMPUESTO_DESCONTABLE", "SI"),
    (70, "De bienes y servicios gravados provenientes de Zonas Francas",
                                                                       "IMPUESTO_DESCONTABLE", "SI"),
    (71, "Por compras de bienes gravados a la tarifa 5%",              "IMPUESTO_DESCONTABLE", "SI"),
    (72, "Por compras de bienes gravados a tarifa general",            "IMPUESTO_DESCONTABLE", "SI"),
    (73, "Por licores, aperitivos, vinos y similares",                 "IMPUESTO_DESCONTABLE", "SI"),
    (74, "Por servicios gravados a la tarifa del 5%",                  "IMPUESTO_DESCONTABLE", "SI"),
    (75, "Por servicios gravados a la tarifa general",                 "IMPUESTO_DESCONTABLE", "SI"),
    (76, "Descuento IVA exploración hidrocarburos Art. 485-2 ET",      "IMPUESTO_DESCONTABLE", "SI"),
    (77, "Total Impuesto pagado o facturado",                          "IMPUESTO_DESCONTABLE", "NO"),
    (78, "IVA retenido por servicios prestados en Colombia por no domiciliados o no residentes",
                                                                       "IMPUESTO_DESCONTABLE", "SI"),
    (79, "IVA resultante por devoluciones en ventas anuladas, rescindidas o resueltas",
                                                                       "IMPUESTO_DESCONTABLE", "SI"),
    (80, "Ajuste impuestos descontables (pérdidas, hurto o castigo de inventarios)",
                                                                       "IMPUESTO_DESCONTABLE", "SI"),
    (81, "Total impuestos descontables",                               "IMPUESTO_DESCONTABLE", "NO"),

    # ---------- Saldos ----------
    (82, "Saldo a pagar por el período fiscal",                        "SALDOS", "NO"),
    (83, "Saldo a favor del período fiscal",                           "SALDOS", "NO"),
    (84, "Saldo a favor del período fiscal anterior",                  "SALDOS", "SI"),
    (85, "Retenciones por IVA que le practicaron",                     "SALDOS", "SI"),
    (86, "Saldo a pagar por impuesto",                                 "SALDOS", "NO"),
    (87, "Sanciones",                                                  "SALDOS", "SI"),
    (88, "Total saldo a pagar",                                        "SALDOS", "NO"),
    (89, "o Total saldo a favor",                                      "SALDOS", "NO"),

    # ---------- Control de saldos ----------
    (90, "Saldo a favor susceptible de devolución y/o compensación por el presente período",
                                                                       "CONTROL_SALDOS", "SI"),
    (91, "Saldo a favor susceptible de ser devuelto y/o compensado a imputar en el período siguiente",
                                                                       "CONTROL_SALDOS", "SI"),
    (92, "Saldo a favor sin derecho a devolución y/o compensación susceptible de ser imputado en el período siguiente",
                                                                       "CONTROL_SALDOS", "SI"),
    (93, "Total saldo a favor a imputar al período siguiente",         "CONTROL_SALDOS", "NO"),

    # ---------- Anticipos IVA Régimen SIMPLE ----------
    (1,   "Anticipo IVA Régimen SIMPLE - bimestre 1",                  "ANTICIPOS_SIMPLE", "SI"),
    (2,   "Anticipo IVA Régimen SIMPLE - bimestre 2",                  "ANTICIPOS_SIMPLE", "SI"),
    (3,   "Anticipo IVA Régimen SIMPLE - bimestre 3",                  "ANTICIPOS_SIMPLE", "SI"),
    (4,   "Anticipo IVA Régimen SIMPLE - bimestre 4",                  "ANTICIPOS_SIMPLE", "SI"),
    (5,   "Anticipo IVA Régimen SIMPLE - bimestre 5",                  "ANTICIPOS_SIMPLE", "SI"),
    (6,   "Anticipo IVA Régimen SIMPLE - bimestre 6",                  "ANTICIPOS_SIMPLE", "SI"),
    (100, "Total anticipos IVA Régimen SIMPLE",                        "ANTICIPOS_SIMPLE", "NO"),
]

print(f"Renglones definidos: {len(RENGLONES_F300)}")

# BLOQUE 4 — Dataset para staging

registros = []

for nro, etiqueta, seccion, editable in RENGLONES_F300:
    registros.append({
        "cod_formulario": COD_FORMULARIO,
        "version":        VERSION_FORM,
        "pagina":         1,
        "nro_renglon":    nro,
        "etiqueta":       etiqueta,
        "tipo_persona":   "",
        "tipo_valor":     "VALOR",
        "grupo_concepto": seccion,
        "seccion":        seccion,
        "editable":       editable,
    })

df_catalogo = pd.DataFrame(registros).sort_values("nro_renglon").reset_index(drop=True)
print(f"Renglones generados: {len(df_catalogo)}")

# BLOQUE 5 — Control de completitud

ESPERADOS = set(range(1, 7)) | set(range(27, 94)) | {100}
en_catalogo = set(df_catalogo["nro_renglon"])

faltantes  = sorted(ESPERADOS - en_catalogo)
sobrantes  = sorted(en_catalogo - ESPERADOS)
duplicados = df_catalogo[df_catalogo.duplicated("nro_renglon")]["nro_renglon"].tolist()

print(f"Renglones: {len(en_catalogo)} de {len(ESPERADOS)} esperados")
print(f"Faltantes:  {faltantes}")
print(f"Sobrantes:  {sobrantes}")
print(f"Duplicados: {duplicados}")

if not faltantes and not sobrantes and not duplicados:
    print("\nCatálogo íntegro.")


# BLOQUE 6 — Envío a STAGING_EXTRACCION

COLUMNAS = ["cod_formulario", "version", "pagina", "nro_renglon", "etiqueta",
            "tipo_persona", "tipo_valor", "grupo_concepto", "seccion", "editable"]

df_salida = df_catalogo[COLUMNAS]

try:
    hoja_staging = libro_registro.worksheet(NOMBRE_HOJA_STAGING)
except gspread.WorksheetNotFound:
    hoja_staging = libro_registro.add_worksheet(
        title=NOMBRE_HOJA_STAGING, rows=400, cols=len(COLUMNAS)
    )

datos = df_salida.fillna("").astype(str).values.tolist()
hoja_staging.clear()
hoja_staging.update(range_name="A1", values=[COLUMNAS] + datos)

print(f"{len(datos)} renglones escritos en {NOMBRE_HOJA_STAGING}")