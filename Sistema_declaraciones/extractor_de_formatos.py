# ============================================================
# BLOQUE 1 — Instalación de dependencias
# ============================================================
!pip install -q pdfplumber gspread --upgrade

# ============================================================
# BLOQUE 2 — Imports, autenticación y configuración
# ============================================================
import os
import re
import pandas as pd
import pdfplumber
import gspread
from google.colab import auth, drive
from google.auth import default

auth.authenticate_user()
credenciales, _ = default()
gc = gspread.authorize(credenciales)

# --- Identificadores del sistema ---
ID_REGISTRO_FORMULARIOS = "1qXjlhZgsK-jOyGxlf0Vk65WLDK-Oypm6yXbF5SF7H5s"
NOMBRE_HOJA_STAGING     = "STAGING_EXTRACCION"

# --- Formulario a dar de alta ---
COD_FORMULARIO = "F350"
VERSION_FORM   = "v2026"

libro_registro = gc.open_by_key(ID_REGISTRO_FORMULARIOS)
print("Conectado a:", libro_registro.title)

# ============================================================
# BLOQUE 3 — Catálogo de renglones del formulario 350
# Cada fila de la matriz agrupa hasta 4 renglones:
# (etiqueta, sección, jur_base, jur_retención, nat_base, nat_retención)
# ============================================================

FILAS_MATRIZ = [
    ("Rentas de trabajo",                                   "CONCEPTOS", None, None,  77,  93),
    ("Rentas de pensiones",                                 "CONCEPTOS", None, None,  78,  94),
    ("Honorarios",                                          "CONCEPTOS",  29,  42,  79,  95),
    ("Comisiones",                                          "CONCEPTOS",  30,  43,  80,  96),
    ("Servicios",                                           "CONCEPTOS",  31,  44,  81,  97),
    ("Rendimientos financieros e intereses",                "CONCEPTOS",  32,  45,  82,  98),
    ("Arrendamientos (muebles e inmuebles)",                "CONCEPTOS",  33,  46,  83,  99),
    ("Regalías y explotación de la propiedad intelectual",  "CONCEPTOS",  34,  47,  84, 100),
    ("Dividendos y participaciones",                        "CONCEPTOS",  35,  48,  85, 101),
    ("Compras",                                             "CONCEPTOS",  36,  49,  86, 102),
    ("Transacciones con tarjetas débito y crédito",         "CONCEPTOS",  37,  50,  87, 103),
    ("Contratos de construcción",                           "CONCEPTOS",  38,  51,  88, 104),
    ("Enajenación de activos fijos de personas naturales ante notarios y autoridades de tránsito",
                                                            "CONCEPTOS", None, None,  89, 105),
    ("Loterías, rifas, apuestas y similares",               "CONCEPTOS",  39,  52,  90, 106),
    ("Hidrocarburos, carbón y demás productos mineros",     "CONCEPTOS",  40,  53,  91, 107),
    ("Otros pagos sujetos a retención",                     "CONCEPTOS",  41,  54,  92, 108),
    ("Pagos o abonos en cuenta al exterior a países sin convenio",
                                                            "EXTERIOR",   55,  57, 109, 111),
    ("Pagos o abonos en cuenta al exterior a países con convenio vigente",
                                                            "EXTERIOR",   56,  58, 110, 112),
    ("Contribuyentes exonerados de aportes (art. 114-1 E.T.)",
                                                            "AUTORRETENCIONES",  59,  68, None, None),
    ("Ventas",                                              "AUTORRETENCIONES",  60,  69, 113, 121),
    ("Honorarios",                                          "AUTORRETENCIONES",  61,  70, 114, 122),
    ("Comisiones",                                          "AUTORRETENCIONES",  62,  71, 115, 123),
    ("Servicios",                                           "AUTORRETENCIONES",  63,  72, 116, 124),
    ("Rendimientos financieros",                            "AUTORRETENCIONES",  64,  73, 117, 125),
    ("Pagos mensuales provisionales de car vol (hidrocarburos y demás productos mineros)",
                                                            "AUTORRETENCIONES",  65,  74, 118, 126),
    ("Exportación de hidrocarburos, carbón y demás productos mineros",
                                                            "AUTORRETENCIONES",  66,  75, 119, 127),
    ("Otros conceptos",                                     "AUTORRETENCIONES",  67,  76, 120, 128),
]

# Renglones únicos, sin desagregación por tipo de persona
FILAS_TOTALES = [
    (129, "Menos retenciones practicadas en exceso o indebidas o por operaciones anuladas, rescindidas o resueltas", "TOTALES_RENTA",   "SI"),
    (130, "Total retenciones renta y complementario",                     "TOTALES_RENTA",   "NO"),
    (131, "A responsables del impuesto sobre las ventas",                 "TOTALES_IVA",     "SI"),
    (132, "Practicadas por servicios a no residentes o no domiciliados",  "TOTALES_IVA",     "SI"),
    (133, "Menos retenciones practicadas en exceso o indebidas o por operaciones anuladas, rescindidas o resueltas", "TOTALES_IVA",     "SI"),
    (134, "Total retenciones IVA",                                        "TOTALES_IVA",     "NO"),
    (135, "Retenciones impuesto timbre nacional",                         "TOTALES_TIMBRE",  "SI"),
    (136, "Total retenciones",                                            "TOTALES_GENERAL", "NO"),
    (137, "Sanciones",                                                    "TOTALES_GENERAL", "SI"),
    (138, "Total retenciones más sanciones",                              "TOTALES_GENERAL", "NO"),
]

# Totales de pagos al exterior (hoja 2)
FILAS_EXTERIOR = [
    (148, "Sin convenio persona jurídica", "JURIDICA", "BASE"),
    (150, "Sin convenio persona jurídica", "JURIDICA", "RETENCION"),
    (149, "Sin convenio persona natural",  "NATURAL",  "BASE"),
    (151, "Sin convenio persona natural",  "NATURAL",  "RETENCION"),
    (152, "Con convenio persona jurídica", "JURIDICA", "BASE"),
    (154, "Con convenio persona jurídica", "JURIDICA", "RETENCION"),
    (153, "Con convenio persona natural",  "NATURAL",  "BASE"),
    (155, "Con convenio persona natural",  "NATURAL",  "RETENCION"),
]

print(f"Filas de matriz: {len(FILAS_MATRIZ)} | Totales: {len(FILAS_TOTALES)} | Exterior: {len(FILAS_EXTERIOR)}")

# ============================================================
# BLOQUE 4 — Expansión del catálogo a un registro por renglón
# ============================================================
registros = []

for orden, (etiqueta, seccion, jb, jr, nb, nr) in enumerate(FILAS_MATRIZ, start=1):
    for nro, persona, valor in [(jb, "JURIDICA", "BASE"),
                                (jr, "JURIDICA", "RETENCION"),
                                (nb, "NATURAL",  "BASE"),
                                (nr, "NATURAL",  "RETENCION")]:
        if nro is None:
            continue
        registros.append({
            "pagina": 1, "nro_renglon": nro, "etiqueta": etiqueta,
            "tipo_persona": persona, "tipo_valor": valor,
            "grupo_concepto": f"{seccion}_{orden:02d}", "seccion": seccion,
            "editable": "SI",
        })

for nro, etiqueta, seccion, editable in FILAS_TOTALES:
    registros.append({
        "pagina": 1, "nro_renglon": nro, "etiqueta": etiqueta,
        "tipo_persona": "", "tipo_valor": "TOTAL",
        "grupo_concepto": seccion, "seccion": seccion,
        "editable": editable,
    })

for nro, etiqueta, persona, valor in FILAS_EXTERIOR:
    registros.append({
        "pagina": 2, "nro_renglon": nro, "etiqueta": etiqueta,
        "tipo_persona": persona, "tipo_valor": valor,
        "grupo_concepto": "TOTALES_EXTERIOR", "seccion": "TOTALES_EXTERIOR",
        "editable": "NO",
    })

df_catalogo = pd.DataFrame(registros).sort_values("nro_renglon").reset_index(drop=True)
print(f"Renglones generados: {len(df_catalogo)}")

# ============================================================
# BLOQUE 5 — Control de completitud del catálogo
# ============================================================
ESPERADOS = set(range(29, 139)) | {148, 149, 150, 151, 152, 153, 154, 155}
en_catalogo = set(df_catalogo["nro_renglon"])

faltantes  = sorted(ESPERADOS - en_catalogo)
sobrantes  = sorted(en_catalogo - ESPERADOS)
duplicados = df_catalogo[df_catalogo.duplicated("nro_renglon")]["nro_renglon"].tolist()

print(f"Renglones en catálogo: {len(en_catalogo)} de {len(ESPERADOS)} esperados")
print(f"Faltantes:  {faltantes}")
print(f"Sobrantes:  {sobrantes}")
print(f"Duplicados: {duplicados}")

if not faltantes and not sobrantes and not duplicados:
    print("\nCatálogo íntegro.")

# ============================================================
# BLOQUE 6 — Verificación cruzada contra el PDF oficial
# ============================================================
drive.mount('/content/drive', force_remount=True)

RUTA_BASE = "/content/drive/MyDrive/SISTEMA_DECLARACIONES/2_FORMULARIOS_BASE"
CARPETA_FORM = os.path.join(RUTA_BASE, COD_FORMULARIO, VERSION_FORM)
os.makedirs(CARPETA_FORM, exist_ok=True)

pdfs = [f for f in os.listdir(CARPETA_FORM) if f.lower().endswith(".pdf")]

if len(pdfs) != 1:
    print(f"Verificación omitida: se esperaba 1 PDF y hay {len(pdfs)} en {CARPETA_FORM}")
else:
    palabras = []
    with pdfplumber.open(os.path.join(CARPETA_FORM, pdfs[0])) as pdf:
        for n_pag, pagina in enumerate(pdf.pages, start=1):
            for p in pagina.extract_words(extra_attrs=["upright"]):
                if p.get("upright", True) is False:
                    continue
                palabras.append({"pagina": n_pag, "texto": p["text"].strip(),
                                 "x0": round(p["x0"], 1)})

    df_pdf = pd.DataFrame(palabras).drop_duplicates()
    df_num = df_pdf[df_pdf["texto"].str.match(r"^\d{1,3}$")].copy()
    df_num["nro"] = df_num["texto"].astype(int)
    df_num = df_num[(df_num["nro"].between(29, 155)) & (df_num["x0"] >= 60)]

    BANDAS = {"JURIDICA_BASE": 154.7, "JURIDICA_RETENCION": 265.7,
              "NATURAL_BASE": 376.7, "NATURAL_RETENCION": 487.7}
    pos = df_num.drop_duplicates("nro").set_index("nro")["x0"].to_dict()

    discrepancias = []
    for _, f in df_catalogo[df_catalogo["pagina"] == 1].iterrows():
        clave = f"{f['tipo_persona']}_{f['tipo_valor']}"
        nro = f["nro_renglon"]
        if clave in BANDAS and nro in pos and abs(pos[nro] - BANDAS[clave]) > 5:
            discrepancias.append((nro, clave, pos[nro]))

    print(f"Renglones contrastados: {len(pos)}")
    print("Discrepancias:", discrepancias if discrepancias else "ninguna")

# ============================================================
# BLOQUE 7 — Envío del catálogo a la hoja de staging
# ============================================================
COLUMNAS = ["cod_formulario", "version", "pagina", "nro_renglon", "etiqueta",
            "tipo_persona", "tipo_valor", "grupo_concepto", "seccion",
            "editable", "confianza", "estado_revision"]

df_salida = df_catalogo.copy()
df_salida["cod_formulario"]  = COD_FORMULARIO
df_salida["version"]         = VERSION_FORM
df_salida["confianza"]       = "ALTA"
df_salida["estado_revision"] = "PENDIENTE"
df_salida = df_salida[COLUMNAS]

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

