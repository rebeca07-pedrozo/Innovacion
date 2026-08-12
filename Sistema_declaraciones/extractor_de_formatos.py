#1
!pip install -q pdfplumber gspread --upgrade

#2
import re
import pandas as pd
import pdfplumber
import gspread
from google.colab import auth, files
from google.auth import default

auth.authenticate_user()
credenciales, _ = default()
gc = gspread.authorize(credenciales)
ID_REGISTRO_FORMULARIOS = "1qXjlhZgsK-jOyGxlf0Vk65WLDK-Oypm6yXbF5SF7H5s"
NOMBRE_HOJA_STAGING     = "STAGING_EXTRACCION"
COD_FORMULARIO = "F350"
VERSION_FORM   = "v2026"
libro_registro = gc.open_by_key(ID_REGISTRO_FORMULARIOS)
print("Conectado a:", libro_registro.title)

#3
ENCABEZADOS_STAGING = [
    "cod_formulario", "version", "pagina", "nro_renglon", "etiqueta",
    "tipo_persona", "tipo_valor", "grupo_concepto", "seccion",
    "x0", "top", "confianza", "estado_revision"
]

try:
    hoja_staging = libro_registro.worksheet(NOMBRE_HOJA_STAGING)
except gspread.WorksheetNotFound:
    hoja_staging = libro_registro.add_worksheet(
        title=NOMBRE_HOJA_STAGING, rows=600, cols=len(ENCABEZADOS_STAGING)
    )
    hoja_staging.update(range_name="A1", values=[ENCABEZADOS_STAGING])

print("Hoja de staging lista:", hoja_staging.title)

#4
import os
from google.colab import drive

drive.mount('/content/drive', force_remount=True)

# --- Ruta de la carpeta de formularios base ---
RUTA_BASE = "/content/drive/MyDrive/SISTEMA_DECLARACIONES/2_FORMULARIOS_BASE"
CARPETA_FORMULARIO = os.path.join(RUTA_BASE, COD_FORMULARIO, VERSION_FORM)

# Se crea la ruta si aún no existe (el bootstrap posterior no la duplicará)
os.makedirs(CARPETA_FORMULARIO, exist_ok=True)

# --- Búsqueda del PDF dentro de la carpeta ---
pdfs = [f for f in os.listdir(CARPETA_FORMULARIO) if f.lower().endswith(".pdf")]

if len(pdfs) == 0:
    raise FileNotFoundError(
        f"No hay ningún PDF en: {CARPETA_FORMULARIO}\n"
        f"Sube allí el formulario oficial y vuelve a ejecutar este bloque."
    )
elif len(pdfs) > 1:
    raise ValueError(
        f"Hay {len(pdfs)} PDF en la carpeta y no se puede determinar cuál usar:\n"
        + "\n".join(f"  - {p}" for p in pdfs)
        + "\nDeja solo el formulario oficial de esta versión."
    )

RUTA_PDF = os.path.join(CARPETA_FORMULARIO, pdfs[0])
print("PDF encontrado:", pdfs[0])

#5
# ============================================================
# BLOQUE 5 — Extracción de palabras con posición (con dedup)
# ============================================================
palabras = []

with pdfplumber.open(RUTA_PDF) as pdf:
    for num_pagina, pagina in enumerate(pdf.pages, start=1):
        for p in pagina.extract_words(extra_attrs=["size", "upright"]):
            # Se descarta texto rotado (marca de agua diagonal)
            if p.get("upright", True) is False:
                continue
            palabras.append({
                "pagina": num_pagina,
                "texto": p["text"].strip(),
                "x0": round(p["x0"], 1),
                "top": round(p["top"], 1),
                "size": round(p.get("size", 0), 1),
            })

df_palabras = pd.DataFrame(palabras)

# El PDF dibuja parte del texto dos veces en la misma coordenada:
# se elimina la capa duplicada antes de cualquier procesamiento.
df_palabras = df_palabras.drop_duplicates(
    subset=["pagina", "texto", "x0", "top"]
).reset_index(drop=True)

print(f"Palabras únicas: {len(df_palabras)} | Páginas: {df_palabras['pagina'].nunique()}")

#6
# ============================================================
# BLOQUE 6 — Identificación de números de renglón
# ============================================================
PATRON_RENGLON = re.compile(r"^\d{1,3}$")
RANGO_RENGLON  = (29, 155)

# Zonas donde NO hay números de renglón (contador de filas de la hoja 2)
X_MINIMO_VALIDO = 60

df_cand = df_palabras[df_palabras["texto"].str.match(PATRON_RENGLON)].copy()
df_cand["valor_num"] = df_cand["texto"].astype(int)

df_cand = df_cand[
    df_cand["valor_num"].between(*RANGO_RENGLON) &
    (df_cand["x0"] >= X_MINIMO_VALIDO)
].copy()

# En la hoja 2 solo son válidos los renglones del bloque de totales (>= 141)
df_cand = df_cand[
    (df_cand["pagina"] == 1) | (df_cand["valor_num"] >= 141)
].copy()

# Un mismo número puede aparecer una sola vez: se conserva el de fuente menor
df_cand = df_cand.sort_values("size").drop_duplicates(
    subset=["pagina", "valor_num"], keep="first"
)

print(f"Renglones detectados: {len(df_cand)}")


#7

UMBRAL_SEPARACION = 25   
def agrupar_bandas(valores_x, umbral):
    """Agrupa posiciones X cercanas en bandas de columna."""
    ordenados = sorted(valores_x)
    bandas, actual = [], [ordenados[0]]
    for x in ordenados[1:]:
        if x - actual[-1] <= umbral:
            actual.append(x)
        else:
            bandas.append(actual)
            actual = [x]
    bandas.append(actual)
    return bandas

for pag in sorted(df_cand["pagina"].unique()):
    x_pagina = df_cand[df_cand["pagina"] == pag]["x0"].tolist()
    bandas = agrupar_bandas(x_pagina, UMBRAL_SEPARACION)
    print(f"\n--- Página {pag} ---")
    for i, b in enumerate(bandas, start=1):
        renglones = df_cand[
            (df_cand["pagina"] == pag) & (df_cand["x0"].between(min(b), max(b)))
        ]["valor_num"].sort_values().tolist()
        print(f"Banda {i}: X {min(b):.0f}-{max(b):.0f} | {len(renglones)} renglones | {renglones[:8]}...")

#8

# ============================================================
# BLOQUE 8 — Mapeo de bandas y secciones (valores reales del F350)
# ============================================================
# Bandas verticales observadas en la extracción
MAPA_BANDAS = [
    {"x_min": 150, "x_max": 160, "tipo_persona": "JURIDICA", "tipo_valor": "BASE"},
    {"x_min": 260, "x_max": 270, "tipo_persona": "JURIDICA", "tipo_valor": "RETENCION"},
    {"x_min": 370, "x_max": 380, "tipo_persona": "NATURAL",  "tipo_valor": "BASE"},
    {"x_min": 482, "x_max": 492, "tipo_persona": "NATURAL",  "tipo_valor": "RETENCION"},
]

# Secciones horizontales por posición vertical
MAPA_SECCIONES = [
    {"pagina": 1, "top_min": 236, "top_max": 452, "seccion": "CONCEPTOS"},
    {"pagina": 1, "top_min": 452, "top_max": 560, "seccion": "AUTORRETENCIONES"},
    {"pagina": 1, "top_min": 560, "top_max": 700, "seccion": "TOTALES"},
    {"pagina": 2, "top_min": 700, "top_max": 780, "seccion": "TOTALES_EXTERIOR"},
]

def asignar_seccion(pagina, top):
    """Devuelve la sección del formulario según la posición vertical."""
    for s in MAPA_SECCIONES:
        if s["pagina"] == pagina and s["top_min"] <= top < s["top_max"]:
            return s["seccion"]
    return ""

def asignar_dimensiones(x0, seccion):
    """Asigna dimensiones solo donde el formulario las tiene."""
    # Los totales ocupan la misma banda X pero no representan tipo de persona
    if seccion in ("TOTALES", "TOTALES_EXTERIOR", ""):
        return "", ""
    for b in MAPA_BANDAS:
        if b["x_min"] <= x0 <= b["x_max"]:
            return b["tipo_persona"], b["tipo_valor"]
    return "", ""

df_cand["seccion"] = df_cand.apply(
    lambda f: asignar_seccion(f["pagina"], f["top"]), axis=1
)
df_cand[["tipo_persona", "tipo_valor"]] = df_cand.apply(
    lambda f: pd.Series(asignar_dimensiones(f["x0"], f["seccion"])), axis=1
)

print(df_cand["seccion"].value_counts().to_string())

#9 
# ============================================================
# BLOQUE 9 — Asociación de etiquetas con agrupación por líneas
# ============================================================
TOL_BANDA = 7.0    # alto de la fila del renglón
TOL_LINEA = 3.0    # separación máxima dentro de una misma línea de texto

# Límite de la columna de etiquetas, distinto en cada página
LIMITE_ETIQUETA = {1: 152.0, 2: 330.0}

df_etiquetas = df_palabras[
    df_palabras["texto"].str.contains(r"[A-Za-zÁÉÍÓÚáéíóúÑñ]", regex=True)
].copy()

def buscar_etiqueta(pagina, top):
    """Reconstruye la etiqueta agrupando primero por línea y luego por posición X."""
    limite = LIMITE_ETIQUETA.get(pagina, 152.0)
    banda = df_etiquetas[
        (df_etiquetas["pagina"] == pagina) &
        (df_etiquetas["x0"] < limite) &
        (df_etiquetas["top"].between(top - TOL_BANDA, top + TOL_BANDA))
    ].sort_values("top")

    if banda.empty:
        return ""

    # Agrupación en líneas de texto por proximidad vertical
    lineas, actual = [], [banda.iloc[0]]
    for _, w in banda.iloc[1:].iterrows():
        if abs(w["top"] - actual[-1]["top"]) <= TOL_LINEA:
            actual.append(w)
        else:
            lineas.append(actual)
            actual = [w]
    lineas.append(actual)

    # Dentro de cada línea sí se ordena horizontalmente
    partes = [
        " ".join(w["texto"] for w in sorted(linea, key=lambda w: w["x0"]))
        for linea in lineas
    ]
    return re.sub(r"\s+", " ", " ".join(partes)).strip()

df_cand["etiqueta"] = df_cand.apply(
    lambda f: buscar_etiqueta(f["pagina"], f["top"]), axis=1
)

df_cand["confianza"] = df_cand.apply(
    lambda f: "ALTA" if (len(f["etiqueta"]) > 3 and f["seccion"]) else "REVISAR",
    axis=1
)

print(df_cand["confianza"].value_counts().to_string())

#10
# ============================================================
# BLOQUE 10 — Armado del dataset para staging
# ============================================================
df_final = pd.DataFrame({
    "cod_formulario": COD_FORMULARIO,
    "version":        VERSION_FORM,
    "pagina":         df_cand["pagina"],
    "nro_renglon":    df_cand["valor_num"],
    "etiqueta":       df_cand["etiqueta"],
    "tipo_persona":   df_cand["tipo_persona"],
    "tipo_valor":     df_cand["tipo_valor"],
    "grupo_concepto": "",          # se completa en la revisión
    "seccion":        "",          # se completa en la revisión
    "x0":             df_cand["x0"],
    "top":            df_cand["top"],
    "confianza":      df_cand["confianza"],
    "estado_revision":"PENDIENTE",
})

df_final = df_final.sort_values(["pagina", "nro_renglon"]).reset_index(drop=True)
print(f"Registros listos para staging: {len(df_final)}")
df_final.head(15)
#10.5
# ============================================================
# BLOQUE 10.5 — Verificación de renglones faltantes
# ============================================================
RENGLONES_ESPERADOS = set(range(29, 139)) | set(range(141, 156))

detectados = set(df_final["nro_renglon"].tolist())
faltantes  = sorted(RENGLONES_ESPERADOS - detectados)
sobrantes  = sorted(detectados - RENGLONES_ESPERADOS)
sin_etiqueta = df_final[df_final["etiqueta"].str.len() < 4]["nro_renglon"].tolist()

print(f"Detectados: {len(detectados)} de {len(RENGLONES_ESPERADOS)}")
print(f"Faltantes: {faltantes}")
print(f"Sobrantes: {sobrantes}")
print(f"Sin etiqueta: {sin_etiqueta}")

#11

datos = df_final.fillna("").astype(str).values.tolist()

hoja_staging.clear()
hoja_staging.update(range_name="A1", values=[ENCABEZADOS_STAGING] + datos)

print(f"{len(datos)} registros escritos en {NOMBRE_HOJA_STAGING}")
