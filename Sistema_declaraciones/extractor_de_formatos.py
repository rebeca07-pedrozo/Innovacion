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
palabras = []

with pdfplumber.open(RUTA_PDF) as pdf:
    for num_pagina, pagina in enumerate(pdf.pages, start=1):
        encontradas = pagina.extract_words(extra_attrs=["size", "upright"])
        for p in encontradas:
            if p.get("upright", True) is False:
                continue
            palabras.append({
                "pagina": num_pagina,
                "texto": p["text"].strip(),
                "x0": round(p["x0"], 1),
                "x1": round(p["x1"], 1),
                "top": round(p["top"], 1),
                "size": round(p.get("size", 0), 1),
            })

df_palabras = pd.DataFrame(palabras)
print(f"Palabras extraídas: {len(df_palabras)} en {df_palabras['pagina'].nunique()} páginas")

#6
PATRON_RENGLON = re.compile(r"^\d{1,3}$")
RANGO_RENGLON  = (20, 200)  

es_numero_corto = df_palabras["texto"].str.match(PATRON_RENGLON)
df_cand = df_palabras[es_numero_corto].copy()
df_cand["valor_num"] = df_cand["texto"].astype(int)

df_cand = df_cand[
    df_cand["valor_num"].between(RANGO_RENGLON[0], RANGO_RENGLON[1])
].copy()

umbral_size = df_cand["size"].median()
df_cand = df_cand[df_cand["size"] <= umbral_size + 0.5].copy()

print(f"Candidatos a renglón: {len(df_cand)}")
print("Rango detectado:", df_cand['valor_num'].min(), "-", df_cand['valor_num'].max())


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

MAPA_BANDAS = {
    1: {"x_min": 235, "x_max": 260, "tipo_persona": "JURIDICA", "tipo_valor": "BASE"},
    2: {"x_min": 400, "x_max": 425, "tipo_persona": "JURIDICA", "tipo_valor": "RETENCION"},
    3: {"x_min": 590, "x_max": 615, "tipo_persona": "NATURAL",  "tipo_valor": "BASE"},
    4: {"x_min": 750, "x_max": 780, "tipo_persona": "NATURAL",  "tipo_valor": "RETENCION"},
}

LIMITE_ETIQUETA_X = 235

def asignar_dimensiones(x0):
    """Devuelve las dimensiones correspondientes a la posición X del renglón."""
    for banda in MAPA_BANDAS.values():
        if banda["x_min"] <= x0 <= banda["x_max"]:
            return banda["tipo_persona"], banda["tipo_valor"]
    return "", ""

df_cand[["tipo_persona", "tipo_valor"]] = df_cand["x0"].apply(
    lambda x: pd.Series(asignar_dimensiones(x))
)

sin_mapear = (df_cand["tipo_persona"] == "").sum()
print(f"Renglones mapeados: {len(df_cand) - sin_mapear} | Sin mapear: {sin_mapear}")

#9 
TOLERANCIA_Y = 6   
df_etiquetas = df_palabras[
    (df_palabras["x0"] < LIMITE_ETIQUETA_X) &
    (df_palabras["texto"].str.contains(r"[A-Za-zÁÉÍÓÚáéíóúÑñ]", regex=True))
].copy()

def buscar_etiqueta(pagina, top):
    """Reconstruye el texto de la etiqueta ubicada en la misma fila."""
    misma_fila = df_etiquetas[
        (df_etiquetas["pagina"] == pagina) &
        (df_etiquetas["top"].between(top - TOLERANCIA_Y, top + TOLERANCIA_Y))
    ].sort_values("x0")
    return " ".join(misma_fila["texto"].tolist()).strip()

df_cand["etiqueta"] = df_cand.apply(
    lambda f: buscar_etiqueta(f["pagina"], f["top"]), axis=1
)

df_cand["confianza"] = df_cand.apply(
    lambda f: "ALTA" if (f["etiqueta"] and f["tipo_persona"]) else "REVISAR", axis=1
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

#11

datos = df_final.fillna("").astype(str).values.tolist()

hoja_staging.clear()
hoja_staging.update(range_name="A1", values=[ENCABEZADOS_STAGING] + datos)

print(f"{len(datos)} registros escritos en {NOMBRE_HOJA_STAGING}")
