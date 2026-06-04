# ============================================================
#  SEPARACIÓN DE COD_AUX POR CANTIDAD DE DÍGITOS
#  - Lee todos los .txt de una carpeta de Drive (por ID)
#  - Grupo 1: menos de 10 dígitos
#  - Grupo 2: 10 dígitos que inician en 1   -> personas NATURALES
#  - Grupo 3: 10 dígitos que NO inician en 1 -> personas JURÍDICAS (las que te interesan)
# ============================================================

!pip install openpyxl -q

import io, re
import pandas as pd
from google.colab import auth
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# ============================================================
#  CONFIGURACIÓN  ->  EDITA AQUÍ
# ============================================================
ID_CARPETA_DRIVE = "PEGA_AQUI_EL_ID_DE_TU_CARPETA"   # <-- ID de la carpeta de Drive
NOMBRE_COLUMNA   = "COD_AUX"        # nombre exacto de la columna a analizar
DELIMITADOR_ENTRADA = None          # None = autodetecta; o pon "\t", ";", ","  o "|"
SEP_SALIDA_TXT      = "\t"          # separador con el que se guardarán los .txt de salida

# ----- OPCIÓN DE OUTPUT (coméntala/descoméntala según lo que definan con tu jefa) -----
FORMATO_SALIDA = "ambos"     # opciones: "txt", "excel", "ambos"
# FORMATO_SALIDA = "txt"
# FORMATO_SALIDA = "excel"
# -------------------------------------------------------------------------------------

# ============================================================
#  AUTENTICACIÓN Y CONEXIÓN A DRIVE
# ============================================================
auth.authenticate_user()
service = build('drive', 'v3')

# ============================================================
#  FUNCIONES AUXILIARES
# ============================================================
def listar_txts(carpeta_id):
    q = f"'{carpeta_id}' in parents and trashed=false"
    res = service.files().list(q=q, fields="files(id,name,mimeType)", pageSize=1000).execute()
    return [f for f in res.get('files', []) if f['name'].lower().endswith('.txt')]

def descargar_texto(file_id):
    request = service.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    data = buf.getvalue()
    for enc in ('utf-8', 'latin-1'):   # intenta utf-8 y si no, latin-1 (común en Colombia)
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode('utf-8', errors='replace')

def leer_df(texto):
    if DELIMITADOR_ENTRADA is None:
        return pd.read_csv(io.StringIO(texto), sep=None, engine='python', dtype=str)
    return pd.read_csv(io.StringIO(texto), sep=DELIMITADOR_ENTRADA, dtype=str)

def obtener_o_crear_subcarpeta(parent_id, nombre):
    q = (f"'{parent_id}' in parents and name='{nombre}' "
         f"and mimeType='application/vnd.google-apps.folder' and trashed=false")
    items = service.files().list(q=q, fields="files(id)").execute().get('files', [])
    if items:
        return items[0]['id']
    meta = {'name': nombre, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [parent_id]}
    return service.files().create(body=meta, fields='id').execute()['id']

def subir_archivo(carpeta_id, nombre, data_bytes, mimetype, reemplazar=True):
    if reemplazar:  # borra archivos con el mismo nombre para no duplicar en cada corrida
        q = f"'{carpeta_id}' in parents and name='{nombre}' and trashed=false"
        for f in service.files().list(q=q, fields="files(id)").execute().get('files', []):
            service.files().delete(fileId=f['id']).execute()
    media = MediaIoBaseUpload(io.BytesIO(data_bytes), mimetype=mimetype, resumable=False)
    service.files().create(body={'name': nombre, 'parents': [carpeta_id]},
                           media_body=media, fields='id').execute()

def clasificar(valor):
    s = re.sub(r'\D', '', '' if pd.isna(valor) else str(valor))  # deja solo dígitos
    n = len(s)
    if n == 0:        return 'sin_dato'
    if n < 10:        return 'menos_de_10'
    if n == 10:       return 'naturales' if s[0] == '1' else 'juridicas'
    return 'mas_de_10'   # por si llega a venir algo con más de 10 dígitos

def exportar(df, nombre_base, carpeta_id):
    if df.empty:
        print(f"   (vacío) {nombre_base} -> no se genera archivo")
        return
    if FORMATO_SALIDA in ('txt', 'ambos'):
        contenido = df.to_csv(index=False, sep=SEP_SALIDA_TXT)
        subir_archivo(carpeta_id, nombre_base + '.txt', contenido.encode('utf-8'), 'text/plain')
    if FORMATO_SALIDA in ('excel', 'ambos'):
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        subir_archivo(carpeta_id, nombre_base + '.xlsx', buf.getvalue(),
                      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    print(f"   OK -> {nombre_base}  ({len(df)} filas)")

# ============================================================
#  PROCESO PRINCIPAL
# ============================================================
archivos = listar_txts(ID_CARPETA_DRIVE)
print(f"Archivos .txt encontrados: {len(archivos)}")
if not archivos:
    raise SystemExit("No hay archivos .txt en la carpeta. Revisa el ID de la carpeta.")

frames = []
for f in archivos:
    print(f" - Leyendo: {f['name']}")
    df = leer_df(descargar_texto(f['id']))
    df.columns = df.columns.str.strip()          # limpia espacios en los nombres de columna
    if NOMBRE_COLUMNA not in df.columns:
        print(f"   ¡OJO! '{NOMBRE_COLUMNA}' no está. Columnas disponibles: {list(df.columns)}")
        continue
    df['archivo_origen'] = f['name']
    frames.append(df)

if not frames:
    raise SystemExit(f"Ningún archivo tiene la columna '{NOMBRE_COLUMNA}'.")

data = pd.concat(frames, ignore_index=True)
data['_grupo'] = data[NOMBRE_COLUMNA].apply(clasificar)

print("\nResumen por grupo:")
print(data['_grupo'].value_counts())

# Separar en los grupos
g_juridicas  = data[data['_grupo'] == 'juridicas' ].drop(columns='_grupo')   # 10 díg, NO inician 1
g_naturales  = data[data['_grupo'] == 'naturales' ].drop(columns='_grupo')   # 10 díg, inician 1
g_menos10    = data[data['_grupo'] == 'menos_de_10'].drop(columns='_grupo')  # menos de 10 díg
g_otros      = data[data['_grupo'].isin(['mas_de_10', 'sin_dato'])].drop(columns='_grupo')

# Carpeta de salida (subcarpeta 'resultados' dentro de la misma carpeta)
carpeta_salida = obtener_o_crear_subcarpeta(ID_CARPETA_DRIVE, 'resultados')

print("\nGenerando archivos de salida...")
exportar(g_juridicas, "JURIDICAS_10digitos_no_inician_1", carpeta_salida)
exportar(g_naturales, "NATURALES_10digitos_inician_1",   carpeta_salida)
exportar(g_menos10,   "MENOS_DE_10_digitos",             carpeta_salida)
exportar(g_otros,     "OTROS_revisar",                   carpeta_salida)  # casos raros, por si acaso

print("\n¡Listo! Revisa la subcarpeta 'resultados' dentro de tu carpeta de Drive.")