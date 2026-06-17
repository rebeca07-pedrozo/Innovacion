#1
!pip install openpyxl -q
print("✅ Dependencias listas.")
#2
from google.colab import auth
auth.authenticate_user()
from googleapiclient.discovery import build
drive_service = build('drive', 'v3')
print("✅ Conectado a la API de Google Drive.")
#3
import ipywidgets as widgets
from IPython.display import display

# ════════════════════════════════════════════════════════════════════
#  👉👉  PEGA AQUÍ LOS IDs DE TUS 3 CARPETAS  (van entre comillas)
#       El ID es lo que sale en la URL después de /folders/
# ════════════════════════════════════════════════════════════════════
ID_CARPETA_EXCEL        = "PEGA_AQUI_ID_CARPETA_DE_EXCELES"      # carpeta con el Excel (BASE 2026)
ID_CARPETA_CONSOLIDADOS = "PEGA_AQUI_ID_CARPETA_DE_CONSOLIDADOS" # carpeta con los consolidados (.txt)
ID_CARPETA_RESULTADOS   = "PEGA_AQUI_ID_CARPETA_DONDE_GUARDAR"   # aquí se creará la subcarpeta 'Resultados'
# ════════════════════════════════════════════════════════════════════

SHEET_MIME    = 'application/vnd.google-apps.spreadsheet'
FOLDER_MIME   = 'application/vnd.google-apps.folder'
GOOGLE_NATIVO = 'application/vnd.google-apps'

def listar_carpeta(folder_id):
    q = f"'{folder_id}' in parents and trashed=false"
    res = drive_service.files().list(
        q=q, fields="files(id,name,mimeType)", pageSize=1000,
        supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
    return res.get('files', [])

def listar_recursivo(folder_id, prefijo=''):
    items = []
    for f in listar_carpeta(folder_id):
        if f['mimeType'] == FOLDER_MIME:
            items += listar_recursivo(f['id'], prefijo + f['name'] + '/')
        else:
            items.append((prefijo + f['name'], f['id'], f['mimeType']))
    return items

# --- Excel: archivos .xlsx/.xls o Google Sheets ---
arch_excel = listar_recursivo(ID_CARPETA_EXCEL)
op_excel = [(n, (fid, mime)) for (n, fid, mime) in arch_excel
            if n.lower().endswith(('.xlsx', '.xls')) or mime == SHEET_MIME]

# --- Consolidados: cualquier archivo que NO sea Excel ni nativo de Google ---
arch_cons = listar_recursivo(ID_CARPETA_CONSOLIDADOS)
op_cons = [(n, (fid, mime)) for (n, fid, mime) in arch_cons
           if not n.lower().endswith(('.xlsx', '.xls')) and not mime.startswith(GOOGLE_NATIVO)]

dd_excel = widgets.Dropdown(options=op_excel, description='EXCEL:',
                            layout=widgets.Layout(width='95%'))
dd_cons  = widgets.Dropdown(options=op_cons, description='CONSOLIDADO:',
                            layout=widgets.Layout(width='95%'))

print(f"Encontrados: {len(op_excel)} Excel y {len(op_cons)} consolidados.")
print("Selecciona en cada menú el archivo que quieres usar:")
display(widgets.VBox([dd_excel, dd_cons]))
#4
import io
from googleapiclient.http import MediaIoBaseDownload

def descargar(file_id, mime, destino):
    if mime == SHEET_MIME:   # si el Excel es un Google Sheet, lo exporta a .xlsx
        req = drive_service.files().export_media(
            fileId=file_id,
            mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    else:
        req = drive_service.files().get_media(fileId=file_id)
    with io.FileIO(destino, 'wb') as fh:
        downloader = MediaIoBaseDownload(fh, req)
        done = False
        while not done:
            _s, done = downloader.next_chunk()

excel_id, excel_mime = dd_excel.value
cons_id,  cons_mime  = dd_cons.value

RUTA_EXCEL = '/content/base.xlsx'
RUTA_TXT   = '/content/consolidado.txt'

descargar(excel_id, excel_mime, RUTA_EXCEL)
descargar(cons_id,  cons_mime,  RUTA_TXT)
print("✅ Archivos descargados:")
print("   Excel       →", dd_excel.label)
print("   Consolidado →", dd_cons.label)
#5
import pandas as pd

NOMBRE_HOJA = "BASE 2026"   # 👈 nombre EXACTO de la pestaña con las cuentas

try:
    df_base = pd.read_excel(RUTA_EXCEL, sheet_name=NOMBRE_HOJA, header=None, dtype=str)
except ValueError:
    import openpyxl
    hojas = openpyxl.load_workbook(RUTA_EXCEL, read_only=True).sheetnames
    raise ValueError(f"No encontré la hoja '{NOMBRE_HOJA}'. Hojas disponibles: {hojas}")

serie = df_base.iloc[1:, 3].dropna().astype(str).str.strip()
serie = serie[serie != '']
serie = serie.str.replace(r'\.0$', '', regex=True)   

cuentas_lista = serie.tolist()
cuentas_set   = set(cuentas_lista)

cuentas_norm = {c.lstrip('0'): c for c in cuentas_set}
claves_norm  = set(cuentas_norm.keys())

print(f"✅ Cuentas leídas: {len(cuentas_lista)} | únicas: {len(cuentas_set)}")
print("Ejemplos:", cuentas_lista[:5])
#6
import re

ENCODING = 'utf-8'    # 👈 cambia a 'latin-1' si ves acentos raros
N_MONTOS = 5          # cuántas columnas numéricas hay al final del reporte

# Encabezados del consolidado, en el MISMO orden del reporte.
# 👉 Si tu reporte trae más/menos columnas, ajústalas aquí.
HEADERS_CONSOLIDADO = [
    "NUM.CUENTA",
    "NUM N.I.T",
    "NOMBRE / RAZON SOCIAL",
    "INGRESOS OPER MES",
    "INGRESOS OPER ANUAL",
    "ING. NO OPER MES",
    "ING. NO OPER ANUAL",
    "IVA MES",
]

sep = re.compile(r'\s{2,}')     # 2+ espacios = separador de campos
resultados, cuentas_encontradas, total = [], set(), 0

print("⏳ Procesando consolidado...")
with open(RUTA_TXT, 'r', encoding=ENCODING, errors='replace') as f:
    for linea in f:
        total += 1
        s  = linea.rstrip('\n').rstrip('\r')
        st = s.strip()
        if not st[:1].isdigit():           # descarta títulos, oficinas, cabeceras
            continue
        campos = sep.split(st)
        if not campos[0].isdigit():
            continue
        clave = campos[0].lstrip('0')      # quita cero(s) a la izquierda
        if clave not in claves_norm:
            continue

        cuenta_orig = cuentas_norm[clave]  # la cuenta tal cual está en el Excel
        cuentas_encontradas.add(cuenta_orig)

        # Reconstrucción robusta: cuenta, nit, [nombre del medio], y los N montos finales
        if len(campos) >= (2 + N_MONTOS + 1):
            cuenta_file = campos[0]
            nit         = campos[1]
            montos      = campos[-N_MONTOS:]
            nombre      = " ".join(campos[2:-N_MONTOS]).strip()
            fila_datos  = [cuenta_file, nit, nombre] + montos
        else:                              # estructura inesperada: rellena
            fila_datos = (campos + [""] * len(HEADERS_CONSOLIDADO))[:len(HEADERS_CONSOLIDADO)]

        resultados.append([cuenta_orig] + fila_datos + [s])

        if total % 300000 == 0:
            print(f"   ... {total:,} líneas | {len(resultados):,} coincidencias")

print(f"✅ Líneas leídas      : {total:,}")
print(f"✅ Cuentas encontradas: {len(cuentas_encontradas)}/{len(cuentas_set)}")
print(f"✅ Registros hallados : {len(resultados):,}")
#7
import os
from datetime import datetime
from googleapiclient.http import MediaFileUpload

# --- DataFrame de resultados ---
columnas = ["Cuenta (BASE 2026)"] + HEADERS_CONSOLIDADO + ["Línea completa (consolidado)"]
df_result = pd.DataFrame(resultados, columns=columnas)
df_no = pd.DataFrame({"Cuenta no encontrada": sorted(cuentas_set - cuentas_encontradas)})

# --- Crear/ubicar subcarpeta 'Resultados' dentro de tu carpeta de resultados ---
def obtener_o_crear_subcarpeta(nombre, parent_id):
    q = (f"'{parent_id}' in parents and name='{nombre}' "
         f"and mimeType='{FOLDER_MIME}' and trashed=false")
    r = drive_service.files().list(q=q, fields='files(id)',
        supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
    if r.get('files'):
        return r['files'][0]['id']
    meta = {'name': nombre, 'mimeType': FOLDER_MIME, 'parents': [parent_id]}
    return drive_service.files().create(body=meta, fields='id',
        supportsAllDrives=True).execute()['id']

def subir(ruta_local, nombre, parent_id):
    meta = {'name': nombre, 'parents': [parent_id]}
    media = MediaFileUpload(ruta_local, resumable=True)
    return drive_service.files().create(body=meta, media_body=media,
        fields='id', supportsAllDrives=True).execute()['id']

id_resultados = obtener_o_crear_subcarpeta("Resultados", ID_CARPETA_RESULTADOS)

marca       = datetime.now().strftime("%Y%m%d_%H%M%S")
nombre_xlsx = f"resultadosBusqueda_{marca}.xlsx"
ruta_local  = f"/content/{nombre_xlsx}"

LIMITE = 1_048_575
with pd.ExcelWriter(ruta_local, engine='openpyxl') as w:
    df_result.head(LIMITE).to_excel(w, sheet_name='Cruce', index=False)
    df_no.to_excel(w, sheet_name='No_encontradas', index=False)

subir(ruta_local, nombre_xlsx, id_resultados)
print(f"✅ Guardado en Drive → carpeta 'Resultados' / {nombre_xlsx}")
#8
print("="*54)
print("             RESUMEN DE LA BÚSQUEDA")
print("="*54)
print(f" Cuentas leídas (Excel)   : {len(cuentas_lista)}")
print(f" Cuentas únicas a buscar  : {len(cuentas_set)}")
print(f" Cuentas ENCONTRADAS      : {len(cuentas_encontradas)}")
print(f" Cuentas NO encontradas   : {len(cuentas_set) - len(cuentas_encontradas)}")
print(f" Registros (filas) traídos: {len(resultados):,}")
print("-"*54)
print(f" Archivo generado : {nombre_xlsx}")
print(f" Ubicación        : subcarpeta 'Resultados'")
print("="*54)