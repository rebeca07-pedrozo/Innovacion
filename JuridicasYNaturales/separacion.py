# ============================================================
#  BLOQUE 1: Instalación, imports y conexión a Drive
# ============================================================
!pip install openpyxl -q

import io, re
import pandas as pd
from google.colab import auth
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

auth.authenticate_user()
service = build('drive', 'v3')
print("✅ Conectado a Drive")

# ============================================================
#  BLOQUE 2: CONFIGURACIÓN  ->  EDITA AQUÍ
# ============================================================
ID_CARPETA_DRIVE = "PEGA_AQUI_EL_ID_DE_TU_CARPETA"   # <-- ID de la carpeta de Drive

# ----- ¿SUMAR VALOR o CONTAR registros? (defínelo con tu compañera) -----
TIPO_REPORTE = "suma"      # opciones: "suma"  ó  "conteo"
# TIPO_REPORTE = "conteo"
# ------------------------------------------------------------------------
print("✅ Configuración cargada")

# ============================================================
#  BLOQUE 3: Funciones de Drive + lectura del .txt
#  (detecta solo si es ANCHO FIJO o TABS exportado por Excel)
# ============================================================
def listar_txts(carpeta_id):
    """Lista .txt incluyendo Unidades Compartidas, con paginación."""
    archivos, token = [], None
    while True:
        res = service.files().list(
            q=f"'{carpeta_id}' in parents and trashed=false",
            fields="nextPageToken, files(id,name)",
            pageSize=1000, supportsAllDrives=True,
            includeItemsFromAllDrives=True, pageToken=token
        ).execute()
        archivos += res.get('files', [])
        token = res.get('nextPageToken')
        if not token:
            break
    return [f for f in archivos if f['name'].lower().endswith('.txt')]

def descargar_texto(file_id):
    request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    data = buf.getvalue()
    for enc in ('utf-8', 'latin-1'):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode('utf-8', errors='replace')

def obtener_o_crear_subcarpeta(parent_id, nombre):
    q = (f"'{parent_id}' in parents and name='{nombre}' "
         f"and mimeType='application/vnd.google-apps.folder' and trashed=false")
    items = service.files().list(q=q, fields="files(id)", supportsAllDrives=True,
                                 includeItemsFromAllDrives=True).execute().get('files', [])
    if items:
        return items[0]['id']
    meta = {'name': nombre, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [parent_id]}
    return service.files().create(body=meta, fields='id', supportsAllDrives=True).execute()['id']

def subir_archivo(carpeta_id, nombre, data_bytes, mimetype, reemplazar=True):
    if reemplazar:
        q = f"'{carpeta_id}' in parents and name='{nombre}' and trashed=false"
        for f in service.files().list(q=q, fields="files(id)", supportsAllDrives=True,
                                      includeItemsFromAllDrives=True).execute().get('files', []):
            service.files().delete(fileId=f['id'], supportsAllDrives=True).execute()
    media = MediaIoBaseUpload(io.BytesIO(data_bytes), mimetype=mimetype, resumable=False)
    service.files().create(body={'name': nombre, 'parents': [carpeta_id]},
                           media_body=media, fields='id', supportsAllDrives=True).execute()

def _es_linea_guiones(l):
    s = l.replace(' ', '').replace('\t', '')
    return len(s) >= 10 and s.count('-') / len(s) >= 0.9

def leer_registros(texto):
    """Devuelve DataFrame con CUENTA, COD_AUX, IDENT, VALOR.
       CUENTA y COD_AUX se anclan al inicio; VALOR e IDENT al final
       (así funciona aunque Excel deje columnas fantasma)."""
    lineas = texto.splitlines()
    idx_header = next((i for i, l in enumerate(lineas) if 'COD_AUX' in l), None)
    if idx_header is None:
        raise ValueError("No encontré encabezado con COD_AUX")
    header = lineas[idx_header]
    idx_guion = next((k for k in range(idx_header + 1, len(lineas))
                      if _es_linea_guiones(lineas[k])), None)
    es_tab = '\t' in header

    if es_tab:
        hcols = [c.strip() for c in header.split('\t')]
        spans = None
    else:
        if idx_guion is None:
            raise ValueError("Sin fila de guiones (----) para ancho fijo")
        g = lineas[idx_guion]; spans = []; j = 0; n = len(g)
        while j < n:
            if g[j] == '-':
                ini = j
                while j < n and g[j] == '-': j += 1
                spans.append((ini, j))
            else:
                j += 1
        spans[-1] = (spans[-1][0], 100000)
        hcols = [header[a:b].strip() for (a, b) in spans]

    idx_cuenta = next((i for i, c in enumerate(hcols) if 'CUENTA-CO' in c.upper()), 1)
    idx_cod    = next((i for i, c in enumerate(hcols) if c.upper() == 'COD_AUX'), 9)

    inicio = (idx_guion + 1) if idx_guion is not None else (idx_header + 1)
    out = []
    for l in lineas[inicio:]:
        if not l.strip():
            continue
        if es_tab:
            campos = [p.strip() for p in l.split('\t')]
            while campos and campos[-1] == '':
                campos.pop()
        else:
            campos = [l[a:b].strip() for (a, b) in spans]
        if len(campos) <= idx_cod:   # título o pie de página -> descartar
            continue
        out.append({
            'CUENTA':  campos[idx_cuenta],
            'COD_AUX': campos[idx_cod],
            'IDENT':   campos[-2] if len(campos) >= 2 else '',
            'VALOR':   campos[-1],
        })
    return pd.DataFrame(out)

print("✅ Bloque 3 listo (lectura ancho fijo + tabs)")

# ============================================================
#  BLOQUE 4: Clasificación de persona y parseo de VALOR
# ============================================================
def _digitos(v):
    """Solo dígitos y SIN ceros a la izquierda (relleno del reporte)."""
    return re.sub(r'\D', '', '' if pd.isna(v) else str(v)).lstrip('0')

def parse_valor(v):
    if pd.isna(v):
        return 0.0
    s = str(v).strip()
    if not s:
        return 0.0
    neg = ('(' in s and ')' in s) or s.endswith('-')
    s = re.sub(r'[^\d.,]', '', s)
    if ',' in s and '.' in s:
        s = s.replace(',', '') if s.rfind('.') > s.rfind(',') else s.replace('.', '').replace(',', '.')
    elif ',' in s:
        p = s.split(',')
        s = s.replace(',', '.') if (len(p) == 2 and len(p[1]) in (1, 2)) else s.replace(',', '')
    try:
        x = float(s) if s else 0.0
    except ValueError:
        x = 0.0
    return -x if neg else x

def clasificar_persona(cod_aux, ident):
    """COD_AUX; si está vacío usa IDENTIFICACION (ambos sin ceros a la izquierda).
       <10 díg -> natural | 10 díg inicia en 1 -> natural | 10 díg no inicia en 1 -> jurídica."""
    s = _digitos(cod_aux)
    if not s:
        s = _digitos(ident)
    if not s:
        return 'vacios'
    nn = len(s)
    if nn < 10:  return 'naturales'
    if nn == 10: return 'naturales' if s[0] == '1' else 'juridicas'
    return 'mas_de_10'

print("✅ Bloque 4 listo (clasificación)")

# ============================================================
#  BLOQUE 5: Reporte por cuenta -> 1 Excel
#  A: Cuenta | B: Naturales | C: Juridicas | D: Vacios
# ============================================================
archivos = listar_txts(ID_CARPETA_DRIVE)
print(f"Archivos .txt encontrados: {len(archivos)}")
if not archivos:
    raise SystemExit("No hay .txt en la carpeta. Revisa el ID.")

frames = []
for f in archivos:
    print(f" - Leyendo: {f['name']}")
    try:
        d = leer_registros(descargar_texto(f['id']))
        d['ARCHIVO'] = f['name']
        frames.append(d)
    except Exception as e:
        print(f"   ✗ Se omite ({e})")

if not frames:
    raise SystemExit("Ningún archivo pudo leerse.")

data = pd.concat(frames, ignore_index=True)

# Solo filas con cuenta de EXACTAMENTE 10 dígitos (descarta títulos, guiones y pies)
data = data[data['CUENTA'].astype(str).str.fullmatch(r'\d{10}')].copy()
if data.empty:
    raise SystemExit("No quedaron filas válidas. Revisa el archivo.")

data['_grupo'] = data.apply(lambda r: clasificar_persona(r['COD_AUX'], r['IDENT']), axis=1)
data['_valor'] = data['VALOR'].apply(parse_valor)

print("\nConteo de registros por grupo:")
print(data['_grupo'].value_counts())
n_raros = int((data['_grupo'] == 'mas_de_10').sum())
if n_raros:
    print(f"⚠ {n_raros} registros con MÁS de 10 dígitos -> columna extra 'Mas_de_10_REVISAR'.")

# Suma (o conteo) por cuenta y grupo
if TIPO_REPORTE == 'conteo':
    data['_metrica'] = 1; metrica = '_metrica'
else:
    metrica = '_valor'

resumen = data.groupby(['CUENTA', '_grupo'])[metrica].sum().unstack(fill_value=0)
for gpo in ['naturales', 'juridicas', 'vacios', 'mas_de_10']:
    if gpo not in resumen.columns:
        resumen[gpo] = 0
orden = ['naturales', 'juridicas', 'vacios'] + (['mas_de_10'] if n_raros else [])
reporte = resumen[orden].reset_index()
reporte.columns = (['Cuenta', 'Naturales', 'Juridicas', 'Vacios']
                   + (['Mas_de_10_REVISAR'] if n_raros else []))
reporte = reporte.sort_values('Cuenta').reset_index(drop=True)

# Fila de TOTAL
cols_num = [c for c in reporte.columns if c != 'Cuenta']
total = {'Cuenta': 'TOTAL', **{c: reporte[c].sum() for c in cols_num}}
reporte = pd.concat([reporte, pd.DataFrame([total])], ignore_index=True)

print(f"\nCuentas en el reporte: {len(reporte) - 1}  |  Tipo: {TIPO_REPORTE.upper()}")

# Guardar 1 solo Excel
# ----- Guardar Excel con FECHA Y HORA en el nombre (no sobreescribe) -----
ahora = datetime.now(pytz.timezone('America/Bogota')).strftime('%Y-%m-%d_%H-%M')
nombre_archivo = f'reporte_por_cuenta_{ahora}.xlsx'

carpeta_salida = obtener_o_crear_subcarpeta(ID_CARPETA_DRIVE, 'resultados')
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine='openpyxl') as writer:
    reporte.to_excel(writer, index=False, sheet_name='reporte')
subir_archivo(carpeta_salida, nombre_archivo, buf.getvalue(),
              'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
              reemplazar=False)   # <-- False = cada corrida crea un archivo nuevo
print(f"\n¡Listo! '{nombre_archivo}' en la subcarpeta 'resultados'.")