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

# Nombres de las columnas tal como vienen en el .txt (ajusta solo si cambian)
COL_CUENTA  = "CUENTA-CO"
COL_COD_AUX = "COD_AUX"
COL_IDENT   = "# IDENTIFICACION"
COL_VALOR   = "VALOR"

# ----- ¿SUMAR VALOR o CONTAR registros? (defínelo con tu compañera) -----
TIPO_REPORTE = "suma"      # opciones: "suma"  ó  "conteo"
# TIPO_REPORTE = "conteo"
# ------------------------------------------------------------------------
print("✅ Configuración cargada")

# ============================================================
#  BLOQUE 3: Funciones de Drive + lectura del .txt (ANCHO FIJO)
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
    for enc in ('utf-8', 'latin-1'):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode('utf-8', errors='replace')

def obtener_o_crear_subcarpeta(parent_id, nombre):
    q = (f"'{parent_id}' in parents and name='{nombre}' "
         f"and mimeType='application/vnd.google-apps.folder' and trashed=false")
    items = service.files().list(q=q, fields="files(id)").execute().get('files', [])
    if items:
        return items[0]['id']
    meta = {'name': nombre, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [parent_id]}
    return service.files().create(body=meta, fields='id').execute()['id']

def subir_archivo(carpeta_id, nombre, data_bytes, mimetype, reemplazar=True):
    if reemplazar:
        q = f"'{carpeta_id}' in parents and name='{nombre}' and trashed=false"
        for f in service.files().list(q=q, fields="files(id)").execute().get('files', []):
            service.files().delete(fileId=f['id']).execute()
    media = MediaIoBaseUpload(io.BytesIO(data_bytes), mimetype=mimetype, resumable=False)
    service.files().create(body={'name': nombre, 'parents': [carpeta_id]},
                           media_body=media, fields='id').execute()

# ---------- helpers de ancho fijo ----------
def _es_linea_guiones(l):
    """True si la línea es básicamente '----' (guiones y espacios)."""
    s = l.replace(' ', '')
    if len(s) < 10:
        return False
    d = s.count('-')
    return d >= 10 and d / len(s) >= 0.9

def _spans_desde_guiones(linea):
    """Posiciones (inicio, fin) de cada columna según los bloques de guiones."""
    spans, j, n = [], 0, len(linea)
    while j < n:
        if linea[j] == '-':
            ini = j
            while j < n and linea[j] == '-':
                j += 1
            spans.append((ini, j))
        else:
            j += 1
    if spans:                       # última columna (VALOR) hasta el final
        a, _ = spans[-1]
        spans[-1] = (a, 100000)
    return spans

def _spans_desde_header(header):
    """Plan B: deduce columnas del encabezado (texto separado por 2+ espacios)."""
    starts, j, n = [], 0, len(header)
    while j < n:
        if header[j] != ' ':
            if j == 0 or (header[j-1] == ' ' and header[j-2] == ' '):
                starts.append(j)
        j += 1
    spans = []
    for i, s in enumerate(starts):
        e = starts[i+1] if i + 1 < len(starts) else 100000
        spans.append((s, e))
    return spans

def leer_df(texto):
    """Lee el reporte de ANCHO FIJO usando la fila de guiones (o el encabezado como plan B)."""
    lineas = texto.splitlines()

    idx_header = next((i for i, l in enumerate(lineas) if 'COD_AUX' in l), None)
    if idx_header is None:
        raise ValueError("No encontré la fila de encabezado con COD_AUX")

    # buscar guiones en TODO el archivo después del encabezado
    idx_guion = next((k for k in range(idx_header + 1, len(lineas))
                      if _es_linea_guiones(lineas[k])), None)

    header = lineas[idx_header]
    if idx_guion is not None:
        spans = _spans_desde_guiones(lineas[idx_guion])
        inicio_datos = idx_guion + 1
    else:
        print("   ⚠ No hallé fila de guiones; uso posiciones del encabezado.")
        spans = _spans_desde_header(header)
        inicio_datos = idx_header + 1

    if not spans:
        raise ValueError("No pude determinar las columnas (ni guiones ni encabezado).")

    # nombres de columna
    cols, vistos = [], {}
    for (a, b) in spans:
        nombre = header[a:b].strip() or f"col_{a}"
        if nombre in vistos:
            vistos[nombre] += 1
            nombre = f"{nombre}_{vistos[nombre]}"
        else:
            vistos[nombre] = 0
        cols.append(nombre)

    # datos
    filas = []
    for l in lineas[inicio_datos:]:
        if not l.strip():
            continue
        filas.append([l[a:b].strip() for (a, b) in spans])

    return pd.DataFrame(filas, columns=cols)

print("✅ Funciones de Drive y lectura (ancho fijo, robusto) listas")
# ============================================================
#  BLOQUE 4: Clasificación de persona y parseo de VALOR
# ============================================================
def resolver_col(df, objetivo, palabra_clave=None):
    if objetivo in df.columns:
        return objetivo
    norm = {c.strip().upper(): c for c in df.columns}
    if objetivo.strip().upper() in norm:
        return norm[objetivo.strip().upper()]
    clave = (palabra_clave or objetivo).strip().upper()
    cands = [c for c in df.columns if clave in c.strip().upper()]
    if len(cands) == 1:
        return cands[0]
    if len(cands) > 1:
        print(f"   ⚠ varias columnas con '{clave}': {cands}; uso la 1ra")
        return cands[0]
    return None

def _digitos(v):
    """Deja solo dígitos y QUITA los ceros de la izquierda (relleno del reporte)."""
    s = re.sub(r'\D', '', '' if pd.isna(v) else str(v))
    return s.lstrip('0')

def parse_valor(v):
    if pd.isna(v):
        return 0.0
    s = str(v).strip()
    if not s:
        return 0.0
    neg = ('(' in s and ')' in s) or s.endswith('-')
    s = re.sub(r'[^\d.,]', '', s)
    if ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
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
    n = len(s)
    if n < 10:  return 'naturales'
    if n == 10: return 'naturales' if s[0] == '1' else 'juridicas'
    return 'mas_de_10'

print("✅ Funciones de clasificación listas")

# ============================================================
#  BLOQUE 5: Reporte por cuenta -> 1 Excel
#  A: Cuenta | B: Naturales | C: Juridicas | D: Vacios
# ============================================================
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
        df = leer_df(descargar_texto(f['id']))
    except Exception as e:
        print(f"   ✗ No se pudo leer ({e}); se omite este archivo.")
        continue
    df.columns = df.columns.str.strip()
    c_cuenta = resolver_col(df, COL_CUENTA,  "CUENTA-CO")
    c_cod    = resolver_col(df, COL_COD_AUX, "COD_AUX")
    c_id     = resolver_col(df, COL_IDENT,   "IDENTIFICACION")
    c_val    = resolver_col(df, COL_VALOR,   "VALOR")
    faltan = [n for n, c in [("CUENTA", c_cuenta), ("COD_AUX", c_cod),
                             ("IDENT", c_id), ("VALOR", c_val)] if c is None]
    if faltan:
        print(f"   ¡OJO! faltan columnas {faltan}. Disponibles: {list(df.columns)}")
        continue
    frames.append(pd.DataFrame({
        'CUENTA':  df[c_cuenta].astype(str).str.strip(),
        'COD_AUX': df[c_cod],
        'IDENT':   df[c_id],
        'VALOR':   df[c_val],
    }))

if not frames:
    raise SystemExit("Ningún archivo trajo las columnas necesarias.")

data = pd.concat(frames, ignore_index=True)

# Conservar solo filas con cuenta NUMÉRICA (descarta títulos, encabezados repetidos, guiones)
data = data[data['CUENTA'].str.fullmatch(r'\d+')].copy()
if data.empty:
    raise SystemExit("No quedaron filas con cuenta numérica. Revisa el archivo.")

data['_grupo'] = data.apply(lambda r: clasificar_persona(r['COD_AUX'], r['IDENT']), axis=1)
data['_valor'] = data['VALOR'].apply(parse_valor)

print("\nConteo de registros por grupo:")
print(data['_grupo'].value_counts())
n_raros = int((data['_grupo'] == 'mas_de_10').sum())
if n_raros:
    print(f"⚠ {n_raros} registros con MÁS de 10 dígitos -> van en columna extra 'Mas_de_10_REVISAR'.")

# ----- Suma (o conteo) por cuenta y grupo -----
if TIPO_REPORTE == 'conteo':
    data['_metrica'] = 1
    metrica = '_metrica'
else:
    metrica = '_valor'

resumen = data.groupby(['CUENTA', '_grupo'])[metrica].sum().unstack(fill_value=0)
for g in ['naturales', 'juridicas', 'vacios', 'mas_de_10']:
    if g not in resumen.columns:
        resumen[g] = 0

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

# ----- Guardar 1 solo Excel -----
carpeta_salida = obtener_o_crear_subcarpeta(ID_CARPETA_DRIVE, 'resultados')
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine='openpyxl') as writer:
    reporte.to_excel(writer, index=False, sheet_name='reporte')
subir_archivo(carpeta_salida, 'reporte_por_cuenta.xlsx', buf.getvalue(),
              'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
print("\n¡Listo! 'reporte_por_cuenta.xlsx' en la subcarpeta 'resultados'.")