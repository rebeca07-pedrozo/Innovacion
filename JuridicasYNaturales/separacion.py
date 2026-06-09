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
    if cands:
        return cands[0]
    return None

def _digitos(v):
    """Deja solo dígitos y QUITA los ceros de la izquierda (relleno del reporte)."""
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
archivos = listar_txts(ID_CARPETA_DRIVE)
print(f"Archivos .txt encontrados: {len(archivos)}")
if not archivos:
    raise SystemExit("No hay .txt en la carpeta. Revisa el ID (y que la carpeta sea la correcta).")

frames = []
for f in archivos:
    print(f" - Leyendo: {f['name']}")
    try:
        df = leer_df(descargar_texto(f['id']))
    except Exception as e:
        print(f"   ✗ Se omite ({e})")
        continue
    df.columns = df.columns.str.strip()
    c_cuenta = resolver_col(df, COL_CUENTA,  "CUENTA-CO")
    c_cod    = resolver_col(df, COL_COD_AUX, "COD_AUX")
    c_id     = resolver_col(df, COL_IDENT,   "IDENTIFICACION")
    c_val    = resolver_col(df, COL_VALOR,   "VALOR")
    faltan = [n for n, c in [("CUENTA", c_cuenta), ("COD_AUX", c_cod),
                             ("IDENT", c_id), ("VALOR", c_val)] if c is None]
    if faltan:
        print(f"   ¡OJO! faltan {faltan}. Columnas: {list(df.columns)}")
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

# Solo filas con cuenta de EXACTAMENTE 10 dígitos (descarta títulos, guiones y filas partidas)
data = data[data['CUENTA'].str.fullmatch(r'\d{10}')].copy()
if data.empty:
    raise SystemExit("No quedaron filas válidas. Revisa el archivo.")

# Aviso si hay filas sin VALOR (posible línea partida en el .txt original)
sin_valor = int((data['VALOR'].astype(str).str.strip() == '').sum())
if sin_valor:
    print(f"⚠ {sin_valor} fila(s) sin VALOR (posibles renglones partidos). Suman 0; revísalas si el total no cuadra.")

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
carpeta_salida = obtener_o_crear_subcarpeta(ID_CARPETA_DRIVE, 'resultados')
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine='openpyxl') as writer:
    reporte.to_excel(writer, index=False, sheet_name='reporte')
subir_archivo(carpeta_salida, 'reporte_por_cuenta.xlsx', buf.getvalue(),
              'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
print("\n¡Listo! 'reporte_por_cuenta.xlsx' en la subcarpeta 'resultados'.")