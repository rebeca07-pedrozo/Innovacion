#1
from google.colab import drive
drive.mount('/content/drive')

import re
import pandas as pd


#2
RUTA_ENTRADA = '/content/drive/MyDrive/EXOGENA/exogena.csv'
RUTA_SALIDA = '/content/drive/MyDrive/EXOGENA/exogena_depurada.csv'
RUTA_AUDITORIA = '/content/drive/MyDrive/EXOGENA/auditoria_completados.csv'

COL_LLAVE = 'concat'
COL_LLAVE_PARTES = ['TIPO_ID', 'NUMERO_ID_BANCO']
SEPARADOR_LLAVE = ''

COL_DIRECCION = 'DIRECCION'
COL_CORREO = 'CORREO_ELECTRONICO'

COLUMNAS_COMPLETITUD = [
    'NUMERO_ID_DIAN',
    'DIGITO_VERIFICACION',
    'NATURALEZA',
    'PRIMER_APELLIDO',
    'SEGUNDO_APELLIDO',
    'PRIMER_NOMBRE',
    'OTROS_NOMBRES',
    'RAZON_SOCIAL',
    'DIRECCION',
    'COD_PAIS',
    'COD_DEPARTAMENTO',
    'COD_MUNICIPIO',
    'CORREO_ELECTRONICO',
    'NUM_FIJO',
    'NUM_CELULAR',
    'X_TD_EXTRANJERO',
    'X_NUM_EXTRANJERO',
    'EX_ID'
]

COLUMNAS_CERO_ES_VACIO = [
    'NUMERO_ID_DIAN',
    'NATURALEZA',
    'PRIMER_APELLIDO',
    'SEGUNDO_APELLIDO',
    'PRIMER_NOMBRE',
    'OTROS_NOMBRES',
    'RAZON_SOCIAL',
    'DIRECCION',
    'COD_PAIS',
    'COD_DEPARTAMENTO',
    'COD_MUNICIPIO',
    'CORREO_ELECTRONICO',
    'NUM_FIJO',
    'NUM_CELULAR',
    'X_TD_EXTRANJERO',
    'X_NUM_EXTRANJERO',
    'EX_ID'
]

VALORES_VACIOS = {
    '', '-', '--', '---', '.', '..', '...', '_', '#', '##',
    'N/A', 'N.A', 'N.A.', 'NA', 'NULL', 'NONE', 'NAN',
    'NO REGISTRA', 'NOREGISTRA', 'NO REPORTA', 'NO TIENE', 'NO APLICA',
    'SIN DATO', 'SIN DATOS', 'SIN INFORMACION', 'SIN INFORMACIÓN',
    'X', 'XX', 'XXX', 'XXXX', 'XXXXX', '@'
}

INCLUIR_LLAVE_EN_SALIDA = False
SEPARADOR_SALIDA = ','
ENCODING_SALIDA = 'utf-8-sig'


#3
def detectar_encoding(ruta):
    for enc in ['utf-8', 'utf-8-sig', 'cp1252', 'latin-1']:
        try:
            with open(ruta, encoding=enc) as f:
                for _ in range(500):
                    if not f.readline():
                        break
            return enc
        except UnicodeDecodeError:
            continue
    return 'latin-1'

def detectar_separador(ruta, encoding):
    with open(ruta, encoding=encoding, errors='replace') as f:
        linea = f.readline()
    return max([';', ',', '|', '\t'], key=linea.count)

ENCODING = detectar_encoding(RUTA_ENTRADA)
SEPARADOR = detectar_separador(RUTA_ENTRADA, ENCODING)

df = pd.read_csv(
    RUTA_ENTRADA,
    sep=SEPARADOR,
    encoding=ENCODING,
    dtype=str,
    na_filter=False,
    on_bad_lines='warn'
)
df.columns = [c.strip() for c in df.columns]

faltantes = [c for c in COL_LLAVE_PARTES + COLUMNAS_COMPLETITUD if c not in df.columns]
if faltantes:
    raise ValueError(f'columnas ausentes en el archivo: {faltantes}')

df[COL_LLAVE] = (
    df[COL_LLAVE_PARTES[0]].str.strip()
    + SEPARADOR_LLAVE
    + df[COL_LLAVE_PARTES[1]].str.strip()
)

print(f'encoding={ENCODING} | separador={repr(SEPARADOR)} | filas={len(df)} | llaves únicas={df[COL_LLAVE].nunique()} | duplicados={len(df) - df[COL_LLAVE].nunique()}')


#4
def normalizar(valor):
    return re.sub(r'\s+', ' ', str(valor)).strip()

def es_vacio(valor, columna):
    v = normalizar(valor).upper()
    if v in VALORES_VACIOS:
        return True
    if columna in COLUMNAS_CERO_ES_VACIO and v and set(v) == {'0'}:
        return True
    return False

llenos = pd.Series(0, index=df.index)
caracteres = pd.Series(0, index=df.index)

for col in COLUMNAS_COMPLETITUD:
    serie = df[col].map(normalizar)
    vacio = serie.map(lambda v, c=col: es_vacio(v, c))
    llenos += (~vacio).astype(int)
    caracteres += serie.where(~vacio, '').str.len()

df['_llenos'] = llenos
df['_caracteres'] = caracteres
df['_orden'] = range(len(df))

print(df['_llenos'].value_counts().sort_index())


#5
df['_dir_norm'] = df[COL_DIRECCION].map(normalizar)
df.loc[df['_dir_norm'].map(lambda v: es_vacio(v, COL_DIRECCION)), '_dir_norm'] = ''
df['_dir_largo'] = df['_dir_norm'].str.len()

df['_cor_norm'] = df[COL_CORREO].map(normalizar)
df.loc[df['_cor_norm'].map(lambda v: es_vacio(v, COL_CORREO)), '_cor_norm'] = ''
df['_cor_vacio'] = (df['_cor_norm'] == '').astype(int)

ganadoras = (
    df.sort_values([COL_LLAVE, '_llenos', '_caracteres', '_orden'],
                   ascending=[True, False, False, True])
      .drop_duplicates(subset=[COL_LLAVE], keep='first')
      .copy()
)

mejor_direccion = (
    df.sort_values([COL_LLAVE, '_dir_largo', '_orden'],
                   ascending=[True, False, True])
      .drop_duplicates(subset=[COL_LLAVE], keep='first')
      [[COL_LLAVE, '_dir_norm', '_dir_largo']]
      .rename(columns={'_dir_norm': '_dir_mejor', '_dir_largo': '_dir_mejor_largo'})
)

mejor_correo = (
    df.sort_values([COL_LLAVE, '_cor_vacio', '_orden'],
                   ascending=[True, True, True])
      .drop_duplicates(subset=[COL_LLAVE], keep='first')
      [[COL_LLAVE, '_cor_norm']]
      .rename(columns={'_cor_norm': '_cor_mejor'})
)

ganadoras = ganadoras.merge(mejor_direccion, on=COL_LLAVE, how='left')
ganadoras = ganadoras.merge(mejor_correo, on=COL_LLAVE, how='left')

ganadoras['_dir_anterior'] = ganadoras['_dir_norm']
ganadoras['_cor_anterior'] = ganadoras['_cor_norm']

reemplazar_dir = ganadoras['_dir_mejor_largo'] > ganadoras['_dir_largo']
reemplazar_cor = (ganadoras['_cor_norm'] == '') & (ganadoras['_cor_mejor'] != '')

ganadoras[COL_DIRECCION] = ganadoras['_dir_norm']
ganadoras.loc[reemplazar_dir, COL_DIRECCION] = ganadoras.loc[reemplazar_dir, '_dir_mejor']

ganadoras[COL_CORREO] = ganadoras['_cor_norm']
ganadoras.loc[reemplazar_cor, COL_CORREO] = ganadoras.loc[reemplazar_cor, '_cor_mejor']

ganadoras['_rep_dir'] = reemplazar_dir
ganadoras['_rep_cor'] = reemplazar_cor
ganadoras['_tipo_completado'] = ''
ganadoras.loc[reemplazar_dir & ~reemplazar_cor, '_tipo_completado'] = 'DIRECCION'
ganadoras.loc[~reemplazar_dir & reemplazar_cor, '_tipo_completado'] = 'CORREO'
ganadoras.loc[reemplazar_dir & reemplazar_cor, '_tipo_completado'] = 'AMBOS'

print(f'filas resultantes={len(ganadoras)} | direcciones completadas={int(reemplazar_dir.sum())} | correos completados={int(reemplazar_cor.sum())} | ambos={int((reemplazar_dir & reemplazar_cor).sum())}')


#6
auditoria = ganadoras.loc[ganadoras['_tipo_completado'] != '', [
    COL_LLAVE, 'TIPO_ID', 'NUMERO_ID_BANCO', 'PRIMER_APELLIDO', 'PRIMER_NOMBRE', 'RAZON_SOCIAL',
    '_tipo_completado', '_dir_anterior', COL_DIRECCION, '_cor_anterior', COL_CORREO
]].rename(columns={
    '_tipo_completado': 'TIPO_COMPLETADO',
    '_dir_anterior': 'DIRECCION_ORIGINAL',
    COL_DIRECCION: 'DIRECCION_FINAL',
    '_cor_anterior': 'CORREO_ORIGINAL',
    COL_CORREO: 'CORREO_FINAL'
})

auditoria = auditoria.sort_values(['TIPO_COMPLETADO', COL_LLAVE])
auditoria.to_csv(RUTA_AUDITORIA, index=False, sep=SEPARADOR_SALIDA, encoding=ENCODING_SALIDA)

print(auditoria['TIPO_COMPLETADO'].value_counts())
display(auditoria.head(20))


#7
columnas_finales = [c for c in df.columns if not c.startswith('_')]
if not INCLUIR_LLAVE_EN_SALIDA:
    columnas_finales = [c for c in columnas_finales if c != COL_LLAVE]

resultado = (
    ganadoras.sort_values('_orden')[columnas_finales]
             .reset_index(drop=True)
)

resultado.to_csv(RUTA_SALIDA, index=False, sep=SEPARADOR_SALIDA, encoding=ENCODING_SALIDA)

print(f'archivo generado: {len(resultado)} filas | columnas={len(columnas_finales)}')



df['concat'] = df['TIPO_ID'].str.strip() + df['NUMERO_ID_BANCO'].str.strip()
print(f'filas={len(df)} | llaves únicas={df["concat"].nunique()} | duplicados={len(df) - df["concat"].nunique()}')