#1
from google.colab import drive
drive.mount('/content/drive')

import re
import pandas as pd


#2
RUTA_ENTRADA = '/content/drive/MyDrive/EXOGENA/exogena.csv'
RUTA_SALIDA = '/content/drive/MyDrive/EXOGENA/exogena_depurada.csv'
RUTA_AUDITORIA = '/content/drive/MyDrive/EXOGENA/auditoria_direcciones.csv'

COL_LLAVE = 'concat'
COL_DIRECCION = 'DIRECCION'

COLUMNAS_COMPLETITUD = [
    'NOMBRES',
    'DIRECCION',
    'dian',
    'CORREO_ELECTRONICO',
    'NUM_CELULAR',
    'DIGITO_VERIFICACION',
    'X_TD_EXTRANJERO',
    'X_NUM_EXTRANJERO',
    'EX_ID'
]

COLUMNAS_CERO_ES_VACIO = [
    'NOMBRES',
    'DIRECCION',
    'dian',
    'CORREO_ELECTRONICO',
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

print(f'encoding={ENCODING} | separador={repr(SEPARADOR)} | filas={len(df)} | llaves únicas={df[COL_LLAVE].nunique()}')


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
      [[COL_LLAVE, '_dir_norm', '_dir_largo', '_orden']]
      .rename(columns={'_dir_norm': '_dir_mejor',
                       '_dir_largo': '_dir_mejor_largo',
                       '_orden': '_orden_dir'})
)

ganadoras = ganadoras.merge(mejor_direccion, on=COL_LLAVE, how='left')

ganadoras['_dir_anterior'] = ganadoras['_dir_norm']
reemplazar = ganadoras['_dir_mejor_largo'] > ganadoras['_dir_largo']

ganadoras[COL_DIRECCION] = ganadoras['_dir_norm']
ganadoras.loc[reemplazar, COL_DIRECCION] = ganadoras.loc[reemplazar, '_dir_mejor']
ganadoras['_reemplazo'] = reemplazar

print(f'filas resultantes={len(ganadoras)} | direcciones reemplazadas={int(reemplazar.sum())}')

#6
auditoria = ganadoras.loc[ganadoras['_reemplazo'], [
    COL_LLAVE, 'TIPO_ID', 'NUMERO_ID_BANCO', 'NOMBRES',
    '_dir_anterior', COL_DIRECCION, '_dir_largo', '_dir_mejor_largo'
]].rename(columns={
    '_dir_anterior': 'DIRECCION_ORIGINAL',
    COL_DIRECCION: 'DIRECCION_FINAL',
    '_dir_largo': 'LARGO_ORIGINAL',
    '_dir_mejor_largo': 'LARGO_FINAL'
})

auditoria['DIFERENCIA'] = auditoria['LARGO_FINAL'] - auditoria['LARGO_ORIGINAL']
auditoria = auditoria.sort_values('DIFERENCIA', ascending=False)
auditoria.to_csv(RUTA_AUDITORIA, index=False, sep=SEPARADOR_SALIDA, encoding=ENCODING_SALIDA)

display(auditoria.head(20))

#7
columnas_finales = [c for c in df.columns if not c.startswith('_')]

resultado = (
    ganadoras.sort_values('_orden')[columnas_finales]
             .reset_index(drop=True)
)

resultado.to_csv(RUTA_SALIDA, index=False, sep=SEPARADOR_SALIDA, encoding=ENCODING_SALIDA)

print(f'archivo generado: {len(resultado)} filas | llaves únicas={resultado[COL_LLAVE].nunique()}')















