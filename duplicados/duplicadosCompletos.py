#1
from google.colab import drive
drive.mount('/content/drive')

import re
import pandas as pd


#2
RUTA_ENTRADA = '/content/drive/MyDrive/EXOGENA/exogena.csv'
RUTA_SALIDA = '/content/drive/MyDrive/EXOGENA/exogena_depurada.csv'
RUTA_REPORTE = '/content/drive/MyDrive/EXOGENA/reporte_consolidacion.csv'

COL_LLAVE = '_llave'
COL_TIPO = 'TIPO_ID'
COL_NUMERO = 'NUMERO_ID_BANCO'
COL_DIRECCION = 'DIRECCION'
COL_CORREO = 'CORREO_ELECTRONICO'

COLUMNAS_DV = ['DIGITO_VERIFICACION']

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

for c in [COL_TIPO, COL_NUMERO]:
    if c not in df.columns:
        raise ValueError(f'falta la columna {c}. columnas encontradas: {list(df.columns)}')

COLUMNAS_ORIGINALES = list(df.columns)
COLUMNAS_DATOS = [c for c in COLUMNAS_ORIGINALES if c not in (COL_TIPO, COL_NUMERO)]

print(f'encoding={ENCODING} | separador={repr(SEPARADOR)} | filas={len(df)} | columnas={len(COLUMNAS_ORIGINALES)}')
print(COLUMNAS_ORIGINALES)


#4
def normalizar_serie(s):
    return s.astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()

def marcar_vacios(s, columna):
    v = s.str.upper()
    vacio = v.isin(VALORES_VACIOS)
    if columna not in COLUMNAS_DV:
        vacio = vacio | v.str.fullmatch(r'0+').fillna(False)
    return vacio

trabajo = pd.DataFrame(index=df.index)
for col in COLUMNAS_ORIGINALES:
    s = normalizar_serie(df[col])
    trabajo[col] = s.mask(marcar_vacios(s, col), '')

trabajo[COL_LLAVE] = trabajo[COL_TIPO] + '|' + trabajo[COL_NUMERO]
trabajo['_orden'] = range(len(trabajo))

sin_llave = int(((trabajo[COL_TIPO] == '') | (trabajo[COL_NUMERO] == '')).sum())
if sin_llave:
    print(f'ATENCION: {sin_llave} filas sin TIPO_ID o NUMERO_ID_BANCO')


#5
filas_ini = len(trabajo)
llaves_ini = trabajo[COL_LLAVE].nunique()
vacios_ini = {c: int((trabajo[c] == '').sum()) for c in COLUMNAS_DATOS}

antes = pd.DataFrame({
    'METRICA': ['Filas totales', 'Registros únicos', 'Filas duplicadas',
                'Sin dirección', 'Sin correo'],
    'VALOR': [filas_ini, llaves_ini, filas_ini - llaves_ini,
              vacios_ini.get(COL_DIRECCION, 0), vacios_ini.get(COL_CORREO, 0)]
})

print('=== ANTES ===')
display(antes)


#6
trabajo = trabajo.sort_values([COL_LLAVE, '_orden']).reset_index(drop=True)
grupos = trabajo.groupby(COL_LLAVE, sort=False)

consolidado = pd.DataFrame({COL_LLAVE: list(grupos.groups.keys())}).set_index(COL_LLAVE)

for col in COLUMNAS_ORIGINALES:
    if col == COL_DIRECCION:
        aux = trabajo[[COL_LLAVE, col, '_orden']].copy()
        aux['_largo'] = aux[col].str.len()
        mejor = (
            aux.sort_values([COL_LLAVE, '_largo', '_orden'], ascending=[True, False, True])
               .drop_duplicates(subset=[COL_LLAVE], keep='first')
               .set_index(COL_LLAVE)[col]
        )
    else:
        aux = trabajo[[COL_LLAVE, col, '_orden']].copy()
        aux['_vacio'] = (aux[col] == '').astype(int)
        aux['_largo'] = aux[col].str.len()
        mejor = (
            aux.sort_values([COL_LLAVE, '_vacio', '_largo', '_orden'],
                            ascending=[True, True, False, True])
               .drop_duplicates(subset=[COL_LLAVE], keep='first')
               .set_index(COL_LLAVE)[col]
        )
    consolidado[col] = mejor.reindex(consolidado.index)

consolidado['_orden'] = grupos['_orden'].min().reindex(consolidado.index)
consolidado = consolidado.sort_values('_orden').reset_index()

print(f'consolidado: {len(consolidado)} registros')


#7
salida = consolidado[COLUMNAS_ORIGINALES].copy()
salida.to_csv(RUTA_SALIDA, index=False, sep=SEPARADOR_SALIDA, encoding=ENCODING_SALIDA)

vacios_fin = {c: int((salida[c] == '').sum()) for c in COLUMNAS_DATOS}

despues = pd.DataFrame({
    'METRICA': ['Registros finales', 'Duplicados eliminados', '% reducción',
                'Sin dirección', 'Sin correo'],
    'VALOR': [len(salida), filas_ini - len(salida),
              f'{(filas_ini - len(salida)) / filas_ini * 100:.1f}%',
              vacios_fin.get(COL_DIRECCION, 0), vacios_fin.get(COL_CORREO, 0)]
})

print('=== DESPUÉS ===')
display(despues)


#8
disponible = grupos[COLUMNAS_DATOS].apply(lambda g: (g != '').any())
disponible = disponible.reindex(consolidado[COL_LLAVE].values)

perdidos = {}
for col in COLUMNAS_DATOS:
    perdidos[col] = int(((salida[col].values == '') & disponible[col].values).sum())

detalle = pd.DataFrame({
    'COLUMNA': COLUMNAS_DATOS,
    'VACIOS_ANTES': [vacios_ini[c] for c in COLUMNAS_DATOS],
    'VACIOS_DESPUES': [vacios_fin[c] for c in COLUMNAS_DATOS],
    'DATOS_PERDIDOS': [perdidos[c] for c in COLUMNAS_DATOS]
})
detalle.to_csv(RUTA_REPORTE, index=False, sep=SEPARADOR_SALIDA, encoding=ENCODING_SALIDA)

total_perdidos = int(detalle['DATOS_PERDIDOS'].sum())
print(f'DATOS PERDIDOS (campo vacío existiendo dato en un duplicado): {total_perdidos}')
display(detalle)
