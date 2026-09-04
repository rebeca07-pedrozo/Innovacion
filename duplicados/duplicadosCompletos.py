
df['concat'] = df['TIPO_ID'].str.strip() + df['NUMERO_ID_BANCO'].str.strip()
print(f'filas={len(df)} | llaves únicas={df["concat"].nunique()} | duplicados={len(df) - df["concat"].nunique()}')


#1
from google.colab import drive
drive.mount('/content/drive')

import re
import pandas as pd


#2
RUTA_ENTRADA = '/content/drive/MyDrive/EXOGENA/exogena.csv'
RUTA_SALIDA = '/content/drive/MyDrive/EXOGENA/exogena_depurada.csv'
RUTA_AUDITORIA = '/content/drive/MyDrive/EXOGENA/auditoria_consolidacion.csv'

COL_LLAVE = 'concat'
COL_LLAVE_PARTES = ['TIPO_ID', 'NUMERO_ID_BANCO']
SEPARADOR_LLAVE = ''

COL_DIRECCION = 'DIRECCION'
COL_CORREO = 'CORREO_ELECTRONICO'

COLUMNAS_IDENTIDAD = [
    'NATURALEZA',
    'PRIMER_APELLIDO',
    'SEGUNDO_APELLIDO',
    'PRIMER_NOMBRE',
    'OTROS_NOMBRES',
    'RAZON_SOCIAL'
]

COLUMNAS_UBICACION = [
    'DIRECCION',
    'COD_PAIS',
    'COD_DEPARTAMENTO',
    'COD_MUNICIPIO'
]

COLUMNAS_INDEPENDIENTES = [
    'NUMERO_ID_DIAN',
    'DIGITO_VERIFICACION',
    'CORREO_ELECTRONICO',
    'NUM_FIJO',
    'NUM_CELULAR',
    'X_TD_EXTRANJERO',
    'X_NUM_EXTRANJERO',
    'EX_ID'
]

COLUMNAS_COMPLETITUD = COLUMNAS_IDENTIDAD + COLUMNAS_UBICACION + COLUMNAS_INDEPENDIENTES

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

df['_orden'] = range(len(df))


#4
def normalizar(valor):
    return re.sub(r'\s+', ' ', str(valor)).strip()

def marcar_vacios(serie, columna):
    v = serie.str.upper()
    vacio = v.isin(VALORES_VACIOS)
    if columna in COLUMNAS_CERO_ES_VACIO:
        vacio = vacio | v.str.fullmatch(r'0+').fillna(False)
    return vacio

norm = pd.DataFrame(index=df.index)
for col in COLUMNAS_COMPLETITUD:
    s = df[col].map(normalizar)
    norm[col] = s.mask(marcar_vacios(s, col), '')

norm[COL_LLAVE] = df[COL_LLAVE]
norm['_orden'] = df['_orden']

llenos_tot = pd.Series(0, index=df.index)
carac_tot = pd.Series(0, index=df.index)
for col in COLUMNAS_COMPLETITUD:
    llenos_tot += (norm[col] != '').astype(int)
    carac_tot += norm[col].str.len()

norm['_llenos'] = llenos_tot
norm['_caracteres'] = carac_tot

id_llenos = pd.Series(0, index=df.index)
id_carac = pd.Series(0, index=df.index)
for col in COLUMNAS_IDENTIDAD:
    id_llenos += (norm[col] != '').astype(int)
    id_carac += norm[col].str.len()

norm['_id_llenos'] = id_llenos
norm['_id_caracteres'] = id_carac

ubi_llenos = pd.Series(0, index=df.index)
for col in COLUMNAS_UBICACION:
    ubi_llenos += (norm[col] != '').astype(int)

norm['_ubi_llenos'] = ubi_llenos
norm['_dir_largo'] = norm[COL_DIRECCION].str.len()
norm['_cor_vacio'] = (norm[COL_CORREO] == '').astype(int)


#5
filas_ini = len(norm)
llaves_ini = norm[COL_LLAVE].nunique()
vacios_ini = {c: int((norm[c] == '').sum()) for c in COLUMNAS_COMPLETITUD}

reporte_antes = pd.DataFrame({
    'METRICA': ['Filas totales', 'Llaves únicas', 'Filas duplicadas', 'Llaves con duplicado',
                'Filas sin correo', 'Filas sin dirección'],
    'VALOR': [filas_ini, llaves_ini, filas_ini - llaves_ini,
              int((norm[COL_LLAVE].duplicated(keep=False).groupby(norm[COL_LLAVE]).any()).sum()),
              vacios_ini[COL_CORREO], vacios_ini[COL_DIRECCION]]
})

print('=== ANTES ===')
display(reporte_antes)


#6
RESCATE_IDENTIDAD = False

norm['_dir_vacio'] = (norm[COL_DIRECCION] == '').astype(int)

base = (
    norm.sort_values([COL_LLAVE, '_cor_vacio', '_llenos', '_caracteres', '_orden'],
                     ascending=[True, True, False, False, True])
        .drop_duplicates(subset=[COL_LLAVE], keep='first')
        .set_index(COL_LLAVE)
)

mejor_id = (
    norm.sort_values([COL_LLAVE, '_id_llenos', '_id_caracteres', '_orden'],
                     ascending=[True, False, False, True])
        .drop_duplicates(subset=[COL_LLAVE], keep='first')
        .set_index(COL_LLAVE)
)

mejor_ubi = (
    norm.sort_values([COL_LLAVE, '_dir_vacio', '_dir_largo', '_ubi_llenos', '_orden'],
                     ascending=[True, True, False, False, True])
        .drop_duplicates(subset=[COL_LLAVE], keep='first')
        .set_index(COL_LLAVE)
)

consolidado = base.copy()

usar_id = mejor_id['_id_llenos'] > base['_id_llenos']
for col in COLUMNAS_IDENTIDAD:
    consolidado.loc[usar_id, col] = mejor_id.loc[usar_id, col]

usar_ubi = ((base['_dir_vacio'] == 1) & (mejor_ubi['_dir_vacio'] == 0)) | \
           ((base['_dir_vacio'] == 0) & (mejor_ubi['_dir_largo'] > base['_dir_largo']))
for col in COLUMNAS_UBICACION:
    consolidado.loc[usar_ubi, col] = mejor_ubi.loc[usar_ubi, col]

columnas_rescate = COLUMNAS_INDEPENDIENTES + COLUMNAS_UBICACION
if RESCATE_IDENTIDAD:
    columnas_rescate = columnas_rescate + COLUMNAS_IDENTIDAD

completados = {}
for col in columnas_rescate:
    aux = norm[[COL_LLAVE, col, '_orden']].copy()
    aux['_vacio'] = (aux[col] == '').astype(int)
    aux['_largo'] = aux[col].str.len()
    mejor = (
        aux.sort_values([COL_LLAVE, '_vacio', '_largo', '_orden'],
                        ascending=[True, True, False, True])
           .drop_duplicates(subset=[COL_LLAVE], keep='first')
           .set_index(COL_LLAVE)[col]
           .reindex(consolidado.index)
    )
    faltante = (consolidado[col] == '') & (mejor != '')
    consolidado.loc[faltante, col] = mejor[faltante]
    completados[col] = int(faltante.sum())

for col in COLUMNAS_COMPLETITUD:
    completados.setdefault(col, 0)

completados['BLOQUE_IDENTIDAD'] = int(usar_id.sum())
completados['BLOQUE_UBICACION'] = int(usar_ubi.sum())

sin_dir_posible = int(
    ((consolidado[COL_DIRECCION] == '') &
     (norm.groupby(COL_LLAVE)['_dir_vacio'].min().reindex(consolidado.index) == 0)).sum()
)

print(f'direcciones rescatadas={completados[COL_DIRECCION]} | bloques ubicación={completados["BLOQUE_UBICACION"]} | casos sin dirección habiendo una disponible={sin_dir_posible}')

#7
resultado = consolidado.reset_index()
resultado = resultado.sort_values('_orden').reset_index(drop=True)

columnas_archivo = [c for c in df.columns if not c.startswith('_') and c != COL_LLAVE]
for col in columnas_archivo:
    if col not in COLUMNAS_COMPLETITUD and col not in COL_LLAVE_PARTES:
        mapa = df.drop_duplicates(subset=[COL_LLAVE], keep='first').set_index(COL_LLAVE)[col]
        resultado[col] = resultado[COL_LLAVE].map(mapa)

resultado[COL_LLAVE_PARTES[0]] = resultado[COL_LLAVE].map(
    df.drop_duplicates(subset=[COL_LLAVE], keep='first').set_index(COL_LLAVE)[COL_LLAVE_PARTES[0]])
resultado[COL_LLAVE_PARTES[1]] = resultado[COL_LLAVE].map(
    df.drop_duplicates(subset=[COL_LLAVE], keep='first').set_index(COL_LLAVE)[COL_LLAVE_PARTES[1]])

columnas_finales = columnas_archivo if not INCLUIR_LLAVE_EN_SALIDA else [COL_LLAVE] + columnas_archivo
salida = resultado[columnas_finales]
salida.to_csv(RUTA_SALIDA, index=False, sep=SEPARADOR_SALIDA, encoding=ENCODING_SALIDA)

vacios_fin = {c: int((resultado[c] == '').sum()) for c in COLUMNAS_COMPLETITUD}

reporte_despues = pd.DataFrame({
    'METRICA': ['Filas resultantes', 'Duplicados eliminados', '% reducción',
                'Filas sin correo', 'Filas sin dirección',
                'Bloques identidad completados', 'Bloques ubicación completados',
                'Correos completados'],
    'VALOR': [len(resultado), filas_ini - len(resultado),
              f'{(filas_ini - len(resultado)) / filas_ini * 100:.1f}%',
              vacios_fin[COL_CORREO], vacios_fin[COL_DIRECCION],
              completados['BLOQUE_IDENTIDAD'], completados['BLOQUE_UBICACION'],
              completados[COL_CORREO]]
})

print('=== DESPUÉS ===')
display(reporte_despues)


#8
detalle = pd.DataFrame({
    'COLUMNA': COLUMNAS_COMPLETITUD,
    'VACIOS_ANTES': [vacios_ini[c] for c in COLUMNAS_COMPLETITUD],
    'VACIOS_DESPUES': [vacios_fin[c] for c in COLUMNAS_COMPLETITUD]
})
detalle['RECUPERADOS'] = detalle['VACIOS_ANTES'] - detalle['VACIOS_DESPUES']
detalle = detalle.sort_values('RECUPERADOS', ascending=False)

detalle.to_csv(RUTA_AUDITORIA, index=False, sep=SEPARADOR_SALIDA, encoding=ENCODING_SALIDA)
display(detalle)

#9
def revisar_llave(valor):
    columnas = COL_LLAVE_PARTES + COLUMNAS_COMPLETITUD
    originales = norm.loc[norm[COL_LLAVE] == valor].copy()
    originales.insert(0, 'FILA', originales['_orden'])
    print(f'--- ORIGINALES ({len(originales)} filas) ---')
    display(originales[['FILA'] + [c for c in COLUMNAS_COMPLETITUD]])
    print('--- CONSOLIDADO ---')
    display(consolidado.loc[[valor], COLUMNAS_COMPLETITUD])

revisar_llave('13123456789')