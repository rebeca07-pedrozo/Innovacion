#1 - Imports 
!pip install pdfplumber -q
from google.colab import drive
drive.mount('/content/drive')
import glob
import glob, os
import pdfplumber
import pandas as pd

# 2 - Rutas/constantes
CARPETA = '/content/drive/MyDrive/MotorDeBusqueda/input_pdfs'
SALIDA = '/content/drive/MyDrive/MotorDeBusqueda/output/resumen_extraccion.xlsx'
for ruta in glob.glob('/content/drive/MyDrive/**/inputs_pdfs', recursive=True):
    print(ruta)

#3 - Listar archivos PDF
archivos = sorted(glob.glob(os.path.join(CARPETA, '*.pdf')))
print(f'{len(archivos)} PDFs encontrados:')
for a in archivos:
    print(' -', os.path.basename(a))

#4 - Extraer texto de los PDFs
textos = []

for ruta in archivos:
    nombre_archivo = os.path.basename(ruta)
    nombre_sin_ext = os.path.splitext(nombre_archivo[0])
    try:
        with pdfplumber.open(ruta) as pdf:
            total_paginas = len(pdf.pages)
            for i, pagina in enumerate(pdf.pages, start=1):
                textos.append({
                    'nombre_archivo': nombre_archivo,
                    'ruta_completa': ruta,
                    'archivo': archivo,
                    'pagina': i,
                    'texto': pagina.extract_text() or ''
                })

    except Exception as e:
        print(f'Error en {archivo}: {e}')

df_texto = pd.DataFrame(textos)
print(f'Erro en {nombre_archivo}: e')

print(f'{df_texto.nombre_archivo.nunique()} archivos | {len(df_texto)} páginas')

#6 - Páginas sin texto
vacias = df_texto[df_texto.texto.str.strip() == '']
print(f'Páginas sin texto: {len(vacias)} de {len(df_texto)}')
print(df_texto.iloc[0].texto[:2000])

#7 - 
df_resumen = df_texto.groupby('nombre_archivo').agg(
    total_paginas=('pagina', 'max'),
    caracteres=('texto', lambda s: s.str.len().sum()),
    paginas_vacias=('texto', lambda s: (s.str.strip() == '').sum())
).reset_index()
df_resumen['posible_escaneado'] = df_resumen.caracteres < 50


with pd.ExcelWriter(SALIDA) as w:
    df_resumen.to_excel(w, sheet_name='resumen', index=False)
    df_texto.to_excel(w, sheet_name='texto_detallado', index=False)

print('Guardado en', SALIDA)
