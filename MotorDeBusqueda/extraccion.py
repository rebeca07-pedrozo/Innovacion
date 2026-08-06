#1 
!pip install pdfplumber -q
from google.colab import drive
drive.mount('/content/drive')


#2
CARPETA = '/content/drive/MyDrive/inputs'   # cámbiala si está en una subcarpeta

#3
import glob
for ruta in glob.glob('/content/drive/MyDrive/**/inputs', recursive=True):
    print(ruta)

#4
import glob, os

archivos = sorted(glob.glob(os.path.join(CARPETA, '*.pdf')))
print(f'{len(archivos)} PDFs encontrados:')
for a in archivos:
    print(' -', os.path.basename(a))


#5
import pdfplumber
import pandas as pd

textos, tablas = [], []

for ruta in archivos:
    archivo = os.path.basename(ruta)
    try:
        with pdfplumber.open(ruta) as pdf:
            for i, pagina in enumerate(pdf.pages, start=1):
                textos.append({
                    'archivo': archivo,
                    'pagina': i,
                    'texto': pagina.extract_text() or ''
                })

                for j, tabla in enumerate(pagina.extract_tables(), start=1):
                    if not tabla or len(tabla) < 2:
                        continue
                    df = pd.DataFrame(tabla[1:], columns=tabla[0])
                    df.insert(0, 'archivo', archivo)
                    df.insert(1, 'pagina', i)
                    df.insert(2, 'tabla_n', j)
                    tablas.append(df)
    except Exception as e:
        print(f'Error en {archivo}: {e}')

df_texto = pd.DataFrame(textos)
df_tablas = pd.concat(tablas, ignore_index=True) if tablas else pd.DataFrame()

print(f'{df_texto.archivo.nunique()} archivos | {len(df_texto)} páginas | {len(tablas)} tablas')

#6
# Páginas sin texto → posible PDF escaneado
vacias = df_texto[df_texto.texto.str.strip() == '']
print(f'Páginas sin texto: {len(vacias)} de {len(df_texto)}')

# Ver el contenido de la primera página del primer archivo
print(df_texto.iloc[0].texto[:2000])


#7
SALIDA = '/content/drive/MyDrive/extraccion_pdfs.xlsx'

with pd.ExcelWriter(SALIDA) as w:
    df_texto.to_excel(w, sheet_name='texto', index=False)
    if not df_tablas.empty:
        df_tablas.to_excel(w, sheet_name='tablas', index=False)

print('Guardado en', SALIDA)
