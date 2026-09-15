# ============================================================
# BLOQUE 1: INSTALACIÓN E IMPORTS
# ============================================================

!pip install pdfplumber -q

from google.colab import drive
drive.mount('/content/drive')

import glob, os
import pdfplumber
import pandas as pd
import gspread
from google.auth import default

print("Entorno listo.")

# ============================================================
# BLOQUE 2: CONFIGURACIÓN
# ============================================================

CARPETA_PDFS = '/content/drive/MyDrive/MotorDeBusqueda/input_pdfs'
SPREADSHEET_ID = '16xGrtTjwQHRkxA4hfuES-RH6WotKgqPDEUbUbOTXob0'
NOMBRE_PESTAÑA = 'texto_detallado'   # <-- ajusta si se llama distinto en tu Sheet

print(f"Carpeta de PDFs: {CARPETA_PDFS}")
print(f"Sheet destino: {SPREADSHEET_ID}")
print(f"Pestaña destino: {NOMBRE_PESTAÑA}")

# ============================================================
# BLOQUE 3: CONEXIÓN AL SHEET Y DETECCIÓN DE PDFS NUEVOS
# ============================================================

# --- Autenticación y conexión al Sheet existente ---
creds, _ = default()
gc = gspread.authorize(creds)
sh = gc.open_by_key(SPREADSHEET_ID)
hoja = sh.worksheet(NOMBRE_PESTAÑA)

# --- Ver qué archivos ya están procesados en el Sheet ---
datos_existentes = hoja.get_all_records()
archivos_ya_procesados = set(fila['nombre_archivo'] for fila in datos_existentes)
print(f'Archivos ya presentes en el Sheet: {len(archivos_ya_procesados)}')

# --- Detectar PDFs en la carpeta que todavía NO están en el Sheet ---
archivos_carpeta = sorted(glob.glob(os.path.join(CARPETA_PDFS, '*.pdf')))
pdfs_nuevos = [a for a in archivos_carpeta if os.path.basename(a) not in archivos_ya_procesados]

print(f'PDFs totales en la carpeta: {len(archivos_carpeta)}')
print(f'PDFs nuevos por procesar: {len(pdfs_nuevos)}')
for a in pdfs_nuevos:
    print('  -', os.path.basename(a))

if not pdfs_nuevos:
    print("\nNo hay PDFs nuevos. Todo lo que está en la carpeta ya existe en el Sheet.")


# ============================================================
# BLOQUE 4: EXTRACCIÓN DE TEXTO Y CARGA AL SHEET
# ============================================================

filas_nuevas = []

if pdfs_nuevos:
    for ruta in pdfs_nuevos:
        nombre_archivo = os.path.basename(ruta)
        try:
            with pdfplumber.open(ruta) as pdf:
                for i, pagina in enumerate(pdf.pages, start=1):
                    filas_nuevas.append([
                        nombre_archivo,
                        ruta,
                        i,
                        pagina.extract_text() or ''
                    ])
            print(f'Procesado: {nombre_archivo}')
        except Exception as e:
            print(f'Error en {nombre_archivo}: {e}')

    # --- Agregar las filas nuevas al FINAL del Sheet, respetando lo que ya hay ---
    if filas_nuevas:
        hoja.append_rows(filas_nuevas, value_input_option='USER_ENTERED')
        print(f'\n✅ Se agregaron {len(filas_nuevas)} fila(s) nueva(s) al Sheet (pestaña "{NOMBRE_PESTAÑA}").')
    else:
        print('\nNo se generaron filas nuevas (revisa si los PDFs nuevos tienen contenido).')
else:
    print('No hay PDFs nuevos que procesar.')