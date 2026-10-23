Bloque 1 — Autenticación e imports
python
from google.colab import auth
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from concurrent.futures import ThreadPoolExecutor
import io, pandas as pd

auth.authenticate_user()
drive = build('drive', 'v3')
Bloque 2 — Configuración (lo único que cambias)
python
ID_BUSCADO = "1000"          # <-- número de ID a buscar
CARPETA = "16P-8IAXVNKCOfwEMj0tEGT05ZnCa4wfw"
PREFIJO = "CDTC_EXOGENA_"
POSICION_ID = 4              # campo 5 (empieza en 0)

CAMPOS = ["be_fecha","be_banco","Titular_1","Tipo ID","Numero de ID","Tipo Persona",
          "Estado del titulo","Oficina de radicacion","Sucursal","Marca tipo cdt",
          "Tipo de movimiento","be_saldo_ini_cont","be_valor_inversion_acum","be_int_liq",
          "Valor retencion en la fuente","be_saldo_final","be_valor_int_cau",
          "Secuencia producto","Numero producto"]
Bloque 3 — Listar los TXT de la carpeta
python
def listar_archivos():
    archivos, token = [], None
    while True:
        r = drive.files().list(
            q=f"'{CARPETA}' in parents and trashed = false",
            fields="nextPageToken, files(id, name)",
            pageSize=1000, pageToken=token).execute()
        archivos += r.get('files', [])
        token = r.get('nextPageToken')
        if not token:
            break
    return [a for a in archivos if a['name'].startswith(PREFIJO)]

archivos = listar_archivos()
print(f"Archivos encontrados: {len(archivos)}")
Bloque 4 — Descargar y buscar en un archivo
python
def buscar(archivo):
    buffer = io.BytesIO()
    bajada = MediaIoBaseDownload(buffer, drive.files().get_media(fileId=archivo['id']))
    listo = False
    while not listo:
        _, listo = bajada.next_chunk()

    datos = buffer.getvalue()
    if ID_BUSCADO.encode() not in datos:   # filtro rápido: si no está, ni recorre
        return []

    try:
        texto = datos.decode('utf-8')
    except UnicodeDecodeError:
        texto = datos.decode('latin-1')

    filas = []
    for num, linea in enumerate(texto.splitlines(), start=1):
        campos = linea.strip().split(';')
        if len(campos) > POSICION_ID and campos[POSICION_ID].strip() == ID_BUSCADO:
            filas.append([archivo['name'], num] + campos[:len(CAMPOS)])
    return filas
Bloque 5 — Ejecutar la búsqueda (en paralelo)
python
resultados = []
with ThreadPoolExecutor(max_workers=8) as pool:
    for filas in pool.map(buscar, archivos):
        resultados += filas

tabla = pd.DataFrame(resultados, columns=["Archivo", "Linea"] + CAMPOS)
print(f"Coincidencias: {len(tabla)}  |  Archivos con el ID: {tabla['Archivo'].nunique() if len(tabla) else 0}")
tabla.head(20)
Bloque 6 — Exportar el reporte
python
salida = f"reporte_{ID_BUSCADO}.xlsx"
tabla.to_excel(salida, index=False)

from google.colab import files
files.download(salida)

Dos notas: el filtro rápido del bloque 4 evita recorrer línea por línea los archivos donde el ID no aparece, y la comparación es exacta sobre el campo 5, así que no te trae falsos positivos de otros campos (por ejemplo un valor que contenga esos dígitos).