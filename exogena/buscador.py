Bloque 1 — Autenticación e imports
from google.colab import auth
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from concurrent.futures import ThreadPoolExecutor
import io, threading, pandas as pd

auth.authenticate_user()
drive = build('drive', 'v3', cache_discovery=False)

# un servicio independiente por hilo (googleapiclient no es thread-safe)
local = threading.local()

def servicio():
    if not hasattr(local, 'drive'):
        local.drive = build('drive', 'v3', cache_discovery=False)
    return local.drive

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
def listar_archivos():
    archivos, token = [], None
    while True:
        r = drive.files().list(
            q=f"'{CARPETA}' in parents and trashed = false",
            fields="nextPageToken, files(id, name)",
            pageSize=1000,
            pageToken=token,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
            corpora="allDrives").execute()
        archivos += r.get('files', [])
        token = r.get('nextPageToken')
        if not token:
            break
    return archivos

todos = listar_archivos()
archivos = [a for a in todos if a['name'].startswith(PREFIJO)]

print(f"Total en la carpeta: {len(todos)}")
print(f"Con el prefijo '{PREFIJO}': {len(archivos)}")
for a in todos[:10]:
    print("  ", repr(a['name']))

DIAGNOSTICO
# DIAGNÓSTICO
print("Archivos con el prefijo:", len(archivos))
print([a['name'] for a in archivos[:5]])

a = archivos[0]
buf = io.BytesIO()
d = MediaIoBaseDownload(buf, drive.files().get_media(fileId=a['id']))
ok = False
while not ok:
    _, ok = d.next_chunk()
datos = buf.getvalue()
try:
    texto = datos.decode('utf-8')
except UnicodeDecodeError:
    texto = datos.decode('latin-1')

lineas = texto.splitlines()
print("\nArchivo:", a['name'], "| líneas:", len(lineas))
print("Primera línea:", repr(lineas[0]))
print("Campos con ';':", len(lineas[0].split(';')))

print("\n¿El ID está en el texto?", ID_BUSCADO in texto)
for n, l in enumerate(lineas, 1):
    if ID_BUSCADO in l:
        campos = l.split(';')
        pos = [i for i, c in enumerate(campos) if c.strip() == ID_BUSCADO]
        print(f"línea {n} -> posiciones exactas: {pos} | campo[4] = {repr(campos[4] if len(campos)>4 else None)}")
        break

Bloque 4 — Descargar y buscar en un archivo
def buscar(archivo):
    api = servicio()
    buffer = io.BytesIO()
    peticion = api.files().get_media(fileId=archivo['id'], supportsAllDrives=True)
    bajada = MediaIoBaseDownload(buffer, peticion)
    listo = False
    while not listo:
        _, listo = bajada.next_chunk()

    datos = buffer.getvalue()
    buffer.close()

    if ID_BUSCADO.encode() not in datos:   # filtro rápido
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
resultados = []
with ThreadPoolExecutor(max_workers=4) as pool:
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