#Explorar estructura
from google.colab import auth
from googleapiclient.discovery import build
from google.auth import default
import google.auth.transport.requests as tr
import requests, csv, io

carpeta_id = "1_S3A_j5ryHjOzqifZgO15UQ9WMpNto5I"

# Autenticar
auth.authenticate_user()
creds, _ = default()
servicio = build("drive", "v3", credentials=creds)

def token_fresco():
    creds.refresh(tr.Request())
    return creds.token

# Listar archivos de la carpeta
respuesta = servicio.files().list(
    q=f"'{carpeta_id}' in parents and trashed=false",
    fields="files(id, name, size, mimeType, shortcutDetails)",
    includeItemsFromAllDrives=True, supportsAllDrives=True
).execute()

archivos = []
for f in respuesta.get("files", []):
    # Si algún item es un acceso directo, usar el archivo real al que apunta
    if f["mimeType"] == "application/vnd.google-apps.shortcut":
        f["id"] = f["shortcutDetails"]["targetId"]
    if f["name"].lower().endswith(".txt"):
        archivos.append(f)

print(f"Archivos .txt encontrados: {len(archivos)}\n")

def primeras_lineas(file_id, cuantas=5):
    url = f"https://www.googleapis.com/drive/v3/files/{file_id}?alt=media&supportsAllDrives=true"
    r = requests.get(url, headers={"Authorization": f"Bearer {token_fresco()}"}, stream=True)
    lineas = []
    for linea in r.iter_lines(chunk_size=1024*1024):
        if linea is not None:
            lineas.append(linea.decode("latin-1"))
        if len(lineas) >= cuantas:
            break
    r.close()
    return lineas

def detectar_delim(muestra):
    candidatos = ["\t", "|", ";", ","]
    mejor, mejor_conteo = None, 0
    for c in candidatos:
        conteo = muestra[0].count(c) if muestra else 0
        if conteo > mejor_conteo:
            mejor, mejor_conteo = c, conteo
    return mejor  # None = sin delimitador claro (ancho fijo)

delim_por_archivo = {}
for f in archivos:
    muestra = primeras_lineas(f["id"])
    delim = detectar_delim(muestra)
    delim_por_archivo[f["id"]] = delim
    peso = round(int(f.get("size", 0)) / 1e9, 2)
    nombre_delim = {"\t": "TAB", "|": "|", ";": ";", ",": ","}.get(delim, "ancho fijo / sin delimitador")
    print(f"📄 {f['name']}  ({peso} GB)  delimitador: {nombre_delim}")
    if delim:
        columnas = muestra[0].split(delim)
        print(f"   {len(columnas)} columnas. Encabezado posible: {columnas[:10]}")
    print(f"   Primeras líneas: {muestra[:2]}\n")

#2 buscar numero 
import time

numero = "74359610"            # <-- tu número a buscar
BLOQUE = 50 * 1024 * 1024      # tamaño de cada bloque (50 MB)

resultados = []

def descargar_rango(file_id, inicio, fin, token):
    url = f"https://www.googleapis.com/drive/v3/files/{file_id}?alt=media&supportsAllDrives=true"
    headers = {"Authorization": f"Bearer {token}", "Range": f"bytes={inicio}-{fin}"}
    for intento in range(5):
        try:
            r = requests.get(url, headers=headers, timeout=120)
            if r.status_code in (200, 206):
                return r.content
            if r.status_code == 401:
                return None            # token vencido
        except requests.exceptions.RequestException:
            pass
        time.sleep(3)
    raise RuntimeError(f"No se pudo descargar el bloque {inicio}-{fin}")

def buscar_en_archivo(f):
    file_id = f["id"]
    tamano = int(f["size"])
    token = token_fresco()
    resto = b""
    n = 0
    inicio = 0
    while inicio < tamano:
        fin = min(inicio + BLOQUE - 1, tamano - 1)
        datos = descargar_rango(file_id, inicio, fin, token)
        if datos is None:              # renueva token y reintenta el mismo bloque
            token = token_fresco()
            continue
        buffer = resto + datos
        partes = buffer.split(b"\n")
        resto = partes.pop()           # guarda la línea que quedó partida
        for linea_bytes in partes:
            n += 1
            linea = linea_bytes.decode("latin-1").rstrip(";").rstrip("\r")
            if numero in linea:
                resultados.append({"archivo": f["name"], "linea": n,
                                   "texto_original": linea,
                                   "legible": " ".join(linea.split())})
        inicio = fin + 1
    if resto:                          # última línea sin salto final
        n += 1
        linea = resto.decode("latin-1").rstrip(";").rstrip("\r")
        if numero in linea:
            resultados.append({"archivo": f["name"], "linea": n,
                               "texto_original": linea,
                               "legible": " ".join(linea.split())})

for f in archivos:
    print(f"Buscando en {f['name']} ...")
    buscar_en_archivo(f)
    print(f"   listo ({sum(1 for x in resultados if x['archivo']==f['name'])} coincidencias)")

print(f"\n✅ Coincidencias totales: {len(resultados)}\n")
for res in resultados:
    print(f"— {res['archivo']} (línea {res['linea']}):")
    print("   ", res["legible"])
    print()

import pandas as pd
tabla = pd.DataFrame(resultados)
tabla.to_csv("/content/coincidencias.csv", index=False)
print("Guardado en /content/coincidencias.csv")