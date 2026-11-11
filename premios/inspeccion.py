# 1. Montar Drive / autenticar
from google.colab import auth
auth.authenticate_user()

from googleapiclient.discovery import build
servicio = build('drive', 'v3')

# 2. Listar todo recursivamente
id_central = 'PON_AQUI_EL_ID'  # ID de la carpeta central
resultados = []

def listar(id_carpeta, ruta):
    token = None
    while True:
        r = servicio.files().list(
            q=f"'{id_carpeta}' in parents and trashed=false",
            fields="nextPageToken, files(id, name, mimeType)",
            pageSize=1000,
            pageToken=token,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()

        for item in r.get('files', []):
            es_carpeta = item['mimeType'] == 'application/vnd.google-apps.folder'
            resultados.append({
                'ruta': ruta,
                'nombre': item['name'],
                'tipo': 'Carpeta' if es_carpeta else 'Archivo',
                'id': item['id']
            })
            if es_carpeta:
                listar(item['id'], ruta + '/' + item['name'])

        token = r.get('nextPageToken')
        if not token:
            break

listar(id_central, '')
print(f"Total encontrado: {len(resultados)}")

# 3. Ver resultado en tabla
import pandas as pd
df = pd.DataFrame(resultados)
df