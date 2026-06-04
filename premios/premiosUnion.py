!pip install xlsxwriter -q

import pandas as pd
import io, os, locale
from datetime import datetime, timedelta

from google.colab import auth
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

# === 1. Autenticación (sin montar Drive) ===
auth.authenticate_user()
service = build('drive', 'v3')

# === 2. IDs de las carpetas ===
ID_ENTRADA = "1vCHEky962KIBuXZw_trg6Ws3XvVHmOac"
ID_SALIDA  = "1YFmxbp94ePNMxLo2C2y5tOPasKVN481J"

MIME_XLSX  = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
MIME_SHEET = "application/vnd.google-apps.spreadsheet"

# === 3. Listar archivos dentro de la carpeta de entrada (por ID) ===
query = f"'{ID_ENTRADA}' in parents and trashed = false"
resultados, page_token = [], None
while True:
    resp = service.files().list(
        q=query,
        fields="nextPageToken, files(id, name, mimeType)",
        pageSize=1000, pageToken=page_token,
        supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()
    resultados.extend(resp.get("files", []))
    page_token = resp.get("nextPageToken")
    if not page_token:
        break

# Solo hojas de cálculo (xlsx nativos o Google Sheets), ignorando temporales
archivos = [f for f in resultados
            if f["mimeType"] in (MIME_XLSX, MIME_SHEET)
            and not f["name"].startswith("~$")]

print(f"Archivos encontrados en Entrada: {len(archivos)}")
for f in archivos:
    print("  ->", f["name"], "|", f["mimeType"])

# === 4. Función para descargar (exporta si es Google Sheet) ===
def descargar_como_xlsx(f):
    if f["mimeType"] == MIME_SHEET:
        request = service.files().export_media(fileId=f["id"], mimeType=MIME_XLSX)
    else:
        request = service.files().get_media(fileId=f["id"])
    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buffer.seek(0)
    return buffer

# === 5. Consolidar datos ===
df_consolidado = pd.DataFrame()
if not archivos:
    print("No se encontraron hojas de cálculo en la carpeta de entrada.")
else:
    for f in archivos:
        try:
            contenido = descargar_como_xlsx(f)
            df = pd.read_excel(contenido, usecols=range(22))  # Columnas A hasta V (22)
            df_consolidado = pd.concat([df_consolidado, df], ignore_index=True)
        except Exception as e:
            print(f"Error leyendo {f['name']}: {e}")

    if not df_consolidado.empty and df_consolidado.shape[1] >= 18:
        # Columna usada para el valor (índice 18-1 = 17), a numérico
        df_consolidado.iloc[:, 18-1] = pd.to_numeric(df_consolidado.iloc[:, 18-1], errors='coerce')

        # === 6. Tarifa de Retención ===
        df_consolidado['Tarifa Retención (%)'] = 0.0
        df_consolidado.loc[(df_consolidado.iloc[:, 10-1] == 'SI') & (df_consolidado.iloc[:, 18-1] > 2514000), 'Tarifa Retención (%)'] = 20.0
        df_consolidado.loc[(df_consolidado.iloc[:, 10-1] == 'NO') & (df_consolidado.iloc[:, 18-1] > 1414000), 'Tarifa Retención (%)'] = 3.5

        # === 7. Valor de la Retención ===
        df_consolidado['Valor Retención'] = df_consolidado.iloc[:, 18-1] * (df_consolidado['Tarifa Retención (%)'] / 100)

        # === 8. Nombre del archivo con el mes anterior ===
        try:
            locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
        except locale.Error:
            try:
                locale.setlocale(locale.LC_TIME, 'es_ES')
            except locale.Error:
                print("Advertencia: no se pudo establecer el locale a español.")

        hoy = datetime.now()
        mes_anterior = hoy.replace(day=1) - timedelta(days=1)
        nombre_mes = mes_anterior.strftime('%B').lower()
        nombre_archivo = f"Consolidado_con_Retencion-{nombre_mes}.xlsx"
        ruta_local = f"/content/{nombre_archivo}"

        # === 9. Generar el Excel con varias hojas (en local) ===
        with pd.ExcelWriter(ruta_local, engine='xlsxwriter') as writer:
            df_consolidado.to_excel(writer, sheet_name='Consolidado General', index=False)

            df_consolidado[df_consolidado['Tarifa Retención (%)'] == 20.0]\
                .to_excel(writer, sheet_name='Premios20%', index=False)

            df_consolidado[df_consolidado['Tarifa Retención (%)'] == 3.5]\
                .to_excel(writer, sheet_name='Otros ingresos tributarios3.5%', index=False)

            df_consolidado[(df_consolidado.iloc[:, 10-1].isin(['NO','No','no'])) & (df_consolidado['Tarifa Retención (%)'] == 0)]\
                .to_excel(writer, sheet_name='Otros Ingresos Sin retencion', index=False)

            df_consolidado[(df_consolidado.iloc[:, 10-1].isin(['SI','Si','si'])) & (df_consolidado['Tarifa Retención (%)'] == 0)]\
                .to_excel(writer, sheet_name='Premios Sin Retencion', index=False)

        # === 10. Subir el resultado a la carpeta de Salida (por ID) ===
        metadata = {"name": nombre_archivo, "parents": [ID_SALIDA]}
        media = MediaFileUpload(ruta_local, mimetype=MIME_XLSX, resumable=True)
        subido = service.files().create(
            body=metadata, media_body=media,
            fields="id, webViewLink", supportsAllDrives=True,
        ).execute()

        print(f"\n✅ Archivo subido a la carpeta de Salida.")
        print(f"   Link: {subido.get('webViewLink')}")
    else:
        print("df_consolidado está vacío o no tiene al menos 18 columnas después de consolidar.")