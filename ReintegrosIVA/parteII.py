#primero 
# Autoriza Colab para que pueda leer/escribir en TU Google Drive.
# Al correrlo, sale una ventana: elige tu cuenta y dale "Permitir".
from google.colab import auth
auth.authenticate_user()

from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

drive = build('drive', 'v3')   # cliente de Google Drive
print("✅ Conectado a Google Drive")

#segundo 
# ===================== AQUÍ PONES TU INFORMACIÓN =====================

FOLDER_ID = ""          # <<<<<< PEGA AQUÍ EL ID DE TU CARPETA DE DRIVE
                        #        (donde dejas los .txt que descargas)

SEPARADOR       = ","   # separador del txt:  ","   "\t"(tab)   "|"   ";"
IVA             = 0.19  # tarifa de IVA
SIGNO_CUENTA_2  = 1     # 1 = (Créd−Déb), cuadra a 0.  Pon -1 si el IVA te llega en Débito
INCLUIR_DETALLE = True  # False = solo hoja Resumen (si un comprobante es enorme)
TOLERANCIA      = 1.0   # diferencia en $ que aceptas como "cuadrado"

# Nombres de columnas en tu txt (déjalos así si no cambian los encabezados)
COL_CUENTA  = "Cuenta"
COL_DEBITO  = "Debito"
COL_CREDITO = "Crédito"
COL_TIPO    = "Tipo Comprobante"

NOMBRE_CARPETA_SALIDA = "SALIDA_COMPROBANTES"   # subcarpeta donde quedan los Excel
# =====================================================================
print("✅ Configuración lista")

#tercero
import os

def listar_archivos(folder_id, extensiones=(".txt", ".csv")):
    """Devuelve los archivos de la carpeta que terminen en .txt o .csv"""
    q = f"'{folder_id}' in parents and trashed = false"
    res = drive.files().list(q=q, fields="files(id, name, mimeType)",
                             pageSize=1000).execute()
    archivos = [f for f in res.get("files", [])
                if f["name"].lower().endswith(extensiones)]
    return archivos

def descargar(file_id, destino_local):
    """Descarga un archivo de Drive al disco temporal de Colab"""
    req = drive.files().get_media(fileId=file_id)
    with open(destino_local, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, req)
        done = False
        while not done:
            _, done = downloader.next_chunk()
    return destino_local

def obtener_o_crear_subcarpeta(nombre, parent_id):
    """Busca la subcarpeta de salida; si no existe, la crea (no duplica)"""
    q = (f"name = '{nombre}' and '{parent_id}' in parents "
         f"and mimeType = 'application/vnd.google-apps.folder' and trashed = false")
    res = drive.files().list(q=q, fields="files(id)").execute().get("files", [])
    if res:
        return res[0]["id"]
    meta = {"name": nombre, "mimeType": "application/vnd.google-apps.folder",
            "parents": [parent_id]}
    return drive.files().create(body=meta, fields="id").execute()["id"]

def subir(ruta_local, folder_id):
    """Sube un archivo local a la carpeta de Drive indicada"""
    meta = {"name": os.path.basename(ruta_local), "parents": [folder_id]}
    media = MediaFileUpload(ruta_local, resumable=True)
    drive.files().create(body=meta, media_body=media, fields="id").execute()

print("✅ Funciones de Drive listas")

#cuarto
import pandas as pd, numpy as np, re, unicodedata

def _norm(s):
    return unicodedata.normalize("NFKD", str(s)).encode("ascii","ignore").decode().lower().strip()

def _col(df, nombre):
    """Encuentra la columna aunque cambien tildes o mayúsculas"""
    obj = _norm(nombre)
    for c in df.columns:
        if _norm(c) == obj:
            return c
    raise KeyError(f"No encuentro la columna '{nombre}'. Disponibles: {list(df.columns)}")

def _a_numero(serie):
    """Convierte texto a número: tolera '1.234,56', '1234.56', '-', vacíos"""
    def conv(x):
        x = str(x).strip().replace(" ", "")
        if x in ("", "-", "nan", "na", "none", "NA", "None"):
            return 0.0
        p, c = "." in x, "," in x
        if p and c:   x = x.replace(".", "").replace(",", ".")   # 1.234.567,89
        elif c:       x = x.replace(",", ".")                    # 1234,56
        try:    return float(x)
        except: return 0.0
    return serie.map(conv)

def leer_txt(ruta):
    try:    return pd.read_csv(ruta, sep=SEPARADOR, dtype=str, encoding="utf-8")
    except UnicodeDecodeError:
            return pd.read_csv(ruta, sep=SEPARADOR, dtype=str, encoding="latin-1")

print("✅ Funciones de cálculo listas")

#quinto
from datetime import datetime, timezone, timedelta

def _hora_bogota():
    # Colombia = UTC-5 (sin horario de verano)
    return datetime.now(timezone(timedelta(hours=-5)))

def _sheet_name(tipo):
    # Excel no permite \ / ? * : [ ] en nombres de hoja, y máximo 31 caracteres
    n = re.sub(r'[\\/*?:\[\]]', "_", str(tipo)).strip()[:31]
    return n or "SIN_TIPO"

def procesar_archivo(ruta_txt, ruta_xlsx):
    df = leer_txt(ruta_txt)
    df.columns = df.columns.str.strip()

    cC = _col(df, COL_CUENTA); cD = _col(df, COL_DEBITO)
    cR = _col(df, COL_CREDITO); cT = _col(df, COL_TIPO)
    deb, cre = _a_numero(df[cD]), _a_numero(df[cR])

    df["Extrae"] = df[cC].astype(str).str.strip().str[0]
    df["Neto"] = np.select(
        [df["Extrae"] == "4", df["Extrae"] == "2"],
        [(deb - cre) * IVA, (cre - deb) * SIGNO_CUENTA_2],
        default=0.0)
    df["_deb"], df["_cre"] = deb, cre

    cols_detalle = [c for c in df.columns if not c.startswith("_")]
    ts = _hora_bogota().strftime("%d/%m/%Y %H:%M:%S")
    resumen_global = []

    with pd.ExcelWriter(ruta_xlsx, engine="openpyxl") as xw:
        # Reservamos la 1ª hoja para el RESUMEN_GENERAL (se llena al final)
        pd.DataFrame().to_excel(xw, sheet_name="RESUMEN_GENERAL", index=False)

        for tipo, g in df.groupby(cT):
            piv = (g.groupby(cC)
                     .agg(Debito=("_deb","sum"), Credito=("_cre","sum"), Neto=("Neto","sum"))
                     .reset_index())
            piv["Dig"] = piv[cC].astype(str).str[0]
            n4  = piv.loc[piv["Dig"] == "4", "Neto"].sum()
            n2  = piv.loc[piv["Dig"] == "2", "Neto"].sum()
            dif = n4 + n2
            piv = piv.drop(columns="Dig")

            sh  = _sheet_name(tipo)
            cur = 3                                   # filas 1-3 = títulos

            # --- RESUMEN (arriba) ---
            piv.to_excel(xw, sheet_name=sh, startrow=cur, index=False)
            cur += len(piv) + 2

            # --- CUADRE (Neto4 / Neto2 / Diferencia) ---
            chk = pd.DataFrame({
                "Concepto": ["Neto cuentas 4", "Neto cuentas 2", "DIFERENCIA (4+2)"],
                "Valor":    [n4, n2, dif]})
            chk.to_excel(xw, sheet_name=sh, startrow=cur, index=False)
            cur += len(chk) + 2

            # --- DETALLE (abajo) ---
            fila_detalle = cur
            if INCLUIR_DETALLE:
                g[cols_detalle].to_excel(xw, sheet_name=sh, startrow=fila_detalle + 1, index=False)

            # Títulos y rótulos (openpyxl usa filas en base 1)
            ws = xw.sheets[sh]
            ws.cell(1, 1, f"COMPROBANTE: {tipo}")
            ws.cell(2, 1, f"Generado: {ts}")
            if INCLUIR_DETALLE:
                ws.cell(fila_detalle + 1, 1, "DETALLE")

            estado = "OK" if abs(dif) <= TOLERANCIA else ">>> REVISAR"
            resumen_global.append({
                "Comprobante": tipo, "Hoja": sh, "Neto_4": n4, "Neto_2": n2,
                "Diferencia": dif, "Filas": len(g), "Estado": estado,
                "Correos destinatarios": ""        # <-- aquí mapeas a quién va cada comprobante
            })
            print(f"   {sh:6}  filas={len(g):>7}  dif={dif:>15,.2f}   {estado}")

        # --- Hoja RESUMEN_GENERAL (al inicio del libro) ---
        pd.DataFrame(resumen_global).to_excel(xw, sheet_name="RESUMEN_GENERAL",
                                              startrow=2, index=False)
        xw.sheets["RESUMEN_GENERAL"].cell(1, 1, f"RESUMEN GENERAL — Generado: {ts}")

    return resumen_global

print("Listo para procesar (un libro por archivo, con hojas por comprobante)")
#sexto

assert FOLDER_ID.strip(), "Falta poner el FOLDER_ID en el Bloque 2"

SALIDA_LOCAL = "/content/salida"
os.makedirs(SALIDA_LOCAL, exist_ok=True)

archivos = listar_archivos(FOLDER_ID)
print(f"Encontré {len(archivos)} archivo(s) en la carpeta:\n")

carpeta_resultados = obtener_o_crear_subcarpeta(NOMBRE_CARPETA_SALIDA, FOLDER_ID)
sello = _hora_bogota().strftime("%Y-%m-%d_%H%M")     # fecha y hora para el nombre

for f in archivos:
    print(f"Procesando: {f['name']}")
    local = descargar(f["id"], f"/content/{f['name']}")

    base = re.sub(r'\.[^.]+$', '', f["name"])         # nombre sin extensión
    nombre_salida = f"COMPROBANTES_{base}_{sello}.xlsx"
    ruta_salida   = os.path.join(SALIDA_LOCAL, nombre_salida)

    procesar_archivo(local, ruta_salida)

    print("Subiendo a Drive...")
    subir(ruta_salida, carpeta_resultados)
    print(f"{nombre_salida}\n")

print(f" Listo. Busca '{NOMBRE_CARPETA_SALIDA}' en tu carpeta de Drive.")