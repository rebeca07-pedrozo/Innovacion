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
def procesar_archivo(ruta_txt, carpeta_salida, prefijo):
    df = leer_txt(ruta_txt)
    df.columns = df.columns.str.strip()

    cC = _col(df, COL_CUENTA); cD = _col(df, COL_DEBITO)
    cR = _col(df, COL_CREDITO); cT = _col(df, COL_TIPO)

    deb, cre = _a_numero(df[cD]), _a_numero(df[cR])

    # Columna "Extrae" = primer dígito de la cuenta
    df["Extrae"] = df[cC].astype(str).str.strip().str[0]
    # Columna "Neto":
    #   cuenta 4 -> (Débito - Crédito) * 19%   (queda negativo)
    #   cuenta 2 -> (Crédito - Débito)         (queda positivo, cuadra con la 4)
    df["Neto"] = np.select(
        [df["Extrae"] == "4", df["Extrae"] == "2"],
        [(deb - cre) * IVA, (cre - deb) * SIGNO_CUENTA_2],
        default=0.0)
    df["_deb"], df["_cre"] = deb, cre

    os.makedirs(carpeta_salida, exist_ok=True)
    cols_detalle = [c for c in df.columns if not c.startswith("_")]
    generados, resumen_global = [], []

    for tipo, g in df.groupby(cT):
        piv = (g.groupby(cC)
                 .agg(Debito=("_deb","sum"), Credito=("_cre","sum"), Neto=("Neto","sum"))
                 .reset_index())
        piv["Dig"] = piv[cC].astype(str).str[0]
        n4  = piv.loc[piv["Dig"] == "4", "Neto"].sum()
        n2  = piv.loc[piv["Dig"] == "2", "Neto"].sum()
        dif = n4 + n2

        nom  = (re.sub(r'[\\/*?:\[\]]', "_", str(tipo)).strip() or "SIN_TIPO")
        ruta = os.path.join(carpeta_salida, f"{prefijo}{nom}.xlsx")
        with pd.ExcelWriter(ruta, engine="openpyxl") as xw:
            if INCLUIR_DETALLE:
                g[cols_detalle].to_excel(xw, sheet_name="Detalle", index=False)
            piv.drop(columns="Dig").to_excel(xw, sheet_name="Resumen", index=False)
            pd.DataFrame({
                "Concepto": ["Neto cuentas 4", "Neto cuentas 2", "DIFERENCIA (4+2)"],
                "Valor":    [n4, n2, dif]
            }).to_excel(xw, sheet_name="Resumen", index=False, startrow=len(piv)+3)

        estado = "OK" if abs(dif) <= TOLERANCIA else ">>> REVISAR"
        generados.append(ruta)
        resumen_global.append({"Comprobante": tipo, "Neto_4": n4, "Neto_2": n2,
                               "Diferencia": dif, "Filas": len(g), "Estado": estado})
        print(f"   {nom:6}  filas={len(g):>7}  dif={dif:>15,.2f}   {estado}")

    rg = os.path.join(carpeta_salida, f"{prefijo}RESUMEN_GENERAL.xlsx")
    pd.DataFrame(resumen_global).to_excel(rg, index=False)
    generados.append(rg)
    return generados

print("✅ Listo para procesar")

#sexto
assert FOLDER_ID.strip(), "❌ Falta poner el FOLDER_ID en el Bloque 2"

SALIDA_LOCAL = "/content/salida"
os.makedirs(SALIDA_LOCAL, exist_ok=True)

archivos = listar_archivos(FOLDER_ID)
print(f"📂 Encontré {len(archivos)} archivo(s) en la carpeta:\n")

carpeta_resultados = obtener_o_crear_subcarpeta(NOMBRE_CARPETA_SALIDA, FOLDER_ID)

for f in archivos:
    print(f"📄 Procesando: {f['name']}")
    local = descargar(f["id"], f"/content/{f['name']}")
    prefijo = re.sub(r'\.[^.]+$', '', f["name"]) + "__"      # ej:  CONSOLIDADO__BR.xlsx
    generados = procesar_archivo(local, SALIDA_LOCAL, prefijo)

    print("   ⬆️  Subiendo resultados a Drive...")
    for ruta in generados:
        subir(ruta, carpeta_resultados)
    print()

print(f"🎉 Listo. Revisa la subcarpeta '{NOMBRE_CARPETA_SALIDA}' dentro de tu carpeta de Drive.")