#primero 
import pandas as pd, numpy as np, re, unicodedata
import os
from google.colab import auth
auth.authenticate_user()
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
from datetime import datetime, timezone, timedelta
drive = build('drive', 'v3')    
print("Conectado a Google Drive")


#segundo 

FOLDER_ID = ""          
SEPARADOR       = ","   
IVA             = 0.19  
SIGNO_CUENTA_2  = 1     
INCLUIR_DETALLE = True 
TOLERANCIA      = 1.0   
COL_CUENTA  = "Cuenta"
COL_DEBITO  = "Debito"
COL_CREDITO = "Crédito"
COL_TIPO    = "Tipo Comprobante"
COL_DESCTRX = "Descripción transacción"   
CARPETA_REPORTES = "REPORTES_POR_CUENTA"        
CARPETA_DETALLE = "DETALLE_COMPROBANTES"       
NOMBRE_CARPETA_SALIDA = "SALIDA_COMPROBANTES"  
print("Configuración lista")


#tercero


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

print("Funciones de Drive listas")


#cuarto

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

print("Funciones de cálculo listas")



#quinto

def _hora_bogota():
    return datetime.now(timezone(timedelta(hours=-5)))   # Colombia UTC-5

def _sheet_name(x):
    n = re.sub(r'[\\/*?:\[\]]', "_", str(x)).strip()[:31]   # Excel: máx 31 chars, sin \/?*:[]
    return n or "SIN_NOMBRE"

def construir_pivot(sub, cols):
    """Tablita dinámica con SOLO las 6 columnas + el cuadre (n4, n2, diferencia)."""
    piv = (sub.groupby([cols["cuenta"], cols["tipo"], cols["desc"]], dropna=False)
              .agg(**{"Suma de Debito":  ("_deb", "sum"),
                      "Suma de Crédito": ("_cre", "sum"),
                      "Suma de Neto":    ("Neto", "sum")})
              .reset_index()
              .rename(columns={cols["cuenta"]: "Cuenta",
                               cols["tipo"]:   "Tipo Comprobante",
                               cols["desc"]:   "Descripción transacción"}))
    n4 = sub.loc[sub["Extrae"] == "4", "Neto"].sum()
    n2 = sub.loc[sub["Extrae"] == "2", "Neto"].sum()
    return piv, n4, n2, n4 + n2

def escribir_hoja(xw, sheet, sub, titulo, ts, cols):
    """Escribe una hoja: título + tablita compacta + cuadre. Devuelve el resumen."""
    piv, n4, n2, dif = construir_pivot(sub, cols)
    cur = 3
    piv.to_excel(xw, sheet_name=sheet, startrow=cur, index=False)
    cur += len(piv) + 2
    pd.DataFrame({"Concepto": ["Neto cuentas 4", "Neto cuentas 2", "DIFERENCIA (4+2)"],
                  "Valor":    [n4, n2, dif]}).to_excel(xw, sheet_name=sheet, startrow=cur, index=False)
    ws = xw.sheets[sheet]
    ws.cell(1, 1, titulo)
    ws.cell(2, 1, f"Generado: {ts}")
    return {"Neto_4": n4, "Neto_2": n2, "Diferencia": dif,
            "Estado": "OK" if abs(dif) <= TOLERANCIA else ">>> REVISAR"}
def construir_pivot_detalle(sub, cols, cuenta_iva):
    """Tablita del detalle por comprobante: Cuenta de IVA fija + Cuenta de ingreso por línea."""
    piv = (sub.groupby([cols["cuenta"], cols["tipo"], cols["desc"]], dropna=False)
              .agg(**{"Suma de Debito":  ("_deb", "sum"),
                      "Suma de Crédito": ("_cre", "sum"),
                      "Suma de Neto":    ("Neto", "sum")})
              .reset_index()
              .rename(columns={cols["cuenta"]: "Cuenta de ingreso",
                               cols["tipo"]:   "Tipo Comprobante",
                               cols["desc"]:   "Descripción transacción"}))
    piv.insert(0, "Cuenta de IVA", cuenta_iva)            # columna fija al inicio (se repite a propósito)
    piv = piv[["Cuenta de IVA", "Cuenta de ingreso", "Tipo Comprobante",
               "Descripción transacción", "Suma de Debito", "Suma de Crédito", "Suma de Neto"]]
    n4 = sub.loc[sub["Extrae"] == "4", "Neto"].sum()
    n2 = sub.loc[sub["Extrae"] == "2", "Neto"].sum()
    return piv, n4, n2, n4 + n2

def escribir_hoja_detalle(xw, sheet, sub, titulo, ts, cols, cuenta_iva):
    """Igual que escribir_hoja pero con la tabla de 7 columnas (IVA + ingreso)."""
    piv, n4, n2, dif = construir_pivot_detalle(sub, cols, cuenta_iva)
    cur = 3
    piv.to_excel(xw, sheet_name=sheet, startrow=cur, index=False)
    cur += len(piv) + 2
    pd.DataFrame({"Concepto": ["Neto cuentas 4", "Neto cuentas 2", "DIFERENCIA (4+2)"],
                  "Valor":    [n4, n2, dif]}).to_excel(xw, sheet_name=sheet, startrow=cur, index=False)
    ws = xw.sheets[sheet]
    ws.cell(1, 1, titulo)
    ws.cell(2, 1, f"Generado: {ts}")
    return {"Neto_4": n4, "Neto_2": n2, "Diferencia": dif,
            "Estado": "OK" if abs(dif) <= TOLERANCIA else ">>> REVISAR"}

def escribir_resumen_general(xw, filas, ts):
    """Hoja RESUMEN_GENERAL al inicio del libro (sin columna de correos)."""
    pd.DataFrame(filas).to_excel(xw, sheet_name="RESUMEN_GENERAL", startrow=2, index=False)
    xw.sheets["RESUMEN_GENERAL"].cell(1, 1, f"RESUMEN GENERAL — Generado: {ts}")

print("Funciones listas (tabla compacta + 2 estructuras)")


#sexto

assert FOLDER_ID.strip(), "Falta poner el FOLDER_ID en el Bloque 2"
frames = []
print(" Leyendo archivos...")
for f in listar_archivos(FOLDER_ID):
    local = descargar(f["id"], f"/content/{f['name']}")
    d = leer_txt(local); d.columns = d.columns.str.strip()
    cuenta_rep = d[_col(d, COL_CUENTA)].astype(str).str.strip().mode().iloc[0]
    d["_CuentaReporte"] = cuenta_rep          
    frames.append(d)
    print(f"   {f['name']}  ->  cuenta {cuenta_rep}  ({len(d)} filas)")
assert frames, "No encontré archivos .txt/.csv en la carpeta"

master = pd.concat(frames, ignore_index=True)
cC = _col(master, COL_CUENTA);  cD = _col(master, COL_DEBITO)
cR = _col(master, COL_CREDITO); cT = _col(master, COL_TIPO); cX = _col(master, COL_DESCTRX)
deb, cre = _a_numero(master[cD]), _a_numero(master[cR])
master["Extrae"] = master[cC].astype(str).str.strip().str[0]
master["Neto"] = np.select(
    [master["Extrae"] == "4", master["Extrae"] == "2"],
    [(deb - cre) * IVA, (cre - deb) * SIGNO_CUENTA_2], default=0.0)
master["_deb"], master["_cre"] = deb, cre
cols = {"cuenta": cC, "tipo": cT, "desc": cX}

ts    = _hora_bogota().strftime("%d/%m/%Y %H:%M:%S")
sello = _hora_bogota().strftime("%Y-%m-%d_%H%M")
raiz  = obtener_o_crear_subcarpeta(NOMBRE_CARPETA_SALIDA, FOLDER_ID)
f_rep = obtener_o_crear_subcarpeta(CARPETA_REPORTES, raiz)
f_det = obtener_o_crear_subcarpeta(CARPETA_DETALLE,  raiz)
os.makedirs("/content/salida/rep", exist_ok=True)
os.makedirs("/content/salida/det", exist_ok=True)
print("\n PASO 1 — Reportes por cuenta")
for cuenta, gc in master.groupby("_CuentaReporte"):
    nombre = f"{cuenta}_{sello}.xlsx"          
    ruta = f"/content/salida/rep/{nombre}"; filas = []
    with pd.ExcelWriter(ruta, engine="openpyxl") as xw:
        pd.DataFrame().to_excel(xw, sheet_name="RESUMEN_GENERAL", index=False)
        for tipo, gt in gc.groupby(cT):
            r = escribir_hoja(xw, _sheet_name(tipo), gt,
                              f"COMPROBANTE: {tipo}  |  CUENTA: {cuenta}", ts, cols)
            filas.append({"Comprobante": tipo, "Hoja": _sheet_name(tipo), **r})
        escribir_resumen_general(xw, filas, ts)
    subir(ruta, f_rep)
    print(f"   {nombre}  ({len(filas)} comprobantes)")
print("\n PASO 2 — Detalle por comprobante")
for tipo, gt in master.groupby(cT):
    nombre = f"{_sheet_name(tipo)}.xlsx"         
    ruta = f"/content/salida/det/{nombre}"; filas = []
    with pd.ExcelWriter(ruta, engine="openpyxl") as xw:
        pd.DataFrame().to_excel(xw, sheet_name="RESUMEN_GENERAL", index=False)
        for cuenta_iva, gc in gt.groupby("_CuentaReporte"):
            r = escribir_hoja_detalle(xw, _sheet_name(cuenta_iva), gc,
                              f"CUENTA IVA: {cuenta_iva}  |  COMPROBANTE: {tipo}", ts, cols, cuenta_iva)
            ingresos = sorted(gc.loc[gc["Extrae"] == "4", cC].astype(str).str.strip().unique())
            cuentas_ingreso = ", ".join(ingresos) if ingresos else "(sin cuentas de ingreso)"
            filas.append({"Cuenta IVA": cuenta_iva,
                          "Cuentas de ingreso": cuentas_ingreso,
                          "Hoja": _sheet_name(cuenta_iva), **r})
        escribir_resumen_general(xw, filas, ts)
    subir(ruta, f_det)
    print(f"    {nombre}  ({len(filas)} cuentas de IVA)")
