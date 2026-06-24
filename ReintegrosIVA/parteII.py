#primero 
import pandas as pd, numpy as np, re, unicodedata
from datetime import datetime, timezone, timedelta
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
TOLERANCIA      = 1.0   
COL_CUENTA  = "Cuenta"
COL_DEBITO  = "Debito"
COL_CREDITO = "Crédito"
COL_TIPO    = "Tipo Comprobante"
COL_DESCTRX = "Descripción transacción"
SUB_REPORTES     = "REPORTES_GENERALES"      
SUB_HOJAS_IVA    = "HOJAS_POR_CUENTA_IVA"    
SUB_PROC_TXT     = "_PROCESADOS_TXT"         
SUB_PROC_DATA    = "_PROCESADOS_DATA"        
ARCHIVO_FASE2    = "HOJAS_POR_CUENTA_IVA.xlsx"
print("✅ Configuración lista")

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
    
def mover_archivo(file_id, nuevo_parent):
    """Mueve un archivo a otra carpeta (para sacar los .txt ya procesados del input)."""
    meta = drive.files().get(fileId=file_id, fields="parents").execute()
    prev = ",".join(meta.get("parents", []))
    drive.files().update(fileId=file_id, addParents=nuevo_parent,
                         removeParents=prev, fields="id").execute()

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
    return datetime.now(timezone(timedelta(hours=-5)))

def _sheet_name(x):
    n = re.sub(r'[\\/*?:\[\]]', "_", str(x)).strip()[:31]
    return n or "SIN_NOMBRE"

def escribir_hoja_comprobante(xw, sheet, sub, titulo, ts, cols):
    piv = (sub.groupby([cols["cuenta"], cols["tipo"], cols["desc"]], dropna=False)
              .agg(**{"Suma de Debito":  ("_deb", "sum"),
                      "Suma de Crédito": ("_cre", "sum"),
                      "Suma de Neto":    ("Neto", "sum")})
              .reset_index()
              .rename(columns={cols["cuenta"]: "Cuenta", cols["tipo"]: "Tipo Comprobante",
                               cols["desc"]: "Descripción transacción"}))
    n4 = sub.loc[sub["Extrae"] == "4", "Neto"].sum()
    n2 = sub.loc[sub["Extrae"] == "2", "Neto"].sum(); dif = n4 + n2
    cur = 3
    piv.to_excel(xw, sheet_name=sheet, startrow=cur, index=False); cur += len(piv) + 2
    pd.DataFrame({"Concepto": ["Neto cuentas 4", "Neto cuentas 2", "DIFERENCIA (4+2)"],
                  "Valor": [n4, n2, dif]}).to_excel(xw, sheet_name=sheet, startrow=cur, index=False)
    ws = xw.sheets[sheet]; ws.cell(1, 1, titulo); ws.cell(2, 1, f"Generado: {ts}")
    return {"Neto_4": n4, "Neto_2": n2, "Diferencia": dif,
            "Estado": "OK" if abs(dif) <= TOLERANCIA else ">>> REVISAR"}

def escribir_hoja_iva(xw, sheet, sub, cuenta_iva, ts):
    """sub = datos ya procesados (columnas fijas del CSV de acumulación)."""
    ing = sub[sub["Extrae"] == "4"]    
    piv = (ing.groupby(["Cuenta", "Descripción transacción"], dropna=False)
              .agg(**{"Suma Débito":  ("Debito", "sum"),
                      "Suma Crédito": ("Credito", "sum"),
                      "Suma Neto":    ("Neto", "sum")})
              .reset_index()
              .rename(columns={"Cuenta": "Cuenta Ingreso",
                               "Descripción transacción": "Descripción Transacción"}))
    piv.insert(0, "Cuenta IVA", cuenta_iva)
    piv = piv[["Cuenta IVA", "Cuenta Ingreso", "Descripción Transacción",
               "Suma Débito", "Suma Crédito", "Suma Neto"]]
    n4 = sub.loc[sub["Extrae"] == "4", "Neto"].sum()
    n2 = sub.loc[sub["Extrae"] == "2", "Neto"].sum(); dif = n4 + n2
    cur = 3
    piv.to_excel(xw, sheet_name=sheet, startrow=cur, index=False); cur += len(piv) + 2
    pd.DataFrame({"Concepto": ["Neto cuentas 4", "Neto cuentas 2", "DIFERENCIA (4+2)"],
                  "Valor": [n4, n2, dif]}).to_excel(xw, sheet_name=sheet, startrow=cur, index=False)
    ws = xw.sheets[sheet]; ws.cell(1, 1, f"CUENTA IVA: {cuenta_iva}"); ws.cell(2, 1, f"Generado: {ts}")
    return {"Cuenta IVA": cuenta_iva,
            "Cuentas de ingreso": ", ".join(sorted(ing["Cuenta"].astype(str).unique())) or "(ninguna)",
            "Neto_4": n4, "Neto_2": n2, "Diferencia": dif,
            "Estado": "OK" if abs(dif) <= TOLERANCIA else ">>> REVISAR"}

def escribir_resumen_general(xw, filas, ts):
    pd.DataFrame(filas).to_excel(xw, sheet_name="RESUMEN_GENERAL", startrow=2, index=False)
    xw.sheets["RESUMEN_GENERAL"].cell(1, 1, f"RESUMEN GENERAL — Generado: {ts}")

print("Funciones de Fase 1 y Fase 2 listas")

#sexto

assert FOLDER_ID.strip(), "Falta el FOLDER_ID en el Bloque 2"

f_rep  = obtener_o_crear_subcarpeta(SUB_REPORTES,  FOLDER_ID)
f_ptxt = obtener_o_crear_subcarpeta(SUB_PROC_TXT,  FOLDER_ID)
f_pdat = obtener_o_crear_subcarpeta(SUB_PROC_DATA, FOLDER_ID)

archivos = listar_archivos(FOLDER_ID)
assert archivos, "No hay .txt/.csv en la carpeta. Sube el lote y vuelve a correr."
print(f"Lote actual: {len(archivos)} archivo(s)")

frames = []
for f in archivos:
    local = descargar(f["id"], f"/content/{f['name']}")
    d = leer_txt(local); d.columns = d.columns.str.strip()
    frames.append(d)
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

cuentas2 = sorted(master.loc[master["Extrae"] == "2", cC].astype(str).str.strip().unique())
assert len(cuentas2) != 0, "Este lote NO tiene ninguna cuenta que empiece por 2 (Cuenta IVA)."
assert len(cuentas2) == 1, f"Hay más de una cuenta que empieza por 2: {cuentas2}. Sube un lote por caso."
cuenta_iva = cuentas2[0]
print(f"🔑 Cuenta IVA del lote: {cuenta_iva}")

ts    = _hora_bogota().strftime("%d/%m/%Y %H:%M:%S")
sello = _hora_bogota().strftime("%Y-%m-%d_%H%M%S")

ruta_rep = f"/content/{cuenta_iva}_{sello}.xlsx"; filas = []
with pd.ExcelWriter(ruta_rep, engine="openpyxl") as xw:
    pd.DataFrame().to_excel(xw, sheet_name="RESUMEN_GENERAL", index=False)
    for tipo, gt in master.groupby(cT):
        r = escribir_hoja_comprobante(xw, _sheet_name(tipo), gt,
                          f"COMPROBANTE: {tipo}  |  CUENTA IVA: {cuenta_iva}", ts, cols)
        filas.append({"Comprobante": tipo, **r})
    escribir_resumen_general(xw, filas, ts)
subir(ruta_rep, f_rep)
print(f"Reporte general subido: {cuenta_iva}_{sello}.xlsx")

slim = pd.DataFrame({
    "Cuenta IVA": cuenta_iva,
    "Cuenta": master[cC].astype(str).str.strip(),
    "Descripción transacción": master[cX].astype(str),
    "Debito": master["_deb"], "Credito": master["_cre"],
    "Neto": master["Neto"], "Extrae": master["Extrae"]})
ruta_csv = f"/content/{cuenta_iva}_{sello}.csv"
slim.to_csv(ruta_csv, index=False)
subir(ruta_csv, f_pdat)

for f in archivos:
    mover_archivo(f["id"], f_ptxt)

print(f" Lote procesado y guardado. Sube el siguiente lote y vuelve a correr esta Parte 1.")

#septimo
f_hiva = obtener_o_crear_subcarpeta(SUB_HOJAS_IVA, FOLDER_ID)
f_pdat = obtener_o_crear_subcarpeta(SUB_PROC_DATA, FOLDER_ID)

# Leer TODO lo acumulado
csvs = listar_archivos(_id_de_subcarpeta := f_pdat, extensiones=(".csv",))
# (si tu listar_archivos no recibe un folder distinto, usa la versión de abajo)
proc = drive.files().list(q=f"'{f_pdat}' in parents and trashed=false",
                          fields="files(id,name)").execute().get("files", [])
assert proc, "No hay datos procesados. Corre la Parte 1 con al menos un lote."

partes = []
for f in proc:
    local = descargar(f["id"], f"/content/{f['name']}")
    partes.append(pd.read_csv(local, dtype={"Cuenta IVA": str, "Cuenta": str, "Extrae": str}))
total = pd.concat(partes, ignore_index=True)
print(f"Consolidando {total['Cuenta IVA'].nunique()} cuenta(s) IVA de {len(proc)} lote(s)")

ts = _hora_bogota().strftime("%d/%m/%Y %H:%M:%S")
ruta = f"/content/{ARCHIVO_FASE2}"; resumen = []
with pd.ExcelWriter(ruta, engine="openpyxl") as xw:
    pd.DataFrame().to_excel(xw, sheet_name="RESUMEN_GENERAL", index=False)
    for cuenta_iva, sub in total.groupby("Cuenta IVA"):
        r = escribir_hoja_iva(xw, _sheet_name(cuenta_iva), sub, cuenta_iva, ts)
        resumen.append(r)
    escribir_resumen_general(xw, resumen, ts)
subir(ruta, f_hiva)
print(f"🎉 Listo: '{ARCHIVO_FASE2}' en la subcarpeta '{SUB_HOJAS_IVA}'.")