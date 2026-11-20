#Bloque A 
import os
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload


def listar_archivos(folder_id, extensiones=(".txt", ".csv")):
    """Archivos .txt/.csv o de tipo texto plano/csv"""
    q = f"'{folder_id}' in parents and trashed = false"
    res = drive.files().list(q=q, fields="files(id, name, mimeType)", pageSize=1000).execute()
    tipos_ok = ("text/plain", "text/csv", "text/comma-separated-values")
    archivos = []
    for f in res.get("files", []):
        if f["name"].lower().endswith(extensiones) or f.get("mimeType") in tipos_ok:
            archivos.append(f)
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
    """Busca la subcarpeta; si no existe, la crea (no duplica)"""
    q = (f"name = '{nombre}' and '{parent_id}' in parents "
         f"and mimeType = 'application/vnd.google-apps.folder' and trashed = false")
    res = drive.files().list(q=q, fields="files(id)").execute().get("files", [])
    if res:
        return res[0]["id"]
    meta = {"name": nombre, "mimeType": "application/vnd.google-apps.folder", "parents": [parent_id]}
    return drive.files().create(body=meta, fields="id").execute()["id"]


def subir(ruta_local, folder_id):
    """Sube un archivo local a la carpeta de Drive indicada"""
    meta = {"name": os.path.basename(ruta_local), "parents": [folder_id]}
    media = MediaFileUpload(ruta_local, resumable=True)
    drive.files().create(body=meta, media_body=media, fields="id").execute()


def buscar_archivo(nombre, folder_id):
    """Devuelve el archivo MÁS RECIENTE con ese nombre, o None."""
    seguro = nombre.replace("'", "\\'")
    q = f"name = '{seguro}' and '{folder_id}' in parents and trashed = false"
    res = drive.files().list(q=q, fields="files(id, name, modifiedTime)",
                             orderBy="modifiedTime desc").execute().get("files", [])
    return res[0] if res else None


def actualizar_o_crear(nombre, folder_id, ruta_local):
    """Si el archivo ya existe en Drive lo ACTUALIZA; si no, lo crea."""
    media = MediaFileUpload(ruta_local, resumable=True)
    ex = buscar_archivo(nombre, folder_id)
    if ex:
        drive.files().update(fileId=ex["id"], media_body=media).execute()
        return ex["id"]
    return drive.files().create(body={"name": nombre, "parents": [folder_id]},
                                media_body=media, fields="id").execute()["id"]


def mover_archivo(file_id, nuevo_parent):
    """Mueve un .txt ya procesado para dejar limpio el input."""
    meta = drive.files().get(fileId=file_id, fields="parents").execute()
    drive.files().update(fileId=file_id, addParents=nuevo_parent,
                         removeParents=",".join(meta.get("parents", [])), fields="id").execute()


print("Funciones de Drive listas")

#Bloque B
import os, re, unicodedata, openpyxl
import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime, timezone, timedelta


INCLUIR_DETALLE   = True
MAX_FILAS_DETALLE = 900000

SELLO_ACTUAL  = ""
_DETALLES_CSV = []


def _hora_bogota(): return datetime.now(timezone(timedelta(hours=-5)))


def _sheet_name(x):
    n = re.sub(r'[\\/*?:[\]]', "_", str(x)).strip()[:31]
    return n or "SIN_NOMBRE"


def nombre_comprobante(tipo):
    """Nombre ESTABLE del archivo por comprobante: sin acentos, sin espacios, mayúsculas."""
    t = unicodedata.normalize("NFKD", str(tipo)).encode("ascii", "ignore").decode()
    t = re.sub(r'[\\/*?:[\]]', "_", t).strip().upper()
    return (t or "SIN_NOMBRE")[:31]


def _v(x):
    return x.item() if hasattr(x, "item") else x


def _cols_txt(df):
    return [c for c in df.columns if not c.startswith("_")]


def _bloque(xw, sheet, df, startrow, gap=2):
    df.to_excel(xw, sheet_name=sheet, startrow=startrow, index=False)
    return startrow + len(df) + 1 + gap


def _cabe(df, fila_inicial):
    return INCLUIR_DETALLE and (fila_inicial + len(df) + 5) <= MAX_FILAS_DETALLE


def _detalle_aparte(df, nombre):
    ruta = f"/content/DETALLE_{nombre}_{SELLO_ACTUAL}.csv"
    df.to_csv(ruta, index=False, encoding="utf-8-sig")
    _DETALLES_CSV.append(ruta)
    return os.path.basename(ruta)


def escribir_hoja_comprobante(xw, sheet, sub, titulo, ts, cols):
    piv = (sub.groupby([cols["cuenta"], cols["tipo"], cols["desc"]], dropna=False)
              .agg(**{"Suma de Debito":  ("_deb", "sum"),
                      "Suma de Crédito": ("_cre", "sum"),
                      "Suma de Neto":    ("Neto", "sum")})
              .reset_index()
              .rename(columns={cols["cuenta"]: "Cuenta",
                               cols["tipo"]:   "Tipo Comprobante",
                               cols["desc"]:   "Descripción transacción"}))
    n4  = float(sub.loc[sub["Extrae"] == "4", "Neto"].sum())
    n2  = float(sub.loc[sub["Extrae"] == "2", "Neto"].sum())
    dif = n4 + n2
    cuadre = pd.DataFrame({"Concepto": ["Neto cuentas 4", "Neto cuentas 2", "DIFERENCIA (4+2)"],
                           "Valor":    [n4, n2, dif]})
    row = 3
    row = _bloque(xw, sheet, piv, row)
    row = _bloque(xw, sheet, cuadre, row)
    ws  = xw.sheets[sheet]

    det = sub[_cols_txt(sub)]
    if _cabe(det, row):
        ws.cell(row + 1, 1, "DETALLE (líneas del TXT)")
        _bloque(xw, sheet, det, row + 1, gap=0)
    else:
        arch = _detalle_aparte(det, sheet)
        ws.cell(row + 1, 1, f"DETALLE: {len(det):,} líneas — no cabe en Excel. Ver CSV: {arch}")

    ws.cell(1, 1, titulo); ws.cell(2, 1, f"Generado: {ts}")
    return {"Comprobante": sheet, "Neto_4": n4, "Neto_2": n2, "Diferencia": dif,
            "Estado": "OK" if abs(dif) <= TOLERANCIA else ">>> REVISAR"}


def escribir_resumen_general(xw, filas, ts, detalle_txt):
    row = 2
    row = _bloque(xw, "RESUMEN_GENERAL", pd.DataFrame(filas), row)
    ws  = xw.sheets["RESUMEN_GENERAL"]
    ws.cell(1, 1, f"RESUMEN GENERAL — Generado: {ts}")

    if _cabe(detalle_txt, row):
        ws.cell(row + 1, 1, "DETALLE (todas las líneas del TXT)")
        _bloque(xw, "RESUMEN_GENERAL", detalle_txt, row + 1, gap=0)
    else:
        arch = _detalle_aparte(detalle_txt, "LOTE_COMPLETO")
        ws.cell(row + 1, 1, f"DETALLE: {len(detalle_txt):,} líneas — no cabe en Excel. Ver CSV: {arch}")


def escribir_hoja_iva_ws(ws, sub, cuenta_iva, titulo, ts, cols):
    ing = sub[sub["Extrae"] == "4"]
    piv = (ing.groupby([cols["cuenta"], cols["desc"]], dropna=False)
              .agg(**{"Suma Débito":  ("_deb", "sum"),
                      "Suma Crédito": ("_cre", "sum"),
                      "Suma Neto":    ("Neto", "sum")})
              .reset_index()
              .rename(columns={cols["cuenta"]: "Cuenta Ingreso",
                               cols["desc"]:   "Descripción Transacción"}))
    piv.insert(0, "Cuenta IVA", cuenta_iva)
    piv = piv[["Cuenta IVA", "Cuenta Ingreso", "Descripción Transacción",
               "Suma Débito", "Suma Crédito", "Suma Neto"]]
    n4  = float(sub.loc[sub["Extrae"] == "4", "Neto"].sum())
    n2  = float(sub.loc[sub["Extrae"] == "2", "Neto"].sum())
    dif = n4 + n2

    ws.cell(1, 1, titulo); ws.cell(2, 1, f"Generado: {ts}")
    fila = 4
    for r in dataframe_to_rows(piv, index=False, header=True):
        for j, val in enumerate(r, start=1): ws.cell(fila, j, _v(val))
        fila += 1
    fila += 1
    for concepto, valor in [("Concepto", "Valor"), ("Neto cuentas 4", n4),
                            ("Neto cuentas 2", n2), ("DIFERENCIA (4+2)", dif)]:
        ws.cell(fila, 1, concepto); ws.cell(fila, 2, valor); fila += 1
    fila += 1

    det = sub[_cols_txt(sub)]
    if _cabe(det, fila):
        ws.cell(fila, 1, "DETALLE (líneas del TXT)"); fila += 1
        for r in dataframe_to_rows(det, index=False, header=True):
            for j, val in enumerate(r, start=1): ws.cell(fila, j, _v(val))
            fila += 1
    else:
        arch = _detalle_aparte(det, f"{ws.title}{_sheet_name(cuenta_iva)}")
        ws.cell(fila, 1, f"DETALLE: {len(det):,} líneas — no cabe en Excel. Ver CSV: {arch}")


print("Funciones de Paso 1 y Paso 2 listas")

#Bloque C 
assert FOLDER_ID.strip(), "Falta el FOLDER_ID"
try:
    f_rep, f_comp, f_proc
except NameError:
    raise RuntimeError("Primero corre la celda 'Preparar carpetas' (una vez por sesión).")

archivos = listar_archivos(FOLDER_ID)
assert archivos, "No hay .txt/.csv en la carpeta. Sube el lote y vuelve a correr."
print(f"Lote actual: {len(archivos)} archivo(s)")

# 1) Consolidar
frames = []
for f in archivos:
    local = descargar(f["id"], f"/content/{f['name']}")
    d = leer_txt(local); d.columns = d.columns.str.strip(); frames.append(d)
master = pd.concat(frames, ignore_index=True)
print(f"Líneas totales del lote: {len(master):,}")

# 2) Extrae y Neto
cC = _col(master, COL_CUENTA);  cD = _col(master, COL_DEBITO)
cR = _col(master, COL_CREDITO); cT = _col(master, COL_TIPO); cX = _col(master, COL_DESCTRX)
deb, cre = _a_numero(master[cD]), _a_numero(master[cR])
master["Extrae"] = master[cC].astype(str).str.strip().str[0]
master["Neto"] = np.select(
    [master["Extrae"] == "4", master["Extrae"] == "2"],
    [(deb - cre) * IVA, (cre - deb) * SIGNO_CUENTA_2], default=0.0)
master["_deb"], master["_cre"] = deb, cre
cols = {"cuenta": cC, "tipo": cT, "desc": cX}

# 3) Cuenta IVA = la ÚNICA que empieza por 2
cuentas2 = sorted(master.loc[master["Extrae"] == "2", cC].astype(str).str.strip().unique())
assert len(cuentas2) != 0, "El lote NO tiene cuenta que empiece por 2."
assert len(cuentas2) == 1, f"Hay más de una cuenta que empieza por 2: {cuentas2}. Sube un lote por caso."
cuenta_iva = cuentas2[0]
print(f"Cuenta IVA del lote: {cuenta_iva}")

ts    = _hora_bogota().strftime("%d/%m/%Y %H:%M:%S")
sello = _hora_bogota().strftime("%Y-%m-%d%H%M%S")
SELLO_ACTUAL = sello
_DETALLES_CSV.clear()

# PASO 1 — reporte general del lote
ruta_rep = f"/content/{cuenta_iva}_{sello}.xlsx"; filas = []
with pd.ExcelWriter(ruta_rep, engine="openpyxl") as xw:
    pd.DataFrame().to_excel(xw, sheet_name="RESUMEN_GENERAL", index=False)
    for tipo, gt in master.groupby(cT):
        filas.append(escribir_hoja_comprobante(xw, _sheet_name(tipo), gt,
                          f"COMPROBANTE: {tipo}  |  CUENTA IVA: {cuenta_iva}", ts, cols))
    escribir_resumen_general(xw, filas, ts, master[_cols_txt(master)])
subir(ruta_rep, f_rep)
print(f"Paso 1 listo: {cuenta_iva}_{sello}.xlsx")

# PASO 2 — UN Excel por comprobante, una HOJA por cuenta IVA (acumulativo)
for tipo, gt in master.groupby(cT):
    base   = nombre_comprobante(tipo)
    nombre = f"{base}.xlsx"
    local  = f"/content/{nombre}"

    ex = buscar_archivo(nombre, f_comp)
    if ex:                                   # ya existe -> lo bajo y le agrego hoja
        descargar(ex["id"], local)
        wb = openpyxl.load_workbook(local)
        print(f"   {nombre}: ya existía, agregando hoja '{cuenta_iva}'")
    else:                                    # no existe -> libro nuevo
        wb = openpyxl.Workbook(); wb.remove(wb.active)
        print(f"   {nombre}: nuevo, creando hoja '{cuenta_iva}'")

    sh = _sheet_name(cuenta_iva)
    if sh in wb.sheetnames: del wb[sh]       # si esa cuenta ya estaba, la reemplazo
    ws = wb.create_sheet(title=sh)
    escribir_hoja_iva_ws(ws, gt, cuenta_iva,
                         f"CUENTA IVA: {cuenta_iva}  |  COMPROBANTE: {tipo}", ts, cols)
    wb.save(local)
    actualizar_o_crear(nombre, f_comp, local)

print(f"Paso 2 listo: hoja '{cuenta_iva}' en {master[cT].nunique()} comprobante(s)")

# Detalles que no cupieron en Excel -> CSV
for ruta_csv in _DETALLES_CSV:
    subir(ruta_csv, f_rep)
if _DETALLES_CSV:
    print(f"Se subieron {len(_DETALLES_CSV)} CSV de detalle a REPORTE_GENERAL")

# Mover los TXT ya procesados
for f in archivos:
    mover_archivo(f["id"], f_proc)
print("Lote terminado.")

