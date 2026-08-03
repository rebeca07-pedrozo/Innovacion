#  1
!pip install xlsxwriter psutil -q

import psutil, gc, os

def ram(etiqueta=""):
    """Muestra RAM usada. Llámala entre bloques para ver dónde se dispara."""
    p = psutil.Process(os.getpid())
    usada = p.memory_info().rss / 1024**3
    v = psutil.virtual_memory()
    total, disp = v.total / 1024**3, v.available / 1024**3
    barra = "#" * int(20 * usada / total) + "." * (20 - int(20 * usada / total))
    alerta = "  <-- CRÍTICO" if disp < 1.5 else ""
    print(f"  RAM [{barra}] {usada:.2f} / {total:.1f} GB  "
          f"| libre {disp:.2f} GB {etiqueta}{alerta}")
    return disp

def liberar(*objetos):
    """Borra objetos y fuerza recolección de basura."""
    for o in objetos:
        try:
            del o
        except Exception:
            pass
    gc.collect()

ram("inicio")

#2

import pandas as pd
import numpy as np
import glob, os, re, shutil, time, gc, unicodedata
from datetime import datetime, timedelta

from google.colab import drive
drive.mount('/content/drive')

pd.set_option("mode.copy_on_write", True)   # evita copias intermedias
print("Librerías cargadas.")
ram("post-imports")


#3
CARPETA_BASE       = "/content/drive/My Drive/Optimizacion-Premios/2026/"
CARPETA_ENTRADA    = os.path.join(CARPETA_BASE, "Entrada")
CARPETA_SALIDA     = os.path.join(CARPETA_BASE, "Salida")
CARPETA_PROCESADOS = os.path.join(CARPETA_BASE, "Procesados")
CARPETA_TEMP       = "/content/temp_premios"      # disco local, NO Drive

ANIO_GRAVABLE = 2026

# === Memoria ===
MODO_BAJO_CONSUMO = True    # lectura por streaming + checkpoints en disco
UMBRAL_RAM_GB     = 1.5     # si baja de esto, vuelca a disco y libera
CHECKPOINT        = True    # guarda avance en parquet por si se cae la sesión

MOVER_PROCESADOS = False
FILA_ENCABEZADO  = None
COL_VALOR = COL_PREMIO = COL_MES = None

UVT = 52_374
TOPE_PREMIOS, TOPE_OTROS = 48 * UVT, 27 * UVT
TARIFA_PREMIOS, TARIFA_OTROS = 20.0, 3.5

for c in (CARPETA_ENTRADA, CARPETA_SALIDA, CARPETA_PROCESADOS, CARPETA_TEMP):
    os.makedirs(c, exist_ok=True)

print(f"Modo bajo consumo: {MODO_BAJO_CONSUMO}   |   Checkpoints: {CHECKPOINT}")
print(f"Temp local: {CARPETA_TEMP}  (más rápido que Drive)")

#4

def sin_tildes(t):
    return "".join(c for c in unicodedata.normalize("NFD", str(t))
                   if unicodedata.category(c) != "Mn")

def limpiar_encabezado(t):
    return re.sub(r"\s+", " ", sin_tildes(t).upper().strip())

def letra_a_indice(l):
    i = 0
    for c in str(l).strip().upper():
        i = i * 26 + (ord(c) - 64)
    return i - 1

def indice_a_letra(i):
    s = ""; i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def normalizar_numero(v):
    """1.234.567,89 | 1,234,567.89 | $ 1.234.567 | (1.234) | 1234"""
    if isinstance(v, (int, float, np.integer, np.floating)):
        return np.nan if pd.isna(v) else float(v)
    if v is None:
        return np.nan
    s = str(v).replace("\xa0", " ").replace("\u2212", "-").replace("–", "-")
    s = re.sub(r"[^\d,\.\-()]", "", s).strip()
    if s in ("", "-", ".", ","):
        return np.nan
    neg = (s.startswith("(") and s.endswith(")")) or s.startswith("-")
    s = s.strip("()").lstrip("-")
    hc, hp = "," in s, "." in s
    if hc and hp:
        s = (s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".")
             else s.replace(",", ""))
    elif hc:
        p = s.split(",")
        s = s.replace(",", ".") if (len(p) == 2 and len(p[1]) != 3) else s.replace(",", "")
    elif hp:
        p = s.split(".")
        if not (len(p) == 2 and len(p[1]) != 3):
            s = s.replace(".", "")
    try:
        x = float(s)
    except ValueError:
        return np.nan
    return -x if neg else x


_SI = {"SI", "S", "1", "1.0", "X", "TRUE", "VERDADERO", "SIP", "SII"}
_NO = {"NO", "N", "0", "0.0", "FALSE", "FALSO", "NA", "N/A", "NINGUNO"}

def normalizar_si_no(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = sin_tildes(v).upper().strip().replace(".", "").replace(" ", "")
    return "SI" if s in _SI else "NO" if s in _NO else s


MESES_NOMBRE = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
                7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",
                11:"Noviembre",12:"Diciembre"}

_ALIAS_MES = {}
for _n, _nom in MESES_NOMBRE.items():
    _b = sin_tildes(_nom).upper()
    _ALIAS_MES[_b] = _ALIAS_MES[_b[:3]] = _ALIAS_MES[f"{_n:02d}"] = _n
_ALIAS_MES.update({"SETIEMBRE":9,"SET":9,"SEPT":9,"SBRE":9,"SEPTBRE":9,
                   "AGTO":8,"AGST":8,"OCTB":10,"OCTBRE":10,"NVBRE":11,
                   "NOVBRE":11,"DICBRE":12,"DBRE":12,
                   "JANUARY":1,"FEBRUARY":2,"MARCH":3,"APRIL":4,"JUNE":6,
                   "JULY":7,"AUGUST":8,"SEPTEMBER":9,"OCTOBER":10,
                   "NOVEMBER":11,"DECEMBER":12})

def normalizar_mes(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return (None, "")
    if isinstance(v, (datetime, pd.Timestamp)):
        return (v.month, MESES_NOMBRE[v.month])
    if isinstance(v, (int, float, np.integer, np.floating)):
        if pd.isna(v):
            return (None, "")
        n = int(v)
        return (n, MESES_NOMBRE[n]) if 1 <= n <= 12 else (None, "")
    s = sin_tildes(v).upper().strip().replace(".", "").replace("_", " ")
    s = re.sub(r"\s+", " ", s).strip()
    if not s:
        return (None, "")
    if s in _ALIAS_MES:
        n = _ALIAS_MES[s]; return (n, MESES_NOMBRE[n])
    for tok in re.split(r"[\s/\-]+", s):
        if tok in _ALIAS_MES:
            n = _ALIAS_MES[tok]; return (n, MESES_NOMBRE[n])
    try:
        f = pd.to_datetime(s, dayfirst=True)
        return (f.month, MESES_NOMBRE[f.month])
    except Exception:
        return (None, "")


def buscar_columna(cols, exactos=(), contiene=(), excluir=()):
    lim = [limpiar_encabezado(c) for c in cols]
    for i, h in enumerate(lim):
        if h in exactos and not any(x in h for x in excluir):
            return i
    for i, h in enumerate(lim):
        if any(t in h for t in contiene) and not any(x in h for x in excluir):
            return i
    return None


# ============================================================ MEMORIA
def optimizar_memoria(df, verbose=False):
    """
    Reduce el consumo del DataFrame:
      - texto repetitivo -> category  (Sucursal, Oficina, Descripción...)
      - enteros -> el int más pequeño que quepa
      - float64 -> float32 donde no se pierde precisión contable
    En un mayor contable típico baja el consumo 60-80 %.
    """
    antes = df.memory_usage(deep=True).sum() / 1024**2

    for col in df.columns:
        s = df[col]
        if s.dtype == "object":
            n = len(s)
            if n and s.nunique(dropna=False) / n < 0.5:   # muy repetitivo
                df[col] = s.astype("category")
        elif pd.api.types.is_integer_dtype(s):
            df[col] = pd.to_numeric(s, downcast="integer")
        elif pd.api.types.is_float_dtype(s):
            df[col] = pd.to_numeric(s, downcast="float")

    despues = df.memory_usage(deep=True).sum() / 1024**2
    if verbose:
        ahorro = 100 * (1 - despues / antes) if antes else 0
        print(f"     memoria: {antes:.1f} -> {despues:.1f} MB  (-{ahorro:.0f} %)")
    return df


def leer_encabezado(archivo, max_filas=25):
    """
    Detecta la fila de encabezado leyendo SOLO las primeras filas.
    Antes se leía el archivo completo dos veces.
    """
    muestra = pd.read_excel(archivo, sheet_name=0, header=None,
                            nrows=max_filas, dtype=object)
    mejor, score = None, 0
    for i in range(len(muestra)):
        fila = muestra.iloc[i]
        textos = sum(1 for v in fila if isinstance(v, str) and v.strip())
        llenas = fila.notna().sum()
        if llenas and textos / max(llenas, 1) > 0.7 and textos > score:
            mejor, score = i, textos
    del muestra; gc.collect()
    return mejor

print("Funciones listas.")
ram("post-funciones")

#5
EXTENSIONES = (".xlsx", ".xlsm", ".xls")

archivos, ignorados = [], []
for f in sorted(os.listdir(CARPETA_ENTRADA)):
    r = os.path.join(CARPETA_ENTRADA, f)
    if not os.path.isfile(r) or f.startswith(("~$", ".")):
        continue
    (archivos if os.path.splitext(f)[1].lower() in EXTENSIONES
     else ignorados).append(r if os.path.splitext(f)[1].lower() in EXTENSIONES else f)

if ignorados:
    print("[AVISO] Ignorados por extensión:")
    for f in ignorados:
        print(f"    {f!r}" + ("  <-- posible Google Sheet nativo"
                              if not os.path.splitext(f)[1] else ""))
    print()

if not archivos:
    raise FileNotFoundError(f"No hay archivos de Excel en: {CARPETA_ENTRADA}")

tam_total = sum(os.path.getsize(a) for a in archivos) / 1024**2
print(f"Archivos: {len(archivos)}  |  {tam_total:.1f} MB en disco")
print(f"(en RAM ocuparán aprox. {tam_total * 6:.0f} MB ya optimizados)\n")

frames, ARCHIVOS_OK, ARCHIVOS_FALLA, checkpoints = [], [], [], []
columnas_referencia = None

for k, archivo in enumerate(archivos, 1):
    nombre = os.path.basename(archivo)
    try:
        fila_enc = ((FILA_ENCABEZADO - 1) if FILA_ENCABEZADO
                    else leer_encabezado(archivo))
        if fila_enc is None:
            raise ValueError("no se detectó fila de encabezado")

        # SIN dtype=object: dejamos que pandas infiera tipos nativos.
        # normalizar_numero() maneja igual texto o número, así que no
        # perdemos nada y ahorramos 10-15x de RAM.
        df = pd.read_excel(archivo, sheet_name=0, header=fila_enc)

        df = df.dropna(how="all").dropna(axis=1, how="all")
        df = df.loc[:, [c for c in df.columns if not str(c).startswith("Unnamed:")]]

        primera = df.columns[0]
        df = df[~df[primera].astype(str).str.upper().str.strip()
                .isin([limpiar_encabezado(primera), "TOTAL", "TOTALES", "NAN", ""])]

        if df.empty:
            raise ValueError("vacío tras la limpieza")

        if columnas_referencia is None:
            columnas_referencia = list(df.columns)
        elif list(df.columns) != columnas_referencia:
            if len(df.columns) == len(columnas_referencia):
                print(f"  [AVISO] {nombre}: encabezados distintos -> alineado por posición.")
                df.columns = columnas_referencia
            else:
                raise ValueError(f"{len(df.columns)} columnas vs "
                                 f"{len(columnas_referencia)} del primero")

        df["Archivo Origen"] = nombre
        df = optimizar_memoria(df)

        print(f"  [{k}/{len(archivos)}] OK  {nombre}  |  fila enc. {fila_enc+1}  |  "
              f"{len(df):,} filas")

        # --- Si la RAM se está agotando, vuelca a disco local y libera ---
        libre = ram()
        if MODO_BAJO_CONSUMO and libre < UMBRAL_RAM_GB and frames:
            ruta_ck = os.path.join(CARPETA_TEMP, f"ck_{len(checkpoints):03d}.parquet")
            pd.concat(frames, ignore_index=True).to_parquet(ruta_ck, index=False)
            checkpoints.append(ruta_ck)
            print(f"     [CHECKPOINT] volcado a disco -> {os.path.basename(ruta_ck)}")
            frames.clear(); gc.collect()

        frames.append(df)
        ARCHIVOS_OK.append(archivo)

    except Exception as e:
        print(f"  [{k}/{len(archivos)}] [ERROR] {nombre}: {e}")
        ARCHIVOS_FALLA.append((archivo, str(e)))
    finally:
        for v in ("df",):
            if v in dir():
                pass
        gc.collect()

if not frames and not checkpoints:
    raise ValueError("Ningún archivo pudo procesarse.")

# --- Unión final: primero disco, luego memoria ---
partes = [pd.read_parquet(c) for c in checkpoints] + frames
df_consolidado = pd.concat(partes, ignore_index=True, copy=False)

liberar(partes, frames)
del partes
frames = []
gc.collect()

for c in checkpoints:
    try: os.remove(c)
    except Exception: pass

df_consolidado = optimizar_memoria(df_consolidado, verbose=True)

print(f"\nConsolidado: {len(df_consolidado):,} filas x {len(df_consolidado.columns)} cols")
print(f"OK: {len(ARCHIVOS_OK)}   |   Falla: {len(ARCHIVOS_FALLA)}")
for r, m in ARCHIVOS_FALLA:
    print(f"    (se queda en Entrada) {os.path.basename(r)}  ->  {m}")
ram("post-consolidación")

#6
cols = [c for c in df_consolidado.columns if c != "Archivo Origen"]

print("=" * 100)
print(f"{'LETRA':<7}{'IDX':<5}{'ENCABEZADO':<40}{'MUESTRA'}")
print("=" * 100)
for i, c in enumerate(cols):
    m = " | ".join(str(v)[:16] for v in df_consolidado[c].dropna().head(3))
    print(f"{indice_a_letra(i):<7}{i:<5}{str(c)[:38]:<40}{m[:44]}")
print("=" * 100 + "\n")

idx_valor = (letra_a_indice(COL_VALOR) if COL_VALOR else buscar_columna(
    cols, exactos=("VALOR","BASE","IMPORTE","MONTO","VALOR BASE","BASE GRAVABLE"),
    contiene=("VALOR","BASE","IMPORTE","MONTO","DEBITO","CREDITO","SALDO"),
    excluir=("RETENCION","TARIFA","IVA","FECHA","CODIGO")))
idx_premio = (letra_a_indice(COL_PREMIO) if COL_PREMIO else buscar_columna(
    cols, exactos=("PREMIO","ES PREMIO","PREMIOS","APLICA PREMIO","SI/NO"),
    contiene=("PREMIO",)))
idx_mes = (letra_a_indice(COL_MES) if COL_MES else buscar_columna(
    cols, exactos=("MES","PERIODO","PERIODO GRAVABLE","MES DECLARADO"),
    contiene=("MES","PERIODO","FECHA")))

print("COLUMNAS RESUELTAS:")
faltan = []
for et, idx, par in [("VALOR / BASE", idx_valor, "COL_VALOR"),
                     ("SI/NO PREMIO", idx_premio, "COL_PREMIO"),
                     ("MES", idx_mes, "COL_MES")]:
    if idx is None:
        print(f"   [X] {et:<14} NO ENCONTRADA -> define {par} en el BLOQUE 2")
        faltan.append(par)
    else:
        print(f"   [OK] {et:<14} col {indice_a_letra(idx)} (idx {idx}) -> '{cols[idx]}'")
        print(f"        muestra: {df_consolidado[cols[idx]].dropna().head(4).tolist()}")
if faltan:
    raise ValueError(f"Define en el BLOQUE 2 y vuelve a correr: {faltan}")

#7
col_valor, col_premio, col_mes = cols[idx_valor], cols[idx_premio], cols[idx_mes]

# .map() sobre una categoría solo evalúa los valores ÚNICOS, no fila por fila.
# En 300.000 filas con 40 valores distintos de mes, son 40 llamadas en vez
# de 300.000. Por eso convertimos primero.
def mapear_eficiente(serie, funcion):
    unicos = pd.Series(serie.dropna().unique())
    tabla = dict(zip(unicos, unicos.map(funcion)))
    return serie.map(tabla)

df_consolidado["Base Normalizada"] = pd.to_numeric(
    mapear_eficiente(df_consolidado[col_valor].astype(object), normalizar_numero),
    errors="coerce", downcast="float")

df_consolidado["Es Premio"] = mapear_eficiente(
    df_consolidado[col_premio].astype(object), normalizar_si_no
).fillna("").astype("category")

_m = mapear_eficiente(df_consolidado[col_mes].astype(object), normalizar_mes)
df_consolidado["Mes Num"] = pd.to_numeric([x[0] if isinstance(x, tuple) else None
                                           for x in _m], errors="coerce")
df_consolidado["Mes"] = pd.Series([x[1] if isinstance(x, tuple) else "" for x in _m],
                                  index=df_consolidado.index).astype("category")
liberar(_m); del _m

df_consolidado["Periodo"] = pd.Series(
    [f"{ANIO_GRAVABLE}-{int(n):02d}" if pd.notna(n) else ""
     for n in df_consolidado["Mes Num"]],
    index=df_consolidado.index).astype("category")

print("--- CONTROL DE CALIDAD ---\n")
n = int(df_consolidado["Base Normalizada"].isna().sum())
if n:
    print(f"[!] {n:,} filas con base no numérica. Crudos:")
    print("   ", df_consolidado.loc[df_consolidado['Base Normalizada'].isna(), col_valor]
          .dropna().astype(str).unique()[:10], "\n")
else:
    print("[OK] Todas las bases se convirtieron a número.\n")

raros = sorted(set(df_consolidado["Es Premio"].cat.categories) - {"SI","NO",""})
print(f"[!] No reconocidos en '{col_premio}': {raros}\n" if raros
      else "[OK] Columna SI/NO sin residuos.\n")

n = int(df_consolidado["Mes Num"].isna().sum())
if n:
    print(f"[!] {n:,} filas sin mes reconocible. Crudos:")
    print("   ", df_consolidado.loc[df_consolidado['Mes Num'].isna(), col_mes]
          .dropna().astype(str).unique()[:10], "\n")
else:
    print("[OK] Todos los meses reconocidos.\n")

print(df_consolidado["Mes"].value_counts().to_string())
print(f"\nSuma de bases: {df_consolidado['Base Normalizada'].sum():,.0f}")
ram("post-normalización")

#8
df_consolidado["Tarifa Retención (%)"] = np.float32(0.0)

es_premio = df_consolidado["Es Premio"].eq("SI")
no_premio = df_consolidado["Es Premio"].eq("NO")
base      = df_consolidado["Base Normalizada"]

df_consolidado.loc[es_premio & base.gt(TOPE_PREMIOS), "Tarifa Retención (%)"] = TARIFA_PREMIOS
df_consolidado.loc[no_premio & base.gt(TOPE_OTROS),   "Tarifa Retención (%)"] = TARIFA_OTROS

df_consolidado["Valor Retención"] = (base * df_consolidado["Tarifa Retención (%)"] / 100).round(0)
df_consolidado.loc[base.isna(), "Valor Retención"] = np.nan

print("--- RESUMEN ---")
print(df_consolidado.groupby("Tarifa Retención (%)", observed=True)
      .agg(Filas=("Base Normalizada","size"), Base=("Base Normalizada","sum"),
           Retencion=("Valor Retención","sum"))
      .to_string(float_format=lambda x: f"{x:,.0f}"))

# Seguro contra caídas: si Colab se reinicia, recuperas con
#   df_consolidado = pd.read_parquet(RUTA_CHECKPOINT)
if CHECKPOINT:
    RUTA_CHECKPOINT = os.path.join(CARPETA_TEMP, "consolidado_calculado.parquet")
    df_consolidado.to_parquet(RUTA_CHECKPOINT, index=False, compression="snappy")
    print(f"\n[CHECKPOINT] {RUTA_CHECKPOINT} "
          f"({os.path.getsize(RUTA_CHECKPOINT)/1024**2:.1f} MB)")
ram("post-cálculo")

#9

def escribir_excel(ruta, hojas, constant_memory=True):
    """
    Escribe y formatea en una sola pasada con xlsxwriter.

    constant_memory=True hace que xlsxwriter descargue cada fila a disco
    apenas la escribe, en vez de mantener todo el libro en RAM. El costo:
    hay que escribir las filas en orden estricto (por eso el encabezado
    va primero, a mano, y los datos después con startrow=1).

    Reemplaza el openpyxl.load_workbook() anterior, que volvía a cargar
    el Excel completo en memoria con un objeto Python por celda.
    """
    with pd.ExcelWriter(ruta, engine="xlsxwriter",
                        engine_kwargs={"options": {
                            "constant_memory": constant_memory,
                            "strings_to_numbers": False,
                            "default_date_format": "yyyy-mm-dd"}}) as writer:

        wb = writer.book
        f_enc = wb.add_format({"bold": True, "font_color": "FFFFFF",
                               "bg_color": "FF0000", "font_name": "Calibri",
                               "font_size": 11, "align": "center",
                               "valign": "vcenter", "text_wrap": True,
                               "border": 1})
        f_num  = wb.add_format({"num_format": "#,##0", "font_name": "Calibri",
                                "font_size": 11, "align": "center", "valign": "vcenter"})
        f_tar  = wb.add_format({"num_format": "0.0", "font_name": "Calibri",
                                "font_size": 11, "align": "center", "valign": "vcenter"})
        f_ent  = wb.add_format({"num_format": "0", "font_name": "Calibri",
                                "font_size": 11, "align": "center", "valign": "vcenter"})
        f_txt  = wb.add_format({"font_name": "Calibri", "font_size": 11,
                                "align": "center", "valign": "vcenter"})

        for nombre, df in hojas.items():
            hoja = nombre[:31]
            ws = wb.add_worksheet(hoja)
            writer.sheets[hoja] = ws

            if df is None or df.empty:
                ws.write(0, 0, "Sin registros", f_enc)
                print(f"  Hoja '{nombre}': 0 filas")
                continue

            # 1) Encabezado PRIMERO (obligatorio con constant_memory)
            for j, col in enumerate(df.columns):
                ws.write(0, j, str(col), f_enc)

            # 2) Ancho y formato por COLUMNA, no celda por celda.
            #    Se muestrean 200 filas en vez de recorrer todo.
            muestra = df.head(200)
            for j, col in enumerate(df.columns):
                titulo = str(col)
                largo = max([len(titulo)] +
                            [len(str(v)) for v in muestra[col].head(200) if v is not None])
                ancho = min(max(largo + 2, 10), 40)
                if "Tarifa" in titulo:
                    fmt = f_tar
                elif "Mes Num" in titulo:
                    fmt = f_ent
                elif pd.api.types.is_numeric_dtype(df[col]):
                    fmt = f_num
                else:
                    fmt = f_txt
                ws.set_column(j, j, ancho, fmt)

            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, len(df), len(df.columns) - 1)

            # 3) Datos, sin encabezado, desde la fila 2
            df.to_excel(writer, sheet_name=hoja, index=False,
                        header=False, startrow=1)
            print(f"  Hoja '{nombre}': {len(df):,} filas")
            gc.collect()

    return ruta

print("Función escribir_excel() lista.")

#10

import os, shutil, time, gc
import pandas as pd, numpy as np
from datetime import datetime, timedelta

req = ["df_consolidado","CARPETA_SALIDA","ANIO_GRAVABLE","MESES_NOMBRE",
       "TARIFA_PREMIOS","TARIFA_OTROS","escribir_excel"]
falt = [v for v in req if v not in globals()]
if falt:
    raise NameError(f"Faltan: {falt}\nSe reinició el entorno. Corre los BLOQUES 0 a 8.\n"
                    f"Si existe el checkpoint, puedes recuperar con:\n"
                    f"  df_consolidado = pd.read_parquet(RUTA_CHECKPOINT)")

os.makedirs(CARPETA_SALIDA, exist_ok=True)

base      = df_consolidado["Base Normalizada"]
es_premio = df_consolidado["Es Premio"].eq("SI")
no_premio = df_consolidado["Es Premio"].eq("NO")
tarifa    = df_consolidado["Tarifa Retención (%)"]

validos = df_consolidado["Mes Num"].dropna()
if len(validos):
    mes_num = int(validos.mode().iloc[0])
    if validos.nunique() > 1:
        print(f"[AVISO] {validos.nunique()} meses distintos. "
              f"Se nombra con el predominante: {MESES_NOMBRE[mes_num]}")
else:
    ant = datetime.now().replace(day=1) - timedelta(days=1)
    mes_num = ant.month
    print("[AVISO] Ningún mes reconocido. Se usa el mes anterior calendario.")

nombre_mes   = MESES_NOMBRE[mes_num]
ETIQUETA_MES = f"{mes_num:02d}_{nombre_mes.lower()}"
ruta_salida  = os.path.join(CARPETA_SALIDA,
    f"Consolidado_con_Retencion-{ETIQUETA_MES}_{ANIO_GRAVABLE}.xlsx")

sin_clasificar = (base.isna() | ~df_consolidado["Es Premio"].isin(["SI","NO"])
                  | df_consolidado["Mes Num"].isna())

# .loc devuelve vistas mientras no se modifiquen: no se duplica el DataFrame
hojas = {
    "Consolidado General":            df_consolidado,
    "Premios20%":                     df_consolidado.loc[tarifa == TARIFA_PREMIOS],
    "Otros ingresos tributarios3.5%": df_consolidado.loc[tarifa == TARIFA_OTROS],
    "Premios Sin Retencion":          df_consolidado.loc[es_premio & tarifa.eq(0) & base.notna()],
    "Otros Ingresos Sin retencion":   df_consolidado.loc[no_premio & tarifa.eq(0) & base.notna()],
    "REVISAR":                        df_consolidado.loc[sin_clasificar],
}

if os.path.exists(ruta_salida):
    hist = os.path.join(CARPETA_SALIDA, "Historico"); os.makedirs(hist, exist_ok=True)
    b, e = os.path.splitext(os.path.basename(ruta_salida))
    shutil.copy2(ruta_salida, os.path.join(hist, f"{b}__{datetime.now():%Y%m%d_%H%M%S}{e}"))
    print("[BACKUP] Versión anterior -> Historico/\n")

# Se escribe primero en disco local (rápido y sin latencia de Drive)
# y solo al final se copia a Drive.
ruta_local = os.path.join(CARPETA_TEMP, os.path.basename(ruta_salida))
escribir_excel(ruta_local, hojas)
liberar(hojas); del hojas; gc.collect()

shutil.copy2(ruta_local, ruta_salida)
os.remove(ruta_local)

EXPORTACION_OK = False
for _ in range(15):
    if os.path.exists(ruta_salida) and os.path.getsize(ruta_salida) > 0:
        EXPORTACION_OK = True
        print(f"\n[OK] {ruta_salida}")
        print(f"     {os.path.getsize(ruta_salida)/1024**2:.2f} MB")
        break
    time.sleep(1)
if not EXPORTACION_OK:
    print("\n[X] NO se confirmó la escritura. NO muevas los archivos todavía.")
ram("post-exportación")

#11
if not globals().get("EXPORTACION_OK", False):
    raise RuntimeError("La exportación no se confirmó. Revisa el BLOQUE 9.")

destino = os.path.join(CARPETA_PROCESADOS, str(ANIO_GRAVABLE), ETIQUETA_MES)

if not MOVER_PROCESADOS:
    print("MOVER_PROCESADOS = False  ->  no se movió nada.\n")
    print(f"Se moverían ({len(ARCHIVOS_OK)}):")
    for r in ARCHIVOS_OK:
        print(f"    {os.path.basename(r)}")
    print(f"\nDestino: {destino}")
    print("\nCambia MOVER_PROCESADOS = True en el BLOQUE 2 y repite este bloque.")
else:
    os.makedirs(destino, exist_ok=True)
    movidos, errores = [], []
    for origen in ARCHIVOS_OK:
        nombre = os.path.basename(origen)
        if not os.path.exists(origen):
            print(f"  [SKIP] {nombre}: ya no está en Entrada."); continue
        rd = os.path.join(destino, nombre)
        if os.path.exists(rd):
            b, e = os.path.splitext(nombre)
            rd = os.path.join(destino, f"{b}__{datetime.now():%Y%m%d_%H%M%S}{e}")
            print(f"  [DUP] {nombre} -> {os.path.basename(rd)}")
        try:
            # copy2 + remove en vez de move: en Drive montado, move puede
            # fallar a mitad y dejar el archivo perdido.
            shutil.copy2(origen, rd)
            if os.path.exists(rd) and os.path.getsize(rd) > 0:
                os.remove(origen); movidos.append(nombre); print(f"  OK  {nombre}")
            else:
                errores.append((nombre, "copia vacía"))
                print(f"  [X] {nombre}: copia vacía, NO se borró el original.")
        except Exception as ex:
            errores.append((nombre, str(ex))); print(f"  [X] {nombre}: {ex}")

    log = os.path.join(CARPETA_PROCESADOS, "bitacora_procesamiento.csv")
    if movidos:
        reg = pd.DataFrame({"Fecha_Proceso": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "Anio": ANIO_GRAVABLE, "Mes": nombre_mes,
                            "Archivo": movidos, "Carpeta_Destino": destino,
                            "Consolidado": os.path.basename(ruta_salida)})
        reg.to_csv(log, mode="a" if os.path.exists(log) else "w",
                   header=not os.path.exists(log), index=False, encoding="utf-8-sig")

    print(f"\n{'='*60}\nMovidos: {len(movidos)}  |  Con error: {len(errores)}")
    print(f"Destino: {destino}\nBitácora: {log}")
    for r, m in ARCHIVOS_FALLA:
        print(f"  (sigue en Entrada) {os.path.basename(r)}  ->  {m}")
    print("=" * 60)

#12
# Libera el consolidado mensual antes de empezar: ya está en disco.
for v in ["df_consolidado", "base", "es_premio", "no_premio", "tarifa"]:
    if v in globals():
        del globals()[v]
gc.collect()
ram("antes del anual")

patron = os.path.join(CARPETA_SALIDA, "Consolidado_con_Retencion-*.xlsx")
archivos_mes = [a for a in sorted(glob.glob(patron))
                if "ANUAL" not in os.path.basename(a).upper()
                and not os.path.basename(a).startswith("~$")]
if not archivos_mes:
    raise FileNotFoundError(f"No hay consolidados mensuales en {CARPETA_SALIDA}")

print(f"Consolidados: {len(archivos_mes)}\n")

# --- PASADA 1: resúmenes, un mes a la vez ---
# Solo se leen las columnas necesarias y se descarta cada mes al terminar.
COLS_RESUMEN = ["Mes Num", "Mes", "Base Normalizada",
                "Valor Retención", "Tarifa Retención (%)"]
resumenes, rutas_parquet = [], []

for a in archivos_mes:
    nom = os.path.basename(a)
    try:
        d = pd.read_excel(a, sheet_name="Consolidado General", usecols=COLS_RESUMEN)
        r = (d.groupby(["Mes Num","Mes","Tarifa Retención (%)"], dropna=False, observed=True)
             .agg(Registros=("Base Normalizada","size"),
                  Base_Total=("Base Normalizada","sum"),
                  Retencion_Total=("Valor Retención","sum")).reset_index())
        resumenes.append(r)
        print(f"  OK  {nom}  ->  {len(d):,} filas")
        del d, r; gc.collect()
    except Exception as e:
        print(f"  [ERROR] {nom}: {e}")

agg = pd.concat(resumenes, ignore_index=True)
del resumenes; gc.collect()

resumen_mes = (agg.groupby(["Mes Num","Mes"], dropna=False, observed=True)
               .agg(Registros=("Registros","sum"), Base_Total=("Base_Total","sum"),
                    Retencion_Total=("Retencion_Total","sum"))
               .reset_index().sort_values("Mes Num"))

resumen_tarifa = (agg.groupby("Tarifa Retención (%)", observed=True)
                  .agg(Registros=("Registros","sum"), Base_Total=("Base_Total","sum"),
                       Retencion_Total=("Retencion_Total","sum")).reset_index())

resumen_mes_tarifa = (agg.pivot_table(index=["Mes Num","Mes"],
                                      columns="Tarifa Retención (%)",
                                      values=["Base_Total","Retencion_Total"],
                                      aggfunc="sum", fill_value=0, observed=True)
                      .reset_index())
resumen_mes_tarifa.columns = [" ".join(str(x) for x in c if str(x) != "").strip()
                              for c in resumen_mes_tarifa.columns.to_flat_index()]

print("\n--- ANUAL POR MES ---")
print(resumen_mes.to_string(index=False, float_format=lambda x: f"{x:,.0f}"))
print("\n--- ANUAL POR TARIFA ---")
print(resumen_tarifa.to_string(index=False, float_format=lambda x: f"{x:,.0f}"))

presentes = set(resumen_mes["Mes Num"].dropna().astype(int))
falt = [MESES_NOMBRE[m] for m in range(1,13) if m not in presentes]
if falt:
    print(f"\n[AVISO] Meses sin datos: {falt}")

# --- PASADA 2: detalle a parquet, mes por mes ---
for a in archivos_mes:
    try:
        d = pd.read_excel(a, sheet_name="Consolidado General")
        d = optimizar_memoria(d)
        p = os.path.join(CARPETA_TEMP, os.path.basename(a).replace(".xlsx",".parquet"))
        d.to_parquet(p, index=False); rutas_parquet.append(p)
        del d; gc.collect()
    except Exception as e:
        print(f"  [ERROR detalle] {os.path.basename(a)}: {e}")

ram("post-parquet")

df_anual = pd.concat([pd.read_parquet(p) for p in rutas_parquet], ignore_index=True)
antes = len(df_anual)
df_anual = df_anual.drop_duplicates()
if antes != len(df_anual):
    print(f"[AVISO] {antes - len(df_anual):,} filas duplicadas eliminadas.")
df_anual = optimizar_memoria(df_anual.sort_values("Mes Num", na_position="last")
                             .reset_index(drop=True), verbose=True)
print(f"Total anual: {len(df_anual):,} filas")
ram("df_anual listo")

t = df_anual["Tarifa Retención (%)"]
ruta_anual = os.path.join(CARPETA_SALIDA, f"CONSOLIDADO_ANUAL_{ANIO_GRAVABLE}.xlsx")
ruta_local = os.path.join(CARPETA_TEMP, os.path.basename(ruta_anual))

escribir_excel(ruta_local, {
    "Resumen por Mes":                resumen_mes,
    "Resumen Mes x Tarifa":           resumen_mes_tarifa,
    "Resumen por Tarifa":             resumen_tarifa,
    "Consolidado Anual":              df_anual,
    "Premios20%":                     df_anual.loc[t == TARIFA_PREMIOS],
    "Otros ingresos tributarios3.5%": df_anual.loc[t == TARIFA_OTROS],
    "REVISAR":                        df_anual.loc[df_anual["Base Normalizada"].isna()
                                                   | ~df_anual["Es Premio"].isin(["SI","NO"])
                                                   | df_anual["Mes Num"].isna()],
})

shutil.copy2(ruta_local, ruta_anual); os.remove(ruta_local)
for p in rutas_parquet:
    try: os.remove(p)
    except Exception: pass

print(f"\n[OK] {ruta_anual}  ({os.path.getsize(ruta_anual)/1024**2:.2f} MB)")
ram("fin")

#13
def inventario(carpeta, nivel=0):
    if not os.path.isdir(carpeta):
        print(f"  (no existe: {carpeta})"); return
    for f in sorted(os.listdir(carpeta)):
        r = os.path.join(carpeta, f); s = "  " * (nivel + 1)
        if os.path.isdir(r):
            print(f"{s}[{f}]"); inventario(r, nivel + 1)
        else:
            mod = datetime.fromtimestamp(os.path.getmtime(r))
            print(f"{s}{f}  ({os.path.getsize(r)/1024**2:.2f} MB, {mod:%Y-%m-%d %H:%M})")

for c, t in [(CARPETA_ENTRADA,"ENTRADA"), (CARPETA_SALIDA,"SALIDA"),
             (CARPETA_PROCESADOS,"PROCESADOS")]:
    print("=" * 70); print(t); print("=" * 70)
    inventario(c); print()

# Limpiar temporales locales (NO toca Drive)
n = 0
for f in os.listdir(CARPETA_TEMP):
    try: os.remove(os.path.join(CARPETA_TEMP, f)); n += 1
    except Exception: pass
print(f"Temporales locales eliminados: {n}")
ram("final")