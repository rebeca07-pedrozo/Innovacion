# BLOQUE 1
!pip install xlsxwriter -q
import pandas as pd, numpy as np, os, re, unicodedata, glob, shutil
from google.colab import drive
drive.mount('/content/drive')

RAIZ      = "/content/drive/My Drive/Optimizacion-Premios/2026/"
ENTRADA   = RAIZ + "Entrada/"
SALIDA    = RAIZ + "Salida/"
PROCESADOS= RAIZ + "Procesados/"
MAESTRO   = RAIZ + "MaestroTerceros.xlsx"   # opcional

ANIO    = 2026
I_MES   = 0
I_VALOR = 17
I_SINO  = 9

TOPE_PREMIOS = 2_514_000
TOPE_OTROS   = 524_000

# --- Formato 1001 ---
CONCEPTO_SI = "5044"   # registrado en Coljuegos
CONCEPTO_NO = "5016"   # cashback y premios no registrados
ET_PAGO     = "pago"
ET_RETE     = "reta"
FORMATO     = "1001"
ORIGEN      = "Premios"
RETE_EN_CERO = False   # True si hay que reportar reta aunque sea 0

# Nombre REAL de la columna en tus archivos. "" = queda en blanco.
COL_DOC = ""
MAPA = {
    "TIPO DE DOCUMENTO": "",
    "DIGITO CHEQUEO":    "",
    "APELLIDO1":         "",
    "APELLIDO2":         "",
    "NOMBRE1":           "",
    "NOMBRE2":           "",
    "RAZON SOCIAL":      "",
    "DIRECCION":         "",
    "DEPTO Y MUNICIPIO": "",
    "PAIS":              "",
}

def num(v):
    if isinstance(v, (int, float, np.integer, np.floating)):
        return np.nan if pd.isna(v) else float(v)
    if v is None:
        return np.nan
    s = re.sub(r"[^\d,\.\-()]", "", str(v)).strip()
    if s in ("", "-", ".", ","):
        return np.nan
    neg = s.startswith("(") or s.startswith("-")
    s = s.strip("()").lstrip("-")
    if "," in s and "." in s:
        s = (s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".")
             else s.replace(",", ""))
    elif "," in s:
        p = s.split(",")
        s = s.replace(",", ".") if (len(p) == 2 and len(p[1]) != 3) else s.replace(",", "")
    elif "." in s:
        p = s.split(".")
        if not (len(p) == 2 and len(p[1]) != 3):
            s = s.replace(".", "")
    try:
        x = float(s)
    except ValueError:
        return np.nan
    return -x if neg else x

def sino(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = "".join(c for c in unicodedata.normalize("NFD", str(v))
                if unicodedata.category(c) != "Mn")
    s = s.upper().replace(".", "").replace(" ", "")
    if s in {"SI", "S", "1", "1.0", "X"}: return "SI"
    if s in {"NO", "N", "0", "0.0"}:      return "NO"
    return s

MESES = {1:"enero",2:"febrero",3:"marzo",4:"abril",5:"mayo",6:"junio",
         7:"julio",8:"agosto",9:"septiembre",10:"octubre",11:"noviembre",12:"diciembre"}
ALIAS = {v.upper(): k for k, v in MESES.items()}
ALIAS.update({"SETIEMBRE":9,"ENE":1,"FEB":2,"MAR":3,"ABR":4,"MAY":5,
              "JUN":6,"JUL":7,"AGO":8,"SEP":9,"OCT":10,"NOV":11,"DIC":12})

def mesnum(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = "".join(c for c in unicodedata.normalize("NFD", str(v))
                if unicodedata.category(c) != "Mn").upper().strip().replace(".", "")
    if s in ALIAS: return ALIAS[s]
    if s.isdigit() and 1 <= int(s) <= 12: return int(s)
    return None

def texto(v):
    if v is None or (isinstance(v, float) and pd.isna(v)): return ""
    return str(v).strip()

for c in (ENTRADA, SALIDA, PROCESADOS):
    os.makedirs(c, exist_ok=True)

archivos = sorted(p for p in glob.glob(os.path.join(ENTRADA, "**", "*.xls*"), recursive=True)
                  if not os.path.basename(p).startswith("~$"))
if not archivos:
    raise FileNotFoundError(f"No hay archivos en {ENTRADA}")

print(f"Archivos encontrados: {len(archivos)}\n")
frames = []
for a in archivos:
    d = pd.read_excel(a, header=0)
    d["Archivo Origen"] = os.path.basename(a)
    frames.append(d)
    print(f"  {os.path.basename(a)} -> {len(d):,} filas")

df = pd.concat(frames, ignore_index=True)
print(f"\nTotal: {len(df):,} filas x {len(df.columns)} columnas\n")

print("="*70)
print("COLUMNAS CLAVE")
for et, i in [("MES", I_MES), ("VALOR", I_VALOR), ("SI/NO", I_SINO)]:
    print(f"  {et:<6} idx {i:>2} -> '{df.columns[i]}'")
    print(f"         {df.iloc[:4, i].tolist()}")
print("\nTODAS LAS COLUMNAS (para llenar COL_DOC y MAPA):")
for i, c in enumerate(df.columns):
    print(f"  [{i:>2}] {c}")
print("="*70)

#BLOQUE 2 
c_mes, c_val, c_sino = df.columns[I_MES], df.columns[I_VALOR], df.columns[I_SINO]

df["Base"]      = df[c_val].map(num)
df["Coljuegos"] = df[c_sino].map(sino)
df["Mes Num"]   = df[c_mes].map(mesnum)

si = df["Coljuegos"].eq("SI")
no = df["Coljuegos"].eq("NO")

df["Tarifa Retención (%)"] = 0.0
df.loc[si & df["Base"].gt(TOPE_PREMIOS), "Tarifa Retención (%)"] = 20.0
df.loc[no & df["Base"].gt(TOPE_OTROS),   "Tarifa Retención (%)"] = 3.5

df["Valor Retención"] = (df["Base"] * df["Tarifa Retención (%)"] / 100).round(0)
df.loc[df["Base"].isna(), "Valor Retención"] = np.nan

df["Concepto"] = ""
df.loc[si, "Concepto"] = CONCEPTO_SI
df.loc[no, "Concepto"] = CONCEPTO_NO

print("--- CONTROL ---")
print(f"Bases no numéricas : {df['Base'].isna().sum():,}")
raros = sorted(set(df["Coljuegos"]) - {"SI", "NO", ""})
print(f"SI/NO no reconocido: {raros}")
print(f"Sin concepto       : {(df['Concepto'] == '').sum():,} filas (van a REVISAR)\n")

print("--- POR CONCEPTO ---")
print(df.groupby("Concepto").agg(Filas=("Base","size"),
                                 Base=("Base","sum"),
                                 Retencion=("Valor Retención","sum"))
        .to_string(float_format=lambda x: f"{x:,.2f}"), "\n")

ctrl = (df.groupby(["Archivo Origen","Mes Num"], dropna=False).size()
          .reset_index(name="Filas"))
ctrl["Mes"] = ctrl["Mes Num"].map(MESES).fillna("SIN MES")
print("--- MES DETECTADO POR ARCHIVO ---")
print(ctrl[["Archivo Origen","Mes","Filas"]].to_string(index=False), "\n")

mezcla = ctrl.groupby("Archivo Origen")["Mes Num"].nunique()
if (mezcla > 1).any():
    print("[AVISO] Archivos con más de un mes:", list(mezcla[mezcla > 1].index))

falta = df[df["Mes Num"].isna()]
if len(falta):
    print(f"[AVISO] {len(falta):,} filas sin mes en:", sorted(falta["Archivo Origen"].unique()))

#BLOQUE 3
def escribir(ruta, hojas):
    with pd.ExcelWriter(ruta, engine="xlsxwriter") as w:
        wb = w.book
        f_enc = wb.add_format({"bold":True,"font_color":"FFFFFF","bg_color":"FF0000",
                               "font_name":"Calibri","font_size":11,"align":"center",
                               "valign":"vcenter","text_wrap":True})
        f_num = wb.add_format({"num_format":"#,##0.00","font_name":"Calibri",
                               "font_size":11,"align":"center"})
        f_txt = wb.add_format({"font_name":"Calibri","font_size":11,"align":"center"})
        for h, d in hojas.items():
            d.to_excel(w, sheet_name=h[:31], index=False)
            ws = w.sheets[h[:31]]
            for j, c in enumerate(d.columns):
                ws.write(0, j, str(c), f_enc)
                largo = max([len(str(c))] + [len(str(x)) for x in d[c].head(200)])
                ws.set_column(j, j, min(max(largo+2, 10), 40),
                              f_num if pd.api.types.is_numeric_dtype(d[c]) else f_txt)
            ws.freeze_panes(1, 0)

print("--- ARCHIVOS GENERADOS ---")
for mes, d in df[df["Mes Num"].notna()].groupby("Mes Num"):
    mes = int(mes)
    nombre  = MESES[mes]
    carpeta = os.path.join(SALIDA, f"{mes:02d}. {nombre.capitalize()}")
    os.makedirs(carpeta, exist_ok=True)
    ruta = os.path.join(carpeta, f"Consolidado_con_Retencion-{mes:02d}_{nombre}_{ANIO}.xlsx")

    t  = d["Tarifa Retención (%)"]
    ps = d["Coljuegos"].eq("SI")
    ns = d["Coljuegos"].eq("NO")
    escribir(ruta, {
        "Consolidado General": d,
        "Premios20%": d[t == 20.0],
        "Otros ingresos tributarios3.5%": d[t == 3.5],
        "Premios Sin Retencion": d[ps & t.eq(0)],
        "Otros Ingresos Sin retencion": d[ns & t.eq(0)],
        "REVISAR": d[d["Base"].isna() | (d["Concepto"] == "")],
    })
    print(f"  {nombre:<12} {len(d):>7,} filas  ->  {os.path.basename(ruta)}")

#BLOQUE 4
mes_arch = (df.dropna(subset=["Mes Num"]).groupby("Archivo Origen")["Mes Num"]
              .agg(lambda s: int(s.mode().iloc[0])).to_dict())

print(f"Moviendo {len(archivos)} archivos...\n")
for a in archivos:
    base = os.path.basename(a)
    m = mes_arch.get(base)
    sub = f"{m:02d}. {MESES[m].capitalize()}" if m else "SIN MES"
    carpeta = os.path.join(PROCESADOS, sub)
    os.makedirs(carpeta, exist_ok=True)

    destino = os.path.join(carpeta, base)
    if os.path.exists(destino):
        raiz, ext = os.path.splitext(base)
        destino = os.path.join(carpeta, f"{raiz}__{pd.Timestamp.now():%Y%m%d_%H%M%S}{ext}")

    shutil.copy2(a, destino)
    if os.path.exists(destino) and os.path.getsize(destino) > 0:
        os.remove(a)
        print(f"  {sub}/{os.path.basename(destino)}")
    else:
        print(f"  [ERROR] No se movió: {base}")

#BLOQUE 5 

archivos_mes = sorted(a for a in glob.glob(
        os.path.join(SALIDA, "**", "Consolidado_con_Retencion-*.xlsx"), recursive=True)
    if "ANUAL" not in os.path.basename(a).upper()
    and "Historico" not in a
    and not os.path.basename(a).startswith("~$"))
if not archivos_mes:
    raise FileNotFoundError(f"No hay consolidados en {SALIDA}")

partes = []
for a in archivos_mes:
    d = pd.read_excel(a, sheet_name="Consolidado General")
    partes.append(d)
    print(f"  OK {os.path.basename(a)} -> {len(d):,} filas")

anual = pd.concat(partes, ignore_index=True).drop_duplicates()
anual = anual.sort_values("Mes Num", na_position="last").reset_index(drop=True)
print(f"\nTotal anual: {len(anual):,} filas")

hay = set(anual["Mes Num"].dropna().astype(int))
sin = [MESES[m] for m in range(1,13) if m not in hay]
if sin: print(f"[AVISO] Meses sin datos: {sin}")

resumen = (anual.groupby("Mes Num", dropna=False)
                .agg(Filas=("Base","size"), Base_Total=("Base","sum"),
                     Retencion_Total=("Valor Retención","sum")).reset_index())
resumen["Mes"] = resumen["Mes Num"].map(MESES)
resumen = resumen[["Mes","Filas","Base_Total","Retencion_Total"]]
print("\n--- RESUMEN ANUAL ---")
print(resumen.to_string(index=False, float_format=lambda x: f"{x:,.2f}"))

# ---------- FORMATO 1001 ----------
CAMPOS = ["ORIGEN INFORMACION","DOCUMENTO","TIPO DE DOCUMENTO","DIGITO CHEQUEO",
          "APELLIDO1","APELLIDO2","NOMBRE1","NOMBRE2","RAZON SOCIAL","DIRECCION",
          "DEPTO Y MUNICIPIO","PAIS","FORMATO","CONCEPTO","ETIQUETA","VALOR"]

hojas_anual = {"Resumen por Mes": resumen}
pendientes = pd.DataFrame()

if not COL_DOC or COL_DOC not in anual.columns:
    print(f"\n[1001] No se generó: falta definir COL_DOC. Hoy vale '{COL_DOC}'.")
    ex = pd.DataFrame(columns=CAMPOS)
else:
    anual["_doc"] = anual[COL_DOC].map(texto)

    # completar con el maestro si existe
    if os.path.exists(MAESTRO):
        maestro = pd.read_excel(MAESTRO)
        maestro["_doc"] = maestro[maestro.columns[0]].map(texto)
        anual = anual.merge(maestro.drop(columns=[maestro.columns[0]]),
                            on="_doc", how="left", suffixes=("", "_m"))
        print(f"[1001] Maestro cargado: {len(maestro):,} terceros")
    else:
        print(f"[1001] Sin maestro en {MAESTRO} (opcional)")

    val = anual[anual["_doc"].ne("") & anual["Concepto"].ne("")]
    descartadas = len(anual) - len(val)
    if descartadas:
        print(f"[1001] {descartadas:,} filas sin documento o sin concepto quedan fuera")

    tot = (val.groupby(["_doc","Concepto"], dropna=False)
              .agg(Pago=("Base","sum"), Rete=("Valor Retención","sum")).reset_index())

    # identificación: primer valor no vacío por tercero
    ident = pd.DataFrame({"_doc": tot["_doc"].unique()})
    for campo, col in MAPA.items():
        if col and col in val.columns:
            aux = (val.assign(_v=val[col].map(texto))
                      .query("_v != ''").groupby("_doc")["_v"].first())
            ident[campo] = ident["_doc"].map(aux).fillna("")
        else:
            ident[campo] = ""

    tot = tot.merge(ident, on="_doc", how="left")

    filas = []
    for _, r in tot.iterrows():
        for et, v in [(ET_PAGO, r["Pago"]), (ET_RETE, r["Rete"])]:
            if et == ET_RETE and not RETE_EN_CERO and (pd.isna(v) or v == 0):
                continue
            f = {c: "" for c in CAMPOS}
            f.update({c: r.get(c, "") for c in MAPA})
            f["ORIGEN INFORMACION"] = ORIGEN
            f["DOCUMENTO"] = r["_doc"]
            f["FORMATO"]   = FORMATO
            f["CONCEPTO"]  = r["Concepto"]
            f["ETIQUETA"]  = et
            f["VALOR"]     = 0 if pd.isna(v) else round(v)
            filas.append(f)

    ex = pd.DataFrame(filas, columns=CAMPOS)
    print(f"[1001] {len(ex):,} filas  |  {tot['_doc'].nunique():,} terceros")
    print(ex.groupby(["CONCEPTO","ETIQUETA"])["VALOR"]
            .agg(["size","sum"]).to_string(float_format=lambda x: f"{x:,.0f}"))

    clave = ["TIPO DE DOCUMENTO","APELLIDO1","NOMBRE1","RAZON SOCIAL","DIRECCION"]
    vacio = ex[clave].apply(lambda c: c.astype(str).str.strip() == "").any(axis=1)
    pendientes = (ex[vacio].groupby("DOCUMENTO")[clave]
                    .first().reset_index())
    if len(pendientes):
        print(f"\n[1001] {len(pendientes):,} terceros SIN DATOS -> consultar en SAS/BigQuery")

    hojas_anual["Formato 1001"] = ex
    if len(pendientes):
        hojas_anual["TERCEROS_SIN_DATOS"] = pendientes

t = anual["Tarifa Retención (%)"]
if (t == 20.0).any(): hojas_anual["Premios20%"] = anual[t == 20.0]
if (t == 3.5).any():  hojas_anual["Otros3.5%"]  = anual[t == 3.5]

ruta_anual = os.path.join(SALIDA, f"CONSOLIDADO_ANUAL_{ANIO}_RESUMEN.xlsx")
escribir(ruta_anual, hojas_anual)
print(f"\nGuardado: {ruta_anual}")

anual.drop(columns=[c for c in ["_doc"] if c in anual.columns]) \
     .to_csv(os.path.join(SALIDA, f"CONSOLIDADO_ANUAL_{ANIO}_COMPLETO.csv"),
             index=False, encoding="utf-8")
print(f"Guardado: CONSOLIDADO_ANUAL_{ANIO}_COMPLETO.csv")

#BLOQUE 6
