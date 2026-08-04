#1

!pip install xlsxwriter -q

import pandas as pd, numpy as np, os, re, unicodedata, glob, shutil
from google.colab import drive
drive.mount('/content/drive')

# ==================== PARÁMETROS ====================
ENTRADA    = "/content/drive/My Drive/Optimizacion-Premios/2026/Entrada/"
SALIDA     = "/content/drive/My Drive/Optimizacion-Premios/2026/Salida/"
PROCESADOS = "/content/drive/My Drive/Optimizacion-Premios/2026/Procesados/"
ANIO       = 2026

I_MES   = 0     # columna MES
I_VALOR = 17    # columna VALOR_TOTAL_DEL_PREMIO_BASE
I_SINO  = 9     # columna REGISTRADO_EN_COL_JUEGOS

TOPE_PREMIOS = 2_514_000    # 20%
TOPE_OTROS   =   524_000    # 3.5%
# ====================================================


def num(v):
    """1.234.567,89 | 1,234,567.89 | $ 1.601,65 | (234)"""
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
    """si / Si / SÍ / 1 / x -> SI      |      no / No / N -> NO"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = "".join(c for c in unicodedata.normalize("NFD", str(v))
                if unicodedata.category(c) != "Mn")
    s = s.upper().replace(".", "").replace(" ", "")
    if s in {"SI", "S", "1", "1.0", "X"}:
        return "SI"
    if s in {"NO", "N", "0", "0.0"}:
        return "NO"
    return s


MESES = {1:"enero", 2:"febrero", 3:"marzo", 4:"abril", 5:"mayo", 6:"junio",
         7:"julio", 8:"agosto", 9:"septiembre", 10:"octubre",
         11:"noviembre", 12:"diciembre"}
ALIAS = {v.upper(): k for k, v in MESES.items()}
ALIAS.update({"SETIEMBRE":9, "ENE":1, "FEB":2, "MAR":3, "ABR":4, "MAY":5,
              "JUN":6, "JUL":7, "AGO":8, "SEP":9, "OCT":10, "NOV":11, "DIC":12})

def mesnum(v):
    """JULIO / julio / Julio / jul / 7 -> 7"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = "".join(c for c in unicodedata.normalize("NFD", str(v))
                if unicodedata.category(c) != "Mn").upper().strip().replace(".", "")
    if s in ALIAS:
        return ALIAS[s]
    if s.isdigit() and 1 <= int(s) <= 12:
        return int(s)
    return None


# ---- Lectura ----
for c in (ENTRADA, SALIDA, PROCESADOS):
    os.makedirs(c, exist_ok=True)

archivos = sorted(os.path.join(ENTRADA, f) for f in os.listdir(ENTRADA)
                  if f.lower().endswith((".xlsx", ".xlsm"))
                  and not f.startswith("~$"))

if not archivos:
    raise FileNotFoundError(f"No hay archivos .xlsx en {ENTRADA}")

print(f"Archivos encontrados: {len(archivos)}\n")
frames = []
for a in archivos:
    d = pd.read_excel(a, header=0)
    d["Archivo Origen"] = os.path.basename(a)
    frames.append(d)
    print(f"  {os.path.basename(a)}  ->  {len(d):,} filas")

df = pd.concat(frames, ignore_index=True)
print(f"\nTotal: {len(df):,} filas x {len(df.columns)} columnas\n")

print("=" * 70)
print("VERIFICA QUE ESTAS SEAN LAS COLUMNAS CORRECTAS")
for et, i in [("MES", I_MES), ("VALOR", I_VALOR), ("SI/NO", I_SINO)]:
    print(f"  {et:<6} idx {i:>2} -> '{df.columns[i]}'")
    print(f"          {df.iloc[:4, i].tolist()}")
print("=" * 70)


#2
c_mes  = df.columns[I_MES]
c_val  = df.columns[I_VALOR]
c_sino = df.columns[I_SINO]

df["Base"]      = df[c_val].map(num)
df["Es Premio"] = df[c_sino].map(sino)
df["Mes Num"]   = df[c_mes].map(mesnum)

print("--- CONTROL ---")
print(f"Bases no numéricas : {df['Base'].isna().sum():,}")
print(f"SI/NO no reconocido: {sorted(set(df['Es Premio']) - {'SI','NO',''})}")
print(f"Meses en los datos : {df[c_mes].value_counts().to_dict()}")
print(f"Base mínima / máxima: {df['Base'].min():,.2f} / {df['Base'].max():,.2f}")
print(f"Topes: premios {TOPE_PREMIOS:,.0f}   otros {TOPE_OTROS:,.0f}\n")

df["Tarifa Retención (%)"] = 0.0
b   = df["Base"]
esp = df["Es Premio"].eq("SI")
nop = df["Es Premio"].eq("NO")

df.loc[esp & b.gt(TOPE_PREMIOS), "Tarifa Retención (%)"] = 20.0
df.loc[nop & b.gt(TOPE_OTROS),   "Tarifa Retención (%)"] = 3.5

df["Valor Retención"] = (b * df["Tarifa Retención (%)"] / 100).round(0)
df.loc[b.isna(), "Valor Retención"] = np.nan

t = df["Tarifa Retención (%)"]
print("--- RESUMEN ---")
print(df.groupby(t).agg(Filas=("Base","size"),
                        Base=("Base","sum"),
                        Retencion=("Valor Retención","sum"))
      .to_string(float_format=lambda x: f"{x:,.2f}"), "\n")

mn = df["Mes Num"].dropna()
mn = int(mn.mode().iloc[0]) if len(mn) else 0
nombre_mes = MESES.get(mn, "sinmes")
ruta = os.path.join(SALIDA,
    f"Consolidado_con_Retencion-{mn:02d}_{nombre_mes}_{ANIO}.xlsx")

hojas = {
    "Consolidado General":            df,
    "Premios20%":                     df[t == 20.0],
    "Otros ingresos tributarios3.5%": df[t == 3.5],
    "Premios Sin Retencion":          df[esp & t.eq(0)],
    "Otros Ingresos Sin retencion":   df[nop & t.eq(0)],
    "REVISAR":                        df[b.isna() | ~df["Es Premio"].isin(["SI","NO"])],
}

with pd.ExcelWriter(ruta, engine="xlsxwriter") as w:
    wb = w.book
    f_enc = wb.add_format({"bold":True, "font_color":"FFFFFF", "bg_color":"FF0000",
                           "font_name":"Calibri", "font_size":11, "align":"center",
                           "valign":"vcenter", "text_wrap":True})
    f_num = wb.add_format({"num_format":"#,##0.00", "font_name":"Calibri",
                           "font_size":11, "align":"center"})
    f_txt = wb.add_format({"font_name":"Calibri", "font_size":11, "align":"center"})

    for h, d in hojas.items():
        d.to_excel(w, sheet_name=h[:31], index=False)
        ws = w.sheets[h[:31]]
        for j, c in enumerate(d.columns):
            ws.write(0, j, str(c), f_enc)
            largo = max([len(str(c))] + [len(str(x)) for x in d[c].head(200)])
            ws.set_column(j, j, min(max(largo + 2, 10), 40),
                          f_num if pd.api.types.is_numeric_dtype(d[c]) else f_txt)
        ws.freeze_panes(1, 0)
        print(f"  {h}: {len(d):,} filas")

print(f"\nGuardado: {ruta}")

#3
print(f"Moviendo {len(archivos)} archivos a Procesados...\n")

for a in archivos:
    destino = os.path.join(PROCESADOS, os.path.basename(a))
    shutil.copy2(a, destino)
    if os.path.exists(destino) and os.path.getsize(destino) > 0:
        os.remove(a)
        print(f"  Movido: {os.path.basename(a)}")
    else:
        print(f"  [ERROR] No se movió: {os.path.basename(a)}")

print("\nListo. Archivos en Procesados:")
for f in os.listdir(PROCESADOS):
    print(f"  {f}")

#4
archivos_mes = sorted([a for a in glob.glob(os.path.join(SALIDA,
                        "Consolidado_con_Retencion-*.xlsx"))
                       if "ANUAL" not in os.path.basename(a).upper()
                       and not os.path.basename(a).startswith("~$")])

if not archivos_mes:
    raise FileNotFoundError(f"No hay consolidados mensuales en {SALIDA}")

print(f"Meses encontrados: {len(archivos_mes)}")
for a in archivos_mes:
    print(f"  {os.path.basename(a)}")

partes = []
for a in archivos_mes:
    d = pd.read_excel(a, sheet_name="Consolidado General")
    partes.append(d)
    print(f"  OK {os.path.basename(a)} -> {len(d):,} filas")

df_anual = pd.concat(partes, ignore_index=True)

antes = len(df_anual)
df_anual = df_anual.drop_duplicates()
if antes != len(df_anual):
    print(f"\n[AVISO] {antes - len(df_anual):,} filas duplicadas eliminadas.")

df_anual = df_anual.sort_values("Mes Num", na_position="last").reset_index(drop=True)
print(f"\nTotal anual: {len(df_anual):,} filas")

presentes = set(df_anual["Mes Num"].dropna().astype(int))
faltantes = [MESES[m] for m in range(1, 13) if m not in presentes]
if faltantes:
    print(f"[AVISO] Meses sin datos: {faltantes}")

t = df_anual["Tarifa Retención (%)"]
resumen = (df_anual.groupby("Mes Num", dropna=False)
           .agg(Filas=("Base","size"),
                Base_Total=("Base","sum"),
                Retencion_Total=("Valor Retención","sum"))
           .reset_index())
resumen["Mes"] = resumen["Mes Num"].map(MESES)

print("\n--- RESUMEN ANUAL ---")
print(resumen[["Mes","Filas","Base_Total","Retencion_Total"]]
      .to_string(index=False, float_format=lambda x: f"{x:,.2f}"))

ruta_anual = os.path.join(SALIDA, f"CONSOLIDADO_ANUAL_{ANIO}.xlsx")

hojas_anual = {
    "Resumen por Mes":                resumen,
    "Consolidado Anual":              df_anual,
    "Premios20%":                     df_anual[t == 20.0],
    "Otros ingresos tributarios3.5%": df_anual[t == 3.5],
    "Premios Sin Retencion":          df_anual[df_anual["Es Premio"].eq("SI") & t.eq(0)],
    "Otros Ingresos Sin retencion":   df_anual[df_anual["Es Premio"].eq("NO") & t.eq(0)],
    "REVISAR":                        df_anual[df_anual["Base"].isna()
                                               | ~df_anual["Es Premio"].isin(["SI","NO"])],
}

with pd.ExcelWriter(ruta_anual, engine="xlsxwriter") as w:
    wb = w.book
    f_enc = wb.add_format({"bold":True, "font_color":"FFFFFF", "bg_color":"FF0000",
                           "font_name":"Calibri", "font_size":11, "align":"center",
                           "valign":"vcenter", "text_wrap":True})
    f_num = wb.add_format({"num_format":"#,##0.00", "font_name":"Calibri",
                           "font_size":11, "align":"center"})
    f_txt = wb.add_format({"font_name":"Calibri", "font_size":11, "align":"center"})

    for h, d in hojas_anual.items():
        d.to_excel(w, sheet_name=h[:31], index=False)
        ws = w.sheets[h[:31]]
        for j, c in enumerate(d.columns):
            ws.write(0, j, str(c), f_enc)
            largo = max([len(str(c))] + [len(str(x)) for x in d[c].head(200)])
            ws.set_column(j, j, min(max(largo + 2, 10), 40),
                          f_num if pd.api.types.is_numeric_dtype(d[c]) else f_txt)
        ws.freeze_panes(1, 0)
        print(f"  {h}: {len(d):,} filas")

print(f"\nGuardado: {ruta_anual}")