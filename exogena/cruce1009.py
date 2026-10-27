# ===========================================================
# COMPARADOR XML FORMATO 1009 (Saldos ctas por pagar)
# Reproceso 1 vs Reproceso 2 - cruce por cpt + nid
# ===========================================================
ID_CARPETA_R1 = "PEGA_AQUI_EL_ID"     # carpeta reproceso 1
ID_CARPETA_R2 = "PEGA_AQUI_EL_ID"     # carpeta reproceso 2
SALIDA = "comparacion_1009.xlsx"

import io
import xml.etree.ElementTree as ET
from collections import Counter
import pandas as pd
from google.colab import auth, files
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

auth.authenticate_user()
drive = build("drive", "v3")

CAMPOS = ["cpt","tdoc","nid","dv","apl1","apl2","nom1","nom2","raz",
          "dir","dpto","mun","pais","sal"]
CAMPOS_CAB = ["Ano","CodCpt","Formato","Version","NumEnvio","FecEnvio","FecInicial",
              "FecFinal","ValorTotal","CantReg"]
NUMERICOS = {"cpt","tdoc","nid","dv","dpto","mun","pais","sal"}   # nid incluido: quita ceros a la izquierda

# ---------- lectura ----------
def limpiar(campo, valor):
    v = (valor or "").strip()
    if campo in NUMERICOS and v.isdigit():
        return str(int(v))
    return " ".join(v.upper().split())

def entero(texto):
    t = (texto or "").strip()
    return int(t) if t.isdigit() else None

def listar_xml(id_carpeta):
    lista, token = [], None
    while True:
        r = drive.files().list(
            q=f"'{id_carpeta}' in parents and trashed=false",
            fields="nextPageToken, files(id,name)", pageSize=1000, pageToken=token,
            supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
        lista += [a for a in r.get("files", []) if a["name"].lower().endswith(".xml")]
        token = r.get("nextPageToken")
        if not token:
            return lista

def bajar(id_archivo):
    buffer = io.BytesIO()
    descarga = MediaIoBaseDownload(buffer, drive.files().get_media(fileId=id_archivo))
    fin = False
    while not fin:
        _, fin = descarga.next_chunk()
    return buffer.getvalue()

def leer_carpeta(id_carpeta, origen):
    registros, cabeceras = [], []
    for archivo in listar_xml(id_carpeta):
        raiz = ET.fromstring(bajar(archivo["id"]))
        for nodo in raiz.iter():                       # por si vienen con namespace
            if "}" in nodo.tag:
                nodo.tag = nodo.tag.split("}", 1)[1]
        cab = {"origen": origen, "archivo": archivo["name"]}
        nodo_cab = next(raiz.iter("Cab"), None)        # busca el Cab en cualquier nivel
        for c in CAMPOS_CAB:
            hijo = nodo_cab.find(c) if nodo_cab is not None else None
            cab[c] = (hijo.text or "").strip() if hijo is not None else ""
        cabeceras.append(cab)
        cuantos = 0
        for i, reg in enumerate(raiz.iter("saldoscp"), start=1):   # en cualquier nivel
            registros.append({"origen": origen, "archivo": archivo["name"], "fila": i,
                              "datos": {c: limpiar(c, reg.get(c)) for c in CAMPOS}})
            cuantos = i
        print(f'{origen:12} | {archivo["name"][:45]:45} | registros: {cuantos}')
    return registros, cabeceras

def firma(reg):
    return tuple(reg["datos"][c] for c in CAMPOS)

def suma_sal(registros):
    total, malos = 0.0, 0
    for r in registros:
        try:
            total += float(r["datos"]["sal"])
        except ValueError:
            malos += 1                                 # valor no numerico: se reporta aparte
    return total, malos

reg1, cab1 = leer_carpeta(ID_CARPETA_R1, "REPROCESO 1")
reg2, cab2 = leer_carpeta(ID_CARPETA_R2, "REPROCESO 2")

# ---------- agrupacion por llave ----------
g1 = {}
for r in reg1:
    g1.setdefault((r["datos"]["cpt"], r["datos"]["nid"]), []).append(r)

g2 = {}
for r in reg2:
    cpt, nid = r["datos"]["cpt"], r["datos"]["nid"]
    if (cpt, nid) in g1:
        clave, como = (cpt, nid), "llave exacta"
    elif len(nid) > 1 and (cpt, nid[:-1]) in g1:       # por si el dv quedo pegado al nid
        clave, como = (cpt, nid[:-1]), "nid sin el dv pegado"
    else:
        clave, como = (cpt, nid), "llave exacta"
    r["como"] = como
    g2.setdefault(clave, []).append(r)

# ---------- comparacion ----------
diferencias, repetidas = [], []
iguales = 0

def nombre_deudor(d):
    return (d["raz"] or f'{d["nom1"]} {d["apl1"]}').strip()

def foto_linea(r):
    d = r["datos"]
    return (f'tdoc={d["tdoc"]} | sal={d["sal"]} | dpto={d["dpto"]} | '
            f'mun={d["mun"]} | pais={d["pais"]}')

def anotar(clave, deudor, tipo, etiqueta, val1, val2, ra=None, rb=None):
    diferencias.append({
        "llave (cpt | nid)": " | ".join(clave), "cpt": clave[0], "nid": clave[1],
        "deudor": deudor, "tipo_diferencia": tipo, "etiqueta": etiqueta,
        "valor_REPROCESO_1": val1, "valor_REPROCESO_2": val2,
        "como_cruzo": rb["como"] if rb else "",
        "archivo_R1": ra["archivo"] if ra else "", "fila_R1": ra["fila"] if ra else "",
        "archivo_R2": rb["archivo"] if rb else "", "fila_R2": rb["fila"] if rb else ""})

for clave in sorted(set(g1) | set(g2)):
    lista1 = list(g1.get(clave, []))
    pend2 = list(g2.get(clave, []))
    pend1 = []
    for ra in lista1:                                  # 1) saca los identicos
        pareja = next((rb for rb in pend2 if firma(rb) == firma(ra)), None)
        if pareja is not None:
            pend2.remove(pareja)
            iguales += 1
        else:
            pend1.append(ra)

    if len(lista1) > 1 or len(g2.get(clave, [])) > 1:  # esta llave no deberia repetirse
        repetidas.append({"llave (cpt | nid)": " | ".join(clave),
                          "veces_en_REPROCESO_1": len(lista1),
                          "veces_en_REPROCESO_2": len(g2.get(clave, []))})

    for ra, rb in zip(pend1, pend2):                   # 2) se cruzan pero difieren
        deudor = nombre_deudor(ra["datos"])
        for c in CAMPOS:
            if ra["datos"][c] != rb["datos"][c]:
                anotar(clave, deudor, "Valor diferente", c,
                       ra["datos"][c] or "(vacio)", rb["datos"][c] or "(vacio)", ra, rb)

    for ra in pend1[len(pend2):]:                      # 3) no existe al otro lado
        anotar(clave, nombre_deudor(ra["datos"]),
               "Registro que NO esta en REPROCESO 2", "saldoscp",
               foto_linea(ra), "(no existe)", ra, None)
    for rb in pend2[len(pend1):]:
        anotar(clave, nombre_deudor(rb["datos"]),
               "Registro que NO esta en REPROCESO 1", "saldoscp",
               "(no existe)", foto_linea(rb), None, rb)

# ---------- archivos y encabezados (los nombres no coinciden, van por separado) ----------
filas_arch = []
for cab, registros in [(cab1, reg1), (cab2, reg2)]:
    for c in cab:
        propios = [r for r in registros if r["archivo"] == c["archivo"]]
        real_cant = len(propios)
        real_sal, malos = suma_sal(propios)
        fila = {"origen": c["origen"], "archivo": c["archivo"]}
        fila.update({k: c[k] for k in CAMPOS_CAB})
        fila["registros_contados"] = real_cant
        fila["suma_sal_real"] = real_sal
        fila["sal_no_numericos"] = malos
        cant_cab = entero(c["CantReg"])
        fila["cuadra_CantReg"] = ("sin dato" if cant_cab is None
                                  else "SI" if cant_cab == real_cant else "NO")
        try:
            fila["cuadra_ValorTotal"] = "SI" if abs(float(c["ValorTotal"]) - real_sal) < 1 else "NO"
        except ValueError:
            fila["cuadra_ValorTotal"] = "sin dato"
        filas_arch.append(fila)

# ---------- resumen ----------
s1, malos1 = suma_sal(reg1)
s2, malos2 = suma_sal(reg2)
conteo = Counter(f["tipo_diferencia"] for f in diferencias)
resumen = [
    {"concepto": "Archivos en REPROCESO 1", "valor": len(cab1)},
    {"concepto": "Archivos en REPROCESO 2", "valor": len(cab2)},
    {"concepto": "Registros saldoscp en REPROCESO 1", "valor": len(reg1)},
    {"concepto": "Registros saldoscp en REPROCESO 2", "valor": len(reg2)},
    {"concepto": "Suma sal REPROCESO 1", "valor": s1},
    {"concepto": "Suma sal REPROCESO 2", "valor": s2},
    {"concepto": "Diferencia en suma sal", "valor": s1 - s2},
    {"concepto": "Valores sal no numericos (R1 / R2)", "valor": f"{malos1} / {malos2}"},
    {"concepto": "Lineas identicas", "valor": iguales},
    {"concepto": "Llaves con alguna diferencia",
     "valor": len({f["llave (cpt | nid)"] for f in diferencias})},
    {"concepto": "Total diferencias encontradas", "valor": len(diferencias)}]
resumen += [{"concepto": f"  - {t}", "valor": n} for t, n in sorted(conteo.items())]
resumen += [
    {"concepto": "Cruces con el nid sin el dv pegado",
     "valor": sum(1 for r in reg2 if r["como"] == "nid sin el dv pegado")},
    {"concepto": "Llaves repetidas dentro de una carpeta", "valor": len(repetidas)}]

# ---------- conteo por archivo ----------
por_archivo = []
for origen, registros in [("REPROCESO 1", reg1), ("REPROCESO 2", reg2)]:
    for nombre in sorted({r["archivo"] for r in registros}):
        por_archivo.append({"origen": origen, "archivo": nombre,
                            "registros_leidos": sum(1 for r in registros if r["archivo"] == nombre)})

# ---------- excel ----------
def guardar(libro, datos, hoja):
    df = pd.DataFrame(datos) if datos else pd.DataFrame([{"resultado": "sin diferencias"}])
    df.to_excel(libro, sheet_name=hoja, index=False)

with pd.ExcelWriter(SALIDA, engine="openpyxl") as libro:
    guardar(libro, resumen, "RESUMEN")
    guardar(libro, diferencias, "DIFERENCIAS")
    guardar(libro, por_archivo, "CONTEO_POR_ARCHIVO")
    guardar(libro, filas_arch, "ARCHIVOS")
    guardar(libro, repetidas, "LLAVES_REPETIDAS")
    for hoja in libro.book.worksheets:
        hoja.freeze_panes = "A2"
        hoja.auto_filter.ref = hoja.dimensions
        for columna in hoja.columns:
            largo = max(len(str(celda.value or "")) for celda in columna)
            hoja.column_dimensions[columna[0].column_letter].width = min(max(largo + 2, 12), 60)

print()
print(pd.DataFrame(resumen).to_string(index=False))
files.download(SALIDA)