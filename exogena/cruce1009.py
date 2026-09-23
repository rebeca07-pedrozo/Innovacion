ID_CARPETA_R1 = "PEGA_AQUI_EL_ID"    
ID_CARPETA_R2 = "PEGA_AQUI_EL_ID"    
SALIDA = "comparacion_1009.xlsx"
CLAVE = ["cpt", "nid"]               
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
NUMERICOS = {"cpt","tdoc","dv","dpto","mun","pais","sal"}
def limpiar(campo, valor):
    v = (valor or "").strip()
    if campo in NUMERICOS and v.isdigit():
        return str(int(v))          
    return " ".join(v.upper().split())
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
        for nodo in raiz.iter():                    
            if "}" in nodo.tag:
                nodo.tag = nodo.tag.split("}", 1)[1]
        cab = {"origen": origen, "archivo": archivo["name"]}
        nodo_cab = raiz.find("Cab")
        for c in CAMPOS_CAB:
            hijo = nodo_cab.find(c) if nodo_cab is not None else None
            cab[c] = (hijo.text or "").strip() if hijo is not None else ""
        cabeceras.append(cab)
        for i, reg in enumerate(raiz.findall("saldoscp"), start=1):
            registros.append({"origen": origen, "archivo": archivo["name"], "fila": i,
                              "datos": {c: limpiar(c, reg.get(c)) for c in CAMPOS}})
    return registros, cabeceras

def firma(reg):
    return tuple(reg["datos"][c] for c in CAMPOS)

def agrupar(registros):
    grupos = {}
    for r in registros:
        clave = tuple(r["datos"][c] for c in CLAVE)
        grupos.setdefault(clave, []).append(r)
    return grupos

reg1, cab1 = leer_carpeta(ID_CARPETA_R1, "REPROCESO 1")
reg2, cab2 = leer_carpeta(ID_CARPETA_R2, "REPROCESO 2")
g1, g2 = agrupar(reg1), agrupar(reg2)

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

    if len(lista1) > 1 or len(g2.get(clave, [])) > 1:  # llave que no deberia repetirse
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

# ---------- archivos y encabezados (nombres distintos, van por separado) ----------
def suma_sal(registros):
    return sum(int(r["datos"]["sal"]) for r in registros if r["datos"]["sal"].isdigit())

filas_arch = []
for cab, registros in [(cab1, reg1), (cab2, reg2)]:
    for c in cab:
        propios = [r for r in registros if r["archivo"] == c["archivo"]]
        real_cant, real_sal = len(propios), suma_sal(propios)
        fila = {"origen": c["origen"], "archivo": c["archivo"]}
        fila.update({k: c[k] for k in CAMPOS_CAB})
        fila["registros_contados"] = real_cant
        fila["suma_sal_real"] = real_sal
        fila["cuadra_CantReg"] = "SI" if str(real_cant) == c["CantReg"].lstrip("0") else "NO"
        try:
            fila["cuadra_ValorTotal"] = "SI" if abs(float(c["ValorTotal"]) - real_sal) < 1 else "NO"
        except ValueError:
            fila["cuadra_ValorTotal"] = "NO"
        filas_arch.append(fila)

# ---------- resumen ----------
conteo = Counter(f["tipo_diferencia"] for f in diferencias)
resumen = [
    {"concepto": "Archivos en REPROCESO 1", "valor": len(cab1)},
    {"concepto": "Archivos en REPROCESO 2", "valor": len(cab2)},
    {"concepto": "Registros saldoscp en REPROCESO 1", "valor": len(reg1)},
    {"concepto": "Registros saldoscp en REPROCESO 2", "valor": len(reg2)},
    {"concepto": "Suma sal REPROCESO 1", "valor": suma_sal(reg1)},
    {"concepto": "Suma sal REPROCESO 2", "valor": suma_sal(reg2)},
    {"concepto": "Diferencia en suma sal", "valor": suma_sal(reg1) - suma_sal(reg2)},
    {"concepto": "Lineas identicas", "valor": iguales},
    {"concepto": "Llaves con alguna diferencia",
     "valor": len({f["llave (cpt | nid)"] for f in diferencias})},
    {"concepto": "Total diferencias encontradas", "valor": len(diferencias)}]
resumen += [{"concepto": f"  - {t}", "valor": n} for t, n in sorted(conteo.items())]
resumen += [{"concepto": "Llaves repetidas dentro de una carpeta", "valor": len(repetidas)}]

# ---------- excel ----------
def guardar(libro, datos, hoja):
    df = pd.DataFrame(datos) if datos else pd.DataFrame([{"resultado": "sin diferencias"}])
    df.to_excel(libro, sheet_name=hoja, index=False)

with pd.ExcelWriter(SALIDA, engine="openpyxl") as libro:
    guardar(libro, resumen, "RESUMEN")
    guardar(libro, diferencias, "DIFERENCIAS")
    guardar(libro, filas_arch, "ARCHIVOS")
    guardar(libro, repetidas, "LLAVES_REPETIDAS")
    for hoja in libro.book.worksheets:                 # ancho de columna y filtro
        hoja.freeze_panes = "A2"
        hoja.auto_filter.ref = hoja.dimensions
        for columna in hoja.columns:
            largo = max(len(str(celda.value or "")) for celda in columna)
            hoja.column_dimensions[columna[0].column_letter].width = min(max(largo + 2, 12), 60)

print(pd.DataFrame(resumen).to_string(index=False))
files.download(SALIDA)