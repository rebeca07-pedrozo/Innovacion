# ===========================================================
# COMPARADOR XML FORMATO 1020 (CDT) - DIAN vs carpeta jefa
# Cruza por nid + ntit, sin importar orden, archivo ni nivel del XML
# ===========================================================
ID_CARPETA_DIAN = "PEGA_AQUI_EL_ID"   # XML que descargué de la DIAN
ID_CARPETA_JEFA = "PEGA_AQUI_EL_ID"   # carpeta a verificar
SALIDA = "comparacion_1020.xlsx"

import io, hashlib
import xml.etree.ElementTree as ET
from collections import Counter
import pandas as pd
from google.colab import auth, files
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

auth.authenticate_user()
drive = build("drive", "v3")

CAMPOS = ["tdoc","nid","dv","apl1","apl2","nom1","nom2","raz","dir","dpto","mun",
          "pais","ntit","ttitu","tmov","salini","inv","vintca","vintpa","retfup","salfin"]
CAMPOS_SEC = ["cpts","tdocs","nids","dvs","apl1s","apl2s","nom1s","nom2s","razs"]
CAMPOS_CAB = ["Ano","CodCpt","Formato","Version","NumEnvio","FecEnvio","FecInicial",
              "FecFinal","ValorTotal","CantReg"]
NUMERICOS = {"tdoc","nid","dv","dpto","mun","pais","ntit","ttitu","tmov","salini","inv",
             "vintca","vintpa","retfup","salfin","cpts","tdocs","nids","dvs"}

# ---------- lectura ----------
def limpiar(campo, valor):
    v = (valor or "").strip()
    if campo in NUMERICOS and v.isdigit():
        return str(int(v))          # quita ceros a la izquierda
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
    por_huella, por_nombre = {}, {}                    # para detectar duplicados de Drive
    for archivo in listar_xml(id_carpeta):
        crudo = bajar(archivo["id"])
        huella = hashlib.md5(crudo).hexdigest()

        if huella in por_huella:                       # mismo contenido: se omite
            print(f'{origen:5} | {archivo["name"][:50]:50} | DUPLICADO, se omite')
            continue
        if archivo["name"] in por_nombre:              # mismo nombre, contenido distinto
            print(f'{origen:5} | {archivo["name"][:50]:50} | OJO: mismo nombre, '
                  f'contenido DISTINTO. Revisar en Drive')
        por_huella[huella] = archivo["name"]
        por_nombre[archivo["name"]] = huella

        raiz = ET.fromstring(crudo)
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
        for i, inv in enumerate(raiz.iter("invcdt"), start=1):   # en cualquier nivel
            registros.append({
                "origen": origen, "archivo": archivo["name"], "fila": i,
                "datos": {c: limpiar(c, inv.get(c)) for c in CAMPOS},
                "sec": [{c: limpiar(c, s.get(c)) for c in CAMPOS_SEC}
                        for s in inv.iter("titSec")]})
            cuantos = i
        print(f'{origen:5} | {archivo["name"][:50]:50} | registros: {cuantos}')
    return registros, cabeceras

def firma_sec(reg):
    return sorted(tuple(s[c] for c in CAMPOS_SEC) for s in reg["sec"])

def firma(reg):
    return (tuple(reg["datos"][c] for c in CAMPOS), tuple(firma_sec(reg)))

reg_dian, cab_dian = leer_carpeta(ID_CARPETA_DIAN, "DIAN")
reg_jefa, cab_jefa = leer_carpeta(ID_CARPETA_JEFA, "JEFA")

# ---------- agrupacion por llave ----------
g_dian = {}
for r in reg_dian:
    g_dian.setdefault((r["datos"]["nid"], r["datos"]["ntit"]), []).append(r)

g_jefa = {}
for r in reg_jefa:
    nid, ntit = r["datos"]["nid"], r["datos"]["ntit"]
    if (nid, ntit) in g_dian:
        clave, como = (nid, ntit), "llave exacta"
    elif len(nid) > 1 and (nid[:-1], ntit) in g_dian:      # el dv quedo pegado al nid
        clave, como = (nid[:-1], ntit), "nid sin el dv pegado"
    else:
        clave, como = (nid, ntit), "llave exacta"
    r["como"] = como
    g_jefa.setdefault(clave, []).append(r)

# ---------- comparacion ----------
diferencias, filas_cab, repetidas = [], [], []
iguales = 0

def nombre_titular(d):
    return (d["raz"] or f'{d["nom1"]} {d["apl1"]}').strip()

def nombre_sec(s):
    nombre = s["razs"] or " ".join(x for x in [s["nom1s"], s["apl1s"], s["apl2s"]] if x)
    return f'{nombre} ({s["nids"]})'.strip()

def foto_linea(r):
    d = r["datos"]
    return (f'ttitu={d["ttitu"]} | tmov={d["tmov"]} | salini={d["salini"]} | '
            f'inv={d["inv"]} | vintca={d["vintca"]} | vintpa={d["vintpa"]} | '
            f'retfup={d["retfup"]} | salfin={d["salfin"]} | titSec={len(r["sec"])}')

def anotar(clave, titular, tipo, etiqueta, val_dian, val_jefa, ra=None, rb=None):
    diferencias.append({
        "llave (nid | ntit)": " | ".join(clave), "nid": clave[0], "ntit": clave[1],
        "titular": titular, "tipo_diferencia": tipo, "etiqueta": etiqueta,
        "valor_DIAN": val_dian, "valor_JEFA": val_jefa,
        "como_cruzo": rb["como"] if rb else "",
        "archivo_DIAN": ra["archivo"] if ra else "", "fila_DIAN": ra["fila"] if ra else "",
        "archivo_JEFA": rb["archivo"] if rb else "", "fila_JEFA": rb["fila"] if rb else ""})

for clave in sorted(set(g_dian) | set(g_jefa)):
    lista_dian = list(g_dian.get(clave, []))
    pend_jefa = list(g_jefa.get(clave, []))
    pend_dian = []
    for ra in lista_dian:                              # 1) saca los identicos
        pareja = next((rb for rb in pend_jefa if firma(rb) == firma(ra)), None)
        if pareja is not None:
            pend_jefa.remove(pareja)
            iguales += 1
        else:
            pend_dian.append(ra)

    if len(lista_dian) > 1 or len(g_jefa.get(clave, [])) > 1:   # llave repetida
        repetidas.append({"llave (nid | ntit)": " | ".join(clave),
                          "veces_en_DIAN": len(lista_dian),
                          "veces_en_JEFA": len(g_jefa.get(clave, []))})

    disponibles = list(pend_jefa)                      # 2) empareja por la linea mas parecida
    parejas = []
    for ra in pend_dian:
        if not disponibles:
            break
        rb = min(disponibles,
                 key=lambda x: sum(1 for c in CAMPOS if ra["datos"][c] != x["datos"][c]))
        disponibles.remove(rb)
        parejas.append((ra, rb))

    for ra, rb in parejas:
        titular = nombre_titular(ra["datos"])
        for c in CAMPOS:
            if ra["datos"][c] != rb["datos"][c]:
                anotar(clave, titular, "Valor diferente", c,
                       ra["datos"][c] or "(vacio)", rb["datos"][c] or "(vacio)", ra, rb)
        sa = Counter(tuple(s[c] for c in CAMPOS_SEC) for s in ra["sec"])
        sb = Counter(tuple(s[c] for c in CAMPOS_SEC) for s in rb["sec"])
        for s in (sb - sa).elements():
            anotar(clave, titular, "Etiqueta que NO esta en el archivo de la DIAN", "titSec",
                   "(no existe)", nombre_sec(dict(zip(CAMPOS_SEC, s))), ra, rb)
        for s in (sa - sb).elements():
            anotar(clave, titular, "Etiqueta que NO esta en el archivo de la JEFA", "titSec",
                   nombre_sec(dict(zip(CAMPOS_SEC, s))), "(no existe)", ra, rb)

    for ra in pend_dian[len(parejas):]:                # 3) no existen al otro lado
        anotar(clave, nombre_titular(ra["datos"]),
               "Registro que NO esta en el archivo de la JEFA", "invcdt",
               foto_linea(ra), "(no existe)", ra, None)
    for rb in disponibles:
        anotar(clave, nombre_titular(rb["datos"]),
               "Registro que NO esta en el archivo de la DIAN", "invcdt",
               "(no existe)", foto_linea(rb), None, rb)

# ---------- encabezados (solo lo que no coincide) ----------
nombres = sorted({c["archivo"] for c in cab_dian} | {c["archivo"] for c in cab_jefa})
for nombre in nombres:
    a = next((c for c in cab_dian if c["archivo"] == nombre), None)
    b = next((c for c in cab_jefa if c["archivo"] == nombre), None)
    for c in CAMPOS_CAB:
        va = a[c] if a else "(no esta en la carpeta DIAN)"
        vb = b[c] if b else "(no esta en la carpeta JEFA)"
        if va != vb:
            filas_cab.append({"archivo": nombre, "etiqueta": c,
                              "valor_DIAN": va, "valor_JEFA": vb})

# ---------- conteo por archivo ----------
por_archivo = []
for origen, registros in [("DIAN", reg_dian), ("JEFA", reg_jefa)]:
    for nombre in sorted({r["archivo"] for r in registros}):
        por_archivo.append({"origen": origen, "archivo": nombre,
                            "registros_leidos": sum(1 for r in registros if r["archivo"] == nombre)})

# ---------- resumen ----------
conteo = Counter(f["tipo_diferencia"] for f in diferencias)
resumen = [
    {"concepto": "Archivos en carpeta DIAN", "valor": len(cab_dian)},
    {"concepto": "Archivos en carpeta JEFA", "valor": len(cab_jefa)},
    {"concepto": "Registros invcdt en DIAN", "valor": len(reg_dian)},
    {"concepto": "Registros invcdt en JEFA", "valor": len(reg_jefa)},
    {"concepto": "Lineas identicas", "valor": iguales},
    {"concepto": "Llaves con alguna diferencia",
     "valor": len({f["llave (nid | ntit)"] for f in diferencias})},
    {"concepto": "Total diferencias encontradas", "valor": len(diferencias)}]
resumen += [{"concepto": f"  - {t}", "valor": n} for t, n in sorted(conteo.items())]
resumen += [
    {"concepto": "Cruces con el nid sin el dv pegado",
     "valor": sum(1 for r in reg_jefa if r["como"] == "nid sin el dv pegado")},
    {"concepto": "Llaves repetidas dentro de una carpeta", "valor": len(repetidas)},
    {"concepto": "Encabezados con diferencia", "valor": len(filas_cab)}]

# ---------- excel ----------
def guardar(libro, datos, hoja):
    df = pd.DataFrame(datos) if datos else pd.DataFrame([{"resultado": "sin diferencias"}])
    df.to_excel(libro, sheet_name=hoja, index=False)

with pd.ExcelWriter(SALIDA, engine="openpyxl") as libro:
    guardar(libro, resumen, "RESUMEN")
    guardar(libro, diferencias, "DIFERENCIAS")
    guardar(libro, por_archivo, "CONTEO_POR_ARCHIVO")
    guardar(libro, repetidas, "LLAVES_REPETIDAS")
    guardar(libro, filas_cab, "ENCABEZADOS")
    for hoja in libro.book.worksheets:                 # ancho de columna y filtro
        hoja.freeze_panes = "A2"
        hoja.auto_filter.ref = hoja.dimensions
        for columna in hoja.columns:
            largo = max(len(str(celda.value or "")) for celda in columna)
            hoja.column_dimensions[columna[0].column_letter].width = min(max(largo + 2, 12), 60)

print()
print(pd.DataFrame(resumen).to_string(index=False))
files.download(SALIDA)