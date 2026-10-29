# ===========================================================
# COMPARADOR XML FORMATO 1009 - DIAN vs carpeta de reproceso
# Cruza por cpt + nid, en las dos direcciones.
# Reporte de una sola hoja: DIFERENCIAS. El resumen va en pantalla.
# ===========================================================
ID_CARPETA_DIAN      = "PEGA_AQUI_EL_ID"   # XML descargados de la DIAN
ID_CARPETA_REPROCESO = "PEGA_AQUI_EL_ID"   # carpeta del reproceso a revisar
NOMBRE_B = "REPROCESO 5"                   # como se llama en el reporte
SALIDA = "diferencias_1009.xlsx"

import io, re, hashlib, unicodedata
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
NUMERICOS = {"cpt","tdoc","nid","dv","dpto","mun","pais","sal"}
MARCAS = ("Ã", "Â", "â€")              # señales de texto partido

# ---------- lectura robusta ----------
def leer_xml(crudo):
    """Devuelve (raiz, lo_que_declara, lo_que_realmente_es).
       No le cree a la declaracion y repara la declaracion dañada."""
    marca = re.search(rb'encoding="([^"]+)"', crudo[:300])
    declara = marca.group(1).decode("ascii", "replace") if marca else "(no declara)"
    try:
        texto, real = crudo.decode("utf-8"), "UTF-8"
    except UnicodeDecodeError:
        texto, real = crudo.decode("latin-1"), "ISO-8859-1"
    texto = texto.lstrip("\ufeff")                              # quita el BOM
    texto = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', texto)  # lo que XML no admite
    roto = re.match(r'^<(.)xml\s+version=.*?\1>', texto, flags=re.DOTALL)
    if roto and roto.group(1) != "?":                           # <Ñxml ... Ñ>
        texto = texto[roto.end():]
    if re.match(r'\s*<\?xml', texto):
        texto = re.sub(r'\s*<\?xml.*?\?>', '<?xml version="1.0" encoding="UTF-8"?>',
                       texto, count=1, flags=re.DOTALL)
    else:
        texto = '<?xml version="1.0" encoding="UTF-8"?>\n' + texto
    return ET.fromstring(texto.encode("utf-8")), declara, real

def arreglar(texto):
    """BRICEÃO -> BRICEÑO. Solo entra si el texto trae las marcas del daño."""
    salida = texto
    for _ in range(3):
        if not any(m in salida for m in MARCAS):
            break
        try:
            nuevo = salida.encode("latin-1").decode("utf-8")
        except (UnicodeEncodeError, UnicodeDecodeError):
            break
        if nuevo == salida:
            break
        salida = nuevo
    return unicodedata.normalize("NFC", salida)

def solo_caracter_perdido(x, y):
    """Los dos textos son el mismo salvo donde un lado tiene el caracter perdido."""
    return len(x) == len(y) and all(a == b or "\ufffd" in (a, b) for a, b in zip(x, y))

def limpiar(campo, valor):
    v = arreglar((valor or "").strip())
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

avisos = []                         # problemas de lectura, duplicados o codificacion

def leer_carpeta(id_carpeta, origen):
    registros, archivos = [], []
    por_huella, por_nombre = {}, {}
    for archivo in listar_xml(id_carpeta):
        crudo = bajar(archivo["id"])
        huella = hashlib.md5(crudo).hexdigest()

        if huella in por_huella:                       # duplicado exacto en Drive
            avisos.append({"origen": origen, "archivo": archivo["name"],
                           "aviso": f'DUPLICADO de {por_huella[huella]}, se omite'})
            print(f'{origen:12} | {archivo["name"][:45]:45} | DUPLICADO, se omite')
            continue
        if archivo["name"] in por_nombre:
            avisos.append({"origen": origen, "archivo": archivo["name"],
                           "aviso": "mismo nombre, contenido DISTINTO"})
        por_huella[huella] = archivo["name"]
        por_nombre[archivo["name"]] = huella

        try:
            raiz, declara, real = leer_xml(crudo)
        except Exception as e:                         # no se cae: reporta y sigue
            avisos.append({"origen": origen, "archivo": archivo["name"],
                           "aviso": f"NO SE PUDO LEER: {e}"})
            print(f'{origen:12} | {archivo["name"][:45]:45} | NO SE PUDO LEER: {e}')
            continue

        for nodo in raiz.iter():                       # por si vienen con namespace
            if "}" in nodo.tag:
                nodo.tag = nodo.tag.split("}", 1)[1]

        archivos.append(archivo["name"])
        cuantos, partidos = 0, 0
        for i, reg in enumerate(raiz.iter("saldoscp"), start=1):
            crudos = {c: reg.get(c) for c in CAMPOS}
            partidos += sum(1 for v in crudos.values() if v and any(m in v for m in MARCAS))
            registros.append({"origen": origen, "archivo": archivo["name"], "fila": i,
                              "datos": {c: limpiar(c, crudos[c]) for c in CAMPOS}})
            cuantos = i

        normal = declara.upper().replace("-", "") == real.upper().replace("-", "")
        if not normal or partidos:
            avisos.append({"origen": origen, "archivo": archivo["name"],
                           "aviso": f"declara {declara} pero es {real}; "
                                    f"valores con texto partido: {partidos}"})
        extra = "" if (normal and not partidos) else f' | CODIFICACION: declara {declara}, es {real}'
        print(f'{origen:12} | {archivo["name"][:45]:45} | registros: {cuantos}{extra}')
    return registros, archivos

reg_dian, arch_dian = leer_carpeta(ID_CARPETA_DIAN, "DIAN")
reg_rep,  arch_rep  = leer_carpeta(ID_CARPETA_REPROCESO, NOMBRE_B)

def firma(reg):
    return tuple(reg["datos"][c] for c in CAMPOS)

# ---------- indices por llave (cpt + nid) ----------
g_dian = {}
for r in reg_dian:
    g_dian.setdefault((r["datos"]["cpt"], r["datos"]["nid"]), []).append(r)

g_rep = {}
for r in reg_rep:
    cpt, nid = r["datos"]["cpt"], r["datos"]["nid"]
    if (cpt, nid) in g_dian:
        clave, como = (cpt, nid), "llave exacta"
    elif len(nid) > 1 and (cpt, nid[:-1]) in g_dian:   # por si el dv quedo pegado al nid
        clave, como = (cpt, nid[:-1]), "nid sin el dv pegado"
    else:
        clave, como = (cpt, nid), "llave exacta"
    r["como"] = como
    g_rep.setdefault(clave, []).append(r)

# ---------- comparacion en las dos direcciones ----------
diferencias = []
iguales = 0

def nombre_deudor(d):
    return (d["raz"] or f'{d["nom1"]} {d["apl1"]}').strip()

def foto_linea(r):
    d = r["datos"]
    return (f'tdoc={d["tdoc"]} | sal={d["sal"]} | dir={d["dir"]} | '
            f'dpto={d["dpto"]} | mun={d["mun"]} | pais={d["pais"]}')

def anotar(clave, deudor, tipo, etiqueta, val_dian, val_rep, ra=None, rb=None):
    diferencias.append({
        "llave (cpt | nid)": " | ".join(clave), "cpt": clave[0], "nid": clave[1],
        "deudor": deudor, "tipo_diferencia": tipo, "etiqueta": etiqueta,
        "valor_DIAN": val_dian, f"valor_{NOMBRE_B}": val_rep,
        "como_cruzo": rb["como"] if rb else "",
        "archivo_DIAN": ra["archivo"] if ra else "", "fila_DIAN": ra["fila"] if ra else "",
        f"archivo_{NOMBRE_B}": rb["archivo"] if rb else "",
        f"fila_{NOMBRE_B}": rb["fila"] if rb else ""})

for clave in sorted(set(g_dian) | set(g_rep)):
    lista_dian = list(g_dian.get(clave, []))
    pend_rep = list(g_rep.get(clave, []))
    pend_dian = []
    for ra in lista_dian:                              # 1) saca las identicas
        pareja = next((rb for rb in pend_rep if firma(rb) == firma(ra)), None)
        if pareja is not None:
            pend_rep.remove(pareja)
            iguales += 1
        else:
            pend_dian.append(ra)

    disponibles = list(pend_rep)                       # 2) empareja por la linea mas parecida
    parejas = []
    for ra in pend_dian:
        if not disponibles:
            break
        rb = min(disponibles,
                 key=lambda x: sum(1 for c in CAMPOS if ra["datos"][c] != x["datos"][c]))
        disponibles.remove(rb)
        parejas.append((ra, rb))

    for ra, rb in parejas:                             # 3) cruzan pero difieren
        deudor = nombre_deudor(ra["datos"])
        for c in CAMPOS:
            x, y = ra["datos"][c], rb["datos"][c]
            if x != y:
                tipo = ("Caracter perdido (no es diferencia de dato)"
                        if solo_caracter_perdido(x, y) else "Valor diferente")
                anotar(clave, deudor, tipo, c, x or "(vacio)", y or "(vacio)", ra, rb)

    for ra in pend_dian[len(parejas):]:                # 4) falta en el reproceso
        anotar(clave, nombre_deudor(ra["datos"]),
               f"Registro que NO esta en {NOMBRE_B}", "saldoscp",
               foto_linea(ra), "(no existe)", ra, None)

    for rb in disponibles:                             # 5) sobra en el reproceso
        anotar(clave, nombre_deudor(rb["datos"]),
               "Registro que NO esta en la DIAN", "saldoscp",
               "(no existe)", foto_linea(rb), None, rb)

# ---------- resumen (solo en pantalla) ----------
conteo = Counter(f["tipo_diferencia"] for f in diferencias)

def suma_sal(registros):
    total, malos = 0.0, 0
    for r in registros:
        try:
            total += float(r["datos"]["sal"])
        except ValueError:
            malos += 1
    return total, malos

sal_dian, malos_dian = suma_sal(reg_dian)
sal_rep,  malos_rep  = suma_sal(reg_rep)

resumen = [
    {"concepto": "Archivos leidos en DIAN", "valor": len(arch_dian)},
    {"concepto": f"Archivos leidos en {NOMBRE_B}", "valor": len(arch_rep)},
    {"concepto": "Registros saldoscp en DIAN", "valor": len(reg_dian)},
    {"concepto": f"Registros saldoscp en {NOMBRE_B}", "valor": len(reg_rep)},
    {"concepto": "Diferencia en cantidad de registros", "valor": len(reg_dian) - len(reg_rep)},
    {"concepto": "Suma sal DIAN", "valor": sal_dian},
    {"concepto": f"Suma sal {NOMBRE_B}", "valor": sal_rep},
    {"concepto": "Diferencia en suma sal", "valor": sal_dian - sal_rep},
    {"concepto": "Valores sal no numericos (DIAN / reproceso)",
     "valor": f"{malos_dian} / {malos_rep}"},
    {"concepto": "Lineas identicas", "valor": iguales},
    {"concepto": "Llaves con alguna diferencia",
     "valor": len({f["llave (cpt | nid)"] for f in diferencias})},
    {"concepto": "Total diferencias en el reporte", "valor": len(diferencias)}]
resumen += [{"concepto": f"  - {t}", "valor": n} for t, n in sorted(conteo.items())]
resumen += [
    {"concepto": "Cruces con el nid sin el dv pegado",
     "valor": sum(1 for r in reg_rep if r["como"] == "nid sin el dv pegado")},
    {"concepto": "Avisos de lectura o codificacion", "valor": len(avisos)}]

print()
print(pd.DataFrame(resumen).to_string(index=False))

if avisos:
    print("\n*** AVISOS: el reporte puede estar incompleto ***")
    for a in avisos:
        print(f'   {a["origen"]:12} | {a["archivo"][:45]:45} | {a["aviso"]}')

# ---------- excel: una sola hoja ----------
df = pd.DataFrame(diferencias) if diferencias else pd.DataFrame(
    [{"resultado": "sin diferencias: las dos carpetas tienen el mismo contenido"}])

with pd.ExcelWriter(SALIDA, engine="openpyxl") as libro:
    df.to_excel(libro, sheet_name="DIFERENCIAS", index=False)
    hoja = libro.book["DIFERENCIAS"]
    hoja.freeze_panes = "A2"
    hoja.auto_filter.ref = hoja.dimensions
    for columna in hoja.columns:
        largo = max(len(str(celda.value or "")) for celda in columna)
        hoja.column_dimensions[columna[0].column_letter].width = min(max(largo + 2, 12), 60)

files.download(SALIDA)