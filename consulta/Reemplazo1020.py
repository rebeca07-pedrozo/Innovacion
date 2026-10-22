# ===========================================================
# CORRECTOR XML FORMATO 1020 - completa datos desde la DIAN
# Copia nid, dv, raz, dir, dpto, mun, pais del XML de la DIAN
# sobre los XML de la carpeta de la jefa
# ===========================================================
ID_CARPETA_DIAN   = "PEGA_AQUI_EL_ID"   # XML descargados de la DIAN (fuente correcta)
ID_CARPETA_JEFA   = "PEGA_AQUI_EL_ID"   # XML a corregir
ID_CARPETA_SALIDA = "PEGA_AQUI_EL_ID"   # carpeta NUEVA y VACIA de Drive; dejala en "" para bajar un zip
CAMPOS_A_CORREGIR = ["nid", "dv", "raz", "dir", "dpto", "mun", "pais"]
SALIDA_LOG = "correccion_1020.xlsx"

import io, os, re, shutil
import xml.etree.ElementTree as ET
import pandas as pd
from google.colab import auth, files
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

auth.authenticate_user()
drive = build("drive", "v3")

# orden oficial de atributos del invcdt
ORDEN = ["tdoc","nid","dv","apl1","apl2","nom1","nom2","raz","dir","dpto","mun",
         "pais","ntit","ttitu","tmov","salini","inv","vintca","vintpa","retfup","salfin"]
CARPETA_SALIDA = "xml_corregidos"

# ---------- utilidades ----------
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

def buscar(raiz, nombre):                      # ignora el prefijo de namespace
    return [n for n in raiz.iter() if n.tag.split("}")[-1] == nombre]

def numero(valor):                             # quita ceros a la izquierda
    v = (valor or "").strip()
    return str(int(v)) if v.isdigit() else v.upper()

# ---------- 1) leer la DIAN y armar los indices ----------
por_llave, por_ntit = {}, {}
total_dian = 0
for archivo in listar_xml(ID_CARPETA_DIAN):
    raiz = ET.fromstring(bajar(archivo["id"]))
    for nodo in buscar(raiz, "invcdt"):
        datos = dict(nodo.attrib)
        nid, ntit = numero(datos.get("nid")), numero(datos.get("ntit"))
        por_llave.setdefault((nid, ntit), []).append(datos)
        por_ntit.setdefault(ntit, []).append(datos)
        total_dian += 1

# ---------- 2) recorrer los XML de la jefa y corregir ----------
os.makedirs(CARPETA_SALIDA, exist_ok=True)
log, sin_cruce = [], []
corregidos = 0

for archivo in listar_xml(ID_CARPETA_JEFA):
    crudo = bajar(archivo["id"])
    texto = crudo.decode("ISO-8859-1", errors="replace")
    for prefijo, uri in re.findall(r'xmlns:?([\w\-]*)="([^"]+)"', texto):
        ET.register_namespace(prefijo, uri)     # conserva los prefijos al reescribir
    codificacion = (re.search(r'encoding="([^"]+)"', texto[:200]) or [None, "ISO-8859-1"])[1]
    raiz = ET.fromstring(crudo)

    for fila, nodo in enumerate(buscar(raiz, "invcdt"), start=1):
        datos = dict(nodo.attrib)
        nid, ntit = numero(datos.get("nid")), numero(datos.get("ntit"))

        # busca la pareja en la DIAN, en tres intentos
        fuente, como = None, ""
        if por_llave.get((nid, ntit)):
            fuente, como = por_llave[(nid, ntit)].pop(0), "llave exacta"
        elif len(nid) > 1 and por_llave.get((nid[:-1], ntit)):
            fuente, como = por_llave[(nid[:-1], ntit)].pop(0), "nid sin el dv pegado"
        elif len(por_ntit.get(ntit, [])) == 1:
            fuente, como = por_ntit[ntit][0], "solo por ntit (unico)"

        if fuente is None:
            sin_cruce.append({"archivo": archivo["name"], "fila": fila,
                              "nid_jefa": datos.get("nid", ""), "ntit": datos.get("ntit", ""),
                              "motivo": "no se encontro esa llave en la carpeta de la DIAN"})
            continue

        cambios = []
        for campo in CAMPOS_A_CORREGIR:
            nuevo = (fuente.get(campo) or "").strip()
            viejo = (datos.get(campo) or "").strip()
            if nuevo and nuevo != viejo:
                datos[campo] = nuevo
                cambios.append(f'{campo}: {viejo or "(no existia)"} -> {nuevo}')

        # reescribe los atributos en el orden oficial
        nodo.attrib.clear()
        for campo in ORDEN:
            if campo in datos:
                nodo.set(campo, datos.pop(campo))
        for campo, valor in datos.items():      # cualquier atributo extra va al final
            nodo.set(campo, valor)

        if cambios:
            corregidos += 1
        log.append({"archivo": archivo["name"], "fila": fila,
                    "nid_original": fuente.get("nid", ""), "ntit": ntit,
                    "titular": fuente.get("raz") or f'{fuente.get("nom1","")} {fuente.get("apl1","")}'.strip(),
                    "como_cruzo": como,
                    "campos_corregidos": len(cambios),
                    "detalle": " | ".join(cambios) if cambios else "ya estaba completo"})

    salida = os.path.join(CARPETA_SALIDA, archivo["name"])
    with open(salida, "wb") as f:
        f.write(ET.tostring(raiz, encoding=codificacion, xml_declaration=True))

# ---------- 3) resumen ----------
resumen = [
    {"concepto": "Registros en los XML de la DIAN", "valor": total_dian},
    {"concepto": "Registros procesados de la jefa", "valor": len(log) + len(sin_cruce)},
    {"concepto": "Registros cruzados", "valor": len(log)},
    {"concepto": "Registros corregidos", "valor": corregidos},
    {"concepto": "Registros que ya estaban completos", "valor": len(log) - corregidos},
    {"concepto": "Registros SIN cruce (quedaron igual)", "valor": len(sin_cruce)}]
for tipo in ["llave exacta", "nid sin el dv pegado", "solo por ntit (unico)"]:
    resumen.append({"concepto": f"  - cruzados por {tipo}",
                    "valor": sum(1 for f in log if f["como_cruzo"] == tipo)})

def guardar(libro, datos, hoja):
    df = pd.DataFrame(datos) if datos else pd.DataFrame([{"resultado": "sin registros"}])
    df.to_excel(libro, sheet_name=hoja, index=False)

with pd.ExcelWriter(SALIDA_LOG, engine="openpyxl") as libro:
    guardar(libro, resumen, "RESUMEN")
    guardar(libro, log, "CORRECCIONES")
    guardar(libro, sin_cruce, "SIN_CRUCE")
    for hoja in libro.book.worksheets:
        hoja.freeze_panes = "A2"
        hoja.auto_filter.ref = hoja.dimensions
        for columna in hoja.columns:
            largo = max(len(str(celda.value or "")) for celda in columna)
            hoja.column_dimensions[columna[0].column_letter].width = min(max(largo + 2, 12), 70)

# ---------- 4) entrega de los XML corregidos ----------
archivos_salida = sorted(os.listdir(CARPETA_SALIDA))
print(pd.DataFrame(resumen).to_string(index=False))

if ID_CARPETA_SALIDA and not ID_CARPETA_SALIDA.startswith("PEGA"):
    for nombre in archivos_salida:
        drive.files().create(
            body={"name": nombre, "parents": [ID_CARPETA_SALIDA]},
            media_body=MediaFileUpload(os.path.join(CARPETA_SALIDA, nombre),
                                       mimetype="application/xml"),
            supportsAllDrives=True).execute()
    print(f"\n{len(archivos_salida)} XML corregidos subidos a la carpeta de Drive")
    files.download(SALIDA_LOG)
else:
    shutil.make_archive("xml_corregidos", "zip", CARPETA_SALIDA)
    print(f"\n{len(archivos_salida)} XML corregidos en xml_corregidos.zip")
    files.download(SALIDA_LOG)
    files.download("xml_corregidos.zip")