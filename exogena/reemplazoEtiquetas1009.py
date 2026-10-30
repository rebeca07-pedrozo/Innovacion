# ===========================================================
# CORRECTOR XML FORMATO 1009 - completa datos demograficos desde la DIAN
# Solo toca el campo cuando el reproceso dice NO REGISTRA o viene vacio.
# Cruza por cpt + nid. El resto de los datos no se tocan.
# ===========================================================
ID_CARPETA_DIAN      = "PEGA_AQUI_EL_ID"   # XML descargados de la DIAN (fuente correcta)
ID_CARPETA_REPROCESO = "PEGA_AQUI_EL_ID"   # XML del reproceso a completar
ID_CARPETA_SALIDA    = ""                  # carpeta NUEVA y VACIA de Drive; vacia = baja un zip
SALIDA_LOG = "correccion_1009.xlsx"

CAMPOS_A_COMPLETAR = ["apl1", "apl2", "nom1", "nom2", "dir", "dpto", "mun", "pais"]
SIN_DATO = {"", "NO REGISTRA", "NOREGISTRA"}      # lo que se considera "falta el dato"

import io, os, re, shutil, hashlib, unicodedata
import xml.etree.ElementTree as ET
import pandas as pd
from google.colab import auth, files
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

auth.authenticate_user()
drive = build("drive", "v3")

# orden oficial de atributos del saldoscp
ORDEN = ["cpt","tdoc","nid","dv","apl1","apl2","nom1","nom2","raz",
         "dir","dpto","mun","pais","sal"]
NUMERICOS = {"cpt","tdoc","nid","dv","dpto","mun","pais","sal"}
MARCAS = ("Ã", "Â", "â€")
CARPETA = "xml_1009_corregidos"

# ---------- lectura robusta ----------
def arreglar(texto):
    """BRICEÃO -> BRICEÑO. Solo entra si el texto trae las marcas del daño."""
    salida = texto or ""
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

def leer_xml(crudo):
    """Detecta la codificacion real, repara la declaracion dañada y devuelve la raiz."""
    try:
        texto = crudo.decode("utf-8")
    except UnicodeDecodeError:
        texto = crudo.decode("latin-1")
    texto = texto.lstrip("\ufeff")
    texto = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', texto)
    roto = re.match(r'^<(.)xml\s+version=.*?\1>', texto, flags=re.DOTALL)
    if roto and roto.group(1) != "?":                  # <Ñxml ... Ñ>
        texto = texto[roto.end():]
    for prefijo, uri in re.findall(r'xmlns:?([\w\-]*)="([^"]+)"', texto):
        ET.register_namespace(prefijo, uri)            # conserva los prefijos al reescribir
    texto = re.sub(r'^\s*<\?xml.*?\?>', '', texto, count=1, flags=re.DOTALL)
    return ET.fromstring(texto)

def clave_de(nodo):
    def norma(campo):
        v = arreglar((nodo.get(campo) or "").strip())
        return str(int(v)) if campo in NUMERICOS and v.isdigit() else v.upper()
    return (norma("cpt"), norma("nid"))

def falta(valor):
    return " ".join(arreglar((valor or "").strip()).upper().split()) in SIN_DATO

# ---------- utilidades de Drive ----------
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

# ---------- 1) indice de la DIAN ----------
por_llave, por_nid = {}, {}
total_dian, vistos = 0, set()
for archivo in listar_xml(ID_CARPETA_DIAN):
    crudo = bajar(archivo["id"])
    huella = hashlib.md5(crudo).hexdigest()
    if huella in vistos:
        print(f'DIAN | {archivo["name"][:45]:45} | DUPLICADO, se omite')
        continue
    vistos.add(huella)
    try:
        raiz = leer_xml(crudo)
    except Exception as e:
        print(f'DIAN | {archivo["name"][:45]:45} | NO SE PUDO LEER: {e}')
        continue
    for nodo in raiz.iter():
        if "}" in nodo.tag:
            nodo.tag = nodo.tag.split("}", 1)[1]
    cuantos = 0
    for reg in raiz.iter("saldoscp"):
        datos = {c: arreglar(reg.get(c) or "") for c in ORDEN}
        clave = clave_de(reg)
        por_llave.setdefault(clave, []).append(datos)
        por_nid.setdefault(clave[1], []).append(datos)
        total_dian += 1
        cuantos += 1
    print(f'DIAN | {archivo["name"][:45]:45} | registros: {cuantos}')

# ---------- 2) completar el reproceso ----------
os.makedirs(CARPETA, exist_ok=True)
log, sin_cruce, sin_fuente = [], [], []
completados = 0

for archivo in listar_xml(ID_CARPETA_REPROCESO):
    crudo = bajar(archivo["id"])
    try:
        raiz = leer_xml(crudo)
    except Exception as e:
        print(f'REPR | {archivo["name"][:45]:45} | NO SE PUDO LEER: {e}')
        continue

    cambios, cuantos = 0, 0
    for fila, nodo in enumerate(
            [n for n in raiz.iter() if n.tag.split("}")[-1] == "saldoscp"], start=1):
        cuantos = fila
        datos = {k: arreglar(v) for k, v in nodo.attrib.items()}
        cpt, nid = clave_de(nodo)

        # busca la pareja en la DIAN
        fuente, como = None, ""
        if por_llave.get((cpt, nid)):
            fuente, como = por_llave[(cpt, nid)][0], "llave exacta"
        elif len(nid) > 1 and por_llave.get((cpt, nid[:-1])):
            fuente, como = por_llave[(cpt, nid[:-1])][0], "nid sin el dv pegado"
        elif len(por_nid.get(nid, [])) == 1:
            fuente, como = por_nid[nid][0], "solo por nid (unico)"

        if fuente is None:
            sin_cruce.append({"archivo": archivo["name"], "fila": fila,
                              "cpt": cpt, "nid": nid,
                              "deudor": datos.get("raz") or
                                        f'{datos.get("nom1","")} {datos.get("apl1","")}'.strip(),
                              "motivo": "esa llave no esta en los XML de la DIAN"})
        else:
            for campo in CAMPOS_A_COMPLETAR:
                if not falta(datos.get(campo)):
                    continue                           # ya tiene dato: no se toca
                nuevo = (fuente.get(campo) or "").strip()
                if falta(nuevo):
                    sin_fuente.append({"archivo": archivo["name"], "fila": fila,
                                       "cpt": cpt, "nid": nid, "campo": campo,
                                       "motivo": "la DIAN tampoco tiene ese dato"})
                    continue
                viejo = datos.get(campo, "")
                datos[campo] = nuevo
                cambios += 1
                log.append({"archivo": archivo["name"], "fila": fila,
                            "cpt": cpt, "nid": nid,
                            "deudor": fuente.get("raz") or
                                      f'{fuente.get("nom1","")} {fuente.get("apl1","")}'.strip(),
                            "como_cruzo": como, "campo": campo,
                            "antes": viejo.strip() or "(vacio)", "despues": nuevo})

        # reescribe los atributos en el orden oficial
        nodo.attrib.clear()
        for campo in ORDEN:
            if campo in datos:
                nodo.set(campo, datos.pop(campo))
        for campo, valor in datos.items():             # cualquier atributo extra va al final
            nodo.set(campo, valor)

        if cambios:
            completados = completados          # contador real abajo

    completados += cambios
    cuerpo = ET.tostring(raiz, encoding="unicode")
    final = '<?xml version="1.0" encoding="ISO-8859-1"?>\n' + cuerpo
    with open(os.path.join(CARPETA, archivo["name"]), "wb") as f:
        f.write(final.encode("ISO-8859-1", errors="xmlcharrefreplace"))
    print(f'REPR | {archivo["name"][:45]:45} | registros: {cuantos} | campos completados: {cambios}')

# ---------- 3) log en excel ----------
resumen = [
    {"concepto": "Registros en los XML de la DIAN", "valor": total_dian},
    {"concepto": "Archivos del reproceso corregidos", "valor": len(os.listdir(CARPETA))},
    {"concepto": "Campos completados", "valor": len(log)},
    {"concepto": "Registros sin cruce en la DIAN", "valor": len(sin_cruce)},
    {"concepto": "Campos que la DIAN tampoco tiene", "valor": len(sin_fuente)}]

def guardar(libro, datos, hoja):
    df = pd.DataFrame(datos) if datos else pd.DataFrame([{"resultado": "sin registros"}])
    df.to_excel(libro, sheet_name=hoja, index=False)

with pd.ExcelWriter(SALIDA_LOG, engine="openpyxl") as libro:
    guardar(libro, resumen, "RESUMEN")
    guardar(libro, log, "COMPLETADOS")
    guardar(libro, sin_cruce, "SIN_CRUCE")
    guardar(libro, sin_fuente, "SIN_DATO_EN_DIAN")
    for hoja in libro.book.worksheets:
        hoja.freeze_panes = "A2"
        hoja.auto_filter.ref = hoja.dimensions
        for columna in hoja.columns:
            largo = max(len(str(celda.value or "")) for celda in columna)
            hoja.column_dimensions[columna[0].column_letter].width = min(max(largo + 2, 12), 60)

print()
print(pd.DataFrame(resumen).to_string(index=False))

# ---------- 4) entrega ----------
archivos_salida = sorted(os.listdir(CARPETA))
if ID_CARPETA_SALIDA and not ID_CARPETA_SALIDA.startswith("PEGA"):
    for nombre in archivos_salida:
        drive.files().create(
            body={"name": nombre, "parents": [ID_CARPETA_SALIDA]},
            media_body=MediaFileUpload(os.path.join(CARPETA, nombre),
                                       mimetype="application/xml"),
            supportsAllDrives=True).execute()
    print(f'\n{len(archivos_salida)} XML corregidos subidos a Drive')
    files.download(SALIDA_LOG)
else:
    shutil.make_archive(CARPETA, "zip", CARPETA)
    print(f'\n{len(archivos_salida)} XML corregidos en {CARPETA}.zip')
    files.download(SALIDA_LOG)
    files.download(f"{CARPETA}.zip")