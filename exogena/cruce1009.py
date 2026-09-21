# ============================================================
# COMPARADOR FORMATO 1009
# DESCARGAS DIAN vs ORIGINALES
# Los archivos NO se llaman igual: se busca cada línea por su llave
# ============================================================

from google.colab import auth, files
from googleapiclient.discovery import build
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from collections import Counter
import xml.etree.ElementTree as ET
import pandas as pd

auth.authenticate_user()
drive = build('drive', 'v3')

# ---------------- CONFIGURACIÓN ----------------

FOLDER_DIAN = "1FOr0h_u3btz-upYimOwFD_9FZRNTkFQ-"       # Carpeta REPROCESO 1
FOLDER_ORIGINAL = "1CAEXJO7c4JeUM7CyYYtcYCOYrCZXTYrp"   # Carpeta REPROCESO 2

NOMBRE_DIAN = "REPROCESO 1"
NOMBRE_ORIGINAL = "REPROCESO 2"
CAMPOS_LLAVE = ['cpt', 'tdoc', 'nid']
ARCHIVO_EXCEL = "COMPARACION_1009.xlsx"

# ---------------- TEXTOS ----------------

SIN_DATO = {'', 'NO REGISTRA'}
IGUAL = "Igual"
CON_DIFERENCIAS = "Línea con diferencias"
SOLO_ESPACIOS = "Solo cambian espacios o mayúsculas"
FALTA_EN_DIAN = f"Línea que no está en {NOMBRE_DIAN}"
FALTA_EN_ORIGINAL = f"Línea que no está en {NOMBRE_ORIGINAL}"

ARCH_DIAN = f"Archivo en {NOMBRE_DIAN}"
ARCH_ORIGINAL = f"Archivo en {NOMBRE_ORIGINAL}"

# ---------------- FUNCIONES ----------------

def listar_xmls(folder_id):
    resultado = drive.files().list(
        q=f"'{folder_id}' in parents and trashed=false",
        spaces='drive', fields='files(id, name)', pageSize=1000
    ).execute()
    archivos = resultado.get('files', [])
    return sorted([f for f in archivos if f['name'].lower().endswith('.xml')],
                  key=lambda x: x['name'])


def descargar_xml(file_id):
    return drive.files().get_media(fileId=file_id).execute()


def unir_campos(campos):
    return " | ".join(f"{k}={v}" for k, v in campos.items())


def leer_xml(contenido):
    """Separa el encabezado (Cab) y las líneas de un XML"""
    root = ET.fromstring(contenido)
    encabezado = {}
    lineas = {}
    for elem in root:
        if elem.tag.lower() == 'cab':
            for campo in elem:
                encabezado[campo.tag] = campo.text or ''
            continue
        llave = (elem.tag,) + tuple(elem.attrib.get(c, '').strip() for c in CAMPOS_LLAVE)
        linea = {
            'campos': dict(elem.attrib),
            'sub': sorted(f"{h.tag}: {unir_campos(h.attrib)}" for h in elem)
        }
        lineas.setdefault(llave, []).append(linea)
    return encabezado, lineas


def leer_carpeta(folder_id, nombre_carpeta):
    """Junta TODAS las líneas de TODOS los archivos de una carpeta"""
    xmls = listar_xmls(folder_id)
    print(f"{nombre_carpeta}: {len(xmls)} archivos")
    registros = {}
    for n, f in enumerate(xmls, start=1):
        if n % 10 == 1:
            print(f"  Leyendo {n}/{len(xmls)}")
        try:
            _, lineas = leer_xml(descargar_xml(f['id']))
        except Exception as e:
            print(f"  Error leyendo {f['name']}: {e}")
            continue
        for llave, lista in lineas.items():
            for linea in lista:
                linea['archivo'] = f['name']
                registros.setdefault(llave, []).append(linea)
    for llave in registros:
        registros[llave].sort(key=lambda l: unir_campos(l['campos']))
    total = sum(len(v) for v in registros.values())
    print(f"  {total} líneas leídas\n")
    return registros, len(xmls), total


def mostrar(valor):
    if valor is None:
        return '(no existe)'
    return valor if valor.strip() else '(vacío)'


def explicar(v_dian, v_orig):
    """Explica en palabras simples por qué un dato es distinto"""
    if v_dian is None:
        return 'falta_dian', f"Está en {NOMBRE_ORIGINAL} pero no en {NOMBRE_DIAN}"
    if v_orig is None:
        return 'falta_orig', f"Está en {NOMBRE_DIAN} pero no en {NOMBRE_ORIGINAL}"
    a, b = v_dian.strip().upper(), v_orig.strip().upper()
    if a == b:
        return 'espacios', "Mismo dato, solo cambian espacios o mayúsculas"
    if a in SIN_DATO:
        return 'sin_dato_dian', f"En {NOMBRE_DIAN} dice {mostrar(v_dian).strip()}, pero en {NOMBRE_ORIGINAL} sí tiene dato"
    if b in SIN_DATO:
        return 'sin_dato_orig', f"En {NOMBRE_ORIGINAL} dice {mostrar(v_orig).strip()}, pero en {NOMBRE_DIAN} sí tiene dato"
    return 'diferente', "El dato es diferente"


def comparar_linea(linea_dian, linea_orig):
    """Compara etiqueta por etiqueta una misma línea"""
    detalles = []
    c_dian, c_orig = linea_dian['campos'], linea_orig['campos']
    etiquetas = list(c_orig) + [e for e in c_dian if e not in c_orig]
    for etiqueta in etiquetas:
        v_dian, v_orig = c_dian.get(etiqueta), c_orig.get(etiqueta)
        if v_dian != v_orig:
            tipo, texto = explicar(v_dian, v_orig)
            detalles.append((etiqueta, v_dian, v_orig, tipo, texto))

    sub_dian, sub_orig = Counter(linea_dian['sub']), Counter(linea_orig['sub'])
    for s in (sub_orig - sub_dian).elements():
        etiqueta, contenido = s.split(': ', 1)
        detalles.append((etiqueta, None, contenido, 'sub_falta_dian',
                         f"Este {etiqueta} está en {NOMBRE_ORIGINAL} pero no en {NOMBRE_DIAN}"))
    for s in (sub_dian - sub_orig).elements():
        etiqueta, contenido = s.split(': ', 1)
        detalles.append((etiqueta, contenido, None, 'sub_falta_orig',
                         f"Este {etiqueta} está en {NOMBRE_DIAN} pero no en {NOMBRE_ORIGINAL}"))
    return detalles


def resumir(detalles):
    frases = {
        'falta_dian': f"Faltan en {NOMBRE_DIAN}",
        'falta_orig': f"Sobran en {NOMBRE_DIAN}",
        'sin_dato_dian': f"Dicen NO REGISTRA o vacío en {NOMBRE_DIAN} pero tienen dato en {NOMBRE_ORIGINAL}",
        'sin_dato_orig': f"Dicen NO REGISTRA o vacío en {NOMBRE_ORIGINAL} pero tienen dato en {NOMBRE_DIAN}",
        'diferente': "Datos diferentes",
        'espacios': "Solo cambian espacios o mayúsculas",
        'sub_falta_dian': f"Subregistros que faltan en {NOMBRE_DIAN}",
        'sub_falta_orig': f"Subregistros que sobran en {NOMBRE_DIAN}",
    }
    grupos = {}
    for etiqueta, _, _, tipo, _ in detalles:
        grupos.setdefault(tipo, []).append(etiqueta)
    return ". ".join(f"{frases[t]} ({len(e)}): {', '.join(e)}" for t, e in grupos.items())


def evaluar(linea_dian, linea_orig):
    if linea_dian is None:
        return FALTA_EN_DIAN, f"La línea completa está en {NOMBRE_ORIGINAL} pero no en {NOMBRE_DIAN}", []
    if linea_orig is None:
        return FALTA_EN_ORIGINAL, f"La línea completa está en {NOMBRE_DIAN} pero no en {NOMBRE_ORIGINAL}", []
    detalles = comparar_linea(linea_dian, linea_orig)
    if not detalles:
        return IGUAL, '', []
    if all(d[3] == 'espacios' for d in detalles):
        return SOLO_ESPACIOS, resumir(detalles), detalles
    return CON_DIFERENCIAS, resumir(detalles), detalles


def emparejar(lista_dian, lista_orig):
    total = max(len(lista_dian), len(lista_orig))
    return [(lista_dian[i] if i < len(lista_dian) else None,
             lista_orig[i] if i < len(lista_orig) else None) for i in range(total)]


def datos_llave(llave):
    return {'Tipo de línea': llave[0], **dict(zip(CAMPOS_LLAVE, llave[1:]))}


def contar(linea):
    return len(linea['campos']) if linea else 0


def texto_linea(linea):
    if linea is None:
        return '(no existe)'
    texto = unir_campos(linea['campos'])
    if linea['sub']:
        texto += "  ||  " + "  ||  ".join(linea['sub'])
    return texto


def fila_linea(base, situacion, resumen, linea_dian, linea_orig):
    fila = dict(base)
    fila.update({
        'Situación': situacion,
        f'Etiquetas en {NOMBRE_DIAN}': contar(linea_dian),
        f'Etiquetas en {NOMBRE_ORIGINAL}': contar(linea_orig),
        'Qué pasó': resumen,
        f'Línea en {NOMBRE_DIAN}': texto_linea(linea_dian),
        f'Línea en {NOMBRE_ORIGINAL}': texto_linea(linea_orig),
    })
    return fila


def fila_detalle(base, etiqueta, v_dian, v_orig, texto):
    fila = dict(base)
    fila.update({
        'Etiqueta': etiqueta,
        f'Valor en {NOMBRE_DIAN}': mostrar(v_dian),
        f'Valor en {NOMBRE_ORIGINAL}': mostrar(v_orig),
        'Qué pasó': texto,
    })
    return fila


def guardar_excel(hojas):
    with pd.ExcelWriter(ARCHIVO_EXCEL, engine='openpyxl') as writer:
        for nombre_hoja, df in hojas.items():
            if df.empty:
                df = pd.DataFrame({'Resultado': ['Sin diferencias']})
            df.to_excel(writer, sheet_name=nombre_hoja, index=False)
            ws = writer.sheets[nombre_hoja]
            for celda in ws[1]:
                celda.fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
                celda.font = Font(bold=True, color="1F3864")
                celda.alignment = Alignment(wrap_text=True, vertical='center')
            for i, col in enumerate(df.columns, start=1):
                largo = max([len(str(col))] + [len(str(v)) for v in df[col].head(200)])
                ws.column_dimensions[get_column_letter(i)].width = min(max(largo + 2, 10), 60)
            ws.freeze_panes = 'A2'
            ws.auto_filter.ref = ws.dimensions

# ---------------- LEER CARPETAS ----------------

print("Leyendo carpetas...\n")
registros_dian, archivos_dian, total_dian = leer_carpeta(FOLDER_DIAN, NOMBRE_DIAN)
registros_orig, archivos_orig, total_orig = leer_carpeta(FOLDER_ORIGINAL, NOMBRE_ORIGINAL)

# ---------------- COMPARACIÓN ----------------

print("Comparando líneas...")
filas_linea, filas_detalle = [], []
conteo = Counter()

for llave in sorted(set(registros_dian) | set(registros_orig)):
    for linea_dian, linea_orig in emparejar(registros_dian.get(llave, []), registros_orig.get(llave, [])):
        situacion, resumen, detalles = evaluar(linea_dian, linea_orig)
        conteo[situacion] += 1
        if situacion == IGUAL:
            continue
        base = {
            ARCH_DIAN: linea_dian['archivo'] if linea_dian else '(no existe)',
            ARCH_ORIGINAL: linea_orig['archivo'] if linea_orig else '(no existe)',
            **datos_llave(llave)
        }
        filas_linea.append(fila_linea(base, situacion, resumen, linea_dian, linea_orig))
        for etiqueta, v_dian, v_orig, _, texto in detalles:
            filas_detalle.append(fila_detalle(base, etiqueta, v_dian, v_orig, texto))

# ---------------- EXCEL ----------------

df_resumen = pd.DataFrame([
    {'Concepto': f'Archivos en {NOMBRE_DIAN}', 'Cantidad': archivos_dian},
    {'Concepto': f'Archivos en {NOMBRE_ORIGINAL}', 'Cantidad': archivos_orig},
    {'Concepto': f'Líneas en {NOMBRE_DIAN}', 'Cantidad': total_dian},
    {'Concepto': f'Líneas en {NOMBRE_ORIGINAL}', 'Cantidad': total_orig},
    {'Concepto': 'Líneas iguales', 'Cantidad': conteo[IGUAL]},
    {'Concepto': 'Líneas con diferencias', 'Cantidad': conteo[CON_DIFERENCIAS]},
    {'Concepto': 'Líneas donde solo cambian espacios', 'Cantidad': conteo[SOLO_ESPACIOS]},
    {'Concepto': f'Líneas que no están en {NOMBRE_DIAN}', 'Cantidad': conteo[FALTA_EN_DIAN]},
    {'Concepto': f'Líneas que no están en {NOMBRE_ORIGINAL}', 'Cantidad': conteo[FALTA_EN_ORIGINAL]},
])
df_linea = pd.DataFrame(filas_linea)
df_detalle = pd.DataFrame(filas_detalle)

print("\nRESUMEN")
print(df_resumen.to_string(index=False))

print("\nGenerando Excel...")
guardar_excel({'Resumen': df_resumen, 'Por línea': df_linea, 'Detalle por etiqueta': df_detalle})
print(f"Excel listo: {ARCHIVO_EXCEL}")
files.download(ARCHIVO_EXCEL)