#1 - Imports 
!pip install pdfplumber -q
from google.colab import drive
drive.mount('/content/drive')
import glob
import glob, os
import pdfplumber
import pandas as pd
import re, unicodedata, hashlib
from google.colab import auth
from googleapiclient.discovery import build
import math
from collections import Counter, defaultdict
import gspread
from google.auth import default


# 2 - Rutas/constantes
CARPETA = '/content/drive/MyDrive/MotorDeBusqueda/input_pdfs'
SALIDA = '/content/drive/MyDrive/MotorDeBusqueda/output/resumen_extraccion.xlsx'
for ruta in glob.glob('/content/drive/MyDrive/**/inputs_pdfs', recursive=True):
    print(ruta)

#3 - Listar archivos PDF
archivos = sorted(glob.glob(os.path.join(CARPETA, '*.pdf')))
print(f'{len(archivos)} PDFs encontrados:')
for a in archivos:
    print(' -', os.path.basename(a))

#4 - Extraer texto de los PDFs
textos = []
for ruta in archivos:
    nombre_archivo = os.path.basename(ruta)
    try:
        with pdfplumber.open(ruta) as pdf:
            for i, pagina in enumerate(pdf.pages, start=1):
                textos.append({
                    'nombre_archivo': nombre_archivo,
                    'ruta_completa': ruta,
                    'pagina': i,
                    'texto': pagina.extract_text() or ''
                })
    except Exception as e:
        print(f'Error en {nombre_archivo}: {e}')

df_texto = pd.DataFrame(textos)
print(f'{df_texto.nombre_archivo.nunique()} archivos | {len(df_texto)} páginas')



#5 - Páginas sin texto
vacias = df_texto[df_texto.texto.str.strip() == '']
print(f'Páginas sin texto: {len(vacias)} de {len(df_texto)}')
print(df_texto.iloc[0].texto[:2000])


#6 - Resumen por archivo
df_resumen = df_texto.groupby('nombre_archivo').agg(
    total_paginas=('pagina', 'max'),
    caracteres=('texto', lambda s: s.str.len().sum()),
    paginas_vacias=('texto', lambda s: (s.str.strip() == '').sum())
).reset_index()
df_resumen['posible_escaneado'] = df_resumen.caracteres < 50
with pd.ExcelWriter(SALIDA) as w:
    df_resumen.to_excel(w, sheet_name='resumen', index=False)
    df_texto.to_excel(w, sheet_name='texto_detallado', index=False)
print('Guardado en', SALIDA)


#7 - funciones de normalización y metadatos
MESES = 'enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre'
RE_FECHA = rf'(\d{{1,2}}\s+de\s+(?:{MESES})\s+de\s+\d{{4}}|(?:{MESES})\s+de\s+\d{{4}})'
RE_RAD_ARCHIVO = r'(\d{4}[A-Z]{2}\d{4,}[A-Z]?\d*)'
RE_RAD_TEXTO = r'Radicado[:\s]+(?:Solicitud\s+)?([A-Z0-9][A-Z0-9\-]{5,24})'

ENTIDADES = [
    ('impuestos de bogota',                       'DIB - Dirección Distrital de Impuestos de Bogotá'),
    ('fortalecimiento institucional territorial', 'MinHacienda - Apoyo Fiscal'),
    ('ministerio de hacienda',                    'MinHacienda'),
    ('secretaria distrital de hacienda',          'SDH - Secretaría Distrital de Hacienda'),
]

def sin_tildes(s):
    return ''.join(c for c in unicodedata.normalize('NFD', s)
                   if unicodedata.category(c) != 'Mn')

def normalizar(s):
    s = sin_tildes(s.lower())
    s = re.sub(r'[^a-z0-9ñ\s]', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()

def campo(cab, etiqueta):
    r = re.search(rf'(?:^|\n)\s*{etiqueta}\s*:?\s*(.+)', cab, re.I)
    return r.group(1).strip(' .:-') if r else None

def metadatos(cab, archivo):
    plano = re.sub(r'\s+', ' ', sin_tildes(cab).lower())

    rad = re.search(RE_RAD_ARCHIVO, archivo) or re.search(RE_RAD_TEXTO, cab)
    f = re.search(RE_FECHA, cab[:500], re.I)
    if f:
        fecha = f.group(1).lower()
    else:
        f2 = re.search(rf'(\d{{1,2}})[-_\s]de[-_\s]({MESES})[-_\s]de[-_\s](\d{{4}})', archivo, re.I)
        fecha = f'{f2.group(1)} de {f2.group(2).lower()} de {f2.group(3)}' if f2 else None

    a = re.search(r'(19|20)\d{2}', fecha or archivo)
    tema = campo(cab, 'Tema') or campo(cab, r'Problema jur[ií]dico') or campo(cab, 'ASUNTO')
    subtema = campo(cab, 'Subtema') or campo(cab, r'Descriptores? especiales?')

    return {
        'radicado': rad.group(1) if rad else None,
        'fecha': fecha,
        'anio': int(a.group(0)) if a else None,
        'entidad': next((v for k, v in ENTIDADES if k in plano), None),
        'tema': tema[:120] if tema else None,
        'subtema': subtema[:160] if subtema else None,
    }


#8 - reflujo de líneas y fragmentación
def reflow(texto):
    parrafos, buf = [], ''
    for l in (x.strip() for x in texto.split('\n')):
        if not l:
            continue
        if buf and buf.endswith(('.', ':', '?', '!')) and len(buf) > 200:
            parrafos.append(buf); buf = l
        else:
            buf = f'{buf} {l}'.strip()
    if buf:
        parrafos.append(buf)
    return parrafos

def trocear(parrafos, objetivo=800, solape=150, minimo=250):
    chunks, buf = [], ''
    for p in parrafos:
        if not buf:
            buf = p
        elif len(buf) + len(p) + 1 <= objetivo:
            buf = f'{buf} {p}'
        else:
            chunks.append(buf)
            cola = buf[-solape:]
            corte = cola.find(' ')
            buf = ((cola[corte+1:] if corte > 0 else '') + ' ' + p).strip()
        while len(buf) > objetivo * 1.5:
            corte = buf.rfind(' ', 0, objetivo)
            chunks.append(buf[:corte])
            buf = buf[max(0, corte - solape):].strip()
    if buf:
        chunks.append(buf)

    fusionados = []
    for c in chunks:
        if fusionados and len(c) < minimo:
            fusionados[-1] = f'{fusionados[-1]} {c}'
        else:
            fusionados.append(c)
    return fusionados


#9 - obtener el file_id de Drive
auth.authenticate_user()
drive = build('drive', 'v3')

def obtener_file_id(nombre):
    n = nombre.replace("'", "\\'")
    r = drive.files().list(
        q=f"name = '{n}' and trashed = false",
        fields='files(id, name)', pageSize=5,
        supportsAllDrives=True, includeItemsFromAllDrives=True
    ).execute()
    archivos = r.get('files', [])
    return archivos[0]['id'] if archivos else None

ids = {a: obtener_file_id(a) for a in df_texto.nombre_archivo.unique()}
for a, i in ids.items():
    print(('OK  ' if i else 'FALTA'), a, '->', i)


#10 - Construir la tabla y guardarla
filas = []
for archivo, g in df_texto.groupby('nombre_archivo', sort=False):
    g = g.sort_values('pagina')
    meta = metadatos(str(g.texto.iloc[0])[:2200], archivo)
    sigla = hashlib.md5(archivo.encode()).hexdigest()[:6]

    for _, pag in g.iterrows():
        for k, ch in enumerate(trocear(reflow(str(pag.texto))), start=1):
            filas.append({
                'chunk_id': f'{sigla}-p{int(pag.pagina)}-{k}',
                'nombre_archivo': archivo,
                'file_id': ids.get(archivo),
                'pagina': int(pag.pagina),
                **meta,
                'texto': ch,
                'texto_norm': normalizar(ch),
            })

df_chunks = pd.DataFrame(filas)
df_chunks.to_excel('/content/drive/MyDrive/fase1_chunks.xlsx', index=False)

print(f'{len(df_texto)} páginas -> {len(df_chunks)} fragmentos')
display(df_chunks.groupby('nombre_archivo')[['radicado','fecha','entidad','tema','subtema']].first())

#11 - Construccion del indice 


STOPWORDS = set("""a al algo algun alguna algunas alguno algunos ante antes aquel aquella aquellas
aquello aquellos aqui asi aun aunque cada como con contra cual cuales cuando cuanto de del desde donde
dos e el ella ellas ello ellos en entre era eran es esa esas ese eso esos esta estan estas este esto
estos fue fueron ha haber habia han hasta hay la las le les lo los mas me mi mientras mismo mucho muy
nada ni no nos nuestra nuestro o os otra otras otro otros para pero poco por porque que quien quienes
se segun ser si sin sobre son su sus tal tambien tanto te tiene tienen toda todas todo todos tras un
una unas uno unos y ya""".split())

def raiz(p):
    if len(p) > 5:
        for suf in ('ciones', 'idades'):
            if p.endswith(suf):
                return p[:-len(suf)] + suf[0]
    if len(p) > 4 and p.endswith('es'):
        return p[:-2]
    if len(p) > 3 and p.endswith('s'):
        return p[:-1]
    return p

def tokenizar(texto):
    return [raiz(p) for p in normalizar(texto).split()
            if len(p) > 2 and p not in STOPWORDS]

PESO_TEMA, K1, B = 3, 1.5, 0.75

postings, longitudes = defaultdict(dict), {}

for _, r in df_chunks.iterrows():
    toks = tokenizar(r.texto)
    encabezado = ' '.join(str(x) for x in [r.tema, r.subtema, r.entidad]
                          if x and str(x) != 'nan')
    tf = Counter(toks)
    for t, n in Counter(tokenizar(encabezado)).items():
        tf[t] += n * PESO_TEMA
    longitudes[r.chunk_id] = len(toks)
    for t, n in tf.items():
        postings[t][r.chunk_id] = n

N = len(df_chunks)
avgdl = sum(longitudes.values()) / N
idf = {t: math.log(1 + (N - len(d) + 0.5) / (len(d) + 0.5)) for t, d in postings.items()}

print(f'{N} fragmentos | {len(postings)} términos | avgdl {avgdl:.1f}')

#12 - Poblacion del ranking 
def buscar(q, top=5):
    puntajes = defaultdict(float)
    for t in tokenizar(q):
        if t not in postings:
            continue
        for cid, f in postings[t].items():
            dl = longitudes[cid]
            puntajes[cid] += idf[t] * (f * (K1 + 1)) / (f + K1 * (1 - B + B * dl / avgdl))
    return sorted(puntajes.items(), key=lambda x: -x[1])[:top]

for q in ['diferencia en cambio', 'estampillas distritales', 'ICA sector financiero']:
    print(f'\n>>> {q}')
    for cid, s in buscar(q):
        r = df_chunks[df_chunks.chunk_id == cid].iloc[0]
        print(f'  {s:6.2f}  p{r.pagina}  {r.nombre_archivo[:45]}')

#13 - Volcado a Google Sheets 


creds, _ = default()
gc = gspread.authorize(creds)

NOMBRE = 'indice_doctrina_tributaria'
FOLDER_ID = 'pega_aqui_el_id_de_MotorDeBusqueda'

try:
    sh = gc.open(NOMBRE)
    print('Reutilizando el existente')
    for ws_viejo in sh.worksheets()[1:]:
        sh.del_worksheet(ws_viejo)
    sh.sheet1.clear()
except gspread.SpreadsheetNotFound:
    sh = gc.create(NOMBRE, folder_id=FOLDER_ID)
    print('Creado nuevo')

print('URL:', sh.url)
print('SHEET_ID:', sh.id) 
padres = drive.files().get(
    fileId=sh.id, fields='parents', supportsAllDrives=True
).execute().get('parents', [])

if FOLDER_ID not in padres:
    drive.files().update(
        fileId=sh.id,
        addParents=FOLDER_ID,
        removeParents=','.join(padres),
        fields='id, parents',
        supportsAllDrives=True
    ).execute()
    print('Movido a la carpeta')
else:
    print('Ya estaba en la carpeta')

cols = ['chunk_id','nombre_archivo','file_id','pagina','radicado','fecha',
        'anio','entidad','tema','subtema','texto']
frag = df_chunks[cols].fillna('').astype(str)
ws = sh.sheet1
ws.update_title('Fragmentos')
ws.update([cols] + frag.values.tolist())

idx = [['termino','df','idf','postings']] + [
    [t, len(d), round(idf[t], 4), ','.join(f'{c}:{n}' for c, n in d.items())]
    for t, d in sorted(postings.items())
]
sh.add_worksheet('Indice', rows=len(idx)+10, cols=4).update(idx)

stats = [['clave','valor'],
         ['N', N], ['avgdl', round(avgdl, 4)],
         ['k1', K1], ['b', B], ['peso_tema', PESO_TEMA]] + \
        [['len:' + c, l] for c, l in longitudes.items()]
sh.add_worksheet('Stats', rows=len(stats)+10, cols=2).update(stats)

print(f'Listo: {len(frag)} fragmentos, {len(idx)-1} términos')


#has not this changed