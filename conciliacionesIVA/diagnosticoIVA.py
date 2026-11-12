q = f"'{FOLDER_ID}' in parents and trashed = false"
res = drive.files().list(q=q, fields="files(id, name, mimeType)", pageSize=1000).execute()
for f in res.get("files", []):
    print(repr(f["name"]), "|", f["mimeType"])

def listar_archivos(folder_id, extensiones=(".txt", ".csv")):
    """Archivos .txt/.csv o de tipo texto plano/csv"""
    q = f"'{folder_id}' in parents and trashed = false"
    res = drive.files().list(q=q, fields="files(id, name, mimeType)", pageSize=1000).execute()
    tipos_ok = ("text/plain", "text/csv", "text/comma-separated-values")
    archivos = []
    for f in res.get("files", []):
        nombre_ok = f["name"].lower().endswith(extensiones)
        tipo_ok = f.get("mimeType") in tipos_ok
        if nombre_ok or tipo_ok:
            archivos.append(f)
    return archivos

import unicodedata

def leer_txt(ruta):
    """Lee el .txt/.csv a DataFrame probando utf-8 y latin-1."""
    for enc in ("utf-8-sig", "latin-1"):
        try:
            return pd.read_csv(ruta, sep=SEPARADOR, dtype=str, encoding=enc, engine="python")
        except Exception:
            continue
    raise ValueError(f"No pude leer {ruta}")

def _col(df, nombre):
    """Encuentra la columna ignorando mayúsculas, acentos y espacios."""
    def norm(s):
        s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
        return s.strip().lower()
    objetivo = norm(nombre)
    for c in df.columns:
        if norm(c) == objetivo:
            return c
    raise KeyError(f"No encontré '{nombre}'. Columnas: {list(df.columns)}")

def _a_numero(serie):
    """Convierte texto a número: quita puntos de miles y usa coma como decimal."""
    s = serie.astype(str).str.strip().str.replace(r"[^\d,.\-]", "", regex=True)
    s = s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    return pd.to_numeric(s, errors="coerce").fillna(0.0)