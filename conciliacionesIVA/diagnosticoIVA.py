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