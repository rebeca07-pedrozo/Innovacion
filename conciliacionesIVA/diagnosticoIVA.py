q = f"'{FOLDER_ID}' in parents and trashed = false"
res = drive.files().list(q=q, fields="files(id, name, mimeType)", pageSize=1000).execute()
for f in res.get("files", []):
    print(repr(f["name"]), "|", f["mimeType"])