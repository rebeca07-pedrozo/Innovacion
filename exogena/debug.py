from google.colab import auth
from googleapiclient.discovery import build

auth.authenticate_user()
drive = build('drive', 'v3')

FOLDER_DESCARGADOS = "1UbGhdnJb2b5RbLRj3gCizK84fYMOGvl5"
FOLDER_JEFA = "1eGZidzTrh19b5M6a30DjhR14X39dq-_a"

print("🔍 VERIFICANDO CARPETA DESCARGAS...\n")

results = drive.files().list(
    q=f"'{FOLDER_DESCARGADOS}' in parents and trashed=false",
    spaces='drive',
    fields='files(id, name, mimeType)',
    pageSize=200
).execute()

files = results.get('files', [])
print(f"Total archivos en DESCARGAS: {len(files)}\n")

if files:
    print("Primeros 10 archivos:")
    for f in files[:10]:
        print(f"  - {f['name']} ({f['mimeType']})")
else:
    print("❌ NO HAY ARCHIVOS O NO TIENES ACCESO")

print("\n" + "="*80)
print("🔍 VERIFICANDO CARPETA JEFA...\n")

results2 = drive.files().list(
    q=f"'{FOLDER_JEFA}' in parents and trashed=false",
    spaces='drive',
    fields='files(id, name, mimeType)',
    pageSize=200
).execute()

files2 = results2.get('files', [])
print(f"Total archivos en JEFA: {len(files2)}\n")

if files2:
    print("Primeros 10 archivos:")
    for f in files2[:10]:
        print(f"  - {f['name']} ({f['mimeType']})")
else:
    print("❌ NO HAY ARCHIVOS O NO TIENES ACCESO")