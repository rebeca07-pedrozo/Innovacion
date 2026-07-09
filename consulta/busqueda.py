# CELDA 1: Autenticación con Google Drive


auth.authenticate_user()
drive_service = build('drive', 'v3')

print("Autenticación exitosa")

# CELDA 2: Listar subcarpetas de una carpeta dada

def listar_subcarpetas(id_carpeta_padre):
    """
    Devuelve una lista de diccionarios con 'id' y 'nombre'
    de las subcarpetas dentro de la carpeta indicada.
    """
    query = f"'{id_carpeta_padre}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    
    resultados = drive_service.files().list(
        q=query,
        fields="files(id, name)",
        pageSize=1000
    ).execute()
    
    subcarpetas = resultados.get('files', [])
    return subcarpetas


def elegir_subcarpeta(id_carpeta_padre):
    """
    Muestra las subcarpetas disponibles y le pide al usuario
    que elija una por número. Devuelve el ID de la subcarpeta elegida.
    """
    subcarpetas = listar_subcarpetas(id_carpeta_padre)
    
    if not subcarpetas:
        print(" No se encontraron subcarpetas dentro de esta carpeta.")
        return None
    
    print(" Subcarpetas encontradas:\n")
    for i, carpeta in enumerate(subcarpetas, start=1):
        print(f"  {i}. {carpeta['name']}")
    
    while True:
        seleccion = input("\n Escribe el número de la subcarpeta que quieres revisar: ")
        try:
            indice = int(seleccion) - 1
            if 0 <= indice < len(subcarpetas):
                elegida = subcarpetas[indice]
                print(f"\n✅ Elegiste: {elegida['name']}")
                return elegida['id'], elegida['name']
            else:
                print(" Número fuera de rango, intenta de nuevo.")
        except ValueError:
            print(" Escribe solo el número, intenta de nuevo.")

# CELDA 3: Pega aquí el ID de la carpeta principal y ejecuta

ID_CARPETA_PRINCIPAL = "PEGA_AQUI_EL_ID_DE_LA_CARPETA"  # <-- reemplaza esto

id_subcarpeta_elegida, nombre_subcarpeta_elegida = elegir_subcarpeta(ID_CARPETA_PRINCIPAL)

# CELDA 4: Buscar PDFs dentro de la subcarpeta y generar Excel

def listar_pdfs(id_carpeta):
    """
    Devuelve una lista de diccionarios con 'id', 'nombre' y 'link'
    de todos los PDFs dentro de la carpeta indicada.
    """
    query = f"'{id_carpeta}' in parents and mimeType = 'application/pdf' and trashed = false"
    
    resultados = drive_service.files().list(
        q=query,
        fields="files(id, name, webViewLink)",
        pageSize=1000
    ).execute()
    
    return resultados.get('files', [])


def crear_excel_con_pdfs(id_subcarpeta, nombre_subcarpeta):
    """
    Crea un Excel con nombre y link de cada PDF encontrado
    y lo sube a la misma subcarpeta en Drive.
    """
    pdfs = listar_pdfs(id_subcarpeta)
    
    if not pdfs:
        print(" No se encontraron archivos PDF en esta subcarpeta.")
        return
    
    print(f" Se encontraron {len(pdfs)} archivos PDF. Generando Excel...")
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Listado PDFs"
    
    ws["A1"] = "Nombre del archivo"
    ws["B1"] = "Link del PDF"
    ws["A1"].font = Font(bold=True)
    ws["B1"].font = Font(bold=True)
    
    for fila, pdf in enumerate(pdfs, start=2):
        nombre = pdf['name']
        link = pdf.get('webViewLink', f"https://drive.google.com/file/d/{pdf['id']}/view")
        
        ws.cell(row=fila, column=1, value=nombre)
        celda_link = ws.cell(row=fila, column=2, value=link)
        celda_link.hyperlink = link
        celda_link.font = Font(color="0563C1", underline="single")
    
    ws.column_dimensions['A'].width = 50
    ws.column_dimensions['B'].width = 60
    
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    
    nombre_archivo_excel = f"Listado_PDFs_{nombre_subcarpeta}.xlsx"
    
    metadata = {
        'name': nombre_archivo_excel,
        'parents': [id_subcarpeta],
        'mimeType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }
    
    media = MediaIoBaseUpload(
        buffer,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        resumable=True
    )
    
    archivo_creado = drive_service.files().create(
        body=metadata,
        media_body=media,
        fields='id, webViewLink'
    ).execute()
    
    print(f"\n Excel creado exitosamente: {nombre_archivo_excel}")
    print(f"🔗 Link del Excel: {archivo_creado.get('webViewLink')}")


if id_subcarpeta_elegida:
    crear_excel_con_pdfs(id_subcarpeta_elegida, nombre_subcarpeta_elegida)