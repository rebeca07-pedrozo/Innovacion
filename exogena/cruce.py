from google.colab import auth
from googleapiclient.discovery import build
import xml.etree.ElementTree as ET
import pandas as pd

auth.authenticate_user()
drive = build('drive', 'v3')

FOLDER_DESCARGADOS = "1UbGhdnJb2b5RbLRj3gCizK84fYMOGvl5"
FOLDER_JEFA = "1eGZidzTrh19b5M6a30DjhR14X39dq-_a"

# ============================================

def listar_xmls(folder_id):
    """Lista todos los archivos y filtra .xml en Python"""
    results = drive.files().list(
        q=f"'{folder_id}' in parents and trashed=false",
        spaces='drive',
        fields='files(id, name)',
        pageSize=1000
    ).execute()
    
    files = results.get('files', [])
    # Filtra .xml en Python, no en la API
    xmls = [f for f in files if f['name'].endswith('.xml')]
    return sorted(xmls, key=lambda x: x['name'])

def descargar_xml(file_id):
    request = drive.files().get_media(fileId=file_id)
    return request.execute()

def contar_titulares(xml_content):
    """Cuenta cuántos titulares hay (elementos Titl)"""
    try:
        root = ET.fromstring(xml_content)
        # Busca TODOS los elementos que contengan "Titl" en el nombre
        titulares = []
        
        def buscar_titulares(elem, path=""):
            # Si el tag contiene "Titl" o "titl", es un titular
            if 'titl' in elem.tag.lower():
                titulares.append({
                    'tag': elem.tag,
                    'path': f"{path}/{elem.tag}",
                    'valores': []
                })
            
            # Buscar en hijos
            for child in elem:
                buscar_titulares(child, f"{path}/{elem.tag}")
        
        buscar_titulares(root)
        return len(titulares), titulares
    except:
        return 0, []

def extraer_todos_datos(xml_content):
    """Extrae TODOS los elementos con valores"""
    try:
        root = ET.fromstring(xml_content)
        datos = {}
        
        def recorrer(elem, path=""):
            current_path = f"{path}/{elem.tag}" if path else elem.tag
            
            # Guardar valor
            valor = (elem.text.strip() if elem.text else "")
            atributos = elem.attrib
            
            datos[current_path] = {
                'valor': valor,
                'atributos': atributos,
                'tag': elem.tag
            }
            
            for child in elem:
                recorrer(child, current_path)
        
        for child in root:
            recorrer(child)
        
        return datos
    except Exception as e:
        print(f"Error: {e}")
        return {}

def comparar_dos_xmls(datos_dian, datos_jefa, nombre_archivo):
    """Compara REALMENTE los datos"""
    diferencias = []
    
    paths_dian = set(datos_dian.keys())
    paths_jefa = set(datos_jefa.keys())
    
    # FALTANTES en DIAN
    faltantes = paths_jefa - paths_dian
    for path in sorted(faltantes):
        elem_j = datos_jefa[path]
        diferencias.append({
            'Archivo': nombre_archivo,
            'Tipo': '❌ FALTA en DIAN',
            'Ruta': path,
            'Tu Archivo (DIAN)': '(NO EXISTE)',
            'Archivo Jefa': elem_j['valor'][:150] if elem_j['valor'] else '(vacío)',
        })
    
    # VALORES DIFERENTES
    comunes = paths_dian & paths_jefa
    for path in sorted(comunes):
        val_d = datos_dian[path]['valor']
        val_j = datos_jefa[path]['valor']
        
        if val_d != val_j:
            diferencias.append({
                'Archivo': nombre_archivo,
                'Tipo': '⚠️ VALOR DIFERENTE',
                'Ruta': path,
                'Tu Archivo (DIAN)': val_d[:150] if val_d else '(vacío)',
                'Archivo Jefa': val_j[:150] if val_j else '(vacío)',
            })
    
    return diferencias

# ============================================
# EJECUCION
# ============================================

print("📥 Buscando XMLs...\n")

xmls_dian = listar_xmls(FOLDER_DESCARGADOS)
xmls_jefa = listar_xmls(FOLDER_JEFA)

print(f"✅ DIAN: {len(xmls_dian)} XMLs encontrados")
print(f"✅ JEFA: {len(xmls_jefa)} XMLs encontrados\n")

if len(xmls_dian) == 0 or len(xmls_jefa) == 0:
    print("❌ NO SE ENCONTRARON XMLS")
    print("Verifica los IDs de las carpetas")
else:
    dian_dict = {f['name']: f for f in xmls_dian}
    jefa_dict = {f['name']: f for f in xmls_jefa}
    
    todos_diffs = []
    contador = 0
    
    print("🔍 COMPARANDO CONTENIDO...\n")
    
    for nombre in sorted(dian_dict.keys()):
        contador += 1
        if contador <= 5 or contador % 20 == 0:
            print(f"  [{contador}/{len(dian_dict)}] {nombre}")
        
        if nombre in jefa_dict:
            # Descargar
            try:
                cont_dian = descargar_xml(dian_dict[nombre]['id'])
                cont_jefa = descargar_xml(jefa_dict[nombre]['id'])
                
                # Extraer datos
                datos_dian = extraer_todos_datos(cont_dian)
                datos_jefa = extraer_todos_datos(cont_jefa)
                
                # Comparar
                diffs = comparar_dos_xmls(datos_dian, datos_jefa, nombre)
                todos_diffs.extend(diffs)
                
            except Exception as e:
                print(f"     ❌ Error: {e}")
    
    print(f"\n" + "="*80)
    print(f"✅ COMPARACION COMPLETADA\n")
    print(f"📊 TOTAL DIFERENCIAS ENCONTRADAS: {len(todos_diffs)}\n")
    
    if len(todos_diffs) > 0:
        # Crear DataFrame y exportar
        df = pd.DataFrame(todos_diffs)
        
        # Contar por tipo
        print("Por TIPO:")
        for tipo in df['Tipo'].unique():
            count = len(df[df['Tipo'] == tipo])
            print(f"  {tipo}: {count}")
        
        # EXPORTAR XLSX
        try:
            from openpyxl import Workbook
            from openpyxl.styles import PatternFill, Font
            
            xlsx_file = "comparacion_1020_DETALLADA.xlsx"
            
            # Crear workbook
            wb = Workbook()
            ws = wb.active
            ws.title = "Diferencias"
            
            # Headers
            headers = list(df.columns)
            ws.append(headers)
            
            # Estilos para header
            header_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF")
            
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
            
            # Agregar datos
            for idx, row in df.iterrows():
                ws.append(list(row))
                
                # Colorear si es FALTA (rojo) o DIFERENTE (amarillo)
                tipo = row['Tipo']
                if '❌' in tipo:
                    fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
                else:
                    fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
                
                for cell in ws[idx + 2]:
                    cell.fill = fill
            
            # Ajustar columnas
            ws.column_dimensions['Ruta'].width = 60
            ws.column_dimensions['Tu Archivo (DIAN)'].width = 40
            ws.column_dimensions['Archivo Jefa'].width = 40
            
            wb.save(xlsx_file)
            print(f"\n✅ XLSX GUARDADO: {xlsx_file}")
            
        except:
            # Si falla openpyxl, usar pandas
            xlsx_file = "comparacion_1020_DETALLADA.xlsx"
            df.to_excel(xlsx_file, index=False, sheet_name='Diferencias')
            print(f"\n✅ XLSX GUARDADO: {xlsx_file}")
        
        # Mostrar primeras 20
        print(f"\n📋 PRIMERAS 20 DIFERENCIAS:")
        print(df[['Archivo', 'Tipo', 'Ruta']].head(20).to_string(index=False))
    
    else:
        print("✅ NO HAY DIFERENCIAS DETECTADAS")
        print("   (Pero verifica manualmente si falta algo)")

print("\n" + "="*80)