from google.colab import auth, files
from googleapiclient.discovery import build
import xml.etree.ElementTree as ET
import pandas as pd

auth.authenticate_user()
drive = build('drive', 'v3')

FOLDER_DESCARGADOS = "1UbGhdnJb2b5RbLRj3gCizK84fYMOGvl5"
FOLDER_JEFA = "1eGZidzTrh19b5M6a30DjhR14X39dq-_a"

# ============================================

def listar_xmls(folder_id):
    results = drive.files().list(
        q=f"'{folder_id}' in parents and trashed=false",
        spaces='drive',
        fields='files(id, name)',
        pageSize=1000
    ).execute()
    files_list = results.get('files', [])
    xmls = [f for f in files_list if f['name'].endswith('.xml')]
    return sorted(xmls, key=lambda x: x['name'])

def descargar_xml(file_id):
    request = drive.files().get_media(fileId=file_id)
    return request.execute()

def extraer_valores(xml_content):
    """Extrae todos los valores exactos del XML"""
    try:
        root = ET.fromstring(xml_content)
        valores = {}
        
        def recorrer(elem, path=""):
            current_path = f"{path}/{elem.tag}" if path else elem.tag
            
            valor_text = elem.text if elem.text else ""
            attrs_str = " | ".join([f"{k}={v}" for k, v in elem.attrib.items()])
            
            valores[current_path] = {
                'valor': valor_text,
                'atributos': attrs_str
            }
            
            for child in elem:
                recorrer(child, current_path)
        
        for child in root:
            recorrer(child)
        
        return valores
    except Exception as e:
        print(f"Error: {e}")
        return {}

def comparar_archivos(val_dian, val_jefa, nombre_archivo):
    """Compara dos archivos XML"""
    diferencias = []
    
    paths_dian = set(val_dian.keys())
    paths_jefa = set(val_jefa.keys())
    
    # 1. ETIQUETAS QUE FALTAN EN DIAN (están en jefa pero NO en dian)
    faltantes = paths_jefa - paths_dian
    for path in sorted(faltantes):
        info_j = val_jefa[path]
        valor_jefa = info_j['valor'] if info_j['valor'] else f"[ATRIBUTOS: {info_j['atributos']}]"
        
        diferencias.append({
            'Archivo': nombre_archivo,
            'Etiqueta': path.split('/')[-1],
            'Qué pasó': '❌ FALTA EN TU ARCHIVO',
            'Ubicación (Ruta)': path,
            'Lo que TÚ tienes': '(NO EXISTE)',
            'Lo que JEFA tiene': valor_jefa[:150]
        })
    
    # 2. VALORES DIFERENTES (están en ambos pero distintos)
    comunes = paths_dian & paths_jefa
    for path in sorted(comunes):
        val_d = val_dian[path]['valor']
        val_j = val_jefa[path]['valor']
        
        if val_d != val_j:
            diferencias.append({
                'Archivo': nombre_archivo,
                'Etiqueta': path.split('/')[-1],
                'Qué pasó': '⚠️ VALOR DIFERENTE',
                'Ubicación (Ruta)': path,
                'Lo que TÚ tienes': val_d[:150] if val_d else '(vacío)',
                'Lo que JEFA tiene': val_j[:150] if val_j else '(vacío)'
            })
    
    # 3. ETIQUETAS QUE SOBRAN EN DIAN (están en dian pero NO en jefa)
    sobrantes = paths_dian - paths_jefa
    for path in sorted(sobrantes):
        info_d = val_dian[path]
        valor_dian = info_d['valor'] if info_d['valor'] else f"[ATRIBUTOS: {info_d['atributos']}]"
        
        diferencias.append({
            'Archivo': nombre_archivo,
            'Etiqueta': path.split('/')[-1],
            'Qué pasó': '⚠️ EXTRA EN TU ARCHIVO',
            'Ubicación (Ruta)': path,
            'Lo que TÚ tienes': valor_dian[:150],
            'Lo que JEFA tiene': '(NO EXISTE)'
        })
    
    return diferencias

# ============================================
# EJECUCION
# ============================================

print("📥 Buscando XMLs...\n")

xmls_dian = listar_xmls(FOLDER_DESCARGADOS)
xmls_jefa = listar_xmls(FOLDER_JEFA)

print(f"✅ TU CARPETA (DIAN): {len(xmls_dian)} XMLs")
print(f"✅ CARPETA JEFA: {len(xmls_jefa)} XMLs\n")

if len(xmls_dian) == 0:
    print("❌ NO ENCONTRÉ XMLS - VERIFICA LOS IDS")
else:
    dian_dict = {f['name']: f for f in xmls_dian}
    jefa_dict = {f['name']: f for f in xmls_jefa}
    
    todos_diffs = []
    contador = 0
    
    print("🔍 COMPARANDO (esto toma un rato)...\n")
    
    for nombre in sorted(dian_dict.keys()):
        contador += 1
        if contador % 10 == 1:
            print(f"  Procesando: {contador}/{len(dian_dict)}")
        
        if nombre in jefa_dict:
            try:
                cont_dian = descargar_xml(dian_dict[nombre]['id'])
                cont_jefa = descargar_xml(jefa_dict[nombre]['id'])
                
                val_dian = extraer_valores(cont_dian)
                val_jefa = extraer_valores(cont_jefa)
                
                diffs = comparar_archivos(val_dian, val_jefa, nombre)
                todos_diffs.extend(diffs)
                
            except Exception as e:
                print(f"     ❌ Error: {e}")
    
    print(f"\n✅ COMPARACION COMPLETADA\n")
    print(f"📊 TOTAL DIFERENCIAS ENCONTRADAS: {len(todos_diffs)}\n")
    
    if len(todos_diffs) > 0:
        df = pd.DataFrame(todos_diffs)
        
        # Resumen
        print("📋 RESUMEN:")
        for tipo in df['Qué pasó'].unique():
            count = len(df[df['Qué pasó'] == tipo])
            print(f"  {tipo}: {count} diferencias")
        
        # Exportar
        xlsx_file = "COMPARACION_1020_DETALLADA.xlsx"
        
        with pd.ExcelWriter(xlsx_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Diferencias', index=False)
            
            # FORMATO
            from openpyxl.styles import PatternFill, Font, Alignment
            ws = writer.sheets['Diferencias']
            
            # Header rojo
            header_fill = PatternFill(start_color="CC0000", end_color="CC0000", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF", size=11)
            
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(wrap_text=True)
            
            # Colorear filas por tipo
            for idx, row in df.iterrows():
                tipo = row['Qué pasó']
                
                if '❌' in tipo:  # FALTA
                    fill = PatternFill(start_color="FF6666", end_color="FF6666", fill_type="solid")
                elif '⚠️' in tipo:  # DIFERENTE
                    fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
                else:
                    fill = None
                
                if fill:
                    for cell in ws[idx + 2]:
                        cell.fill = fill
                        cell.alignment = Alignment(wrap_text=True, vertical='top')
            
            # Ajustar anchos
            ws.column_dimensions['Archivo'].width = 45
            ws.column_dimensions['Ubicación (Ruta)'].width = 60
            ws.column_dimensions['Lo que TÚ tienes'].width = 45
            ws.column_dimensions['Lo que JEFA tiene'].width = 45
            ws.column_dimensions['Qué pasó'].width = 25
            
            # Altura filas
            ws.row_dimensions[1].height = 30
            for row in ws.iter_rows(min_row=2, max_row=len(df)+1):
                ws.row_dimensions[row[0].row].height = 40
        
        print(f"\n✅ EXCEL GUARDADO: {xlsx_file}")
        print("📥 DESCARGANDO AUTOMÁTICAMENTE...\n")
        
        files.download(xlsx_file)
        
        print("✅ ¡DESCARGA LISTA!\n")
        print("="*80)
        print("PRIMERAS 20 DIFERENCIAS:")
        print("="*80 + "\n")
        print(df[['Archivo', 'Qué pasó', 'Etiqueta', 'Lo que TÚ tienes', 'Lo que JEFA tiene']].head(20).to_string(index=False))
    
    else:
        print("✅ SIN DIFERENCIAS - TODO ESTÁ IGUAL")

print("\n" + "="*80)