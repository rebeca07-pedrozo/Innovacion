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

def extraer_valores_exactos(xml_content, archivo_nombre):
    """Extrae TODOS los valores de TODAS las etiquetas exactamente como están"""
    try:
        root = ET.fromstring(xml_content)
        valores = {}
        
        def recorrer(elem, path=""):
            current_path = f"{path}/{elem.tag}" if path else elem.tag
            
            # Valor EXACTO (sin strip, sin transformaciones)
            valor_text = elem.text if elem.text else ""
            
            # Si hay atributos, incluirlos
            attrs_str = " | ".join([f"{k}={v}" for k, v in elem.attrib.items()])
            
            valores[current_path] = {
                'valor': valor_text,
                'atributos': attrs_str,
                'tiene_contenido': bool(valor_text or elem.attrib)
            }
            
            for child in elem:
                recorrer(child, current_path)
        
        for child in root:
            recorrer(child)
        
        return valores
    except Exception as e:
        print(f"Error en {archivo_nombre}: {e}")
        return {}

def comparar_exacto(val_dian, val_jefa, nombre_archivo):
    """Compara valores EXACTAMENTE iguales"""
    diferencias = []
    
    paths_dian = set(val_dian.keys())
    paths_jefa = set(val_jefa.keys())
    
    # ETIQUETAS QUE FALTAN EN DIAN
    faltantes = paths_jefa - paths_dian
    for path in sorted(faltantes):
        info_j = val_jefa[path]
        diferencias.append({
            'Archivo': nombre_archivo,
            'Etiqueta': path.split('/')[-1],
            'Tipo': '❌ FALTA EN DIAN',
            'Ruta': path,
            'Tu Valor': '(NO EXISTE)',
            'Valor Jefa': info_j['valor'][:100] if info_j['valor'] else '(vacío)',
            'Atributos Jefa': info_j['atributos'][:100] if info_j['atributos'] else '-'
        })
    
    # ETIQUETAS QUE SOBRAN EN DIAN
    sobrantes = paths_dian - paths_jefa
    for path in sorted(sobrantes):
        info_d = val_dian[path]
        diferencias.append({
            'Archivo': nombre_archivo,
            'Etiqueta': path.split('/')[-1],
            'Tipo': '⚠️ EXTRA EN DIAN',
            'Ruta': path,
            'Tu Valor': info_d['valor'][:100] if info_d['valor'] else '(vacío)',
            'Valor Jefa': '(NO EXISTE)',
            'Atributos Jefa': '-'
        })
    
    # VALORES DIFERENTES EN ETIQUETAS COMUNES
    comunes = paths_dian & paths_jefa
    for path in sorted(comunes):
        val_d = val_dian[path]['valor']
        val_j = val_jefa[path]['valor']
        attr_d = val_dian[path]['atributos']
        attr_j = val_jefa[path]['atributos']
        
        # Compara valor exacto
        if val_d != val_j:
            diferencias.append({
                'Archivo': nombre_archivo,
                'Etiqueta': path.split('/')[-1],
                'Tipo': '⚠️ VALOR DIFERENTE',
                'Ruta': path,
                'Tu Valor': val_d[:100] if val_d else '(vacío)',
                'Valor Jefa': val_j[:100] if val_j else '(vacío)',
                'Atributos Jefa': attr_j[:100] if attr_j else '-'
            })
        
        # Compara atributos
        if attr_d != attr_j:
            diferencias.append({
                'Archivo': nombre_archivo,
                'Etiqueta': path.split('/')[-1],
                'Tipo': '⚠️ ATRIBUTOS DIFERENTES',
                'Ruta': path,
                'Tu Valor': attr_d[:100] if attr_d else '(sin atributos)',
                'Valor Jefa': attr_j[:100] if attr_j else '(sin atributos)',
                'Atributos Jefa': '-'
            })
    
    return diferencias

# ============================================
# MAIN
# ============================================

print("📥 Buscando XMLs...\n")

xmls_dian = listar_xmls(FOLDER_DESCARGADOS)
xmls_jefa = listar_xmls(FOLDER_JEFA)

print(f"✅ DIAN: {len(xmls_dian)} XMLs")
print(f"✅ JEFA: {len(xmls_jefa)} XMLs\n")

if len(xmls_dian) == 0:
    print("❌ NO SE ENCONTRARON XMLS - VERIFICA IDS")
else:
    dian_dict = {f['name']: f for f in xmls_dian}
    jefa_dict = {f['name']: f for f in xmls_jefa}
    
    todos_diffs = []
    contador = 0
    
    print("🔍 COMPARANDO (ESTO TOMA UN TIEMPO)...\n")
    
    for nombre in sorted(dian_dict.keys()):
        contador += 1
        if contador % 10 == 1:
            print(f"  {contador}/{len(dian_dict)}")
        
        if nombre in jefa_dict:
            try:
                cont_dian = descargar_xml(dian_dict[nombre]['id'])
                cont_jefa = descargar_xml(jefa_dict[nombre]['id'])
                
                val_dian = extraer_valores_exactos(cont_dian, nombre)
                val_jefa = extraer_valores_exactos(cont_jefa, nombre)
                
                diffs = comparar_exacto(val_dian, val_jefa, nombre)
                todos_diffs.extend(diffs)
                
            except Exception as e:
                print(f"     ❌ Error: {e}")
    
    print(f"\n✅ COMPLETADO\n")
    print(f"📊 DIFERENCIAS ENCONTRADAS: {len(todos_diffs)}\n")
    
    if len(todos_diffs) > 0:
        df = pd.DataFrame(todos_diffs)
        
        # Contar por tipo
        print("Por TIPO:")
        for tipo in df['Tipo'].unique():
            count = len(df[df['Tipo'] == tipo])
            print(f"  {tipo}: {count}")
        
        # Exportar
        xlsx_file = "COMPARACION_1020_EXACTA.xlsx"
        
        with pd.ExcelWriter(xlsx_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Diferencias', index=False)
            
            # Formato
            from openpyxl.styles import PatternFill, Font, Alignment
            ws = writer.sheets['Diferencias']
            
            # Header rojo
            header_fill = PatternFill(start_color="CC0000", end_color="CC0000", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF", size=11)
            
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(wrap_text=True)
            
            # Colorear filas
            for idx, row in df.iterrows():
                tipo = row['Tipo']
                if '❌' in tipo:
                    fill = PatternFill(start_color="FF6666", end_color="FF6666", fill_type="solid")
                elif '⚠️' in tipo:
                    fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
                else:
                    fill = None
                
                if fill:
                    for cell in ws[idx + 2]:
                        cell.fill = fill
                        cell.alignment = Alignment(wrap_text=True)
            
            # Ajustar ancho
            ws.column_dimensions['Ruta'].width = 50
            ws.column_dimensions['Tu Valor'].width = 40
            ws.column_dimensions['Valor Jefa'].width = 40
            ws.column_dimensions['Archivo'].width = 40
            
            # Alto filas
            ws.row_dimensions[1].height = 25
        
        print(f"✅ EXCEL GUARDADO: {xlsx_file}\n")
        print("📥 DESCARGANDO...\n")
        
        # DESCARGAR AUTOMATICO
        files.download(xlsx_file)
        
        print("✅ DESCARGA COMPLETA")
        print("\n📋 PRIMERAS 15 DIFERENCIAS:")
        print(df[['Archivo', 'Tipo', 'Etiqueta', 'Tu Valor', 'Valor Jefa']].head(15).to_string(index=False))
    
    else:
        print("✅ SIN DIFERENCIAS - TODO IDÉNTICO")

print("\n" + "="*80)