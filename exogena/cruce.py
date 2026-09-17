from google.colab import auth
from googleapiclient.discovery import build
import xml.etree.ElementTree as ET
import pandas as pd

auth.authenticate_user()
drive = build('drive', 'v3')

# ============================================
# EDITA ESTOS 2 IDS
# ============================================
FOLDER_DESCARGADOS = "TU_ID_CARPETA_DESCARGAS"  # Los que bajaste de DIAN
FOLDER_JEFA = "TU_ID_CARPETA_JEFA"             # Los que tu jefa tiene

# ============================================

def listar_xmls(folder_id):
    results = drive.files().list(
        q=f"'{folder_id}' in parents and name contains '.xml' and trashed=false",
        spaces='drive',
        fields='files(id, name)',
        pageSize=1000
    ).execute()
    return sorted(results.get('files', []), key=lambda x: x['name'])

def descargar_xml(file_id):
    request = drive.files().get_media(fileId=file_id)
    return request.execute()

def recorrer_xml(elem, parent_path="", lista_elementos=None):
    """Recorre TODO el XML y guarda cada etiqueta con su valor completo"""
    if lista_elementos is None:
        lista_elementos = []
    
    # Path actual
    path = f"{parent_path}/{elem.tag}" if parent_path else elem.tag
    
    # Valor
    valor = (elem.text.strip() if elem.text else "").replace('\n', ' ')
    
    # Atributos
    atributos_str = " | ".join([f"{k}={v}" for k, v in elem.attrib.items()])
    
    # Guardar este elemento
    lista_elementos.append({
        'path': path,
        'tag': elem.tag,
        'valor': valor,
        'atributos': atributos_str,
        'tiene_datos': bool(valor or elem.attrib)
    })
    
    # Recorrer hijos
    for child in elem:
        recorrer_xml(child, path, lista_elementos)
    
    return lista_elementos

def parsear_xml_completo(xml_content):
    """Parsea el XML y retorna lista de TODOS los elementos"""
    try:
        root = ET.fromstring(xml_content)
        elementos = recorrer_xml(root)
        return elementos
    except Exception as e:
        print(f"Error parseando: {e}")
        return []

def comparar_dos_xmls(elem_dian, elem_jefa, nombre_archivo):
    """Compara dos listas de elementos y encuentra diferencias"""
    diferencias = []
    
    # Crear dict por path para buscar rápido
    jefa_dict = {e['path']: e for e in elem_jefa}
    dian_dict = {e['path']: e for e in elem_dian}
    
    # 1. Elementos que FALTAN en DIAN (están en Jefa pero no en DIAN)
    paths_jefa = set(jefa_dict.keys())
    paths_dian = set(dian_dict.keys())
    
    faltantes = paths_jefa - paths_dian
    for path in sorted(faltantes):
        elem_j = jefa_dict[path]
        diferencias.append({
            'Archivo': nombre_archivo,
            'Tipo': '❌ FALTA en DIAN',
            'Etiqueta': elem_j['tag'],
            'Ruta': path,
            'Valor en Tu Archivo (DIAN)': '(NO EXISTE)',
            'Valor en Archivo Jefa': elem_j['valor'][:100] if elem_j['valor'] else '(vacío)',
            'Crítico': '🔴 CRÍTICO'
        })
    
    # 2. Elementos que están en AMBOS pero con VALORES DIFERENTES
    comunes = paths_jefa & paths_dian
    for path in sorted(comunes):
        elem_j = jefa_dict[path]
        elem_d = dian_dict[path]
        
        val_j = elem_j['valor']
        val_d = elem_d['valor']
        
        # Comparar valores
        if val_j != val_d:
            diferencias.append({
                'Archivo': nombre_archivo,
                'Tipo': '⚠️ VALOR DIFERENTE',
                'Etiqueta': elem_j['tag'],
                'Ruta': path,
                'Valor en Tu Archivo (DIAN)': val_d[:100] if val_d else '(vacío)',
                'Valor en Archivo Jefa': val_j[:100] if val_j else '(vacío)',
                'Crítico': '🔴 CRÍTICO'
            })
        
        # Comparar atributos
        if elem_j['atributos'] != elem_d['atributos']:
            diferencias.append({
                'Archivo': nombre_archivo,
                'Tipo': '⚠️ ATRIBUTOS DIFERENTES',
                'Etiqueta': elem_j['tag'],
                'Ruta': path,
                'Valor en Tu Archivo (DIAN)': elem_d['atributos'][:100] if elem_d['atributos'] else '(sin atributos)',
                'Valor en Archivo Jefa': elem_j['atributos'][:100] if elem_j['atributos'] else '(sin atributos)',
                'Crítico': '🟡 REVISAR'
            })
    
    return diferencias

# ============================================
# EJECUCION
# ============================================

print("📥 Descargando listas de archivos...\n")

xmls_dian = listar_xmls(FOLDER_DESCARGADOS)
xmls_jefa = listar_xmls(FOLDER_JEFA)

print(f"✅ DIAN: {len(xmls_dian)} archivos")
print(f"✅ JEFA: {len(xmls_jefa)} archivos\n")

# Crear dicts por nombre
dian_dict = {f['name']: f for f in xmls_dian}
jefa_dict = {f['name']: f for f in xmls_jefa}

todos_diffs = []

print("🔍 COMPARANDO CONTENIDO...\n")

for nombre in sorted(dian_dict.keys()):
    if nombre in jefa_dict:
        print(f"  📄 {nombre}")
        
        # Descargar
        cont_dian = descargar_xml(dian_dict[nombre]['id'])
        cont_jefa = descargar_xml(jefa_dict[nombre]['id'])
        
        # Parsear
        elem_dian = parsear_xml_completo(cont_dian)
        elem_jefa = parsear_xml_completo(cont_jefa)
        
        # Comparar
        diffs = comparar_dos_xmls(elem_dian, elem_jefa, nombre)
        todos_diffs.extend(diffs)
        
        if diffs:
            print(f"     ⚠️ {len(diffs)} diferencia(s) encontrada(s)")
        else:
            print(f"     ✅ Idéntico")

# ============================================
# EXPORTAR
# ============================================

print("\n" + "="*80)

if todos_diffs:
    df = pd.DataFrame(todos_diffs)
    
    # Ordenar por archivo y criticidad
    df['orden_critico'] = df['Crítico'].map({'🔴 CRÍTICO': 0, '🟡 REVISAR': 1})
    df = df.sort_values(['Archivo', 'orden_critico']).drop('orden_critico', axis=1)
    
    # Exportar XLSX
    xlsx_name = "comparacion_1020_DETALLADA.xlsx"
    df.to_excel(xlsx_name, index=False, sheet_name='Diferencias', engine='openpyxl')
    
    print(f"✅ ARCHIVO EXPORTADO: {xlsx_name}\n")
    print(f"📊 RESUMEN GENERAL:")
    print(f"   Total diferencias: {len(todos_diffs)}")
    print(f"\n   Por TIPO:")
    for tipo in df['Tipo'].unique():
        count = len(df[df['Tipo'] == tipo])
        print(f"      {tipo}: {count}")
    
    print(f"\n   Por CRITICIDAD:")
    print(f"      🔴 CRÍTICO: {len(df[df['Crítico'] == '🔴 CRÍTICO'])}")
    print(f"      🟡 REVISAR: {len(df[df['Crítico'] == '🟡 REVISAR'])}")
    
    print(f"\n📋 PRIMERAS 15 DIFERENCIAS:")
    cols_mostrar = ['Archivo', 'Tipo', 'Etiqueta', 'Valor en Tu Archivo (DIAN)', 'Valor en Archivo Jefa']
    print(df[cols_mostrar].head(15).to_string(index=False))
    
else:
    print("✅ NO HAY DIFERENCIAS - Todos los archivos son idénticos")

print("\n" + "="*80)