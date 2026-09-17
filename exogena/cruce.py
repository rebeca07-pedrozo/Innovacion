
from google.colab import auth
from googleapiclient.discovery import build
import xml.etree.ElementTree as ET
import pandas as pd
from io import BytesIO

# Autentica
auth.authenticate_user()
drive = build('drive', 'v3')

# ============================================
# EDITA ESTOS 2 IDS
# ============================================
FOLDER_DESCARGADOS = "TU_ID_CARPETA_DESCARGAS"  # Los que bajaste de DIAN
FOLDER_JEFA = "TU_ID_CARPETA_JEFA"             # Los que tu jefa tiene

# ============================================
# Funciones
# ============================================

def listar_xmls(folder_id):
    """List todos los XMLs en una carpeta"""
    results = drive.files().list(
        q=f"'{folder_id}' in parents and name contains '.xml' and trashed=false",
        spaces='drive',
        fields='files(id, name)',
        pageSize=1000
    ).execute()
    
    files = results.get('files', [])
    # Ordena por nombre para que queden secuenciales
    return sorted(files, key=lambda x: x['name'])

def descargar_xml(file_id):
    """Descarga contenido del XML"""
    request = drive.files().get_media(fileId=file_id)
    return request.execute()

def extraer_elementos(xml_content):
    """Parsea XML y retorna dict de elementos con sus valores"""
    try:
        root = ET.fromstring(xml_content)
        elementos = {}
        
        def recorrer(elem, path=""):
            current_path = f"{path}/{elem.tag}"
            # Guarda el elemento con su valor y atributos
            elementos[current_path] = {
                'valor': elem.text.strip() if elem.text else '',
                'atributos': elem.attrib,
                'tag': elem.tag
            }
            for child in elem:
                recorrer(child, current_path)
        
        for child in root:
            recorrer(child)
        
        return elementos
    except:
        return {}

def comparar_xmls(elementos_dian, elementos_jefa, nombre_archivo):
    """Compara dos XMLs y retorna diferencias"""
    diferencias = []
    
    # 1. Etiquetas que faltan en DIAN (están en jefa pero no en DIAN)
    para_jefa = set(elementos_jefa.keys())
    para_dian = set(elementos_dian.keys())
    
    faltantes = para_jefa - para_dian
    for elemento in faltantes:
        valor_jefa = elementos_jefa[elemento]['valor']
        diferencias.append({
            'Archivo': nombre_archivo,
            'Tipo': 'FALTA en DIAN',
            'Etiqueta': elemento.split('/')[-1],
            'Ruta completa': elemento,
            'Valor en Jefa': valor_jefa[:80] if valor_jefa else '(vacío)',
            'Valor en DIAN': 'NO EXISTE',
            'Crítico': '🔴 SÍ' if valor_jefa else '🟡 Revisar'
        })
    
    # 2. Valores diferentes en etiquetas comunes
    comunes = para_jefa & para_dian
    for elemento in comunes:
        val_jefa = elementos_jefa[elemento]['valor']
        val_dian = elementos_dian[elemento]['valor']
        
        if val_jefa != val_dian:
            diferencias.append({
                'Archivo': nombre_archivo,
                'Tipo': 'VALOR DIFERENTE',
                'Etiqueta': elemento.split('/')[-1],
                'Ruta completa': elemento,
                'Valor en Jefa': val_jefa[:80],
                'Valor en DIAN': val_dian[:80],
                'Crítico': '🔴 SÍ'
            })
    
    return diferencias

# ============================================
# EJECUCION
# ============================================

print("🔄 Leyendo archivos de ambas carpetas...")

xmls_dian = listar_xmls(FOLDER_DESCARGADOS)
xmls_jefa = listar_xmls(FOLDER_JEFA)

print(f"✅ Encontrados {len(xmls_dian)} archivos en DIAN")
print(f"✅ Encontrados {len(xmls_jefa)} archivos de Jefa")

# Emparejar por nombre
dian_dict = {f['name']: f for f in xmls_dian}
jefa_dict = {f['name']: f for f in xmls_jefa}

todos_los_diffs = []

print("\n🔍 Comparando archivos...\n")

for nombre in sorted(dian_dict.keys()):
    if nombre in jefa_dict:
        print(f"  Comparando: {nombre}")
        
        # Descargar ambos
        contenido_dian = descargar_xml(dian_dict[nombre]['id'])
        contenido_jefa = descargar_xml(jefa_dict[nombre]['id'])
        
        # Extraer elementos
        elem_dian = extraer_elementos(contenido_dian)
        elem_jefa = extraer_elementos(contenido_jefa)
        
        # Comparar
        diffs = comparar_xmls(elem_dian, elem_jefa, nombre)
        todos_los_diffs.extend(diffs)
    else:
        print(f"  ⚠️ {nombre} NO EXISTE en carpeta de Jefa")

# ============================================
# Exportar a XLSX
# ============================================

if todos_los_diffs:
    df = pd.DataFrame(todos_los_diffs)
    
    # Guardar en CSV primero (más fácil en Colab)
    csv_name = "comparacion_1020_DIAN_vs_JEFA.csv"
    df.to_csv(csv_name, index=False, encoding='utf-8-sig')
    
    # Si quieres XLSX, también lo hacemos
    xlsx_name = "comparacion_1020_DIAN_vs_JEFA.xlsx"
    df.to_excel(xlsx_name, index=False, sheet_name='Diferencias')
    
    print(f"\n✅ EXPORTADO!")
    print(f"   CSV: {csv_name}")
    print(f"   XLSX: {xlsx_name}")
    print(f"\n📊 RESUMEN:")
    print(f"   Total diferencias encontradas: {len(todos_los_diffs)}")
    print(f"\n   Por tipo:")
    for tipo in df['Tipo'].unique():
        count = len(df[df['Tipo'] == tipo])
        criticos = len(df[(df['Tipo'] == tipo) & (df['Crítico'] == '🔴 SÍ')])
        print(f"      {tipo}: {count} ({criticos} crítica(s))")
    
    # Vista previa de las primeras 10
    print(f"\n📋 VISTA PREVIA (primeras 10):")
    print(df[['Archivo', 'Tipo', 'Etiqueta', 'Valor en Jefa', 'Valor en DIAN', 'Crítico']].head(10).to_string(index=False))
else:
    print("\n✅ ¡Sin diferencias! Todos los archivos están iguales")

print("\n" + "="*80)
print("✨ Listo")
print("="*80)

