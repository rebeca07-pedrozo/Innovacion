"""
COMPARADOR XML FORMATO 1020 - Davivienda
Compara archivo enviado a DIAN (Drive) vs archivo descargado de DIAN
"""

# ==============================================================================
# PASO 0: CONFIGURAR EN COLAB
# ==============================================================================
# Ejecuta esto primero:
# !pip install google-auth-oauthlib google-auth-httplib2 google-api-python-client lxml

from google.colab import auth
from googleapiclient.discovery import build
import xml.etree.ElementTree as ET
from io import BytesIO
import pandas as pd

# Autentica con Google Drive
auth.authenticate_user()
drive_service = build('drive', 'v3')

# ==============================================================================
# PASO 1: EDITA ESTOS IDS (son los que necesitas conseguir)
# ==============================================================================

# ID de la CARPETA en Drive donde guardaste el archivo que ENVIASTE a DIAN
FOLDER_ID_DRIVE_ENVIADO = "REEMPLAZA_CON_ID_CARPETA_DRIVE_1020"  # ← CAMBIAR AQUI

# ID de la CARPETA donde bajaste el archivo DESCARGADO de DIAN
FOLDER_ID_DESCARGAS_DIAN = "REEMPLAZA_CON_ID_CARPETA_DESCARGAS"  # ← CAMBIAR AQUI

# ==============================================================================
# PASO 2: FUNCIONES HELPER
# ==============================================================================

def obtener_primer_xml_carpeta(folder_id):
    """Obtiene el primer archivo .xml de una carpeta"""
    try:
        results = drive_service.files().list(
            q=f"'{folder_id}' in parents and name contains '.xml' and trashed=false",
            spaces='drive',
            fields='files(id, name, createdTime)',
            pageSize=10,
            orderBy='createdTime desc'
        ).execute()
        
        files = results.get('files', [])
        
        if not files:
            return None, None
        
        # Retorna el primero (más reciente)
        file_info = files[0]
        return file_info['id'], file_info['name']
    
    except Exception as e:
        print(f"❌ Error al acceder a carpeta {folder_id}: {e}")
        return None, None

def descargar_archivo_drive(file_id):
    """Descarga contenido del archivo desde Drive"""
    try:
        request = drive_service.files().get_media(fileId=file_id)
        file_content = request.execute()
        return file_content
    except Exception as e:
        print(f"❌ Error descargando archivo {file_id}: {e}")
        return None

def parsear_xml(contenido):
    """Parsea XML y retorna árbol + diccionario de elementos"""
    try:
        root = ET.fromstring(contenido)
        return root
    except Exception as e:
        print(f"❌ Error parseando XML: {e}")
        return None

def extraer_estructura_xml(root):
    """Extrae estructura y campos del XML"""
    estructura = {
        'raiz': root.tag,
        'elementos_nivel_1': [],
        'atributos_raiz': root.attrib,
        'total_elementos': len(list(root)),
    }
    
    for elem in root:
        estructura['elementos_nivel_1'].append({
            'tag': elem.tag,
            'atributos': elem.attrib,
            'texto': elem.text[:50] if elem.text else '',
            'subelementos': len(list(elem))
        })
    
    return estructura

def comparar_xmls(xml1, xml2, nombre1, nombre2):
    """Compara dos XMLs y genera reporte de diferencias"""
    
    print("\n" + "="*80)
    print(f"📊 COMPARACION XML FORMATO 1020")
    print(f"   Archivo 1 (DRIVE/ENVIADO): {nombre1}")
    print(f"   Archivo 2 (DESCARGAS/DIAN): {nombre2}")
    print("="*80 + "\n")
    
    struct1 = extraer_estructura_xml(xml1)
    struct2 = extraer_estructura_xml(xml2)
    
    diferencias = []
    
    # 1. Comparar elemento raíz
    if struct1['raiz'] != struct2['raiz']:
        diferencias.append({
            'Tipo': 'ESTRUCTURA',
            'Campo': 'Elemento Raíz',
            'Archivo 1 (Drive)': struct1['raiz'],
            'Archivo 2 (DIAN)': struct2['raiz'],
            'Crítico': '🔴 SÍ'
        })
    
    # 2. Comparar cantidad de elementos
    diff_cantidad = struct2['total_elementos'] - struct1['total_elementos']
    if diff_cantidad != 0:
        diferencias.append({
            'Tipo': 'CANTIDAD',
            'Campo': f"Total elementos (diferencia: {diff_cantidad:+d})",
            'Archivo 1 (Drive)': struct1['total_elementos'],
            'Archivo 2 (DIAN)': struct2['total_elementos'],
            'Crítico': '🟡 Revisar' if abs(diff_cantidad) > 5 else '🟢 Minor'
        })
    
    # 3. Comparar atributos de raíz
    attrs1 = set(struct1['atributos_raiz'].keys())
    attrs2 = set(struct2['atributos_raiz'].keys())
    
    attrs_faltantes = attrs1 - attrs2
    attrs_nuevos = attrs2 - attrs1
    
    if attrs_faltantes:
        diferencias.append({
            'Tipo': 'ATRIBUTOS',
            'Campo': f"Atributos FALTANTES en DIAN",
            'Archivo 1 (Drive)': ', '.join(attrs_faltantes),
            'Archivo 2 (DIAN)': '(no presente)',
            'Crítico': '🔴 SÍ'
        })
    
    if attrs_nuevos:
        diferencias.append({
            'Tipo': 'ATRIBUTOS',
            'Campo': f"Atributos NUEVOS en DIAN",
            'Archivo 1 (Drive)': '(no presente)',
            'Archivo 2 (DIAN)': ', '.join(attrs_nuevos),
            'Crítico': '🟡 Revisar'
        })
    
    # 4. Comparar elementos nivel 1
    tags1 = {e['tag']: e for e in struct1['elementos_nivel_1']}
    tags2 = {e['tag']: e for e in struct2['elementos_nivel_1']}
    
    tags_faltantes = set(tags1.keys()) - set(tags2.keys())
    tags_nuevos = set(tags2.keys()) - set(tags1.keys())
    
    if tags_faltantes:
        for tag in tags_faltantes:
            diferencias.append({
                'Tipo': 'ELEMENTOS',
                'Campo': f"Elemento <{tag}> FALTANTE en DIAN",
                'Archivo 1 (Drive)': 'Presente',
                'Archivo 2 (DIAN)': 'FALTA',
                'Crítico': '🔴 SÍ'
            })
    
    if tags_nuevos:
        for tag in tags_nuevos:
            diferencias.append({
                'Tipo': 'ELEMENTOS',
                'Campo': f"Elemento <{tag}> NUEVO en DIAN",
                'Archivo 1 (Drive)': 'No existe',
                'Archivo 2 (DIAN)': 'Presente',
                'Crítico': '🟡 Revisar'
            })
    
    # 5. Comparar valores de elementos comunes
    for tag in tags1.keys() & tags2.keys():
        if tags1[tag]['texto'] != tags2[tag]['texto']:
            diferencias.append({
                'Tipo': 'VALORES',
                'Campo': f"<{tag}> - Contenido diferente",
                'Archivo 1 (Drive)': tags1[tag]['texto'][:40],
                'Archivo 2 (DIAN)': tags2[tag]['texto'][:40],
                'Crítico': '🟡 Revisar'
            })
    
    return diferencias

# ==============================================================================
# PASO 3: EJECUCION PRINCIPAL
# ==============================================================================

print("🔍 Buscando archivos XML en Drive...")

# Obtener primer XML de carpeta 1 (Drive/Enviado)
file_id_1, nombre_1 = obtener_primer_xml_carpeta(FOLDER_ID_DRIVE_ENVIADO)
if not file_id_1:
    print(f"❌ No se encontró XML en carpeta Drive {FOLDER_ID_DRIVE_ENVIADO}")
    print("   Verifica que el ID sea correcto y que contenga archivos .xml")
else:
    print(f"✅ Archivo Drive: {nombre_1}")

# Obtener primer XML de carpeta 2 (Descargas/DIAN)
file_id_2, nombre_2 = obtener_primer_xml_carpeta(FOLDER_ID_DESCARGAS_DIAN)
if not file_id_2:
    print(f"❌ No se encontró XML en carpeta Descargas {FOLDER_ID_DESCARGAS_DIAN}")
    print("   Verifica que el ID sea correcto y que contenga archivos .xml")
else:
    print(f"✅ Archivo DIAN: {nombre_2}")

if file_id_1 and file_id_2:
    # Descargar ambos archivos
    print("\n📥 Descargando archivos...")
    contenido_1 = descargar_archivo_drive(file_id_1)
    contenido_2 = descargar_archivo_drive(file_id_2)
    
    if contenido_1 and contenido_2:
        # Parsear XMLs
        print("⚙️ Parseando XMLs...")
        xml1 = parsear_xml(contenido_1)
        xml2 = parsear_xml(contenido_2)
        
        if xml1 and xml2:
            # Comparar
            lista_difs = comparar_xmls(xml1, xml2, nombre_1, nombre_2)
            
            # Mostrar resultados
            if not lista_difs:
                print("✅ ¡No hay diferencias! Archivos idénticos.")
            else:
                print(f"\n⚠️ SE ENCONTRARON {len(lista_difs)} DIFERENCIA(S):\n")
                
                # Crear DataFrame
                df = pd.DataFrame(lista_difs)
                print(df.to_string(index=False))
                
                # Resumen por tipo
                print("\n" + "="*80)
                print("📋 RESUMEN POR TIPO:")
                print("="*80)
                for tipo in df['Tipo'].unique():
                    count = len(df[df['Tipo'] == tipo])
                    criticos = len(df[(df['Tipo'] == tipo) & (df['Crítico'] == '🔴 SÍ')])
                    print(f"  {tipo}: {count} diferencia(s) ({criticos} crítica(s))")
                
                # Exportar a CSV (opcional)
                csv_filename = f"comparacion_1020_{nombre_1.replace('.xml', '')}_vs_{nombre_2.replace('.xml', '')}.csv"
                df.to_csv(csv_filename, index=False, encoding='utf-8-sig')
                print(f"\n💾 Reporte exportado a: {csv_filename}")
        else:
            print("❌ Error al parsear uno de los XMLs")
    else:
        print("❌ Error descargando archivos")

print("\n" + "="*80)
print("✨ Proceso terminado")
print("="*80)
