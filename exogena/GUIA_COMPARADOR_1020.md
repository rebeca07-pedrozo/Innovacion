# 📊 Guía: Comparador XML Formato 1020 - Davivienda

## 🚀 En 5 pasos

### **PASO 1: Conseguir IDs de tus carpetas**

#### Carpeta 1: Drive (donde guardaste el 1020 enviado)
1. Abre Google Drive → encuentra la carpeta
2. Click derecho → **"Abrir"** (la URL change)
3. En la URL verás: `https://drive.google.com/drive/folders/AQUI_VA_EL_ID`
4. Copia solo la parte **AQUI_VA_EL_ID** (es un string largo)
5. **Ejemplo real:**
   ```
   https://drive.google.com/drive/folders/1a2b3c4d5e6f7g8h9i0j1k2l3m4n5o6p7
   
   ID = 1a2b3c4d5e6f7g8h9i0j1k2l3m4n5o6p7
   ```

#### Carpeta 2: Descargas (donde bajaste lo de DIAN)
- Mismo procedimiento

---

### **PASO 2: Abre Colab**

1. Ve a [colab.research.google.com](https://colab.research.google.com)
2. Click "Nuevo notebook"
3. Renómbralo: `Comparador_1020`

---

### **PASO 3: Copia el script**

1. **En la primera celda**, copia TODO esto:

```python
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
```

---

### **PASO 4: Edita los IDs**

En el script, busca estas líneas:

```python
FOLDER_ID_DRIVE_ENVIADO = "REEMPLAZA_CON_ID_CARPETA_DRIVE_1020"
FOLDER_ID_DESCARGAS_DIAN = "REEMPLAZA_CON_ID_CARPETA_DESCARGAS"
```

**Cambia por tus IDs reales:**

```python
FOLDER_ID_DRIVE_ENVIADO = "1a2b3c4d5e6f7g8h9i0j1k2l3m4n5o6p7"
FOLDER_ID_DESCARGAS_DIAN = "9z8y7x6w5v4u3t2s1r0q9p8o7n6m5l4k"
```

---

### **PASO 5: Ejecuta**

1. Click en la celda → `Ctrl+Enter` (o botón Play)
2. Cuando pida autenticación, acepta acceso a Drive
3. **¡Listo!** Verás la comparación

---

## 📋 Qué ve como resultado

### **Si NO hay diferencias:**
```
✅ ¡No hay diferencias! Archivos idénticos.
```

### **Si hay diferencias:**

```
📊 COMPARACION XML FORMATO 1020
   Archivo 1 (DRIVE/ENVIADO): Dmuisca_010120240001.xml
   Archivo 2 (DESCARGAS/DIAN): Dmuisca_010120240001.xml

Tipo          | Campo                              | Archivo 1 (Drive)     | Archivo 2 (DIAN)    | Crítico
--------------|------------------------------------|-----------------------|---------------------|----------
ESTRUCTURA    | Elemento Raíz                      | mas                   | mas                 | 🟢 No
CANTIDAD      | Total elementos (+5)               | 100                   | 105                 | 🟡 Revisar
ATRIBUTOS     | Atributos FALTANTES en DIAN       | version               | (no presente)       | 🔴 SÍ
ELEMENTOS     | Elemento <Cab> FALTANTE en DIAN   | Presente              | FALTA               | 🔴 SÍ
VALORES       | <Ano> - Contenido diferente       | 2024                  | 2023                | 🟡 Revisar

📋 RESUMEN POR TIPO:
  ESTRUCTURA: 1 diferencia(s) (0 crítica(s))
  CANTIDAD: 1 diferencia(s) (0 crítica(s))
  ATRIBUTOS: 1 diferencia(s) (1 crítica(s))
  ELEMENTOS: 2 diferencia(s) (2 crítica(s))
  VALORES: 1 diferencia(s) (0 crítica(s))

💾 Reporte exportado a: comparacion_1020_Dmuisca...csv
```

---

## 🎯 Leyenda de críticos

| Símbolo | Significado | Acción |
|---------|------------|--------|
| 🔴 SÍ | Crítico - Blocante | ⚠️ Revisar ya - Bloquea validación DIAN |
| 🟡 Revisar | Importante - Revisar | ✅ Revisar antes de enviar |
| 🟢 Minor | Menor - OK | ✅ Puede ignorarse |

---

## 🔧 Si hay problemas

### "No se encontró XML en carpeta"
- Verifica que el ID sea correcto
- Que la carpeta tenga archivos `.xml`
- Que no esté en Trash

### "Error de autenticación"
- Colab pedirá acceso a Drive
- Acepta cuando pida permiso
- Asegúrate de estar logueado en Google

### "Error parseando XML"
- El archivo podría estar corrupto
- Abre el XML en el navegador para verificar

---

## 💡 Tips

**Si quieres comparar específicos, edita esto:**

```python
# Cambiar de:
orderBy='createdTime desc'  # Más reciente

# A:
orderBy='name'  # Orden alfabético
```

**Para ver más archivos:**

```python
# Cambiar pageSize
pageSize=10  # a  pageSize=50
```

---

## ✅ Checklist antes de usar

- [ ] IDs de carpetas copiados correctamente
- [ ] Los archivos XML están en Drive
- [ ] Colab abierto
- [ ] Conectado a internet
- [ ] Google account activo

---

**¡Listo! Ahora ejecuta y mira las diferencias.** 🚀
