# INSTALACIÓN SISTEMA DE REQUERIMIENTOS - Guía Paso a Paso

## 📥 PASO 1: Descargar el Excel (2 minutos)

1. Descarga el archivo: **`Requerimientos_Legales_Template.xlsx`**
   - Este archivo ya tiene:
     - ✅ Todas las hojas creadas
     - ✅ Todas las columnas
     - ✅ Headers formateados (azul oscuro con letras blancas)
     - ✅ Datos iniciales de SLA
     - ✅ Instrucciones incluidas

2. Guarda en tu carpeta de descargas o donde prefieras

---

## ☁️ PASO 2: Subir a Google Drive (3 minutos)

1. Abre Google Drive: https://drive.google.com
2. En la sección izquierda, ve a donde guardas cosas de Davivienda
3. Click en **"Nuevo"** (arriba izquierda)
4. Selecciona **"Subir archivo"**
5. Busca y selecciona `Requerimientos_Legales_Template.xlsx`
6. Espera a que se cargue

**Importante:** Google Drive detectará que es un Excel y te ofrecerá abrir con Google Sheets automáticamente.

---

## 🔄 PASO 3: Convertir a Google Sheets (2 minutos)

1. El archivo ya debería estar en tu Drive
2. Haz **doble click** para abrirlo
3. Google Drive te pregunta: **"¿Abrir con Google Sheets?"**
4. Click **"Aceptar"** o **"Open with Google Sheets"**
5. ¡Listo! Ahora tienes un Google Sheets limpio y formateado

**Nota:** Aún no es automático. Seguimos con AppScript.

---

## 📝 PASO 4: Abrir Google Apps Script (2 minutos)

1. En tu nuevo Google Sheets, ve al menú superior
2. Click en **"Extensiones"** (entre Herramientas e Insertar)
3. Selecciona **"Apps Script"**
4. Se abrirá una nueva pestaña con el editor
5. **ELIMINA TODO EL CÓDIGO** que aparezca por defecto
   - Selecciona todo: `Ctrl+A`
   - Borra: `Delete` o `Backspace`

---

## 💾 PASO 5: Pegar el Código AppScript (3 minutos)

1. Abre el archivo `appscript-codigo-completo.gs`
2. Copia TODO el contenido (Ctrl+A, Ctrl+C)
3. Vuelve al editor de Apps Script
4. En el área en blanco, pega: `Ctrl+V`
5. Espera a que se cargue el código
6. Click **"Guardar"** (Ctrl+S)

**Nota:** Si ves errores subrayados en rojo, no importa. Son advertencias, no errores.

---

## ⚡ PASO 6: Ejecutar el Trigger Automático (2 minutos)

### Esto es CRÍTICO para que el sistema funcione

1. En el editor de Apps Script, arriba encontrarás un dropdown que dice: **"Ejecutar"**
2. Haz click en el dropdown
3. Busca y selecciona: **`crearTriggerAutomatico`**
4. Presiona el botón **▶️ Play** (justo a la izquierda del dropdown)
5. Te preguntará por permisos. Click en **"Autorizar"**
6. Elige tu cuenta de Google
7. Te dirá "Davivienda necesita acceso"
8. Click **"Permitir"**

**Si todo funciona:**
- Verás en la pantalla: `✅ Trigger creado: verificará correos cada 5 minutos`
- En los Logs verás mensajes verdes con ✅

---

## 🧪 PASO 7: Primer Test con Template 1 (5 minutos)

### Ahora probamos si el sistema captura correos

1. Abre **Gmail**
2. Click en **"Redactar"** (o "Compose")
3. Completa:
   - **Para:** `notificacionesjudiciales@davivienda.co`
   - **Asunto:** Copia exactamente:
     ```
     Oficio No. PC 0156 del 18 de septiembre de 2024 - Pliego de Cargos
     ```
   - **Cuerpo:** Copia todo del Template 1 (de testing-templates-correos.md):
     ```
     Estimado Banco Davivienda,

     Por este medio, la Secretaría de Hacienda de Medellín, le comunica 
     que se ha expedido PLIEGO DE CARGOS en contra de la entidad 
     relacionado con incumplimiento en la presentación de declaraciones 
     y pagos de impuestos municipales.

     DETALLES DEL REQUERIMIENTO:
     Número de Pliego: PC 0156
     Fecha de expedición: 18 de septiembre de 2024
     Contribuyente: Banco Davivienda S.A.
     Período gravable: 2024

     Para efectos de respuesta, se concede un término de VEINTE (20) DÍAS 
     contados a partir de la fecha de notificación.

     Fecha de vencimiento para presentación de descargos: 08 de octubre de 2024

     Atentamente,

     Dr. Juan Carlos López Rodríguez
     Secretario de Hacienda
     Secretaría de Hacienda Municipal
     Medellín, Antioquia
     ```

4. Click **"Enviar"**
5. **ESPERA 5 MINUTOS** (el trigger se ejecuta cada 5 min)
6. Vuelve a tu Sheets
7. Abre la hoja **"REQ LEGALES SEC HACIENDA"**
8. ¿APARECE UNA FILA NUEVA? 👀

---

## 🔍 PASO 8: Revisar qué se Extrajo (5 minutos)

**En tu Sheets:**

| Columna | ¿Qué debería tener? | ¿Sí o No? |
|---------|-------------------|----------|
| ID_REQ | REQ-2024-09-XX-XXX | ✅ |
| STATUS | por-completar | ✅ |
| EMAIL_TIMESTAMP | 18/09/2024 14:30 | ✅ |
| ASUNTO_CORREO | Oficio No. PC 0156... | ✅ |
| CIUDAD_MUNICIPIO | Medellín | ✅ |
| DEPARTAMENTO | [vacío o Antioquia?] | ❓ |
| TIPO_OFICIO | Pliego de Cargos | ✅ |
| FECHA_VENCIMIENTO | 08/10/2024 | ✅ |

**Para ver los LOGS (el detalle de qué hizo el AppScript):**

1. Vuelve a la pestaña de **Apps Script Editor**
2. Abajo a la izquierda, click en **"Ejecuciones"**
3. Busca la ejecución más reciente de `verificaCorreosNuevos`
4. Haz click para abrir los logs
5. Lee qué dice:
   ```
   ✅ Municipio extraído: "Medellín"
   ✅ Tipo Oficio: "Pliego de Cargos"
   ✅ Fila agregada a Sheets - ID: REQ-2024-09-18-001
   ```

---

## 📋 PASO 9: Documentar Resultados (5 minutos)

Crea un documento con esto:

```
=== TEST 1: MEDELLÍN - PLIEGO DE CARGOS ===
Fecha de test: [HOY]
Hora de envío del email: [HORA]

✅ APARICIÓN EN SHEETS: SÍ / NO

CAMPOS RELLENADOS AUTOMÁTICAMENTE:
✅ ID_REQ = REQ-2024-09-18-XXX
✅ ASUNTO = Oficio No. PC 0156...
✅ MUNICIPIO = Medellín
✅ TIPO OFICIO = Pliego de Cargos
✅ VENCIMIENTO = 08/10/2024

CAMPOS QUE FALTARON (normales, se llenan manualmente):
- RESPONSABLE_ÁREA (vacío)
- CASO (vacío)
- ACCIÓN_EJECUTADA (vacío)

ERRORES EN LOGS:
[Si aparecen, cópialo aquí]

OBSERVACIONES:
[Describe qué viste, si fue extraño, etc.]
```

---

## 🎯 PASO 10: Próximos Tests (Opcional pero Recomendado)

Si el Template 1 funcionó, continúa con:

**Día 2:** Template 2 (Bogotá - Emplazamiento)
- Mismo proceso
- Documenta si se extrae bien el tipo "Emplazamiento"

**Día 3:** Template 3 (Cali - Requerimiento)
- Prueba otro tipo de oficio

**Día 4+:** Templates 4, 5, 6
- Válida edge cases

---

## ⚠️ SOLUCIÓN DE PROBLEMAS

### "No aparece fila nueva en Sheets"

**Checklist:**
- [ ] ¿Ejecutaste `crearTriggerAutomatico()` y autorizaste?
- [ ] ¿Enviaste el email a `notificacionesjudiciales@davivienda.co`?
- [ ] ¿Esperaste 5 minutos?
- [ ] ¿El Sheets está abierto en otra pestaña?

**Si nada funcionó:**
1. Ve a Apps Script → Ejecuciones
2. ¿Dice "ERROR"?
3. Haz click y lee el error completo
4. **Cópialo y envíamelo**

---

### "El AppScript no ejecuta"

**Solución:**
1. Apps Script Editor
2. Arriba, dropdown dice "Ejecutar"
3. Abrirlo y seleccionar: `crearTriggerAutomatico`
4. Presiona ▶️ Play
5. Autoriza si pide

---

### "¿Dónde veo los Logs?"

**Así:**
1. Apps Script → Logs (abajo izquierda, dice "Ejecuciones")
2. Busca la ejecución más reciente
3. Haz click para expandir
4. Lee los mensajes

**Verás algo como:**
```
📧 Correos nuevos encontrados: 1
[PROCESANDO CORREO 1]
Asunto: Oficio No. PC 0156...
[REGEX] Buscando municipio...
  ✅ Encontrado (patrón 1): "Medellín"
[REGEX] Buscando tipo de oficio...
  ✅ Tipo: Pliego de Cargos
✅ Fila agregada - ID: REQ-2024-09-18-001
✅ Procesamiento completado: 1 correo(s)
```

---

## 📞 RESUMEN DE ARCHIVOS A TENER

1. **Requerimientos_Legales_Template.xlsx** ← Descargas y subes a Drive
2. **appscript-codigo-completo.gs** ← Copias en Apps Script
3. **testing-templates-correos.md** ← Copias los correos de prueba
4. **requerimientos-sistema-completo.md** ← Referencia técnica (leer si hay dudas)

---

## ✅ CHECKLIST FINAL

- [ ] Descargué el Excel
- [ ] Lo subí a Drive
- [ ] Lo abrí como Google Sheets
- [ ] Abrí Apps Script
- [ ] Peguè el código
- [ ] Ejecuté `crearTriggerAutomatico()` y autorizé
- [ ] Envié email de prueba a notificacionesjudiciales@davivienda.co
- [ ] Esperé 5 minutos
- [ ] Volví a Sheets y chequeé si aparece fila nueva
- [ ] Documenté qué se extrajo vs qué no

**Si todo está OK, el sistema está listo para usar.** 🎉

Si algo falló, documenta y comparte conmigo qué error ves.
