# Correos de Prueba - Rebeca se envía a sí misma

## SETUP PREVIO (1 min)

1. Abre Gmail
2. Ve a Configuración (rueda arriba a la derecha)
3. Click "Crear label"
4. Nombre: **`Testing_Reqs`**
5. Guardar

Listo. Ahora tienes un label para los correos de prueba.

---

## CÓMO ENVIAR CADA PRUEBA

```
DESDE:    rebeca.pedrozo@davivienda.com (tu correo)
PARA:     rebeca.pedrozo@davivienda.com (tu mismo correo)
ASUNTO:   [COPIA EXACTO DE ABAJO]
CUERPO:   [COPIA EXACTO DE ABAJO]
LABELS:   Después de enviar, marca con label "Testing_Reqs"
```

---

## PRUEBA 1: MEDELLÍN - PLIEGO DE CARGOS

**ASUNTO (copia exacto):**
```
REQ-TEST-001 | Oficio No. PC 0156 del 18 de septiembre de 2024
```

**CUERPO (copia exacto):**
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

**PASOS POST-ENVÍO:**
1. Envía el email
2. En Gmail, busca el email que acabas de enviar
3. Abre
4. Click las 3 puntitos (arriba a la derecha)
5. "Aplicar etiqueta" → Selecciona "Testing_Reqs"
6. ✅ Marcado
7. Espera 5 minutos (AppScript busca cada 5 min)
8. Ve a Sheets → ¿Aparece fila?

---

## PRUEBA 2: BOGOTÁ - EMPLAZAMIENTO

**ASUNTO (copia exacto):**
```
REQ-TEST-002 | Emplazamiento para Declarar No. 2024-5847
```

**CUERPO (copia exacto):**
```
Banco Davivienda S.A.

La Secretaría de Hacienda de Bogotá-Cundinamarca, mediante presente 
EMPLAZAMIENTO, requiere al Banco para que declare y presente las 
declaraciones de Impuesto de Industria y Comercio (ICA) 
correspondientes a los períodos no reportados.

NÚMERO DE EMPLAZAMIENTO: 2024-5847
PERIODOS FALTANTES: 01-2024, 02-2024, 03-2024
CONCEPTO: ICA omitido

Deberá presentar respuesta dentro del término legal de DIEZ (10) DÍAS.

Vencimiento: 28 de septiembre de 2024

Secretaría de Hacienda
Bogotá, Cundinamarca
```

---

## PRUEBA 3: CALI - REQUERIMIENTO

**ASUNTO (copia exacto):**
```
REQ-TEST-003 | Oficio 140565123 - Requerimiento de Información Exógena
```

**CUERPO (copia exacto):**
```
Banco Davivienda,

Por este medio, la Secretaría de Hacienda de Cali le requiere para 
que suministre información exógena relacionada con retenciones en 
la fuente practicadas durante el año 2023.

Oficio: 140565123
Fecha: 15 de septiembre de 2024
Solicitado por: Secretaría de Hacienda Municipal

Se solicita envío de información en formato Excel con los siguientes datos:
- NIT del retenedor
- Valor retenciones
- Conceptos

Vencimiento para respuesta: 15 de octubre de 2024

Secretaría de Hacienda
Cali, Valle del Cauca
```

---

## PRUEBA 4: BARRANQUILLA - CERTIFICACIÓN

**ASUNTO (copia exacto):**
```
REQ-TEST-004 | Solicitud de Certificación - No. 2024-08956
```

**CUERPO (copia exacto):**
```
Estimados,

La administración de Barranquilla solicita que se expida CERTIFICACIÓN 
sobre el estado de las obligaciones tributarias del Banco.

Número de solicitud: 2024-08956
Fecha de solicitud: 20 de septiembre de 2024
Objeto: Certificación de no adeudo

La certificación deberá ser presentada dentro de QUINCE (15) DÍAS 
a partir de esta comunicación.

Vencimiento: 05 de octubre de 2024

Secretaría de Hacienda
Barranquilla, Atlántico
```

---

## PRUEBA 5: SANTA MARTA - MÚLTIPLES FECHAS

**ASUNTO (copia exacto):**
```
REQ-TEST-005 | Oficio No. 5678 - Seguimiento a Requerimiento Anterior
```

**CUERPO (copia exacto):**
```
Banco Davivienda S.A.,

En seguimiento al requerimiento radicado en fecha 18/07/2024, la 
Secretaría de Hacienda de Santa Marta comunica lo siguiente:

Oficio No: 5678
Municipio: Santa Marta
Departamento: Magdalena
Fecha de radicación original: 18 de julio de 2024
Fecha de esta comunicación: 22 de septiembre de 2024

Se concede prórroga de TREINTA (30) DÍAS para la presentación de 
la información solicitada.

Nueva fecha de vencimiento: 22 de octubre de 2024

Atentamente,
Secretaría de Hacienda
Santa Marta, Magdalena
```

---

## PRUEBA 6: CARTAGENA - FORMATO DD/MM/YYYY

**ASUNTO (copia exacto):**
```
REQ-TEST-006 | PC 0089 del 22/09/2024 - Pliego de Cargos
```

**CUERPO (copia exacto):**
```
Banco Davivienda,

Mediante Pliego de Cargos emitido el día 22/09/2024, esta 
Secretaría de Hacienda de Cartagena notifica sobre presuntas 
infracciones en materia tributaria.

Número de Pliego: PC 0089
Fecha emisión: 22/09/2024
Vencimiento para respuesta: 12/10/2024

Favor presentar descargos ante esta dependencia.

Secretaría de Hacienda
Cartagena, Bolívar
```

---

## ✅ PLAN MAÑANA

### **Día 1: Prueba 1 (Medellín - Básico)**

1. Copia asunto de Prueba 1
2. Redacta nuevo email en Gmail
   - **De:** rebeca.pedrozo@davivienda.com
   - **Para:** rebeca.pedrozo@davivienda.com
   - **Asunto:** Pega exacto
   - **Cuerpo:** Pega exacto
3. Envía
4. En Gmail, busca el email que acabas de enviar
5. Abre el email
6. Click 3 puntitos (⋮) arriba a la derecha
7. "Aplicar etiqueta" → **Testing_Reqs**
8. Espera 5 minutos
9. Ve a tu Sheets (REQ LEGALES SEC HACIENDA)
10. ¿Aparece fila nueva? 👀

### **Día 2: Prueba 2 (Bogotá)**

Repite pasos 1-10 con Prueba 2

### **Día 3: Prueba 3 (Cali)**

Repite con Prueba 3

### **Día 4+: Pruebas 4, 5, 6**

Repite para validar edge cases

---

## 📝 DOCUMENTO PARA REGISTRAR

Crea archivo: `TEST_RESULTS_[FECHA].txt`

```
=== PRUEBA 1: MEDELLÍN - PLIEGO DE CARGOS ===
Fecha: [HOY]
Hora envío: [HORA]
Label aplicado: Testing_Reqs ✓

✅ APARECE FILA NUEVA: SÍ / NO

SI SÍ:
CAMPOS RELLENADOS AUTOMÁTICAMENTE:
☐ ID_REQ: REQ-2024-09-20-XXX
☐ EMAIL_TIMESTAMP: 20/09/2024 HH:MM
☐ ASUNTO_CORREO: REQ-TEST-001 | Oficio No...
☐ CIUDAD_MUNICIPIO: Medellín ✓
☐ DEPARTAMENTO: Antioquia ✓
☐ TIPO_OFICIO: Pliego de Cargos ✓
☐ FECHA_VENCIMIENTO: 08/10/2024 ✓

CAMPOS VACÍOS (normal):
- RESPONSABLE_ÁREA
- CASO
- ACCIÓN_EJECUTADA

SI NO:
☐ Chequear Logs de AppScript
☐ ¿Está el label "Testing_Reqs" bien escrito?
☐ ¿Esperé 5 minutos completos?
☐ Copiar error de Logs aquí: [...]

OBSERVACIONES:
[Describe algo raro]
```

---

## ✨ RESUMEN

**Lo que cambió:**
- ✅ Envías **de ti a ti** (rebeca → rebeca)
- ✅ Aplicas label **"Testing_Reqs"** manualmente
- ✅ AppScript busca en ese **label**
- ✅ Procesa los correos de prueba
- ✅ Sin afectar `notificacionesjudiciales@`

**Cuándo cambiar a producción:**
Una vez valides que TODO funciona, cambias el AppScript para:
- Buscar en label `Procesado` (de verdad)
- Y en las direcciones reales: notificacionesjudiciales@davivienda.co

---

**¿Claro? Mañana te envías a ti misma y listo.** ✅
