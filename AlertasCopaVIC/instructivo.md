# Instructivo — Alertas Copa VIC (Vencimientos DIAN)

Guía rápida para: (1) correr el ETL que limpia EXTRACT y llena LOAD (la hoja que alimenta Data Studio), y (2) probar el envío de correos sin que le lleguen a los encargados, jefes ni gerentes reales.

> Dónde se ejecuta todo: estos `.js` son el código de un proyecto de **Google Apps Script** ligado a la hoja de cálculo. Para correr cualquier función, abre la hoja de cálculo → menú **Extensiones → Apps Script**, selecciona el archivo indicado, elige la función en el desplegable de la barra superior (al lado del botón ▶️) y dale **Ejecutar**. La primera vez puede pedir autorización de permisos: acéptala.

---

## 1. Correr el ETL (Extract → Transform → Load)

Flujo de hojas: **EXTRACT** (donde pegas los datos crudos) → **TRANSFORM** (datos limpios) → **LOAD** (la que lee Data Studio).

### Opción rápida (recomendada)
- Archivo: `load.js`
- Función: **`ejecutarProcesoCompletoETL()`**

Esta función hace las 3 cosas en un solo clic:
1. `transformarDatosETL()` → limpia EXTRACT y llena TRANSFORM.
2. `cargarDatosLoad()` → toma TRANSFORM y llena/actualiza LOAD.
3. `asegurarHojaParametros()` → verifica que exista la hoja `PARAMETROS_ALERTA`.

Al terminar te sale un mensaje (toast) tipo "Proceso ETL completo... terminado".

### Opción paso a paso (si quieres revisar cada etapa)
1. Archivo `ETL.js` → función **`transformarDatosETL()`**. Limpia lo que pegaste en EXTRACT y lo pasa a TRANSFORM (normaliza nombre de compañía, valida fechas, separa nombre/correo de encargado y jefes, detecta duplicados exactos, etc.).
   - Revisa la hoja **`LOG_ANOMALIAS`** después de correrla: ahí quedan registradas filas con fecha inválida, sin persona encargada, sin municipio cuando el impuesto lo requiere, o filas duplicadas.
2. Archivo `load.js` → función **`cargarDatosLoad()`**. Toma lo que quedó en TRANSFORM y lo pasa a LOAD, calculando semáforo, días restantes, estado actual (Pendiente/Notificado/En proceso/Presentado), etc. **Esta es la hoja conectada a Data Studio.**

⚠️ Importante: siempre corre primero Transform y después Load (o usa `ejecutarProcesoCompletoETL()` que ya respeta ese orden). Si corres `cargarDatosLoad()` sin haber corrido antes el Transform, va a tomar datos viejos de TRANSFORM.

---

## 2. Probar los correos SIN que le lleguen a encargados/jefes/gerentes

El proyecto ya tenía pensado esto: existen versiones "producción" (mandan correos reales) y versiones "prueba" que leen de una hoja distinta llamada **`TRANSFORM2`** en vez de LOAD.

| Tipo de correo | Función de PRODUCCIÓN (correos reales) | Función de PRUEBA | Archivo |
|---|---|---|---|
| Recordatorio semanal a encargados | `enviarCorreosDiariosProduccion()` | `enviarCorreosDiariosPrueba()` | `notificaciones.js` |
| Reporte PDF semanal a jefes | `enviarReportesJefesProduccion()` | `enviarReportePruebaJefes()` | `reportesJefes.js` |

**Ojo con el detalle clave:** las funciones de "prueba" no filtran destinatarios automáticamente — simplemente leen los correos que estén escritos en las columnas `Encargado Email`, `Jefe1 Email` y `Jefe2 Email` **de la hoja `TRANSFORM2`**. O sea, la seguridad de que "solo te llegue a ti" depende de qué correos haya en esa hoja, no de la función en sí.

### Pasos para que SOLO te lleguen a ti

1. **Prepara la hoja `TRANSFORM2`:**
   - Si no existe, duplica la hoja `LOAD` (clic derecho en la pestaña → Duplicar) y renómbrala exactamente `TRANSFORM2` (debe tener las mismas columnas que TRANSFORM/LOAD).
   - Si ya existe, actualízala: copia ahí los datos actuales de LOAD para que la prueba sea representativa.

2. **Reemplaza los correos de esa hoja de prueba por el tuyo:**
   - Selecciona toda la columna **`Encargado Email`** y reemplaza los valores por tu propio correo.
   - Haz lo mismo con **`Jefe1 Email`** y **`Jefe2 Email`**.
   - Deja al menos una fila con `Estado Actual` distinto de `Presentado` y con `Fecha máxima de presentación` dentro de la semana actual (o ya vencida) — si no, la función no encuentra nada que enviar (así filtra `enviarCorreosDiarios`/`enviarReportesJefes`: solo manda si está dentro de la semana o ya vencida, y si no está "Presentado").

3. **Corre las funciones de prueba:**
   - `notificaciones.js` → **`enviarCorreosDiariosPrueba()`** → te llega el correo de recordatorio semanal (el de las tarjetitas con botones Notificado / En proceso / Presentado).
   - `reportesJefes.js` → **`enviarReportePruebaJefes()`** → te llega el PDF adjunto tipo el que reciben los jefes.

4. **Verifica que no se tocó producción:** como ambas funciones de prueba leen y escriben sobre `TRANSFORM2` (incluida la columna "Última Fecha Envío Recordatorio"), la hoja `LOAD` real queda intacta. Nada se marca como enviado en producción.

### 🚫 Qué NO correr mientras pruebas
No ejecutes `enviarCorreosDiariosProduccion()` ni `enviarReportesJefesProduccion()` — esas sí leen la hoja `LOAD` real y le llegan correos de verdad a encargados, jefes y gerentes.

---

## Resumen ultra rápido

- **¿Quiero pasar datos de Extract a Load (para Data Studio)?** → `load.js` → `ejecutarProcesoCompletoETL()`
- **¿Quiero ver cómo llegaría un correo sin molestar a nadie?** → primero deja tu correo en las columnas de email de `TRANSFORM2`, luego corre `enviarCorreosDiariosPrueba()` (notificaciones.js) y/o `enviarReportePruebaJefes()` (reportesJefes.js).
- **¿Ya probé y quiero mandar en serio?** → `enviarCorreosDiariosProduccion()` y `enviarReportesJefesProduccion()` (¡ojo, ahí sí llegan a todos!).
