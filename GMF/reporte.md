# Reporte de auditoría — `GMF/revision.r`

**Alcance:** revisión estática del archivo completo (2119 líneas). No se modificó el código, solo se identificaron los problemas.

**Resumen ejecutivo:** el script tiene **dos errores de sintaxis** que, en la práctica, impiden que R lo pueda siquiera *parsear* (analizar) completo. Cuando R ejecuta un archivo con `source()` (o `Rscript`), primero intenta convertir **todo el archivo** en código válido antes de correr una sola línea. Si hay un error de sintaxis en cualquier parte —aunque esté "escondido" dentro de una función que nunca se llama— el script entero falla a leerse y no se ejecuta nada, ni siquiera los `library(...)` del inicio. Estos dos puntos (hallazgos #1 y #2) son la prioridad máxima. Después de esos, hay bugs funcionales reales (variables no definidas, una condición que siempre da el mismo resultado, y un `unlink()` que borra el propio directorio de trabajo) y varios problemas de mantenibilidad (rutas e IDs hardcodeados repetidos decenas de veces).

---

## 🔴 CRÍTICOS — rompen la ejecución del script completo

### 1. Línea 490–496 — Nombres de columna sin comillas invertidas (backticks) → error de sintaxis

```r
datos_redenciones <- read_delim(redenciones_individuales, delim = ";", escape_double = FALSE, col_types = cols(NRO. ID COBIS. = col_character(), 
                                                                                                                       NRO CDT/CDTAS. = col_character(), 
                                                                                                                       VAL REDENCION CAPITAL = col_character(), 
                                                                                                                       VALOR GMF = col_character(), 
                                                                                                                       FEC PAGO = col_date(format = "%Y%m%d"),
                                                                                                                       NRO. ID CTA  = col_character(),
                                                                                                                       RECHAZOS Y OBSERVACIONES = col_character()),  trim_ws = TRUE, skip = 1)
```

**Qué está mal:** dentro de `cols(...)` se están usando nombres de argumento como `NRO. ID COBIS.`, `NRO CDT/CDTAS.`, `VAL REDENCION CAPITAL`, `VALOR GMF`, `FEC PAGO`, `NRO. ID CTA` y `RECHAZOS Y OBSERVACIONES`. En R, el nombre de un argumento en una llamada a función **debe ser un identificador válido** (sin espacios, puntos consecutivos raros, ni `/`), o si no, debe ir entre comillas invertidas (`` ` `` `). Ninguno de estos nombres está entre backticks. Esto es un **error de sintaxis** (`unexpected symbol`), no un error de ejecución.

**Por qué importa:** esta expresión vive dentro del cuerpo de la función `CDT <- function(datos, i, fondo_de_inversion, redenciones_individuales) {...}` (líneas 462–985), que R trata como **una sola expresión** al parsear. Un error de sintaxis aquí impide parsear toda la definición de la función `CDT`, y como el archivo se parsea completo antes de ejecutarse, **todo el script falla al cargar**, no solo esta parte. El `tryCatch` que envuelve este bloque (línea 487) no sirve de nada porque el error ocurre en la fase de *parsing*, antes de que exista la posibilidad de ejecutar el `tryCatch`.

**Cómo corregirlo:** envolver cada nombre de columna en backticks:

```r
datos_redenciones <- read_delim(redenciones_individuales, delim = ";", escape_double = FALSE,
  col_types = cols(
    `NRO. ID COBIS.` = col_character(),
    `NRO CDT/CDTAS.` = col_character(),
    `VAL REDENCION CAPITAL` = col_character(),
    `VALOR GMF` = col_character(),
    `FEC PAGO` = col_date(format = "%Y%m%d"),
    `NRO. ID CTA` = col_character(),
    `RECHAZOS Y OBSERVACIONES` = col_character()
  ), trim_ws = TRUE, skip = 1)
```

> Nota: además hay que confirmar que estos nombres coincidan **exactamente** (incluidos espacios y puntos) con el encabezado real del archivo `.DAT`/`.TXT` de redenciones, porque `read_delim` con `cols()` falla si el nombre de columna especificado no existe tal cual en el archivo.

---

### 2. Línea 2082 — Texto suelto sin `#` → error de sintaxis

```r
  tryCatch({writeData(wb, sheet = "Rtran_VS_redenciones", x = get(paste0("Cruza_rtran_redenciones_",indice), envir = .GlobalEnv) , startCol = 1, startRow = 1) },error = function(e){ 
    mensaje = paste('El archivo', "Cruza_rtran_redenciones_", 'no esta disponible')
    cat(mensaje, "\n")})
le estilo a las salidas
  addStyle(wb,  sheet = "Resumen", rows = 1, cols = 1, style = createStyle (textDecoration = "bold", fontSize = 12,  halign = "center"))
```

**Qué está mal:** la línea `le estilo a las salidas` no tiene `#` delante, por lo que R intenta interpretarla como código en lugar de como comentario. `le estilo a las salidas` no es una expresión R válida (son varios identificadores seguidos sin ningún operador entre ellos), así que produce un error de sintaxis. Por el texto, es evidente que era un comentario en español que quedó truncado (probablemente decía algo como `#Colocamos el estilo a las salidas`) y perdió el símbolo `#` al editar.

**Por qué importa:** igual que en el hallazgo #1, esta línea está dentro del cuerpo de `Funcion_Control_CDT <- function(indice) {...}` (líneas 2017–2112), una única expresión para el parser de R. Este segundo error de sintaxis **también** hace que el archivo completo no se pueda cargar.

**Cómo corregirlo:** anteponer `#`:

```r
#Colocamos el estilo a las salidas
```

---

## 🟠 ALTOS — bugs funcionales (una vez arreglados los errores de sintaxis)

### 3. Líneas 1014 y 1029 — Variables no definidas `Fondo_Inversion_Semana_6` y `Redenciones_CDTS_Semana_6`

```r
CDT(CDT_Completo, 6,Fondo_Inversion_Semana_6,Redenciones_CDTS_Semana_6) 
...
CDT(CDT_Completo, 6,Fondo_Inversion_Semana_6,Redenciones_CDTS_Completo_TXT)
```

**Qué está mal:** en las líneas 995–1000 y 1002–1007 se definen `Fondo_Inversion_Semana_1..5` y `Fondo_Inversion_Completo`, y `Redenciones_CDTS_Semana_1..5` y `Redenciones_CDTS_Completo` (y su variante `_TXT`). **Nunca se define** `Fondo_Inversion_Semana_6` ni `Redenciones_CDTS_Semana_6`; el nombre correcto para el índice 6 ("Completo") es `Fondo_Inversion_Completo` y `Redenciones_CDTS_Completo` / `Redenciones_CDTS_Completo_TXT`.

**Por qué importa:** cuando la función `CDT()` intenta usar el parámetro `fondo_de_inversion` (que recibe la promesa `Fondo_Inversion_Semana_6`), R lanza `object 'Fondo_Inversion_Semana_6' not found`. Ese error sí queda atrapado por el `tryCatch` interno de la sección "Cruce fondo de inversión" (línea 735), así que no detiene el script, pero **el cruce contra el fondo de inversión para la corrida "Completo" (índice 6) nunca se ejecuta**, y el mensaje de error que se imprime (`'El archivo', fondo_de_inversion, 'no esta disponible'`) es engañoso porque el problema no es que falte el archivo, sino que la variable con el nombre del archivo no existe.

**Cómo corregirlo:**

```r
CDT(CDT_Completo, 6, Fondo_Inversion_Completo, Redenciones_CDTS_Completo) 
...
CDT(CDT_Completo, 6, Fondo_Inversion_Completo, Redenciones_CDTS_Completo_TXT)
```

---

### 4. Línea 1530 — `ifelse(is.na(.), 0, 0)` siempre devuelve 0, sin importar el valor original

```r
comprobante_impuestos_funcion = data.frame(comprobante_impuestos_funcion) %>% mutate(across(c(18:19), ~ as.numeric(.)),
                                                                                     across(c(18:19), ~ ifelse(is.na(.),0,0)),
                                                                                     FECHA.CONTABLE = fecha_impuestos,
                                                                                     DESCRIPCION = descripcion_impuestos)
```

**Qué está mal:** el patrón `ifelse(is.na(x), 0, 0)` es lógicamente equivalente a escribir simplemente `0`: sea el valor `NA` o no, el resultado siempre es `0`. Esto contrasta con el patrón correcto usado unas líneas antes para `comprobante_1` (línea 1389): `ifelse(is.na(.), 0, .)`, que sí conserva el valor original cuando no es `NA` y solo reemplaza los `NA` por `0`. Todo indica que esta línea es una copia de ese patrón con un error de tecleo: el tercer argumento debería ser `.` (el valor original), no `0`.

**Por qué importa:** las columnas 18 y 19 son `DEBITO` y `CREDITO` de la plantilla `comprobante_impuestos` (ver nombres en línea 48–50). Con el código actual, **cualquier valor que ya viniera cargado en esas columnas desde la hoja de Google Sheets `Estructura_Impuestos` se borra y se reemplaza por 0**, en lugar de solo limpiar las celdas vacías. Si la plantilla de origen alguna vez trae valores por defecto en `DEBITO`/`CREDITO` para alguna fila que el código no sobrescribe explícitamente más adelante (líneas 1536 en adelante solo tocan filas puntuales, no todas), esos valores se pierden silenciosamente.

**Cómo corregirlo:**

```r
across(c(18:19), ~ ifelse(is.na(.), 0, .))
```

Si la intención real era "resetear siempre a 0 sin importar el valor previo" (limpieza intencional de la plantilla), entonces el código no es un bug pero sí es engañoso: debería escribirse directamente como `across(c(18:19), ~ 0)` para dejar clara la intención y evitar que un futuro lector piense que se está comprobando algo.

---

### 5. Línea 1503 — `unlink(local_directory, recursive = TRUE)` borra el directorio de trabajo actual, y luego se guardan archivos ahí mismo

```r
#================================= Datos comprobante impuestos ======================
#Eliminar los insumos de la corrida pasada
unlink(local_directory, recursive = TRUE)
```

**Qué está mal:** en la línea 167–170, `input <- local_directory` y `setwd(input)` hacen que **el directorio de trabajo actual sea exactamente `local_directory`**. En la línea 1503 (ya avanzado el script, después de haber usado todos los insumos) se ejecuta `unlink(local_directory, recursive = TRUE)`, es decir, se borra recursivamente la carpeta que es al mismo tiempo el directorio de trabajo activo. Inmediatamente después, el script sigue usando rutas relativas para **guardar** archivos ahí (`saveWorkbook(wb, 'Comprobante_impuestos.xlsx', ...)` en la línea 1676, y varios `saveWorkbook` más adelante), asumiendo que ese directorio sigue existiendo y es válido como destino de escritura.

**Por qué importa:** en Windows, borrar (o intentar borrar) el directorio que es el "current working directory" de un proceso puede fallar parcialmente, dejar el directorio en un estado inconsistente, o (si tiene éxito) hacer que las escrituras relativas posteriores fallen porque la carpeta ya no existe. Es una secuencia de operaciones riesgosa: la limpieza de insumos se hace **a mitad de script**, cuando el propio proceso R sigue "parado" dentro de esa carpeta y todavía necesita escribir en ella.

**Cómo corregirlo:** dos opciones seguras:
- Mover el `unlink()` de los insumos de entrada al **final** del script, después de que todos los `saveWorkbook`/`drive_upload` hayan terminado; o
- Antes de borrar, hacer `setwd()` a una carpeta distinta (p. ej. una carpeta temporal o la carpeta padre), borrar `local_directory`, y luego crear de nuevo la carpeta de salida antes de escribir en ella:

```r
setwd(dirname(local_directory))
unlink(local_directory, recursive = TRUE)
dir.create(local_directory)
setwd(local_directory)
```

---

## 🟡 MEDIOS — mantenibilidad y fragilidad (no rompen la ejecución, pero son fuente probable de errores futuros)

### 6. Ruta absoluta `"D:/Mis Documentos/GMF/Comprobante_Completo/ENTRADA"` hardcodeada en 7 lugares distintos

Líneas: **150** (`local_directory <- "..."`), **1678**, **1773**, **1860**, **1944**, **2015**, **2108**.

**Qué está mal:** el script ya define la variable `local_directory` (línea 150) precisamente para no repetir la ruta, pero los seis `drive_upload(...)` posteriores (líneas 1678, 1773, 1860, 1944, 2015, y el `paste0(...)` de la línea 2108) vuelven a escribir la ruta completa como texto literal en lugar de reutilizar `local_directory`.

**Por qué importa:** el propio comentario del encabezado del archivo (líneas 1–4) le pide al usuario que reemplace esta ruta con Ctrl+F por la carpeta de su equipo. Si alguien reemplaza solo la línea 150 (donde parece estar centralizada la ruta) y no se da cuenta de que hay 6 apariciones más, esos `drive_upload()` intentarán leer un archivo que no existe en la ruta vieja/incorrecta, y fallarán en silencio dentro de bloques sin manejo de error explícito.

**Cómo corregirlo:** usar `local_directory` (o `file.path(local_directory, "Comprobante_impuestos.xlsx")`) en cada uno de esos `drive_upload()`, por ejemplo:

```r
drive_upload(file.path(local_directory, "Comprobante_impuestos.xlsx"), path = as_id(folder_id_salida), name = "Comprobante_impuestos.xlsx", overwrite = TRUE)
```

Y en la línea 2108:

```r
directorio = file.path(local_directory, paste0("Control_CDT_Semana_", indice, ".xlsx"))
```

---

### 7. ID del Google Sheet `"1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM"` repetido 36 veces

**Qué está mal:** el mismo ID de spreadsheet aparece hardcodeado como literal de texto en 36 llamadas distintas a `read_sheet(...)` a lo largo de las líneas 43–131.

**Por qué importa:** si algún día cambia el spreadsheet de origen (por ejemplo, se hace una copia para otro equipo o se migra), hay que editar 36 lugares uno por uno, con alto riesgo de dejar alguno desactualizado y que el script lea de dos hojas distintas sin que nadie lo note.

**Cómo corregirlo:** declarar una constante al inicio del script y reutilizarla:

```r
SHEET_ID <- "1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM"
...
comprobante_1 = read_sheet(SHEET_ID, sheet = 'Estructura', ...)
```

---

### 8. Líneas 829, 851 y 854 — Comparación `rpt_tip_id == "03"` nunca se cumple (código muerto)

```r
Resumen_cancelacion_plazo_FIJO = VIVOS_G3_CDT %>% mutate(naturaleza = ifelse(rpt_tip_id == "03" | rpt_tip_id == 3,"Juridico", "Natural" ))
```

**Qué está mal:** en la definición de la función `CDT` (línea 476), la columna `rpt_tip_id` se lee explícitamente como `col_number()`, es decir, es **numérica**, no texto. Al comparar un número contra el string `"03"` (`rpt_tip_id == "03"`), R convierte el número a texto para comparar (`as.character(3)` → `"3"`), y `"3" != "03"`, por lo que esa mitad de la condición **nunca es verdadera**. La condición funciona únicamente gracias al segundo término (`rpt_tip_id == 3`), que sí compara número contra número.

**Por qué importa:** no es un bug que cambie el resultado (porque el segundo término cubre el caso), pero es código muerto que confunde a cualquiera que lea el script pensando que ambas ramas aportan algo, y sugiere que en algún momento `rpt_tip_id` se leía como texto (donde `"03"` sí habría tenido sentido) y el tipo de columna cambió sin limpiar esta comparación. Aparece igual en tres lugares (829, 851, 854), así que si se corrige uno hay que corregir los tres.

**Cómo corregirlo:** simplificar, dejando solo la comparación numérica:

```r
naturaleza = ifelse(rpt_tip_id == 3, "Juridico", "Natural")
```

---

### 9. Uso extensivo de `assign()`/`get()` con nombres construidos dinámicamente y `<<-` para efectos globales

Ejemplos: líneas 305–318 (`Cheques_oficina`), 938–970 (`CDT`), 1211–1219 (`R_cheques_funcion`), 1393–1398 (`control`), etc.

**Qué está mal:** en vez de que las funciones devuelvan (`return`) sus resultados y el llamador los asigne a una variable con nombre claro, el script construye nombres de variable como texto (`paste("R_cheques_oficina_", index, sep = "")`) y los inyecta al entorno global con `assign(..., envir = .GlobalEnv)`. De forma simétrica, funciones como `control()` y `R_cheques_funcion()` modifican `comprobante_1` directamente con `<<-` en vez de recibirlo y devolverlo.

**Por qué importa:** este patrón hace que:
- Sea imposible saber, con solo mirar la firma de una función, qué variables va a crear o modificar (hay que leer todo el cuerpo).
- Cualquier error de tipeo en el nombre generado (como ya ocurrió en el hallazgo #3) sea invisible para R en tiempo de "parseo" y solo se note en ejecución, o ni siquiera se note si queda atrapado por un `tryCatch`.
- No se pueda probar (`test`) ninguna de estas funciones de forma aislada, porque dependen y modifican el estado global del script completo.

**Cómo corregirlo (a mediano plazo, no urgente):** hacer que cada función `return()` una lista con sus resultados, y que el código que la llama la asigne explícitamente a variables con nombre fijo, o a una lista indexada (`resultados[[i]] <- CDT(...)`), evitando `assign`/`get`/`<<-`. Esto es un cambio de diseño más grande, se menciona aquí como recomendación, no como corrección puntual de una línea.

---

### 10. Índices de fila/columna "mágicos" repetidos por todo el script sin constantes con nombre

Ejemplos: línea 115–122 (listas de índices de fila para las fechas de cada semana), línea 1704–1762 (decenas de `addStyle` con listas idénticas de números de fila repetidas), línea 291–292 (`resultado_cheques_oficina[5,4]`), línea 1536 en adelante (`comprobante_impuestos_funcion[9,19]`, `[10,18]`, etc.).

**Qué está mal:** el diseño del comprobante depende de que ciertas filas y columnas de la hoja `Estructura` (Google Sheets) mantengan siempre el mismo orden y posición (por ejemplo, la fila 5 de `resultado_cheques_oficina` siempre es "reposiciones", o la fila 9 de `comprobante_impuestos_funcion` siempre corresponde a "Cheques girados 9710"). Estos números están escritos directamente como literales en decenas de lugares, sin ningún comentario que documente su significado salvo comentarios sueltos como `#fila_comprobante_me = 1` (línea 1535) que ni siquiera son código, son anotaciones.

**Por qué importa:** si alguien reordena o agrega una fila en la hoja `Estructura` de Google Sheets (por ejemplo, para agregar un nuevo concepto), **todos** estos índices numéricos quedan desalineados de forma silenciosa: el script no lanzará ningún error, simplemente escribirá los valores en la fila equivocada del comprobante final, generando un reporte financiero incorrecto sin ninguna señal de que algo falló.

**Cómo corregirlo:** no es una corrección de una línea puntual; se recomienda centralizar estos índices en constantes con nombre al inicio del script (p. ej. `FILA_CHEQUES_9710 <- 1`) o, mejor aún, ubicar las filas dinámicamente por el contenido de la columna `Consecutivo`/`MOTIVO` en vez de por número fijo de fila.

---

## Resumen priorizado

| # | Línea(s) | Severidad | Tipo | Efecto |
|---|----------|-----------|------|--------|
| 1 | 490–496 | 🔴 Crítico | Error de sintaxis | El archivo completo no se puede parsear/ejecutar |
| 2 | 2082 | 🔴 Crítico | Error de sintaxis | El archivo completo no se puede parsear/ejecutar |
| 3 | 1014, 1029 | 🟠 Alto | Variable no definida | El cruce de fondo de inversión para "Completo" nunca corre |
| 4 | 1530 | 🟠 Alto | Lógica incorrecta | Borra valores válidos de DEBITO/CREDITO en vez de solo limpiar NA |
| 5 | 1503 | 🟠 Alto | Operación riesgosa | Borra el directorio de trabajo activo antes de escribir en él |
| 6 | 150, 1678, 1773, 1860, 1944, 2015, 2108 | 🟡 Medio | Duplicación | Ruta hardcodeada 7 veces, riesgo de quedar desincronizada |
| 7 | 43–131 (36 veces) | 🟡 Medio | Duplicación | ID de Sheet hardcodeado 36 veces |
| 8 | 829, 851, 854 | 🟡 Medio | Código muerto | Comparación que nunca se cumple, confunde la lectura |
| 9 | Múltiples | 🟡 Medio | Diseño | `assign`/`get`/`<<-` dificulta debug y pruebas |
| 10 | Múltiples | 🟡 Medio | Diseño | Índices de fila/columna mágicos, frágiles ante cambios de la hoja fuente |

**Recomendación de orden de trabajo:** corregir primero los hallazgos #1 y #2 (sin esto, nada más se puede siquiera probar), luego #3, #4 y #5 (bugs funcionales con impacto directo en las cifras y en la estabilidad de la corrida), y finalmente considerar los puntos #6–#10 como mejoras de mantenibilidad a abordar cuando haya tiempo, no como bloqueantes.
