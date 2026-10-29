# Guion de sustentación — Prueba Técnica Analítica

**Duración objetivo:** 10–12 min
**Cómo usarlo:** lee los bloques en orden. Lo que está en `>` es para decir casi literal. Lo demás es apoyo.

**Regla de oro:** nunca des una cifra sola. Siempre di la cifra **y qué significa para el negocio**. Esa es toda la diferencia entre "sé usar pandas" y "sé analizar".

---

## BLOQUE 1 — Apertura (30 seg)

> "La base son 1.000 cotizaciones de enero a septiembre de 2026, cuatro segmentos de vehículos. Lo que voy a mostrar sigue el orden de la prueba: primero cómo dejé la base confiable, luego los indicadores, y al final lo que esos indicadores dicen del negocio.
>
> Adelanto la conclusión para que la tengan en mente mientras explico: **el segmento que más vende es el que más plata pierde, y el que menos vende es el más rentable.** Todo lo demás sale de ahí."

💡 **Por qué abrir así:** cuando sueltas el hallazgo de entrada, el que escucha deja de evaluar si sabes código y empieza a seguir el argumento. Ganas autoridad en los primeros 20 segundos.

---

## BLOQUE 2 — Por qué la limpieza no es trámite (1 min)

> "Empecé por calidad de datos, y no como formalidad. En esta base hay un registro que, si no lo corriges, te daña el indicador principal de toda la prueba.
>
> Es la cotización COT-00731: está marcada como **no vendida**, pero tiene 8 millones y medio en siniestros incurridos. Eso es imposible. Sin póliza emitida no puede haber siniestro.
>
> Si lo dejas pasar, ese siniestro entra al numerador de la siniestralidad sin que entre ninguna prima al denominador. Inflas la siniestralidad con un costo que no tiene ingreso detrás."

💡 **Insight para rematar:**
> "Ese registro es el que separa a alguien que revisa nulos y duplicados de alguien que revisa **coherencia lógica entre campos**. Los nulos los encuentra `isna()`. Eso no."

---

## BLOQUE 3 — Qué encontré en calidad (1.5 min)

> "Encontré seis tipos de problema:
>
> Tres duplicados exactos: COT-00027, 428 y 762. Idénticos en todas las columnas, no son versiones distintas. Los quité y dejé la primera ocurrencia.
>
> La columna SEGMENTO traía **siete categorías cuando en realidad son cuatro**: 'AUTOMOVILES' con espacios, 'camionetas' en minúscula, 'MOTO' en singular. Si no normalizas eso, agrupas sobre siete grupos y todos los indicadores quedan calculados sobre bases incompletas.
>
> Modelo de vehículo en 0 y en 2031. Una prima negativa de menos 500.000 y otra en cero.
>
> Y una prima de 35 millones en un carro asegurado por 80 millones."

🔑 **Este es tu mejor momento técnico. No lo pases rápido:**
> "Esa última la quiero explicar porque ahí está el criterio. Un IQR normal te la marca como atípico, sí, pero no te dice si es un riesgo caro o un error de digitación. Entonces calculé la **tasa técnica**: prima sobre valor asegurado.
>
> Esa cotización da 43,7%. La mediana del portafolio es 2,2%. Nadie asegura un carro cobrando el 44% de su valor. No es un riesgo caro, es un dedazo.
>
> Y lo contrario también: había tres valores asegurados altísimos en camionetas y motos que el IQR marcaba como atípicos, pero su tasa técnica era normal. Esos son vehículos de gama alta, son reales. **Los dejé.**"

💡 **El insight que vale:** un atípico estadístico no es lo mismo que un error. La tasa técnica es la que te deja distinguirlos. Eliminar los tres valores altos habría sido botar información buena.

---

## BLOQUE 4 — Criterio de limpieza (30 seg)

> "Una decisión transversal: **no eliminé ninguna fila por tener un dato malo.** Anulo el campo afectado y dejo el registro vivo.
>
> La razón es directa: si una cotización tiene la prima mal pero sí sabemos si se vendió o no, esa fila todavía sirve para calcular conversión. Si la borro completa, le estoy metiendo sesgo a la tasa de conversión para arreglar un problema de prima.
>
> La base queda en 997 registros. Solo se perdieron los tres duplicados."

---

## BLOQUE 5 — Los faltantes (30 seg)

> "Quedaron 7 faltantes entre valor asegurado y prima: menos del 0,6% de la base.
>
> **No los imputé**, y es a propósito. Imputar una prima con el promedio del segmento es inventar ingreso donde no lo hay, y ese ingreso inventado entra directo al denominador de la siniestralidad, que es el indicador más delicado de toda la prueba.
>
> Con un volumen así de bajo, sacarlos del cálculo es más transparente que estimarlos."

💡 **Si te preguntan "¿y si fueran el 20%?":**
> "Ahí sí imputaría, pero no con el promedio general: con la mediana por segmento y rango de valor asegurado, y además dejaría una bandera marcando cuáles fueron imputados para poder medir después si movieron el resultado."

---

## BLOQUE 6 — Los números (1.5 min)

> "Con la base limpia, el portafolio completo: **36,61% de conversión y 78,78% de siniestralidad.**
>
> Ese 78,78% suena manejable, pero no incluye gastos ni comisiones. En seguros eso pesa normalmente entre 25 y 30 puntos. O sea que el portafolio global está operando **cerca del punto de equilibrio técnico**, no cómodo."

Luego vas segmento por segmento:

| Segmento | Conversión | Siniestralidad | Resultado técnico |
|---|---|---|---|
| Pesados | 43,33% | 81,88% | +69.636.000 |
| **Motos** | **42,11%** | **112,02%** | **−13.312.000** |
| Automóviles | 37,43% | 70,00% | +72.412.000 |
| Camionetas | 26,00% | 61,13% | +60.735.000 |

> "Miren la tabla de arriba hacia abajo: la conversión va bajando y la siniestralidad también. **Van en sentido contrario.** Los que más venden son los que peor resultado técnico dejan."

🔑 **Frase clave, dila despacio:**
> "Ese cruce es todo el análisis. Si miras solo conversión, motos y pesados parecen los mejores segmentos del portafolio. Mirando las dos columnas juntas, la lectura se voltea completa."

---

## BLOQUE 7 — Hallazgo 1: Motos (1.5 min)

> "Motos es el único segmento con siniestralidad por encima de 100%. **112,02%**: por cada peso de prima que emite, paga 1,12 pesos en siniestros. Y eso es antes de gastos.
>
> Aporta el 12,4% de las primas del portafolio, pero el 17,6% de los siniestros. Deja un resultado técnico de menos 13 millones.
>
> Y al mismo tiempo es el segundo segmento que más convierte: 42,11%, muy por encima del promedio de 36,6%."

🔑 **El insight fuerte — este es el que te distingue:**
> "Aquí hay algo que quiero señalar porque cambia la lectura. Motos **ya es el segmento con la tarifa relativamente más alta**: su tasa técnica es 4,06% del valor asegurado, contra 1,65% en camionetas. O sea, en motos ya se cobra proporcionalmente el doble que en el resto.
>
> Y aun así pierde. Entonces el problema no es que se cobre poco en términos generales: es que **la frecuencia de siniestro es del 41,3%** y la tarifa no la está recogiendo.
>
> Y la conversión alta en un producto que pierde plata no es buena noticia, es una señal de **antiselección**: estás barato frente al mercado y por eso te está entrando el riesgo que los demás están rechazando."

💡 **Remate:** "El crecimiento de motos no es un resultado comercial. Es exposición mal tarificada."

---

## BLOQUE 8 — Hallazgo 2: Pesados y motos fallan distinto (1 min)

> "El segundo hallazgo es que **pesados y motos tienen siniestralidad alta por razones opuestas**, y eso importa porque no se arreglan igual.
>
> Pesados: la **menor** frecuencia de siniestro de todos, 32,3%. Pero su siniestro promedio es de **15 millones**, cinco veces el de cualquier otro segmento.
>
> Motos: siniestros de menos de 3 millones, pero pasan todo el tiempo.
>
> La consecuencia práctica: la siniestralidad de pesados es **volátil**, depende de poquísimos eventos grandes, un siniestro más la mueve varios puntos. La de motos es **estable**, o sea estructural, o sea predecible."

🔑 **El cierre del bloque:**
> "Por eso la recomendación es distinta para cada uno. Pesados no necesita subir tarifa, necesita **reaseguro y control de acumulación**. Motos sí necesita **retarifación y reglas de suscripción**.
>
> Si uno mira solo la columna de siniestralidad, 112% y 82% se ven como el mismo problema de grado distinto. Frecuencia y severidad muestran que son problemas de naturaleza diferente."

---

## BLOQUE 9 — El hallazgo extra: antigüedad (1 min)

> "Hice un corte adicional que no pedía la prueba, por antigüedad del vehículo:"

| Antigüedad | Conversión | Siniestralidad |
|---|---|---|
| 0–2 años | 40,7% | **35,0%** |
| 3–5 años | 38,6% | 103,5% |
| 6–8 años | 33,6% | 53,8% |
| 9+ años | 34,2% | **125,0%** |

> "Los vehículos de 9 años o más dan **125% de siniestralidad**. Los de 0 a 2 años dan **35%**. Es una diferencia de casi cuatro veces en el costo del riesgo.
>
> Pero miren la columna de conversión: **40,7% contra 34,2%.** Prácticamente igual.
>
> Eso quiere decir que la compañía está cobrando casi lo mismo por riesgos que cuestan cuatro veces más."

🔑 **Por qué este hallazgo vale oro:**
> "Y esto es lo importante: como la conversión es plana, **hay espacio para subir precio en los vehículos viejos sin perder el negocio bueno.** El dato ya está en la base, solo que no está en la tarifa. Esa es la palanca más concreta que encontré en todo el análisis."

---

## BLOQUE 10 — La pregunta trampa (1 min)

La prueba pregunta si crecer en ventas es necesariamente buen resultado. **Es la pregunta que separa perfiles.**

> "No. Y esta base lo demuestra sin que haya que teorizar.
>
> Motos convierte 42,11% y deja **menos** 13 millones. Camionetas convierte 26% —el peor comercialmente de los cuatro— y deja **más** 60 millones.
>
> O sea: el peor segmento comercial es el mejor segmento técnico. Son literalmente opuestos.
>
> La razón de fondo es que en seguros **el ingreso se reconoce cuando emites, pero el costo real se conoce después.** Por eso una venta puede verse bien hoy y ser pérdida en seis meses. Crecer solo es buen resultado si la siniestralidad esperada de ese negocio nuevo alcanza para cubrir siniestros y gastos."

💡 **Remate de negocio, no de analista:**
> "Y eso deja una recomendación concreta: si quieren crecer sin dañar el resultado, el crecimiento tiene que irse **a camionetas**, donde el riesgo sí está bien pagado. Ahí están dejando negocio sobre la mesa. Y motos se corrige por tarifa y suscripción, no por volumen."

---

## BLOQUE 11 — SQL (40 seg)

> "Para la integración usé **LEFT JOIN**, encadenado: de cotizaciones a información de cliente por ID_COTIZACION, y de ahí a vehículo por CODIGO_VEHICULO, porque ese campo vive en la tabla de cliente.
>
> Tiene que ser LEFT y no INNER por una razón de negocio: con INNER perderías las cotizaciones que no tienen datos complementarios, y la tasa de conversión te quedaría calculada sobre una base incompleta. Los nulos que quedan no son un problema, son información: te dicen qué tan completa está la data de apoyo."

🔑 **Agrega esto, casi nadie lo menciona:**
> "Y antes de usar esa consulta en producción hay que validar algo: que INFORMACION_CLIENTE tenga **una sola fila por cotización**. Si tiene más de una, el join te duplica cotizaciones y te infla primas y siniestros sin que te des cuenta. Dejé el query de validación en el documento."

💡 Si quieres cerrar con altura: *"Y las variables que trae ese cruce —edad, ciudad, marca, línea— son justo las que me hacían falta para pasar de un diagnóstico por segmento a una tarifa por perfil."*

---

## BLOQUE 12 — IA (40 seg)

> "Para el punto de inteligencia artificial propuse un modelo de pricing técnico separado en dos: **frecuencia y severidad.**
>
> Y lo separé a propósito, conectado con el hallazgo 2: como motos falla por frecuencia y pesados por severidad, un modelo único los mezclaría y no resolvería ninguno de los dos bien.
>
> Para el algoritmo propondría GLM como línea base y gradient boosting para capturar interacciones, pero con **interpretabilidad obligatoria**: en tarifa cada factor tiene que poder sustentarse ante el regulador. Un modelo que sube el precio y no puede explicar por qué, no es usable en seguros así prediga bien."

🔑 **Honestidad que suma, no resta:**
> "Y algo que dejé explícito: con 997 registros y nueve meses **no se puede entrenar ese modelo todavía.** Lo que sí se puede hacer hoy es el diagnóstico y el diseño. El modelo necesita 24 a 36 meses de histórico."

💡 Decir "con estos datos no alcanza" te suma mucho más de lo que crees. Es la diferencia entre alguien que vende humo y alguien que sabe medir el alcance de sus datos.

---

## BLOQUE 13 — Cierre (30 seg)

> "Resumo en tres frases:
>
> **Uno:** motos es el problema urgente, único segmento en pérdida técnica y además creciendo.
>
> **Dos:** camionetas es la oportunidad que no están aprovechando, mejor siniestralidad y peor conversión.
>
> **Tres:** la antigüedad del vehículo es una variable de riesgo que ya está en los datos y no está en el precio.
>
> La recomendación de una línea: **mover el crecimiento hacia camionetas mientras se retarifica motos por antigüedad y valor asegurado.**"

---

# Preguntas que te van a hacer

**"¿Por qué no imputaste los faltantes?"**
> "Porque eran menos del 0,6% y porque imputar prima contamina directo la siniestralidad. Con ese volumen, el riesgo de distorsionar es mayor que el beneficio de completar."

**"¿Por qué motos y no pesados, si pesados tiene el 43% de las primas?"**
> "Por el signo, no por el tamaño. Pesados tiene siniestralidad alta pero sigue dejando resultado positivo, 69 millones. Motos es el único en negativo. Y además pesados falla por severidad, que se maneja con reaseguro; motos falla por tarifa, que es un problema estructural del producto."

**"¿Y si el siniestro de esa cotización no vendida era real?"**
> "Podría ser un error de marcación en el campo VENTA en vez de un error en el siniestro. Por eso lo dejé documentado como un supuesto explícito y no como una corrección silenciosa: en un caso real eso se valida contra el sistema de emisión antes de tocarlo. Elegí la opción conservadora, que es no meter costo sin prima."

**"¿Esto es significativo estadísticamente?"**
> "Con 997 registros, las diferencias grandes sí: 112% contra 61% de siniestralidad no es ruido. Las chiquitas no las forcé — por eso en la conversión mensual, que se mueve entre 30% y 40%, dije explícitamente que **no hay evidencia de tendencia**, porque con 111 cotizaciones por mes esa variación cabe dentro del azar."

**"¿Usaste IA?"**
> "Sí, como apoyo para revisar el enfoque y pulir redacción. Las reglas de limpieza, el cálculo de indicadores y la interpretación son mías. Cada cifra del documento sale del notebook y es reproducible corriéndolo."

---

# Cifras que debes tener en la cabeza

| Dato | Valor |
|---|---|
| Base | 1.000 → **997** |
| Conversión portafolio | **36,61%** |
| Siniestralidad portafolio | **78,78%** |
| Motos | 42,11% conv / **112,02%** sin / **−13,3 M** |
| Camionetas | 26,00% conv / **61,13%** sin / +60,7 M |
| Pesados | severidad **15 M** (5x el resto) |
| Antigüedad | 9+ años: **125%** vs 0–2: **35%** |
| Tasa técnica motos | **4,06%** (la más alta y aun así pierde) |
| Dedazo detectado | tasa **43,7%** vs mediana **2,2%** |

---

# Tres cosas que NO hagas

1. **No recites la limpieza paso a paso.** Di el criterio (no borro filas, anulo campos) y salta al caso de COT-00731 y al de la tasa técnica. Eso es lo que demuestra criterio; lo demás es `drop_duplicates()`.

2. **No des conversión sin siniestralidad.** Nunca. Ese cruce es el eje de todo tu análisis, y mencionarlas por separado lo desarma.

3. **No infles lo de IA.** Decir "con estos datos no alcanza todavía" te hace ver más maduro que proponer un modelo que no se puede entrenar.
